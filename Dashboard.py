import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import tempfile
import os
import base64
from io import BytesIO

# 1. Configuracion institucional
st.set_page_config(page_title="Proyección Comercial CE", layout="wide")

COLOR_ACTUAL = "#7F8C8D" 
COLOR_CE = "#27AE60"     

# Imagen institucional definida en el codigo, no cargada por el usuario.
RUTA_IMAGEN_TITULO = os.path.join(os.path.dirname(__file__), "LogoCENS.png")
with open(RUTA_IMAGEN_TITULO, "rb") as archivo_imagen:
    IMAGEN_TITULO = base64.b64encode(archivo_imagen.read()).decode("ascii")

with st.container():
    st.markdown(
        """
        <style>
        :root {
            --verde-cens: #27AE60;
            --verde-oscuro-cens: #145A32;
            --verde-claro-cens: #E8F5E9;
            --gris-fondo-cens: #F4F6F5;
            --gris-borde-cens: #DDE5E0;
        }
        .stApp {
            background: linear-gradient(135deg, #FFFFFF 0%, var(--gris-fondo-cens) 100%);
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
        }
        .encabezado-app {
            align-items: center;
            background: linear-gradient(110deg, #FFFFFF 0%, var(--verde-claro-cens) 100%);
            border: 1px solid #CFE7D5;
            border-right: 8px solid var(--verde-cens);
            border-radius: 14px;
            box-shadow: 0 8px 24px rgba(20, 90, 50, 0.08);
            display: flex;
            gap: 18px;
            justify-content: space-between;
            margin-bottom: 20px;
            min-height: 142px;
            padding: 24px 30px;
        }
        .encabezado-app img {
            flex: 0 0 auto;
            height: 94px;
            object-fit: contain;
            order: 2;
            width: 94px;
        }
        .encabezado-app-contenido {
            min-width: 0;
            order: 1;
        }
        .encabezado-app h1 {
            color: var(--verde-oscuro-cens);
            font-size: clamp(1.8rem, 3vw, 2.7rem);
            line-height: 1.12;
            margin: 0;
        }
        .encabezado-app p {
            color: #1B4332;
            margin: 12px 0 0;
            font-size: 1.05rem;
        }
        h2, h3 {
            color: var(--verde-oscuro-cens);
        }
        hr {
            border-color: var(--gris-borde-cens);
            margin: 1.5rem 0;
        }
        [data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid var(--gris-borde-cens);
            border-left: 4px solid var(--verde-cens);
            border-radius: 10px;
            box-shadow: 0 4px 14px rgba(34, 64, 48, 0.06);
            padding: 16px 18px;
        }
        [data-testid="stMetricLabel"] {
            color: #5D6D64;
        }
        [data-testid="stSidebar"] {
            background: #F1F4F2;
            border-right: 1px solid var(--gris-borde-cens);
        }
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: var(--verde-oscuro-cens);
        }
        .stButton > button,
        .stDownloadButton > button {
            background-color: var(--verde-cens);
            border: 1px solid var(--verde-cens);
            border-radius: 8px;
            color: #FFFFFF;
            font-weight: 600;
        }
        .stButton > button:hover,
        .stDownloadButton > button:hover {
            background-color: var(--verde-oscuro-cens);
            border-color: var(--verde-oscuro-cens);
            color: #FFFFFF;
        }
        @media (max-width: 640px) {
            .encabezado-app {
                align-items: flex-start;
                min-height: 0;
                padding: 20px;
            }
            .encabezado-app img {
                height: 68px;
                width: 68px;
            }
            .encabezado-app h1 {
                font-size: 1.65rem;
            }
            .encabezado-app p {
                font-size: 0.95rem;
            }
        }
        </style>
        <div class="encabezado-app">
            <img src="data:image/png;base64,{IMAGEN_TITULO}" alt="Logo CENS">
            <div class="encabezado-app-contenido">
                <h1>Proyección Comercial: Beneficios de la Comunidad Energética</h1>
                <p>Simulación dinámica semestral y análisis financiero de ahorros.</p>
            </div>
        </div>
        """.replace("{IMAGEN_TITULO}", IMAGEN_TITULO),
        unsafe_allow_html=True,
    )
st.markdown("---")

if "mostrar_desglose" not in st.session_state:
    st.session_state.mostrar_desglose = False

columnas_mensuales = [
    "Mes", "Consumo_kWh", "Tarifa_Aplicada_COP_kWh",
    "Tarifa_CE_COP_kWh", "Cobertura_CE_pct", "Alumbrado_Publico_pct"
]
datos_mensuales_predeterminados = pd.DataFrame({
    "Mes": [f"Mes {i}" for i in range(1, 7)],
    "Consumo_kWh": [7000, 7200, 6900, 7100, 7300, 7000],
    "Tarifa_Aplicada_COP_kWh": [1043.93] * 6,
    "Tarifa_CE_COP_kWh": [913.36] * 6,
    "Cobertura_CE_pct": [80.0] * 6,
    "Alumbrado_Publico_pct": [13.0] * 6,
})

if "datos_mensuales" not in st.session_state:
    st.session_state.datos_mensuales = datos_mensuales_predeterminados.copy()
if "archivo_excel_cargado" not in st.session_state:
    st.session_state.archivo_excel_cargado = None
if "mostrar_editor_manual" not in st.session_state:
    st.session_state.mostrar_editor_manual = False

def toggle_desglose():
    st.session_state.mostrar_desglose = not st.session_state.mostrar_desglose

def toggle_editor_manual():
    st.session_state.mostrar_editor_manual = not st.session_state.mostrar_editor_manual

# 2. Panel lateral para el ingreso de datos
with st.sidebar:
    st.header("Datos mensuales")
    st.caption("La tarifa aplicada debe incluir la contribución. El alumbrado se ingresa aparte.")
    archivo_excel = st.file_uploader("Cargar Excel mensual", type=["xlsx", "xls"])

    if archivo_excel is not None:
        identificador_archivo = f"{archivo_excel.name}:{archivo_excel.size}"
        if identificador_archivo != st.session_state.archivo_excel_cargado:
            try:
                datos_excel = pd.read_excel(archivo_excel)
                columnas_faltantes = [columna for columna in columnas_mensuales if columna not in datos_excel.columns]
                if columnas_faltantes:
                    st.error(f"Faltan columnas: {', '.join(columnas_faltantes)}")
                elif len(datos_excel) != 6:
                    st.error("El Excel debe contener exactamente seis filas, una por cada mes.")
                else:
                    datos_excel = datos_excel[columnas_mensuales].copy()
                    columnas_numericas = columnas_mensuales[1:]
                    for columna in columnas_numericas:
                        datos_excel[columna] = pd.to_numeric(datos_excel[columna], errors="coerce")
                    if datos_excel[columnas_numericas].isna().any().any():
                        st.error("Las columnas numericas del Excel no pueden tener valores vacios o no numericos.")
                    elif (datos_excel[columnas_numericas] < 0).any().any():
                        st.error("Las cantidades y tarifas del Excel no pueden ser negativas.")
                    else:
                        st.session_state.datos_mensuales = datos_excel
                        st.session_state.archivo_excel_cargado = identificador_archivo
                        st.success("Excel cargado correctamente.")
            except Exception as error:
                st.error(f"No fue posible leer el Excel: {error}")

    plantilla_excel = BytesIO()
    with pd.ExcelWriter(plantilla_excel, engine="openpyxl") as escritor:
        datos_mensuales_predeterminados.to_excel(escritor, index=False, sheet_name="Datos mensuales")
    st.download_button(
        "Descargar plantilla Excel",
        data=plantilla_excel.getvalue(),
        file_name="plantilla_datos_mensuales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.button(
        "Ingresar valores manualmente" if not st.session_state.mostrar_editor_manual else "Ocultar ingreso manual",
        on_click=toggle_editor_manual,
        use_container_width=True,
        type="primary",
    )

    if st.session_state.mostrar_editor_manual:
        st.caption("Edite los valores de los seis meses directamente en la tabla.")
        editor_key = f"editor_mensual_{st.session_state.archivo_excel_cargado or 'manual'}"
        datos_editados = st.data_editor(
            st.session_state.datos_mensuales,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "Consumo_kWh": st.column_config.NumberColumn("Consumo (kWh)", min_value=0.0),
                "Tarifa_Aplicada_COP_kWh": st.column_config.NumberColumn("Tarifa aplicada (COP/kWh)", min_value=0.0),
                "Tarifa_CE_COP_kWh": st.column_config.NumberColumn("Tarifa CE (COP/kWh)", min_value=0.0),
                "Cobertura_CE_pct": st.column_config.NumberColumn("Cobertura CE (%)", min_value=0.0, max_value=100.0),
                "Alumbrado_Publico_pct": st.column_config.NumberColumn("Alumbrado publico (%)", min_value=0.0, max_value=100.0),
            },
            key=editor_key,
        )
        columnas_numericas = columnas_mensuales[1:]
        for columna in columnas_numericas:
            datos_editados[columna] = pd.to_numeric(datos_editados[columna], errors="coerce")
        if datos_editados[columnas_numericas].isna().any().any() or (datos_editados[columnas_numericas] < 0).any().any():
            st.error("Complete las cantidades y tarifas con valores numericos no negativos.")
            datos_editados = st.session_state.datos_mensuales.copy()
        st.session_state.datos_mensuales = datos_editados

# 3. Motor de Calculo Dinamico
def calcular_mes(consumo, tarifa_aplicada, tarifa_ce, pct_cobertura, pct_alumbrado):
    consumo_activa = consumo * tarifa_aplicada
    alumbrado = consumo_activa * pct_alumbrado
    total_actual = consumo_activa + alumbrado
    
    energia_asignada = consumo * pct_cobertura
    compra_energia_asignada = - (energia_asignada * tarifa_aplicada)
    cobro_energia_ce = energia_asignada * tarifa_ce
    alumbrado_ce = alumbrado 

    total_ce = consumo_activa + compra_energia_asignada + cobro_energia_ce + alumbrado_ce
    ahorro = total_actual - total_ce
    
    return {
        "Consumo": consumo,
        "Tarifa Aplicada": tarifa_aplicada,
        "Tarifa CE": tarifa_ce,
        "Total Actual": total_actual,
        "Total CE": total_ce,
        "Ahorro": ahorro,
        "Detalle Actual": [consumo_activa, 0.0, 0.0, 0.0, alumbrado, total_actual],
        "Detalle CE": [consumo_activa, compra_energia_asignada, cobro_energia_ce, 0.0, alumbrado_ce, total_ce]
    }

# Ejecución mensual
resultados = [
    calcular_mes(
        fila["Consumo_kWh"], fila["Tarifa_Aplicada_COP_kWh"], fila["Tarifa_CE_COP_kWh"],
        fila["Cobertura_CE_pct"] / 100.0, fila["Alumbrado_Publico_pct"] / 100.0
    )
    for _, fila in st.session_state.datos_mensuales.iterrows()
]
meses_labels = st.session_state.datos_mensuales["Mes"].astype(str).tolist()

facturas_sin_ce = [r["Total Actual"] for r in resultados]
facturas_con_ce = [r["Total CE"] for r in resultados]
ahorros_mensuales = [r["Ahorro"] for r in resultados]

total_semestre_sin_ce = sum(facturas_sin_ce)
total_semestre_con_ce = sum(facturas_con_ce)
total_ahorro_semestre = sum(ahorros_mensuales)
valor_tarifa_aplicada = sum(r["Consumo"] * r["Tarifa Aplicada"] for r in resultados)
valor_tarifa_ce = sum(r["Consumo"] * r["Tarifa CE"] for r in resultados)
ahorro_tarifa_porcentual = ((valor_tarifa_aplicada - valor_tarifa_ce) / valor_tarifa_aplicada) * 100 if valor_tarifa_aplicada > 0 else 0
ahorro_factura_porcentual = (total_ahorro_semestre / total_semestre_sin_ce) * 100 if total_semestre_sin_ce > 0 else 0

# 4. Panel de Impacto Financiero
st.subheader("Resumen de Impacto Semestral")
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric(label="Facturación Tradicional Proyectada", value=f"${total_semestre_sin_ce:,.0f}")
with col2:
    st.metric(label="Facturación con Comunidad Energética", value=f"${total_semestre_con_ce:,.0f}", delta=f"-${total_ahorro_semestre:,.0f}", delta_color="inverse")
with col3:
    st.metric(label="Ahorro sobre Tarifa", value=f"{ahorro_tarifa_porcentual:.1f}%")
with col4:
    st.metric(label="Ahorro sobre Factura", value=f"{ahorro_factura_porcentual:.1f}%")

resultado_mes_1 = resultados[0]
st.subheader(f"Resumen de Impacto Mensual - {meses_labels[0]}")
mes1_col1, mes1_col2, mes1_col3, mes1_col4 = st.columns(4)
with mes1_col1:
    st.metric(label="Facturación Tradicional", value=f"${resultado_mes_1['Total Actual']:,.0f}")
with mes1_col2:
    st.metric(label="Facturación con CE", value=f"${resultado_mes_1['Total CE']:,.0f}")
with mes1_col3:
    st.metric(label="Ahorro del Mes", value=f"${resultado_mes_1['Ahorro']:,.0f}")
with mes1_col4:
    ahorro_mes_1_pct = (resultado_mes_1["Ahorro"] / resultado_mes_1["Total Actual"]) * 100 if resultado_mes_1["Total Actual"] > 0 else 0
    st.metric(label="Ahorro sobre Factura", value=f"{ahorro_mes_1_pct:.1f}%")

st.markdown("---")

# 5. Cuadros de Resumen Textual por Mes
st.subheader("Historial de Consumo y Ahorro Mensual")
cols = st.columns(3)
for i, (res, label) in enumerate(zip(resultados, meses_labels)):
    with cols[i % 3]:
        st.success(
            f"**{label}**\n\n"
            f"Consumo: **{res['Consumo']:,.0f} kWh**\n\n"
            f"Tarifa aplicada (con contribución): **${res['Tarifa Aplicada']:,.2f}/kWh**\n\n"
            f"Tarifa CE: **${res['Tarifa CE']:,.2f}/kWh**\n\n"
            f"Facturación Tradicional: **${res['Total Actual']:,.0f}**\n\n"
            f"Facturación en CE: **${res['Total CE']:,.0f}**\n\n"
            f"Ahorro del mes: **${res['Ahorro']:,.0f}**"
        )

st.markdown("---")

# 6. Representacion Grafica
st.subheader("Análisis Comparativo del Periodo")
col_graf1, col_graf2 = st.columns(2)

with col_graf1:
    fig_lineas = go.Figure()
    fig_lineas.add_trace(go.Scatter(x=meses_labels, y=facturas_sin_ce, name="Sin CE", line=dict(color=COLOR_ACTUAL, width=3, dash='dot'), mode='lines+markers'))
    fig_lineas.add_trace(go.Scatter(x=meses_labels, y=facturas_con_ce, name="Con CE", line=dict(color=COLOR_CE, width=4), mode='lines+markers', fill='tonexty', fillcolor='rgba(39, 174, 96, 0.2)'))
    fig_lineas.update_layout(title="Tendencia de Facturación", xaxis_title="Periodo", yaxis_title="Valor (COP)", hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig_lineas, use_container_width=True)

with col_graf2:
    df_barras = pd.DataFrame({"Mes": meses_labels * 2, "Escenario": ["Tradicional"] * 6 + ["Comunidad Energética"] * 6, "Costo ($)": facturas_sin_ce + facturas_con_ce})
    fig_barras = px.bar(df_barras, x="Mes", y="Costo ($)", color="Escenario", barmode="group", color_discrete_map={"Tradicional": COLOR_ACTUAL, "Comunidad Energética": COLOR_CE}, text_auto='.2s', title="Comparación Mensual Directa")
    fig_barras.update_layout(xaxis_title="Periodo", yaxis_title="Valor (COP)", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig_barras, use_container_width=True)

st.markdown("---")

# 7. Desglose y Generación de PDF
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    st.button("Ver desglose de facturación", on_click=toggle_desglose, type="primary")

def generar_pdf(datos_mensuales, reduccion_tarifa_pdf, ahorro_total_pdf):
    pdf = FPDF()
    pdf.add_page()
    
    # --- Encabezado Corporativo ---
    pdf.set_fill_color(39, 174, 96) 
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(190, 15, "Reporte Comercial: Comunidad Energetica", ln=True, align='C', fill=True)
    pdf.ln(10)
    
    # --- Resumen Tarifario (Enfoque Mensual) ---
    promedio_ahorro_mensual = ahorro_total_pdf / len(datos_mensuales)
    consumo_total_pdf = sum(res["Consumo"] for res in datos_mensuales)
    tarifa_aplicada_pdf = sum(res["Consumo"] * res["Tarifa Aplicada"] for res in datos_mensuales) / consumo_total_pdf if consumo_total_pdf else 0
    tarifa_ce_pdf = sum(res["Consumo"] * res["Tarifa CE"] for res in datos_mensuales) / consumo_total_pdf if consumo_total_pdf else 0
    # Extraemos específicamente el ahorro del primer mes
    ahorro_mes_1 = datos_mensuales[0]["Ahorro"] 
    
    pdf.set_text_color(44, 62, 80)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(190, 10, "Resumen Tarifario y Proyeccion Mensual", ln=True)
    
    pdf.set_font("Arial", '', 12)
    pdf.cell(190, 8, f"Tarifa Aplicada Promedio (con contribucion): ${tarifa_aplicada_pdf:,.2f} COP/kWh", ln=True)
    pdf.cell(190, 8, f"Tarifa Comunidad Energetica: ${tarifa_ce_pdf:,.2f} COP/kWh", ln=True)
    
    pdf.set_font("Arial", 'B', 12)
    pdf.set_text_color(39, 174, 96)
    pdf.cell(190, 8, f"Reduccion de la Tarifa: {reduccion_tarifa_pdf:.1f}%", ln=True)
    # Mostramos el ahorro del primer mes y el promedio de los 6 meses
    pdf.cell(190, 8, f"Ahorro Estimado (Mes 1): ${ahorro_mes_1:,.0f} COP", ln=True)
    pdf.cell(190, 8, f"Ahorro Promedio Mensual Estimado: ${promedio_ahorro_mensual:,.0f} COP", ln=True)
    pdf.ln(10)
    
    # --- Tabla de Desglose Mensual ---
    pdf.set_text_color(44, 62, 80)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(190, 10, "Desglose Mensual de Consumo y Ahorro", ln=True)
    
    pdf.set_fill_color(230, 240, 230)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_font("Arial", 'B', 7)
    pdf.cell(18, 10, "Mes", border=1, align='C', fill=True)
    pdf.cell(25, 10, "Consumo", border=1, align='C', fill=True)
    pdf.cell(35, 10, "Tarifa aplicada", border=1, align='C', fill=True)
    pdf.cell(30, 10, "Tarifa CE", border=1, align='C', fill=True)
    pdf.cell(27, 10, "Tradicional", border=1, align='C', fill=True)
    pdf.cell(27, 10, "Comunidad", border=1, align='C', fill=True)
    pdf.cell(28, 10, "Ahorro", border=1, align='C', fill=True)
    pdf.ln()
    
    pdf.set_font("Arial", '', 10)
    for i, res in enumerate(datos_mensuales):
        pdf.set_text_color(0, 0, 0)
        pdf.cell(18, 10, meses_labels[i], border=1, align='C')
        pdf.cell(25, 10, f"{res['Consumo']:,.0f}", border=1, align='C')
        pdf.cell(35, 10, f"${res['Tarifa Aplicada']:,.2f}", border=1, align='C')
        pdf.cell(30, 10, f"${res['Tarifa CE']:,.2f}", border=1, align='C')
        pdf.cell(27, 10, f"${res['Total Actual']:,.0f}", border=1, align='C')
        pdf.cell(27, 10, f"${res['Total CE']:,.0f}", border=1, align='C')
        
        pdf.set_text_color(39, 174, 96)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(28, 10, f"${res['Ahorro']:,.0f}", border=1, align='C')
        pdf.set_font("Arial", '', 10)
        pdf.ln()
        
    # --- Pie de página ---
    pdf.ln(15)
    pdf.set_text_color(127, 140, 141)
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(190, 5, "Documento generado automaticamente.", ln=True, align='C')
    pdf.cell(190, 5, "Esta proyeccion es una simulacion comercial; los valores finales pueden variar segun la regulacion tarifaria vigente.", ln=True, align='C')
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf.output(tmp.name)
        with open(tmp.name, "rb") as f:
            datos_pdf = f.read()
    os.remove(tmp.name)
    return datos_pdf

with col_btn2:
    st.download_button(
        label="Descargar Reporte en PDF",
        data=generar_pdf(resultados, ahorro_tarifa_porcentual, total_ahorro_semestre),
        file_name="Reporte_Comunidad_Energetica.pdf",
        mime="application/pdf",
        type="primary"
    )

if st.session_state.mostrar_desglose:
    st.subheader("Análisis Detallado por Concepto (Estructura de Modelo)")
    conceptos = ["Consumo Energía Activa", "Compra Energía Asignada CE (-)", "Cobro Energía Comunidad Energética", "Contribución (incluida en tarifa)", "Alumbrado Público", "TOTAL FACTURA"]
    def formato_moneda(lista): return [f"${val:,.0f}" for val in lista]
    
    df_desglose = pd.DataFrame({
        "Concepto Facturado": conceptos,
        "Mes 1 (Sin CE)": formato_moneda(resultados[0]["Detalle Actual"]), "Mes 1 (Con CE)": formato_moneda(resultados[0]["Detalle CE"]),
        "Mes 2 (Sin CE)": formato_moneda(resultados[1]["Detalle Actual"]), "Mes 2 (Con CE)": formato_moneda(resultados[1]["Detalle CE"]),
        "Mes 3 (Sin CE)": formato_moneda(resultados[2]["Detalle Actual"]), "Mes 3 (Con CE)": formato_moneda(resultados[2]["Detalle CE"]),
        "Mes 4 (Sin CE)": formato_moneda(resultados[3]["Detalle Actual"]), "Mes 4 (Con CE)": formato_moneda(resultados[3]["Detalle CE"]),
        "Mes 5 (Sin CE)": formato_moneda(resultados[4]["Detalle Actual"]), "Mes 5 (Con CE)": formato_moneda(resultados[4]["Detalle CE"]),
        "Mes 6 (Sin CE)": formato_moneda(resultados[5]["Detalle Actual"]), "Mes 6 (Con CE)": formato_moneda(resultados[5]["Detalle CE"]),
    })
    st.dataframe(df_desglose, use_container_width=True, hide_index=True)