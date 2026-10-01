import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import tempfile
import os
import base64
from io import BytesIO
from PIL import Image

# 1. Configuracion institucional
st.set_page_config(page_title="Proyección Comercial CE", layout="wide")

COLOR_ACTUAL = "#7F8C8D" 
COLOR_CE = "#27AE60"     

# Imagen institucional definida en el codigo, no cargada por el usuario.
RUTA_IMAGEN_TITULO = os.path.join(os.path.dirname(__file__), "LogoCENS.png")
with open(RUTA_IMAGEN_TITULO, "rb") as archivo_imagen:
    IMAGEN_TITULO = base64.b64encode(archivo_imagen.read()).decode("ascii")
RUTA_IMAGEN_SOLAR = os.path.join(os.path.dirname(__file__), "Planta Solar.jpg")

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
    "Tarifa_CE_COP_kWh", "Cobertura_CE_pct", "Alumbrado_Publico_pct",
    "Paga contribucion",
]
datos_mensuales_predeterminados = pd.DataFrame({
    "Mes": [f"Mes {indice}" for indice in range(1, 7)],
    "Consumo_kWh": [7000, 7200, 6900, 7100, 7300, 7000],
    "Tarifa_Aplicada_COP_kWh": [1043.93] * 6,
    "Tarifa_CE_COP_kWh": [913.36] * 6,
    "Cobertura_CE_pct": [80.0] * 6,
    "Alumbrado_Publico_pct": [13.0] * 6,
    "Paga contribucion": ["NO"] * 6,
})
if "datos_mensuales" not in st.session_state:
    st.session_state.datos_mensuales = datos_mensuales_predeterminados.copy()
if "archivo_excel_cargado" not in st.session_state:
    st.session_state.archivo_excel_cargado = None

def toggle_desglose():
    st.session_state.mostrar_desglose = not st.session_state.mostrar_desglose

if "mostrar_editor_manual" not in st.session_state:
    st.session_state.mostrar_editor_manual = False

def toggle_editor_manual():
    st.session_state.mostrar_editor_manual = not st.session_state.mostrar_editor_manual

# 2. Panel lateral para el ingreso de datos
with st.sidebar:
    st.header("Datos mensuales")
    st.caption("La tarifa aplicada debe incluir la contribución. El alumbrado se ingresa aparte.")
    nombre_cliente = st.text_input("Nombre del Cliente / Razón Social")
    dias_retiro = st.number_input("Días de preaviso para retiro", min_value=0, value=90, step=1)
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
                    datos_excel["Paga contribucion"] = datos_excel["Paga contribucion"].astype(str).str.strip().str.upper()
                    columnas_numericas = columnas_mensuales[1:]
                    columnas_numericas.remove("Paga contribucion")
                    for columna in columnas_numericas:
                        datos_excel[columna] = pd.to_numeric(datos_excel[columna], errors="coerce")
                    if not datos_excel["Paga contribucion"].isin({"SI", "NO"}).all():
                        st.error('La columna "Paga contribucion" solo puede contener SI o NO.')
                    elif datos_excel[columnas_numericas].isna().any().any():
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
        "Descargar plantilla Excel", data=plantilla_excel.getvalue(),
        file_name="plantilla_datos_mensuales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.button(
        "Ingresar valores manualmente" if not st.session_state.mostrar_editor_manual else "Ocultar ingreso manual",
        on_click=toggle_editor_manual, width="stretch", type="primary",
    )

    if st.session_state.mostrar_editor_manual:
        st.caption("Edite los valores de los seis meses directamente en la tabla.")
        editor_key = f"editor_mensual_{st.session_state.archivo_excel_cargado or 'manual'}"
        datos_editados = st.data_editor(
            st.session_state.datos_mensuales, hide_index=True, width="stretch",
            num_rows="fixed",
            column_config={
                "Consumo_kWh": st.column_config.NumberColumn("Consumo (kWh)", min_value=0.0),
                "Tarifa_Aplicada_COP_kWh": st.column_config.NumberColumn("Tarifa aplicada (COP/kWh)", min_value=0.0),
                "Tarifa_CE_COP_kWh": st.column_config.NumberColumn("Tarifa CE (COP/kWh)", min_value=0.0),
                "Cobertura_CE_pct": st.column_config.NumberColumn("Cobertura CE (%)", min_value=0.0, max_value=100.0),
                "Alumbrado_Publico_pct": st.column_config.NumberColumn("Alumbrado publico (%)", min_value=0.0, max_value=100.0),
                "Paga contribucion": st.column_config.SelectboxColumn("Paga contribucion", options=["SI", "NO"], required=True),
            },
            key=editor_key,
        )
        columnas_numericas = columnas_mensuales[1:]
        columnas_numericas.remove("Paga contribucion")
        for columna in columnas_numericas:
            datos_editados[columna] = pd.to_numeric(datos_editados[columna], errors="coerce")
        datos_editados["Paga contribucion"] = datos_editados["Paga contribucion"].astype(str).str.strip().str.upper()
        if (
            datos_editados[columnas_numericas].isna().any().any()
            or (datos_editados[columnas_numericas] < 0).any().any()
            or not datos_editados["Paga contribucion"].isin({"SI", "NO"}).all()
        ):
            st.error("Complete las cantidades y tarifas con valores numericos no negativos.")
            datos_editados = st.session_state.datos_mensuales.copy()
        st.session_state.datos_mensuales = datos_editados

# 3. Motor de Calculo Dinamico
def calcular_mes(consumo, tarifa_aplicada, tarifa_ce, pct_cobertura, pct_alumbrado, paga_contribucion):
    contribuye = paga_contribucion == "SI"
    consumo_activa = consumo * tarifa_aplicada / 1.2 if contribuye else consumo * tarifa_aplicada
    contribucion = consumo_activa * 0.2 if contribuye else 0.0
    alumbrado = consumo_activa * pct_alumbrado
    total_actual = consumo_activa + contribucion + alumbrado
    energia_asignada = consumo * pct_cobertura
    compra_energia_asignada = -(energia_asignada * tarifa_aplicada)
    cobro_energia_ce = consumo * pct_cobertura * tarifa_ce if contribuye else energia_asignada * tarifa_ce
    consumo_activa_ce = consumo * tarifa_aplicada / 1.2 if contribuye else consumo_activa
    compra_energia_ce = consumo * pct_cobertura * -(tarifa_aplicada / 1.2) if contribuye else 0.0
    contribucion_ce = (consumo_activa_ce + compra_energia_ce) * 0.2 if contribuye else 0.0
    alumbrado_ce = consumo_activa_ce * pct_alumbrado
    total_ce = (
        consumo_activa_ce + compra_energia_ce + cobro_energia_ce + contribucion_ce + alumbrado_ce
        if contribuye else consumo_activa + compra_energia_asignada + cobro_energia_ce + alumbrado_ce
    )
    ahorro = total_actual - total_ce
    return {
        "Consumo": consumo, "Tarifa Aplicada": tarifa_aplicada, "Tarifa CE": tarifa_ce,
        "Cobertura CE": pct_cobertura * 100, "Paga contribucion": paga_contribucion,
        "Total Actual": total_actual, "Total CE": total_ce, "Ahorro": ahorro,
        "Detalle Actual": [consumo_activa, 0.0, 0.0, contribucion, alumbrado, total_actual],
        "Detalle CE": [
            consumo_activa_ce, compra_energia_asignada if not contribuye else compra_energia_ce,
            cobro_energia_ce, contribucion_ce, alumbrado_ce, total_ce,
        ],
    }

resultados = [
    calcular_mes(
        fila["Consumo_kWh"], fila["Tarifa_Aplicada_COP_kWh"], fila["Tarifa_CE_COP_kWh"],
        fila["Cobertura_CE_pct"] / 100.0, fila["Alumbrado_Publico_pct"] / 100.0,
        fila["Paga contribucion"],
    )
    for _, fila in st.session_state.datos_mensuales.iterrows()
]
meses_labels = st.session_state.datos_mensuales["Mes"].astype(str).tolist()
facturas_sin_ce = [resultado["Total Actual"] for resultado in resultados]
facturas_con_ce = [resultado["Total CE"] for resultado in resultados]
ahorros_mensuales = [resultado["Ahorro"] for resultado in resultados]

total_semestre_sin_ce = sum(facturas_sin_ce)
total_semestre_con_ce = sum(facturas_con_ce)
total_ahorro_semestre = sum(ahorros_mensuales)
valor_tarifa_aplicada = sum(resultado["Consumo"] * resultado["Tarifa Aplicada"] for resultado in resultados)
valor_tarifa_ce = sum(resultado["Consumo"] * resultado["Tarifa CE"] for resultado in resultados)
ahorro_tarifa_porcentual = (
    (valor_tarifa_aplicada - valor_tarifa_ce) / valor_tarifa_aplicada * 100
    if valor_tarifa_aplicada > 0 else 0
)
ahorro_factura_porcentual = (
    total_ahorro_semestre / total_semestre_sin_ce * 100
    if total_semestre_sin_ce > 0 else 0
)

st.subheader("Resumen de Impacto Semestral")
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Facturación Tradicional Proyectada", f"${total_semestre_sin_ce:,.0f}")
with col2:
    st.metric(
        "Facturación con Comunidad Energética", f"${total_semestre_con_ce:,.0f}",
        delta=f"-${total_ahorro_semestre:,.0f}", delta_color="inverse",
    )
with col3:
    st.metric("Ahorro sobre Tarifa", f"{ahorro_tarifa_porcentual:.1f}%")
with col4:
    st.metric("Ahorro sobre Factura", f"{ahorro_factura_porcentual:.1f}%")

resultado_mes_1 = resultados[0]
st.subheader(f"Resumen de Impacto Mensual - {meses_labels[0]}")
mes1_col1, mes1_col2, mes1_col3, mes1_col4 = st.columns(4)
with mes1_col1:
    st.metric("Facturación Tradicional", f"${resultado_mes_1['Total Actual']:,.0f}")
with mes1_col2:
    st.metric("Facturación con CE", f"${resultado_mes_1['Total CE']:,.0f}")
with mes1_col3:
    st.metric("Ahorro del Mes", f"${resultado_mes_1['Ahorro']:,.0f}")
with mes1_col4:
    ahorro_mes_1_pct = (
        resultado_mes_1["Ahorro"] / resultado_mes_1["Total Actual"] * 100
        if resultado_mes_1["Total Actual"] > 0 else 0
    )
    st.metric("Ahorro sobre Factura", f"{ahorro_mes_1_pct:.1f}%")

st.markdown("---")
st.subheader("Historial de Consumo y Ahorro Mensual")
cols = st.columns(3)
for indice, (resultado, etiqueta) in enumerate(zip(resultados, meses_labels)):
    with cols[indice % 3]:
        st.success(
            f"**{etiqueta}**\n\n"
            f"Consumo: **{resultado['Consumo']:,.0f} kWh**\n\n"
            f"Tarifa aplicada: **${resultado['Tarifa Aplicada']:,.2f}/kWh**\n\n"
            f"Tarifa CE: **${resultado['Tarifa CE']:,.2f}/kWh**\n\n"
            f"Facturación Tradicional: **${resultado['Total Actual']:,.0f}**\n\n"
            f"Facturación en CE: **${resultado['Total CE']:,.0f}**\n\n"
            f"Ahorro del mes: **${resultado['Ahorro']:,.0f}**"
        )

st.markdown("---")
st.subheader("Análisis Comparativo del Periodo")
col_graf1, col_graf2 = st.columns(2)
with col_graf1:
    fig_lineas = go.Figure()
    fig_lineas.add_trace(go.Scatter(
        x=meses_labels, y=facturas_sin_ce, name="Sin CE",
        line=dict(color=COLOR_ACTUAL, width=3, dash="dot"), mode="lines+markers",
    ))
    fig_lineas.add_trace(go.Scatter(
        x=meses_labels, y=facturas_con_ce, name="Con CE",
        line=dict(color=COLOR_CE, width=4), mode="lines+markers",
        fill="tonexty", fillcolor="rgba(39, 174, 96, 0.2)",
    ))
    fig_lineas.update_layout(
        title="Tendencia de Facturación", xaxis_title="Periodo",
        yaxis_title="Valor (COP)", hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig_lineas, width="stretch")

with col_graf2:
    df_barras = pd.DataFrame({
        "Mes": meses_labels * 2,
        "Escenario": ["Tradicional"] * len(resultados) + ["Comunidad Energética"] * len(resultados),
        "Costo ($)": facturas_sin_ce + facturas_con_ce,
    })
    fig_barras = px.bar(
        df_barras, x="Mes", y="Costo ($)", color="Escenario", barmode="group",
        color_discrete_map={"Tradicional": COLOR_ACTUAL, "Comunidad Energética": COLOR_CE},
        text_auto=".2s", title="Comparación Mensual Directa",
    )
    fig_barras.update_layout(
        xaxis_title="Periodo", yaxis_title="Valor (COP)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig_barras, width="stretch")

st.markdown("---")
col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    st.button("Ver desglose de facturación", on_click=toggle_desglose, type="primary")

@st.cache_data(show_spinner=False)
def generar_pdf(datos_mensuales, meses, nombre_cliente, dias_retiro):
    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=False)
    cliente = nombre_cliente.strip() or "Cliente por definir"
    total_tradicional = sum(resultado["Total Actual"] for resultado in datos_mensuales)
    total_ce = sum(resultado["Total CE"] for resultado in datos_mensuales)
    ahorro_total = sum(resultado["Ahorro"] for resultado in datos_mensuales)
    consumo_total = sum(resultado["Consumo"] for resultado in datos_mensuales)
    cobertura_media = (
        sum(resultado["Consumo"] * resultado["Cobertura CE"] for resultado in datos_mensuales) / consumo_total
        if consumo_total else 0
    )
    temporal_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    ruta_logo = temporal_logo.name
    temporal_logo.close()
    temporal_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    ruta_pdf = temporal_pdf.name
    temporal_pdf.close()

    def encabezado(titulo, subtitulo):
        pdf.set_fill_color(20, 90, 50)
        pdf.rect(0, 0, 210, 48, style="F")
        pdf.set_fill_color(39, 174, 96)
        pdf.rect(183, 0, 27, 48, style="F")
        pdf.set_fill_color(232, 245, 233)
        pdf.rect(12, 8, 31, 31, style="F")
        with Image.open(RUTA_IMAGEN_TITULO) as logo:
            logo_rgba = logo.convert("RGBA")
            fondo_logo = Image.new("RGB", logo_rgba.size, (255, 255, 255))
            fondo_logo.paste(logo_rgba, mask=logo_rgba.getchannel("A"))
            fondo_logo.save(ruta_logo, format="PNG")
        pdf.image(ruta_logo, x=13, y=9, w=29, h=29)
        pdf.set_text_color(204, 238, 216)
        pdf.set_xy(50, 7)
        pdf.set_font("Arial", "B", 8)
        pdf.cell(126, 5, "ENERGÍA CENS  /  PROPUESTA COMERCIAL")
        pdf.set_text_color(255, 255, 255)
        pdf.set_xy(50, 14)
        pdf.set_font("Arial", "B", 16)
        pdf.cell(126, 9, titulo)
        pdf.set_xy(50, 27)
        pdf.set_font("Arial", "", 8)
        pdf.multi_cell(126, 5, subtitulo)

    def titulo_seccion(y, titulo, detalle=""):
        pdf.set_xy(12, y)
        pdf.set_text_color(20, 90, 50)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(186, 6, titulo)
        if detalle:
            pdf.set_xy(12, y + 6)
            pdf.set_text_color(100, 115, 106)
            pdf.set_font("Arial", "", 7)
            pdf.cell(186, 4, detalle)

    def bloque_condicion(x, y, numero, titulo, texto):
        pdf.set_fill_color(232, 245, 233)
        pdf.rect(x, y, 90, 58, style="F")
        pdf.set_fill_color(39, 174, 96)
        pdf.ellipse(x + 4, y + 4, 8, 8, style="F")
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 8)
        pdf.set_xy(x + 4, y + 5)
        pdf.cell(8, 5, str(numero), align="C")
        pdf.set_text_color(20, 90, 50)
        pdf.set_font("Arial", "B", 9)
        pdf.set_xy(x + 15, y + 4)
        pdf.cell(71, 6, titulo)
        pdf.set_text_color(48, 65, 54)
        pdf.set_font("Arial", "", 7.5)
        pdf.set_xy(x + 4, y + 15)
        pdf.multi_cell(82, 4.2, texto)

    try:
        pdf.add_page()
        encabezado("COMUNIDAD ENERGÉTICA", f"Propuesta personalizada para: {cliente}")

        tarjetas = [
            ("COSTO SIN COMUNIDAD", total_tradicional, "$"),
            ("COSTO CON COMUNIDAD", total_ce, "CE"),
            ("AHORRO PROYECTADO", ahorro_total, "%"),
        ]
        for indice, (etiqueta, valor, icono) in enumerate(tarjetas):
            x = 12 + indice * 63
            pdf.set_fill_color(232, 245, 233)
            pdf.rect(x, 54, 58, 31, style="F")
            pdf.set_draw_color(39, 174, 96)
            pdf.ellipse(x + 5, 58, 11, 11)
            pdf.set_text_color(20, 90, 50)
            pdf.set_font("Arial", "B", 7)
            pdf.set_xy(x + 5, 61)
            pdf.cell(11, 5, icono, align="C")
            pdf.set_xy(x + 19, 58)
            pdf.cell(36, 5, etiqueta)
            if indice == 2:
                pdf.set_text_color(39, 174, 96)
            else:
                pdf.set_text_color(20, 90, 50)
            pdf.set_font("Arial", "B", 10 if indice == 2 else 9)
            pdf.set_xy(x + 5, 72)
            pdf.cell(49, 7, f"${valor:,.0f}", align="C")
            pdf.set_text_color(90, 105, 96)
            pdf.set_font("Arial", "", 6)
            pdf.set_xy(x + 5, 79)
            pdf.cell(49, 4, "COP  |  PROYECCIÓN SEMESTRAL", align="C")

        titulo_seccion(91, "Resumen comercial de la propuesta", "Proyección semestral estimada para el consumo informado por el cliente")
        pdf.set_xy(12, 102)
        pdf.set_fill_color(244, 248, 245)
        pdf.rect(12, 101, 186, 15, style="F")
        pdf.set_text_color(48, 65, 54)
        pdf.set_font("Arial", "", 7)
        pdf.set_xy(16, 103)
        pdf.multi_cell(
            178, 4,
            f"La propuesta vincula a {cliente} a una comunidad de suministro con energía renovable asignada según la generación disponible. El modelo proyecta un diferencial esperado del 25% en el precio de la energía asignada, una cobertura mínima de referencia del 80% y un ahorro estimado de ${ahorro_total:,.0f} COP durante seis meses, sujeto a las condiciones técnicas y contractuales.",
        )
        maximo = max((max(r["Total Actual"], r["Total CE"]) for r in datos_mensuales), default=0) or 1
        eje_y = 155
        pdf.set_draw_color(198, 211, 202)
        pdf.line(22, eje_y, 198, eje_y)
        pdf.set_fill_color(20, 90, 50)
        pdf.rect(145, 120, 4, 3, style="F")
        pdf.set_xy(151, 119)
        pdf.set_text_color(70, 80, 74)
        pdf.set_font("Arial", "", 6)
        pdf.cell(22, 4, "Tradicional")
        pdf.set_fill_color(39, 174, 96)
        pdf.rect(175, 120, 4, 3, style="F")
        pdf.set_xy(181, 119)
        pdf.cell(14, 4, "CE")
        for indice, resultado in enumerate(datos_mensuales):
            centro = 29 + indice * 29
            alto_actual = 21 * resultado["Total Actual"] / maximo
            alto_ce = 21 * resultado["Total CE"] / maximo
            pdf.set_fill_color(20, 90, 50)
            pdf.rect(centro, eje_y - alto_actual, 8, alto_actual, style="F")
            pdf.set_fill_color(39, 174, 96)
            pdf.rect(centro + 10, eje_y - alto_ce, 8, alto_ce, style="F")
            pdf.set_text_color(48, 65, 54)
            pdf.set_font("Arial", "B", 4)
            pdf.set_xy(centro - 6, eje_y - alto_actual - 2.5)
            pdf.cell(20, 3, f"${resultado['Total Actual']:,.0f}", align="C")
            pdf.set_xy(centro + 4, eje_y - alto_ce - 2.5)
            pdf.cell(20, 3, f"${resultado['Total CE']:,.0f}", align="C")

            ahorro_es_positivo = resultado["Ahorro"] >= 0
            x_barra_menor = centro + (10 if ahorro_es_positivo else 0)
            alto_barra_menor = min(alto_actual, alto_ce)
            y_barra_menor = eje_y - alto_barra_menor
            pdf.set_draw_color(39, 174, 96)
            pdf.set_line_width(0.6)
            pdf.line(x_barra_menor - 1, y_barra_menor - 8, x_barra_menor + 9, y_barra_menor - 8)
            pdf.set_text_color(20, 90, 50)
            pdf.set_font("Arial", "B", 4)
            pdf.set_xy(centro - 5, y_barra_menor - 11)
            etiqueta_ahorro = "AHORRO" if ahorro_es_positivo else "COSTO EXTRA"
            pdf.cell(26, 3, f"{etiqueta_ahorro} ${abs(resultado['Ahorro']):,.0f}", align="C")

            pdf.set_xy(centro - 3, eje_y + 1)
            pdf.set_text_color(70, 80, 74)
            pdf.set_font("Arial", "", 6)
            pdf.cell(25, 4, meses[indice], align="C")

        titulo_seccion(163, "Desglose mensual", "Facturación estimada para cada uno de los seis períodos")
        anchos = [21, 25, 23, 28, 30, 29, 30]
        encabezados_tabla = ["Mes", "Consumo kWh", "Cobertura", "Tarifa CE", "Tradicional", "Con CE", "Ahorro"]
        pdf.set_xy(12, 174)
        pdf.set_fill_color(20, 90, 50)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "B", 6.5)
        for ancho, texto in zip(anchos, encabezados_tabla):
            pdf.cell(ancho, 7, texto, align="C", fill=True)
        pdf.ln()
        for indice, resultado in enumerate(datos_mensuales):
            pdf.set_x(12)
            if indice % 2 == 0:
                pdf.set_fill_color(244, 247, 245)
            else:
                pdf.set_fill_color(255, 255, 255)
            pdf.set_text_color(44, 62, 50)
            pdf.set_font("Arial", "", 6.5)
            valores = [
                meses[indice], f"{resultado['Consumo']:,.0f}", f"{resultado['Cobertura CE']:,.0f}%",
                f"${resultado['Tarifa CE']:,.0f}", f"${resultado['Total Actual']:,.0f}",
                f"${resultado['Total CE']:,.0f}", f"${resultado['Ahorro']:,.0f}",
            ]
            for columna, (ancho, valor) in enumerate(zip(anchos, valores)):
                if columna == len(anchos) - 1:
                    pdf.set_text_color(39, 174, 96)
                    pdf.set_font("Arial", "B", 6.5)
                pdf.cell(ancho, 7, valor, align="C", fill=True)
            pdf.ln()

        titulo_seccion(228, "Beneficios comerciales")
        for indice, (valor, etiqueta) in enumerate([
            ("25%", "Diferencial sobre kWh asignado"),
            ("80%", f"Cobertura mínima de referencia (promedio simulado: {cobertura_media:.0f}%)"),
            ("120,268 kWh", "Consumo respaldado del proyecto"),
        ]):
            x = 12 + indice * 63
            pdf.set_fill_color(232, 245, 233)
            pdf.rect(x, 237, 58, 20, style="F")
            pdf.set_text_color(20, 90, 50)
            pdf.set_font("Arial", "B", 10 if indice < 2 else 9)
            pdf.set_xy(x + 3, 238)
            pdf.cell(52, 7, valor, align="C")
            pdf.set_text_color(70, 85, 75)
            pdf.set_font("Arial", "", 6)
            pdf.set_xy(x + 3, 246)
            pdf.multi_cell(52, 4, etiqueta, align="C")

        pdf.set_xy(12, 263)
        pdf.set_text_color(95, 108, 99)
        pdf.set_font("Arial", "I", 7)
        pdf.multi_cell(186, 4, "El diferencial del 25% aplica al precio de la Energía Comunitaria Asignada; no representa una reducción garantizada sobre la factura total.")
        pdf.set_fill_color(20, 90, 50)
        pdf.rect(0, 285, 210, 12, style="F")
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", "", 7)
        pdf.set_xy(12, 288)
        pdf.cell(186, 5, "Proyección comercial semestral  |  COP  |  Comunidad Energética", align="C")

        pdf.add_page()
        encabezado("PROPUESTA COMERCIAL", f"Condiciones para: {cliente}")
        titulo_seccion(54, "Características de la propuesta y condiciones del contrato", "Síntesis informativa basada en la minuta de suministro para usuario comercial")
        condiciones = [
            ("Precio y beneficio tarifario", "El precio previsto equivale al 75% de la Tarifa Aplicada por cada kWh efectivamente asignado: un diferencial esperado del 25% sobre esa energía. No garantiza una reducción del 25% en la factura total, que incluye energía no cubierta y otros componentes."),
            ("Asignación y cobertura", "La asignación y el PDE pueden variar cada mes según generación exportada, consumos, composición de la comunidad y medición. No exceden el consumo real; el PDE individual no debe superar el 10%. La cobertura mínima del 80% del consumo base depende de energía disponible y de las condiciones técnicas, operativas y regulatorias de la minuta."),
            ("Facturación y pago", "El cobro puede ser directo o reflejarse en la factura habilitada. Debe identificar kWh asignados, precio unitario, valor y período. Si se factura por separado, el plazo previsto es de siete (7) días hábiles desde la expedición; si se integra a otra factura, aplica el plazo de esta última."),
            ("Continuidad del suministro", "La minuta no garantiza un volumen fijo y constante de energía mensual. Ante fallas, indisponibilidad o generación insuficiente, el usuario cubre la energía no recibida con su comercializador convencional. Una compensación por indisponibilidad imputable al generador requiere acuerdo en el ACE o anexo económico."),
            ("Vigencia y retiro voluntario", f"La minuta prevé una vigencia inicial de ocho (8) años, con prórrogas automáticas iguales salvo aviso escrito de no renovación con treinta (30) días de anticipación. Esta propuesta indica {dias_retiro} días de preaviso para retiro voluntario; la minuta fija un mínimo de treinta (30) días, salvo plazo superior en el ACE o contrato. El retiro requiere estar al día o acordar el pago y completar los ajustes operativos."),
            ("Medición y permanencia", "El usuario debe facilitar medición y telemedida, conservar equipos y permitir verificaciones. La minuta restringe vincularse a autogeneración u otros esquemas que desplacen materialmente el consumo sin autorización escrita. Una reducción sostenida por debajo del 70% del consumo evaluado puede dar lugar a revisión de permanencia y retiro conforme al contrato."),
        ]
        ubicaciones = [(12, 70), (108, 70), (12, 132), (108, 132), (12, 194), (108, 194)]
        for indice, ((titulo, texto), (x, y)) in enumerate(zip(condiciones, ubicaciones), start=1):
            bloque_condicion(x, y, indice, titulo, texto)

        pdf.set_fill_color(20, 90, 50)
        pdf.rect(0, 270, 210, 27, style="F")
        pdf.set_text_color(255, 255, 255)
        pdf.set_xy(12, 275)
        pdf.set_font("Arial", "B", 8)
        pdf.cell(186, 5, "NOTA IMPORTANTE")
        pdf.set_xy(12, 281)
        pdf.set_font("Arial", "", 7)
        pdf.multi_cell(186, 4, "Este reporte resume una propuesta comercial y no reemplaza el contrato firmado, el ACE ni sus anexos. Las condiciones definitivas deben constar en los documentos suscritos por las partes.")

        pdf.output(ruta_pdf)
        with open(ruta_pdf, "rb") as archivo_pdf:
            return archivo_pdf.read()
    finally:
        for ruta_temporal in (ruta_logo, ruta_pdf):
            if os.path.exists(ruta_temporal):
                os.remove(ruta_temporal)

with col_btn2:
    st.download_button(
        label="Descargar Reporte en PDF",
        data=generar_pdf(resultados, meses_labels, nombre_cliente, dias_retiro),
        file_name="Reporte_Comunidad_Energetica.pdf",
        mime="application/pdf",
        type="primary",
        on_click="ignore",
    )

if st.session_state.mostrar_desglose:
    st.subheader("Análisis Detallado por Concepto (Estructura de Modelo)")
    conceptos = [
        "Consumo Energía Activa", "Compra Energía Asignada CE (-)",
        "Cobro Energía Comunidad Energética", "Contribución", "Alumbrado Público", "TOTAL FACTURA",
    ]
    def formato_moneda(lista):
        return [f"${valor:,.0f}" for valor in lista]

    df_desglose = pd.DataFrame({"Concepto Facturado": conceptos})
    for indice, resultado in enumerate(resultados, start=1):
        df_desglose[f"Mes {indice} (Sin CE)"] = formato_moneda(resultado["Detalle Actual"])
        df_desglose[f"Mes {indice} (Con CE)"] = formato_moneda(resultado["Detalle CE"])
    st.dataframe(df_desglose, width="stretch", hide_index=True)