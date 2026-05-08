import streamlit as st
import pandas as pd
from pathlib import Path
from engine_descuentos import calcular_descuentos, dataframe_to_excel_bytes

st.set_page_config(
    page_title="Skechers Ecommerce Operations",
    page_icon="👟",
    layout="wide"
)

DATA_DIR = Path("data")
ASSETS_DIR = Path("assets")
LOGO_PATH = ASSETS_DIR / "SKECHERS_logo.png"

FILES = {
    "Listado Prod. 1P": DATA_DIR / "Listado Prod. 1P.xlsx",
    "Descuentos Retail": DATA_DIR / "Descuentos Retail.xlsx",
    "Tech Sports": DATA_DIR / "Tech Sports.xlsx",
    "Compras Retail-Ecomm-BTS": DATA_DIR / "Compras Retail-Ecomm-BTS.xlsx",
    "Disponible": DATA_DIR / "Disponible.xlsx",
}

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container{
    padding-top: 1rem;
    padding-left: 2rem;
    padding-right: 2rem;
    max-width: 1260px;
}

[data-testid="stSidebar"]{
    background: #061426;
}

[data-testid="stSidebar"] *{
    color: white;
}

.sidebar-logo{
    text-align:center;
    margin-top:35px;
    margin-bottom:55px;
}

.sidebar-section{
    background:#082B3A;
    border-radius:12px;
    padding:14px;
    font-weight:700;
    margin-bottom:36px;
}

.hero{
    background: linear-gradient(90deg, #061426 0%, #0B2B4C 100%);
    border-radius:18px;
    padding:22px 26px;
    margin-bottom:18px;
    color:white;
    display:flex;
    align-items:center;
    gap:24px;
}

.hero img{
    width:190px;
    max-height:70px;
    object-fit:contain;
}

.hero-title{
    font-size:30px;
    font-weight:850;
    margin-bottom:4px;
}

.hero-sub{
    font-size:13px;
    color:#CBD5E1;
}

.metric-card{
    background:#FFFFFF;
    border:1px solid #E5E7EB;
    border-radius:14px;
    padding:13px 16px;
    box-shadow:0 1px 5px rgba(15,23,42,0.05);
}

.metric-title{
    font-size:11px;
    color:#64748B;
    text-transform:uppercase;
    font-weight:800;
    letter-spacing:.04em;
}

.metric-value{
    font-size:21px;
    color:#061426;
    font-weight:850;
}

.card-section{
    background:#FFFFFF;
    border:1px solid #E5E7EB;
    border-radius:15px;
    padding:18px 20px;
    margin-bottom:18px;
    box-shadow:0 1px 6px rgba(15,23,42,0.04);
}

.section-label{
    font-size:11px;
    color:#64748B;
    text-transform:uppercase;
    font-weight:850;
    letter-spacing:.05em;
    margin-bottom:4px;
}

.section-heading{
    font-size:21px;
    font-weight:850;
    color:#061426;
    margin-bottom:4px;
}

.executive-note{
    color:#64748B;
    font-size:13px;
}

.rule-row{
    background:#F8FAFC;
    border:1px solid #E5E7EB;
    border-radius:12px;
    padding:10px 14px;
    margin-bottom:8px;
}

.rule-title{
    font-size:15px;
    font-weight:750;
    color:#061426;
}

.stButton>button{
    background-color:#0B2B4C;
    color:white;
    border-radius:10px;
    border:none;
    font-weight:750;
    padding:8px 16px;
}

.stButton>button:hover{
    background-color:#123E6B;
    color:white;
}

div[data-testid="stDownloadButton"] button{
    background-color:#061426;
    color:white;
    border-radius:10px;
    border:none;
    font-weight:750;
}

[data-testid="stFileUploader"]{
    background:#F8FAFC;
    border-radius:12px;
}

h1, h2, h3{
    color:#061426;
}

hr{
    margin-top:20px;
    margin-bottom:20px;
}
</style>
""", unsafe_allow_html=True)


def norm(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def load_1p_options(path):
    df = pd.read_excel(path, header=0, usecols="D:H")
    vals = set()
    for col in df.columns:
        vals.update(df[col].dropna().astype(str).str.strip())
    return sorted([v for v in vals if v and v != "-"])


def load_tech_options(path):
    df = pd.read_excel(path, header=0, usecols="A,I,K")
    df.columns = ["StyleColor", "Season Actual", "Detalle"]
    df["Detalle"] = df["Detalle"].map(norm)
    df["Season Actual"] = df["Season Actual"].map(norm)

    detalles = sorted([v for v in df["Detalle"].unique() if v and v != "-"])
    apto_df = df[df["Detalle"].str.upper() == "APTO PARA DSCTO"]
    seasons_apto = sorted([v for v in apto_df["Season Actual"].unique() if v and v != "-"])

    return detalles, seasons_apto


def load_compras_options(path):
    df = pd.read_excel(path, header=0, usecols="B,M:Y")
    df.columns = [
        "StyleColor",
        "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"
    ]

    grupos = {
        "COMPRA RETAIL BTS-26": ["M"],
        "COMPRA RETAIL 2026-1": ["N", "O", "P"],
        "COMPRA ECOMM 2026-1": ["Q"],
        "COMPRA ECOMM BTS-26": ["R"],
        "COMPRA RETAIL 2025-2": ["S", "T", "U"],
        "COMPRA RETAIL 2026-2": ["V", "W", "X"],
        "COMPRA ECOMM 2026-2": ["Y"],
    }

    encontrados = []
    for grupo, cols in grupos.items():
        for c in cols:
            serie = df[c].fillna("-").astype(str).str.strip()
            if (serie != "-").any():
                encontrados.append(grupo)
                break

    return encontrados


with st.sidebar:
    st.markdown("<div class='sidebar-logo'>", unsafe_allow_html=True)
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
    else:
        st.markdown("## SKECHERS")
    st.markdown("</div>", unsafe_allow_html=True)

    st.caption("Ecommerce Operations")
    st.markdown("<div class='sidebar-section'>Cálculo Descuentos Ecommerce</div>", unsafe_allow_html=True)
    st.caption("From: Franco Opazo")


if LOGO_PATH.exists():
    st.markdown("<div class='hero'>", unsafe_allow_html=True)
    c_logo, c_text = st.columns([1.1, 4])
    with c_logo:
        st.image(str(LOGO_PATH), use_container_width=True)
    with c_text:
        st.markdown("""
        <div class="hero-title">Ecommerce Operations</div>
        <div class="hero-sub">Motor corporativo para cálculo de descuentos ecommerce · From Franco Opazo</div>
        """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="hero">
        <div>
            <div class="hero-title">SKECHERS · Ecommerce Operations</div>
            <div class="hero-sub">Motor corporativo para cálculo de descuentos ecommerce · From Franco Opazo</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="metric-card">
        <div class="metric-title">Estado</div>
        <div class="metric-value">Online</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="metric-card">
        <div class="metric-title">Motor</div>
        <div class="metric-value">Python</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="metric-card">
        <div class="metric-title">Output</div>
        <div class="metric-value">Excel</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    archivos_ok = sum(1 for p in FILES.values() if p.exists())
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">Archivos base</div>
        <div class="metric-value">{archivos_ok}/5 OK</div>
    </div>
    """, unsafe_allow_html=True)

with st.expander("Ver estado técnico de archivos internos", expanded=False):
    for name, path in FILES.items():
        if path.exists():
            st.success(f"{name}: OK")
        else:
            st.error(f"{name}: No encontrado → {path}")

st.markdown("<hr>", unsafe_allow_html=True)

st.markdown("""
<div class="card-section">
    <div class="section-label">Input principal</div>
    <div class="section-heading">Carga de archivo Style-Color</div>
    <div class="executive-note">Sube el archivo base con los Style-Color que serán evaluados por el motor.</div>
</div>
""", unsafe_allow_html=True)

archivo = st.file_uploader("Archivo Style-Color", type=["xlsx"])
df_style = None

if archivo:
    df_style = pd.read_excel(archivo)
    st.success(f"Archivo cargado: {archivo.name}")
    st.info(f"Filas encontradas: {len(df_style)}")
    with st.expander("Ver vista previa del archivo cargado", expanded=False):
        st.dataframe(df_style.head(20), use_container_width=True)

st.markdown("<hr>", unsafe_allow_html=True)

st.markdown("""
<div class="card-section">
    <div class="section-label">Reglas de negocio</div>
    <div class="section-heading">Prioridad de evaluación</div>
    <div class="executive-note">La primera regla que aplique define el resultado final del producto.</div>
</div>
""", unsafe_allow_html=True)

reglas_base = [
    "Descuento Retail",
    "Tech Sports",
    "Compras Retail-Ecomm-BTS",
    "Listado Prod. 1P",
    "Key Initiative",
]

if "prioridades" not in st.session_state:
    st.session_state.prioridades = reglas_base.copy()

for i, regla in enumerate(st.session_state.prioridades):
    c1, c2, c3 = st.columns([8, 1, 1])
    with c1:
        st.markdown(f"""
        <div class="rule-row">
            <div class="rule-title">{i + 1}. {regla}</div>
        </div>
        """, unsafe_allow_html=True)

    if c2.button("⬆️", key=f"up_{i}", disabled=i == 0):
        st.session_state.prioridades[i - 1], st.session_state.prioridades[i] = (
            st.session_state.prioridades[i],
            st.session_state.prioridades[i - 1],
        )
        st.rerun()

    if c3.button("⬇️", key=f"down_{i}", disabled=i == len(st.session_state.prioridades) - 1):
        st.session_state.prioridades[i + 1], st.session_state.prioridades[i] = (
            st.session_state.prioridades[i],
            st.session_state.prioridades[i + 1],
        )
        st.rerun()

if st.button("Restablecer orden"):
    st.session_state.prioridades = reglas_base.copy()
    st.rerun()

st.markdown("<hr>", unsafe_allow_html=True)

st.markdown("""
<div class="card-section">
    <div class="section-label">Parámetros dinámicos</div>
    <div class="section-heading">Configuración de descuentos</div>
    <div class="executive-note">Define qué valores permiten descuento y el porcentaje correspondiente.</div>
</div>
""", unsafe_allow_html=True)

tab_1p, tab_tech, tab_compras = st.tabs([
    "Listado Prod. 1P",
    "Tech Sports",
    "Compras Retail-Ecomm-BTS",
])

config_1p = pd.DataFrame()
config_tech = pd.DataFrame()
config_tech_season = pd.DataFrame()
config_compras = pd.DataFrame()

with tab_1p:
    if FILES["Listado Prod. 1P"].exists():
        vals_1p = load_1p_options(FILES["Listado Prod. 1P"])
        df_1p = pd.DataFrame({
            "Valor detectado": vals_1p,
            "Permite descuento": [False] * len(vals_1p),
            "%": [0.0] * len(vals_1p),
        })
        config_1p = st.data_editor(df_1p, use_container_width=True, key="config_1p")
    else:
        st.error("No se encontró Listado Prod. 1P.")

with tab_tech:
    if FILES["Tech Sports"].exists():
        detalles, seasons_apto = load_tech_options(FILES["Tech Sports"])

        df_tech = pd.DataFrame({
            "Valor detectado": detalles,
            "Permite descuento": [v.upper() == "APTO PARA DSCTO" for v in detalles],
            "%": [0.0] * len(detalles),
        })

        config_tech = st.data_editor(df_tech, use_container_width=True, key="config_tech")

        apto_row = config_tech[
            config_tech["Valor detectado"].astype(str).str.upper().str.strip() == "APTO PARA DSCTO"
        ]

        apto_permite = False
        if not apto_row.empty:
            apto_permite = bool(apto_row.iloc[0]["Permite descuento"])

        if apto_permite:
            st.markdown("##### Porcentaje por Season para APTO PARA DSCTO")
            df_season = pd.DataFrame({
                "Season Actual": seasons_apto,
                "%": [0.0] * len(seasons_apto),
            })
            config_tech_season = st.data_editor(
                df_season,
                use_container_width=True,
                key="config_tech_season",
            )
        else:
            st.info("Activa 'Permite descuento' en APTO PARA DSCTO para configurar porcentaje por Season.")
            config_tech_season = pd.DataFrame(columns=["Season Actual", "%"])
    else:
        st.error("No se encontró Tech Sports.")

with tab_compras:
    if FILES["Compras Retail-Ecomm-BTS"].exists():
        grupos_compras = load_compras_options(FILES["Compras Retail-Ecomm-BTS"])
        df_compras = pd.DataFrame({
            "Grupo detectado": grupos_compras,
            "Permite descuento": [False] * len(grupos_compras),
            "%": [0.0] * len(grupos_compras),
        })
        config_compras = st.data_editor(
            df_compras,
            use_container_width=True,
            key="config_compras",
        )
    else:
        st.error("No se encontró Compras Retail-Ecomm-BTS.")

st.markdown("<hr>", unsafe_allow_html=True)

st.markdown("""
<div class="card-section">
    <div class="section-label">Resultado</div>
    <div class="section-heading">Ejecutar cálculo y descargar output</div>
    <div class="executive-note">Procesa las reglas configuradas y descarga el archivo final en formato Excel.</div>
</div>
""", unsafe_allow_html=True)

if st.button("Ejecutar cálculo", type="primary"):
    if archivo is None:
        st.error("Primero debes cargar el archivo Style-Color.")
    else:
        try:
            with st.spinner("Procesando cálculo de descuentos..."):
                output_df = calcular_descuentos(
                    style_file=archivo,
                    files=FILES,
                    prioridades=st.session_state.prioridades,
                    config_1p_df=config_1p,
                    config_tech_df=config_tech,
                    config_tech_season_df=config_tech_season,
                    config_compras_df=config_compras,
                )

                excel_bytes = dataframe_to_excel_bytes(output_df)

            st.success("Cálculo terminado correctamente.")

            c1, c2, c3 = st.columns(3)
            c1.metric("Productos procesados", len(output_df))
            c2.metric("Habilitados", int((output_df["Habilitado/No Habilitado"] == "Habilitado para tener Descuento").sum()))
            c3.metric("No habilitados", int((output_df["Habilitado/No Habilitado"] != "Habilitado para tener Descuento").sum()))

            with st.expander("Ver vista previa del output", expanded=True):
                st.dataframe(output_df.head(50), use_container_width=True)

            st.download_button(
                label="Descargar output_descuentos_ecommerce.xlsx",
                data=excel_bytes,
                file_name="output_descuentos_ecommerce.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error ejecutando cálculo: {e}")

st.caption("Plataforma Ecommerce Operations · Franco Opazo")
