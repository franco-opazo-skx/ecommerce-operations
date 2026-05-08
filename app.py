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
    padding-top: 1.3rem;
    padding-left: 2rem;
    padding-right: 2rem;
    max-width: 1280px;
}

[data-testid="stSidebar"]{
    background: #071527;
}

[data-testid="stSidebar"] *{
    color: white;
}

.hero{
    background: linear-gradient(90deg, #071527 0%, #0B2B4C 100%);
    border-radius: 18px;
    padding: 24px 30px;
    margin-bottom: 22px;
    color: white;
}

.hero-title{
    font-size: 32px;
    font-weight: 850;
    margin-bottom: 4px;
}

.hero-sub{
    font-size: 14px;
    color: #D1D5DB;
}

.metric-card{
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 14px;
    padding: 15px 18px;
    box-shadow: 0 1px 6px rgba(15,23,42,0.05);
}

.metric-title{
    font-size: 12px;
    color: #64748B;
    text-transform: uppercase;
    font-weight: 700;
}

.metric-value{
    font-size: 23px;
    color: #071527;
    font-weight: 800;
}

.card-section{
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 16px;
    padding: 22px;
    margin-bottom: 22px;
    box-shadow: 0 2px 8px rgba(15,23,42,0.04);
}

.section-label{
    font-size: 12px;
    color: #64748B;
    text-transform: uppercase;
    font-weight: 800;
    letter-spacing: .04em;
}

.section-heading{
    font-size: 22px;
    font-weight: 850;
    color: #071527;
    margin-bottom: 6px;
}

.executive-note{
    color: #64748B;
    font-size: 14px;
    margin-bottom: 14px;
}

.stButton>button{
    background-color: #0B2B4C;
    color: white;
    border-radius: 10px;
    border: none;
    font-weight: 700;
}

.stButton>button:hover{
    background-color: #123E6B;
    color: white;
}

div[data-testid="stDownloadButton"] button{
    background-color: #071527;
    color: white;
    border-radius: 10px;
    border: none;
    font-weight: 700;
}

h1, h2, h3{
    color: #071527;
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
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
    else:
        st.markdown("## SKECHERS")

    st.caption("Ecommerce Operations")
    st.divider()
    st.success("Cálculo Descuentos Ecommerce")
    st.divider()
    st.caption("From: Franco Opazo")


st.markdown("""
<div class="hero">
    <div class="hero-title">SKECHERS · Ecommerce Operations</div>
    <div class="hero-sub">Motor corporativo para cálculo de descuentos ecommerce · From Franco Opazo</div>
</div>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

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

st.divider()

with st.expander("Estado de conexión archivos internos", expanded=False):
    for name, path in FILES.items():
        if path.exists():
            st.success(f"{name}: OK")
        else:
            st.error(f"{name}: No encontrado → {path}")

st.markdown("""
<div class="card-section">
    <div class="section-label">Paso 1</div>
    <div class="section-heading">Carga de archivo Style-Color</div>
    <div class="executive-note">Sube el archivo base con los Style-Color que serán evaluados por el motor.</div>
</div>
""", unsafe_allow_html=True)

archivo = st.file_uploader("Archivo Style-Color", type=["xlsx"])
df_style = None

if archivo:
    df_style = pd.read_excel(archivo)
    st.success(f"Archivo cargado: {archivo.name}")
    st.dataframe(df_style.head(20), use_container_width=True)
    st.info(f"Filas encontradas: {len(df_style)}")

st.divider()

st.markdown("""
<div class="card-section">
    <div class="section-label">Paso 2</div>
    <div class="section-heading">Prioridad de reglas</div>
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
    c1.markdown(f"**{i + 1}. {regla}**")

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

st.divider()

st.markdown("""
<div class="card-section">
    <div class="section-label">Paso 3</div>
    <div class="section-heading">Configuración dinámica de reglas</div>
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
            st.markdown("##### % por Season para APTO PARA DSCTO")
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
            st.info("Activa 'Permite descuento' en APTO PARA DSCTO para configurar % por Season.")
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

st.divider()

st.markdown("""
<div class="card-section">
    <div class="section-label">Paso 4</div>
    <div class="section-heading">Ejecutar y descargar output</div>
    <div class="executive-note">Procesa las reglas configuradas y descarga el archivo final en Excel.</div>
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

            st.success("✅ Cálculo terminado correctamente.")
            st.dataframe(output_df.head(50), use_container_width=True)

            st.download_button(
                label="📥 Descargar output_descuentos_ecommerce.xlsx",
                data=excel_bytes,
                file_name="output_descuentos_ecommerce.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error ejecutando cálculo: {e}")

st.caption("Plataforma Ecommerce Operations · Franco Opazo")
