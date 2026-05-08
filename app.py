import streamlit as st
import pandas as pd
import streamlit as st
import pandas as pd
from pathlib import Path
from engine_descuentos import calcular_descuentos, dataframe_to_excel_bytes

# =========================================
# CONFIG
# =========================================

st.set_page_config(
    page_title="Skechers Ecommerce Operations",
    page_icon="👟",
    layout="wide"
)

# =========================================
# CSS CORPORATIVO
# =========================================

st.markdown("""
<style>

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

.block-container{
    padding-top: 2rem;
    padding-bottom: 2rem;
    padding-left: 2rem;
    padding-right: 2rem;
}

[data-testid="stSidebar"]{
    background: #0F172A;
}

[data-testid="stSidebar"] *{
    color: white;
}

.skechers-title{
    font-size: 42px;
    font-weight: 800;
    color: #111827;
    margin-bottom: 0;
}

.skechers-sub{
    color: #6B7280;
    font-size: 15px;
    margin-top: -10px;
}

.metric-card{
    background: white;
    border-radius: 16px;
    padding: 22px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    border: 1px solid #E5E7EB;
}

.metric-title{
    color: #6B7280;
    font-size: 14px;
}

.metric-value{
    font-size: 34px;
    font-weight: 700;
    color: #111827;
}

.section-title{
    font-size: 32px;
    font-weight: 700;
    margin-top: 30px;
    margin-bottom: 15px;
}

.stButton>button{
    background-color: #EF4444;
    color: white;
    border-radius: 12px;
    border: none;
    padding: 12px 22px;
    font-weight: 600;
}

.stButton>button:hover{
    background-color: #DC2626;
    color: white;
}

hr{
    margin-top: 35px;
    margin-bottom: 35px;
}

.logo-wrap{
    display:flex;
    align-items:center;
    gap:20px;
    margin-bottom:20px;
}

.logo-wrap img{
    width:220px;
}

.small-status{
    font-size:12px;
    color:#6B7280;
    margin-top:5px;
}

</style>
""", unsafe_allow_html=True)

# =========================================
# PATHS
# =========================================

DATA_DIR = Path("data")

FILES = {
    "Listado Prod. 1P": DATA_DIR / "Listado Prod. 1P.xlsx",
    "Descuentos Retail": DATA_DIR / "Descuentos Retail.xlsx",
    "Tech Sports": DATA_DIR / "Tech Sports.xlsx",
    "Compras Retail-Ecomm-BTS": DATA_DIR / "Compras Retail-Ecomm-BTS.xlsx",
    "Disponible": DATA_DIR / "Disponible.xlsx",
}

# =========================================
# SIDEBAR
# =========================================

with st.sidebar:

    st.markdown("""
    <div style='margin-top:20px;margin-bottom:30px'>
        <img src='https://upload.wikimedia.org/wikipedia/commons/thumb/e/e2/Skechers_logo.svg/512px-Skechers_logo.svg.png' width='180'>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Ecommerce Operations")

    st.markdown("---")

    st.success("Cálculo Descuentos Ecommerce")

# =========================================
# HEADER
# =========================================

st.markdown("""
<div class='logo-wrap'>
    <img src='https://upload.wikimedia.org/wikipedia/commons/thumb/e/e2/Skechers_logo.svg/512px-Skechers_logo.svg.png'>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class='skechers-title'>
Plataforma Ecommerce Operations
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class='skechers-sub'>
Skechers Chile · Discount Engine Platform
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =========================================
# CARDS
# =========================================

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class='metric-card'>
        <div class='metric-title'>Estado</div>
        <div class='metric-value'>Online</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class='metric-card'>
        <div class='metric-title'>Motor</div>
        <div class='metric-value'>Python</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class='metric-card'>
        <div class='metric-title'>Deployment</div>
        <div class='metric-value'>Streamlit Cloud</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)
from pathlib import Path

from engine_descuentos import calcular_descuentos, dataframe_to_excel_bytes

st.set_page_config(page_title="Descuentos Ecommerce", page_icon="🔥", layout="wide")

DATA_DIR = Path("data")

FILES = {
    "Listado Prod. 1P": DATA_DIR / "Listado Prod. 1P.xlsx",
    "Descuentos Retail": DATA_DIR / "Descuentos Retail.xlsx",
    "Tech Sports": DATA_DIR / "Tech Sports.xlsx",
    "Compras Retail-Ecomm-BTS": DATA_DIR / "Compras Retail-Ecomm-BTS.xlsx",
    "Disponible": DATA_DIR / "Disponible.xlsx",
}

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

st.sidebar.title("Skechers")
st.sidebar.caption("Ecommerce Operations")
st.sidebar.divider()
st.sidebar.success("Cálculo Descuentos Ecommerce")
st.sidebar.divider()
st.sidebar.caption("From: Franco Opazo")

st.title("🔥 Plataforma Ecommerce Operations")
st.caption("From: Franco Opazo")
st.header("Cálculo de Descuentos Ecommerce")

col1, col2, col3 = st.columns(3)
col1.metric("Estado", "Operativo")
col2.metric("Motor", "Python + Streamlit")
col3.metric("Output", "Excel")

st.divider()

with st.expander("Estado de conexión archivos internos", expanded=False):
    for name, path in FILES.items():
        if path.exists():
            st.success(f"{name}: OK")
        else:
            st.error(f"{name}: No encontrado → {path}")

st.divider()

st.subheader("1. Cargar archivo Style-Color")

archivo = st.file_uploader("Carga archivo Style-Color", type=["xlsx"])
df_style = None

if archivo:
    df_style = pd.read_excel(archivo)
    st.success(f"Archivo cargado: {archivo.name}")
    st.dataframe(df_style.head(20), use_container_width=True)
    st.info(f"Filas encontradas: {len(df_style)}")

st.divider()

st.subheader("2. Prioridad de reglas")

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

st.subheader("3. Configuración dinámica de reglas")

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
    st.write("#### Configuración 1P")

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
    st.write("#### Configuración Tech Sports")

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
            st.write("##### % por Season para APTO PARA DSCTO")

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
    st.write("#### Configuración Compras")

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

st.subheader("4. Ejecutar cálculo")

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

            st.write("#### Vista previa output")
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
