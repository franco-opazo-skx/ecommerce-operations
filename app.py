import streamlit as st
import pandas as pd
from pathlib import Path

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
col1.metric("Estado", "En desarrollo")
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
    if df_style is None:
        st.error("Primero debes cargar el archivo Style-Color.")
    else:
        st.success("Siguiente paso: conectar esta configuración al motor real.")
        st.write("Prioridad actual:")
        st.write(st.session_state.prioridades)

st.caption("Plataforma Ecommerce Operations · Franco Opazo")
