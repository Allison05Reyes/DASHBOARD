import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config("Dashboard Atenciones", layout="wide")

# 1) Configura un dataset por defecto alojado en tu repo (evita abrir explorador)
DEFAULT_URL = "https://raw.githubusercontent.com/Allison05Reyes/DASHBOARD1/main/data/atenciones.parquet"

st.sidebar.subheader("Origen de datos")
origen = st.sidebar.radio(
    "Elige cómo cargar tus datos",
    ["URL (recomendado)", "Archivo del repositorio", "Subir archivo (drag & drop)"],
    index=0
)

# ---- Lectores cacheados ----
@st.cache_data(ttl=600)
def read_parquet_from_url(url: str) -> pd.DataFrame:
    return pd.read_parquet(url, engine="pyarrow")

@st.cache_data(ttl=600)
def read_csv_from_url(url: str) -> pd.DataFrame:
    return pd.read_csv(url)

@st.cache_data(ttl=600)
def read_excel_once(content: bytes) -> pd.DataFrame:
    bio = BytesIO(content)
    return pd.read_excel(bio)  # openpyxl por defecto para .xlsx

def tipa(df: pd.DataFrame) -> pd.DataFrame:
    # Ajusta a tus columnas reales
    cat_cols = ["SEDE","EPS","PROGRAMA","ESPECIALIDAD","PROFESIONAL","GENERO"]
    date_cols = ["FECHA_ATENCION","MES_ATENCION"]
    for c in cat_cols:
        if c in df.columns:
            df[c] = df[c].astype("category")
    for c in date_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

df = None

if origen == "URL (recomendado)":
    st.sidebar.write("Pega un enlace a un **CSV o Parquet** (GitHub raw / Google Drive público).")
    url = st.sidebar.text_input("URL de datos", value=DEFAULT_URL, help="Soporta .csv o .parquet")
    if url:
        with st.spinner("Cargando desde URL..."):
            if url.lower().endswith(".parquet"):
                df = read_parquet_from_url(url)
            elif url.lower().endswith(".csv"):
                df = read_csv_from_url(url)
            else:
                st.warning("Extensión no reconocida. Usa .csv o .parquet")

elif origen == "Archivo del repositorio":
    st.sidebar.write("Usando archivo incluido en tu repo (sin explorador del sistema).")
    with st.spinner("Cargando archivo del repo..."):
        df = read_parquet_from_url(DEFAULT_URL)  # cambia si usas CSV

else:  # "Subir archivo (drag & drop)"
    upl = st.file_uploader("Arrastra aquí tu CSV/XLSX (evita el diálogo de carpetas)", type=["csv","xlsx"])
    if upl is not None:
        # límite suave para no reventar la RAM del plan gratis
        if upl.size > 80 * 1024 * 1024:
            st.error("Archivo grande (>80 MB). Convierte a CSV/Parquet o usa la opción por URL.")
            st.stop()
        with st.spinner("Leyendo archivo..."):
            if upl.name.lower().endswith(".csv"):
                df = pd.read_csv(upl)
            else:
                df = read_excel_once(upl.getvalue())

if df is None:
    st.info("Selecciona un origen de datos o pega una URL válida para continuar.")
    st.stop()

# Tipado y vista rápida
df = tipa(df)
st.success(f"Datos cargados: {len(df):,} filas, {len(df.columns)} columnas")
st.dataframe(df.head(100), use_container_width=True)

# --- Ejemplo de filtros rápidos (ajusta a tus columnas) ---
st.sidebar.header("Filtros")
def pick(col):
    if col in df.columns:
        opts = ["(Todos)"] + sorted([str(x) for x in df[col].dropna().unique()])
        return st.sidebar.selectbox(col, opts, index=0)
    return None

f_eps = pick("EPS")
f_esp = pick("ESPECIALIDAD")
f_prog = pick("PROGRAMA")
f_sede = pick("SEDE")

mask = pd.Series(True, index=df.index)
if f_eps and f_eps != "(Todos)":  mask &= df["EPS"].astype(str).eq(f_eps)
if f_esp and f_esp != "(Todos)":  mask &= df["ESPECIALIDAD"].astype(str).eq(f_esp)
if f_prog and f_prog != "(Todos)": mask &= df["PROGRAMA"].astype(str).eq(f_prog)
if f_sede and f_sede != "(Todos)": mask &= df["SEDE"].astype(str).eq(f_sede)

dff = df.loc[mask]
st.markdown(f"**Registros filtrados:** {len(dff):,}")
