import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# Ejemplo de URL por defecto (cámbiala por la tuya)
DEFAULT_URL = "https://raw.githubusercontent.com/usuario/repo/main/data/archivo.parquet"

st.set_page_config("Dashboard Atenciones", layout="wide")

@st.cache_data(ttl=600)
def fetch_bytes(url: str, *, github_token: str | None = None) -> bytes:
    """Descarga bytes con soporte para repos privados (token opcional)."""
    headers = {"User-Agent": "streamlit-app"}
    if github_token:
        headers["Authorization"] = f"token {github_token}"
    r = requests.get(url, headers=headers, timeout=30)
    try:
        r.raise_for_status()
    except requests.HTTPError as e:
        raise RuntimeError(f"HTTP {r.status_code} al descargar {url}") from e
    return r.content

@st.cache_data(ttl=600)
def read_df_from_url(url: str, *, github_token: str | None = None) -> pd.DataFrame:
    """
    Lee DataFrame desde URL en formatos:
    - .parquet  (pyarrow)
    - .csv      (pandas)
    - .xlsx/.xls (pandas + openpyxl/xlrd)
    Si la extensión no coincide, intenta por contenido.
    """
    raw = fetch_bytes(url, github_token=github_token)
    bio = BytesIO(raw)
    lower = url.lower()

    # Helper internos
    def try_parquet(b):
        b.seek(0)
        return pd.read_parquet(b, engine="pyarrow")

    def try_csv(b):
        b.seek(0)
        return pd.read_csv(b)

    def try_excel(b):
        b.seek(0)
        # pandas detecta engine: openpyxl para .xlsx, xlrd para .xls (si disponible)
        return pd.read_excel(b)

    # Selección por extensión + fallback inteligente
    try:
        if lower.endswith(".parquet"):
            try:
                return try_parquet(bio)
            except Exception:
                # Por si en realidad era CSV/Excel con extensión mal puesta
                try:
                    return try_csv(bio)
                except Exception:
                    return try_excel(bio)
        elif lower.endswith(".csv"):
            try:
                return try_csv(bio)
            except Exception:
                # Intento parquet o excel si el contenido no era CSV
                try:
                    return try_parquet(bio)
                except Exception:
                    return try_excel(bio)
        elif lower.endswith(".xlsx") or lower.endswith(".xls"):
            try:
                return try_excel(bio)
            except Exception:
                # Intento parquet o csv si el contenido no era Excel
                try:
                    return try_parquet(bio)
                except Exception:
                    return try_csv(bio)
        else:
            # Extensión desconocida: detectar por contenido
            for reader in (try_parquet, try_csv, try_excel):
                try:
                    return reader(bio)
                except Exception:
                    continue
            raise RuntimeError("No se pudo leer el archivo como Parquet/CSV/Excel. Revisa la URL RAW y el formato.")
    except Exception as e:
        raise RuntimeError(f"No se pudo procesar la URL '{url}'. Detalle: {e}") from e

# ---------- UI mínima de ejemplo ----------
st.sidebar.subheader("Origen de datos")
url = st.sidebar.text_input("URL RAW (.parquet / .csv / .xlsx / .xls)", value=DEFAULT_URL)

# Si tu repo es privado, define en Secrets: GITHUB_TOKEN="tu_pat"
gh_token = st.secrets.get("GITHUB_TOKEN", None)

if not url:
    st.stop()

with st.spinner("Cargando datos..."):
    try:
        df = read_df_from_url(url, github_token=gh_token)
    except Exception as e:
        st.error(str(e))
        st.stop()

st.success(f"Datos cargados: {len(df):,} filas × {len(df.columns)} columnas")
st.dataframe(df.head(100), use_container_width=True)
