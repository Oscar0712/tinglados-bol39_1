import streamlit as st
import pandas as pd
import unicodedata
import re
from io import BytesIO
from pathlib import Path

# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="BOL-39/2024 | Detalle de Proyectos",
    page_icon="🏫",
    layout="wide"
)

st.title("🏫 BOL-39/2024 — Detalle de Proyectos")
st.caption("Consulta de proyectos priorizados y no priorizados")

ARCHIVO_BASE = "Tabla_resumen_Tinglados_v5.xlsx"

# ============================================================
# ESTILO
# ============================================================
st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem;}
@media (max-width: 768px) {
    .block-container {padding-left: .75rem; padding-right: .75rem;}
    h1 {font-size: 1.5rem !important;}
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# FUNCIONES
# ============================================================
def norm(s):
    s = str(s).strip().upper()
    s = ''.join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[_\s]+", " ", s)
    return s

ALIASES = {
    "DEPARTAMENTO": ["DEPARTAMENTO"],
    "PROVINCIA": ["PROVINCIA"],
    "MUNICIPIO": ["MUNICIPIO"],

    "NOMBRE DE LA UNIDAD EDUCATIVA": [
        "NOMBRE DE LA UNIDAD EDUCATIVA",
        "UNIDAD EDUCATIVA",
        "NOMBRE UE"
    ],

    "ESTADO DE PRIORIZACIÓN": [
        "ESTADO DE PRIORIZACIÓN",
        "ESTADO DE PRIORIZACION",
        "ESTADO SELECCION",
        "ESTADO_SELECCION"
    ],

    "INDICE PRIORIZACION": [
        "INDICE PRIORIZACION",
        "ÍNDICE PRIORIZACIÓN",
        "INDICE_PRIORIZACION"
    ],

    "MATRÍCULA": [
        "MATRÍCULA",
        "MATRICULA",
        "INSCRITO SUM",
        "INSCRITO_SUM"
    ],

    "INDICADOR NBI": [
        "INDICADOR NBI",
        "INDICADOR_NBI"
    ],

    "ESTUDIANTES A 1KM": [
        "ESTUDIANTES A 1KM",
        "ESTUDIANTES 1KM",
        "ESTUDIANTES_1KM"
    ],

    "DENTRO POLIGONO RADIACIÓN": [
        "DENTRO POLIGONO RADIACIÓN",
        "DENTRO POLIGONO RADIACION",
        "DENTRO_POLIGONO_RADIACION"
    ],

    "MONTO PROYECTOS PRIORIZADOS": [
        "MONTO PROYECTOS PRIORIZADOS",
        "MONTO PRIORIZADO SELECCIONADO",
        "MONTO PRIORIZADO_SELECCIONADO"
    ]
}

def standardize_columns(df):
    normalized = {norm(c): c for c in df.columns}
    rename = {}

    for target, candidates in ALIASES.items():
        for candidate in candidates:
            key = norm(candidate)
            if key in normalized:
                rename[normalized[key]] = target
                break

    return df.rename(columns=rename)

def excel_bytes(df):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Detalle")
    return bio.getvalue()

@st.cache_data
def cargar_base(ruta):
    return pd.read_excel(ruta)

# ============================================================
# CARGA DE BASE
# ============================================================
ruta_base = Path(__file__).parent / ARCHIVO_BASE

if not ruta_base.exists():
    st.error(
        f"No se encontró el archivo '{ARCHIVO_BASE}'. "
        "Debe estar en la misma carpeta que app.py."
    )
    st.stop()

try:
    df = cargar_base(ruta_base)
except Exception as e:
    st.error(f"No se pudo leer la base: {e}")
    st.stop()

df = standardize_columns(df)

# ============================================================
# ESTADO DE PRIORIZACIÓN
# ============================================================
if "ESTADO DE PRIORIZACIÓN" not in df.columns:
    st.error("No se encontró la variable ESTADO DE PRIORIZACIÓN.")
    st.stop()

estado = (
    df["ESTADO DE PRIORIZACIÓN"]
    .fillna("")
    .astype(str)
    .str.strip()
)

estado_norm = (
    estado
    .str.upper()
    .str.normalize("NFKD")
    .str.encode("ascii", errors="ignore")
    .str.decode("utf-8")
)

pri = estado_norm.eq("PRIORIZADO")
npri = estado_norm.isin(["NO PRIORIZADO", "NO PRIORIZADA"])

df["ESTADO DE PRIORIZACIÓN"] = ""
df.loc[pri, "ESTADO DE PRIORIZACIÓN"] = "Priorizado"
df.loc[npri, "ESTADO DE PRIORIZACIÓN"] = "No priorizado"

# SOLO se conservan Priorizado y No priorizado
df = df[
    df["ESTADO DE PRIORIZACIÓN"].isin(
        ["Priorizado", "No priorizado"]
    )
].copy()

# ============================================================
# VALIDACIÓN TERRITORIAL
# ============================================================
territorial = ["DEPARTAMENTO", "PROVINCIA", "MUNICIPIO"]
faltan = [c for c in territorial if c not in df.columns]

if faltan:
    st.error("Faltan campos territoriales: " + ", ".join(faltan))
    st.stop()

# ============================================================
# FILTROS
# ============================================================
st.subheader("Filtros")

c1, c2, c3 = st.columns(3)

deps = sorted(df["DEPARTAMENTO"].dropna().astype(str).unique())
dep = c1.selectbox("Departamento", ["Todos"] + deps)

d1 = df if dep == "Todos" else df[df["DEPARTAMENTO"].astype(str) == dep]

provs = sorted(d1["PROVINCIA"].dropna().astype(str).unique())
prov = c2.selectbox("Provincia", ["Todos"] + provs)

d2 = d1 if prov == "Todos" else d1[d1["PROVINCIA"].astype(str) == prov]

muns = sorted(d2["MUNICIPIO"].dropna().astype(str).unique())
mun = c3.selectbox("Municipio", ["Todos"] + muns)

base = d2 if mun == "Todos" else d2[d2["MUNICIPIO"].astype(str) == mun]
base = base.copy()

# ============================================================
# FILTRO DE ESTADO Y BÚSQUEDA
# ============================================================
f1, f2 = st.columns(2)

estado_filtro = f1.selectbox(
    "Estado de priorización",
    ["Todos", "Priorizado", "No priorizado"]
)

if estado_filtro != "Todos":
    base = base[
        base["ESTADO DE PRIORIZACIÓN"] == estado_filtro
    ].copy()

buscar_ue = f2.text_input("Buscar Unidad Educativa")

if buscar_ue and "NOMBRE DE LA UNIDAD EDUCATIVA" in base.columns:
    base = base[
        base["NOMBRE DE LA UNIDAD EDUCATIVA"]
        .astype(str)
        .str.contains(buscar_ue, case=False, na=False)
    ].copy()

# ============================================================
# TABLA DE DETALLE
# ============================================================
st.subheader("Detalle de proyectos")

columnas = [
    "DEPARTAMENTO",
    "PROVINCIA",
    "MUNICIPIO",
    "NOMBRE DE LA UNIDAD EDUCATIVA",
    "ESTADO DE PRIORIZACIÓN",
    "INDICE PRIORIZACION",
    "MATRÍCULA",
    "INDICADOR NBI",
    "ESTUDIANTES A 1KM",
    "DENTRO POLIGONO RADIACIÓN",
    "MONTO PROYECTOS PRIORIZADOS"
]

disponibles = [c for c in columnas if c in base.columns]

st.dataframe(
    base[disponibles],
    use_container_width=True,
    hide_index=True
)

st.caption(f"Registros mostrados: {len(base):,}")

st.download_button(
    "⬇️ Descargar detalle en Excel",
    data=excel_bytes(base[disponibles]),
    file_name="detalle_proyectos_tinglados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Ministerio de Educación — BOL-39/2024")
