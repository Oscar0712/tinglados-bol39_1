import streamlit as st
import pandas as pd
import unicodedata
import re
from io import BytesIO
from pathlib import Path

st.set_page_config(page_title="BOL-39/2024 | Consulta de Proyectos", page_icon="🏫", layout="wide")
st.title("🏫 BOL-39/2024 — Consulta de Proyectos")
st.caption("Consulta territorial y detalle de proyectos de tinglados")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem;}
div[data-testid="stMetric"] {
    border: 1px solid rgba(128,128,128,.25);
    border-radius: 12px;
    padding: 10px;
}
.mobile-card {
    border: 1px solid rgba(128,128,128,.28);
    border-radius: 14px;
    padding: 14px 16px;
    margin-bottom: 12px;
    background: rgba(128,128,128,.04);
}
.mobile-card h4 {margin: 0 0 8px 0; font-size: 1.05rem;}
.mobile-card p {margin: 3px 0; line-height: 1.35;}
@media (max-width: 768px) {
    .block-container {padding-left: .75rem; padding-right: .75rem; padding-top: .6rem;}
    h1 {font-size: 1.5rem !important; line-height: 1.15 !important;}
    h2, h3 {font-size: 1.15rem !important;}
    div[data-testid="stMetric"] {padding: 8px;}
    div[data-testid="stMetricValue"] {font-size: 1.1rem;}
}
</style>
""", unsafe_allow_html=True)


ARCHIVO_BASE = "Tabla_resumen_Tinglados_v5.xlsx"

def norm(s):
    s = str(s).strip().upper()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return re.sub(r'[_\\s]+', ' ', s)

ALIASES = {
    "DEPARTAMENTO": ["DEPARTAMENTO"],
    "PROVINCIA": ["PROVINCIA"],
    "MUNICIPIO": ["MUNICIPIO"],
    "CÓDIGO RUE": ["CÓDIGO RUE", "CODIGO RUE", "COD RUE", "RUE"],
    "NOMBRE DE LA UNIDAD EDUCATIVA": ["NOMBRE DE LA UNIDAD EDUCATIVA", "UNIDAD EDUCATIVA", "NOMBRE UE"],
    "ESTADO DE PRIORIZACIÓN": ["ESTADO DE PRIORIZACIÓN", "ESTADO DE PRIORIZACION", "ESTADO SELECCION", "ESTADO_SELECCION"],
    "INDICE PRIORIZACION": ["INDICE PRIORIZACION", "ÍNDICE PRIORIZACIÓN", "INDICE_PRIORIZACION"],
    "MATRÍCULA": ["MATRÍCULA", "MATRICULA", "INSCRITO SUM", "INSCRITO_SUM"],
    "INDICADOR NBI": ["INDICADOR NBI", "INDICADOR_NBI"],
    "ESTUDIANTES A 1KM": ["ESTUDIANTES A 1KM", "ESTUDIANTES 1KM", "ESTUDIANTES_1KM"],
    "DENTRO POLIGONO RADIACIÓN": ["DENTRO POLIGONO RADIACIÓN", "DENTRO POLIGONO RADIACION", "DENTRO_POLIGONO_RADIACION"],
    "PROYECTOS PRESENTADOS": ["PROYECTOS PRESENTADOS", "PRESENTADOS"],
    "ENVIADOS A FPS": ["ENVIADOS A FPS", "ENVIADOS FPS"],
    "APROBADOS POR FPS": ["APROBADOS POR FPS", "APROBADOS FPS"],
    "PROYECTOS PRIORIZADOS": ["PROYECTOS PRIORIZADOS", "SELECCIONADOS"],
    "MONTO PROYECTOS PRIORIZADOS": ["MONTO PROYECTOS PRIORIZADOS", "MONTO PRIORIZADO SELECCIONADO", "MONTO PRIORIZADO_SELECCIONADO"]
}

def standardize_columns(df):
    normalized = {norm(c): c for c in df.columns}
    rename = {}
    for target, candidates in ALIASES.items():
        for candidate in candidates:
            if norm(candidate) in normalized:
                rename[normalized[norm(candidate)]] = target
                break
    return df.rename(columns=rename)

def excel_bytes(df, sheet_name):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return bio.getvalue()

@st.cache_data
def cargar_base(ruta):
    return pd.read_excel(ruta)

ruta_base = Path(__file__).parent / ARCHIVO_BASE
if not ruta_base.exists():
    st.error(f"No se encontró '{ARCHIVO_BASE}' en el repositorio. Súbalo a GitHub en la misma carpeta que app.py.")
    st.stop()

try:
    df = cargar_base(ruta_base)
except Exception as e:
    st.error(f"No se pudo leer la base: {e}")
    st.stop()

df = standardize_columns(df)

# AJUSTE 1: todo registro sin Priorizado/No priorizado pasa a No cumple
if "ESTADO DE PRIORIZACIÓN" not in df.columns:
    df["ESTADO DE PRIORIZACIÓN"] = "No cumple"
else:
    estado = df["ESTADO DE PRIORIZACIÓN"].fillna("").astype(str).str.strip()
    estado_norm = estado.str.upper().str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8")
    pri = estado_norm.eq("PRIORIZADO")
    npri = estado_norm.isin(["NO PRIORIZADO", "NO PRIORIZADA"])
    df.loc[pri, "ESTADO DE PRIORIZACIÓN"] = "Priorizado"
    df.loc[npri, "ESTADO DE PRIORIZACIÓN"] = "No priorizado"
    df.loc[~(pri | npri), "ESTADO DE PRIORIZACIÓN"] = "No cumple"

territorial = ["DEPARTAMENTO", "PROVINCIA", "MUNICIPIO"]
faltan = [c for c in territorial if c not in df.columns]
if faltan:
    st.error("Faltan campos territoriales: " + ", ".join(faltan))
    st.write("Columnas encontradas:", list(df.columns))
    st.stop()

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

kpis = ["PROYECTOS PRESENTADOS", "ENVIADOS A FPS", "APROBADOS POR FPS", "PROYECTOS PRIORIZADOS", "MONTO PROYECTOS PRIORIZADOS"]
for c in kpis:
    if c not in base.columns:
        base[c] = 0
    base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Presentados", f"{base['PROYECTOS PRESENTADOS'].sum():,.0f}")
k2.metric("Enviados FPS", f"{base['ENVIADOS A FPS'].sum():,.0f}")
k3.metric("Aprobados FPS", f"{base['APROBADOS POR FPS'].sum():,.0f}")
k4.metric("Priorizados", f"{base['PROYECTOS PRIORIZADOS'].sum():,.0f}")
k5.metric("Monto priorizado (Bs)", f"{base['MONTO PROYECTOS PRIORIZADOS'].sum():,.2f}")

tab1, tab2 = st.tabs(["📊 Resumen territorial", "📋 Detalle de proyectos"])

with tab1:
    g = ["DEPARTAMENTO", "PROVINCIA", "MUNICIPIO"]
    a = ["PROYECTOS PRESENTADOS", "ENVIADOS A FPS", "APROBADOS POR FPS", "PROYECTOS PRIORIZADOS", "MONTO PROYECTOS PRIORIZADOS"]
    resumen = base.groupby(g, dropna=False)[a].sum().reset_index()
    st.dataframe(resumen, use_container_width=True, hide_index=True)
    st.download_button("⬇️ Descargar resumen en Excel", excel_bytes(resumen, "Resumen"),
                       "reporte_resumen_tinglados.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab2:
    detail = base.copy()

    # En el detalle se muestran EXCLUSIVAMENTE proyectos
    # con estado "Priorizado" o "No priorizado".
    estado_detalle = (
        detail["ESTADO DE PRIORIZACIÓN"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )
    detail = detail[
        estado_detalle.isin(["priorizado", "no priorizado"])
    ].copy()

    f1, f2, f3 = st.columns(3)

    estados = [e for e in ["Priorizado", "No priorizado"]
               if e in detail["ESTADO DE PRIORIZACIÓN"].astype(str).unique()]
    ef = f1.selectbox("Estado de priorización", ["Todos"] + estados)
    if ef != "Todos":
        detail = detail[detail["ESTADO DE PRIORIZACIÓN"].astype(str) == ef]

    if "NOMBRE DE LA UNIDAD EDUCATIVA" in detail.columns:
        ue = f2.text_input("Buscar Unidad Educativa")
        if ue:
            detail = detail[detail["NOMBRE DE LA UNIDAD EDUCATIVA"].astype(str).str.contains(ue, case=False, na=False)]

    cols = ["DEPARTAMENTO", "PROVINCIA", "MUNICIPIO",
            "NOMBRE DE LA UNIDAD EDUCATIVA", "ESTADO DE PRIORIZACIÓN",
            "INDICE PRIORIZACION", "MATRÍCULA", "INDICADOR NBI",
            "ESTUDIANTES A 1KM", "DENTRO POLIGONO RADIACIÓN",
            "MONTO PROYECTOS PRIORIZADOS"]
    disponibles = [c for c in cols if c in detail.columns]
    st.dataframe(detail[disponibles], use_container_width=True, hide_index=True)
    st.markdown("#### 📱 Vista para celular")
    st.caption("Fichas compactas para una lectura más cómoda desde el teléfono.")

    def formato(v, dec=None):
        if pd.isna(v) or v == "":
            return "—"
        if dec is not None:
            try:
                return f"{float(v):,.{dec}f}"
            except Exception:
                pass
        return str(v)

    for _, row in detail[disponibles].head(100).iterrows():
        st.markdown(f"""
        <div class="mobile-card">
          <h4>{formato(row.get("NOMBRE DE LA UNIDAD EDUCATIVA", "—"))}</h4>
          <p><b>Municipio:</b> {formato(row.get("MUNICIPIO", "—"))}</p>
          <p><b>Estado:</b> {formato(row.get("ESTADO DE PRIORIZACIÓN", "—"))}</p>
          <p><b>Índice de priorización:</b> {formato(row.get("INDICE PRIORIZACION", "—"), 3)}</p>
          <p><b>Matrícula:</b> {formato(row.get("MATRÍCULA", "—"), 0)}</p>
          <p><b>Indicador NBI:</b> {formato(row.get("INDICADOR NBI", "—"), 3)}</p>
          <p><b>Estudiantes a 1 km:</b> {formato(row.get("ESTUDIANTES A 1KM", "—"), 0)}</p>
          <p><b>Dentro polígono radiación:</b> {formato(row.get("DENTRO POLIGONO RADIACIÓN", "—"))}</p>
          <p><b>Monto priorizado:</b> Bs {formato(row.get("MONTO PROYECTOS PRIORIZADOS", "—"), 2)}</p>
        </div>
        """, unsafe_allow_html=True)

    if len(detail) > 100:
        st.info("Se muestran las primeras 100 fichas. Utilice los filtros para reducir los resultados.")

    st.download_button("⬇️ Descargar detalle en Excel", excel_bytes(detail[disponibles], "Detalle"),
                       "reporte_detalle_tinglados.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Ministerio de Educación — herramienta de consulta BOL-39/2024")
