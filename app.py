"""
App Streamlit: Tabla de Resultados (Rex+) -> Archivo de carga Liquidaciones Detalle.

Ejecutar:
    pip install -r requirements.txt
    streamlit run app.py
"""
import io
import pandas as pd
import streamlit as st

# ─────────────────────────── Lógica de transformación ───────────────────────────

CONCEPTOS_EXCLUIDOS = {"variables_empleado"}

COLUMNAS_SALIDA = [
    "Fecha de proceso", "Id empleado", "Número de contrato", "Id del concepto",
    "Monto del concepto", "Afecto", "Id de institución", "Cotización de jubilación",
    "Días de licencias", "Días trabajados", "Fecha de aplicación", "Empresa",
    "Total de rebajas por LLSS", "Rentas no gravadas", "Rebaja por zona extrema",
    "Jornada", "Días de vacaciones", "Monto Init", "Fase", "Parcial 7", "Parcial 8",
]

COLUMNAS_REQUERIDAS = [
    "fechaProc", "empleado", "contrato", "concepto", "monto", "baseCalc",
    "institucion", "parcial1", "parcial2", "parcial3", "parcial4", "parcial5",
    "parcial7", "parcial8", "fechaAplic", "empresa", "montoInit", "fase_id",
]


def leer_tabla_resultados(contenido: bytes) -> pd.DataFrame:
    """Lee el Excel ubicando la fila de encabezados (la que contiene 'concepto')."""
    crudo = pd.read_excel(io.BytesIO(contenido), header=None)
    filas = crudo.index[crudo.apply(
        lambda r: r.astype(str).str.strip().eq("concepto").any(), axis=1)]
    if len(filas) == 0:
        raise ValueError("No se encontró la fila de encabezados (columna 'concepto').")
    df = pd.read_excel(io.BytesIO(contenido), header=filas[0])
    df.columns = [str(c).strip() for c in df.columns]
    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en la Tabla de Resultados: {', '.join(faltantes)}")
    for c in df.select_dtypes(include=["object", "string"]).columns:
        df[c] = df[c].apply(lambda v: v.strip() if isinstance(v, str) else v)
        df[c] = df[c].replace("", None)
    df = df[df["concepto"].notna()]
    return df


def num(v, default=0):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def tasa(v):
    """Vacío si no hay dato o es 0; decimales con coma (1.78 -> '1,78')."""
    v = num(v, None)
    if v is None or v == 0:
        return None
    if isinstance(v, float):
        return f"{v:g}".replace(".", ",")
    return v


def dias_trabajados_por_contrato(df: pd.DataFrame) -> dict:
    """Días trabajados = parcial4 de sueldoBase (respaldo: totalesEmpl)."""
    out = {}
    for concepto in ("totalesEmpl", "sueldoBase"):  # sueldoBase sobrescribe
        for _, r in df[df["concepto"] == concepto].iterrows():
            out[(r["empleado"], r["contrato"], r["fechaProc"])] = num(r["parcial4"])
    return out


def transformar(df: pd.DataFrame) -> pd.DataFrame:
    df = df[~df["concepto"].isin(CONCEPTOS_EXCLUIDOS)].copy()
    dias_trab = dias_trabajados_por_contrato(df)
    filas = []
    for _, r in df.iterrows():
        clave = (r["empleado"], r["contrato"], r["fechaProc"])
        es_impuesto = r["concepto"] == "impuesto"
        filas.append({
            "Fecha de proceso": r["fechaProc"],
            "Id empleado": r["empleado"],
            "Número de contrato": num(r["contrato"], None),
            "Id del concepto": r["concepto"],
            "Monto del concepto": num(r["monto"]),
            "Afecto": num(r["baseCalc"]),
            "Id de institución": r["institucion"] if not pd.isna(r["institucion"]) else None,
            "Cotización de jubilación": tasa(r["parcial1"]),
            "Días de licencias": num(r["parcial5"]),
            "Días trabajados": dias_trab.get(clave, 0),
            "Fecha de aplicación": r["fechaAplic"],
            "Empresa": r["empresa"],
            "Total de rebajas por LLSS": num(r["parcial2"]) if es_impuesto else 0,
            "Rentas no gravadas": num(r["parcial3"]) if es_impuesto else 0,
            "Rebaja por zona extrema": 0,
            "Jornada": None,
            "Días de vacaciones": 0,
            "Monto Init": num(r["montoInit"]),
            "Fase": num(r["fase_id"], 1),
            "Parcial 7": num(r["parcial7"]),
            "Parcial 8": num(r["parcial8"]),
        })
    return pd.DataFrame(filas, columns=COLUMNAS_SALIDA)


def a_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Liquidaciones", index=False)
    return buf.getvalue()


# ─────────────────────────────────── Interfaz ───────────────────────────────────

st.set_page_config(page_title="Tabla de Resultados → Liq. Detalle", page_icon="📄", layout="wide")
st.markdown("""
<style>
.stApp { background-color: #F0F4FA; }
h1, h2, h3 { color: #1565C0; }
.stDownloadButton button, .stButton button {
    background-color: #1565C0; color: white; border: none; border-radius: 6px;
}
.stDownloadButton button:hover, .stButton button:hover { background-color: #0D47A1; color: white; }
</style>
""", unsafe_allow_html=True)

st.title("Tabla de Resultados → Liquidaciones Detalle")
st.caption("Carga la consulta de la Tabla de Resultados de Rex+ y descarga el archivo de carga de liquidaciones.")

archivo = st.file_uploader("Tabla de Resultados (.xlsx)", type=["xlsx", "xls"])

if archivo is not None:
    try:
        entrada = leer_tabla_resultados(archivo.getvalue())
        salida = transformar(entrada)
    except Exception as e:
        st.error(f"No se pudo procesar el archivo: {e}")
        st.stop()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Filas leídas", len(entrada))
    c2.metric("Filas generadas", len(salida))
    c3.metric("Empleados", salida["Id empleado"].nunique())
    c4.metric("Períodos", ", ".join(sorted(salida["Fecha de proceso"].astype(str).unique())))

    excluidas = len(entrada) - len(salida)
    if excluidas:
        st.info(f"Se excluyeron {excluidas} fila(s) de conceptos no cargables: "
                f"{', '.join(sorted(CONCEPTOS_EXCLUIDOS))}.")

    st.subheader("Vista previa")
    st.dataframe(salida, use_container_width=True, hide_index=True)

    periodo = "_".join(sorted(salida["Fecha de proceso"].astype(str).unique()))
    st.download_button(
        "⬇️ Descargar archivo de carga",
        data=a_excel(salida),
        file_name=f"liq_detalle_{periodo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
