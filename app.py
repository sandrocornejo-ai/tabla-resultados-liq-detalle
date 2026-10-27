"""
App Streamlit: Tabla de Resultados (Rex+) -> Archivo de carga Liquidaciones Detalle.

Ejecutar:
    pip install -r requirements.txt
    streamlit run app.py
"""
import io
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import pandas as pd
import streamlit as st

# ─────────────────────────── Lógica de transformación ───────────────────────────

CONCEPTOS_EXCLUIDOS = {"variables_empleado"}

# Conceptos que NO se pueden cargar: el usuario debe crear un concepto nuevo y reemplazarlos
CONCEPTOS_NO_PERMITIDOS = [
    "recargoMi", "remvarMi", "tratoMi", "bonsupMi", "benefMi", "otremMi", "pagosindMi",
    "sobresueldoMi", "comisionMi", "AguinaldoMi", "BonoMi", "participacionMi", "becaMi",
    "otingrMi", "viaticoMi", "gastoMi", "salaCMi", "indemMMi", "indemNMi", "indvolMi",
    "indcontMi", "asigPerdCajaMi", "AsigDesgHerrMi", "AsigTeletrabajoMi", "TraslacionMi",
    "alojamientoTrabajoMi", "imptoindMi", "cuotavivMi", "credMi", "otrdescMi", "CuotaSindMi",
    "otrosDesctoMi", "AnticipoPrestamoMi",
]


def detectar_no_permitidos(df: pd.DataFrame) -> pd.Series:
    """Cantidad de filas por cada concepto no permitido presente en la entrada."""
    sub = df[df["concepto"].isin(CONCEPTOS_NO_PERMITIDOS)]
    return sub["concepto"].value_counts().reindex(
        [c for c in CONCEPTOS_NO_PERMITIDOS if c in set(sub["concepto"])])


def aplicar_reemplazos(df: pd.DataFrame, reemplazos: dict) -> pd.DataFrame:
    """Reemplaza el Id de concepto según {concepto_no_permitido: concepto_nuevo}."""
    if not reemplazos:
        return df
    df = df.copy()
    df["concepto"] = df["concepto"].replace(reemplazos)
    return df

COLUMNAS_ENTERAS = [
    "Número de contrato", "Monto del concepto", "Afecto", "Días de licencias",
    "Días trabajados", "Total de rebajas por LLSS", "Rentas no gravadas",
    "Rebaja por zona extrema", "Días de vacaciones", "Monto Init", "Fase",
    "Parcial 7", "Parcial 8",
]

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
    """Valor entero sin decimales (redondeo al entero más cercano, .5 hacia arriba)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    try:
        return int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError):
        return v


def tasa(v):
    """Vacío si no hay dato o es 0; decimales con coma (1.78 -> '1,78')."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == 0:
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, float):
        return f"{v:g}".replace(".", ",")
    return v


def leer_horas_semanales(fuente) -> dict:
    """Lee 'Horas semanales' y devuelve {(empleado, contrato): horasSema}."""
    crudo = pd.read_excel(fuente, header=None)
    filas = crudo.index[crudo.apply(
        lambda r: r.astype(str).str.strip().eq("horasSema").any(), axis=1)]
    if len(filas) == 0:
        raise ValueError("No se encontró la columna 'horasSema' en el archivo de Horas semanales.")
    if hasattr(fuente, "seek"):
        fuente.seek(0)
    h = pd.read_excel(fuente, header=filas[0])
    h.columns = [str(c).strip() for c in h.columns]
    h["empleado"] = h["empleado"].astype(str).str.strip().str.upper()
    h["contrato"] = pd.to_numeric(h["contrato"], errors="coerce")
    h["horasSema"] = pd.to_numeric(h["horasSema"], errors="coerce")
    return {(e, c): hs for e, c, hs in zip(h["empleado"], h["contrato"], h["horasSema"])}


def jornada(horas):
    """'C' si las horas semanales son mayores a 30, si no 'P'. Vacío si no hay dato."""
    if horas is None or pd.isna(horas):
        return None
    return "C" if horas > 30 else "P"


def dias_trabajados_por_contrato(df: pd.DataFrame) -> dict:
    """Días trabajados = parcial4 de sueldoBase (respaldo: totalesEmpl)."""
    out = {}
    for concepto in ("totalesEmpl", "sueldoBase"):  # sueldoBase sobrescribe
        for _, r in df[df["concepto"] == concepto].iterrows():
            out[(r["empleado"], r["contrato"], r["fechaProc"])] = num(r["parcial4"])
    return out


def transformar(df: pd.DataFrame, horas: dict | None = None) -> pd.DataFrame:
    horas = horas or {}
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
            "Jornada": jornada(horas.get(
                (str(r["empleado"]).strip().upper(), pd.to_numeric(r["contrato"], errors="coerce")))),
            "Días de vacaciones": 0,
            "Monto Init": num(r["montoInit"]),
            "Fase": num(r["fase_id"], 1),
            "Parcial 7": num(r["parcial7"]),
            "Parcial 8": num(r["parcial8"]),
        })
    out = pd.DataFrame(filas, columns=COLUMNAS_SALIDA)
    # Todas las columnas numéricas como enteros (sin decimales), salvo Cotización de jubilación
    for c in COLUMNAS_ENTERAS:
        out[c] = pd.to_numeric(out[c], errors="coerce").round().astype("Int64")
    return out


def a_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Liquidaciones", index=False)
        ws = w.sheets["Liquidaciones"]
        for j, c in enumerate(COLUMNAS_SALIDA, start=1):
            if c in COLUMNAS_ENTERAS:
                for fila in range(2, ws.max_row + 1):
                    ws.cell(row=fila, column=j).number_format = "0"
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
st.caption("Carga la Tabla de Resultados de Rex+ y el archivo de Horas semanales, y descarga el archivo de carga de liquidaciones.")

col_a, col_b = st.columns(2)
archivo = col_a.file_uploader("1. Tabla de Resultados (.xlsx)", type=["xlsx", "xls"])
archivo_horas = col_b.file_uploader("2. Horas semanales (.xlsx)", type=["xlsx", "xls"])

if archivo is not None and archivo_horas is None:
    st.warning("Falta cargar el archivo de Horas semanales para definir la Jornada (C > 30 hrs, P ≤ 30 hrs).")

if archivo is not None and archivo_horas is not None:
    try:
        entrada = leer_tabla_resultados(archivo.getvalue())
        horas = leer_horas_semanales(io.BytesIO(archivo_horas.getvalue()))
    except Exception as e:
        st.error(f"No se pudo procesar el archivo: {e}")
        st.stop()

    # ── Validación de conceptos no permitidos ──
    id_archivo = f"{archivo.name}-{archivo.size}"
    if st.session_state.get("archivo_actual") != id_archivo:
        st.session_state["archivo_actual"] = id_archivo
        st.session_state["reemplazos"] = {}
    reemplazos = st.session_state["reemplazos"]

    no_permitidos = detectar_no_permitidos(entrada)
    pendientes = [c for c in no_permitidos.index if c not in reemplazos]

    if len(no_permitidos):
        st.subheader("Conceptos no permitidos")
        if pendientes:
            st.error(f"El archivo trae {len(no_permitidos)} concepto(s) que no se pueden cargar. "
                     "Crea un concepto nuevo en Rex+ para cada uno, escribe su ID y presiona "
                     "**Aplicar cambio**.")
        else:
            st.success("Todos los conceptos no permitidos fueron reemplazados.")

        h1c, h2c, h3c, h4c = st.columns([3, 1.3, 3, 1.6])
        h1c.markdown("**Concepto no permitido**")
        h2c.markdown("**Filas**")
        h3c.markdown("**ID del concepto nuevo**")
        for concepto, n in no_permitidos.items():
            c_1, c_2, c_3, c_4 = st.columns([3, 1.3, 3, 1.6])
            c_1.markdown(f"`{concepto}`")
            c_2.markdown(f"{n}")
            nuevo = c_3.text_input("ID nuevo", key=f"nuevo_{concepto}",
                                   value=reemplazos.get(concepto, ""),
                                   label_visibility="collapsed", placeholder="Ej: AguinaldoNuevo")
            if c_4.button("Aplicar cambio", key=f"aplicar_{concepto}"):
                nuevo = (nuevo or "").strip()
                if not nuevo:
                    c_3.error("Escribe el ID del concepto nuevo.")
                elif nuevo in CONCEPTOS_NO_PERMITIDOS:
                    c_3.error(f"'{nuevo}' también es un concepto no permitido.")
                else:
                    reemplazos[concepto] = nuevo
                    st.rerun()
            if concepto in reemplazos:
                c_4.markdown(f"✅ → `{reemplazos[concepto]}`")

    try:
        salida = transformar(aplicar_reemplazos(entrada, reemplazos), horas)
    except Exception as e:
        st.error(f"No se pudo procesar el archivo: {e}")
        st.stop()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Filas leídas", len(entrada))
    c2.metric("Filas generadas", len(salida))
    c3.metric("Empleados", salida["Id empleado"].nunique())
    c4.metric("Períodos", ", ".join(sorted(salida["Fecha de proceso"].astype(str).unique())))

    contratos = salida.drop_duplicates(["Id empleado", "Número de contrato"])
    j1, j2, j3 = st.columns(3)
    j1.metric("Contratos jornada C", int((contratos["Jornada"] == "C").sum()))
    j2.metric("Contratos jornada P", int((contratos["Jornada"] == "P").sum()))
    sin_horas = contratos[contratos["Jornada"].isna()]
    j3.metric("Sin horas semanales", len(sin_horas))
    if len(sin_horas):
        st.warning("Estos empleados/contratos no están en el archivo de Horas semanales; "
                   "su Jornada queda en blanco:")
        st.dataframe(sin_horas[["Id empleado", "Número de contrato"]], hide_index=True)

    excluidas = len(entrada) - len(salida)
    if excluidas:
        st.info(f"Se excluyeron {excluidas} fila(s) de conceptos no cargables: "
                f"{', '.join(sorted(CONCEPTOS_EXCLUIDOS))}.")

    st.subheader("Vista previa")
    vista = salida.copy()
    vista["Cotización de jubilación"] = vista["Cotización de jubilación"].map(
        lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
    st.dataframe(vista, width="stretch", hide_index=True)

    if pendientes:
        st.warning("La descarga está bloqueada hasta reemplazar todos los conceptos no permitidos: "
                   + ", ".join(pendientes))
    else:
        periodo = "_".join(sorted(salida["Fecha de proceso"].astype(str).unique()))
        st.download_button(
            "⬇️ Descargar archivo de carga",
            data=a_excel(salida),
            file_name=f"liq_detalle_{periodo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
