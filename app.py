from datetime import date

import streamlit as st

from it2 import (HOJA_15_DEFECTO, MESES, cruzar_con_base, filtrar_mes, hojas_excel,
                       leer_base, leer_mtto_15, leer_mtto_20)

st.set_page_config(page_title="IT2 Mantenimientos", page_icon="🔧", layout="wide")

st.title("Formato IT2 – Mantenimientos")
st.caption("Paso 1: carga de archivos, filtro por mes/año y cruce con la base.")


@st.cache_data(show_spinner="Leyendo base...")
def _base(contenido: bytes):
    return leer_base(contenido)


@st.cache_data(show_spinner="Leyendo listado 1.5...")
def _m15(contenido: bytes, hoja: str):
    return leer_mtto_15(contenido, hoja)


@st.cache_data(show_spinner="Leyendo listado 2.0...")
def _m20(contenido: bytes):
    return leer_mtto_20(contenido)


# ---------- Período ----------
c1, c2 = st.columns(2)
hoy = date.today()
mes = c1.selectbox("Mes", list(MESES), format_func=MESES.get, index=hoy.month - 1)
anio = c2.number_input("Año", min_value=2023, max_value=2100, value=hoy.year, step=1)

# ---------- Archivos ----------
f_base = st.file_uploader("Base (Excel con NIU, cod_localidad, dane, TIPO_UC...)", type=["xlsx"])
f_15 = st.file_uploader("Listado de mantenimientos 1.5", type=["xls", "xlsx"])
hoja_15 = None
if f_15:
    hojas = hojas_excel(f_15.getvalue(), "xlrd" if f_15.name.lower().endswith(".xls") else None)
    idx = hojas.index(HOJA_15_DEFECTO) if HOJA_15_DEFECTO in hojas else 0
    hoja_15 = st.selectbox("Hoja del listado 1.5", hojas, index=idx)
f_20 = st.file_uploader("Listado de mantenimientos 2.0", type=["xls", "xlsx"])

if not f_base or not (f_15 or f_20):
    st.info("Sube la base y al menos un listado de mantenimientos para continuar.")
    st.stop()

base = _base(f_base.getvalue())
st.write(f"**Base:** {len(base):,} filas · {base['NIU'].nunique():,} NIU únicos")

# ---------- Filtro por mes y cruce ----------
resultados = []
if f_15:
    resultados.append(("1.5", _m15(f_15.getvalue(), hoja_15)))
if f_20:
    resultados.append(("2.0", _m20(f_20.getvalue())))

st.subheader(f"{MESES[mes]} {int(anio)}")
cruzados, descartados = [], []
for version, df in resultados:
    del_mes = filtrar_mes(df, int(anio), mes)
    ok, no = cruzar_con_base(del_mes, base)
    cruzados.append(ok)
    descartados.append(no)
    a, b, c, d = st.columns(4)
    a.metric(f"{version} · del mes", f"{len(del_mes):,}")
    b.metric("Cruzan con la base", f"{len(ok):,}")
    c.metric("Descartados", f"{len(no):,}")
    d.metric("NUI únicos", f"{ok['NUI_NORM'].nunique():,}")

with st.expander("Vista previa: mantenimientos que cruzan"):
    for (version, _), ok in zip(resultados, cruzados):
        st.write(f"Versión {version}")
        st.dataframe(ok[["NUI_NORM", "FECHA_REF"]].head(50), use_container_width=True)

with st.expander("Vista previa: descartados (NUI que no están en la base)"):
    for (version, _), no in zip(resultados, descartados):
        st.write(f"Versión {version}")
        st.dataframe(no[["NUI_NORM", "FECHA_REF"]].head(50), use_container_width=True)
