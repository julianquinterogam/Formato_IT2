from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st

from it2 import (HOJA_15_DEFECTO, MESES, a_formato_final, corregir_tipo_elemento, construir_it2,
                 construir_it2_final, cruzar_con_base, cruzar_localidad_it1, csv_it2, excel_it2,
                 fechas_20_inconsistentes, filtrar_mes, hojas_excel, leer_base, leer_it1, leer_mtto_15,
                 leer_mtto_20, leer_mtto_sytex, unificar_por_nui, valores_inversion)

st.set_page_config(page_title="IT2 Mantenimientos", page_icon="🔧", layout="wide")


@st.cache_data(show_spinner="Leyendo base...")
def _base(contenido: bytes):
    return leer_base(contenido)


@st.cache_data(show_spinner="Leyendo listado 1.5...")
def _m15(contenido: bytes, hoja: str):
    return leer_mtto_15(contenido, hoja)


@st.cache_data(show_spinner="Leyendo valores de inversión...")
def _valores(contenido: bytes, anio: int, mes: int):
    return valores_inversion(contenido, anio, mes)


@st.cache_data(show_spinner="Leyendo listado 2.0...")
def _m20(contenido: bytes):
    return leer_mtto_20(contenido)


@st.cache_data(show_spinner="Leyendo listado SYTEX...")
def _sytex(contenido: bytes):
    return leer_mtto_sytex(contenido)


@st.cache_data(show_spinner="Leyendo archivo IT2...")
def _it2_subido(contenido: bytes):
    return pd.read_excel(BytesIO(contenido), engine="openpyxl")


@st.cache_data(show_spinner="Leyendo IT1...")
def _it1(contenido: bytes):
    return leer_it1(contenido)


def tab_generar():
    st.caption("Carga de archivos, filtro por mes/año, cruce con la base y descarga del formato.")

    # ---------- Período ----------
    c1, c2 = st.columns(2)
    hoy = date.today()
    mes = c1.selectbox("Mes", list(MESES), format_func=MESES.get, index=hoy.month - 1, key="g_mes")
    anio = c2.number_input("Año", min_value=2023, max_value=2100, value=hoy.year, step=1, key="g_anio")

    # ---------- Archivos ----------
    f_base = st.file_uploader("Base (Excel con NIU, cod_localidad, dane, TIPO_UC...)", type=["xlsx"])
    f_15 = st.file_uploader("Listado de mantenimientos 1.5", type=["xls", "xlsx"])
    hoja_15 = None
    if f_15:
        hojas = hojas_excel(f_15.getvalue())
        idx = hojas.index(HOJA_15_DEFECTO) if HOJA_15_DEFECTO in hojas else 0
        hoja_15 = st.selectbox("Hoja del listado 1.5", hojas, index=idx)
    f_20 = st.file_uploader("Listado de mantenimientos 2.0", type=["xls", "xlsx"])
    f_sytex = st.file_uploader("Listado de mantenimientos SYTEX (EJECUTADO_SYTEX)", type=["xls", "xlsx"])
    f_it2 = st.file_uploader("Excel IT2 (hoja 'Valor inversion', para la columna VALOR)", type=["xlsx"])

    if not f_base or not (f_15 or f_20 or f_sytex):
        st.info("Sube la base y al menos un listado de mantenimientos para continuar.")
        return

    base = _base(f_base.getvalue())
    st.write(f"**Base:** {len(base):,} filas · {base['NIU'].nunique():,} NIU únicos")

    # ---------- Filtro por mes y cruce ----------
    resultados = []
    if f_15:
        resultados.append(("1.5", _m15(f_15.getvalue(), hoja_15)))
    if f_20:
        resultados.append(("2.0", _m20(f_20.getvalue())))
    if f_sytex:
        resultados.append(("SYTEX", _sytex(f_sytex.getvalue())))

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

    # ---------- Formato IT2 y descarga ----------
    st.divider()
    st.subheader("Formato IT2")
    mttos = pd.concat(cruzados, ignore_index=True)
    if not mttos.empty:
        mttos, repetidos = unificar_por_nui(mttos)
        if repetidos:
            st.info(f"{repetidos} registros del mismo NUI y del mismo día en el mes; se dejó uno solo por día "
                    "(el primero del listado). Cuando un NUI tiene mantenimientos en días distintos, se conservan todos.")
    if mttos.empty:
        st.warning("No hay mantenimientos que crucen con la base en este mes; no hay nada que descargar.")
        return

    valores = None
    if f_it2:
        try:
            valores = _valores(f_it2.getvalue(), int(anio), mes)
        except ValueError as e:
            st.error(str(e))
    else:
        st.info("Sube el Excel IT2 para que la columna VALOR se llene con el valor del mes seleccionado.")
    it2 = construir_it2(mttos, base, valores)
    if valores is not None and it2["VALOR"].isna().any():
        st.warning(f"{int(it2['VALOR'].isna().sum())} filas quedaron sin VALOR: su TIPO_UC_9995 no está en la hoja 'Valor inversion'.")
    par = it2.attrs.get("resumen_fechas", {})
    if par.get("grupos_en_paralelo"):
        st.info(f"Horas simuladas: en {par['grupos_en_paralelo']} combinaciones técnico-día "
                f"({par['registros_en_paralelo']} mantenimientos) los mantenimientos no caben en una sola jornada "
                "de 08:00 a 17:00 con las duraciones y desplazamientos definidos, así que se repartieron en "
                "cuadrillas que trabajan en paralelo.")
    malas = fechas_20_inconsistentes(mttos)
    if len(malas):
        st.info(f"{len(malas)} mantenimientos del listado 2.0 traían la fecha fin anterior al inicio, igual al "
                "inicio (misma hora redondeada al minuto) o en otro día; se reprogramaron con las mismas reglas "
                "del 1.5 (mismo día de inicio, jornada de 08:00 a 17:00 y duración según el tipo de mantenimiento).")
        with st.expander("Ver esos NUI (fechas originales y nuevas)"):
            nuevas = (it2.drop_duplicates("NUI_MANTENIMIENTO")
                      .set_index("NUI_MANTENIMIENTO")[["FECHA INICIO", "FECHA FIN"]]
                      .rename(columns={"FECHA INICIO": "INICIO NUEVO", "FECHA FIN": "FIN NUEVO"}))
            malas = malas.assign(NUI=pd.to_numeric(malas["NUI"])).rename(
                columns={"FECHA INICIO": "INICIO ORIGINAL", "FECHA FIN": "FIN ORIGINAL"})
            st.dataframe(malas.join(nuevas, on="NUI"), use_container_width=True)

    st.write(f"**{len(it2):,} filas** de **{it2['NUI_MANTENIMIENTO'].nunique():,} NUI** con mantenimiento, ordenadas por NUI.")
    st.dataframe(it2.head(100), use_container_width=True)

    nombre_editable = f"IT2_{mes:02d}_{int(anio)}_editable.xlsx"
    nombre_final = f"IT2_{mes:02d}_{int(anio)}.xlsx"
    it2_final = construir_it2_final(it2)

    c1, c2 = st.columns(2)
    c1.download_button(
        f"Descargar {nombre_editable}",
        data=excel_it2(it2),
        file_name=nombre_editable,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    c2.download_button(
        f"Descargar {nombre_final}",
        data=excel_it2(it2_final),
        file_name=nombre_final,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Deja el editable disponible para la pestaña "IT2 final (con IT1)"
    st.session_state["editable_generado"] = it2
    st.session_state["editable_generado_nombre"] = f"{MESES[mes]} {int(anio)} (recién generado arriba)"


def tab_final():
    st.caption("Toma un IT2 (editable o final), corrige TIPO_UC_9995 = 14 → 13 y reemplaza el Código Localidad "
               "con el que encuentre en el IT1, cruzando por el serial. Genera el IT2 final para descargar.")

    c1, c2 = st.columns(2)
    hoy = date.today()
    mes = c1.selectbox("Mes", list(MESES), format_func=MESES.get, index=hoy.month - 1, key="f_mes")
    anio = c2.number_input("Año", min_value=2023, max_value=2100, value=hoy.year, step=1, key="f_anio")

    origen = None
    generado = st.session_state.get("editable_generado")
    if generado is not None:
        usar_generado = st.checkbox(
            f"Usar el IT2 recién generado en la otra pestaña ({st.session_state.get('editable_generado_nombre')})",
            value=True,
        )
        if usar_generado:
            origen = generado

    f_manual = st.file_uploader("O sube un IT2 manualmente (editable de 17 columnas, o final de 9)",
                                 type=["xlsx"])
    if f_manual:
        origen = _it2_subido(f_manual.getvalue())

    f_it1 = st.file_uploader("IT1 (con las columnas 'serial' y 'cod_localidad')", type=["xlsx"])

    if origen is None or f_it1 is None:
        st.info("Elige un IT2 (el recién generado o uno subido manualmente) y sube el IT1 para continuar.")
        return

    try:
        final = a_formato_final(origen)
    except ValueError as e:
        st.error(str(e))
        return
    it1 = _it1(f_it1.getvalue())

    final, cambios14 = corregir_tipo_elemento(final)
    if cambios14:
        st.info(f"{cambios14} filas tenían Tipo de Elemento = 14 (poste); se corrigieron a 13.")

    final, sin_cruce = cruzar_localidad_it1(final, it1)
    if len(sin_cruce):
        st.warning(f"{len(sin_cruce)} filas con un serial que no aparece en el IT1; se dejó el Código Localidad "
                   "que ya traían.")
        with st.expander("Ver esos seriales"):
            st.dataframe(sin_cruce.rename("Serial del Elemento").drop_duplicates(), use_container_width=True)

    st.write(f"**{len(final):,} filas** en el IT2 final.")
    st.dataframe(final.head(100), use_container_width=True)

    nombre = f"IT2_{mes:02d}_{int(anio)}"
    c1, c2 = st.columns(2)
    c1.download_button(
        f"Descargar {nombre}.xlsx",
        data=excel_it2(final),
        file_name=f"{nombre}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    c2.download_button(
        f"Descargar {nombre}.csv",
        data=csv_it2(final),
        file_name=f"{nombre}.csv",
        mime="text/csv",
    )


st.title("Formato IT2 – Mantenimientos")
tab1, tab2 = st.tabs(["Generar IT2", "IT2 final (con IT1)"])
with tab1:
    tab_generar()
with tab2:
    tab_final()
