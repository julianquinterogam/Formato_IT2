"""Lógica del formato IT2 Mantenimientos (Paso 1: carga, filtro por mes y cruce con la base)."""
import re
import zlib
from io import BytesIO

import numpy as np
import pandas as pd

MESES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
         7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

HOJA_15_DEFECTO = "mttos 23_24_25"


def normalizar_nui(valor):
    """'44430-1027' -> '444301027'; 444301014.0 -> '444301014'; vacío -> None."""
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    solo_digitos = re.sub(r"\D", "", texto)
    return solo_digitos or None


def _motor(contenido: bytes) -> str:
    """.xlsx es un zip (empieza con 'PK') -> openpyxl; .xls antiguo (OLE2) -> xlrd."""
    return "openpyxl" if contenido[:2] == b"PK" else "xlrd"


def hojas_excel(contenido: bytes) -> list[str]:
    return pd.ExcelFile(BytesIO(contenido), engine=_motor(contenido)).sheet_names


def leer_base(contenido: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(contenido), engine="openpyxl")
    df["NIU"] = df["NIU"].map(normalizar_nui)
    return df


def leer_mtto_15(contenido: bytes, hoja: str) -> pd.DataFrame:
    """Listado 1.5: el NUI viene en 'NUI' y la fecha (sin hora) en 'fecha'."""
    df = pd.read_excel(BytesIO(contenido), sheet_name=hoja, engine=_motor(contenido))
    df["NUI_NORM"] = df["NUI"].map(normalizar_nui)
    df["FECHA_REF"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["VERSION"] = "1.5"
    return df


def leer_mtto_20(contenido: bytes) -> pd.DataFrame:
    """Listado 2.0: el NUI viene en 'Responsable' y la fecha con hora en 'Fecha_Inicio'."""
    df = pd.read_excel(BytesIO(contenido), engine=_motor(contenido))
    df["NUI_NORM"] = df["Responsable"].map(normalizar_nui)
    df["FECHA_REF"] = pd.to_datetime(df["Fecha_Inicio"], errors="coerce")
    df["VERSION"] = "2.0"
    return df


def filtrar_mes(df: pd.DataFrame, anio: int, mes: int) -> pd.DataFrame:
    """Deja únicamente los registros cuya fecha cae en el mes y año seleccionados."""
    f = df["FECHA_REF"]
    return df[(f.dt.year == anio) & (f.dt.month == mes)].copy()


def cruzar_con_base(mtto: pd.DataFrame, base: pd.DataFrame):
    """Separa los mantenimientos cuyo NUI existe como NIU en la base de los que no."""
    nius = set(base["NIU"].dropna())
    coincide = mtto["NUI_NORM"].isin(nius)
    return mtto[coincide].copy(), mtto[~coincide].copy()


# ---------------------------------------------------------------------------
# Paso 2: armado del formato IT2 (columnas de la base, homologaciones, tipo de mantenimiento)
# ---------------------------------------------------------------------------
COLUMNAS_IT2 = [
    "NIU", "COD_LOCALIDAD", "DANE", "TIPO_UC", "TIPO_UC_9995", "CAPACIDAD_UC", "MARCA",
    "SERIAL_INTERNO", "NUI_MANTENIMIENTO", "TIPO DE MANTENIMIENTO", "DECO TIPO MANTENIMIENTO",
    "MANTENIMIENTO REALIZADO", "FECHA INICIO", "FECHA FIN", "ESTADO", "DECO ESTADO", "VALOR",
]

# TIPO_UC -> TIPO_UC_9995
HOMOLOGACION_TIPO_UC = {1: 5, 2: 6, 3: 7, 4: 10, 5: 14, 6: 21, 7: 21, 8: 21, 9: 21, 10: 21, 11: 19}
DECO_TIPO_MTTO = {"Preventivo": 1, "Correctivo": 2}


def tipo_mantenimiento(df: pd.DataFrame) -> pd.Series:
    """1.5: columna Maintenance_Type. 2.0: columnas Preventivo / Correctivo marcadas con True."""
    es_15 = df["VERSION"] == "1.5"
    tipo_15 = df["Maintenance_Type"].astype("string").str.strip().str.capitalize() if "Maintenance_Type" in df else None
    prev = df["Preventivo"].eq(True) if "Preventivo" in df else pd.Series(False, index=df.index)
    corr = df["Correctivo"].eq(True) if "Correctivo" in df else pd.Series(False, index=df.index)
    tipo_20 = pd.Series(pd.NA, index=df.index, dtype="string")
    tipo_20[corr] = "Correctivo"
    tipo_20[prev] = "Preventivo"  # si vinieran ambos marcados, prevalece Preventivo
    if tipo_15 is None:
        return tipo_20
    return tipo_15.where(es_15, tipo_20)


def unificar_por_nui(mttos: pd.DataFrame):
    """Un solo mantenimiento por NUI en el mes. Se prefiere un registro que tenga tipo de mantenimiento
    y, entre esos, el más antiguo (a igual fecha, el primero del listado).
    Devuelve (mantenimientos_unicos, cantidad_de_registros_descartados)."""
    ordenado = (mttos.assign(_sin_tipo=tipo_mantenimiento(mttos).isna())
                .sort_values(["_sin_tipo", "FECHA_REF"], kind="stable"))
    unicos = ordenado.drop_duplicates(subset="NUI_NORM", keep="first").drop(columns="_sin_tipo")
    return unicos.reset_index(drop=True), len(mttos) - len(unicos)


# --- Fechas ---------------------------------------------------------------
# Listado 2.0: se toman tal cual (Fecha_Inicio / Fecha_Finalizacion, sin segundos).
# Listado 1.5: solo trae la fecha; las horas se simulan con estas reglas.
DURACION_MIN = {"Preventivo": (90, 120), "Correctivo": (120, 180)}  # minutos
VIAJE_MIN = (30, 60)                                               # minutos entre mantenimientos
VENTANA_MIN = 9 * 60                                               # jornada 08:00 - 17:00
HORA_APERTURA = pd.Timedelta(hours=8)


def _minimo_jornada(tipos) -> int:
    return sum(DURACION_MIN[t][0] for t in tipos) + VIAJE_MIN[0] * (len(tipos) - 1)


def _programar_cuadrilla(tipos, rng):
    """Horas (minutos desde las 08:00) de una cuadrilla que hace los mantenimientos en secuencia."""
    n = len(tipos)
    dur = [int(rng.integers(DURACION_MIN[t][0], DURACION_MIN[t][1] + 1)) for t in tipos]
    viaje = [int(rng.integers(VIAJE_MIN[0], VIAJE_MIN[1] + 1)) for _ in range(n - 1)]
    total = sum(dur) + sum(viaje)
    if total > VENTANA_MIN:  # jornada muy cargada: se acerca a los mínimos hasta que quepa
        minimo = _minimo_jornada(tipos)
        f = (total - VENTANA_MIN) / (total - minimo)
        dur = [DURACION_MIN[t][0] + int((d - DURACION_MIN[t][0]) * (1 - f)) for t, d in zip(tipos, dur)]
        viaje = [VIAJE_MIN[0] + int((v - VIAJE_MIN[0]) * (1 - f)) for v in viaje]
        total = sum(dur) + sum(viaje)
    t = int(rng.integers(0, VENTANA_MIN - total + 1))  # hora de arranque aleatoria dentro de la holgura
    tramos = []
    for i in range(n):
        tramos.append((t, t + dur[i]))
        t += dur[i] + (viaje[i] if i < n - 1 else 0)
    return tramos


def _programar_grupo(grupo: pd.DataFrame, dia: pd.Timestamp, semilla: str):
    """Mantenimientos 1.5 de un mismo técnico y día. Si no caben en una sola jornada respetando
    duraciones y desplazamientos, se reparten en cuadrillas que trabajan en paralelo."""
    rng = np.random.default_rng(zlib.crc32(semilla.encode()))
    g = grupo.assign(
        _vereda=grupo["vereda"].fillna("").astype(str) if "vereda" in grupo else "",
        _ts=pd.to_numeric(grupo["Id_Encuesta"].astype(str).str.split("-").str[-1], errors="coerce").fillna(0),
    ).sort_values(["_vereda", "_ts"], kind="stable")  # misma vereda junta, y en el orden real de registro
    tipos = [t if t in DURACION_MIN else "Preventivo" for t in tipo_mantenimiento(g)]
    n = len(tipos)
    k = 1
    while True:
        partes = np.array_split(np.arange(n), k)
        if all(_minimo_jornada([tipos[i] for i in p]) <= VENTANA_MIN for p in partes):
            break
        k += 1
    inicio, fin = {}, {}
    for p in partes:
        for idx, (a, b) in zip(p, _programar_cuadrilla([tipos[i] for i in p], rng)):
            etiqueta = g.index[idx]
            inicio[etiqueta] = dia + HORA_APERTURA + pd.Timedelta(minutes=a)
            fin[etiqueta] = dia + HORA_APERTURA + pd.Timedelta(minutes=b)
    return inicio, fin, k


def asignar_fechas(mttos: pd.DataFrame):
    """Devuelve (inicio, fin, resumen). Reproducible: el mismo mes siempre genera las mismas horas."""
    inicio = pd.Series(pd.NaT, index=mttos.index, dtype="datetime64[ns]")
    fin = pd.Series(pd.NaT, index=mttos.index, dtype="datetime64[ns]")
    es20 = mttos["VERSION"] == "2.0"
    if es20.any():
        inicio[es20] = pd.to_datetime(mttos.loc[es20, "Fecha_Inicio"]).dt.floor("min")
        fin[es20] = pd.to_datetime(mttos.loc[es20, "Fecha_Finalizacion"]).dt.floor("min")
    grupos_paralelo = registros_paralelo = 0
    m15 = mttos[~es20]
    if len(m15):
        dias = m15["FECHA_REF"].dt.normalize()
        for (usuario, dia), grupo in m15.groupby([m15["UserName"].fillna(""), dias], sort=False):
            ini_g, fin_g, k = _programar_grupo(grupo, dia, f"{usuario}|{dia:%Y-%m-%d}")
            for etiqueta in ini_g:
                inicio[etiqueta], fin[etiqueta] = ini_g[etiqueta], fin_g[etiqueta]
            if k > 1:
                grupos_paralelo += 1
                registros_paralelo += len(grupo)
    resumen = {"grupos_en_paralelo": grupos_paralelo, "registros_en_paralelo": registros_paralelo}
    return inicio, fin, resumen


def fechas_20_inconsistentes(mttos: pd.DataFrame) -> pd.DataFrame:
    """Registros 2.0 cuya fecha fin es anterior al inicio o cae en otro día."""
    m = mttos[mttos["VERSION"] == "2.0"]
    ini = pd.to_datetime(m["Fecha_Inicio"]); fin = pd.to_datetime(m["Fecha_Finalizacion"])
    malo = fin.isna() | (fin < ini) | (fin.dt.normalize() != ini.dt.normalize())
    return pd.DataFrame({"NUI": m.loc[malo, "NUI_NORM"], "FECHA INICIO": ini[malo], "FECHA FIN": fin[malo]})


def construir_it2(mttos: pd.DataFrame, base: pd.DataFrame) -> pd.DataFrame:
    """Cada mantenimiento que cruzó se expande a todas las filas de la base de su NIU.
    El resultado queda ordenado por NUI."""
    mttos = mttos.reset_index(drop=True)
    inicio, fin, resumen = asignar_fechas(mttos)
    m = pd.DataFrame({
        "NUI_MANTENIMIENTO": mttos["NUI_NORM"].to_numpy(),
        "TIPO DE MANTENIMIENTO": tipo_mantenimiento(mttos).to_numpy(),
        "FECHA INICIO": inicio.to_numpy(),
        "FECHA FIN": fin.to_numpy(),
    })
    m["_orden"] = pd.to_numeric(m["NUI_MANTENIMIENTO"])
    m = m.sort_values("_orden", kind="stable")  # bloques por NUI; dentro de cada uno queda el orden de la base
    df = m.merge(base, left_on="NUI_MANTENIMIENTO", right_on="NIU", how="left")

    out = pd.DataFrame(index=df.index, columns=COLUMNAS_IT2, dtype=object)
    out["NIU"] = pd.to_numeric(df["NIU"])
    out["COD_LOCALIDAD"] = df["cod_localidad"]
    out["DANE"] = df["dane"]
    out["TIPO_UC"] = df["TIPO_UC"]
    out["TIPO_UC_9995"] = df["TIPO_UC"].map(HOMOLOGACION_TIPO_UC)
    out["CAPACIDAD_UC"] = df["CAPACIDAD_UC"].astype("Int64")
    out["MARCA"] = df["MARCA"]
    out["SERIAL_INTERNO"] = df["SERIAL_INTERNO"]
    out["NUI_MANTENIMIENTO"] = pd.to_numeric(df["NUI_MANTENIMIENTO"])
    out["TIPO DE MANTENIMIENTO"] = df["TIPO DE MANTENIMIENTO"]
    out["DECO TIPO MANTENIMIENTO"] = df["TIPO DE MANTENIMIENTO"].map(DECO_TIPO_MTTO)
    out["FECHA INICIO"] = df["FECHA INICIO"]
    out["FECHA FIN"] = df["FECHA FIN"]
    out = out.reset_index(drop=True)
    out.attrs["resumen_fechas"] = resumen
    return out


def excel_it2(df: pd.DataFrame) -> bytes:
    """Excel editable: encabezado en negrita, primera fila fija y anchos de columna ajustados."""
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="dd-mm-yyyy hh:mm") as writer:
        df.to_excel(writer, sheet_name="IT2", index=False)
        ws = writer.sheets["IT2"]
        ws.freeze_panes = "A2"
        for i, col in enumerate(df.columns, start=1):
            ws.cell(row=1, column=i).font = Font(bold=True)
            largo = df[col].astype(str).str.len().head(500).max() if len(df) else 0
            ws.column_dimensions[get_column_letter(i)].width = min(max(len(col), largo) + 2, 45)
            if col in ("FECHA INICIO", "FECHA FIN"):
                for celda in ws[get_column_letter(i)][1:]:
                    celda.number_format = "dd-mm-yyyy hh:mm"
    return buf.getvalue()
