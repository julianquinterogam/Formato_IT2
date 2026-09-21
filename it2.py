"""Lógica del formato IT2 Mantenimientos (Paso 1: carga, filtro por mes y cruce con la base)."""
import datetime
import re
import unicodedata
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
DECO_ESTADO = {"Funcional": 1, "No Funcional": 2}
CAMPOS_ESTADO_15 = ["estadoGabinete", "paneles", "puestaTierra", "inversor", "bateria", "protecciones", "mppt", "soporte"]
# Campo del listado 1.5 que describe el estado de cada elemento (TIPO_UC de la base)
CAMPO_ESTADO_POR_TIPO_UC = {
    1: "paneles",          # PANEL
    2: "inversor",         # INVERSOR
    3: "mppt",             # CONTROLADOR
    4: "bateria",          # BATERIAS
    5: "soporte",          # POSTE
    6: "paneles",          # RED PANEL
    7: "estadoGabinete",   # GABINETE
    8: "estadoGabinete",   # RED BATERIA GABINETE
    9: "puestaTierra",     # PUESTA TIERRA
    10: "protecciones",    # RED DOMICILIARIA
    11: "protecciones",    # MEDIDOR
}


def tipo_mantenimiento(df: pd.DataFrame) -> pd.Series:
    """1.5: columna Maintenance_Type. 2.0: columnas Preventivo / Correctivo marcadas con True."""
    es_15 = df["VERSION"] == "1.5"
    tipo_15 = df["Maintenance_Type"].astype("string").str.strip().str.capitalize() if "Maintenance_Type" in df else None
    prev = df["Preventivo"].eq(True) if "Preventivo" in df else pd.Series(False, index=df.index)
    corr = df["Correctivo"].eq(True) if "Correctivo" in df else pd.Series(False, index=df.index)
    tipo_20 = pd.Series(pd.NA, index=df.index, dtype="string")
    tipo_20[corr] = "Correctivo"
    tipo_20[prev] = "Preventivo"  # si vinieran ambos marcados, prevalece Preventivo
    tipo = tipo_20 if tipo_15 is None else tipo_15.where(es_15, tipo_20)
    return tipo.fillna("Preventivo")  # sin dato (ni Preventivo ni Correctivo marcado) -> se asume Preventivo


def _a_estado(col: pd.Series) -> pd.Series:
    """Listado 1.5: 'Bueno' -> Funcional. Todo lo demás ('Malo...', 'No Tiene', 'No Existe' o vacío) -> No Funcional."""
    v = col.astype("string").str.strip().str.lower()
    est = pd.Series("No Funcional", index=col.index, dtype="string")
    est[v == "bueno"] = "Funcional"
    return est


CAMPOS_TEXTO_ESTADO_20 = ["Entrega", "Estado_Entrega_instalacion", "Estado_Instalacion", "Hallazgos", "Observaciones"]


def estado_entrega_20(df: pd.DataFrame) -> pd.Series:
    """Listado 2.0: estado general del NUI. Se busca primero en 'Entrega' (Funcional / No Funcional); si viene
    vacío o con 'undefined', se busca la palabra 'funcional' en 'Estado_Entrega_instalacion', 'Estado_Instalacion',
    'Hallazgos' y 'Observaciones', en ese orden (se revisa 'no funcional' antes que 'funcional' para no
    confundirlas). Si ninguna trae información, se asume No Funcional, para que quede marcado y se revise."""
    est = pd.Series(pd.NA, index=df.index, dtype="string")
    for campo in CAMPOS_TEXTO_ESTADO_20:
        if campo not in df.columns:
            continue
        falta = est.isna()
        if not falta.any():
            break
        texto = df.loc[falta, campo].astype("string").str.strip().str.lower()
        est.loc[falta[falta].index[texto.str.contains("no funcional", na=False)]] = "No Funcional"
        falta = est.isna()
        texto = df.loc[falta, campo].astype("string").str.strip().str.lower()
        est.loc[falta[falta].index[texto.str.contains("funcional", na=False)]] = "Funcional"
    return est.fillna("No Funcional")  # sin ningún dato de estado: se deja marcado para revisión manual


def estado_por_fila(df: pd.DataFrame) -> pd.Series:
    """df: filas ya expandidas con la base. 1.5: cada elemento toma el estado del campo que le corresponde
    (CAMPO_ESTADO_POR_TIPO_UC). 2.0: todas las filas del NUI toman el estado general de 'Entrega'."""
    est = pd.DataFrame({c: _a_estado(df[f"_{c}"]) for c in CAMPOS_ESTADO_15})
    pos = df["TIPO_UC"].map(CAMPO_ESTADO_POR_TIPO_UC).map({c: i for i, c in enumerate(CAMPOS_ESTADO_15)})
    valido = pos.notna().to_numpy()
    arr = est.to_numpy(dtype=object)
    res = np.full(len(df), pd.NA, dtype=object)
    res[valido] = arr[np.flatnonzero(valido), pos[valido].astype(int).to_numpy()]
    est_15 = pd.Series(res, index=df.index, dtype="string")
    return est_15.where(df["_VERSION"] == "1.5", df["_ESTADO_20"])


def _sin_tildes(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn").lower()


def _columna_mes(columnas, anio: int, mes: int):
    for c in columnas:
        if isinstance(c, (pd.Timestamp, datetime.datetime, datetime.date)) and (c.year, c.month) == (anio, mes):
            return c
    return None


def valores_inversion(contenido: bytes, anio: int, mes: int) -> pd.Series:
    """Serie {TIPO_ELEMENTO: valor} tomada de la columna del mes seleccionado en la hoja 'Valor inversion'.
    Si esa columna no trae valores calculados, se calcula como VALOR_INVERSION_BASE x (IPP del mes / IPP base)
    con la hoja 'index'."""
    xl = pd.ExcelFile(BytesIO(contenido), engine="openpyxl")
    hoja = next((h for h in xl.sheet_names if "valor" in _sin_tildes(h) and "inversion" in _sin_tildes(h)), None)
    if hoja is None:
        raise ValueError("El Excel IT2 no tiene una hoja 'Valor inversion'.")
    df = xl.parse(hoja)
    col = _columna_mes(df.columns, anio, mes)
    if col is None:
        raise ValueError(f"La hoja '{hoja}' no tiene una columna para {MESES[mes]} {anio}.")
    valor = pd.to_numeric(df[col], errors="coerce")
    if valor.isna().any():  # sin valores cacheados: se calcula con el índice IPP
        h_idx = next((h for h in xl.sheet_names if _sin_tildes(h).strip() == "index"), None)
        if h_idx is None:
            raise ValueError(f"La columna de {MESES[mes]} {anio} está vacía y no existe la hoja 'index' para calcularla.")
        idx = xl.parse(h_idx)
        col_idx = _columna_mes(idx.columns, anio, mes)
        ipp_base = pd.to_numeric(idx["Ipp base*"], errors="coerce").iloc[0]
        ipp_mes = pd.to_numeric(idx[col_idx], errors="coerce").iloc[0] if col_idx is not None else np.nan
        if pd.isna(ipp_base) or pd.isna(ipp_mes):
            raise ValueError(f"No hay valor ni índice IPP para {MESES[mes]} {anio}.")
        valor = pd.to_numeric(df["VALOR_INVERSION_BASE"], errors="coerce") * (ipp_mes / ipp_base)
    tipo = pd.to_numeric(df["TIPO_ELEMENTO"], errors="coerce")
    return pd.Series(valor.to_numpy(), index=tipo.to_numpy())[lambda x: x.index.notna()].rename("VALOR")


def unificar_por_nui(mttos: pd.DataFrame):
    """Si un NUI tiene mantenimientos en fechas (días) distintas dentro del mes, se conservan todos: cada uno
    genera su propio bloque de filas. Si coinciden en el mismo día, se deja uno solo, el primero del listado.
    Devuelve (mantenimientos_unicos, cantidad_de_registros_descartados)."""
    ordenado = mttos.assign(_dia=mttos["FECHA_REF"].dt.normalize()).sort_values("FECHA_REF", kind="stable")
    unicos = ordenado.drop_duplicates(subset=["NUI_NORM", "_dia"], keep="first").drop(columns="_dia")
    return unicos.reset_index(drop=True), len(mttos) - len(unicos)


# --- Descripción del mantenimiento -------------------------------------------
# Se asigna al azar una de estas 6 a cada NUI y se repite igual en todas las filas de ese NUI.
DESCRIPCIONES_MANTENIMIENTO = [
    "Inspeccion visual de los paneles fotovoltaicos  Limpieza superficie Ajuste de torques inspeccion de mastil y pernos de anclaje con inspeccion de torque revision de conexiones y puesta a tierra estado de fusibles sulfatacion prueba de voltaje y corrientes en circuito abierto y cerrado breakers en DC estado conexion y conductividad de cableado",
    "revision estado general de la planta revision de circuito interno adecuaciones no autorizadas estado de conexion verificacion de conexiones fraudulentas limpieza de gabinete capacitacion del usuario recomendaciones generales llenado planilla control de actividades verificacion de protecciones electricas y puesta a tierra en medicion y condicion optima verificado por usuario",
    "Funcionamiento de lectura de variables electricas y torque medicion de elemento patron en DC y AC equipo patro calibrado y certificado revision firmware y estado general de conexion funcionamiento tarjeta protocolo de evaluacion de datos de medida de disponibilidad",
    "Revision de los componentes electronicos verificacion de bms si aplica pruebas de voltaje y corriente en circuito abierto y cerrado verificacion conectores lubricacion y adecuacion inspeccion visual para detertar anomalias parte fisicas Limpieza de partes sulfatadas y torque protocolo carga y descarga",
    "Verificacion de voltajes y corrientes en circuito abierto y cerrado verificacion y configuracion d eocntrolador de carga a especificaciones del proyecto limpieza e inspeccion general del componente evaluacion de estado conectores y verificacion torques",
    "Configuracion limpieza analisis de medida en circuto abierto y cerrado prueba de conductividad inspeccion visual verificacion de conectores en etapa DC sujecion y torques limpieza sistema de ventilacion verificacion de circuito de potencia y iluminacion verificacion tablero distribucion prueba en circuito cerrado con carga protocolo de pruebas de estres al sistema",
]


def descripcion_por_nui(mttos: pd.DataFrame) -> pd.Series:
    """Una descripción al azar por NUI. Reproducible: depende solo del NUI y del mes del mantenimiento."""
    def elegir(nui, fecha):
        semilla = zlib.crc32(f"desc|{nui}|{fecha:%Y-%m}".encode())
        return DESCRIPCIONES_MANTENIMIENTO[int(np.random.default_rng(semilla).integers(len(DESCRIPCIONES_MANTENIMIENTO)))]
    return pd.Series([elegir(n, f) for n, f in zip(mttos["NUI_NORM"], mttos["FECHA_REF"])], index=mttos.index)


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
    zona = grupo["vereda"] if "vereda" in grupo else pd.Series(np.nan, index=grupo.index)
    if "Municipio" in grupo:  # el listado 2.0 no trae vereda: se agrupa por municipio
        zona = zona.fillna(grupo["Municipio"])
    g = grupo.assign(
        _vereda=zona.fillna("").astype(str),
        _ts=(pd.to_numeric(grupo["Id_Encuesta"].astype("string").str.split("-").str[-1], errors="coerce").fillna(0)
             if "Id_Encuesta" in grupo else 0),
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


def _fechas_20_malas(mttos: pd.DataFrame) -> pd.Series:
    """True en los registros 2.0 cuya fecha fin falta, es anterior al inicio o cae en otro día."""
    ini = pd.to_datetime(mttos["Fecha_Inicio"]).dt.floor("min")
    fin = pd.to_datetime(mttos["Fecha_Finalizacion"]).dt.floor("min")
    malo = fin.isna() | (fin < ini) | (fin.dt.normalize() != ini.dt.normalize())
    return (mttos["VERSION"] == "2.0") & malo


def asignar_fechas(mttos: pd.DataFrame):
    """Devuelve (inicio, fin, resumen). Reproducible: el mismo mes siempre genera las mismas horas.
    - 2.0 con fechas consistentes: se toman tal cual (sin segundos).
    - 1.5 (solo trae el día) y 2.0 con fin en otro día: horas simuladas con las reglas de jornada,
      duración, desplazamiento y cuadrillas; el día es el de inicio."""
    inicio = pd.Series(pd.NaT, index=mttos.index, dtype="datetime64[ns]")
    fin = pd.Series(pd.NaT, index=mttos.index, dtype="datetime64[ns]")
    es20 = mttos["VERSION"] == "2.0"
    if es20.any():
        inicio[es20] = pd.to_datetime(mttos.loc[es20, "Fecha_Inicio"]).dt.floor("min")
        fin[es20] = pd.to_datetime(mttos.loc[es20, "Fecha_Finalizacion"]).dt.floor("min")
    malas20 = _fechas_20_malas(mttos)
    simular = mttos[~es20 | malas20]
    grupos_paralelo = registros_paralelo = 0
    if len(simular):
        dias = simular["FECHA_REF"].dt.normalize()
        for (version, usuario, dia), grupo in simular.groupby(
                [simular["VERSION"], simular["UserName"].fillna(""), dias], sort=False):
            semilla = f"{usuario}|{dia:%Y-%m-%d}" if version == "1.5" else f"2.0|{dia:%Y-%m-%d}"
            ini_g, fin_g, k = _programar_grupo(grupo, dia, semilla)
            for etiqueta in ini_g:
                inicio[etiqueta], fin[etiqueta] = ini_g[etiqueta], fin_g[etiqueta]
            if k > 1:
                grupos_paralelo += 1
                registros_paralelo += len(grupo)
    resumen = {"grupos_en_paralelo": grupos_paralelo, "registros_en_paralelo": registros_paralelo,
               "reprogramados_20": int(malas20.sum())}
    return inicio, fin, resumen


def fechas_20_inconsistentes(mttos: pd.DataFrame) -> pd.DataFrame:
    """Registros 2.0 cuya fecha fin es anterior al inicio o cae en otro día."""
    m = mttos[mttos["VERSION"] == "2.0"]
    ini = pd.to_datetime(m["Fecha_Inicio"]); fin = pd.to_datetime(m["Fecha_Finalizacion"])
    malo = fin.isna() | (fin < ini) | (fin.dt.normalize() != ini.dt.normalize())
    return pd.DataFrame({"NUI": m.loc[malo, "NUI_NORM"], "FECHA INICIO": ini[malo], "FECHA FIN": fin[malo]})


def construir_it2(mttos: pd.DataFrame, base: pd.DataFrame, valores: pd.Series | None = None) -> pd.DataFrame:
    """Cada mantenimiento que cruzó se expande a todas las filas de la base de su NIU.
    El resultado queda ordenado por NUI."""
    mttos = mttos.reset_index(drop=True)
    inicio, fin, resumen = asignar_fechas(mttos)
    m = pd.DataFrame({
        "NUI_MANTENIMIENTO": mttos["NUI_NORM"].to_numpy(),
        "TIPO DE MANTENIMIENTO": tipo_mantenimiento(mttos).to_numpy(),
        "MANTENIMIENTO REALIZADO": descripcion_por_nui(mttos).to_numpy(),
        "_VERSION": mttos["VERSION"].to_numpy(),
        "_ESTADO_20": estado_entrega_20(mttos).to_numpy(),
        "FECHA INICIO": inicio.to_numpy(),
        "FECHA FIN": fin.to_numpy(),
    })
    for c in CAMPOS_ESTADO_15:
        m[f"_{c}"] = mttos[c].to_numpy() if c in mttos.columns else None
    m["_orden"] = pd.to_numeric(m["NUI_MANTENIMIENTO"])
    m = m.sort_values("_orden", kind="stable")  # bloques por NUI; dentro de cada uno queda el orden de la base
    df = m.merge(base, left_on="NUI_MANTENIMIENTO", right_on="NIU", how="left")
    df["ESTADO"] = estado_por_fila(df)

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
    out["MANTENIMIENTO REALIZADO"] = df["MANTENIMIENTO REALIZADO"]
    out["FECHA INICIO"] = df["FECHA INICIO"]
    out["FECHA FIN"] = df["FECHA FIN"]
    out["ESTADO"] = df["ESTADO"]
    out["DECO ESTADO"] = df["ESTADO"].map(DECO_ESTADO)
    if valores is not None:  # cruce TIPO_UC_9995 <-> TIPO_ELEMENTO
        out["VALOR"] = out["TIPO_UC_9995"].map(valores)
    out = out.reset_index(drop=True)
    out.attrs["resumen_fechas"] = resumen
    return out


# --- Formato final (9 columnas homologadas, para envío) --------------------
COLUMNAS_IT2_FINAL = [
    "Codigo Localidad", "Tipo de Elemento", "Serial del Elemento", "Tipo de Mantenimiento",
    "Mantenimiento Realizado", "Fecha y hora de inicio", "Fecha y hora fin", "Estado",
    "Valor de la Intervencion",
]
HOMOLOGACION_FINAL = {
    "Codigo Localidad": "COD_LOCALIDAD",
    "Tipo de Elemento": "TIPO_UC_9995",
    "Serial del Elemento": "SERIAL_INTERNO",
    "Tipo de Mantenimiento": "DECO TIPO MANTENIMIENTO",
    "Mantenimiento Realizado": "MANTENIMIENTO REALIZADO",
    "Fecha y hora de inicio": "FECHA INICIO",
    "Fecha y hora fin": "FECHA FIN",
    "Estado": "DECO ESTADO",
    "Valor de la Intervencion": "VALOR",
}


def construir_it2_final(it2: pd.DataFrame) -> pd.DataFrame:
    """A partir del formato editable, deja solo las 9 columnas homologadas que se envían."""
    out = pd.DataFrame({destino: it2[origen] for destino, origen in HOMOLOGACION_FINAL.items()})
    out["Codigo Localidad"] = pd.to_numeric(out["Codigo Localidad"]).astype("Int64")
    return out[COLUMNAS_IT2_FINAL]


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
            if col in ("VALOR", "Valor de la Intervencion"):
                for celda in ws[get_column_letter(i)][1:]:
                    celda.number_format = "#,##0.00"
            if col in ("FECHA INICIO", "FECHA FIN", "Fecha y hora de inicio", "Fecha y hora fin"):
                for celda in ws[get_column_letter(i)][1:]:
                    celda.number_format = "dd-mm-yyyy hh:mm"
    return buf.getvalue()
