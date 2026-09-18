"""Lógica del formato IT2 Mantenimientos (Paso 1: carga, filtro por mes y cruce con la base)."""
import re
from io import BytesIO

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


def construir_it2(mttos: pd.DataFrame, base: pd.DataFrame) -> pd.DataFrame:
    """Cada mantenimiento que cruzó se expande a todas las filas de la base de su NIU."""
    m = pd.DataFrame({
        "NUI_MANTENIMIENTO": mttos["NUI_NORM"].to_numpy(),
        "TIPO DE MANTENIMIENTO": tipo_mantenimiento(mttos).to_numpy(),
    })
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
    return out.reset_index(drop=True)


def excel_it2(df: pd.DataFrame) -> bytes:
    """Excel editable: encabezado en negrita, primera fila fija y anchos de columna ajustados."""
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="IT2", index=False)
        ws = writer.sheets["IT2"]
        ws.freeze_panes = "A2"
        for i, col in enumerate(df.columns, start=1):
            ws.cell(row=1, column=i).font = Font(bold=True)
            largo = df[col].astype(str).str.len().head(500).max() if len(df) else 0
            ws.column_dimensions[get_column_letter(i)].width = min(max(len(col), largo) + 2, 45)
    return buf.getvalue()
