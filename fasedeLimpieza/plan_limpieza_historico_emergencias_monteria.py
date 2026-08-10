"""Plan de limpieza ejecutable para Historico Emergencias Monteria.

Este script no se ejecuta automaticamente sobre el archivo original.
Su objetivo es documentar y dejar listo un flujo reproducible de limpieza
para la hoja detallada Hoja1.

Estrategia general:
- Cargar el archivo original.
- Estandarizar tipos y nombres de columnas.
- Normalizar fechas y texto.
- Marcar inconsistencias cronologicas y fechas corruptas.
- Derivar variables utiles para analisis y modelado.
- Guardar una version limpia en un archivo nuevo.

Reglas conservadoras:
- No eliminar registros sin auditoria previa.
- No inventar contratos ni fechas.
- Mantener flags de calidad para trazabilidad.
"""

from __future__ import annotations

from pathlib import Path
import re
import unicodedata
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "Historico Emergencias Monteria.xlsx"
OUTPUT_FILE = BASE_DIR / "Historico_Emergencias_Monteria_limpio.xlsx"
REPORT_FILE = BASE_DIR / "reporte_calidad_historico_emergencias_monteria.txt"

DATE_COLS = [
    "Fecha registro orden",
    "Fecha asignación",
    "Fecha_llegada",
    "Fecha_control",
    "Fecha_normalizacion",
    "Fecha_finalizacion",
    "Fecha solicitud",
]

TEXT_COLS = [
    "Dirección",
    "Persona ejecuta",
    "Persona legaliza",
    "Nombre unidad operativa",
    "Causal solicitud",
    "Observacion solicitud",
    "Observación",
    "Contacto",
]

KEY_COLS = [
    "Contrato",
    "Orden",
    "Nro_solicitud",
    "Interaccion",
    "Codigo unidad de  operativa",
]


def normalize_text(value: object) -> object:
    if pd.isna(value):
        return pd.NA
    text = str(value)
    text = text.replace("_x000D_", " ")
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text)
    return text.strip().upper()


def extract_barrio(text: object) -> object:
    if pd.isna(text):
        return pd.NA
    value = str(text)
    patterns = [
        r"BARRIO:\s*([^_\n\r]+)",
        r"BARRIO\s*[:\-]\s*([^_\n\r]+)",
        r"\bBARRIO\b\s*([^_\n\r]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, value, flags=re.IGNORECASE)
        if match:
            return normalize_text(match.group(1)).strip(" .;,-")
    return pd.NA


def add_temporal_features(df: pd.DataFrame) -> pd.DataFrame:
    registro = df["Fecha registro orden"]
    df["anio"] = registro.dt.year
    df["trimestre"] = registro.dt.to_period("Q").astype(str)
    df["mes"] = registro.dt.month
    df["semana"] = registro.dt.isocalendar().week.astype("Int64")
    df["dia_semana"] = registro.dt.day_name()
    df["hora"] = registro.dt.hour
    return df


def add_operational_times(df: pd.DataFrame) -> pd.DataFrame:
    df["tiempo_respuesta_h"] = (df["Fecha_llegada"] - df["Fecha asignación"]).dt.total_seconds() / 3600
    df["tiempo_control_h"] = (df["Fecha_control"] - df["Fecha_llegada"]).dt.total_seconds() / 3600
    df["duracion_total_h"] = (df["Fecha_finalizacion"] - df["Fecha registro orden"]).dt.total_seconds() / 3600
    return df


def add_quality_flags(df: pd.DataFrame) -> pd.DataFrame:
    df["flag_resp_negativa"] = df["tiempo_respuesta_h"] < 0
    df["flag_control_negativo"] = df["tiempo_control_h"] < 0
    df["flag_total_negativa"] = df["duracion_total_h"] < 0
    df["flag_inconsistencia_temporal"] = (
        (df["Fecha registro orden"] > df["Fecha asignación"])
        | (df["Fecha asignación"] > df["Fecha_llegada"])
        | (df["Fecha_llegada"] > df["Fecha_control"])
        | (df["Fecha_control"] > df["Fecha_finalizacion"])
    )
    return df


def build_text_features(df: pd.DataFrame) -> pd.DataFrame:
    combined = (
        df["Observacion solicitud"].fillna("").astype(str)
        + " "
        + df["Observación"].fillna("").astype(str)
    )
    combined = combined.str.replace("_x000D_", " ", regex=False)
    combined = combined.str.replace(r"\s+", " ", regex=True).str.strip()
    df["texto_completo_limpio"] = combined.str.upper()
    df["barrio_extraido"] = df["Observacion solicitud"].map(extract_barrio)

    families = {
        "flag_excavaciones": ["EXCAV", "ZANJA", "RETRO", "PERFOR", "OBRA", "ANDAMIO"],
        "flag_robo": ["ROBO", "HURTO", "HURT", "SUSTRA", "LADRON"],
        "flag_dano_tuberia": ["TUBER", "ROT", "FISUR", "FUGA", "VALVUL", "ACOMETIDA"],
        "flag_obras_civiles": ["OBRA", "CONSTRUCC", "PAVIMENT", "CEMENTO", "PISO"],
    }

    for flag_name, keywords in families.items():
        pattern = "|".join(keywords)
        df[flag_name] = df["texto_completo_limpio"].str.contains(pattern, case=False, regex=True, na=False)

    return df


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for column in DATE_COLS:
        df[column] = pd.to_datetime(df[column], errors="coerce")

    for column in TEXT_COLS:
        df[column] = df[column].map(normalize_text)

    for column in KEY_COLS:
        if column in df.columns:
            df[column] = df[column].astype("string").str.strip()

    df["Contrato"] = df["Contrato"].astype("string")
    df["Orden"] = df["Orden"].astype("string")
    df["Nro_solicitud"] = df["Nro_solicitud"].astype("string")
    df["Interaccion"] = df["Interaccion"].astype("string")
    df["Codigo unidad de  operativa"] = df["Codigo unidad de  operativa"].astype("string")

    df = add_temporal_features(df)
    df = add_operational_times(df)
    df = add_quality_flags(df)
    df = build_text_features(df)

    return df


def export_quality_report(df: pd.DataFrame) -> None:
    lines = []
    lines.append("REPORTE DE CALIDAD Y LIMPIEZA")
    lines.append(f"Registros: {len(df)}")
    lines.append("")
    lines.append("Problemas temporales:")
    lines.append(f"- Inconsistencias temporales: {int(df['flag_inconsistencia_temporal'].sum())}")
    lines.append(f"- Respuesta negativa: {int(df['flag_resp_negativa'].sum())}")
    lines.append(f"- Control negativo: {int(df['flag_control_negativo'].sum())}")
    lines.append(f"- Duracion total negativa: {int(df['flag_total_negativa'].sum())}")
    lines.append("")
    lines.append("Nulos por columna:")
    for column in df.columns:
        nulls = int(df[column].isna().sum())
        if nulls:
            lines.append(f"- {column}: {nulls}")

    REPORT_FILE.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"No se encontro el archivo de entrada: {INPUT_FILE}")

    raw = pd.read_excel(INPUT_FILE, sheet_name="Hoja1")
    cleaned = clean_dataframe(raw)
    export_quality_report(cleaned)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="raw", index=False)
        cleaned.to_excel(writer, sheet_name="cleaned", index=False)

    print(f"Archivo limpio generado en: {OUTPUT_FILE}")
    print(f"Reporte generado en: {REPORT_FILE}")


if __name__ == "__main__":
    main()