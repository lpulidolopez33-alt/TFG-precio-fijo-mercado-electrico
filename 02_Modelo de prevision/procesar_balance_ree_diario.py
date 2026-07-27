from datetime import date, datetime
from pathlib import Path
import unicodedata

import pandas as pd

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    REE_BALANCE_DIARIO_DIR,
    RUTA_GENERACION_MANUAL,
)

from datos_base import crear_indice_horario


MESES_ES = {
    "ene": 1,
    "feb": 2,
    "mar": 3,
    "abr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dic": 12,
}


MAPEO_TECNOLOGIAS = {
    "hidraulica": "generacion_hidraulica",
    "eolica": "generacion_eolica",
    "solar fotovoltaica": "generacion_solar_fotovoltaica",
    "solar termica": "generacion_solar_termica",
    "nuclear": "generacion_nuclear",
    "ciclo combinado": "generacion_ciclo_combinado",
    "carbon": "generacion_carbon",
}


COLUMNAS_GENERACION = [
    "generacion_solar_fotovoltaica",
    "generacion_solar_termica",
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "generacion_carbon",
]


def normalizar_texto(texto: str) -> str:
    """
    Normaliza texto para comparar nombres sin acentos ni mayúsculas.
    """

    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )

    return texto


def parsear_fecha_ree(valor) -> date | None:
    """
    Convierte las fechas del Excel de REE a fecha de Python.

    Ejemplo:
    '01/ene/23' -> 2023-01-01
    """

    if pd.isna(valor):
        return None

    if isinstance(valor, pd.Timestamp):
        return valor.date()

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    texto = str(valor).strip().lower().replace(".", "")

    partes = texto.split("/")

    if len(partes) != 3:
        return None

    dia = int(partes[0])
    mes_txt = partes[1][:3]
    año_txt = partes[2]

    if mes_txt not in MESES_ES:
        return None

    mes = MESES_ES[mes_txt]

    año = int(año_txt)

    if año < 100:
        año = 2000 + año

    return date(año, mes, dia)


def obtener_horas_del_dia(fecha: date) -> int:
    """
    Devuelve el número de horas reales de un día en horario español.

    Normalmente son 24, pero en cambios horarios pueden ser 23 o 25.
    """

    fecha_txt = fecha.strftime("%Y-%m-%d")

    indice = crear_indice_horario(fecha_txt, fecha_txt)

    return len(indice)


def leer_balance_ree_diario(ruta_excel: Path) -> pd.DataFrame:
    """
    Lee un Excel de balance eléctrico de REE y lo convierte a formato diario.

    El archivo original viene con:
    - tecnologías en filas
    - días en columnas
    - magnitudes en GWh
    """

    df_raw = pd.read_excel(
        ruta_excel,
        sheet_name="data",
        header=None,
    )

    fila_fechas = 4

    fechas_por_columna = {}

    for columna in range(1, df_raw.shape[1]):
        fecha = parsear_fecha_ree(df_raw.iloc[fila_fechas, columna])

        if fecha is not None:
            fechas_por_columna[columna] = fecha

    registros = []

    for fila in range(fila_fechas + 1, df_raw.shape[0]):
        tecnologia_original = df_raw.iloc[fila, 0]
        tecnologia_normalizada = normalizar_texto(tecnologia_original)

        if tecnologia_normalizada not in MAPEO_TECNOLOGIAS:
            continue

        columna_modelo = MAPEO_TECNOLOGIAS[tecnologia_normalizada]

        for columna, fecha in fechas_por_columna.items():
            valor_gwh = pd.to_numeric(
                df_raw.iloc[fila, columna],
                errors="coerce",
            )

            registros.append(
                {
                    "fecha": fecha,
                    "columna_modelo": columna_modelo,
                    "valor_gwh": valor_gwh,
                    "archivo_origen": ruta_excel.name,
                }
            )

    df_largo = pd.DataFrame(registros)

    if df_largo.empty:
        raise ValueError(
            f"No se han encontrado tecnologías útiles en el archivo {ruta_excel}"
        )

    df_diario = (
        df_largo
        .pivot_table(
            index="fecha",
            columns="columna_modelo",
            values="valor_gwh",
            aggfunc="sum",
        )
        .reset_index()
    )

    df_diario.columns.name = None

    return df_diario


def convertir_gwh_diarios_a_mw_medios(df_diario: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte generación diaria en GWh a potencia media diaria en MW.
    """

    df = df_diario.copy()

    df["horas_dia"] = df["fecha"].apply(obtener_horas_del_dia)

    for columna in COLUMNAS_GENERACION:
        if columna not in df.columns:
            df[columna] = None

    columnas_gwh = [
        "generacion_solar_fotovoltaica",
        "generacion_solar_termica",
        "generacion_eolica",
        "generacion_hidraulica",
        "generacion_nuclear",
        "generacion_ciclo_combinado",
        "generacion_carbon",
    ]

    for columna in columnas_gwh:
        df[columna] = pd.to_numeric(df[columna], errors="coerce")
        df[columna] = df[columna] * 1000 / df["horas_dia"]

    df["generacion_solar"] = (
        df["generacion_solar_fotovoltaica"].fillna(0)
        + df["generacion_solar_termica"].fillna(0)
    )

    hay_solar = (
        df["generacion_solar_fotovoltaica"].notna()
        | df["generacion_solar_termica"].notna()
    )

    df.loc[~hay_solar, "generacion_solar"] = None

    df = df[
        [
            "fecha",
            "generacion_solar_fotovoltaica",
            "generacion_solar_termica",
            "generacion_solar",
            "generacion_eolica",
            "generacion_hidraulica",
            "generacion_nuclear",
            "generacion_ciclo_combinado",
            "generacion_carbon",
        ]
    ]

    return df


def construir_generacion_horaria_desde_diaria(df_diario_mw: pd.DataFrame) -> pd.DataFrame:
    """
    Construye una tabla horaria repitiendo para cada hora el valor medio diario.
    """

    indice_horario = crear_indice_horario(
        FECHA_INICIO_HISTORICO,
        FECHA_FIN_HISTORICO,
    )

    df_horario = pd.DataFrame()

    df_horario["fecha_hora_local"] = indice_horario.astype(str)
    df_horario["fecha_hora_utc"] = indice_horario.tz_convert("UTC").astype(str)
    df_horario["fecha"] = indice_horario.date
    df_horario["año"] = indice_horario.year
    df_horario["mes"] = indice_horario.month
    df_horario["dia"] = indice_horario.day
    df_horario["hora"] = indice_horario.hour

    df_horario = df_horario.merge(
        df_diario_mw,
        on="fecha",
        how="left",
    )

    return df_horario


def procesar_balance_ree_diario() -> None:
    """
    Procesa los Excel diarios de balance eléctrico de REE y genera
    generacion_manual.xlsx compatible con el resto del programa.
    """

    print("\nPROCESADO DE BALANCE REE DIARIO")
    print("-" * 50)

    REE_BALANCE_DIARIO_DIR.mkdir(parents=True, exist_ok=True)

    archivos = sorted(REE_BALANCE_DIARIO_DIR.glob("*.xlsx"))

    if not archivos:
        raise FileNotFoundError(
            f"No hay archivos Excel en {REE_BALANCE_DIARIO_DIR}. "
            f"Copia ahí los balance-electrico_*.xlsx descargados de REE."
        )

    tablas_diarias = []

    for archivo in archivos:
        df_archivo = leer_balance_ree_diario(archivo)
        tablas_diarias.append(df_archivo)

        print(f"Archivo leído: {archivo.name} - {len(df_archivo)} días")

    df_diario = pd.concat(tablas_diarias, ignore_index=True)

    df_diario = df_diario.drop_duplicates(
        subset=["fecha"],
        keep="last",
    )

    df_diario = df_diario.sort_values("fecha")

    df_diario_mw = convertir_gwh_diarios_a_mw_medios(df_diario)

    df_horario = construir_generacion_horaria_desde_diaria(df_diario_mw)

    try:
        df_horario.to_excel(RUTA_GENERACION_MANUAL, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {RUTA_GENERACION_MANUAL}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Generación manual generada desde balance diario: {RUTA_GENERACION_MANUAL}")
    print(f"Número de filas horarias generadas: {len(df_horario)}")

    for columna in COLUMNAS_GENERACION:
        datos_disponibles = df_horario[columna].notna().sum()
        print(f"{columna}: {datos_disponibles} datos disponibles")

    print("-" * 50)