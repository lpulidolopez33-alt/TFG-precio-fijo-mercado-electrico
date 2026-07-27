from io import StringIO
from pathlib import Path

import pandas as pd
import requests

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    RAW_DATA_DIR,
    PROCESSED_DATA_DIR,
    MODO_PRUEBA_OMIE,
    DIAS_PRUEBA_OMIE,
    VERSIONES_OMIE,
)

from datos_base import crear_indice_horario


OMIE_RAW_DIR = RAW_DATA_DIR / "omie"


def convertir_numero_omie(valor: str) -> float:
    """
    Convierte un número en formato OMIE a float de Python.

    OMIE suele usar coma decimal, por ejemplo:
    '72,35' -> 72.35
    """

    texto = str(valor).strip()

    if texto == "":
        raise ValueError("Valor vacío")

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    return float(texto)


def construir_nombre_fichero_omie(fecha: pd.Timestamp, version: int) -> str:
    """
    Construye el nombre del fichero diario de OMIE.
    """

    fecha_txt = fecha.strftime("%Y%m%d")

    return f"marginalpdbc_{fecha_txt}.{version}"


def descargar_fichero_omie(fecha: pd.Timestamp) -> tuple[str, str]:
    """
    Descarga el fichero diario de precios horarios de OMIE.

    Devuelve:
    - el texto del fichero
    - el nombre del fichero descargado
    """

    OMIE_RAW_DIR.mkdir(parents=True, exist_ok=True)

    url = "https://www.omie.es/es/file-download"

    for version in VERSIONES_OMIE:
        nombre_fichero = construir_nombre_fichero_omie(fecha, version)
        ruta_local = OMIE_RAW_DIR / f"{nombre_fichero}.txt"

        if ruta_local.exists():
            texto = ruta_local.read_text(encoding="latin-1")
            return texto, nombre_fichero

        parametros = {
            "parents[0]": "marginalpdbc",
            "filename": nombre_fichero,
        }

        respuesta = requests.get(
            url,
            params=parametros,
            timeout=30,
            headers={"User-Agent": "Mozilla/5.0"},
        )

        if respuesta.status_code == 200 and len(respuesta.text) > 100:
            texto = respuesta.text
            ruta_local.write_text(texto, encoding="latin-1")
            return texto, nombre_fichero

    raise FileNotFoundError(f"No se ha encontrado fichero OMIE para {fecha.date()}")


def detectar_filas_datos_omie(texto: str) -> list[tuple[list[str], int]]:
    """
    Detecta las filas reales de datos dentro del fichero OMIE.

    Devuelve una lista de tuplas:
    - celdas de la fila
    - posición donde empieza el año dentro de la fila

    Algunos ficheros pueden empezar directamente por el año:
    2026;4;29;1;1;precio...

    Otros podrían incluir una primera columna de texto:
    MARGINALPDBC;2026;4;29;1;1;precio...
    """

    filas_datos = []

    for linea in StringIO(texto):
        celdas = [celda.strip() for celda in linea.strip().split(";")]

        while celdas and celdas[-1] == "":
            celdas.pop()

        if not celdas:
            continue

        posicion_año = None

        if len(celdas) >= 5 and celdas[0].isdigit() and len(celdas[0]) == 4:
            posicion_año = 0

        elif len(celdas) >= 6 and celdas[1].isdigit() and len(celdas[1]) == 4:
            posicion_año = 1

        if posicion_año is None:
            continue

        filas_datos.append((celdas, posicion_año))

    return filas_datos


def extraer_precios_desde_texto_omie(texto: str) -> list[float]:
    """
    Extrae los precios de España desde el texto descargado de OMIE.

    El programa distingue entre:
    - ficheros horarios: año; mes; día; hora; precio España; ...
    - ficheros cuarto-horarios: año; mes; día; hora; cuarto; precio España; ...
    """

    filas_datos = detectar_filas_datos_omie(texto)

    if not filas_datos:
        raise ValueError("No se han encontrado filas de datos en el fichero OMIE.")

    numero_filas = len(filas_datos)

    if numero_filas in [23, 24, 25]:
        fichero_cuarto_horario = False

    elif numero_filas in [92, 96, 100]:
        fichero_cuarto_horario = True

    else:
        raise ValueError(
            f"Número inesperado de filas en el fichero OMIE: {numero_filas}. "
            f"No se puede determinar si el fichero es horario o cuarto-horario."
        )

    precios = []

    for celdas, posicion_año in filas_datos:
        if fichero_cuarto_horario:
            indice_precio_espana = posicion_año + 5
        else:
            indice_precio_espana = posicion_año + 4

        if indice_precio_espana >= len(celdas):
            raise ValueError(
                f"No existe columna de precio España en la fila: {celdas}"
            )

        precio = convertir_numero_omie(celdas[indice_precio_espana])
        precios.append(precio)

    return precios

def convertir_precios_a_horarios(precios: list[float], horas_esperadas: int) -> list[float]:
    """
    Convierte los precios descargados de OMIE a formato horario.

    Si OMIE devuelve ya 24 precios, se dejan igual.
    Si OMIE devuelve 96 precios, se agrupan de 4 en 4 haciendo la media.
    """

    if len(precios) == horas_esperadas:
        return precios

    if len(precios) == horas_esperadas * 4:
        precios_horarios = []

        for i in range(0, len(precios), 4):
            bloque_cuarto_horario = precios[i:i + 4]
            precio_horario = sum(bloque_cuarto_horario) / len(bloque_cuarto_horario)
            precios_horarios.append(precio_horario)

        return precios_horarios

    raise ValueError(
        f"No se puede convertir la serie de precios. "
        f"Valores recibidos: {len(precios)}. "
        f"Horas esperadas: {horas_esperadas}."
    )

def parsear_fichero_omie(texto: str, fecha: pd.Timestamp, nombre_fichero: str) -> pd.DataFrame:
    """
    Convierte el fichero diario de OMIE en una tabla horaria.
    """

    precios_originales = extraer_precios_desde_texto_omie(texto)

    fecha_txt = fecha.strftime("%Y-%m-%d")
    indice_horario = crear_indice_horario(fecha_txt, fecha_txt)

    precios_horarios = convertir_precios_a_horarios(
        precios=precios_originales,
        horas_esperadas=len(indice_horario),
    )

    if len(precios_horarios) != len(indice_horario):
        raise ValueError(
            f"El número de precios horarios no coincide con las horas del día. "
            f"Fecha: {fecha_txt}. "
            f"Precios horarios: {len(precios_horarios)}. "
            f"Horas esperadas: {len(indice_horario)}."
        )

    df = pd.DataFrame()

    df["fecha_hora_local"] = indice_horario.astype(str)
    df["fecha_hora_utc"] = indice_horario.tz_convert("UTC").astype(str)
    df["fecha"] = fecha_txt
    df["precio_omie"] = precios_horarios
    df["numero_valores_originales_omie"] = len(precios_originales)
    df["fichero_origen_omie"] = nombre_fichero

    return df


def obtener_rango_fechas_omie() -> pd.DatetimeIndex:
    """
    Define el rango de fechas que se va a descargar de OMIE.
    """

    fecha_fin = pd.Timestamp(FECHA_FIN_HISTORICO)

    if MODO_PRUEBA_OMIE:
        fecha_inicio = fecha_fin - pd.Timedelta(days=DIAS_PRUEBA_OMIE - 1)
    else:
        fecha_inicio = pd.Timestamp(FECHA_INICIO_HISTORICO)

    return pd.date_range(
        start=fecha_inicio,
        end=fecha_fin,
        freq="D",
    )


def descargar_historico_omie() -> None:
    """
    Descarga y guarda el histórico horario de precios OMIE.
    """

    fechas = obtener_rango_fechas_omie()

    tablas_diarias = []
    errores = []

    print("\nDESCARGA DE PRECIOS OMIE")
    print("-" * 50)
    print(f"Fecha inicial OMIE: {fechas[0].date()}")
    print(f"Fecha final OMIE: {fechas[-1].date()}")
    print(f"Número de días a descargar: {len(fechas)}")
    print("-" * 50)

    for fecha in fechas:
        try:
            texto, nombre_fichero = descargar_fichero_omie(fecha)
            df_dia = parsear_fichero_omie(texto, fecha, nombre_fichero)
            tablas_diarias.append(df_dia)

            print(f"OK {fecha.date()} - {len(df_dia)} horas")

        except Exception as error:
            errores.append(
                {
                    "fecha": fecha.strftime("%Y-%m-%d"),
                    "error": str(error),
                }
            )

            print(f"ERROR {fecha.date()} - {error}")

    if not tablas_diarias:
        raise RuntimeError("No se ha podido descargar ningún dato de OMIE.")

    df_omie = pd.concat(tablas_diarias, ignore_index=True)

    ruta_salida = PROCESSED_DATA_DIR / "precios_omie_horarios.xlsx"
    df_omie.to_excel(ruta_salida, index=False)

    print("-" * 50)
    print(f"Archivo OMIE guardado: {ruta_salida}")
    print(f"Número de filas horarias OMIE: {len(df_omie)}")

    if errores:
        ruta_errores = PROCESSED_DATA_DIR / "errores_descarga_omie.xlsx"
        pd.DataFrame(errores).to_excel(ruta_errores, index=False)

        print(f"Días con error guardados en: {ruta_errores}")