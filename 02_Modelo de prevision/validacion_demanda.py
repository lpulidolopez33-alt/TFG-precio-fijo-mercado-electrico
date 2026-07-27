
# Esto sirve para:

# 1. Comprobar si faltan horas de demanda.
# 2. Comprobar si hay horas duplicadas.
# 3. Calcular demanda mínima, máxima, media y mediana.
# 4. Sacar demanda media mensual.
# 5. Sacar perfil horario medio de demanda.
# 6. Guardar un Excel de control y dos gráficas.

import pandas as pd
import matplotlib.pyplot as plt

from config import PROCESSED_DATA_DIR, GRAPH_DIR


def cargar_base_modelo() -> pd.DataFrame:
    """
    Carga la base del modelo con OMIE y demanda ya integrados.
    """

    ruta_base = PROCESSED_DATA_DIR / "base_modelo.xlsx"

    if not ruta_base.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_base}")

    df = pd.read_excel(ruta_base)

    return df


def preparar_base_demanda(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la base para poder analizar la demanda eléctrica.
    """

    df = df.copy()

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["fecha_hora_utc_dt"] = pd.to_datetime(df["fecha_hora_utc"], utc=True)
    df["demanda"] = pd.to_numeric(df["demanda"], errors="coerce")

    return df


def crear_resumen_general_demanda(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea un resumen general de calidad de la demanda REE.
    """

    total_filas = len(df)
    filas_con_demanda = df["demanda"].notna().sum()
    filas_sin_demanda = df["demanda"].isna().sum()
    horas_duplicadas_utc = df["fecha_hora_utc"].duplicated().sum()

    resumen = {
        "total_filas": total_filas,
        "filas_con_demanda": filas_con_demanda,
        "filas_sin_demanda": filas_sin_demanda,
        "horas_utc_duplicadas": horas_duplicadas_utc,
        "demanda_minima": df["demanda"].min(),
        "demanda_maxima": df["demanda"].max(),
        "demanda_media": df["demanda"].mean(),
        "demanda_mediana": df["demanda"].median(),
    }

    df_resumen = pd.DataFrame([resumen])

    return df_resumen


def crear_resumen_mensual_demanda(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula la demanda media mensual e identifica si el mes está completo.
    """

    df = df.copy()

    df["año_mes"] = df["fecha"].dt.to_period("M").astype(str)

    resumen_mensual = (
        df.groupby("año_mes", as_index=False)
        .agg(
            demanda_media=("demanda", "mean"),
            demanda_minima=("demanda", "min"),
            demanda_maxima=("demanda", "max"),
            horas_con_dato=("demanda", "count"),
            primera_fecha=("fecha", "min"),
            ultima_fecha=("fecha", "max"),
        )
    )

    resumen_mensual["dias_mes"] = pd.to_datetime(
        resumen_mensual["año_mes"] + "-01"
    ).dt.days_in_month

    resumen_mensual["horas_teoricas_mes"] = resumen_mensual["dias_mes"] * 24

    resumen_mensual["mes_completo"] = (
        resumen_mensual["horas_con_dato"]
        >= resumen_mensual["horas_teoricas_mes"] - 1
    ).astype(int)

    return resumen_mensual


def crear_resumen_horario_demanda(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula la demanda media por hora del día.
    """

    resumen_horario = (
        df.groupby("hora", as_index=False)
        .agg(
            demanda_media=("demanda", "mean"),
            demanda_minima=("demanda", "min"),
            demanda_maxima=("demanda", "max"),
            horas_con_dato=("demanda", "count"),
        )
    )

    return resumen_horario


def guardar_control_calidad_demanda(
    resumen_general: pd.DataFrame,
    resumen_mensual: pd.DataFrame,
    resumen_horario: pd.DataFrame,
) -> None:
    """
    Guarda los resúmenes de demanda en un archivo Excel.
    """

    ruta_salida = PROCESSED_DATA_DIR / "control_calidad_demanda.xlsx"

    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            resumen_general.to_excel(writer, sheet_name="resumen_general", index=False)
            resumen_mensual.to_excel(writer, sheet_name="resumen_mensual", index=False)
            resumen_horario.to_excel(writer, sheet_name="resumen_horario", index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Control de calidad de demanda guardado en: {ruta_salida}")


def graficar_evolucion_mensual_demanda(resumen_mensual: pd.DataFrame) -> None:
    """
    Genera una gráfica con la evolución mensual de la demanda eléctrica.

    Para evitar interpretaciones erróneas, solo representa los meses completos.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    resumen_grafica = resumen_mensual[
        resumen_mensual["mes_completo"] == 1
    ].copy()

    plt.figure(figsize=(12, 5))
    plt.plot(
        resumen_grafica["año_mes"],
        resumen_grafica["demanda_media"],
        marker="o",
    )
    plt.xticks(rotation=90)
    plt.title("Evolución mensual de la demanda eléctrica peninsular")
    plt.xlabel("Mes")
    plt.ylabel("Demanda media [MW]")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "evolucion_mensual_demanda.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_perfil_horario_demanda(resumen_horario: pd.DataFrame) -> None:
    """
    Genera una gráfica con la demanda media por hora del día.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    plt.figure(figsize=(10, 5))
    plt.plot(
        resumen_horario["hora"],
        resumen_horario["demanda_media"],
        marker="o",
    )
    plt.title("Perfil horario medio de la demanda eléctrica peninsular")
    plt.xlabel("Hora del día")
    plt.ylabel("Demanda media [MW]")
    plt.xticks(range(0, 24))
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "perfil_horario_demanda.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def ejecutar_validacion_demanda() -> None:
    """
    Ejecuta el proceso completo de validación de la demanda REE.
    """

    print("\nVALIDACIÓN DE DEMANDA REE")
    print("-" * 50)

    df = cargar_base_modelo()
    df = preparar_base_demanda(df)

    resumen_general = crear_resumen_general_demanda(df)
    resumen_mensual = crear_resumen_mensual_demanda(df)
    resumen_horario = crear_resumen_horario_demanda(df)

    guardar_control_calidad_demanda(
        resumen_general=resumen_general,
        resumen_mensual=resumen_mensual,
        resumen_horario=resumen_horario,
    )

    graficar_evolucion_mensual_demanda(resumen_mensual)
    graficar_perfil_horario_demanda(resumen_horario)

    print("Validación de demanda completada correctamente.")
    print("-" * 50)
