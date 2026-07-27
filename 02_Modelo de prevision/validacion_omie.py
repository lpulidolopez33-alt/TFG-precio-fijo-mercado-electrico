# Esto sirve para:

# 1. Si hay horas sin precio OMIE.
# 2. Si hay horas duplicadas.
# 3. Cuál es el precio mínimo, máximo y medio.
# 4. Cuántos días venían con datos horarios y cuántos con datos cuarto-horarios.
# 5. Cómo evoluciona el precio mensual.
# 6. Cómo cambia el precio medio por hora del día.

import pandas as pd
import matplotlib.pyplot as plt

from config import PROCESSED_DATA_DIR, GRAPH_DIR


def cargar_base_modelo() -> pd.DataFrame:
    """
    Carga la base del modelo con los precios OMIE ya integrados.
    """

    ruta_base = PROCESSED_DATA_DIR / "base_modelo.xlsx"

    if not ruta_base.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_base}")

    df = pd.read_excel(ruta_base)

    return df


def preparar_base_omie(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la base OMIE para validaciones y gráficas.
    """

    df = df.copy()

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["fecha_hora_utc_dt"] = pd.to_datetime(df["fecha_hora_utc"], utc=True)
    df["precio_omie"] = pd.to_numeric(df["precio_omie"], errors="coerce")

    return df


def crear_resumen_general(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea un resumen general de calidad de los precios OMIE.
    """

    total_filas = len(df)
    filas_con_precio = df["precio_omie"].notna().sum()
    filas_sin_precio = df["precio_omie"].isna().sum()
    horas_duplicadas_utc = df["fecha_hora_utc"].duplicated().sum()

    resumen = {
        "total_filas": total_filas,
        "filas_con_precio_omie": filas_con_precio,
        "filas_sin_precio_omie": filas_sin_precio,
        "horas_utc_duplicadas": horas_duplicadas_utc,
        "precio_minimo": df["precio_omie"].min(),
        "precio_maximo": df["precio_omie"].max(),
        "precio_medio": df["precio_omie"].mean(),
        "precio_mediana": df["precio_omie"].median(),
        "numero_precios_negativos": (df["precio_omie"] < 0).sum(),
        "numero_precios_cero": (df["precio_omie"] == 0).sum(),
    }

    df_resumen = pd.DataFrame([resumen])

    return df_resumen


def crear_resumen_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el precio OMIE medio mensual e identifica si el mes está completo.
    """

    df = df.copy()

    df["año_mes"] = df["fecha"].dt.to_period("M").astype(str)

    resumen_mensual = (
        df.groupby("año_mes", as_index=False)
        .agg(
            precio_medio_omie=("precio_omie", "mean"),
            precio_minimo_omie=("precio_omie", "min"),
            precio_maximo_omie=("precio_omie", "max"),
            horas_con_dato=("precio_omie", "count"),
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


def crear_resumen_horario(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el precio OMIE medio por hora del día.
    """

    resumen_horario = (
        df.groupby("hora", as_index=False)
        .agg(
            precio_medio_omie=("precio_omie", "mean"),
            precio_minimo_omie=("precio_omie", "min"),
            precio_maximo_omie=("precio_omie", "max"),
            horas_con_dato=("precio_omie", "count"),
        )
    )

    return resumen_horario


def crear_resumen_resolucion_omie(df: pd.DataFrame) -> pd.DataFrame:
    """
    Resume cuántas horas proceden de ficheros OMIE horarios o cuarto-horarios.
    """

    if "numero_valores_originales_omie" not in df.columns:
        return pd.DataFrame()

    resumen_resolucion = (
        df.groupby("numero_valores_originales_omie", as_index=False)
        .agg(
            numero_horas=("precio_omie", "count"),
            precio_medio_omie=("precio_omie", "mean"),
        )
    )

    return resumen_resolucion


def guardar_control_calidad(
    resumen_general: pd.DataFrame,
    resumen_mensual: pd.DataFrame,
    resumen_horario: pd.DataFrame,
    resumen_resolucion: pd.DataFrame,
) -> None:
    """
    Guarda los resúmenes de control de calidad en un archivo Excel.
    """

    ruta_salida = PROCESSED_DATA_DIR / "control_calidad_omie.xlsx"

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        resumen_general.to_excel(writer, sheet_name="resumen_general", index=False)
        resumen_mensual.to_excel(writer, sheet_name="resumen_mensual", index=False)
        resumen_horario.to_excel(writer, sheet_name="resumen_horario", index=False)

        if not resumen_resolucion.empty:
            resumen_resolucion.to_excel(writer, sheet_name="resolucion_omie", index=False)

    print(f"Control de calidad OMIE guardado en: {ruta_salida}")


def graficar_evolucion_mensual(resumen_mensual: pd.DataFrame) -> None:
    """
    Genera una gráfica con la evolución mensual del precio OMIE.

    Para evitar interpretaciones erróneas, solo representa los meses completos.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    resumen_grafica = resumen_mensual[
        resumen_mensual["mes_completo"] == 1
    ].copy()

    plt.figure(figsize=(12, 5))
    plt.plot(
        resumen_grafica["año_mes"],
        resumen_grafica["precio_medio_omie"],
        marker="o",
    )
    plt.xticks(rotation=90)
    plt.title("Evolución mensual del precio OMIE")
    plt.xlabel("Mes")
    plt.ylabel("Precio medio OMIE [€/MWh]")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "evolucion_mensual_omie.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_perfil_horario(resumen_horario: pd.DataFrame) -> None:
    """
    Genera una gráfica con el precio medio OMIE por hora del día.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    plt.figure(figsize=(10, 5))
    plt.plot(
        resumen_horario["hora"],
        resumen_horario["precio_medio_omie"],
        marker="o",
    )
    plt.title("Perfil horario medio del precio OMIE")
    plt.xlabel("Hora del día")
    plt.ylabel("Precio medio OMIE [€/MWh]")
    plt.xticks(range(0, 24))
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "perfil_horario_omie.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def ejecutar_validacion_omie() -> None:
    """
    Ejecuta el proceso completo de validación de precios OMIE.
    """

    print("\nVALIDACIÓN DE PRECIOS OMIE")
    print("-" * 50)

    df = cargar_base_modelo()
    df = preparar_base_omie(df)

    resumen_general = crear_resumen_general(df)
    resumen_mensual = crear_resumen_mensual(df)
    resumen_horario = crear_resumen_horario(df)
    resumen_resolucion = crear_resumen_resolucion_omie(df)

    guardar_control_calidad(
        resumen_general=resumen_general,
        resumen_mensual=resumen_mensual,
        resumen_horario=resumen_horario,
        resumen_resolucion=resumen_resolucion,
    )

    graficar_evolucion_mensual(resumen_mensual)
    graficar_perfil_horario(resumen_horario)

    print("Validación OMIE completada correctamente.")
    print("-" * 50)