import pandas as pd

from config import (
    OUTPUT_DIR,
    RUTA_OMIP_MANUAL,
    RUTA_COMPARACION_OMIP,
)


def cargar_prevision_modelo() -> pd.DataFrame:
    """
    Carga la previsión horaria generada por el modelo.
    """

    ruta_prevision = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"

    if not ruta_prevision.exists():
        raise FileNotFoundError(
            f"No existe el archivo de previsión: {ruta_prevision}. "
            f"Ejecuta primero la opción 4 para generar una previsión."
        )

    df = pd.read_excel(
        ruta_prevision,
        sheet_name="prevision_horaria",
    )

    return df


def cargar_omip_manual() -> pd.DataFrame:
    """
    Carga los precios OMIP introducidos manualmente.
    """

    if not RUTA_OMIP_MANUAL.exists():
        raise FileNotFoundError(
            f"No existe el archivo OMIP manual: {RUTA_OMIP_MANUAL}. "
            f"Crea una copia de plantilla_omip_manual.xlsx y llámala omip_manual.xlsx."
        )

    df = pd.read_excel(RUTA_OMIP_MANUAL)

    return df


def preparar_prevision_modelo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la previsión horaria del modelo.
    """

    df = df.copy()

    columnas_obligatorias = [
        "fecha",
        "precio_omie_previsto",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna '{columna}' en prevision_commodity_modelo.xlsx"
            )

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["precio_omie_previsto"] = pd.to_numeric(
        df["precio_omie_previsto"],
        errors="coerce",
    )

    return df


def preparar_omip_manual(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la tabla manual de OMIP.
    """

    df = df.copy()

    columnas_obligatorias = [
        "producto",
        "fecha_inicio",
        "fecha_fin",
        "precio_omip_EUR_MWh",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna '{columna}' en omip_manual.xlsx"
            )

    df["fecha_inicio"] = pd.to_datetime(df["fecha_inicio"])
    df["fecha_fin"] = pd.to_datetime(df["fecha_fin"])

    df["precio_omip_EUR_MWh"] = pd.to_numeric(
        df["precio_omip_EUR_MWh"],
        errors="coerce",
    )

    df = df.dropna(
        subset=[
            "producto",
            "fecha_inicio",
            "fecha_fin",
            "precio_omip_EUR_MWh",
        ]
    )

    df = df.sort_values("fecha_inicio")

    return df


def calcular_comparacion_omip(
    df_prevision: pd.DataFrame,
    df_omip: pd.DataFrame,
) -> pd.DataFrame:
    """
    Compara la media prevista por el modelo con el precio OMIP de cada producto.
    """

    registros = []

    fecha_min_modelo = df_prevision["fecha"].min()
    fecha_max_modelo = df_prevision["fecha"].max()

    for _, fila in df_omip.iterrows():
        producto = fila["producto"]
        fecha_inicio = fila["fecha_inicio"]
        fecha_fin = fila["fecha_fin"]
        precio_omip = fila["precio_omip_EUR_MWh"]

        mascara = (
            (df_prevision["fecha"] >= fecha_inicio)
            & (df_prevision["fecha"] <= fecha_fin)
        )

        df_periodo = df_prevision.loc[mascara].copy()

        dias_esperados = (
            fecha_fin.normalize()
            - fecha_inicio.normalize()
        ).days + 1

        dias_disponibles = df_periodo["fecha"].dt.date.nunique()

        cobertura = (
            dias_disponibles / dias_esperados
            if dias_esperados > 0
            else 0
        )

        if df_periodo.empty:
            precio_modelo = None
            diferencia = None
            diferencia_pct = None
            estado = "Sin cobertura en la previsión actual"

        else:
            precio_modelo = df_periodo["precio_omie_previsto"].mean()
            diferencia = precio_modelo - precio_omip

            diferencia_pct = (
                diferencia / precio_omip * 100
                if precio_omip != 0
                else None
            )

            if cobertura >= 0.99:
                estado = "Comparación completa"
            else:
                estado = "Comparación parcial"

        registros.append(
            {
                "producto": producto,
                "tipo_producto": fila.get("tipo_producto", None),
                "fecha_inicio": fecha_inicio.date(),
                "fecha_fin": fecha_fin.date(),
                "precio_omip_EUR_MWh": precio_omip,
                "precio_modelo_EUR_MWh": precio_modelo,
                "diferencia_modelo_menos_omip_EUR_MWh": diferencia,
                "diferencia_modelo_menos_omip_%": diferencia_pct,
                "dias_esperados": dias_esperados,
                "dias_disponibles_en_prevision": dias_disponibles,
                "cobertura": cobertura,
                "estado": estado,
                "fecha_minima_prevision_modelo": fecha_min_modelo.date(),
                "fecha_maxima_prevision_modelo": fecha_max_modelo.date(),
                "fecha_cotizacion_omip": fila.get("fecha_cotizacion", None),
                "fuente": fila.get("fuente", "OMIP"),
                "observaciones": fila.get("observaciones", None),
            }
        )

    df_comparacion = pd.DataFrame(registros)

    return df_comparacion


def guardar_comparacion_omip(
    df_comparacion: pd.DataFrame,
    df_omip: pd.DataFrame,
) -> None:
    """
    Guarda la comparación modelo vs OMIP.
    """

    try:
        with pd.ExcelWriter(RUTA_COMPARACION_OMIP, engine="openpyxl") as writer:
            df_comparacion.to_excel(
                writer,
                sheet_name="comparacion_modelo_omip",
                index=False,
            )

            df_omip.to_excel(
                writer,
                sheet_name="omip_manual",
                index=False,
            )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_COMPARACION_OMIP}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Comparación modelo vs OMIP guardada en: {RUTA_COMPARACION_OMIP}")


def mostrar_resumen_comparacion(df_comparacion: pd.DataFrame) -> None:
    """
    Muestra por pantalla el resumen de comparación.
    """

    columnas_resumen = [
        "producto",
        "precio_omip_EUR_MWh",
        "precio_modelo_EUR_MWh",
        "diferencia_modelo_menos_omip_EUR_MWh",
        "estado",
    ]

    print(
        df_comparacion[columnas_resumen].to_string(
            index=False,
            justify="center",
        )
    )


def generar_comparacion_modelo_omip() -> None:
    """
    Genera la comparación entre la previsión del modelo y los precios OMIP.
    """

    print("\nCOMPARACIÓN MODELO VS OMIP")
    print("-" * 50)

    df_prevision = cargar_prevision_modelo()
    df_prevision = preparar_prevision_modelo(df_prevision)

    df_omip = cargar_omip_manual()
    df_omip = preparar_omip_manual(df_omip)

    df_comparacion = calcular_comparacion_omip(
        df_prevision=df_prevision,
        df_omip=df_omip,
    )

    guardar_comparacion_omip(
        df_comparacion=df_comparacion,
        df_omip=df_omip,
    )

    mostrar_resumen_comparacion(df_comparacion)

    print("Comparación modelo vs OMIP completada correctamente.")
    print("-" * 50)