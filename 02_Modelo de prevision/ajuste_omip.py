import pandas as pd

from config import (
    OUTPUT_DIR,
    RUTA_COMPARACION_OMIP,
    RUTA_PREVISION_HIBRIDA_OMIP,
    RUTA_RESUMEN_HIBRIDO_OMIP,
    UMBRAL_DIFERENCIA_MODELO_OMIP_BAJO,
    UMBRAL_DIFERENCIA_MODELO_OMIP_ALTO,
    FACTOR_AJUSTE_OMIP_BAJO,
    FACTOR_AJUSTE_OMIP_MEDIO,
    FACTOR_AJUSTE_OMIP_ALTO,
)


def cargar_prevision_modelo() -> pd.DataFrame:
    """
    Carga la previsión horaria generada por el modelo.
    """

    ruta_prevision = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"

    if not ruta_prevision.exists():
        raise FileNotFoundError(
            f"No existe el archivo de previsión: {ruta_prevision}. "
            f"Ejecuta primero la opción 4."
        )

    df = pd.read_excel(
        ruta_prevision,
        sheet_name="prevision_horaria",
    )

    return df


def cargar_comparacion_omip() -> pd.DataFrame:
    """
    Carga la comparación modelo vs OMIP.
    """

    if not RUTA_COMPARACION_OMIP.exists():
        raise FileNotFoundError(
            f"No existe el archivo de comparación OMIP: {RUTA_COMPARACION_OMIP}. "
            f"Ejecuta primero la comparación modelo vs OMIP."
        )

    df = pd.read_excel(
        RUTA_COMPARACION_OMIP,
        sheet_name="comparacion_modelo_omip",
    )

    return df


def preparar_prevision(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la previsión horaria para aplicar el ajuste OMIP.
    """

    df = df.copy()

    if "fecha" not in df.columns:
        raise ValueError("Falta la columna 'fecha' en la previsión del modelo.")

    if "precio_omie_previsto" not in df.columns:
        raise ValueError(
            "Falta la columna 'precio_omie_previsto' en la previsión del modelo."
        )

    df["fecha"] = pd.to_datetime(df["fecha"])

    df["precio_omie_previsto"] = pd.to_numeric(
        df["precio_omie_previsto"],
        errors="coerce",
    )

    return df


def preparar_comparacion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la comparación modelo vs OMIP.
    """

    df = df.copy()

    columnas_obligatorias = [
        "producto",
        "fecha_inicio",
        "fecha_fin",
        "precio_omip_EUR_MWh",
        "precio_modelo_EUR_MWh",
        "diferencia_modelo_menos_omip_EUR_MWh",
        "cobertura",
        "estado",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna '{columna}' en comparacion_modelo_omip.xlsx"
            )

    df["fecha_inicio"] = pd.to_datetime(df["fecha_inicio"])
    df["fecha_fin"] = pd.to_datetime(df["fecha_fin"])

    df["precio_omip_EUR_MWh"] = pd.to_numeric(
        df["precio_omip_EUR_MWh"],
        errors="coerce",
    )

    df["precio_modelo_EUR_MWh"] = pd.to_numeric(
        df["precio_modelo_EUR_MWh"],
        errors="coerce",
    )

    df["diferencia_modelo_menos_omip_EUR_MWh"] = pd.to_numeric(
        df["diferencia_modelo_menos_omip_EUR_MWh"],
        errors="coerce",
    )

    df["cobertura"] = pd.to_numeric(
        df["cobertura"],
        errors="coerce",
    )

    return df


def seleccionar_productos_omip_validos(df_comparacion: pd.DataFrame) -> pd.DataFrame:
    """
    Selecciona los productos OMIP con comparación completa.
    """

    df = df_comparacion.copy()

    df = df[
        (df["precio_omip_EUR_MWh"].notna())
        & (df["precio_modelo_EUR_MWh"].notna())
        & (df["cobertura"] >= 0.99)
    ].copy()

    if df.empty:
        return df

    df["dias_producto"] = (
        df["fecha_fin"].dt.normalize()
        - df["fecha_inicio"].dt.normalize()
    ).dt.days + 1

    df = df.sort_values(
        by=["dias_producto", "fecha_inicio"],
        ascending=[True, True],
    )

    return df


def seleccionar_factor_ajuste(
    diferencia_modelo_menos_omip: float,
) -> tuple[float, str, str]:
    """
    Selecciona el factor de corrección parcial hacia OMIP.

    diferencia_modelo_menos_omip = media_modelo - precio_OMIP
    """

    diferencia_abs = abs(diferencia_modelo_menos_omip)

    if diferencia_abs <= UMBRAL_DIFERENCIA_MODELO_OMIP_BAJO:
        return (
            FACTOR_AJUSTE_OMIP_BAJO,
            "Modelo alineado con OMIP",
            "Se mantiene el modelo puro porque la diferencia con OMIP es reducida.",
        )

    if diferencia_abs <= UMBRAL_DIFERENCIA_MODELO_OMIP_ALTO:
        return (
            FACTOR_AJUSTE_OMIP_MEDIO,
            "Diferencia moderada",
            "Se aplica una corrección parcial moderada hacia OMIP.",
        )

    return (
        FACTOR_AJUSTE_OMIP_ALTO,
        "Alerta por diferencia elevada",
        "Se aplica una corrección parcial elevada hacia OMIP, sin sustituir completamente al modelo.",
    )


def aplicar_ajuste_omip(
    df_prevision: pd.DataFrame,
    df_productos_validos: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Aplica una corrección parcial hacia OMIP a la previsión horaria del modelo.
    """

    df = df_prevision.copy()

    df["precio_corregido_omip_EUR_MWh"] = pd.NA
    df["precio_hibrido_omip_EUR_MWh"] = pd.NA
    df["ajuste_omip_EUR_MWh"] = pd.NA
    df["factor_ajuste_omip"] = pd.NA
    df["producto_omip_aplicado"] = pd.NA
    df["precio_omip_producto_EUR_MWh"] = pd.NA
    df["media_modelo_producto_EUR_MWh"] = pd.NA
    df["diferencia_modelo_menos_omip_EUR_MWh"] = pd.NA

    registros_resumen = []

    if df_productos_validos.empty:
        df["precio_corregido_omip_EUR_MWh"] = df["precio_omie_previsto"]
        df["precio_hibrido_omip_EUR_MWh"] = df["precio_omie_previsto"]
        df["ajuste_omip_EUR_MWh"] = 0.0
        df["factor_ajuste_omip"] = 0.0
        df["producto_omip_aplicado"] = "Sin OMIP aplicable"
        df["criterio_ajuste"] = "Sin producto OMIP aplicable; se mantiene modelo puro."

        resumen = pd.DataFrame(
            [
                {
                    "producto": None,
                    "estado": "Sin producto OMIP aplicable",
                    "precio_modelo_EUR_MWh": df["precio_omie_previsto"].mean(),
                    "precio_omip_EUR_MWh": None,
                    "diferencia_modelo_menos_omip_EUR_MWh": None,
                    "factor_ajuste_omip": 0.0,
                    "ajuste_omip_EUR_MWh": 0.0,
                    "precio_corregido_EUR_MWh": df["precio_omie_previsto"].mean(),
                    "recomendacion": "Modelo puro",
                    "criterio_recomendacion": "No hay referencia OMIP completa para el rango previsto.",
                }
            ]
        )

        return df, resumen

    horas_asignadas = pd.Series(False, index=df.index)

    for _, producto in df_productos_validos.iterrows():
        nombre_producto = producto["producto"]
        fecha_inicio = producto["fecha_inicio"]
        fecha_fin = producto["fecha_fin"]
        precio_omip = producto["precio_omip_EUR_MWh"]

        mascara_periodo = (
            (df["fecha"] >= fecha_inicio)
            & (df["fecha"] <= fecha_fin)
            & (~horas_asignadas)
        )

        df_periodo = df.loc[mascara_periodo].copy()

        if df_periodo.empty:
            continue

        media_modelo = df_periodo["precio_omie_previsto"].mean()

        diferencia_modelo_menos_omip = media_modelo - precio_omip

        factor_ajuste, recomendacion, criterio_recomendacion = seleccionar_factor_ajuste(
            diferencia_modelo_menos_omip=diferencia_modelo_menos_omip
        )

        ajuste_total_hasta_omip = precio_omip - media_modelo

        ajuste_parcial = factor_ajuste * ajuste_total_hasta_omip

        precio_corregido = (
            df.loc[mascara_periodo, "precio_omie_previsto"]
            + ajuste_parcial
        )

        df.loc[mascara_periodo, "precio_corregido_omip_EUR_MWh"] = precio_corregido
        df.loc[mascara_periodo, "precio_hibrido_omip_EUR_MWh"] = precio_corregido
        df.loc[mascara_periodo, "ajuste_omip_EUR_MWh"] = ajuste_parcial
        df.loc[mascara_periodo, "factor_ajuste_omip"] = factor_ajuste
        df.loc[mascara_periodo, "producto_omip_aplicado"] = nombre_producto
        df.loc[mascara_periodo, "precio_omip_producto_EUR_MWh"] = precio_omip
        df.loc[mascara_periodo, "media_modelo_producto_EUR_MWh"] = media_modelo
        df.loc[
            mascara_periodo,
            "diferencia_modelo_menos_omip_EUR_MWh",
        ] = diferencia_modelo_menos_omip

        horas_asignadas.loc[mascara_periodo] = True

        precio_corregido_medio = (
            df.loc[mascara_periodo, "precio_corregido_omip_EUR_MWh"]
            .astype(float)
            .mean()
        )

        registros_resumen.append(
            {
                "producto": nombre_producto,
                "fecha_inicio": fecha_inicio.date(),
                "fecha_fin": fecha_fin.date(),
                "precio_modelo_EUR_MWh": media_modelo,
                "precio_omip_EUR_MWh": precio_omip,
                "diferencia_modelo_menos_omip_EUR_MWh": diferencia_modelo_menos_omip,
                "factor_ajuste_omip": factor_ajuste,
                "ajuste_total_hasta_omip_EUR_MWh": ajuste_total_hasta_omip,
                "ajuste_omip_EUR_MWh": ajuste_parcial,
                "precio_corregido_EUR_MWh": precio_corregido_medio,
                "precio_hibrido_EUR_MWh": precio_corregido_medio,
                "horas_ajustadas": int(mascara_periodo.sum()),
                "recomendacion": recomendacion,
                "criterio_recomendacion": criterio_recomendacion,
            }
        )

    mascara_sin_ajuste = df["precio_corregido_omip_EUR_MWh"].isna()

    df.loc[mascara_sin_ajuste, "precio_corregido_omip_EUR_MWh"] = (
        df.loc[mascara_sin_ajuste, "precio_omie_previsto"]
    )

    df.loc[mascara_sin_ajuste, "precio_hibrido_omip_EUR_MWh"] = (
        df.loc[mascara_sin_ajuste, "precio_omie_previsto"]
    )

    df.loc[mascara_sin_ajuste, "ajuste_omip_EUR_MWh"] = 0.0
    df.loc[mascara_sin_ajuste, "factor_ajuste_omip"] = 0.0
    df.loc[mascara_sin_ajuste, "producto_omip_aplicado"] = "Sin OMIP aplicable"

    df["criterio_ajuste"] = df["producto_omip_aplicado"].apply(
        lambda x: (
            "Modelo corregido parcialmente con OMIP"
            if x != "Sin OMIP aplicable"
            else "Modelo puro por falta de OMIP aplicable"
        )
    )

    resumen = pd.DataFrame(registros_resumen)

    if resumen.empty:
        resumen = pd.DataFrame(
            [
                {
                    "producto": None,
                    "estado": "Sin producto OMIP aplicable",
                    "precio_modelo_EUR_MWh": df["precio_omie_previsto"].mean(),
                    "precio_omip_EUR_MWh": None,
                    "diferencia_modelo_menos_omip_EUR_MWh": None,
                    "factor_ajuste_omip": 0.0,
                    "ajuste_omip_EUR_MWh": 0.0,
                    "precio_corregido_EUR_MWh": df[
                        "precio_corregido_omip_EUR_MWh"
                    ].astype(float).mean(),
                    "recomendacion": "Modelo puro",
                    "criterio_recomendacion": "No hay producto OMIP completo aplicable.",
                }
            ]
        )

    return df, resumen


def calcular_resumen_global(df_prevision_ajustada: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula un resumen global de la previsión pura y corregida.
    """

    df = df_prevision_ajustada.copy()

    resumen = pd.DataFrame(
        [
            {
                "fecha_inicio_prevision": df["fecha"].min().date(),
                "fecha_fin_prevision": df["fecha"].max().date(),
                "commodity_media_modelo_EUR_MWh": df["precio_omie_previsto"].mean(),
                "commodity_media_corregida_omip_EUR_MWh": df[
                    "precio_corregido_omip_EUR_MWh"
                ].astype(float).mean(),
                "commodity_media_hibrida_omip_EUR_MWh": df[
                    "precio_hibrido_omip_EUR_MWh"
                ].astype(float).mean(),
                "ajuste_medio_omip_EUR_MWh": df[
                    "ajuste_omip_EUR_MWh"
                ].astype(float).mean(),
                "factor_ajuste_medio_omip": df[
                    "factor_ajuste_omip"
                ].astype(float).mean(),
                "horas_totales": len(df),
                "horas_con_ajuste_omip": (
                    df["producto_omip_aplicado"] != "Sin OMIP aplicable"
                ).sum(),
                "horas_sin_ajuste_omip": (
                    df["producto_omip_aplicado"] == "Sin OMIP aplicable"
                ).sum(),
            }
        ]
    )

    return resumen


def guardar_ajuste_omip(
    df_prevision_ajustada: pd.DataFrame,
    df_resumen_productos: pd.DataFrame,
    df_resumen_global: pd.DataFrame,
) -> None:
    """
    Guarda la previsión corregida y sus resúmenes.
    """

    try:
        with pd.ExcelWriter(RUTA_PREVISION_HIBRIDA_OMIP, engine="openpyxl") as writer:
            df_prevision_ajustada.to_excel(
                writer,
                sheet_name="prevision_horaria_corregida",
                index=False,
            )

            df_resumen_productos.to_excel(
                writer,
                sheet_name="resumen_productos",
                index=False,
            )

            df_resumen_global.to_excel(
                writer,
                sheet_name="resumen_global",
                index=False,
            )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_PREVISION_HIBRIDA_OMIP}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    try:
        with pd.ExcelWriter(RUTA_RESUMEN_HIBRIDO_OMIP, engine="openpyxl") as writer:
            df_resumen_global.to_excel(
                writer,
                sheet_name="resumen_global",
                index=False,
            )

            df_resumen_productos.to_excel(
                writer,
                sheet_name="resumen_productos",
                index=False,
            )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_RESUMEN_HIBRIDO_OMIP}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Previsión corregida OMIP guardada en: {RUTA_PREVISION_HIBRIDA_OMIP}")
    print(f"Resumen corregido OMIP guardado en: {RUTA_RESUMEN_HIBRIDO_OMIP}")


def mostrar_resumen_ajuste(
    df_resumen_productos: pd.DataFrame,
    df_resumen_global: pd.DataFrame,
) -> None:
    """
    Muestra por pantalla el resumen del ajuste OMIP.
    """

    print("\nRESUMEN GLOBAL CORRECCIÓN OMIP")

    df_global = df_resumen_global.copy()

    columnas_global = {
        "commodity_media_modelo_EUR_MWh": "Modelo",
        "commodity_media_corregida_omip_EUR_MWh": "Corregido",
        "ajuste_medio_omip_EUR_MWh": "Ajuste",
        "factor_ajuste_medio_omip": "Factor",
        "horas_totales": "Horas",
    }

    df_global = df_global[list(columnas_global.keys())].rename(columns=columnas_global)

    for columna in ["Modelo", "Corregido", "Ajuste", "Factor"]:
        df_global[columna] = df_global[columna].astype(float).round(2)

    print(df_global.to_string(index=False))

    print("\nRESUMEN POR PRODUCTO OMIP")

    if df_resumen_productos.empty:
        print("No hay productos OMIP aplicables.")
        return

    df_mostrar = pd.DataFrame(
        {
            "Producto": df_resumen_productos["producto"],
            "Modelo": df_resumen_productos["precio_modelo_EUR_MWh"].astype(float).round(2),
            "OMIP": df_resumen_productos["precio_omip_EUR_MWh"].astype(float).round(2),
            "Dif": df_resumen_productos[
                "diferencia_modelo_menos_omip_EUR_MWh"
            ].astype(float).round(2),
            "Factor": df_resumen_productos["factor_ajuste_omip"].astype(float).round(2),
            "Ajuste": df_resumen_productos["ajuste_omip_EUR_MWh"].astype(float).round(2),
            "Corregido": df_resumen_productos[
                "precio_corregido_EUR_MWh"
            ].astype(float).round(2),
            "Recomendacion": df_resumen_productos["recomendacion"],
        }
    )

    print(df_mostrar.to_string(index=False))


def generar_ajuste_omip() -> None:
    """
    Genera la previsión corregida OMIP-modelo.
    """

    print("\nCORRECCIÓN PARCIAL OMIP-MODELO")
    print("-" * 50)

    df_prevision = cargar_prevision_modelo()
    df_prevision = preparar_prevision(df_prevision)

    df_comparacion = cargar_comparacion_omip()
    df_comparacion = preparar_comparacion(df_comparacion)

    df_productos_validos = seleccionar_productos_omip_validos(df_comparacion)

    df_prevision_ajustada, df_resumen_productos = aplicar_ajuste_omip(
        df_prevision=df_prevision,
        df_productos_validos=df_productos_validos,
    )

    df_resumen_global = calcular_resumen_global(df_prevision_ajustada)

    guardar_ajuste_omip(
        df_prevision_ajustada=df_prevision_ajustada,
        df_resumen_productos=df_resumen_productos,
        df_resumen_global=df_resumen_global,
    )

    mostrar_resumen_ajuste(
        df_resumen_productos=df_resumen_productos,
        df_resumen_global=df_resumen_global,
    )

    print("Corrección parcial OMIP-modelo completada correctamente.")
    print("-" * 50)