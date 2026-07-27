import numpy as np
import pandas as pd

from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score

from config import (
    PROCESSED_DATA_DIR,
    RUTA_BACKTESTING_TEMPORAL,
    COLUMNA_OBJETIVO_MODELO,
)


def pedir_fechas_backtesting() -> tuple[pd.Timestamp, pd.Timestamp]:
    """
    Pide al usuario el periodo histórico que quiere validar.
    """

    print("\nCONFIGURACIÓN DE BACKTESTING TEMPORAL")
    print("-" * 50)

    fecha_inicio_txt = input(
        "Introduce fecha inicio del periodo a validar [2025-06-01]: "
    ).strip()

    fecha_fin_txt = input(
        "Introduce fecha fin del periodo a validar [2025-06-30]: "
    ).strip()

    if fecha_inicio_txt == "":
        fecha_inicio_txt = "2025-06-01"

    if fecha_fin_txt == "":
        fecha_fin_txt = "2025-06-30"

    fecha_inicio = pd.to_datetime(
        fecha_inicio_txt,
        dayfirst=True,
        errors="raise",
    )

    fecha_fin = pd.to_datetime(
        fecha_fin_txt,
        dayfirst=True,
        errors="raise",
    )

    if fecha_fin < fecha_inicio:
        raise ValueError("La fecha fin no puede ser anterior a la fecha inicio.")

    print(f"Fecha inicio validación: {fecha_inicio.date()}")
    print(f"Fecha fin validación: {fecha_fin.date()}")
    print("-" * 50)

    return fecha_inicio, fecha_fin


def cargar_base_entrenamiento() -> pd.DataFrame:
    """
    Carga la base de entrenamiento ya preparada.
    """

    ruta_base = PROCESSED_DATA_DIR / "base_entrenamiento.xlsx"

    if not ruta_base.exists():
        raise FileNotFoundError(
            f"No existe {ruta_base}. Ejecuta primero la opción 2."
        )

    df = pd.read_excel(ruta_base)

    return df


def preparar_base_backtesting(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara fechas y ordena la base.
    """

    df = df.copy()

    if "fecha" not in df.columns:
        raise ValueError("Falta la columna 'fecha' en base_entrenamiento.xlsx")

    if "fecha_hora_utc" not in df.columns:
        raise ValueError("Falta la columna 'fecha_hora_utc' en base_entrenamiento.xlsx")

    if COLUMNA_OBJETIVO_MODELO not in df.columns:
        raise ValueError(
            f"Falta la columna objetivo '{COLUMNA_OBJETIVO_MODELO}' "
            f"en base_entrenamiento.xlsx"
        )

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["fecha_hora_utc"] = pd.to_datetime(
        df["fecha_hora_utc"],
        utc=True,
        errors="coerce",
    )

    df = df.sort_values("fecha_hora_utc")

    return df


def obtener_columnas_variables_backtesting(df: pd.DataFrame) -> list[str]:
    """
    Selecciona automáticamente las variables numéricas válidas para entrenar.
    """

    columnas_excluir = {
        COLUMNA_OBJETIVO_MODELO,
        "precio_omie_real",
        "precio_omie_previsto",
        "precio_hibrido_omip_EUR_MWh",
        "precio_corregido_omip_EUR_MWh",
        "error",
        "error_EUR_MWh",
        "error_absoluto_EUR_MWh",
    }

    columnas_fecha = {
        "fecha",
        "fecha_hora_utc",
        "fecha_hora_local",
    }

    columnas_excluir = columnas_excluir.union(columnas_fecha)

    columnas_numericas = df.select_dtypes(
        include=["number", "bool"]
    ).columns.tolist()

    columnas_variables = [
        columna
        for columna in columnas_numericas
        if columna not in columnas_excluir
    ]

    if not columnas_variables:
        raise ValueError("No se han encontrado variables numéricas para entrenar.")

    return columnas_variables


def dividir_backtesting(
    df: pd.DataFrame,
    fecha_inicio_validacion: pd.Timestamp,
    fecha_fin_validacion: pd.Timestamp,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Divide la base en entrenamiento histórico y periodo de validación.
    """

    fecha_corte_entrenamiento = fecha_inicio_validacion - pd.Timedelta(days=1)

    df_train = df[df["fecha"] <= fecha_corte_entrenamiento].copy()

    df_test = df[
        (df["fecha"] >= fecha_inicio_validacion)
        & (df["fecha"] <= fecha_fin_validacion)
    ].copy()

    if df_train.empty:
        raise ValueError(
            "No hay datos de entrenamiento anteriores al periodo de validación."
        )

    if df_test.empty:
        raise ValueError(
            "No hay datos reales para el periodo de validación seleccionado."
        )

    return df_train, df_test


def limpiar_train_test(
    df_train: pd.DataFrame,
    df_test: pd.DataFrame,
    columnas_variables: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Elimina filas con datos faltantes en variables o precio objetivo.
    """

    columnas_necesarias = columnas_variables + [COLUMNA_OBJETIVO_MODELO]

    df_train = df_train.dropna(subset=columnas_necesarias).copy()
    df_test = df_test.dropna(subset=columnas_necesarias).copy()

    if df_train.empty:
        raise ValueError("El entrenamiento queda vacío tras eliminar nulos.")

    if df_test.empty:
        raise ValueError("La validación queda vacía tras eliminar nulos.")

    return df_train, df_test


def entrenar_modelo_backtesting(
    df_train: pd.DataFrame,
    columnas_variables: list[str],
) -> HistGradientBoostingRegressor:
    """
    Entrena un modelo HistGradientBoosting con la configuración optimizada.
    """

    modelo = HistGradientBoostingRegressor(
        l2_regularization=0.0,
        learning_rate=0.03,
        max_iter=1000,
        max_leaf_nodes=31,
        min_samples_leaf=20,
        random_state=42,
    )

    X_train = df_train[columnas_variables]
    y_train = df_train[COLUMNA_OBJETIVO_MODELO]

    modelo.fit(X_train, y_train)

    return modelo


def calcular_predicciones_backtesting(
    modelo: HistGradientBoostingRegressor,
    df_test: pd.DataFrame,
    columnas_variables: list[str],
) -> pd.DataFrame:
    """
    Calcula predicciones y errores sobre el periodo validado.
    """

    df = df_test.copy()

    X_test = df[columnas_variables]

    df["precio_previsto_backtesting"] = modelo.predict(X_test)

    df["precio_real_omie"] = df[COLUMNA_OBJETIVO_MODELO]

    df["error_EUR_MWh"] = (
        df["precio_previsto_backtesting"]
        - df["precio_real_omie"]
    )

    df["error_absoluto_EUR_MWh"] = df["error_EUR_MWh"].abs()

    df["error_cuadratico"] = df["error_EUR_MWh"] ** 2

    return df


def calcular_resumen_global(
    df_validacion: pd.DataFrame,
    df_train: pd.DataFrame,
    columnas_variables: list[str],
) -> pd.DataFrame:
    """
    Calcula métricas globales del backtesting.
    """

    y_real = df_validacion["precio_real_omie"]
    y_pred = df_validacion["precio_previsto_backtesting"]

    mae = mean_absolute_error(y_real, y_pred)
    rmse = np.sqrt(mean_squared_error(y_real, y_pred))
    r2 = r2_score(y_real, y_pred)

    resumen = pd.DataFrame(
        [
            {
                "fecha_inicio_entrenamiento": df_train["fecha"].min().date(),
                "fecha_fin_entrenamiento": df_train["fecha"].max().date(),
                "fecha_inicio_validacion": df_validacion["fecha"].min().date(),
                "fecha_fin_validacion": df_validacion["fecha"].max().date(),
                "horas_entrenamiento": len(df_train),
                "horas_validacion": len(df_validacion),
                "numero_variables": len(columnas_variables),
                "precio_real_medio_EUR_MWh": y_real.mean(),
                "precio_previsto_medio_EUR_MWh": y_pred.mean(),
                "error_medio_EUR_MWh": df_validacion["error_EUR_MWh"].mean(),
                "MAE_EUR_MWh": mae,
                "RMSE_EUR_MWh": rmse,
                "R2": r2,
                "error_maximo_EUR_MWh": df_validacion["error_EUR_MWh"].max(),
                "error_minimo_EUR_MWh": df_validacion["error_EUR_MWh"].min(),
            }
        ]
    )

    return resumen


def calcular_resumen_mensual(df_validacion: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula métricas mensuales del backtesting.
    """

    df = df_validacion.copy()

    df["año_mes"] = df["fecha"].dt.to_period("M").astype(str)

    resumen = (
        df.groupby("año_mes", as_index=False)
        .agg(
            horas=("precio_real_omie", "count"),
            precio_real_medio_EUR_MWh=("precio_real_omie", "mean"),
            precio_previsto_medio_EUR_MWh=("precio_previsto_backtesting", "mean"),
            error_medio_EUR_MWh=("error_EUR_MWh", "mean"),
            MAE_EUR_MWh=("error_absoluto_EUR_MWh", "mean"),
            RMSE_EUR_MWh=("error_cuadratico", lambda x: np.sqrt(x.mean())),
        )
    )

    return resumen


def guardar_backtesting(
    resumen_global: pd.DataFrame,
    resumen_mensual: pd.DataFrame,
    df_validacion: pd.DataFrame,
    columnas_variables: list[str],
) -> None:
    """
    Guarda el resultado del backtesting.
    """

    df_variables = pd.DataFrame(
        {
            "variable": columnas_variables,
        }
    )

    columnas_validacion = [
        "fecha_hora_utc",
        "fecha",
        "precio_real_omie",
        "precio_previsto_backtesting",
        "error_EUR_MWh",
        "error_absoluto_EUR_MWh",
    ]

    columnas_validacion = [
        columna
        for columna in columnas_validacion
        if columna in df_validacion.columns

    ]

            # Excel no admite datetimes con zona horaria.
    df_validacion_excel = df_validacion.copy()

    if "fecha_hora_utc" in df_validacion_excel.columns:
        df_validacion_excel["fecha_hora_utc"] = pd.to_datetime(
            df_validacion_excel["fecha_hora_utc"],
            utc=True,
            errors="coerce",
        ).dt.tz_localize(None)

    if "fecha" in df_validacion_excel.columns:
        df_validacion_excel["fecha"] = pd.to_datetime(
            df_validacion_excel["fecha"],
            errors="coerce",
        ).dt.tz_localize(None)
  

    try:
        with pd.ExcelWriter(RUTA_BACKTESTING_TEMPORAL, engine="openpyxl") as writer:
            resumen_global.to_excel(
                writer,
                sheet_name="resumen_global",
                index=False,
            )

            resumen_mensual.to_excel(
                writer,
                sheet_name="resumen_mensual",
                index=False,
            )

            df_validacion_excel[columnas_validacion].to_excel(
                writer,
                sheet_name="validacion_horaria",
                index=False,
            )

            df_variables.to_excel(
                writer,
                sheet_name="variables_usadas",
                index=False,
            )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_BACKTESTING_TEMPORAL}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Backtesting temporal guardado en: {RUTA_BACKTESTING_TEMPORAL}")


def mostrar_resumen_backtesting(resumen_global: pd.DataFrame) -> None:
    """
    Muestra un resumen limpio en terminal.
    """

    print("\nRESUMEN BACKTESTING TEMPORAL")

    df = resumen_global.copy()

    columnas_mostrar = [
        "fecha_inicio_entrenamiento",
        "fecha_fin_entrenamiento",
        "fecha_inicio_validacion",
        "fecha_fin_validacion",
        "precio_real_medio_EUR_MWh",
        "precio_previsto_medio_EUR_MWh",
        "error_medio_EUR_MWh",
        "MAE_EUR_MWh",
        "RMSE_EUR_MWh",
        "R2",
    ]

    df = df[columnas_mostrar].copy()

    columnas_redondear = [
        "precio_real_medio_EUR_MWh",
        "precio_previsto_medio_EUR_MWh",
        "error_medio_EUR_MWh",
        "MAE_EUR_MWh",
        "RMSE_EUR_MWh",
        "R2",
    ]

    for columna in columnas_redondear:
        df[columna] = df[columna].astype(float).round(3)

    print(df.to_string(index=False))


def ejecutar_backtesting_temporal() -> None:
    """
    Ejecuta el backtesting temporal sin fuga de información.
    """

    print("\nBACKTESTING TEMPORAL DEL MODELO")
    print("-" * 50)

    fecha_inicio_validacion, fecha_fin_validacion = pedir_fechas_backtesting()

    df = cargar_base_entrenamiento()
    df = preparar_base_backtesting(df)

    columnas_variables = obtener_columnas_variables_backtesting(df)

    df_train, df_test = dividir_backtesting(
        df=df,
        fecha_inicio_validacion=fecha_inicio_validacion,
        fecha_fin_validacion=fecha_fin_validacion,
    )

    df_train, df_test = limpiar_train_test(
        df_train=df_train,
        df_test=df_test,
        columnas_variables=columnas_variables,
    )

    print(f"Entrenamiento hasta: {df_train['fecha'].max().date()}")
    print(f"Horas entrenamiento: {len(df_train)}")
    print(f"Horas validación: {len(df_test)}")
    print(f"Variables utilizadas: {len(columnas_variables)}")

    modelo = entrenar_modelo_backtesting(
        df_train=df_train,
        columnas_variables=columnas_variables,
    )

    df_validacion = calcular_predicciones_backtesting(
        modelo=modelo,
        df_test=df_test,
        columnas_variables=columnas_variables,
    )

    resumen_global = calcular_resumen_global(
        df_validacion=df_validacion,
        df_train=df_train,
        columnas_variables=columnas_variables,
    )

    resumen_mensual = calcular_resumen_mensual(df_validacion)

    guardar_backtesting(
        resumen_global=resumen_global,
        resumen_mensual=resumen_mensual,
        df_validacion=df_validacion,
        columnas_variables=columnas_variables,
    )

    mostrar_resumen_backtesting(resumen_global)

    print("Backtesting temporal completado correctamente.")
    print("-" * 50)