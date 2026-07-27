import pandas as pd

from config import (
    OUTPUT_DIR,
    RUTA_CURVA_CLIENTE,
    ENERGIA_TOTAL_CLIENTE_MWH,
    RUTA_PERFILES_REE_2026,
    PERFIL_REE_CLIENTE,
    RUTA_CURVAS_CLIENTES_SIMULADOS,
)


PERFILES_OFICIALES_REE = {
    "2.0TD": "P2.0TD",
    "3.0TD": "P3.0TD",
    "3.0TDVE": "P3.0TDVE",
}


PERFILES_DISPONIBLES = [
    "2.0TD",
    "3.0TD",
    "3.0TDVE",
    "6.XTD",
]


def cargar_prevision_actual() -> pd.DataFrame:
    """
    Carga la última previsión generada para usar sus horas como base.
    """

    ruta_prevision = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"

    if not ruta_prevision.exists():
        raise FileNotFoundError(
            f"No existe {ruta_prevision}. Ejecuta primero la opción 4 "
            f"para generar la previsión del periodo de oferta."
        )

    df = pd.read_excel(
        ruta_prevision,
        sheet_name="prevision_horaria",
    )

    if "fecha_hora_utc" not in df.columns:
        raise ValueError(
            "La previsión no contiene la columna 'fecha_hora_utc'."
        )

    df["fecha_hora_utc"] = pd.to_datetime(
        df["fecha_hora_utc"],
        utc=True,
        errors="coerce",
    )

    df = df.dropna(subset=["fecha_hora_utc"]).copy()

    # Fecha-hora local peninsular para cruzar con el perfil REE.
    # Esto corrige el problema de que antes todas las horas salían como hora 0.
    df["fecha_hora_local"] = (
        df["fecha_hora_utc"]
        .dt.tz_convert("Europe/Madrid")
        .dt.tz_localize(None)
    )

    return df


def cargar_perfil_ree_2026() -> pd.DataFrame:
    """
    Carga directamente el archivo de perfiles iniciales 2026 de REE.

    El archivo descargado tiene:
    - Hoja: Perfiles Iniciales 2026
    - Cabecera en la fila 2
    - Hora en formato 1-24
    """

    if not RUTA_PERFILES_REE_2026.exists():
        raise FileNotFoundError(
            f"No existe {RUTA_PERFILES_REE_2026}. "
            f"Guarda el archivo descargado de REE como perfiles_iniciales_2026.xlsx."
        )

    df = pd.read_excel(
        RUTA_PERFILES_REE_2026,
        sheet_name="Perfiles Iniciales 2026",
        header=1,
    )

    df.columns = [
        str(columna).strip()
        for columna in df.columns
    ]

    columnas_obligatorias = [
        "Mes",
        "Día",
        "Hora",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"No se encuentra la columna '{columna}' en el archivo REE."
            )

    columnas_perfiles = [
        columna
        for columna in df.columns
        if "P2.0TD" in str(columna)
        or "P3.0TD,0m,d,h" in str(columna)
        or "P3.0TDVE" in str(columna)
    ]

    if not columnas_perfiles:
        raise ValueError(
            "No se han encontrado columnas de perfiles 2.0TD, 3.0TD o 3.0TDVE "
            "en el archivo REE."
        )

    df_base = df[
        [
            "Mes",
            "Día",
            "Hora",
        ]
        + columnas_perfiles
    ].copy()

    df_base = df_base.rename(
        columns={
            "Mes": "mes",
            "Día": "dia",
            "Hora": "hora_ree",
        }
    )

    df_base["mes"] = pd.to_numeric(
        df_base["mes"],
        errors="coerce",
    )

    df_base["dia"] = pd.to_numeric(
        df_base["dia"],
        errors="coerce",
    )

    df_base["hora_ree"] = pd.to_numeric(
        df_base["hora_ree"],
        errors="coerce",
    )

    for columna in columnas_perfiles:
        df_base[columna] = pd.to_numeric(
            df_base[columna],
            errors="coerce",
        )

    df_base = df_base.dropna(
        subset=[
            "mes",
            "dia",
            "hora_ree",
        ]
    ).copy()

    df_base["mes"] = df_base["mes"].astype(int)
    df_base["dia"] = df_base["dia"].astype(int)
    df_base["hora_ree"] = df_base["hora_ree"].astype(int)

    # REE usa horas 1-24. El modelo trabaja con 0-23.
    df_base["hora"] = df_base["hora_ree"] - 1

    print(
        "Perfil REE cargado correctamente. "
        f"Columnas detectadas: {', '.join(columnas_perfiles)}"
    )

    return df_base


def encontrar_columna_perfil(
    df_perfil_ree: pd.DataFrame,
    perfil: str,
) -> str:
    """
    Encuentra la columna del perfil oficial REE correspondiente.
    """

    if perfil not in PERFILES_OFICIALES_REE:
        raise ValueError(
            f"El perfil {perfil} no es oficial REE en este archivo."
        )

    texto_busqueda = PERFILES_OFICIALES_REE[perfil].lower()

    columnas = [
        columna
        for columna in df_perfil_ree.columns
        if texto_busqueda in str(columna).lower()
    ]

    if not columnas:
        raise ValueError(
            f"No se ha encontrado la columna del perfil {perfil}."
        )

    return columnas[0]


def crear_tabla_coeficientes_perfil(
    df_perfil_ree: pd.DataFrame,
    perfil: str,
) -> pd.DataFrame:
    """
    Crea una tabla mes/día/hora/coeficiente para el perfil elegido.
    """

    if perfil in PERFILES_OFICIALES_REE:
        columna_perfil = encontrar_columna_perfil(
            df_perfil_ree=df_perfil_ree,
            perfil=perfil,
        )

        df = df_perfil_ree[
            [
                "mes",
                "dia",
                "hora",
                columna_perfil,
            ]
        ].copy()

        df = df.rename(
            columns={
                columna_perfil: "coeficiente_perfil",
            }
        )

        df["origen_perfil"] = "Perfil oficial REE"

        return df

    if perfil == "6.XTD":
        return crear_perfil_6xtd_sintetico(df_perfil_ree)

    raise ValueError(
        f"Perfil no reconocido: {perfil}. "
        f"Opciones: {PERFILES_DISPONIBLES}"
    )


def crear_perfil_6xtd_sintetico(
    df_perfil_ree: pd.DataFrame,
) -> pd.DataFrame:
    """
    Crea un perfil sintético para 6.XTD.

    El archivo oficial REE usado contiene perfiles 2.0TD, 3.0TD y 3.0TDVE,
    pero no un perfil 6.XTD. Para simular un consumidor industrial,
    se toma como base el perfil 3.0TD y se ajusta para representar:
    - Mayor estabilidad horaria.
    - Más peso en días laborables.
    - Menor caída nocturna que un cliente comercial.
    """

    columna_3td = encontrar_columna_perfil(
        df_perfil_ree=df_perfil_ree,
        perfil="3.0TD",
    )

    df = df_perfil_ree[
        [
            "mes",
            "dia",
            "hora",
            columna_3td,
        ]
    ].copy()

    df = df.rename(
        columns={
            columna_3td: "coeficiente_base_3td",
        }
    )

    # Creamos una fecha ficticia de 2026 para obtener día de la semana.
    df["fecha_ficticia"] = pd.to_datetime(
        {
            "year": 2026,
            "month": df["mes"],
            "day": df["dia"],
        },
        errors="coerce",
    )

    df["dia_semana"] = df["fecha_ficticia"].dt.dayofweek

    # Factor horario industrial:
    # consumo más estable que 3.0TD, pero con mayor peso en horario productivo.
    def factor_horario_industrial(hora: int) -> float:
        if 0 <= hora <= 5:
            return 0.85
        if 6 <= hora <= 7:
            return 1.00
        if 8 <= hora <= 18:
            return 1.18
        if 19 <= hora <= 22:
            return 1.00
        return 0.90

    # Factor semanal:
    # fuerte actividad de lunes a viernes, menor en fin de semana.
    def factor_semana_industrial(dia_semana: int) -> float:
        if dia_semana <= 4:
            return 1.12
        if dia_semana == 5:
            return 0.72
        return 0.55

    # Factor mensual:
    # ligera caída en agosto y algo más de consumo en invierno.
    def factor_mensual_industrial(mes: int) -> float:
        factores = {
            1: 1.08,
            2: 1.06,
            3: 1.03,
            4: 1.00,
            5: 1.00,
            6: 1.02,
            7: 0.95,
            8: 0.78,
            9: 1.00,
            10: 1.04,
            11: 1.07,
            12: 1.03,
        }

        return factores.get(mes, 1.0)

    df["factor_horario_industrial"] = df["hora"].apply(
        factor_horario_industrial
    )

    df["factor_semana_industrial"] = df["dia_semana"].apply(
        factor_semana_industrial
    )

    df["factor_mensual_industrial"] = df["mes"].apply(
        factor_mensual_industrial
    )

    df["coeficiente_perfil"] = (
        df["coeficiente_base_3td"]
        * df["factor_horario_industrial"]
        * df["factor_semana_industrial"]
        * df["factor_mensual_industrial"]
    )

    df["origen_perfil"] = (
        "Perfil sintético industrial 6.XTD basado en P3.0TD REE"
    )

    df = df[
        [
            "mes",
            "dia",
            "hora",
            "coeficiente_perfil",
            "origen_perfil",
        ]
    ].copy()

    return df


def generar_curva_horaria(
    df_horas: pd.DataFrame,
    df_coeficientes: pd.DataFrame,
    energia_total_mwh: float,
) -> pd.DataFrame:
    """
    Genera una curva horaria escalada a la energía total del cliente.

    Si alguna hora no cruza exactamente con el perfil, se rellena con la media
    del mismo mes y misma hora. Esto resuelve desajustes por cambio horario
    entre el patrón 2026 y ofertas que incluyen 2027.
    """

    df = df_horas.copy()

    df["mes"] = df["fecha_hora_local"].dt.month
    df["dia"] = df["fecha_hora_local"].dt.day
    df["hora"] = df["fecha_hora_local"].dt.hour

    df = df.merge(
        df_coeficientes,
        on=[
            "mes",
            "dia",
            "hora",
        ],
        how="left",
    )

    df["criterio_coeficiente_perfil"] = "Cruce directo mes-día-hora"

    horas_sin_perfil = df["coeficiente_perfil"].isna().sum()

    if horas_sin_perfil > 0:
        print("\nHoras sin cruce directo de perfil:")
        print(
            df.loc[
                df["coeficiente_perfil"].isna(),
                [
                    "fecha_hora_local",
                    "mes",
                    "dia",
                    "hora",
                ],
            ].head(30).to_string(index=False)
        )

        print(
            f"\nSe rellenan {horas_sin_perfil} horas usando la media "
            f"del mismo mes y hora del perfil."
        )

        fallback_mes_hora = (
            df_coeficientes.groupby(
                [
                    "mes",
                    "hora",
                ],
                as_index=False,
            )
            .agg(
                coeficiente_fallback_mes_hora=("coeficiente_perfil", "mean")
            )
        )

        df = df.merge(
            fallback_mes_hora,
            on=[
                "mes",
                "hora",
            ],
            how="left",
        )

        mascara_fallback = df["coeficiente_perfil"].isna()

        df.loc[
            mascara_fallback,
            "coeficiente_perfil",
        ] = df.loc[
            mascara_fallback,
            "coeficiente_fallback_mes_hora",
        ]

        df.loc[
            mascara_fallback,
            "criterio_coeficiente_perfil",
        ] = "Relleno por media del mismo mes y hora"

        if "origen_perfil" in df.columns:
            df.loc[
                mascara_fallback,
                "origen_perfil",
            ] = df.loc[
                mascara_fallback,
                "origen_perfil",
            ].fillna("Perfil con relleno por media mes-hora")

        df = df.drop(
            columns=[
                "coeficiente_fallback_mes_hora",
            ]
        )

        horas_sin_perfil_final = df["coeficiente_perfil"].isna().sum()

        if horas_sin_perfil_final > 0:
            raise ValueError(
                f"Después del relleno siguen quedando "
                f"{horas_sin_perfil_final} horas sin coeficiente de perfil."
            )

    suma_coeficientes = df["coeficiente_perfil"].sum()

    if suma_coeficientes <= 0:
        raise ValueError(
            "La suma de coeficientes del perfil es cero."
        )

    df["consumo_MWh"] = (
        df["coeficiente_perfil"]
        / suma_coeficientes
        * energia_total_mwh
    )

    df_salida = df[
        [
            "fecha_hora_utc",
            "fecha_hora_local",
            "consumo_MWh",
            "coeficiente_perfil",
            "criterio_coeficiente_perfil",
            "origen_perfil",
        ]
    ].copy()

    # Excel no admite fechas con zona horaria.
    df_salida["fecha_hora_utc"] = pd.to_datetime(
        df_salida["fecha_hora_utc"],
        utc=True,
        errors="coerce",
    ).dt.tz_localize(None)

    df_salida["fecha_hora_local"] = pd.to_datetime(
        df_salida["fecha_hora_local"],
        errors="coerce",
    )

    # Para compatibilidad con decision_cobertura.py,
    # dejamos también una columna llamada fecha.
    df_salida["fecha"] = df_salida["fecha_hora_local"]

    df_salida = df_salida[
        [
            "fecha_hora_utc",
            "fecha_hora_local",
            "fecha",
            "consumo_MWh",
            "coeficiente_perfil",
            "criterio_coeficiente_perfil",
            "origen_perfil",
        ]
    ].copy()

    return df_salida


def calcular_resumen_curva(
    df_curva: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Calcula resúmenes mensual, horario y por día de la semana.
    """

    df = df_curva.copy()

    df["fecha_hora_local"] = pd.to_datetime(df["fecha_hora_local"])

    df["año_mes"] = df["fecha_hora_local"].dt.to_period("M").astype(str)
    df["hora"] = df["fecha_hora_local"].dt.hour
    df["dia_semana"] = df["fecha_hora_local"].dt.dayofweek

    resumen_mensual = (
        df.groupby("año_mes", as_index=False)
        .agg(
            energia_MWh=("consumo_MWh", "sum"),
            consumo_medio_horario_MWh=("consumo_MWh", "mean"),
            horas=("consumo_MWh", "count"),
        )
    )

    resumen_mensual["peso_energia"] = (
        resumen_mensual["energia_MWh"]
        / resumen_mensual["energia_MWh"].sum()
    )

    resumen_horario = (
        df.groupby("hora", as_index=False)
        .agg(
            energia_MWh=("consumo_MWh", "sum"),
            consumo_medio_MWh=("consumo_MWh", "mean"),
            horas=("consumo_MWh", "count"),
        )
    )

    resumen_horario["peso_energia"] = (
        resumen_horario["energia_MWh"]
        / resumen_horario["energia_MWh"].sum()
    )

    resumen_dia_semana = (
        df.groupby("dia_semana", as_index=False)
        .agg(
            energia_MWh=("consumo_MWh", "sum"),
            consumo_medio_MWh=("consumo_MWh", "mean"),
            horas=("consumo_MWh", "count"),
        )
    )

    resumen_dia_semana["peso_energia"] = (
        resumen_dia_semana["energia_MWh"]
        / resumen_dia_semana["energia_MWh"].sum()
    )

    return resumen_mensual, resumen_horario, resumen_dia_semana


def guardar_curvas_todos_perfiles(
    curvas: dict,
) -> None:
    """
    Guarda todas las curvas generadas en un único Excel.
    """

    try:
        with pd.ExcelWriter(
            RUTA_CURVAS_CLIENTES_SIMULADOS,
            engine="openpyxl",
        ) as writer:
            for perfil, datos in curvas.items():
                nombre = perfil.replace(".", "_").replace("/", "_")

                datos["curva"].to_excel(
                    writer,
                    sheet_name=f"curva_{nombre}"[:31],
                    index=False,
                )

                datos["resumen_mensual"].to_excel(
                    writer,
                    sheet_name=f"mensual_{nombre}"[:31],
                    index=False,
                )

                datos["resumen_horario"].to_excel(
                    writer,
                    sheet_name=f"horario_{nombre}"[:31],
                    index=False,
                )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_CURVAS_CLIENTES_SIMULADOS}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Curvas simuladas guardadas en: {RUTA_CURVAS_CLIENTES_SIMULADOS}")


def guardar_curva_cliente_seleccionada(
    df_curva: pd.DataFrame,
    resumen_mensual: pd.DataFrame,
    resumen_horario: pd.DataFrame,
    resumen_dia_semana: pd.DataFrame,
) -> None:
    """
    Guarda la curva seleccionada como curva_cliente.xlsx.
    """

    try:
        with pd.ExcelWriter(RUTA_CURVA_CLIENTE, engine="openpyxl") as writer:
            df_curva.to_excel(
                writer,
                sheet_name="curva_horaria",
                index=False,
            )

            resumen_mensual.to_excel(
                writer,
                sheet_name="resumen_mensual",
                index=False,
            )

            resumen_horario.to_excel(
                writer,
                sheet_name="perfil_horario",
                index=False,
            )

            resumen_dia_semana.to_excel(
                writer,
                sheet_name="perfil_dia_semana",
                index=False,
            )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_CURVA_CLIENTE}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Curva seleccionada guardada en: {RUTA_CURVA_CLIENTE}")


def generar_curva_cliente_no_plana() -> None:
    """
    Genera curvas horarias para varios perfiles y deja seleccionada
    la curva definida en PERFIL_REE_CLIENTE.
    """

    print("\nGENERACIÓN DE CURVAS HORARIAS DE CLIENTE")
    print("-" * 50)

    if PERFIL_REE_CLIENTE not in PERFILES_DISPONIBLES:
        raise ValueError(
            f"PERFIL_REE_CLIENTE = {PERFIL_REE_CLIENTE} no es válido. "
            f"Opciones disponibles: {PERFILES_DISPONIBLES}"
        )

    df_horas = cargar_prevision_actual()

    df_perfil_ree = cargar_perfil_ree_2026()

    curvas = {}

    for perfil in PERFILES_DISPONIBLES:
        print(f"\nGenerando curva para perfil {perfil}...")

        df_coeficientes = crear_tabla_coeficientes_perfil(
            df_perfil_ree=df_perfil_ree,
            perfil=perfil,
        )

        df_curva = generar_curva_horaria(
            df_horas=df_horas,
            df_coeficientes=df_coeficientes,
            energia_total_mwh=ENERGIA_TOTAL_CLIENTE_MWH,
        )

        resumen_mensual, resumen_horario, resumen_dia_semana = calcular_resumen_curva(
            df_curva
        )

        curvas[perfil] = {
            "curva": df_curva,
            "resumen_mensual": resumen_mensual,
            "resumen_horario": resumen_horario,
            "resumen_dia_semana": resumen_dia_semana,
        }

        print(f"Energía total {perfil}: {df_curva['consumo_MWh'].sum():,.2f} MWh")

    guardar_curvas_todos_perfiles(curvas)

    datos_seleccionados = curvas[PERFIL_REE_CLIENTE]

    guardar_curva_cliente_seleccionada(
        df_curva=datos_seleccionados["curva"],
        resumen_mensual=datos_seleccionados["resumen_mensual"],
        resumen_horario=datos_seleccionados["resumen_horario"],
        resumen_dia_semana=datos_seleccionados["resumen_dia_semana"],
    )

    print(f"\nPerfil seleccionado para curva_cliente.xlsx: {PERFIL_REE_CLIENTE}")
    print(f"Energía total generada: {datos_seleccionados['curva']['consumo_MWh'].sum():,.2f} MWh")
    print(f"Fecha inicio curva: {datos_seleccionados['curva']['fecha_hora_local'].min()}")
    print(f"Fecha fin curva: {datos_seleccionados['curva']['fecha_hora_local'].max()}")

    print("\nRESUMEN MENSUAL DEL PERFIL SELECCIONADO")
    print(
        datos_seleccionados["resumen_mensual"][
            [
                "año_mes",
                "energia_MWh",
                "peso_energia",
            ]
        ].round(4).to_string(index=False)
    )

    print("\nPERFIL HORARIO MEDIO DEL PERFIL SELECCIONADO")
    print(
        datos_seleccionados["resumen_horario"][
            [
                "hora",
                "consumo_medio_MWh",
                "peso_energia",
            ]
        ].round(4).to_string(index=False)
    )

    print("Curvas horarias de cliente generadas correctamente.")
    print("-" * 50)