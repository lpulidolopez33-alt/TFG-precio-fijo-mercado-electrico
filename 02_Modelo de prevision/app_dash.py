from __future__ import annotations

import os
import re
import sys
import subprocess
from pathlib import Path
from datetime import date

import pandas as pd
import plotly.graph_objects as go
from dash import Dash, html, dcc, dash_table, Input, Output, State, ctx, no_update

import config


# ============================================================
# RUTAS
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = PROJECT_DIR / "config.py"

OUTPUT_DIR = config.OUTPUT_DIR
RUTA_DECISION_COBERTURA = config.RUTA_DECISION_COBERTURA
RUTA_OMIP_MANUAL = config.RUTA_OMIP_MANUAL
RUTA_CURVA_CLIENTE = config.RUTA_CURVA_CLIENTE
MANUAL_DATA_DIR = RUTA_OMIP_MANUAL.parent

RUTA_PREVISION = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"
RUTA_COMPARACION_OMIP = OUTPUT_DIR / "comparacion_modelo_omip.xlsx"
RUTA_RESULTADOS_ENTRENAMIENTO = OUTPUT_DIR / "resultados_entrenamiento_modelos.xlsx"
RUTA_BACKTESTING = OUTPUT_DIR / "backtesting_temporal.xlsx"

RUTA_CURVAS_CLIENTES_SIMULADOS = getattr(
    config,
    "RUTA_CURVAS_CLIENTES_SIMULADOS",
    MANUAL_DATA_DIR / "curvas_clientes_simulados.xlsx",
)


# ============================================================
# APP
# ============================================================

app = Dash(__name__, suppress_callback_exceptions=True)
server = app.server


# ============================================================
# COLORES
# ============================================================

COLORS = {
    "bg": "#0B1120",
    "panel": "#111827",
    "panel_2": "#172033",
    "panel_3": "#1F2A44",
    "border": "#2B3A55",
    "text": "#F8FAFC",
    "muted": "#A7B0C0",
    "blue": "#3B82F6",
    "cyan": "#22D3EE",
    "green": "#22C55E",
    "orange": "#F59E0B",
    "red": "#EF4444",
    "purple": "#8B5CF6",
}


# ============================================================
# UTILIDADES
# ============================================================

def format_number(value, decimals: int = 2) -> str:
    if value is None or pd.isna(value):
        return "—"

    text = f"{float(value):,.{decimals}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def read_excel_safe(path: Path, sheet_name: str | None = None) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()

    try:
        if sheet_name is None:
            return pd.read_excel(path)

        return pd.read_excel(path, sheet_name=sheet_name)

    except Exception:
        return pd.DataFrame()


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    for column in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[column]):
            df[column] = df[column].astype(str)

        elif pd.api.types.is_float_dtype(df[column]):
            df[column] = df[column].round(3)

    return df


def data_table(
    df: pd.DataFrame,
    page_size: int = 12,
    visible_columns: list[str] | None = None,
) -> dash_table.DataTable:
    """
    Crea una tabla interactiva.
    Permite ordenar, filtrar y ocultar columnas desde la cabecera.
    """

    df = clean_dataframe(df)

    if visible_columns:
        visible_columns = [c for c in visible_columns if c in df.columns]
        if visible_columns:
            df = df[visible_columns].copy()

    if df.empty:
        return dash_table.DataTable(
            columns=[],
            data=[],
            style_table={
                "height": "250px",
                "overflowX": "auto",
            },
        )

    return dash_table.DataTable(
        columns=[
            {
                "name": str(column),
                "id": str(column),
                "hideable": True,
            }
            for column in df.columns
        ],
        data=df.to_dict("records"),
        page_size=page_size,
        sort_action="native",
        filter_action="native",
        column_selectable="multi",
        style_table={
            "overflowX": "auto",
            "border": "1px solid #2B3A55",
            "borderRadius": "14px",
        },
        style_header={
            "backgroundColor": "#1F2A44",
            "color": "#F8FAFC",
            "fontWeight": "700",
            "border": "1px solid #2B3A55",
        },
        style_filter={
            "backgroundColor": "#0F172A",
            "color": "#F8FAFC",
            "border": "1px solid #2B3A55",
        },
        style_cell={
            "backgroundColor": "#111827",
            "color": "#E5E7EB",
            "border": "1px solid #26344E",
            "padding": "10px",
            "fontFamily": "Segoe UI",
            "fontSize": "13px",
            "minWidth": "120px",
            "maxWidth": "360px",
            "whiteSpace": "normal",
            "textAlign": "left",
        },
        style_data_conditional=[
            {
                "if": {"row_index": "odd"},
                "backgroundColor": "#152033",
            },
            {
                "if": {"state": "selected"},
                "backgroundColor": "#243B63",
                "border": "1px solid #3B82F6",
            },
        ],
    )


def make_kpi(title: str, value: str, unit: str, accent: str = "blue") -> html.Div:
    return html.Div(
        className=f"kpi-card kpi-{accent}",
        children=[
            html.Div(title, className="kpi-title"),
            html.Div(value, className="kpi-value"),
            html.Div(unit, className="kpi-unit"),
        ],
    )


def make_empty_figure(title: str) -> go.Figure:
    fig = go.Figure()

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,0.35)",
        title=dict(text=title, x=0.02),
        font=dict(color=COLORS["text"]),
        height=360,
        margin=dict(l=40, r=25, t=60, b=40),
    )

    return fig


def style_figure(fig: go.Figure, title: str, height: int = 360) -> go.Figure:
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,0.35)",
        title=dict(text=title, x=0.02),
        font=dict(color=COLORS["text"]),
        height=height,
        margin=dict(l=40, r=25, t=60, b=40),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
    )

    fig.update_xaxes(gridcolor="rgba(255,255,255,0.08)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.08)")

    return fig


# ============================================================
# CONFIGURACIÓN DESDE LA APP
# ============================================================

def replace_or_add_config_var(text: str, var_name: str, literal: str) -> str:
    pattern = rf"^{var_name}\s*=.*$"
    new_line = f"{var_name} = {literal}"

    if re.search(pattern, text, flags=re.MULTILINE):
        return re.sub(pattern, new_line, text, flags=re.MULTILINE)

    return text.rstrip() + "\n\n" + new_line + "\n"


def apply_config(use_curve: bool, profile: str, energy_mwh: float) -> None:
    text = CONFIG_PATH.read_text(encoding="utf-8")

    text = replace_or_add_config_var(
        text,
        "USAR_CURVA_CLIENTE",
        "True" if use_curve else "False",
    )

    text = replace_or_add_config_var(
        text,
        "PERFIL_REE_CLIENTE",
        repr(profile),
    )

    text = replace_or_add_config_var(
        text,
        "ENERGIA_TOTAL_CLIENTE_MWH",
        str(float(energy_mwh)),
    )

    CONFIG_PATH.write_text(text, encoding="utf-8")


# ============================================================
# EJECUCIÓN DE MAIN.PY
# ============================================================

def run_main_option(
    option: str,
    start_date: str | None = None,
    end_date: str | None = None,
    timeout: int = 1800,
) -> tuple[bool, str]:
    inputs = [option]

    if option in ["1", "4"]:
        if start_date:
            inputs.append(start_date)

        if end_date:
            inputs.append(end_date)

    input_text = "\n".join(inputs) + "\n"

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    try:
        process = subprocess.run(
            [sys.executable, "main.py"],
            cwd=PROJECT_DIR,
            input=input_text,
            text=True,
            capture_output=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            env=env,
        )

        log = ""

        if process.stdout:
            log += process.stdout

        if process.stderr:
            log += "\n\nAVISOS / ERRORES:\n" + process.stderr

        return process.returncode == 0, log

    except subprocess.TimeoutExpired:
        return False, "El proceso ha superado el tiempo máximo permitido."

    except Exception as error:
        return False, f"Error ejecutando main.py: {error}"


def run_option_4_workflow(
    start_date: str,
    end_date: str,
    energy_mwh: float,
    customer_mode: str,
    profile: str,
) -> tuple[bool, str]:
    logs = []

    if customer_mode == "Perfil plano":
        apply_config(
            use_curve=False,
            profile=profile,
            energy_mwh=energy_mwh,
        )

        ok, log = run_main_option(
            option="4",
            start_date=start_date,
            end_date=end_date,
        )

        logs.append("OPCIÓN 4 - PERFIL PLANO")
        logs.append(log)

        return ok, "\n\n".join(logs)

    apply_config(
        use_curve=False,
        profile=profile,
        energy_mwh=energy_mwh,
    )

    ok_base, log_base = run_main_option(
        option="4",
        start_date=start_date,
        end_date=end_date,
    )

    logs.append("PASO 1 - PREVISIÓN BASE SIN CURVA")
    logs.append(log_base)

    if not ok_base:
        return False, "\n\n".join(logs)

    ok_curve, log_curve = run_main_option(option="6")

    logs.append("PASO 2 - GENERACIÓN DE CURVA HORARIA")
    logs.append(log_curve)

    if not ok_curve:
        return False, "\n\n".join(logs)

    apply_config(
        use_curve=True,
        profile=profile,
        energy_mwh=energy_mwh,
    )

    ok_final, log_final = run_main_option(
        option="4",
        start_date=start_date,
        end_date=end_date,
    )

    logs.append("PASO 3 - CÁLCULO FINAL CON CURVA")
    logs.append(log_final)

    return ok_final, "\n\n".join(logs)


# ============================================================
# CARGA DE RESULTADOS
# ============================================================

def load_results() -> dict:
    decision = read_excel_safe(RUTA_DECISION_COBERTURA, "resumen_decision")
    plan = read_excel_safe(RUTA_DECISION_COBERTURA, "plan_cobertura")
    products = read_excel_safe(RUTA_DECISION_COBERTURA, "productos_omip_usados")
    detail = read_excel_safe(RUTA_DECISION_COBERTURA, "detalle_horario")
    forecast = read_excel_safe(RUTA_PREVISION, "prevision_horaria")
    curve = read_excel_safe(RUTA_CURVA_CLIENTE, "curva_horaria")
    monthly_curve = read_excel_safe(RUTA_CURVA_CLIENTE, "resumen_mensual")
    hourly_profile = read_excel_safe(RUTA_CURVA_CLIENTE, "perfil_horario")
    omip = read_excel_safe(RUTA_OMIP_MANUAL)

    if plan.empty and not products.empty:
        plan = products.copy()

    return {
        "decision": decision,
        "plan": plan,
        "products": products,
        "detail": detail,
        "forecast": forecast,
        "curve": curve,
        "monthly_curve": monthly_curve,
        "hourly_profile": hourly_profile,
        "omip": omip,
    }


# ============================================================
# EXCELS Y HOJAS
# ============================================================

def get_excel_files() -> dict[str, Path]:
    return {
        "Previsión commodity": RUTA_PREVISION,
        "Decisión commodity y cobertura": RUTA_DECISION_COBERTURA,
        "Comparación modelo vs OMIP": RUTA_COMPARACION_OMIP,
        "Curva cliente": RUTA_CURVA_CLIENTE,
        "Curvas clientes simulados": RUTA_CURVAS_CLIENTES_SIMULADOS,
        "OMIP manual": RUTA_OMIP_MANUAL,
        "Resultados entrenamiento": RUTA_RESULTADOS_ENTRENAMIENTO,
        "Backtesting temporal": RUTA_BACKTESTING,
    }


def get_excel_sheets(path: Path | None) -> list[str]:
    if path is None or not path.exists():
        return []

    try:
        excel = pd.ExcelFile(path)
        return excel.sheet_names
    except Exception:
        return []


def read_selected_excel_table(file_label: str, sheet_name: str) -> pd.DataFrame:
    files = get_excel_files()
    path = files.get(file_label)

    if path is None or not path.exists():
        return pd.DataFrame()

    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


# ============================================================
# GRÁFICAS
# ============================================================

def build_figures(
    data: dict,
    price_chart_type: str,
    cover_chart_type: str,
    consumption_chart_type: str,
) -> dict:
    plan = data["plan"]
    monthly_curve = data["monthly_curve"]
    hourly_profile = data["hourly_profile"]
    forecast = data["forecast"]
    detail = data["detail"]

    # --------------------------------------------------------
    # Cobertura
    # --------------------------------------------------------

    fig_cover = make_empty_figure("Energía cubierta y abierta")

    if not plan.empty:
        product_col = "producto_cobertura" if "producto_cobertura" in plan.columns else None

        if product_col and cover_chart_type in ["bar_stack", "bar_group"]:
            fig_cover = go.Figure()

            if "energia_a_cubrir_MWh" in plan.columns:
                fig_cover.add_trace(
                    go.Bar(
                        x=plan[product_col].astype(str),
                        y=plan["energia_a_cubrir_MWh"],
                        name="Energía cubierta",
                        marker_color=COLORS["blue"],
                    )
                )

            if "energia_abierta_MWh" in plan.columns:
                fig_cover.add_trace(
                    go.Bar(
                        x=plan[product_col].astype(str),
                        y=plan["energia_abierta_MWh"],
                        name="Energía abierta",
                        marker_color=COLORS["orange"],
                    )
                )

            fig_cover.update_layout(
                barmode="stack" if cover_chart_type == "bar_stack" else "group"
            )

            fig_cover = style_figure(
                fig_cover,
                "Energía cubierta y abierta por producto",
            )

        elif cover_chart_type == "donut":
            covered = plan["energia_a_cubrir_MWh"].sum() if "energia_a_cubrir_MWh" in plan.columns else 0
            open_energy = plan["energia_abierta_MWh"].sum() if "energia_abierta_MWh" in plan.columns else 0

            fig_cover = go.Figure(
                data=[
                    go.Pie(
                        labels=["Cubierta", "Abierta"],
                        values=[covered, open_energy],
                        hole=0.62,
                        marker=dict(
                            colors=[
                                COLORS["blue"],
                                COLORS["orange"],
                            ]
                        ),
                    )
                ]
            )

            fig_cover = style_figure(
                fig_cover,
                "Distribución cubierta / abierta",
            )

    # --------------------------------------------------------
    # Modelo vs OMIP
    # --------------------------------------------------------

    fig_model = make_empty_figure("Modelo vs OMIP")

    if not forecast.empty and {"fecha", "precio_omie_previsto"}.issubset(forecast.columns):
        df = forecast.copy()
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        df = df.dropna(subset=["fecha"])

        daily_forecast = (
            df.groupby("fecha", as_index=False)
            .agg(precio_modelo=("precio_omie_previsto", "mean"))
        )

        fig_model = go.Figure()

        if price_chart_type == "area":
            fig_model.add_trace(
                go.Scatter(
                    x=daily_forecast["fecha"],
                    y=daily_forecast["precio_modelo"],
                    mode="lines",
                    name="Modelo",
                    fill="tozeroy",
                    line=dict(color=COLORS["cyan"], width=3),
                )
            )
        else:
            fig_model.add_trace(
                go.Scatter(
                    x=daily_forecast["fecha"],
                    y=daily_forecast["precio_modelo"],
                    mode="lines",
                    name="Modelo",
                    line=dict(color=COLORS["cyan"], width=3),
                )
            )

        if not detail.empty and {"fecha", "precio_omip_EUR_MWh"}.issubset(detail.columns):
            df_detail = detail.copy()
            df_detail["fecha"] = pd.to_datetime(df_detail["fecha"], errors="coerce")

            daily_omip = (
                df_detail.dropna(subset=["precio_omip_EUR_MWh"])
                .groupby("fecha", as_index=False)
                .agg(precio_omip=("precio_omip_EUR_MWh", "mean"))
            )

            if not daily_omip.empty:
                fig_model.add_trace(
                    go.Scatter(
                        x=daily_omip["fecha"],
                        y=daily_omip["precio_omip"],
                        mode="lines",
                        name="OMIP asignado",
                        line=dict(color=COLORS["orange"], width=2, dash="dash"),
                    )
                )

        fig_model = style_figure(fig_model, "Modelo vs OMIP asignado")

    # --------------------------------------------------------
    # Consumo mensual
    # --------------------------------------------------------

    fig_monthly = make_empty_figure("Consumo mensual")

    if not monthly_curve.empty and {"año_mes", "energia_MWh"}.issubset(monthly_curve.columns):
        fig_monthly = go.Figure()

        if consumption_chart_type == "line":
            fig_monthly.add_trace(
                go.Scatter(
                    x=monthly_curve["año_mes"],
                    y=monthly_curve["energia_MWh"],
                    mode="lines+markers",
                    name="Energía mensual",
                    line=dict(color=COLORS["purple"], width=3),
                )
            )

        elif consumption_chart_type == "area":
            fig_monthly.add_trace(
                go.Scatter(
                    x=monthly_curve["año_mes"],
                    y=monthly_curve["energia_MWh"],
                    mode="lines",
                    fill="tozeroy",
                    name="Energía mensual",
                    line=dict(color=COLORS["purple"], width=3),
                )
            )

        elif consumption_chart_type == "donut":
            fig_monthly.add_trace(
                go.Pie(
                    labels=monthly_curve["año_mes"],
                    values=monthly_curve["energia_MWh"],
                    hole=0.55,
                )
            )

        else:
            fig_monthly.add_trace(
                go.Bar(
                    x=monthly_curve["año_mes"],
                    y=monthly_curve["energia_MWh"],
                    name="Energía",
                    marker_color=COLORS["purple"],
                )
            )

        fig_monthly = style_figure(fig_monthly, "Consumo mensual del cliente")

    # --------------------------------------------------------
    # Perfil horario
    # --------------------------------------------------------

    fig_hourly = make_empty_figure("Perfil horario")

    if not hourly_profile.empty and {"hora", "consumo_medio_MWh"}.issubset(hourly_profile.columns):
        fig_hourly = go.Figure()

        fig_hourly.add_trace(
            go.Scatter(
                x=hourly_profile["hora"],
                y=hourly_profile["consumo_medio_MWh"],
                mode="lines+markers",
                name="Consumo medio",
                line=dict(color=COLORS["green"], width=3),
            )
        )

        fig_hourly.update_xaxes(dtick=1)

        fig_hourly = style_figure(fig_hourly, "Perfil horario medio")

    return {
        "cover": fig_cover,
        "model": fig_model,
        "monthly": fig_monthly,
        "hourly": fig_hourly,
    }


# ============================================================
# KPIS
# ============================================================

def build_kpis(decision: pd.DataFrame) -> list:
    if decision.empty:
        return [
            make_kpi("Commodity modelo", "—", "€/MWh", "blue"),
            make_kpi("OMIP equivalente", "—", "€/MWh", "cyan"),
            make_kpi("Commodity recomendada", "—", "€/MWh", "green"),
            make_kpi("Energía cliente", "—", "MWh", "purple"),
            make_kpi("Cobertura total", "—", "%", "blue"),
            make_kpi("Parte abierta", "—", "%", "orange"),
            make_kpi("Energía cubierta", "—", "MWh", "green"),
            make_kpi("Importe recomendado", "—", "€", "purple"),
        ]

    row = decision.iloc[0]

    return [
        make_kpi(
            "Commodity modelo",
            format_number(row.get("commodity_modelo_total_EUR_MWh"), 2),
            "€/MWh",
            "blue",
        ),
        make_kpi(
            "OMIP equivalente",
            format_number(row.get("omip_equivalente_EUR_MWh"), 2),
            "€/MWh",
            "cyan",
        ),
        make_kpi(
            "Commodity recomendada",
            format_number(row.get("commodity_recomendada_EUR_MWh"), 2),
            "€/MWh",
            "green",
        ),
        make_kpi(
            "Energía cliente",
            format_number(row.get("energia_total_cliente_MWh"), 0),
            "MWh",
            "purple",
        ),
        make_kpi(
            "Cobertura total",
            format_number(100 * row.get("porcentaje_cobertura_total", 0), 1),
            "%",
            "blue",
        ),
        make_kpi(
            "Parte abierta",
            format_number(100 * row.get("porcentaje_abierto_total", 0), 1),
            "%",
            "orange",
        ),
        make_kpi(
            "Energía cubierta",
            format_number(row.get("energia_cubierta_total_MWh"), 0),
            "MWh",
            "green",
        ),
        make_kpi(
            "Importe recomendado",
            format_number(row.get("importe_recomendado_total_EUR"), 0),
            "€",
            "purple",
        ),
    ]


# ============================================================
# LAYOUT
# ============================================================

app.layout = html.Div(
    className="app-shell",
    children=[
        dcc.Store(id="refresh-token", data=0),
        dcc.Store(id="log-store", data=""),
        dcc.Store(id="status-store", data="Listo para ejecutar"),

        html.Div(
            className="sidebar",
            children=[
                html.Div(
                    className="logo-row",
                    children=[
                        html.Div("⚡", className="logo"),
                        html.Div(
                            children=[
                                html.Div("Commodity", className="logo-title"),
                                html.Div("Pricing Dashboard", className="logo-subtitle"),
                            ]
                        ),
                    ],
                ),

                html.Div("Parámetros de oferta", className="section-title"),

                html.Div("Periodo de oferta", className="input-label"),

                dcc.DatePickerRange(
                    id="date-range",
                    start_date=date(2026, 7, 1),
                    end_date=date(2027, 6, 30),
                    display_format="DD/MM/YYYY",
                    className="date-picker",
                    minimum_nights=0,
                    with_portal=True,
                    day_size=39,
                ),

                html.Div("Consumo total cliente [MWh]", className="input-label"),

                dcc.Input(
                    id="energy-input",
                    type="number",
                    value=12000,
                    min=1,
                    step=100,
                    className="input-box",
                ),

                html.Div("Tipo de cliente", className="input-label"),

                dcc.RadioItems(
                    id="customer-mode",
                    options=[
                        {"label": "Perfil plano", "value": "Perfil plano"},
                        {"label": "Perfil no plano", "value": "Perfil no plano"},
                    ],
                    value="Perfil no plano",
                    className="radio-box",
                    labelStyle={"display": "block", "marginTop": "8px"},
                ),

                html.Div("Tarifa / perfil", className="input-label"),

                dcc.Dropdown(
                    id="profile-select",
                    options=[
                        {"label": "2.0TD · Doméstico", "value": "2.0TD"},
                        {"label": "3.0TD · Pyme/comercial", "value": "3.0TD"},
                        {"label": "3.0TDVE · Vehículo eléctrico", "value": "3.0TDVE"},
                        {"label": "6.XTD · Industrial sintético", "value": "6.XTD"},
                    ],
                    value="3.0TD",
                    clearable=False,
                    className="dash-dropdown",
                ),

                html.Div("Visualización", className="section-title"),

                html.Div("Gráfico de cobertura", className="input-label"),

                dcc.Dropdown(
                    id="cover-chart-type",
                    options=[
                        {"label": "Barras apiladas", "value": "bar_stack"},
                        {"label": "Barras agrupadas", "value": "bar_group"},
                        {"label": "Dónut cubierta/abierta", "value": "donut"},
                    ],
                    value="bar_stack",
                    clearable=False,
                    className="dash-dropdown",
                ),

                html.Div("Gráfico de consumo", className="input-label"),

                dcc.Dropdown(
                    id="consumption-chart-type",
                    options=[
                        {"label": "Barras", "value": "bar"},
                        {"label": "Línea", "value": "line"},
                        {"label": "Área", "value": "area"},
                        {"label": "Dónut mensual", "value": "donut"},
                    ],
                    value="bar",
                    clearable=False,
                    className="dash-dropdown",
                ),

                html.Div("Gráfico de precio", className="input-label"),

                dcc.Dropdown(
                    id="price-chart-type",
                    options=[
                        {"label": "Línea", "value": "line"},
                        {"label": "Área", "value": "area"},
                    ],
                    value="line",
                    clearable=False,
                    className="dash-dropdown",
                ),

                html.Div("Órdenes del programa", className="section-title"),

                dcc.Loading(
                    type="circle",
                    color="#22D3EE",
                    children=[
                        html.Div(
                            id="execution-feedback",
                            className="execution-feedback",
                            children="Pulsa una opción para ejecutar el programa.",
                        )
                    ],
                ),

                html.Div(
                    className="action-grid",
                    children=[
                        html.Button("1 · Completo", id="btn-1", n_clicks=0, className="small-button"),
                        html.Button("2 · Datos", id="btn-2", n_clicks=0, className="small-button"),
                        html.Button("3 · Entrenar", id="btn-3", n_clicks=0, className="small-button"),
                        html.Button("4 · Previsión", id="btn-4", n_clicks=0, className="small-button primary-small"),
                        html.Button("5 · Backtest", id="btn-5", n_clicks=0, className="small-button"),
                        html.Button("6 · Curvas", id="btn-6", n_clicks=0, className="small-button"),
                    ],
                ),

                html.Button(
                    "Actualizar vista",
                    id="btn-refresh",
                    n_clicks=0,
                    className="secondary-button",
                ),

                html.Div("Notas", className="section-title"),

                html.Div(
                    className="small-note",
                    children=[
                        html.Div("• La app ejecuta las mismas opciones que el terminal."),
                        html.Div("• La opción 4 usa las fechas y parámetros elegidos arriba."),
                        html.Div("• En perfil no plano, la opción 4 genera curva y recalcula la oferta."),
                        html.Div("• Cierra los Excel de salida antes de ejecutar procesos."),
                    ],
                ),
            ],
        ),

        html.Div(
            className="main",
            children=[
                html.Div(
                    className="hero",
                    children=[
                        html.Div(
                            children=[
                                html.H1("Calculadora de Commodity y Cobertura"),
                                html.P(
                                    "Interfaz visual del programa Python: ejecución, terminal, tablas, gráficos y archivos generados."
                                ),
                            ]
                        ),
                        html.Div(id="status-pill", className="status-pill"),
                    ],
                ),

                dcc.Loading(
                    id="loading-status",
                    type="circle",
                    color="#22D3EE",
                    children=[
                        html.Div(id="status-box", className="status-box"),
                    ],
                ),

                html.Div(id="kpi-grid", className="kpi-grid"),

                dcc.Tabs(
                    id="tabs",
                    value="tab-log",
                    className="tabs-container",
                    children=[
                        dcc.Tab(label="Registro", value="tab-log", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Todas las tablas", value="tab-tablas", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Resumen", value="tab-resumen", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Plan de cobertura", value="tab-plan", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Curva cliente", value="tab-curva", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Previsión", value="tab-prevision", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="OMIP", value="tab-omip", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Archivos", value="tab-archivos", className="tab", selected_className="tab-selected"),
                    ],
                ),

                dcc.Loading(
                    id="loading-tabs",
                    type="circle",
                    color="#22D3EE",
                    children=[
                        html.Div(
                            id="tab-content",
                            className="tab-content",
                        ),
                    ],
                ),
            ],
        ),
    ],
)


# ============================================================
# CALLBACK EJECUCIÓN
# ============================================================

@app.callback(
    Output("status-store", "data"),
    Output("log-store", "data"),
    Output("refresh-token", "data"),
    Output("execution-feedback", "children"),
    Input("btn-refresh", "n_clicks"),
    Input("btn-1", "n_clicks"),
    Input("btn-2", "n_clicks"),
    Input("btn-3", "n_clicks"),
    Input("btn-4", "n_clicks"),
    Input("btn-5", "n_clicks"),
    Input("btn-6", "n_clicks"),
    State("date-range", "start_date"),
    State("date-range", "end_date"),
    State("energy-input", "value"),
    State("customer-mode", "value"),
    State("profile-select", "value"),
    State("refresh-token", "data"),
    prevent_initial_call=True,
)
def execute_action(
    n_refresh,
    n_1,
    n_2,
    n_3,
    n_4,
    n_5,
    n_6,
    start_date,
    end_date,
    energy_mwh,
    customer_mode,
    profile,
    token,
):
    trigger = ctx.triggered_id

    if trigger == "btn-refresh":
        message = "Vista actualizada sin ejecutar procesos."
        return "Vista actualizada", message, (token or 0) + 1, message

    if not start_date or not end_date:
        message = "Faltan fechas de oferta."
        return "Faltan fechas", message, (token or 0) + 1, message

    if not energy_mwh or float(energy_mwh) <= 0:
        message = "El consumo total debe ser mayor que cero."
        return "Consumo no válido", message, (token or 0) + 1, message

    if trigger == "btn-1":
        apply_config(
            use_curve=(customer_mode == "Perfil no plano"),
            profile=profile,
            energy_mwh=float(energy_mwh),
        )

        ok, log = run_main_option(
            option="1",
            start_date=start_date,
            end_date=end_date,
        )

        status = "Opción 1 completada" if ok else "Opción 1 con errores"

        return status, log, (token or 0) + 1, status

    if trigger == "btn-2":
        apply_config(
            use_curve=(customer_mode == "Perfil no plano"),
            profile=profile,
            energy_mwh=float(energy_mwh),
        )

        ok, log = run_main_option(option="2")

        status = "Opción 2 completada" if ok else "Opción 2 con errores"

        return status, log, (token or 0) + 1, status

    if trigger == "btn-3":
        apply_config(
            use_curve=(customer_mode == "Perfil no plano"),
            profile=profile,
            energy_mwh=float(energy_mwh),
        )

        ok, log = run_main_option(option="3")

        status = "Opción 3 completada" if ok else "Opción 3 con errores"

        return status, log, (token or 0) + 1, status

    if trigger == "btn-4":
        ok, log = run_option_4_workflow(
            start_date=start_date,
            end_date=end_date,
            energy_mwh=float(energy_mwh),
            customer_mode=customer_mode,
            profile=profile,
        )

        status = "Opción 4 completada" if ok else "Opción 4 con errores"

        return status, log, (token or 0) + 1, status

    if trigger == "btn-5":
        apply_config(
            use_curve=(customer_mode == "Perfil no plano"),
            profile=profile,
            energy_mwh=float(energy_mwh),
        )

        ok, log = run_main_option(option="5")

        status = "Opción 5 completada" if ok else "Opción 5 con errores"

        return status, log, (token or 0) + 1, status

    if trigger == "btn-6":
        apply_config(
            use_curve=False,
            profile=profile,
            energy_mwh=float(energy_mwh),
        )

        ok_base, log_base = run_main_option(
            option="4",
            start_date=start_date,
            end_date=end_date,
        )

        if not ok_base:
            status = "Error generando previsión base"
            return status, log_base, (token or 0) + 1, status

        ok_curve, log_curve = run_main_option(option="6")

        log = (
            "PASO 1 - PREVISIÓN BASE\n\n"
            + log_base
            + "\n\nPASO 2 - GENERACIÓN DE CURVAS\n\n"
            + log_curve
        )

        status = "Opción 6 completada" if ok_curve else "Opción 6 con errores"

        return status, log, (token or 0) + 1, status

    return no_update, no_update, no_update, no_update


# ============================================================
# CALLBACK CABECERA
# ============================================================

@app.callback(
    Output("kpi-grid", "children"),
    Output("status-pill", "children"),
    Output("status-box", "children"),
    Input("refresh-token", "data"),
    Input("status-store", "data"),
)
def update_header(_, status_text):
    data = load_results()
    decision = data["decision"]

    kpis = build_kpis(decision)

    if decision.empty:
        mode = "Sin resultados"
    else:
        mode = str(decision.iloc[0].get("modo_ponderacion", ""))

    status_box = html.Div(
        className="status-message",
        children=[
            html.Div("Estado de ejecución", className="status-title"),
            html.Div(status_text or "Listo para ejecutar", className="status-text"),
        ],
    )

    return kpis, mode, status_box


# ============================================================
# CALLBACK SELECTORES DE TABLAS
# ============================================================

@app.callback(
    Output("sheet-selector", "options"),
    Output("sheet-selector", "value"),
    Input("file-selector", "value"),
    prevent_initial_call=True,
)
def update_sheet_selector(file_label):
    files = get_excel_files()
    path = files.get(file_label)

    sheets = get_excel_sheets(path)

    options = [
        {
            "label": sheet,
            "value": sheet,
        }
        for sheet in sheets
    ]

    value = sheets[0] if sheets else None

    return options, value


@app.callback(
    Output("column-selector", "options"),
    Output("column-selector", "value"),
    Input("file-selector", "value"),
    Input("sheet-selector", "value"),
    prevent_initial_call=True,
)
def update_column_selector(file_label, sheet_name):
    if not file_label or not sheet_name:
        return [], []

    df = read_selected_excel_table(file_label, sheet_name)

    if df.empty:
        return [], []

    options = [
        {
            "label": column,
            "value": column,
        }
        for column in df.columns
    ]

    value = list(df.columns)

    return options, value


@app.callback(
    Output("selected-table-container", "children"),
    Input("file-selector", "value"),
    Input("sheet-selector", "value"),
    Input("column-selector", "value"),
    prevent_initial_call=True,
)
def update_selected_table(file_label, sheet_name, visible_columns):
    if not file_label or not sheet_name:
        return html.P("Selecciona un archivo y una hoja.", className="small-note")

    df = read_selected_excel_table(file_label, sheet_name)

    if df.empty:
        return html.P(
            "La hoja seleccionada está vacía o no se ha podido leer.",
            className="small-note",
        )

    return html.Div(
        children=[
            html.H4(f"{file_label} · {sheet_name}"),
            data_table(df, page_size=15, visible_columns=visible_columns),
        ]
    )


# ============================================================
# CALLBACK TABS
# ============================================================

@app.callback(
    Output("tab-content", "children"),
    Input("tabs", "value"),
    Input("refresh-token", "data"),
    Input("cover-chart-type", "value"),
    Input("consumption-chart-type", "value"),
    Input("price-chart-type", "value"),
    State("log-store", "data"),
)
def render_tab(
    tab_value,
    _,
    cover_chart_type,
    consumption_chart_type,
    price_chart_type,
    log_text,
):
    data = load_results()

    figures = build_figures(
        data=data,
        price_chart_type=price_chart_type,
        cover_chart_type=cover_chart_type,
        consumption_chart_type=consumption_chart_type,
    )

    if tab_value == "tab-log":
        return html.Div(
            className="card",
            children=[
                html.H3("Registro de ejecución"),
                html.P(
                    "Aquí aparece la salida de Python al ejecutar procesos desde la app. Es la versión visual del terminal.",
                    className="small-note",
                ),
                html.Pre(
                    log_text or "Todavía no se ha ejecutado ningún proceso desde la app.",
                    className="log-box",
                ),
            ],
        )

    if tab_value == "tab-tablas":
        files = get_excel_files()

        file_options = [
            {
                "label": name,
                "value": name,
            }
            for name, path in files.items()
            if path.exists()
        ]

        if not file_options:
            return html.Div(
                className="card",
                children=[
                    html.H3("Todas las tablas"),
                    html.P(
                        "Todavía no hay archivos Excel generados.",
                        className="small-note",
                    ),
                ],
            )

        default_file = file_options[0]["value"]
        default_sheets = get_excel_sheets(files[default_file])
        default_sheet = default_sheets[0] if default_sheets else None

        df_default = (
            read_selected_excel_table(default_file, default_sheet)
            if default_sheet
            else pd.DataFrame()
        )

        default_columns = list(df_default.columns) if not df_default.empty else []

        return html.Div(
            className="card",
            children=[
                html.H3("Todas las tablas generadas por el programa"),
                html.P(
                    "Selecciona cualquier Excel, hoja y columnas visibles. Así puedes revisar todos los datos sin abrir Excel.",
                    className="small-note",
                ),

                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(
                            children=[
                                html.Div("Archivo", className="input-label"),
                                dcc.Dropdown(
                                    id="file-selector",
                                    options=file_options,
                                    value=default_file,
                                    clearable=False,
                                    className="dash-dropdown",
                                ),
                            ]
                        ),
                        html.Div(
                            children=[
                                html.Div("Hoja", className="input-label"),
                                dcc.Dropdown(
                                    id="sheet-selector",
                                    options=[
                                        {
                                            "label": sheet,
                                            "value": sheet,
                                        }
                                        for sheet in default_sheets
                                    ],
                                    value=default_sheet,
                                    clearable=False,
                                    className="dash-dropdown",
                                ),
                            ]
                        ),
                    ],
                ),

                html.Div("Columnas visibles", className="input-label"),

                dcc.Dropdown(
                    id="column-selector",
                    options=[
                        {
                            "label": column,
                            "value": column,
                        }
                        for column in default_columns
                    ],
                    value=default_columns,
                    multi=True,
                    className="dash-dropdown",
                ),

                html.Br(),

                html.Div(
                    id="selected-table-container",
                    children=[
                        html.H4(f"{default_file} · {default_sheet}"),
                        data_table(df_default, page_size=15, visible_columns=default_columns),
                    ],
                ),
            ],
        )

    if tab_value == "tab-resumen":
        decision = data["decision"]

        return html.Div(
            children=[
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(
                            className="card",
                            children=[
                                dcc.Graph(figure=figures["model"]),
                            ],
                        ),
                        html.Div(
                            className="card",
                            children=[
                                dcc.Graph(figure=figures["cover"]),
                            ],
                        ),
                    ],
                ),

                html.Div(
                    className="card",
                    children=[
                        html.H3("Resumen numérico"),
                        data_table(decision, page_size=5),
                    ],
                ),
            ],
        )

    if tab_value == "tab-plan":
        plan = data["plan"]

        preferred_cols = [
            "producto_cobertura",
            "fecha_inicio",
            "fecha_fin",
            "decision_operativa",
            "energia_producto_MWh",
            "energia_a_cubrir_MWh",
            "energia_abierta_MWh",
            "cobertura_recomendada_pct",
            "exposicion_abierta_pct",
            "precio_modelo_producto_EUR_MWh",
            "precio_omip_producto_EUR_MWh",
            "commodity_producto_recomendada_EUR_MWh",
            "importe_producto_recomendado_EUR",
            "señal",
        ]

        cols = [
            column
            for column in preferred_cols
            if column in plan.columns
        ]

        plan_view = plan[cols].copy() if cols else plan

        return html.Div(
            children=[
                html.Div(
                    className="card",
                    children=[
                        dcc.Graph(figure=figures["cover"]),
                    ],
                ),

                html.Div(
                    className="card",
                    children=[
                        html.H3("Plan operativo de cobertura"),
                        html.P(
                            "Decisión por producto temporal: mes, trimestre o CAL.",
                            className="small-note",
                        ),
                        data_table(plan_view, page_size=10),
                    ],
                ),
            ],
        )

    if tab_value == "tab-curva":
        curve = data["curve"]
        monthly_curve = data["monthly_curve"]
        hourly_profile = data["hourly_profile"]

        return html.Div(
            children=[
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(
                            className="card",
                            children=[
                                dcc.Graph(figure=figures["monthly"]),
                            ],
                        ),
                        html.Div(
                            className="card",
                            children=[
                                dcc.Graph(figure=figures["hourly"]),
                            ],
                        ),
                    ],
                ),

                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(
                            className="card",
                            children=[
                                html.H3("Resumen mensual"),
                                data_table(monthly_curve, page_size=12),
                            ],
                        ),
                        html.Div(
                            className="card",
                            children=[
                                html.H3("Perfil horario"),
                                data_table(hourly_profile, page_size=12),
                            ],
                        ),
                    ],
                ),

                html.Div(
                    className="card",
                    children=[
                        html.H3("Curva horaria"),
                        html.P(
                            "Se muestran las primeras 1.500 filas.",
                            className="small-note",
                        ),
                        data_table(curve.head(1500), page_size=15),
                    ],
                ),
            ],
        )

    if tab_value == "tab-prevision":
        forecast = data["forecast"]
        detail = data["detail"]

        forecast_cols = [
            column
            for column in [
                "fecha_hora_utc",
                "fecha",
                "precio_omie_previsto",
            ]
            if column in forecast.columns
        ]

        detail_cols = [
            column
            for column in [
                "fecha_hora_utc",
                "fecha",
                "precio_omie_previsto",
                "producto_omip",
                "precio_omip_EUR_MWh",
                "commodity_recomendada_hora_EUR_MWh",
                "energia_cliente_MWh",
            ]
            if column in detail.columns
        ]

        forecast_view = forecast[forecast_cols].head(1500) if forecast_cols else forecast.head(1500)
        detail_view = detail[detail_cols].head(1500) if detail_cols else detail.head(1500)

        return html.Div(
            children=[
                html.Div(
                    className="card",
                    children=[
                        dcc.Graph(figure=figures["model"]),
                    ],
                ),

                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(
                            className="card",
                            children=[
                                html.H3("Previsión horaria del modelo"),
                                data_table(forecast_view, page_size=15),
                            ],
                        ),
                        html.Div(
                            className="card",
                            children=[
                                html.H3("Detalle horario con cobertura"),
                                data_table(detail_view, page_size=15),
                            ],
                        ),
                    ],
                ),
            ],
        )

    if tab_value == "tab-omip":
        omip = data["omip"]
        products = data["products"]

        return html.Div(
            className="dashboard-grid-2",
            children=[
                html.Div(
                    className="card",
                    children=[
                        html.H3("OMIP manual"),
                        data_table(omip, page_size=14),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Productos usados en la oferta"),
                        data_table(products, page_size=14),
                    ],
                ),
            ],
        )

    if tab_value == "tab-archivos":
        files = get_excel_files()

        df_files = pd.DataFrame(
            [
                {
                    "archivo": path.name,
                    "existe": path.exists(),
                    "ruta": str(path),
                }
                for path in files.values()
            ]
        )

        return html.Div(
            className="card",
            children=[
                html.H3("Archivos generados y utilizados"),
                data_table(df_files, page_size=10),
            ],
        )

    return html.Div("Pestaña no reconocida.")


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app.run(debug=False, port=8050)