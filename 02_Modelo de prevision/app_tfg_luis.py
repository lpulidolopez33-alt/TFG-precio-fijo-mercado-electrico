
from __future__ import annotations

import base64
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
# RUTAS BASE
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = PROJECT_DIR / "config.py"

OUTPUT_DIR = config.OUTPUT_DIR
MANUAL_DATA_DIR = config.MANUAL_DATA_DIR
GRAPH_DIR = config.GRAPH_DIR

RUTA_PREVISION = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"
RUTA_DECISION = getattr(config, "RUTA_DECISION_COBERTURA", OUTPUT_DIR / "decision_commodity_cobertura.xlsx")
RUTA_COMPARACION_OMIP = getattr(config, "RUTA_COMPARACION_OMIP", OUTPUT_DIR / "comparacion_modelo_omip.xlsx")
RUTA_COMMODITY_PERIODOS = OUTPUT_DIR / "commodity_por_periodos.xlsx"
RUTA_BACKTESTING = getattr(config, "RUTA_BACKTESTING_TEMPORAL", OUTPUT_DIR / "backtesting_temporal_modelo.xlsx")
RUTA_HIBRIDA = getattr(config, "RUTA_PREVISION_HIBRIDA_OMIP", OUTPUT_DIR / "prevision_commodity_hibrida_omip.xlsx")
RUTA_RESUMEN_HIBRIDO = getattr(config, "RUTA_RESUMEN_HIBRIDO_OMIP", OUTPUT_DIR / "resumen_commodity_hibrida_omip.xlsx")

RUTA_OMIP_MANUAL = getattr(config, "RUTA_OMIP_MANUAL", MANUAL_DATA_DIR / "omip_manual.xlsx")
RUTA_CURVA_CLIENTE = getattr(config, "RUTA_CURVA_CLIENTE", MANUAL_DATA_DIR / "curva_cliente.xlsx")
RUTA_CURVAS_CLIENTES_SIMULADOS = getattr(config, "RUTA_CURVAS_CLIENTES_SIMULADOS", MANUAL_DATA_DIR / "curvas_clientes_simulados.xlsx")


# ============================================================
# APP
# ============================================================

app = Dash(
    __name__,
    suppress_callback_exceptions=True,
    title="TFG Commodity Hub",
)
server = app.server


# ============================================================
# ESTILO DE GRÁFICAS
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
    "pink": "#EC4899",
}


# ============================================================
# UTILIDADES
# ============================================================

def fmt(value, decimals: int = 2, dash: str = "—") -> str:
    if value is None or pd.isna(value):
        return dash
    try:
        text = f"{float(value):,.{decimals}f}"
        return text.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def pct(value, decimals: int = 1) -> str:
    if value is None or pd.isna(value):
        return "—"
    try:
        value = float(value)
        if abs(value) <= 1.5:
            value *= 100
        return fmt(value, decimals) + "%"
    except Exception:
        return "—"


def read_excel_safe(path: Path, sheet_name: str | None = None) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        if sheet_name is None:
            return pd.read_excel(path)
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


def excel_sheets(path: Path) -> list[str]:
    if not path.exists():
        return []
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception:
        return []


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].astype(str)
        elif pd.api.types.is_float_dtype(out[c]):
            out[c] = out[c].round(4)
    return out


def table(df: pd.DataFrame, page_size: int = 12, cols: list[str] | None = None) -> dash_table.DataTable:
    df = clean_df(df)
    if cols:
        cols = [c for c in cols if c in df.columns]
        if cols:
            df = df[cols].copy()

    if df.empty:
        return dash_table.DataTable(
            columns=[],
            data=[],
            style_table={"height": "220px", "overflowX": "auto"},
        )

    return dash_table.DataTable(
        columns=[{"name": str(c), "id": str(c), "hideable": True} for c in df.columns],
        data=df.to_dict("records"),
        page_size=page_size,
        sort_action="native",
        filter_action="native",
        column_selectable="multi",
        style_table={"overflowX": "auto", "borderRadius": "16px"},
        style_header={
            "backgroundColor": "#1F2A44",
            "color": "#F8FAFC",
            "fontWeight": "800",
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
            "fontFamily": "Segoe UI, Arial",
            "fontSize": "13px",
            "minWidth": "115px",
            "maxWidth": "360px",
            "whiteSpace": "normal",
            "textAlign": "left",
        },
        style_data_conditional=[
            {"if": {"row_index": "odd"}, "backgroundColor": "#152033"},
            {"if": {"state": "selected"}, "backgroundColor": "#243B63", "border": "1px solid #3B82F6"},
        ],
    )


def fig_base(title: str, height: int = 360) -> go.Figure:
    fig = go.Figure()
    return style_fig(fig, title, height)


def style_fig(fig: go.Figure, title: str, height: int = 360) -> go.Figure:
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,0.35)",
        title={"text": title, "x": 0.02, "font": {"size": 18}},
        font={"color": COLORS["text"]},
        height=height,
        margin={"l": 45, "r": 25, "t": 58, "b": 45},
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.08)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.08)")
    return fig


def kpi_card(title: str, value: str, unit: str = "", accent: str = "blue", note: str = ""):
    return html.Div(
        className=f"kpi-card kpi-{accent}",
        children=[
            html.Div(title, className="kpi-title"),
            html.Div(value, className="kpi-value"),
            html.Div(unit, className="kpi-unit"),
            html.Div(note, className="kpi-note") if note else None,
        ],
    )


def insight_card(icon: str, title: str, text: str, pill: str | None = None, accent: str = "blue"):
    return html.Div(
        className=f"insight-card insight-{accent}",
        children=[
            html.Div(
                className="insight-head",
                children=[
                    html.Div(icon, className="insight-icon"),
                    html.Div(
                        children=[
                            html.H4(title),
                            html.P(text),
                        ]
                    ),
                ],
            ),
            html.Span(pill, className=f"mini-pill mini-{accent}") if pill else None,
        ],
    )


def replace_or_add_config_var(text: str, var_name: str, literal: str) -> str:
    pattern = rf"^{var_name}\s*=.*$"
    new_line = f"{var_name} = {literal}"

    if re.search(pattern, text, flags=re.MULTILINE):
        return re.sub(pattern, new_line, text, flags=re.MULTILINE)

    return text.rstrip() + "\n\n" + new_line + "\n"


def apply_config(use_curve: bool, profile: str, energy_mwh: float) -> None:
    text = CONFIG_PATH.read_text(encoding="utf-8")

    text = replace_or_add_config_var(text, "USAR_CURVA_CLIENTE", "True" if use_curve else "False")
    text = replace_or_add_config_var(text, "PERFIL_REE_CLIENTE", repr(profile))
    text = replace_or_add_config_var(text, "ENERGIA_TOTAL_CLIENTE_MWH", str(float(energy_mwh)))

    CONFIG_PATH.write_text(text, encoding="utf-8")


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


def run_offer_workflow(start_date: str, end_date: str, energy_mwh: float, customer_mode: str, profile: str) -> tuple[bool, str]:
    logs = []

    if customer_mode == "Perfil plano":
        apply_config(use_curve=False, profile=profile, energy_mwh=energy_mwh)
        ok, log = run_main_option("4", start_date, end_date)
        return ok, "OPCIÓN 4 - PERFIL PLANO\n\n" + log

    apply_config(use_curve=False, profile=profile, energy_mwh=energy_mwh)

    ok_base, log_base = run_main_option("4", start_date, end_date)
    logs.append("PASO 1 - PREVISIÓN BASE SIN CURVA")
    logs.append(log_base)

    if not ok_base:
        return False, "\n\n".join(logs)

    ok_curve, log_curve = run_main_option("6")
    logs.append("PASO 2 - GENERACIÓN DE CURVA HORARIA")
    logs.append(log_curve)

    if not ok_curve:
        return False, "\n\n".join(logs)

    apply_config(use_curve=True, profile=profile, energy_mwh=energy_mwh)

    ok_final, log_final = run_main_option("4", start_date, end_date)
    logs.append("PASO 3 - CÁLCULO FINAL CON CURVA HORARIA")
    logs.append(log_final)

    return ok_final, "\n\n".join(logs)


# ============================================================
# CARGA DE RESULTADOS
# ============================================================

def load_results() -> dict[str, pd.DataFrame]:
    return {
        "decision": read_excel_safe(RUTA_DECISION, "resumen_decision"),
        "plan": read_excel_safe(RUTA_DECISION, "plan_cobertura"),
        "products": read_excel_safe(RUTA_DECISION, "productos_omip_usados"),
        "detail": read_excel_safe(RUTA_DECISION, "detalle_horario"),
        "forecast_global": read_excel_safe(RUTA_PREVISION, "resumen_global"),
        "forecast_monthly": read_excel_safe(RUTA_PREVISION, "resumen_mensual"),
        "forecast_daily": read_excel_safe(RUTA_PREVISION, "resumen_diario"),
        "forecast_hourly": read_excel_safe(RUTA_PREVISION, "prevision_horaria"),
        "period_table": read_excel_safe(RUTA_COMMODITY_PERIODOS, "tabla_calculadora"),
        "period_summary": read_excel_safe(RUTA_COMMODITY_PERIODOS, "resumen_periodos"),
        "period_hourly": read_excel_safe(RUTA_COMMODITY_PERIODOS, "prevision_horaria_periodos"),
        "comparison": read_excel_safe(RUTA_COMPARACION_OMIP, "comparacion_modelo_omip"),
        "omip": read_excel_safe(RUTA_OMIP_MANUAL),
        "curve": read_excel_safe(RUTA_CURVA_CLIENTE, "curva_horaria"),
        "curve_monthly": read_excel_safe(RUTA_CURVA_CLIENTE, "resumen_mensual"),
        "curve_hourly": read_excel_safe(RUTA_CURVA_CLIENTE, "perfil_horario"),
        "curve_weekday": read_excel_safe(RUTA_CURVA_CLIENTE, "perfil_dia_semana"),
        "backtest_global": read_excel_safe(RUTA_BACKTESTING, "resumen_global"),
        "backtest_monthly": read_excel_safe(RUTA_BACKTESTING, "resumen_mensual"),
        "backtest_hourly": read_excel_safe(RUTA_BACKTESTING, "validacion_horaria"),
        "backtest_vars": read_excel_safe(RUTA_BACKTESTING, "variables_usadas"),
        "hybrid_global": read_excel_safe(RUTA_RESUMEN_HIBRIDO, "resumen_global"),
        "hybrid_products": read_excel_safe(RUTA_RESUMEN_HIBRIDO, "resumen_productos"),
    }


def available_files() -> dict[str, Path]:
    return {
        "Decisión commodity y cobertura": RUTA_DECISION,
        "Previsión commodity modelo": RUTA_PREVISION,
        "Commodity por periodos": RUTA_COMMODITY_PERIODOS,
        "Comparación modelo vs OMIP": RUTA_COMPARACION_OMIP,
        "Curva cliente": RUTA_CURVA_CLIENTE,
        "Curvas clientes simulados": RUTA_CURVAS_CLIENTES_SIMULADOS,
        "OMIP manual": RUTA_OMIP_MANUAL,
        "Backtesting temporal": RUTA_BACKTESTING,
        "Previsión híbrida OMIP": RUTA_HIBRIDA,
        "Resumen híbrido OMIP": RUTA_RESUMEN_HIBRIDO,
    }


# ============================================================
# GRÁFICAS
# ============================================================

def build_price_figure(data: dict, chart_type: str = "line") -> go.Figure:
    forecast = data["forecast_hourly"]
    detail = data["detail"]
    fig = fig_base("Precio previsto modelo vs OMIP")

    if forecast.empty or "precio_omie_previsto" not in forecast.columns:
        return fig

    df = forecast.copy()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df = df.dropna(subset=["fecha"])

    daily = df.groupby("fecha", as_index=False).agg(precio_modelo=("precio_omie_previsto", "mean"))

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=daily["fecha"],
            y=daily["precio_modelo"],
            mode="lines",
            fill="tozeroy" if chart_type == "area" else None,
            name="Modelo",
            line={"color": COLORS["cyan"], "width": 3},
        )
    )

    if not detail.empty and {"fecha", "precio_omip_EUR_MWh"}.issubset(detail.columns):
        d = detail.copy()
        d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce")
        omip = (
            d.dropna(subset=["fecha", "precio_omip_EUR_MWh"])
            .groupby("fecha", as_index=False)
            .agg(precio_omip=("precio_omip_EUR_MWh", "mean"))
        )
        if not omip.empty:
            fig.add_trace(
                go.Scatter(
                    x=omip["fecha"],
                    y=omip["precio_omip"],
                    mode="lines",
                    name="OMIP asignado",
                    line={"color": COLORS["orange"], "width": 2, "dash": "dash"},
                )
            )

    return style_fig(fig, "Precio diario previsto: modelo vs OMIP", 390)


def build_cover_figure(data: dict, chart_type: str = "bar_stack") -> go.Figure:
    plan = data["plan"]
    fig = fig_base("Plan de cobertura")

    if plan.empty:
        return fig

    product_col = "producto_cobertura" if "producto_cobertura" in plan.columns else plan.columns[0]

    if chart_type == "donut":
        covered = plan["energia_a_cubrir_MWh"].sum() if "energia_a_cubrir_MWh" in plan.columns else 0
        open_energy = plan["energia_abierta_MWh"].sum() if "energia_abierta_MWh" in plan.columns else 0

        fig = go.Figure(
            go.Pie(
                labels=["Cubierta", "Abierta"],
                values=[covered, open_energy],
                hole=0.63,
                marker={"colors": [COLORS["green"], COLORS["orange"]]},
            )
        )
        return style_fig(fig, "Distribución de cobertura", 390)

    fig = go.Figure()

    if "energia_a_cubrir_MWh" in plan.columns:
        fig.add_trace(
            go.Bar(
                x=plan[product_col].astype(str),
                y=plan["energia_a_cubrir_MWh"],
                name="Energía cubierta",
                marker_color=COLORS["green"],
            )
        )

    if "energia_abierta_MWh" in plan.columns:
        fig.add_trace(
            go.Bar(
                x=plan[product_col].astype(str),
                y=plan["energia_abierta_MWh"],
                name="Energía abierta",
                marker_color=COLORS["orange"],
            )
        )

    fig.update_layout(barmode="group" if chart_type == "bar_group" else "stack")
    return style_fig(fig, "Energía cubierta y abierta por producto", 390)


def build_monthly_curve_figure(data: dict, chart_type: str = "bar") -> go.Figure:
    df = data["curve_monthly"]
    fig = fig_base("Consumo mensual del cliente")

    if df.empty or not {"año_mes", "energia_MWh"}.issubset(df.columns):
        return fig

    fig = go.Figure()

    if chart_type == "line":
        fig.add_trace(go.Scatter(x=df["año_mes"], y=df["energia_MWh"], mode="lines+markers", name="Energía", line={"color": COLORS["purple"], "width": 3}))
    elif chart_type == "area":
        fig.add_trace(go.Scatter(x=df["año_mes"], y=df["energia_MWh"], mode="lines", fill="tozeroy", name="Energía", line={"color": COLORS["purple"], "width": 3}))
    elif chart_type == "donut":
        fig.add_trace(go.Pie(labels=df["año_mes"], values=df["energia_MWh"], hole=0.55))
    else:
        fig.add_trace(go.Bar(x=df["año_mes"], y=df["energia_MWh"], name="Energía", marker_color=COLORS["purple"]))

    return style_fig(fig, "Consumo mensual del cliente", 390)


def build_hourly_profile_figure(data: dict) -> go.Figure:
    df = data["curve_hourly"]
    fig = fig_base("Perfil horario medio")

    if df.empty or not {"hora", "consumo_medio_MWh"}.issubset(df.columns):
        return fig

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df["hora"],
            y=df["consumo_medio_MWh"],
            mode="lines+markers",
            name="Consumo medio",
            line={"color": COLORS["green"], "width": 3},
        )
    )
    fig.update_xaxes(dtick=1)
    return style_fig(fig, "Perfil horario medio del cliente", 360)


def build_comparison_figure(data: dict) -> go.Figure:
    df = data["comparison"]
    fig = fig_base("Comparación modelo vs OMIP por producto")

    if df.empty or not {"producto", "precio_modelo_EUR_MWh", "precio_omip_EUR_MWh"}.issubset(df.columns):
        return fig

    fig = go.Figure()
    fig.add_trace(go.Bar(x=df["producto"].astype(str), y=df["precio_modelo_EUR_MWh"], name="Modelo", marker_color=COLORS["cyan"]))
    fig.add_trace(go.Bar(x=df["producto"].astype(str), y=df["precio_omip_EUR_MWh"], name="OMIP", marker_color=COLORS["orange"]))
    fig.update_layout(barmode="group")
    return style_fig(fig, "Modelo vs OMIP por producto", 390)


def build_backtest_figure(data: dict) -> go.Figure:
    df = data["backtest_hourly"]
    fig = fig_base("Backtesting: precio real vs previsto")

    if df.empty or not {"fecha", "precio_real_omie", "precio_previsto_backtesting"}.issubset(df.columns):
        return fig

    d = df.copy()
    d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce")
    daily = (
        d.dropna(subset=["fecha"])
        .groupby("fecha", as_index=False)
        .agg(
            real=("precio_real_omie", "mean"),
            previsto=("precio_previsto_backtesting", "mean"),
        )
    )

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=daily["fecha"], y=daily["real"], name="Real OMIE", mode="lines", line={"color": COLORS["green"], "width": 3}))
    fig.add_trace(go.Scatter(x=daily["fecha"], y=daily["previsto"], name="Previsto", mode="lines", line={"color": COLORS["cyan"], "width": 3}))
    return style_fig(fig, "Backtesting diario: real vs previsto", 390)


def build_period_heatmap(data: dict) -> go.Figure:
    df = data["period_summary"]
    fig = fig_base("Commodity por tarifa y periodo")

    if df.empty or not {"tarifa", "periodo", "commodity_modelo_EUR_MWh"}.issubset(df.columns):
        return fig

    pivot = df.pivot_table(index="tarifa", columns="periodo", values="commodity_modelo_EUR_MWh", aggfunc="mean")
    fig = go.Figure(
        data=go.Heatmap(
            z=pivot.values,
            x=pivot.columns.astype(str),
            y=pivot.index.astype(str),
            colorscale="Blues",
            colorbar={"title": "€/MWh"},
            text=[[fmt(v, 2) for v in row] for row in pivot.values],
            texttemplate="%{text}",
        )
    )
    return style_fig(fig, "Mapa de calor: commodity por tarifa y periodo", 360)


# ============================================================
# KPIS E INSIGHTS
# ============================================================

def build_kpis(decision: pd.DataFrame) -> list:
    if decision.empty:
        return [
            kpi_card("Commodity modelo", "—", "€/MWh", "blue"),
            kpi_card("OMIP equivalente", "—", "€/MWh", "cyan"),
            kpi_card("Commodity recomendada", "—", "€/MWh", "green"),
            kpi_card("Diferencia modelo-OMIP", "—", "€/MWh", "orange"),
            kpi_card("Energía cliente", "—", "MWh", "purple"),
            kpi_card("Cobertura", "—", "%", "green"),
            kpi_card("Abierto", "—", "%", "orange"),
            kpi_card("Importe recomendado", "—", "€", "purple"),
        ]

    row = decision.iloc[0]
    saving = None
    if "importe_modelo_total_EUR" in row and "importe_recomendado_total_EUR" in row:
        saving = row.get("importe_modelo_total_EUR") - row.get("importe_recomendado_total_EUR")

    return [
        kpi_card("Commodity modelo", fmt(row.get("commodity_modelo_total_EUR_MWh"), 2), "€/MWh", "blue"),
        kpi_card("OMIP equivalente", fmt(row.get("omip_equivalente_EUR_MWh"), 2), "€/MWh", "cyan"),
        kpi_card("Commodity recomendada", fmt(row.get("commodity_recomendada_EUR_MWh"), 2), "€/MWh", "green"),
        kpi_card("Diferencia modelo-OMIP", fmt(row.get("diferencia_modelo_menos_omip_EUR_MWh"), 2), "€/MWh", "orange"),
        kpi_card("Energía cliente", fmt(row.get("energia_total_cliente_MWh"), 0), "MWh", "purple"),
        kpi_card("Cobertura", pct(row.get("porcentaje_cobertura_total"), 1), "energía cubierta", "green"),
        kpi_card("Abierto", pct(row.get("porcentaje_abierto_total"), 1), "exposición abierta", "orange"),
        kpi_card("Ahorro vs modelo", fmt(saving, 0), "€", "purple", "modelo - recomendada" if saving is not None else ""),
    ]


def build_executive_insights(data: dict) -> list:
    decision = data["decision"]
    comparison = data["comparison"]
    backtest = data["backtest_global"]

    cards = []

    if not decision.empty:
        row = decision.iloc[0]
        diff = row.get("diferencia_modelo_menos_omip_EUR_MWh")
        coverage = row.get("porcentaje_cobertura_total")
        recommended = row.get("commodity_recomendada_EUR_MWh")

        if pd.notna(diff) and diff > 0:
            cards.append(insight_card("🟢", "Señal de cobertura favorable", f"El modelo queda {fmt(diff, 2)} €/MWh por encima de la referencia OMIP, por lo que la app recomienda cubrir el tramo disponible.", "Cubrir", "green"))
        elif pd.notna(diff):
            cards.append(insight_card("🟠", "Señal de mercado abierta", f"El modelo queda {fmt(abs(diff), 2)} €/MWh por debajo de OMIP; conviene revisar si interesa dejar parte abierta.", "Revisar", "orange"))
        else:
            cards.append(insight_card("🔎", "Sin diferencia calculada", "No se ha podido leer la diferencia entre modelo y OMIP en el resumen de decisión.", "Sin dato", "blue"))

        cards.append(insight_card("🛡️", "Cobertura recomendada", f"La cobertura total calculada es {pct(coverage, 1)} y la commodity recomendada se sitúa en {fmt(recommended, 2)} €/MWh.", "Plan operativo", "blue"))

    else:
        cards.append(insight_card("⚠️", "No hay resultados cargados", "Ejecuta la previsión desde el panel lateral o revisa que existan los Excel de salida.", "Pendiente", "orange"))

    if not comparison.empty and "estado" in comparison.columns:
        complete = (comparison["estado"].astype(str).str.contains("completa", case=False, na=False)).sum()
        total = len(comparison)
        cards.append(insight_card("📈", "Cobertura OMIP disponible", f"{complete} de {total} productos tienen comparación completa entre modelo y OMIP.", "OMIP", "cyan"))

    if not backtest.empty:
        row = backtest.iloc[0]
        cards.append(insight_card("🧪", "Validación del modelo", f"MAE: {fmt(row.get('MAE_EUR_MWh'), 2)} €/MWh · RMSE: {fmt(row.get('RMSE_EUR_MWh'), 2)} €/MWh · R²: {fmt(row.get('R2'), 3)}.", "Backtesting", "purple"))

    return cards


# ============================================================
# UPLOADS
# ============================================================

def save_uploaded_file(contents: str, filename: str, target_path: Path) -> str:
    if not contents:
        return "No se ha recibido ningún archivo."

    try:
        content_type, content_string = contents.split(",", 1)
        decoded = base64.b64decode(content_string)
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(decoded)
        return f"Archivo guardado como {target_path.name}."
    except Exception as error:
        return f"No se ha podido guardar {filename}: {error}"


# ============================================================
# LAYOUT
# ============================================================

app.layout = html.Div(
    className="app-shell",
    children=[
        dcc.Store(id="refresh-token", data=0),
        dcc.Store(id="log-store", data=""),
        dcc.Store(id="status-store", data="Listo para ejecutar"),
        dcc.Download(id="download-output"),

        html.Div(
            className="sidebar",
            children=[
                html.Div(
                    className="logo-row",
                    children=[
                        html.Div("⚡", className="logo"),
                        html.Div(
                            children=[
                                html.Div("TFG Commodity", className="logo-title"),
                                html.Div("Pricing · OMIP · Cobertura", className="logo-subtitle"),
                            ]
                        ),
                    ],
                ),

                html.Div("Oferta", className="section-title"),

                html.Div("Periodo de oferta", className="input-label"),
                dcc.DatePickerRange(
                    id="date-range",
                    start_date=date(2026, 6, 1),
                    end_date=date(2027, 5, 31),
                    display_format="DD/MM/YYYY",
                    minimum_nights=0,
                    with_portal=True,
                    day_size=39,
                    className="date-picker",
                ),

                html.Div("Consumo total cliente [MWh]", className="input-label"),
                dcc.Input(id="energy-input", type="number", value=12000, min=1, step=100, className="input-box"),

                html.Div("Tipo de cliente", className="input-label"),
                dcc.RadioItems(
                    id="customer-mode",
                    options=[
                        {"label": "Perfil plano", "value": "Perfil plano"},
                        {"label": "Perfil no plano / curva cliente", "value": "Perfil no plano"},
                    ],
                    value="Perfil no plano",
                    className="radio-box",
                    labelStyle={"display": "block", "marginTop": "8px"},
                ),

                html.Div("Tarifa / perfil de curva", className="input-label"),
                dcc.Dropdown(
                    id="profile-select",
                    options=[
                        {"label": "2.0TD · Doméstico", "value": "2.0TD"},
                        {"label": "3.0TD · Pyme/comercial", "value": "3.0TD"},
                        {"label": "3.0TDVE · Vehículo eléctrico", "value": "3.0TDVE"},
                        {"label": "6.XTD · Industrial", "value": "6.XTD"},
                    ],
                    value="3.0TD",
                    clearable=False,
                    className="dash-dropdown",
                ),

                html.Div("Ejecución", className="section-title"),
                dcc.Loading(
                    type="circle",
                    color="#22D3EE",
                    children=[
                        html.Div(id="execution-feedback", className="execution-feedback", children="Listo para ejecutar el cálculo."),
                    ],
                ),

                html.Button("Calcular oferta completa", id="btn-calc-offer", n_clicks=0, className="main-button"),

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

                html.Button("Actualizar vista", id="btn-refresh", n_clicks=0, className="secondary-button"),

                html.Div("Visualización", className="section-title"),
                html.Div("Cobertura", className="input-label"),
                dcc.Dropdown(
                    id="cover-chart-type",
                    options=[
                        {"label": "Barras apiladas", "value": "bar_stack"},
                        {"label": "Barras agrupadas", "value": "bar_group"},
                        {"label": "Dónut cubierta / abierta", "value": "donut"},
                    ],
                    value="bar_stack",
                    clearable=False,
                    className="dash-dropdown",
                ),

                html.Div("Consumo", className="input-label"),
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

                html.Div("Precio", className="input-label"),
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

                html.Div("Notas", className="section-title"),
                html.Div(
                    className="small-note",
                    children=[
                        html.Div("• Esta app ejecuta el mismo main.py que usabas en terminal."),
                        html.Div("• Cierra los Excel de salida antes de calcular."),
                        html.Div("• La opción de perfil no plano genera curva y recalcula."),
                        html.Div("• Las tablas permiten filtrar, ordenar y ocultar columnas."),
                    ],
                ),
            ],
        ),

        html.Div(
            className="main",
            children=[
                html.Div(
                    className="hero-premium",
                    children=[
                        html.Div(
                            children=[
                                html.Div("Calculadora local para tu TFG", className="eyebrow"),
                                html.H1("Commodity Hub"),
                                html.P("Interfaz visual para ejecutar el modelo Python, consultar la commodity recomendada, comparar con OMIP, revisar cobertura, curva del cliente, backtesting y archivos generados."),
                            ]
                        ),
                        html.Div(id="status-pill", className="status-pill"),
                    ],
                ),

                dcc.Loading(
                    type="circle",
                    color="#22D3EE",
                    children=[html.Div(id="status-box", className="status-box")],
                ),

                html.Div(id="kpi-grid", className="kpi-grid"),

                dcc.Tabs(
                    id="tabs",
                    value="tab-general",
                    className="tabs-container",
                    children=[
                        dcc.Tab(label="Vista general", value="tab-general", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Ejecución", value="tab-ejecucion", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Commodity", value="tab-commodity", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Cobertura y OMIP", value="tab-cobertura", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Curva cliente", value="tab-curva", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Modelo y riesgo", value="tab-modelo", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Datos de entrada", value="tab-entrada", className="tab", selected_className="tab-selected"),
                        dcc.Tab(label="Archivos y tablas", value="tab-archivos", className="tab", selected_className="tab-selected"),
                    ],
                ),

                dcc.Loading(
                    id="loading-tabs",
                    type="circle",
                    color="#22D3EE",
                    children=[html.Div(id="tab-content", className="tab-content")],
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
    Input("btn-calc-offer", "n_clicks"),
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
    n_calc,
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
    token = (token or 0) + 1

    if trigger == "btn-refresh":
        message = "Vista actualizada sin ejecutar procesos."
        return "Vista actualizada", message, token, message

    if not start_date or not end_date:
        message = "Faltan fechas de oferta."
        return "Faltan fechas", message, token, message

    if not energy_mwh or float(energy_mwh) <= 0:
        message = "El consumo total debe ser mayor que cero."
        return "Consumo no válido", message, token, message

    try:
        energy_mwh = float(energy_mwh)
    except Exception:
        message = "El consumo total no es válido."
        return "Consumo no válido", message, token, message

    if trigger == "btn-calc-offer":
        ok, log = run_offer_workflow(start_date, end_date, energy_mwh, customer_mode, profile)
        status = "Oferta calculada correctamente" if ok else "Oferta con errores"
        return status, log, token, status

    option_map = {
        "btn-1": "1",
        "btn-2": "2",
        "btn-3": "3",
        "btn-4": "4",
        "btn-5": "5",
        "btn-6": "6",
    }

    option = option_map.get(trigger)
    if not option:
        return no_update, no_update, no_update, no_update

    if option == "4":
        ok, log = run_offer_workflow(start_date, end_date, energy_mwh, customer_mode, profile)
    elif option == "6":
        apply_config(use_curve=False, profile=profile, energy_mwh=energy_mwh)
        ok_base, log_base = run_main_option("4", start_date, end_date)
        if not ok_base:
            ok, log = False, "PASO 1 - PREVISIÓN BASE\n\n" + log_base
        else:
            ok_curve, log_curve = run_main_option("6")
            ok = ok_curve
            log = "PASO 1 - PREVISIÓN BASE\n\n" + log_base + "\n\nPASO 2 - GENERACIÓN DE CURVA\n\n" + log_curve
    else:
        apply_config(use_curve=(customer_mode == "Perfil no plano"), profile=profile, energy_mwh=energy_mwh)
        ok, log = run_main_option(option, start_date, end_date)

    status = f"Opción {option} completada" if ok else f"Opción {option} con errores"
    return status, log, token, status


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
        pill = "Sin resultados"
    else:
        mode = str(decision.iloc[0].get("modo_ponderacion", "Resultados cargados"))
        pill = mode if len(mode) < 48 else mode[:45] + "..."

    status_box = html.Div(
        className="status-message",
        children=[
            html.Div("Estado actual", className="status-title"),
            html.Div(status_text or "Listo para ejecutar", className="status-text"),
        ],
    )

    return kpis, pill, status_box


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
def render_tab(tab_value, _, cover_chart_type, consumption_chart_type, price_chart_type, log_text):
    data = load_results()

    fig_price = build_price_figure(data, price_chart_type)
    fig_cover = build_cover_figure(data, cover_chart_type)
    fig_monthly = build_monthly_curve_figure(data, consumption_chart_type)
    fig_hourly = build_hourly_profile_figure(data)
    fig_comparison = build_comparison_figure(data)
    fig_backtest = build_backtest_figure(data)
    fig_periods = build_period_heatmap(data)

    if tab_value == "tab-general":
        return html.Div(
            children=[
                html.Div(className="insights-grid", children=build_executive_insights(data)),
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(className="card", children=[dcc.Graph(figure=fig_price)]),
                        html.Div(className="card", children=[dcc.Graph(figure=fig_cover)]),
                    ],
                ),
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(className="card", children=[dcc.Graph(figure=fig_periods)]),
                        html.Div(
                            className="card",
                            children=[
                                html.H3("Tabla para calculadora por periodos"),
                                html.P("Vista limpia de los valores P1-P6 que pueden trasladarse a la calculadora.", className="small-note"),
                                table(data["period_table"], page_size=8),
                            ],
                        ),
                    ],
                ),
            ]
        )

    if tab_value == "tab-ejecucion":
        return html.Div(
            children=[
                html.Div(
                    className="workflow-grid",
                    children=[
                        insight_card("1", "Actualizar datos", "Descarga/procesa histórico, valida OMIE, demanda, generación, gas y CO₂.", "Opción 2", "blue"),
                        insight_card("2", "Entrenar modelo", "Actualiza el modelo predictivo y el análisis de riesgo.", "Opción 3", "purple"),
                        insight_card("3", "Previsión oferta", "Calcula la previsión futura y commodity por periodos.", "Opción 4", "green"),
                        insight_card("4", "Cobertura", "Compara con OMIP y genera la decisión operativa.", "Automático", "orange"),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Registro de ejecución"),
                        html.P("Aquí aparece lo mismo que antes veías en el terminal.", className="small-note"),
                        html.Pre(log_text or "Todavía no se ha ejecutado ningún proceso desde la app.", className="log-box"),
                    ],
                ),
            ]
        )

    if tab_value == "tab-commodity":
        return html.Div(
            children=[
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(className="card", children=[dcc.Graph(figure=fig_price)]),
                        html.Div(className="card", children=[dcc.Graph(figure=fig_periods)]),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Resumen mensual de previsión"),
                        table(data["forecast_monthly"], page_size=14),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Commodity por periodos"),
                        table(data["period_summary"], page_size=18),
                    ],
                ),
            ]
        )

    if tab_value == "tab-cobertura":
        preferred = [
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
            "diferencia_modelo_menos_omip_EUR_MWh",
            "commodity_producto_recomendada_EUR_MWh",
            "importe_producto_recomendado_EUR",
            "señal",
        ]

        return html.Div(
            children=[
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(className="card", children=[dcc.Graph(figure=fig_cover)]),
                        html.Div(className="card", children=[dcc.Graph(figure=fig_comparison)]),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Plan operativo de cobertura"),
                        html.P("Tabla resumida, ordenable y filtrable. Puedes ocultar columnas desde la cabecera.", className="small-note"),
                        table(data["plan"], page_size=14, cols=preferred),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Productos OMIP usados"),
                        table(data["products"], page_size=14),
                    ],
                ),
            ]
        )

    if tab_value == "tab-curva":
        return html.Div(
            children=[
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(className="card", children=[dcc.Graph(figure=fig_monthly)]),
                        html.Div(className="card", children=[dcc.Graph(figure=fig_hourly)]),
                    ],
                ),
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(className="card", children=[html.H3("Resumen mensual curva cliente"), table(data["curve_monthly"], page_size=14)]),
                        html.Div(className="card", children=[html.H3("Perfil por día de la semana"), table(data["curve_weekday"], page_size=7)]),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Curva horaria del cliente"),
                        html.P("Muestra la curva generada o importada. Se pagina para no colapsar la pantalla.", className="small-note"),
                        table(data["curve"], page_size=18),
                    ],
                ),
            ]
        )

    if tab_value == "tab-modelo":
        bg = data["backtest_global"]
        bt_cards = []
        if not bg.empty:
            row = bg.iloc[0]
            bt_cards = [
                kpi_card("MAE", fmt(row.get("MAE_EUR_MWh"), 2), "€/MWh", "orange"),
                kpi_card("RMSE", fmt(row.get("RMSE_EUR_MWh"), 2), "€/MWh", "red"),
                kpi_card("R²", fmt(row.get("R2"), 3), "bondad de ajuste", "green"),
                kpi_card("Error medio", fmt(row.get("error_medio_EUR_MWh"), 2), "€/MWh", "blue"),
            ]

        graph_cards = []
        graph_names = [
            ("modelo_precio_real_vs_estimado.png", "Real vs estimado"),
            ("modelo_distribucion_errores.png", "Distribución de errores"),
            ("modelo_importancia_variables.png", "Importancia de variables"),
            ("modelo_prima_incertidumbre.png", "Prima de incertidumbre"),
        ]

        for file_name, title in graph_names:
            path = GRAPH_DIR / file_name
            if path.exists():
                graph_cards.append(html.Div(className="image-card", children=[html.H4(title), html.Img(src=f"/assets_generated/{file_name}", className="model-image")]))

        return html.Div(
            children=[
                html.Div(className="mini-kpi-grid", children=bt_cards),
                html.Div(className="card", children=[dcc.Graph(figure=fig_backtest)]),
                html.Div(className="card", children=[html.H3("Backtesting mensual"), table(data["backtest_monthly"], page_size=12)]),
                html.Div(className="image-grid", children=graph_cards) if graph_cards else html.Div(className="card", children=[html.H3("Gráficas del modelo"), html.P("No se han encontrado imágenes en la carpeta graficas.", className="small-note")]),
            ]
        )

    if tab_value == "tab-entrada":
        return html.Div(
            children=[
                html.Div(
                    className="dashboard-grid-2",
                    children=[
                        html.Div(
                            className="card",
                            children=[
                                html.H3("Subir OMIP manual"),
                                html.P("Permite sustituir el Excel de OMIP manual sin entrar a la carpeta.", className="small-note"),
                                dcc.Upload(
                                    id="upload-omip",
                                    children=html.Div(["Arrastra o selecciona ", html.B("omip_manual.xlsx")]),
                                    className="upload-box",
                                    multiple=False,
                                ),
                                html.Div(id="upload-omip-status", className="upload-status"),
                            ],
                        ),
                        html.Div(
                            className="card",
                            children=[
                                html.H3("Subir curva cliente"),
                                html.P("Permite sustituir la curva cliente si quieres trabajar con una curva propia.", className="small-note"),
                                dcc.Upload(
                                    id="upload-curve",
                                    children=html.Div(["Arrastra o selecciona ", html.B("curva_cliente.xlsx")]),
                                    className="upload-box",
                                    multiple=False,
                                ),
                                html.Div(id="upload-curve-status", className="upload-status"),
                            ],
                        ),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("OMIP manual cargado"),
                        table(data["omip"], page_size=10),
                    ],
                ),
                html.Div(
                    className="card",
                    children=[
                        html.H3("Rutas usadas por la app"),
                        html.Div(
                            className="path-grid",
                            children=[
                                html.Div([html.B("OMIP manual"), html.P(str(RUTA_OMIP_MANUAL), className="small-note")]),
                                html.Div([html.B("Curva cliente"), html.P(str(RUTA_CURVA_CLIENTE), className="small-note")]),
                                html.Div([html.B("Salidas"), html.P(str(OUTPUT_DIR), className="small-note")]),
                                html.Div([html.B("Gráficas"), html.P(str(GRAPH_DIR), className="small-note")]),
                            ],
                        ),
                    ],
                ),
            ]
        )

    if tab_value == "tab-archivos":
        files = available_files()
        existing = {k: v for k, v in files.items() if v.exists()}
        default_file = next(iter(existing.keys()), None)
        default_path = existing.get(default_file) if default_file else None
        sheets = excel_sheets(default_path) if default_path else []
        default_sheet = sheets[0] if sheets else None
        df_default = read_excel_safe(default_path, default_sheet) if default_path and default_sheet else pd.DataFrame()

        return html.Div(
            children=[
                html.Div(
                    className="card",
                    children=[
                        html.H3("Explorador de Excel generados"),
                        html.P("Selecciona archivo, hoja y columnas visibles para revisar todo sin abrir Excel.", className="small-note"),
                        html.Div(
                            className="dashboard-grid-3",
                            children=[
                                html.Div(
                                    children=[
                                        html.Div("Archivo", className="input-label"),
                                        dcc.Dropdown(
                                            id="file-selector",
                                            options=[{"label": k, "value": k} for k in existing.keys()],
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
                                            options=[{"label": s, "value": s} for s in sheets],
                                            value=default_sheet,
                                            clearable=False,
                                            className="dash-dropdown",
                                        ),
                                    ]
                                ),
                                html.Div(
                                    children=[
                                        html.Div("Descarga", className="input-label"),
                                        html.Button("Descargar Excel seleccionado", id="btn-download-file", n_clicks=0, className="secondary-button compact-button"),
                                    ]
                                ),
                            ],
                        ),
                        html.Div("Columnas visibles", className="input-label"),
                        dcc.Dropdown(
                            id="column-selector",
                            options=[{"label": c, "value": c} for c in df_default.columns],
                            value=list(df_default.columns),
                            multi=True,
                            className="dash-dropdown",
                        ),
                        html.Br(),
                        html.Div(
                            id="selected-table-container",
                            children=[
                                html.H4(f"{default_file or 'Sin archivo'} · {default_sheet or 'Sin hoja'}"),
                                table(df_default, page_size=15, cols=list(df_default.columns)),
                            ],
                        ),
                    ],
                ),
            ]
        )

    return html.Div(className="card", children=[html.H3("Pestaña no disponible")])


# ============================================================
# CALLBACKS TABLAS / ARCHIVOS
# ============================================================

@app.callback(
    Output("sheet-selector", "options"),
    Output("sheet-selector", "value"),
    Input("file-selector", "value"),
    prevent_initial_call=True,
)
def update_sheet_selector(file_label):
    files = available_files()
    path = files.get(file_label)
    sheets = excel_sheets(path) if path else []
    return [{"label": s, "value": s} for s in sheets], sheets[0] if sheets else None


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

    path = available_files().get(file_label)
    df = read_excel_safe(path, sheet_name) if path else pd.DataFrame()
    options = [{"label": c, "value": c} for c in df.columns]
    return options, list(df.columns)


@app.callback(
    Output("selected-table-container", "children"),
    Input("file-selector", "value"),
    Input("sheet-selector", "value"),
    Input("column-selector", "value"),
    prevent_initial_call=True,
)
def update_selected_table(file_label, sheet_name, columns):
    if not file_label or not sheet_name:
        return html.P("Selecciona archivo y hoja.", className="small-note")

    path = available_files().get(file_label)
    df = read_excel_safe(path, sheet_name) if path else pd.DataFrame()

    return html.Div(
        children=[
            html.H4(f"{file_label} · {sheet_name}"),
            table(df, page_size=15, cols=columns),
        ]
    )


@app.callback(
    Output("download-output", "data"),
    Input("btn-download-file", "n_clicks"),
    State("file-selector", "value"),
    prevent_initial_call=True,
)
def download_selected_file(n_clicks, file_label):
    path = available_files().get(file_label)
    if path and path.exists():
        return dcc.send_file(str(path))
    return no_update


# ============================================================
# CALLBACKS UPLOADS
# ============================================================

@app.callback(
    Output("upload-omip-status", "children"),
    Output("refresh-token", "data", allow_duplicate=True),
    Input("upload-omip", "contents"),
    State("upload-omip", "filename"),
    State("refresh-token", "data"),
    prevent_initial_call=True,
)
def upload_omip(contents, filename, token):
    if not contents:
        return no_update, no_update
    message = save_uploaded_file(contents, filename or "omip_manual.xlsx", RUTA_OMIP_MANUAL)
    return message, (token or 0) + 1


@app.callback(
    Output("upload-curve-status", "children"),
    Output("refresh-token", "data", allow_duplicate=True),
    Input("upload-curve", "contents"),
    State("upload-curve", "filename"),
    State("refresh-token", "data"),
    prevent_initial_call=True,
)
def upload_curve(contents, filename, token):
    if not contents:
        return no_update, no_update
    message = save_uploaded_file(contents, filename or "curva_cliente.xlsx", RUTA_CURVA_CLIENTE)
    return message, (token or 0) + 1


# ============================================================
# SERVIR IMÁGENES GENERADAS DEL MODELO
# ============================================================

# Dash sirve automáticamente /assets. Para no copiar imágenes cada vez,
# copiamos al inicio las gráficas relevantes a assets/assets_generated.
def sync_generated_assets() -> None:
    target_dir = PROJECT_DIR / "assets" / "assets_generated"
    target_dir.mkdir(parents=True, exist_ok=True)

    if not GRAPH_DIR.exists():
        return

    for png in GRAPH_DIR.glob("*.png"):
        target = target_dir / png.name
        try:
            if not target.exists() or png.stat().st_mtime > target.stat().st_mtime:
                target.write_bytes(png.read_bytes())
        except Exception:
            pass


sync_generated_assets()


if __name__ == "__main__":
    app.run(debug=True, port=8050)
