# -*- coding: utf-8 -*-
"""
Vision grafica del patrimonio, organizada en 4 categorias (Bancos, Activos,
Renta Variable y Cripto): snapshot actual por posicion y evolucion semanal
historica (patrimonio, cartera/riesgo/uso de cripto, categoria de RV).

Lee EXCLUSIVAMENTE en modo lectura G:\\Mi unidad\\Marmol\\3-Finances\\finance.xlsx
(ver data_loader.py). No escribe nada fuera de "Finance Dashboard App".
"""

import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data_loader import (
    CARTERA_ORDER,
    EXCEL_PATH,
    RIESGO_ORDER,
    USUFRUCTO_ORDER,
    load_data_checked,
)

# --------------------------------------------------------------------------
# Paleta (ver skill "dataviz"): un color fijo por categoria, consistente en
# toda la pagina; verde/rojo reservados para ganancia/perdida.
# El dashboard es solo modo oscuro (decision del usuario, sin toggle de tema).
# --------------------------------------------------------------------------
COLORS = {
    "surface": "#1a1a19", "page": "#0d0d0d", "text": "#ffffff",
    "text_muted": "#c3c2b7", "axis": "#898781", "grid": "#2c2c2a",
    "card_bg": "#232322", "card_border": "rgba(255,255,255,0.10)",
    "blue": "#3987e5", "orange": "#d95926", "aqua": "#199e70",
    "yellow": "#c98500", "magenta": "#d55181", "good": "#0ca30c", "critical": "#e66767",
}

CAT_COLOR = {"Banca": "blue", "Cripto": "orange", "Renta Variable": "aqua", "Otros activos": "yellow"}

# Cartera y riesgo cripto: color fijo por categoria (nunca se reasigna segun
# filtros), en el orden que ya usa data_loader (CARTERA_ORDER / RIESGO_ORDER).
CARTERA_COLOR_KEYS = {"Cartera Old": "blue", "Cartera New": "orange", "RD26": "aqua", "DEFI": "yellow", "HYPE": "magenta"}
RIESGO_COLOR_KEYS = {"Liquided": "blue", "LP": "aqua", "Medium": "yellow", "Gem": "orange", "Crash": "critical"}
USUFRUCTO_COLOR_KEYS = {"Hold": "blue", "Stake": "aqua", "Farm": "yellow"}
CATEGORIA_ORDER = ["LIQUIDED", "FONDOS", "MP", "ETF", "ACC"]
CATEGORIA_COLOR_KEYS = {"LIQUIDED": "blue", "FONDOS": "orange", "MP": "aqua", "ETF": "yellow", "ACC": "magenta"}
# Nombres formales de las categorias de renta variable (los codigos del Excel
# no son autoexplicativos para el usuario).
CATEGORIA_LABELS = {
    "LIQUIDED": "Liquidez", "FONDOS": "Fondos", "MP": "Materias primas",
    "ETF": "ETF", "ACC": "Bolsa",
}

st.set_page_config(page_title="Patrimonio · Posiciones", page_icon="\U0001F4B0", layout="wide")

# --------------------------------------------------------------------------
# Estado / helpers
# --------------------------------------------------------------------------
if "data" not in st.session_state:
    st.session_state.data = None
    st.session_state.loaded_at = None


def fmt_eur(value, decimals=0, signed=False):
    if value is None or pd.isna(value):
        return "—"
    sign = "-" if value < 0 else ("+" if (signed and value > 0) else "")
    value = abs(value)
    s = f"{value:,.{decimals}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    return f"{sign}{s} €"


def fmt_pct(value, decimals=1, signed=False):
    if value is None or pd.isna(value):
        return "—"
    sign = "+" if (signed and value > 0) else ""
    return f"{sign}{value:.{decimals}f}%".replace(".", ",")


def fmt_num(value, decimals=4):
    if value is None or pd.isna(value):
        return "—"
    return f"{value:,.{decimals}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def load(force=False):
    if force or st.session_state.data is None:
        st.session_state.data = load_data_checked()
        st.session_state.loaded_at = datetime.datetime.now()
    return st.session_state.data


with st.sidebar:
    st.markdown("### ⚙️ Panel")
    refresh_clicked = st.button("🔄 Actualizar datos desde finance.xlsx", use_container_width=True)

c = COLORS


def apply_layout(fig, height=360, showlegend=False):
    fig.update_layout(
        paper_bgcolor=c["surface"], plot_bgcolor=c["surface"],
        font=dict(color=c["text"], family="system-ui, -apple-system, Segoe UI, sans-serif", size=13),
        legend=dict(font=dict(color=c["text_muted"], size=12), orientation="h", yanchor="bottom", y=-0.2),
        margin=dict(l=10, r=10, t=10, b=10), height=height, showlegend=showlegend,
    )
    fig.update_xaxes(gridcolor=c["grid"], zerolinecolor=c["axis"], linecolor=c["axis"], color=c["text_muted"])
    fig.update_yaxes(gridcolor=c["grid"], zerolinecolor=c["axis"], linecolor=c["axis"], color=c["text_muted"])
    return fig


st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {c['page']}; }}
    .kpi-card {{ background: {c['card_bg']}; border: 1px solid {c['card_border']};
                 border-radius: 12px; padding: 16px 20px; margin-bottom: 10px; }}
    .kpi-label {{ color: {c['text_muted']}; font-size: 0.82rem; text-transform: uppercase; letter-spacing: .04em; }}
    .kpi-value {{ color: {c['text']}; font-size: 1.35rem; font-weight: 700; margin-top: 2px; white-space: nowrap; }}
    .kpi-sub {{ color: {c['text_muted']}; font-size: 0.8rem; margin-top: 2px; }}
    .kpi-compact {{ padding: 9px 14px; margin-bottom: 8px; }}
    .kpi-compact .kpi-label {{ font-size: 0.68rem; }}
    .kpi-compact .kpi-value {{ font-size: 1.05rem; margin-top: 1px; }}
    .kpi-compact .kpi-sub {{ font-size: 0.68rem; }}
    .hero-row {{ display: flex; align-items: baseline; gap: 12px; flex-wrap: wrap; }}
    .hero-total {{ color: {c['text']}; font-size: 1.9rem; font-weight: 800; white-space: nowrap; }}
    .hero-label {{ color: {c['text_muted']}; font-size: 0.8rem; text-transform: uppercase; letter-spacing: .04em; white-space: nowrap; }}
    .data-table-wrap {{ background: {c['card_bg']}; border: 1px solid {c['card_border']};
                         border-radius: 12px; overflow: hidden; }}
    .data-table {{ width: 100%; border-collapse: collapse; font-size: 0.92rem; }}
    .data-table th {{ color: {c['text_muted']}; font-size: 0.72rem; text-transform: uppercase;
                       letter-spacing: .04em; font-weight: 600; text-align: left;
                       padding: 10px 16px; border-bottom: 1px solid {c['card_border']}; }}
    .data-table td {{ color: {c['text']}; padding: 9px 16px; border-bottom: 1px solid {c['grid']}; }}
    .data-table tbody tr:last-child td {{ border-bottom: none; }}
    .data-table td.num, .data-table th.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .data-table td.muted {{ color: {c['text_muted']}; }}
    .kpi-mini {{ padding: 5px 10px; margin-bottom: 4px; }}
    .kpi-mini .kpi-label {{ font-size: 0.6rem; }}
    .kpi-mini .kpi-value {{ font-size: 0.8rem; margin-top: 0; }}
    .kpi-mini .kpi-sub {{ font-size: 0.58rem; }}
    .data-table-wrap.data-table-compact {{ width: fit-content; max-width: 100%; margin: 0 auto; }}
    .data-table-compact .data-table {{ width: auto; }}
    .data-table-compact th, .data-table-compact td {{ text-align: center; padding: 10px 30px; white-space: nowrap; }}
    .data-table-compact th.num, .data-table-compact td.num {{ text-align: center; }}
    .data-table-dense th, .data-table-dense td {{ padding-left: 14px; padding-right: 14px; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def hex_to_rgba(hex_color, alpha):
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alpha})"


def render_table(df, columns, row_bg_col=None, row_bg_map=None, compact_width=False):
    """Tabla HTML formal. columns: lista de (clave, cabecera, align, formatter).
    align 'right' activa tabular-nums (para columnas numericas); formatter
    recibe el valor crudo y devuelve el string ya formateado.
    row_bg_col/row_bg_map: colorea el fondo de toda la fila segun el valor de
    esa columna (p.ej. distinguir tipo de cuenta).
    compact_width: la tabla se ajusta a su contenido y se centra (en vez de
    ocupar el ancho completo de la columna), con el texto centrado — pensado
    para tablas de pocas columnas donde el 100% de ancho deja huecos.
    """
    head = "".join(
        f'<th class="{"num" if align == "right" else ""}">{header}</th>'
        for _, header, align, _ in columns
    )
    body = []
    for _, row in df.iterrows():
        cells = "".join(
            f'<td class="{"num" if align == "right" else ""}">'
            f'{(fmt(row[key]) if fmt else row[key])}</td>'
            for key, _, align, fmt in columns
        )
        row_style = ""
        if row_bg_col and row_bg_map:
            bg = row_bg_map.get(row[row_bg_col])
            if bg:
                row_style = f' style="background:{bg};"'
        body.append(f"<tr{row_style}>{cells}</tr>")
    wrap_class = "data-table-wrap data-table-compact" if compact_width else "data-table-wrap"
    st.markdown(
        f"""<div class="{wrap_class}"><table class="data-table">
                <thead><tr>{head}</tr></thead>
                <tbody>{''.join(body)}</tbody>
            </table></div>""",
        unsafe_allow_html=True,
    )


def render_cartera_perf_table(filas, compact_width=False):
    """Tabla de rendimiento por cartera (Actual/Invertido/Profit/+/- 1S..1A),
    con el fondo de cada celda de porcentaje coloreado segun signo (verde/rojo).
    filas: lista de dicts con las claves de _load_cripto_rendimiento.
    """
    def pct_cell(v):
        if v is None:
            return '<td class="num muted">—</td>'
        bg = hex_to_rgba(c["good"], 0.20) if v >= 0 else hex_to_rgba(c["critical"], 0.20)
        fg = c["good"] if v >= 0 else c["critical"]
        return f'<td class="num" style="background:{bg};color:{fg};font-weight:600;">{fmt_pct(v * 100, 2, signed=True)}</td>'

    def eur_cell(v):
        if v is None or abs(v) < 0.005:
            return '<td class="num muted">—</td>'
        return f'<td class="num">{fmt_eur(v, 2)}</td>'

    headers = ["Cartera", "Actual", "Invertido", "Profit", "+/- 1S", "+/- 1M", "+/- 3M", "+/- 6M", "+/- 1A"]
    head = "".join(f'<th class="{"" if h == "Cartera" else "num"}">{h}</th>' for h in headers)
    body = []
    for f in filas:
        cells = f'<td>{f["cartera"]}</td>'
        cells += eur_cell(f["actual"]) + eur_cell(f["invertido"]) + pct_cell(f["profit_pct"])
        cells += "".join(pct_cell(f[k]) for k in ("ret_1s", "ret_1m", "ret_3m", "ret_6m", "ret_1a"))
        body.append(f"<tr>{cells}</tr>")
    wrap_class = "data-table-wrap data-table-compact data-table-dense" if compact_width else "data-table-wrap"
    st.markdown(
        f"""<div class="{wrap_class}"><table class="data-table">
                <thead><tr>{head}</tr></thead>
                <tbody>{''.join(body)}</tbody>
            </table></div>""",
        unsafe_allow_html=True,
    )


def render_evolucion_apilada(series_dict, order, color_keys, value_key="valor", label_map=None, dtick="M3", height=380):
    """Grafico de area apilada semana a semana a partir de un dict
    {clave: [{"fecha": ..., value_key: ...}, ...]} — el formato que
    devuelven las funciones `*_evolucion` de data_loader.py. Una serie por
    clave de `order`, coloreada con `color_keys`. Comun a las 4 pestanas de
    evolucion por categoria (cartera, riesgo y uso de cripto, categoria de RV).
    """
    df = None
    for key in order:
        serie = pd.DataFrame(series_dict[key])[["fecha", value_key]]
        serie = serie.rename(columns={value_key: key})
        df = serie if df is None else df.merge(serie, on="fecha", how="outer")
    df = df.sort_values("fecha").fillna(0)

    fig = go.Figure()
    for key in order:
        color = c[color_keys.get(key, "text_muted")]
        label = (label_map or {}).get(key, key)
        fig.add_trace(go.Scatter(
            x=df["fecha"], y=df[key], name=label,
            mode="lines", stackgroup="one", line=dict(width=0.5, color=color),
            fillcolor=hex_to_rgba(color, 0.65),
            hovertemplate=f"{label}: " + "%{y:,.0f} €<extra></extra>",
        ))
    fig.update_layout(separators=",.", hovermode="x unified")
    fig.update_xaxes(tickformat="%b %Y", dtick=dtick, showgrid=False)
    fig.update_yaxes(tickformat=",.0f", ticksuffix=" €")
    apply_layout(fig, height=height, showlegend=True)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def kpi_card(label, value_str, sub=None, value_color=None, dot_color=None, compact=False, mini=False):
    style = f' style="color:{value_color}"' if value_color else ""
    dot = (
        f'<span style="display:inline-block;width:8px;height:8px;border-radius:50%;'
        f'background:{dot_color};margin-right:7px;"></span>'
        if dot_color else ""
    )
    card_class = "kpi-card"
    if compact:
        card_class += " kpi-compact"
    if mini:
        card_class += " kpi-mini"
    st.markdown(
        f"""<div class="{card_class}">
                <div class="kpi-label">{dot}{label}</div>
                <div class="kpi-value"{style}>{value_str}</div>
                <div class="kpi-sub">{sub or ""}</div>
            </div>""",
        unsafe_allow_html=True,
    )


if refresh_clicked:
    load(force=True)
    st.rerun()

with st.sidebar:
    st.caption(f"Fuente: `{EXCEL_PATH}`")
    if st.session_state.loaded_at:
        st.caption(f"Ultima lectura: {st.session_state.loaded_at.strftime('%d/%m/%Y %H:%M:%S')}")
    st.divider()
    st.caption("Snapshot del patrimonio actual, con evolución histórica semanal donde el Excel la registra.")

try:
    data = load()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()

if data.get("warnings"):
    with st.container():
        st.warning(
            "⚠️ Algunos totales no cuadran al leer finance.xlsx ahora mismo — probable "
            "lectura a medio-sincronizar (el archivo esta en Google Drive y se detecto "
            "abierto/editandose). Prueba a **guardar el Excel** y pulsar "
            "**'🔄 Actualizar datos'** en la barra lateral.",
            icon="⚠️",
        )
        with st.expander("Ver detalle tecnico de lo que no cuadra"):
            for w in data["warnings"]:
                st.markdown(f"- {w}")

pat = data["patrimonio"]

# ==========================================================================
# PATRIMONIO TOTAL — cifra total + tarjetas KPI (sin grafico de composicion)
# ==========================================================================
st.markdown(
    f'<div class="hero-row">'
    f'<span class="hero-label">Patrimonio total</span>'
    f'<span class="hero-total">{fmt_eur(pat["total"], 2)}</span>'
    f'</div>',
    unsafe_allow_html=True,
)

st.write("")
# Las 4 categorias en una sola fila horizontal (sin grafico de distribucion
# al lado): mas simple que el grid 2x2 + barra que habia antes.
col1, col2, col3, col4 = st.columns(4)
with col1:
    kpi_card("🏦 Bancos", fmt_eur(pat["banca"]), fmt_pct(pat["banca"] / pat["total"] * 100) + " del total",
              dot_color=c[CAT_COLOR["Banca"]], compact=True, mini=True)
with col2:
    kpi_card("💼 Activos", fmt_eur(pat["activos_otros"]), fmt_pct(pat["activos_otros"] / pat["total"] * 100) + " del total",
              dot_color=c[CAT_COLOR["Otros activos"]], compact=True, mini=True)
with col3:
    kpi_card("📈 Renta variable", fmt_eur(pat["renta_variable"]), fmt_pct(pat["renta_variable"] / pat["total"] * 100) + " del total",
              dot_color=c[CAT_COLOR["Renta Variable"]], compact=True, mini=True)
with col4:
    kpi_card("🪙 Cripto", fmt_eur(pat["cripto"]), fmt_pct(pat["cripto"] / pat["total"] * 100) + " del total",
              dot_color=c[CAT_COLOR["Cripto"]], compact=True, mini=True)

# --------------------------------------------------------------------------
# EVOLUCION DEL PATRIMONIO — replica de "Grafico 1" (hoja CashFlow del
# Excel): area apilada Bancos (liquidez) + Activos (cripto + renta variable
# + otros activos), semana a semana desde el primer registro. Mismos datos
# y misma forma que el original; solo cambia el formato visual. Se queda
# dentro de la seccion "Patrimonio total" (Parte 1), antes del divider.
# --------------------------------------------------------------------------
df_evo = pd.DataFrame(data["evolucion"])
with st.expander(f"📈 Evolución del patrimonio ({df_evo['fecha'].min().year} – hoy)"):
    fig_evo = go.Figure()
    fig_evo.add_trace(go.Scatter(
        x=df_evo["fecha"], y=df_evo["banca"], name="Bancos", mode="lines",
        stackgroup="one", line=dict(width=1, color=c["blue"]),
        fillcolor=hex_to_rgba(c["blue"], 0.55),
        hovertemplate="Bancos: %{y:,.0f} €<extra></extra>",
    ))
    fig_evo.add_trace(go.Scatter(
        x=df_evo["fecha"], y=df_evo["activos"], name="Activos (cripto + RV + otros)", mode="lines",
        stackgroup="one", line=dict(width=1, color=c["magenta"]),
        fillcolor=hex_to_rgba(c["magenta"], 0.55),
        hovertemplate="Activos: %{y:,.0f} €<extra></extra>",
    ))
    fig_evo.update_layout(separators=",.", hovermode="x unified")
    fig_evo.update_xaxes(tickformat="%Y", dtick="M12", showgrid=False)
    fig_evo.update_yaxes(tickformat=",.0f", ticksuffix=" €")
    apply_layout(fig_evo, height=380, showlegend=True)
    st.plotly_chart(fig_evo, use_container_width=True, config={"displayModeBar": False})
    st.caption(
        "Replica 'Gráfico 1' de la hoja CashFlow del Excel: patrimonio total "
        "apilado en Bancos (liquidez) y Activos (cripto + renta variable + "
        "otros activos), semana a semana desde el primer registro."
    )

st.divider()

# ==========================================================================
# POSICIONES POR CATEGORIA (grafico, no tablas)
# ==========================================================================
st.markdown(f'<div class="hero-label" style="margin-bottom:8px;">TUS POSICIONES</div>', unsafe_allow_html=True)

tab_bancos, tab_activos, tab_rv, tab_cripto = st.tabs(
    ["🏦 Bancos", "💼 Activos", "📈 Renta variable", "🪙 Cripto"]
)


with tab_bancos:
    df_b = pd.DataFrame(data["bancos"])

    # La deuda/credito (importes negativos) ya esta incluida en el total
    # bancario oficial; se excluye de la tabla de cuentas para no mezclar
    # patrimonio con deuda, y no se muestra como KPI aparte.
    df_cuentas = df_b[df_b["importe"] >= 0].copy()
    df_cuentas["pct"] = df_cuentas["importe"] / pat["banca"] * 100 if pat["banca"] else 0
    df_cuentas = df_cuentas.sort_values("importe", ascending=False).reset_index(drop=True)

    st.caption(f"{len(df_cuentas)} cuentas · Total: {fmt_eur(pat['banca'], 2)}")

    # Fondo de fila por tipo de cuenta (tinte sutil sobre el tema oscuro,
    # coherente con los colores ya usados en el resto del dashboard).
    TIPO_ROW_BG = {
        "BANCO": hex_to_rgba(c["aqua"], 0.16),
        "CASH": hex_to_rgba(c["orange"], 0.16),
        "SHOP": hex_to_rgba(c["blue"], 0.16),
    }
    render_table(
        df_cuentas,
        columns=[
            ("proveedor", "Tipo", "left", None),
            ("detalle", "Cuenta", "left", None),
            ("importe", "Importe", "right", lambda v: fmt_eur(v, 2)),
            ("pct", "% del total", "right", lambda v: fmt_pct(v, 1)),
        ],
        row_bg_col="proveedor",
        row_bg_map=TIPO_ROW_BG,
        compact_width=True,
    )

with tab_activos:
    df_a = pd.DataFrame(data["activos_otros"])
    df_a["pct"] = df_a["importe"] / pat["activos_otros"] * 100 if pat["activos_otros"] else 0
    df_a = df_a.sort_values("importe", ascending=False).reset_index(drop=True)
    st.caption(f"{len(df_a)} posiciones · Total: {fmt_eur(pat['activos_otros'], 2)}")
    render_table(
        df_a,
        columns=[
            ("tipo", "Tipo", "left", None),
            ("detalle", "Detalle", "left", None),
            ("importe", "Importe", "right", lambda v: fmt_eur(v, 2)),
            ("pct", "% del total", "right", lambda v: fmt_pct(v, 1)),
        ],
        compact_width=True,
    )

with tab_rv:
    df_rv_raw = pd.DataFrame(data["rv_posiciones"])
    rv_resumen = data["rv_resumen"]
    invertido_total = rv_resumen["invertido"]
    rendimiento_total = rv_resumen["rendimiento_pct"]

    # 3 tarjetas mini (izquierda) + grafico de categorias (derecha), sin
    # titulo: mismo patron de alineacion pixel-perfect que la fila de
    # patrimonio total (altura del grafico ajustada al alto real del bloque
    # de tarjetas, medido en el navegador).
    col_kpis, col_cat = st.columns([1, 2])
    with col_kpis:
        kpi_card("Capital actual", fmt_eur(pat["renta_variable"], 2), compact=True, mini=True)
        kpi_card("Capital invertido", fmt_eur(invertido_total, 2), "total oficial (hoja GUI)", compact=True, mini=True)
        gain_color = c["good"] if (rendimiento_total or 0) >= 0 else c["critical"]
        kpi_card("Rendimiento", fmt_pct(rendimiento_total, 2, signed=True), "sobre el capital invertido",
                  value_color=gain_color, compact=True, mini=True)

    with col_cat:
        df_rvcat = pd.DataFrame(data["rv_categoria"])
        df_rvcat["label"] = df_rvcat["categoria"].map(CATEGORIA_LABELS).fillna(df_rvcat["categoria"])
        df_rvcat = df_rvcat.sort_values("importe", ascending=True).reset_index(drop=True)
        total_cat = df_rvcat["importe"].sum()
        cat_colors = [c[CATEGORIA_COLOR_KEYS.get(cat, "text_muted")] for cat in df_rvcat["categoria"]]
        # Mismo formato de etiqueta que el grafico de posiciones de abajo:
        # "importe (porcentaje)".
        cat_text = [
            f"{fmt_eur(v, 0)} ({fmt_pct(v / total_cat * 100 if total_cat else 0, 1)})"
            for v in df_rvcat["importe"]
        ]
        fig_cat = go.Figure(go.Bar(
            x=df_rvcat["importe"], y=df_rvcat["label"], orientation="h",
            marker=dict(color=cat_colors), text=cat_text, textposition="outside", cliponaxis=False,
        ))
        fig_cat.update_xaxes(range=[0, df_rvcat["importe"].max() * 1.35])
        # Altura igualada al bloque de 3 tarjetas mini de la izquierda
        # (medida con getBoundingClientRect, igual que en patrimonio total).
        apply_layout(fig_cat, height=187)
        fig_cat.update_layout(margin=dict(l=0, r=0, t=20, b=0))
        st.plotly_chart(fig_cat, use_container_width=True, config={"displayModeBar": False})

    st.write("")
    st.caption(
        "Cada posicion, agrupada y coloreada por categoria — la etiqueta muestra la "
        "posicion total y, entre parentesis, el % y el importe de beneficio/perdida "
        "sobre el capital invertido en esa posicion"
    )

    # --------------------------------------------------------------------
    # EVOLUCION DE RENTA VARIABLE — replica de "Grafico 1" (hoja GUI del
    # Excel): linea de valor actual vs. capital invertido, semana a semana
    # (rango reciente que registra esa hoja). Mismos datos y misma forma que
    # el original; solo cambia el formato visual.
    # --------------------------------------------------------------------
    df_rvevo = pd.DataFrame(data["rv_evolucion"])
    with st.expander(f"📈 Evolución de Renta Variable ({df_rvevo['fecha'].min():%b %Y} – hoy)"):
        fig_rvevo = go.Figure()
        fig_rvevo.add_trace(go.Scatter(
            x=df_rvevo["fecha"], y=df_rvevo["invertido"], name="Capital invertido",
            mode="lines", line=dict(width=1.5, color=c["text_muted"], dash="dash"),
            hovertemplate="Invertido: %{y:,.0f} €<extra></extra>",
        ))
        fig_rvevo.add_trace(go.Scatter(
            x=df_rvevo["fecha"], y=df_rvevo["actual"], name="Valor actual",
            mode="lines", line=dict(width=2.5, color=c[CAT_COLOR["Renta Variable"]]),
            hovertemplate="Valor actual: %{y:,.0f} €<extra></extra>",
        ))
        fig_rvevo.update_layout(separators=",.", hovermode="x unified")
        fig_rvevo.update_xaxes(tickformat="%b %Y", dtick="M1", showgrid=False)
        fig_rvevo.update_yaxes(tickformat=",.0f", ticksuffix=" €")
        apply_layout(fig_rvevo, height=340, showlegend=True)
        st.plotly_chart(fig_rvevo, use_container_width=True, config={"displayModeBar": False})
        st.caption(
            "Replica 'Gráfico 1' de la hoja GUI del Excel: valor actual de renta "
            "variable frente a capital invertido, semana a semana."
        )

    # --------------------------------------------------------------------
    # EVOLUCION POR CATEGORIA DE RV — Evolucion!S:W (ver
    # _load_rv_categoria_evolucion en data_loader.py), mismo rango de
    # fechas que "Evolución de Renta Variable".
    # --------------------------------------------------------------------
    with st.expander("📈 Evolución de RV por categoría"):
        render_evolucion_apilada(
            data["rv_categoria_evolucion"], CATEGORIA_ORDER, CATEGORIA_COLOR_KEYS,
            label_map=CATEGORIA_LABELS, dtick="M1", height=340,
        )
        st.caption(
            "Valor por categoría apilado, semana a semana — leído de la hoja "
            "Evolucion del Excel."
        )

    # Orden: por categoria (fijo) y, dentro de cada categoria, de mayor a menor valor.
    df_rv = df_rv_raw.copy()
    # 'activo' se repite alguna vez (p.ej. 'EUR' en dos brokers distintos);
    # sin desambiguar, un eje categorico con la misma etiqueta dos veces
    # fusiona esas filas en una sola barra. Añadimos el campo entre parentesis
    # solo en esos casos para no repetir etiqueta.
    dup_activo = df_rv["activo"].duplicated(keep=False)
    df_rv["etiqueta"] = df_rv["activo"]
    df_rv.loc[dup_activo, "etiqueta"] = df_rv["activo"] + " (" + df_rv["campo"] + ")"
    df_rv["cat_order"] = df_rv["categoria"].map(lambda x: CATEGORIA_ORDER.index(x) if x in CATEGORIA_ORDER else 99)
    df_rv = df_rv.sort_values(["cat_order", "valor_actual"], ascending=[True, True]).reset_index(drop=True)

    # Beneficio (% e importe) calculado SIEMPRE a partir del capital invertido
    # de cada posicion (no del total oficial de arriba, que es algo superior
    # por las comisiones que cobra el broker y no se reparten por posicion).
    tiene_coste = df_rv["invertido"].notna() & (df_rv["invertido"] != 0)
    df_rv["gain_abs"] = (df_rv["valor_actual"] - df_rv["invertido"]).where(tiene_coste)
    df_rv["gain_pct"] = (df_rv["gain_abs"] / df_rv["invertido"] * 100).where(tiene_coste)

    bar_colors = [c[CATEGORIA_COLOR_KEYS.get(cat, "text_muted")] for cat in df_rv["categoria"]]
    text_colors = [
        c["good"] if pd.notna(p) and p >= 0 else (c["critical"] if pd.notna(p) else c["text_muted"])
        for p in df_rv["gain_pct"]
    ]
    # Etiqueta unificada: "POSICION TOTAL € (% BENEFICIO | BENEFICIO ABSOLUTO €)".
    # Sin capital invertido (TRADE REP/MYINV: efectivo aun sin invertir) no hay
    # beneficio que calcular, asi que se queda solo en la posicion total.
    text_labels = [
        f"{fmt_eur(v, 0)} ({fmt_pct(p, 1, signed=True)} | {fmt_eur(a, 0, signed=True)})"
        if pd.notna(p) else fmt_eur(v, 0)
        for v, p, a in zip(df_rv["valor_actual"], df_rv["gain_pct"], df_rv["gain_abs"])
    ]
    # La etiqueta NO se ancla al final de la barra (Bar textposition="outside"):
    # en una posicion perdedora la marca de capital invertido cae mas a la
    # derecha que la propia barra, justo donde empezaria el texto, y lo tapa.
    # En su lugar se coloca en un trace de texto aparte, siempre a la derecha
    # de lo que quede mas lejos entre la barra y la marca.
    label_x = df_rv[["valor_actual", "invertido"]].max(axis=1)
    axis_max = max(label_x.max(), df_rv["valor_actual"].max())
    text_pad = axis_max * 0.02

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_rv["valor_actual"], y=df_rv["etiqueta"], orientation="h",
        marker=dict(color=bar_colors), name="Valor actual",
    ))
    fig.add_trace(go.Scatter(
        x=df_rv["invertido"], y=df_rv["etiqueta"], mode="markers",
        marker=dict(symbol="line-ns", size=16, line=dict(color=c["text_muted"], width=2)),
        name="Capital invertido",
    ))
    fig.add_trace(go.Scatter(
        x=label_x + text_pad, y=df_rv["etiqueta"], mode="text",
        text=text_labels, textposition="middle right", textfont=dict(color=text_colors),
        showlegend=False, hoverinfo="skip",
    ))
    fig.update_xaxes(range=[0, axis_max * 1.55])
    apply_layout(fig, height=max(420, 30 * len(df_rv)), showlegend=True)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    st.caption(
        "El color de la barra indica la categoria (ver grafico de categorias de arriba); "
        "el texto en verde/rojo es el beneficio/perdida sobre el capital invertido "
        "(marca gris). 'TRADE REP' y 'MYINV' son efectivo aun sin invertir en el "
        "broker: su capital invertido coincide con su valor actual, por eso muestran 0%. "
        f"Nota: la suma del coste de entrada de estas posiciones da "
        f"{fmt_eur(df_rv['invertido'].sum(skipna=True), 2)}, algo menos que el "
        f"'Capital invertido' oficial de arriba ({fmt_eur(invertido_total, 2)}) — la "
        "diferencia es el efecto acumulado de las comisiones que cobra el broker, "
        "que engordan el capital invertido oficial pero no se reparten por posicion."
    )

with tab_cripto:
    # KPIs (3, verticales, mini) + tabla de rendimiento por cartera al lado
    # (en vez de debajo): mismo patron que la fila de Renta variable.
    df_rend = pd.DataFrame(data["cripto_rendimiento"])
    invertido_total_cripto = df_rend["invertido"].fillna(0).sum()
    rendimiento_total_cripto = (
        (pat["cripto"] - invertido_total_cripto) / invertido_total_cripto * 100
        if invertido_total_cripto else None
    )

    col_kpis, col_perf = st.columns([1, 4])
    with col_kpis:
        kpi_card("Valor actual", fmt_eur(pat["cripto"], 2), compact=True, mini=True)
        kpi_card("Capital invertido", fmt_eur(invertido_total_cripto, 2), "suma por cartera (hoja GUI)",
                  compact=True, mini=True)
        gain_color = c["good"] if (rendimiento_total_cripto or 0) >= 0 else c["critical"]
        kpi_card("Rentabilidad",
                  fmt_pct(rendimiento_total_cripto, 2, signed=True) if rendimiento_total_cripto is not None else "—",
                  "sobre el capital invertido", value_color=gain_color, compact=True, mini=True)

    with col_perf:
        render_cartera_perf_table(data["cripto_rendimiento"], compact_width=True)

    st.write("")
    col_bar1, col_bar2, col_bar3 = st.columns(3)
    with col_bar1:
        st.caption("Distribución por cartera")
        df_cart = pd.DataFrame(data["cripto_cartera"]).sort_values("actual", ascending=True)
        total_cart = df_cart["actual"].sum()
        colors = [c[CARTERA_COLOR_KEYS.get(x, "text_muted")] for x in df_cart["cartera"]]
        text = [f"{fmt_eur(v, 0)} ({fmt_pct(v / total_cart * 100 if total_cart else 0, 1)})" for v in df_cart["actual"]]
        fig = go.Figure(go.Bar(
            x=df_cart["actual"], y=df_cart["cartera"], orientation="h",
            marker=dict(color=colors), text=text, textposition="outside", cliponaxis=False,
        ))
        fig.update_xaxes(range=[0, df_cart["actual"].max() * 1.35])
        apply_layout(fig, height=max(220, 44 * len(df_cart)))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    with col_bar2:
        st.caption("Distribución por riesgo")
        df_r = pd.DataFrame(data["cripto_riesgo"]).sort_values("actual", ascending=True)
        total_r = df_r["actual"].sum()
        colors = [c[RIESGO_COLOR_KEYS.get(x, "text_muted")] for x in df_r["riesgo"]]
        text = [f"{fmt_eur(v, 0)} ({fmt_pct(v / total_r * 100 if total_r else 0, 1)})" for v in df_r["actual"]]
        fig = go.Figure(go.Bar(
            x=df_r["actual"], y=df_r["riesgo"], orientation="h",
            marker=dict(color=colors), text=text, textposition="outside", cliponaxis=False,
        ))
        fig.update_xaxes(range=[0, df_r["actual"].max() * 1.35])
        apply_layout(fig, height=max(220, 44 * len(df_r)))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    with col_bar3:
        st.caption("Distribución por uso")
        df_uso_all = pd.DataFrame(data["cripto_uso_cross"]).groupby("uso")["valor"].sum()
        df_uso_all = df_uso_all.reindex(USUFRUCTO_ORDER, fill_value=0).sort_values(ascending=True)
        total_uso_all = df_uso_all.sum()
        colors = [c[USUFRUCTO_COLOR_KEYS.get(x, "text_muted")] for x in df_uso_all.index]
        text = [f"{fmt_eur(v, 0)} ({fmt_pct(v / total_uso_all * 100 if total_uso_all else 0, 1)})" for v in df_uso_all.values]
        fig = go.Figure(go.Bar(
            x=df_uso_all.values, y=df_uso_all.index, orientation="h",
            marker=dict(color=colors), text=text, textposition="outside", cliponaxis=False,
        ))
        fig.update_xaxes(range=[0, df_uso_all.max() * 1.35])
        apply_layout(fig, height=max(220, 44 * len(df_uso_all)))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.write("")

    # --------------------------------------------------------------------
    # EVOLUCION HISTORICA GENERAL — area apilada con el valor actual de las
    # 5 carteras semana a semana (Evolucion!X:AJ, ver cartera_evolucion),
    # mismo patron que "Evolución de Renta Variable".
    # --------------------------------------------------------------------
    with st.expander(f"📈 Evolución histórica de Cripto"):
        render_evolucion_apilada(data["cartera_evolucion"], CARTERA_ORDER, CARTERA_COLOR_KEYS, value_key="actual")
        st.caption(
            "Valor actual de cada cartera apilado, semana a semana — leído de la "
            "hoja Evolucion del Excel."
        )

    # --------------------------------------------------------------------
    # EVOLUCION HISTORICA POR RIESGO — igual que el anterior pero por nivel
    # de riesgo en vez de por cartera (Evolucion!AL:AP, ver
    # _load_cripto_riesgo_evolucion en data_loader.py).
    # --------------------------------------------------------------------
    with st.expander("📈 Evolución histórica por riesgo"):
        render_evolucion_apilada(data["cripto_riesgo_evolucion"], RIESGO_ORDER, RIESGO_COLOR_KEYS)
        st.caption(
            "Valor por nivel de riesgo apilado, semana a semana — leído de la "
            "hoja Evolucion del Excel."
        )

    # --------------------------------------------------------------------
    # EVOLUCION HISTORICA POR USO — mismo planteamiento que la de riesgo,
    # pero para usufructo (Hold/Stake/Farm), Evolucion!AQ:AS (ver
    # _load_cripto_uso_evolucion en data_loader.py).
    # --------------------------------------------------------------------
    with st.expander("📈 Evolución histórica por uso"):
        render_evolucion_apilada(data["cripto_uso_evolucion"], USUFRUCTO_ORDER, USUFRUCTO_COLOR_KEYS)
        st.caption(
            "Valor por tipo de uso apilado, semana a semana — leído de la hoja "
            "Evolucion del Excel."
        )

    st.write("")

    # --------------------------------------------------------------------
    # TOP 10 POSICIONES (todas las carteras) — mismo mapa de calor en forma
    # de barra que el de "Detalle por cartera" mas abajo, pero sobre
    # data["cripto_posiciones"] (Portfolio!M9:Q164) sin filtrar por cartera,
    # para ver de un vistazo las mayores posiciones de todo el patrimonio
    # cripto.
    # --------------------------------------------------------------------
    st.caption("Top 10 posiciones (todas las carteras)")
    df_top_all = pd.DataFrame(data["cripto_posiciones"])
    df_top_all = df_top_all[df_top_all["capital"] > 0.01]
    df_top_all = df_top_all.sort_values("capital", ascending=True).tail(10)
    fig_heat_all = go.Figure(go.Bar(
        x=df_top_all["capital"], y=df_top_all["coin"], orientation="h",
        marker=dict(color=df_top_all["capital"], colorscale="Inferno",
                    line=dict(color=c["page"], width=1)),
        text=[f"{fmt_eur(v, 0)} ({fmt_pct(p, 1)})" for v, p in zip(df_top_all["capital"], df_top_all["pct_categoria"])],
        textposition="outside", cliponaxis=False,
        hovertemplate="%{y}: %{x:,.0f} €<extra></extra>",
    ))
    fig_heat_all.update_xaxes(range=[0, df_top_all["capital"].max() * 1.35])
    apply_layout(fig_heat_all, height=max(320, 34 * len(df_top_all)))
    st.plotly_chart(fig_heat_all, use_container_width=True, config={"displayModeBar": False})

    st.divider()

    # --------------------------------------------------------------------
    # DETALLE POR CARTERA — selector; riesgo y usufructo exactos por cartera
    # individual (de tablas propias de GUI), lado a lado.
    # --------------------------------------------------------------------
    st.markdown('<div class="hero-label" style="margin-bottom:8px;">DETALLE POR CARTERA</div>', unsafe_allow_html=True)
    cartera_sel = st.selectbox("Selecciona una cartera", CARTERA_ORDER, key="cartera_detalle")

    col_riesgo, col_uso = st.columns(2)
    with col_riesgo:
        st.caption(f"Riesgo — {cartera_sel}")
        df_cross_sel = pd.DataFrame(data["cripto_cross"])
        df_cross_sel = df_cross_sel[df_cross_sel["cartera"] == cartera_sel]
        # Mismas filas (RIESGO_ORDER) siempre, con 0 si esta cartera no tiene
        # posiciones en esa categoria — para que el grafico no cambie de alto
        # ni de orden al cambiar de cartera.
        s_riesgo = df_cross_sel.set_index("riesgo")["valor"].reindex(RIESGO_ORDER, fill_value=0)
        total_riesgo_sel = s_riesgo.sum()
        colors = [c[RIESGO_COLOR_KEYS.get(x, "text_muted")] for x in s_riesgo.index]
        text = [f"{fmt_eur(v, 0)} ({fmt_pct(v / total_riesgo_sel * 100 if total_riesgo_sel else 0, 1)})" for v in s_riesgo.values]
        fig = go.Figure(go.Bar(
            x=s_riesgo.values, y=s_riesgo.index, orientation="h",
            marker=dict(color=colors), text=text, textposition="outside", cliponaxis=False,
        ))
        fig.update_xaxes(range=[0, max(s_riesgo.max(), 1) * 1.4])
        fig.update_yaxes(autorange="reversed")
        apply_layout(fig, height=42 * len(s_riesgo) + 20)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with col_uso:
        st.caption(f"Usufructo — {cartera_sel}")
        df_uso_sel = pd.DataFrame(data["cripto_uso_cross"])
        df_uso_sel = df_uso_sel[df_uso_sel["cartera"] == cartera_sel]
        s_uso = df_uso_sel.set_index("uso")["valor"].reindex(USUFRUCTO_ORDER, fill_value=0)
        total_uso_sel = s_uso.sum()
        colors = [c[USUFRUCTO_COLOR_KEYS.get(x, "text_muted")] for x in s_uso.index]
        text = [f"{fmt_eur(v, 0)} ({fmt_pct(v / total_uso_sel * 100 if total_uso_sel else 0, 1)})" for v in s_uso.values]
        fig = go.Figure(go.Bar(
            x=s_uso.values, y=s_uso.index, orientation="h",
            marker=dict(color=colors), text=text, textposition="outside", cliponaxis=False,
        ))
        fig.update_xaxes(range=[0, max(s_uso.max(), 1) * 1.4])
        fig.update_yaxes(autorange="reversed")
        apply_layout(fig, height=42 * len(s_uso) + 20)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # --------------------------------------------------------------------
    # EVOLUCION ECONOMICA DE LA CARTERA SELECCIONADA — Evolucion!D205:AJ317
    # (aprox): valor actual y, cuando la hoja lo registra, capital invertido,
    # semana a semana, para la cartera elegida arriba (ver
    # CARTERA_EVOLUCION_COLS en data_loader.py).
    # --------------------------------------------------------------------
    df_cart_evo = pd.DataFrame(data["cartera_evolucion"][cartera_sel])
    st.caption(f"Evolución económica — {cartera_sel}")
    fig_cart_evo = go.Figure()
    tiene_invertido = bool(df_cart_evo["invertido"].notna().any())
    if tiene_invertido:
        fig_cart_evo.add_trace(go.Scatter(
            x=df_cart_evo["fecha"], y=df_cart_evo["invertido"], name="Capital invertido",
            mode="lines", line=dict(width=1.5, color=c["text_muted"], dash="dash"),
            hovertemplate="Invertido: %{y:,.0f} €<extra></extra>",
        ))
    fig_cart_evo.add_trace(go.Scatter(
        x=df_cart_evo["fecha"], y=df_cart_evo["actual"], name="Valor actual",
        mode="lines", line=dict(width=2.5, color=c[CARTERA_COLOR_KEYS.get(cartera_sel, "text_muted")]),
        hovertemplate="Valor actual: %{y:,.0f} €<extra></extra>",
    ))
    fig_cart_evo.update_layout(separators=",.", hovermode="x unified")
    fig_cart_evo.update_xaxes(tickformat="%b %Y", dtick="M3", showgrid=False)
    fig_cart_evo.update_yaxes(tickformat=",.0f", ticksuffix=" €")
    apply_layout(fig_cart_evo, height=340, showlegend=tiene_invertido)
    st.plotly_chart(fig_cart_evo, use_container_width=True, config={"displayModeBar": False})
    if not tiene_invertido:
        st.caption(f"'{cartera_sel}' no tiene capital invertido registrado en esta hoja.")

    # --------------------------------------------------------------------
    # MAPA DE CALOR — top 10 posiciones de la cartera seleccionada, por
    # capital (Portfolio, columna de capital propia de cada cartera — ver
    # CARTERA_CAPITAL_COL en data_loader.py). Barra horizontal con escala de
    # color continua (en vez de Treemap/Heatmap nativos: se probaron ambos y
    # esta version de Plotly/Streamlit los renderiza en blanco de forma
    # intermitente — un Bar normal es el tipo mas fiable en todo el resto
    # del dashboard, coloreado aqui por magnitud para dar el mismo efecto
    # visual de "mapa de calor").
    # --------------------------------------------------------------------
    st.caption(f"Top 10 posiciones — {cartera_sel}")
    df_top_sel = pd.DataFrame(data["cripto_top_por_cartera"].get(cartera_sel, []))
    if df_top_sel.empty:
        st.info("Sin posiciones para esta cartera.")
    else:
        df_top_sel = df_top_sel.sort_values("capital", ascending=True).tail(10)
        fig_heat = go.Figure(go.Bar(
            x=df_top_sel["capital"], y=df_top_sel["coin"], orientation="h",
            marker=dict(color=df_top_sel["capital"], colorscale="Inferno",
                        line=dict(color=c["page"], width=1)),
            text=[fmt_eur(v, 0) for v in df_top_sel["capital"]], textposition="outside", cliponaxis=False,
            hovertemplate="%{y}: %{x:,.0f} €<extra></extra>",
        ))
        fig_heat.update_xaxes(range=[0, df_top_sel["capital"].max() * 1.35])
        apply_layout(fig_heat, height=max(320, 34 * len(df_top_sel)))
        st.plotly_chart(fig_heat, use_container_width=True, config={"displayModeBar": False})
