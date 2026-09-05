# -*- coding: utf-8 -*-
"""
Vision grafica del patrimonio actual (snapshot, sin evolucion temporal),
organizada en 4 categorias: Bancos, Activos, Renta Variable y Cripto.

Lee EXCLUSIVAMENTE en modo lectura G:\\Mi unidad\\Marmol\\3-Finances\\finance.xlsx
(ver data_loader.py). No escribe nada fuera de "Finance Dashboard App".
"""

import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data_loader import CARTERA_ORDER, EXCEL_PATH, RIESGO_ORDER, load_data_checked

# --------------------------------------------------------------------------
# Paleta (ver skill "dataviz"): un color fijo por categoria, consistente en
# toda la pagina; verde/rojo reservados para ganancia/perdida.
# --------------------------------------------------------------------------
PALETTE = {
    "light": {
        "surface": "#fcfcfb", "page": "#f9f9f7", "text": "#0b0b0b",
        "text_muted": "#52514e", "axis": "#898781", "grid": "#e1e0d9",
        "card_bg": "#ffffff", "card_border": "rgba(11,11,11,0.10)",
        "blue": "#2a78d6", "orange": "#eb6834", "aqua": "#1baf7a",
        "yellow": "#eda100", "magenta": "#e87ba4", "good": "#0ca30c", "critical": "#d03b3b",
    },
    "dark": {
        "surface": "#1a1a19", "page": "#0d0d0d", "text": "#ffffff",
        "text_muted": "#c3c2b7", "axis": "#898781", "grid": "#2c2c2a",
        "card_bg": "#232322", "card_border": "rgba(255,255,255,0.10)",
        "blue": "#3987e5", "orange": "#d95926", "aqua": "#199e70",
        "yellow": "#c98500", "magenta": "#d55181", "good": "#0ca30c", "critical": "#e66767",
    },
}

CAT_COLOR = {"Banca": "blue", "Cripto": "orange", "Renta Variable": "aqua", "Otros activos": "yellow"}

# Cartera y riesgo cripto: color fijo por categoria (nunca se reasigna segun
# filtros), en el orden que ya usa data_loader (CARTERA_ORDER / RIESGO_ORDER).
CARTERA_COLOR_KEYS = {"Cartera Old": "blue", "Cartera New": "orange", "RD26": "aqua", "DEFI": "yellow", "HYPE": "magenta"}
RIESGO_COLOR_KEYS = {"Liquided": "blue", "LP": "aqua", "Medium": "yellow", "Gem": "orange", "Crash": "critical"}
CATEGORIA_ORDER = ["LIQUIDED", "FONDOS", "MP", "ETF", "ACC"]
CATEGORIA_COLOR_KEYS = {"LIQUIDED": "blue", "FONDOS": "orange", "MP": "aqua", "ETF": "yellow", "ACC": "magenta"}

st.set_page_config(page_title="Patrimonio · Posiciones", page_icon="\U0001F4B0", layout="wide")

# --------------------------------------------------------------------------
# Estado / helpers
# --------------------------------------------------------------------------
if "theme" not in st.session_state:
    st.session_state.theme = "light"
if "data" not in st.session_state:
    st.session_state.data = None
    st.session_state.loaded_at = None


def fmt_eur(value, decimals=0):
    if value is None or pd.isna(value):
        return "—"
    sign = "-" if value < 0 else ""
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
    st.session_state.theme = st.radio("Tema", ["light", "dark"], index=0 if st.session_state.theme == "light" else 1,
                                       format_func=lambda x: "Claro" if x == "light" else "Oscuro", horizontal=True)
    refresh_clicked = st.button("🔄 Actualizar datos desde finance.xlsx", use_container_width=True)

# COLORS se resuelve DESPUES del radio de arriba para que la CSS de este
# mismo rerun ya refleje el tema recien elegido.
COLORS = PALETTE[st.session_state.theme]
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


def make_donut(labels, values, color_keys, height=300, center_total=None):
    """Donut con colores fijos por etiqueta (color_keys: dict label -> clave de PALETTE)."""
    colors = [c[color_keys.get(lb, "text_muted")] for lb in labels]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.58,
        marker=dict(colors=colors, line=dict(color=c["surface"], width=2)),
        textinfo="percent", textfont=dict(color="#ffffff", size=12), sort=False,
    ))
    if center_total is not None:
        fig.add_annotation(text=f"<b>{center_total}</b>", showarrow=False, font=dict(size=15, color=c["text"]))
    apply_layout(fig, height=height, showlegend=True)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {c['page']}; }}
    .kpi-card {{ background: {c['card_bg']}; border: 1px solid {c['card_border']};
                 border-radius: 12px; padding: 16px 20px; height: 100%; }}
    .kpi-label {{ color: {c['text_muted']}; font-size: 0.82rem; text-transform: uppercase; letter-spacing: .04em; }}
    .kpi-value {{ color: {c['text']}; font-size: 1.35rem; font-weight: 700; margin-top: 2px; white-space: nowrap; }}
    .kpi-sub {{ color: {c['text_muted']}; font-size: 0.8rem; margin-top: 2px; }}
    .hero-total {{ color: {c['text']}; font-size: 2.6rem; font-weight: 800; }}
    .hero-label {{ color: {c['text_muted']}; font-size: 0.95rem; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def kpi_card(label, value_str, sub=None, value_color=None):
    style = f' style="color:{value_color}"' if value_color else ""
    st.markdown(
        f"""<div class="kpi-card">
                <div class="kpi-label">{label}</div>
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
    st.caption("Snapshot del patrimonio en el instante actual (sin evolucion temporal).")

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
# PATRIMONIO TOTAL
# ==========================================================================
st.markdown('<div class="hero-label">PATRIMONIO TOTAL ACTUAL</div>', unsafe_allow_html=True)
st.markdown(f'<div class="hero-total">{fmt_eur(pat["total"], 2)}</div>', unsafe_allow_html=True)
st.write("")

col1, col2, col3, col4 = st.columns(4)
with col1:
    kpi_card("🏦 Bancos", fmt_eur(pat["banca"]), fmt_pct(pat["banca"] / pat["total"] * 100) + " del total")
with col2:
    kpi_card("💼 Activos", fmt_eur(pat["activos_otros"]), fmt_pct(pat["activos_otros"] / pat["total"] * 100) + " del total")
with col3:
    kpi_card("📈 Renta variable", fmt_eur(pat["renta_variable"]), fmt_pct(pat["renta_variable"] / pat["total"] * 100) + " del total")
with col4:
    kpi_card("🪙 Cripto", fmt_eur(pat["cripto"]), fmt_pct(pat["cripto"] / pat["total"] * 100) + " del total")

st.write("")
col_donut, _ = st.columns([1, 2])
with col_donut:
    cats = ["Banca", "Cripto", "Renta Variable", "Otros activos"]
    vals = [pat["banca"], pat["cripto"], pat["renta_variable"], pat["activos_otros"]]
    colors = [c[CAT_COLOR[k]] for k in cats]
    fig = go.Figure(go.Pie(labels=cats, values=vals, hole=0.62,
                            marker=dict(colors=colors, line=dict(color=c["surface"], width=2)),
                            textinfo="percent", textfont=dict(color="#ffffff", size=12), sort=False))
    fig.add_annotation(text=f"<b>{fmt_eur(pat['total'], 0)}</b>", showarrow=False, font=dict(size=16, color=c["text"]))
    apply_layout(fig, height=280, showlegend=True)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

st.divider()

# ==========================================================================
# POSICIONES POR CATEGORIA (grafico, no tablas)
# ==========================================================================
st.markdown(f'<div class="hero-label" style="margin-bottom:8px;">TUS POSICIONES</div>', unsafe_allow_html=True)

tab_bancos, tab_activos, tab_rv, tab_cripto = st.tabs(
    ["🏦 Bancos", "💼 Activos", "📈 Renta variable", "🪙 Cripto"]
)


def _dedupe_labels(labels):
    """Si una etiqueta se repite, un eje categorico de Plotly fusiona esas
    filas en una sola barra. Anadimos un sufijo (2), (3)... a partir de la
    segunda aparicion para que cada posicion se vea por separado.
    """
    seen = {}
    out = []
    for lb in labels:
        seen[lb] = seen.get(lb, 0) + 1
        out.append(lb if seen[lb] == 1 else f"{lb} ({seen[lb]})")
    return out


def position_bar(df, label_col, value_col, color, height, value_fmt=fmt_eur):
    """Barra horizontal simple: una posicion por fila, ordenada por valor."""
    df = df.sort_values(value_col, ascending=True).copy()
    df[label_col] = _dedupe_labels(df[label_col].tolist())
    bar_colors = [c["critical"] if v < 0 else color for v in df[value_col]]
    fig = go.Figure(go.Bar(
        x=df[value_col], y=df[label_col], orientation="h",
        marker=dict(color=bar_colors),
        text=[value_fmt(v) for v in df[value_col]], textposition="outside",
        cliponaxis=False,
    ))
    lo, hi = df[value_col].min(), df[value_col].max()
    pad = max(abs(lo), abs(hi)) * 0.28 or 1
    fig.update_xaxes(range=[min(0, lo) - pad, hi + pad])
    apply_layout(fig, height=height)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


with tab_bancos:
    df_b = pd.DataFrame(data["bancos"])
    df_b["etiqueta"] = df_b["proveedor"] + " · " + df_b["detalle"]
    df_b.loc[df_b["proveedor"] == df_b["detalle"], "etiqueta"] = df_b["proveedor"]
    deuda_total = df_b.loc[df_b["importe"] < 0, "importe"].sum()

    if deuda_total < 0:
        st.caption(
            f"{len(df_b)} posiciones · Total bancario: {fmt_eur(pat['banca'], 2)} "
            f"(ya incluye la deuda de abajo)"
        )
        kd1, _ = st.columns(2)
        with kd1:
            kpi_card("Deuda / crédito", fmt_eur(deuda_total, 2), "no es patrimonio: es dinero prestado, pendiente de devolver", value_color=c["critical"])
        st.write("")
    else:
        st.caption(f"{len(df_b)} posiciones · Total: {fmt_eur(pat['banca'], 2)}")

    position_bar(df_b, "etiqueta", "importe", c["blue"], height=max(280, 32 * len(df_b)))
    st.caption(
        "Todas las cuentas se tratan igual salvo la que aparece en rojo (si la hay): "
        "un saldo negativo es una deuda/crédito, no un activo — resta del total, no suma."
    )

with tab_activos:
    df_a = pd.DataFrame(data["activos_otros"])
    st.caption(f"{len(df_a)} posiciones · Total: {fmt_eur(pat['activos_otros'], 2)}")
    position_bar(df_a, "detalle", "importe", c["yellow"], height=max(220, 60 * len(df_a)))

with tab_rv:
    df_rv_raw = pd.DataFrame(data["rv_posiciones"])
    rv_resumen = data["rv_resumen"]
    invertido_total = rv_resumen["invertido"]
    rendimiento_total = rv_resumen["rendimiento_pct"]

    k1, k2, k3 = st.columns(3)
    with k1:
        kpi_card("Capital invertido", fmt_eur(invertido_total, 2), "total oficial (hoja GUI); el desglose de abajo usa el coste por posicion")
    with k2:
        kpi_card("Valor actual", fmt_eur(pat["renta_variable"], 2))
    with k3:
        gain_color = c["good"] if (rendimiento_total or 0) >= 0 else c["critical"]
        kpi_card("Rendimiento", fmt_pct(rendimiento_total, 2, signed=True), "sobre el capital invertido", value_color=gain_color)

    st.write("")
    col_donut_rv, col_gap_rv = st.columns([1, 2])
    with col_donut_rv:
        st.caption("Por categoria (tipo de producto)")
        df_rvcat = pd.DataFrame(data["rv_categoria"])
        make_donut(df_rvcat["categoria"], df_rvcat["importe"], CATEGORIA_COLOR_KEYS,
                   height=280, center_total=fmt_eur(pat["renta_variable"], 0))

    st.write("")
    st.caption("Cada posicion, agrupada y coloreada por categoria — el texto muestra valor y % de rendimiento")
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

    bar_colors = [c[CATEGORIA_COLOR_KEYS.get(cat, "text_muted")] for cat in df_rv["categoria"]]
    text_colors = [
        c["good"] if pd.notna(r) and r >= 0 else (c["critical"] if pd.notna(r) else c["text_muted"])
        for r in df_rv["rendimiento_pct"]
    ]
    text_labels = [
        f"{fmt_eur(v, 0)}  ({fmt_pct(r, 1, signed=True)})" if pd.notna(r) and pd.notna(df_rv['invertido'].iloc[i]) else fmt_eur(v, 0)
        for i, (v, r) in enumerate(zip(df_rv["valor_actual"], df_rv["rendimiento_pct"]))
    ]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_rv["valor_actual"], y=df_rv["etiqueta"], orientation="h",
        marker=dict(color=bar_colors), text=text_labels, textposition="outside", cliponaxis=False,
        textfont=dict(color=text_colors), name="Valor actual",
    ))
    fig.add_trace(go.Scatter(
        x=df_rv["invertido"], y=df_rv["etiqueta"], mode="markers",
        marker=dict(symbol="line-ns", size=16, line=dict(color=c["text_muted"], width=2)),
        name="Capital invertido",
    ))
    fig.update_xaxes(range=[0, df_rv["valor_actual"].max() * 1.35])
    apply_layout(fig, height=max(420, 30 * len(df_rv)), showlegend=True)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    st.caption(
        "El color de la barra indica la categoria (ver donut de arriba); el texto en "
        "verde/rojo es el % de rendimiento sobre lo invertido (marca gris). 'TRADE REP' "
        "y 'MYINV' son efectivo aun sin invertir en el broker, por eso no tienen "
        "invertido/rendimiento. "
        f"Nota: la suma del coste de entrada de estas posiciones da "
        f"{fmt_eur(df_rv['invertido'].sum(skipna=True), 2)}, algo menos que el "
        f"'Capital invertido' oficial de arriba ({fmt_eur(invertido_total, 2)}) — la "
        "diferencia vendra de movimientos que la hoja GUI arrastra y que no estan "
        "en la tabla de posiciones abiertas (p.ej. posiciones ya vendidas)."
    )

with tab_cripto:
    k1, k2 = st.columns(2)
    with k1:
        kpi_card("Valor actual", fmt_eur(pat["cripto"], 2))
    with k2:
        kpi_card("Capital invertido", "—", "el Excel no lo registra de forma fiable por posicion (ver nota en 'Posiciones')")

    st.write("")
    sub_cartera, sub_riesgo, sub_cruce, sub_posiciones = st.tabs(
        ["Por cartera", "Por riesgo", "Cruce cartera × riesgo", "Todas las posiciones"]
    )

    with sub_cartera:
        df_cart = pd.DataFrame(data["cripto_cartera"])
        st.caption("Valor actual por cartera / estrategia (5 carteras independientes)")
        make_donut(df_cart["cartera"], df_cart["actual"], CARTERA_COLOR_KEYS,
                   height=340, center_total=fmt_eur(pat["cripto"], 0))

    with sub_riesgo:
        df_r = pd.DataFrame(data["cripto_riesgo"])
        st.caption("Valor actual por nivel de riesgo/calidad")
        make_donut(df_r["riesgo"], df_r["actual"], RIESGO_COLOR_KEYS,
                   height=340, center_total=fmt_eur(pat["cripto"], 0))

    with sub_cruce:
        st.caption("Cada barra es una cartera; los tramos de color son su reparto por nivel de riesgo")
        df_cross = pd.DataFrame(data["cripto_cross"])
        totales_cartera = {x["cartera"]: x["actual"] for x in data["cripto_cartera"]}
        carteras_presentes = [x for x in CARTERA_ORDER if x in totales_cartera]

        fig = go.Figure()
        for riesgo in RIESGO_ORDER:
            sub = df_cross[df_cross["riesgo"] == riesgo].set_index("cartera")["valor"]
            xs = [sub.get(cart, 0) for cart in carteras_presentes]
            fig.add_trace(go.Bar(
                x=xs, y=carteras_presentes, orientation="h", name=riesgo,
                marker=dict(color=c[RIESGO_COLOR_KEYS[riesgo]]),
            ))
        max_total = max(totales_cartera.values())
        for cart in carteras_presentes:
            fig.add_annotation(
                x=totales_cartera[cart] + max_total * 0.02, y=cart, xanchor="left",
                text=f"<b>{fmt_eur(totales_cartera[cart], 0)}</b>", showarrow=False,
                font=dict(size=12, color=c["text"]),
            )
        fig.update_layout(barmode="stack")
        fig.update_xaxes(range=[0, max_total * 1.22])
        fig.update_yaxes(autorange="reversed")
        apply_layout(fig, height=320, showlegend=True)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        st.caption(
            "Leido de la hoja GUI (cruce cartera × riesgo) y convertido a EUR con el "
            "mismo tipo de cambio que el resto de cripto; reconcilia practicamente al "
            "euro con el desglose 'Por cartera' y 'Por riesgo' de al lado."
        )

    with sub_posiciones:
        show_dust = st.checkbox("Incluir posiciones residuales (< 1 €)", value=False, key="dust_cripto")
        df_cr = pd.DataFrame(data["cripto_posiciones"])
        if not show_dust:
            df_cr = df_cr[df_cr["capital"] >= 1]
        df_cr = df_cr.sort_values("capital", ascending=True).reset_index(drop=True)
        position_bar(
            df_cr.rename(columns={"coin": "Activo", "capital": "Valor"})[["Activo", "Valor"]],
            "Activo", "Valor", c["orange"], height=max(420, 22 * len(df_cr)),
        )
        st.caption(
            f"{len(df_cr)} posiciones con valor mostradas. El Excel no guarda el capital "
            "invertido por posicion en cripto (solo un total agregado que ademas no cuadra "
            "de forma fiable), asi que de momento solo se muestra el valor actual — lo "
            "resolveremos cuando definamos como registrar el coste de entrada por posicion."
        )
