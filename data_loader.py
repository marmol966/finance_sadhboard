# -*- coding: utf-8 -*-
"""
Lectura de finance.xlsx (SOLO LECTURA) y transformacion a estructuras
listas para pintar en el dashboard.

finance.xlsx vive FUERA de esta carpeta (G:\\Mi unidad\\Marmol\\3-Finances\\finance.xlsx).
Este modulo nunca escribe en ese archivo ni en ningun otro fuera de
"Finance Dashboard App".

Nota de rendimiento: las hojas se leen con openpyxl en modo read_only, que
esta optimizado para iteracion SECUENCIAL (iter_rows). Acceder a celdas
sueltas por coordenada (ws["H12"]) en una hoja read_only obliga a openpyxl a
re-recorrer el XML desde el principio cada vez que "saltamos hacia atras",
lo que en una hoja como Portfolio (varias tablas independientes, saltos de
fila constantes) puede tardar minutos. Por eso aqui cada hoja se vuelca UNA
sola vez a una lista en memoria (`_Grid`) y toda la logica de parseo trabaja
sobre esa copia en memoria con acceso O(1).
"""

import os
import re

import openpyxl
from openpyxl.utils import column_index_from_string

# finance.xlsx esta un nivel por encima de "Finance Dashboard App"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.normpath(os.path.join(APP_DIR, "..", "finance.xlsx"))

CARTERA_ORDER = ["Cartera Old", "Cartera New", "RD26", "DEFI", "HYPE"]
RIESGO_ORDER = ["Liquided", "LP", "Medium", "Gem", "Crash"]

_COORD_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


class _Grid:
    """Copia en memoria (lista de listas) de una hoja, con acceso tipo A1."""

    def __init__(self, rows):
        self.rows = rows  # tuple de tuplas, 0-indexado

    def get(self, coord):
        m = _COORD_RE.match(coord)
        col_letters, row_str = m.group(1), m.group(2)
        row_idx = int(row_str) - 1
        col_idx = column_index_from_string(col_letters) - 1
        if row_idx >= len(self.rows):
            return None
        row = self.rows[row_idx]
        if col_idx >= len(row):
            return None
        return row[col_idx]


def _norm_riesgo(valor):
    if valor is None:
        return None
    v = str(valor).strip()
    if v.upper() == "GEM":
        return "Gem"
    return v


def load_workbook():
    """Abre finance.xlsx en modo solo-lectura (nunca se guarda)."""
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"No se encuentra finance.xlsx en: {EXCEL_PATH}")
    return openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)


def _read_grid(wb, sheet_name, max_row, max_col):
    ws = wb[sheet_name]
    rows = tuple(ws.iter_rows(max_row=max_row, max_col=max_col, values_only=True))
    return _Grid(rows)


def load_data():
    wb = load_workbook()
    # Una sola pasada secuencial por cada hoja que necesitamos (rapido); el
    # resto de la funcion trabaja sobre estas copias en memoria.
    g_cashflow = _read_grid(wb, "CashFlow", max_row=20, max_col=12)
    g_portfolio = _read_grid(wb, "Portfolio", max_row=165, max_col=27)
    g_gui = _read_grid(wb, "GUI", max_row=95, max_col=11)
    wb.close()

    # Portfolio!D3 = tipo de cambio USD->EUR (los precios/capitales de las
    # posiciones cripto individuales, columnas O/Q, estan en USD; el resto
    # de la hoja -bancos, RV, totales de CashFlow- ya esta en EUR).
    fx_usd_eur = g_portfolio.get("D3") or 1.0

    data = {}
    data["fx_usd_eur"] = fx_usd_eur
    data["patrimonio"] = _load_patrimonio(g_cashflow)
    data["bancos"] = _load_bancos(g_portfolio)
    data["activos_otros"] = _load_activos_otros(g_portfolio)
    data["cripto_cartera"] = _load_cripto_cartera(g_portfolio)
    data["cripto_posiciones"] = _load_cripto_posiciones(g_portfolio, fx_usd_eur)
    data["cripto_riesgo"] = _aggregate_riesgo(data["cripto_posiciones"])
    data["cripto_cross"] = _load_cripto_cross(g_gui, fx_usd_eur)
    data["rv_posiciones"] = _load_rv_posiciones(g_portfolio)
    data["rv_categoria"] = _load_rv_categoria(g_portfolio)
    data["rv_resumen"] = _load_rv_resumen(g_gui)

    total = data["patrimonio"]["total"]
    total_cripto = data["patrimonio"]["cripto"]
    total_rv = data["patrimonio"]["renta_variable"]

    for pos in data["cripto_posiciones"]:
        pos["pct_categoria"] = (pos["capital"] / total_cripto * 100) if total_cripto else 0
        pos["pct_patrimonio"] = (pos["capital"] / total * 100) if total else 0

    for pos in data["rv_posiciones"]:
        pos["pct_categoria"] = (pos["valor_actual"] / total_rv * 100) if total_rv else 0
        pos["pct_patrimonio"] = (pos["valor_actual"] / total * 100) if total else 0

    data["warnings"] = _validate(data)
    return data


def _validate(data):
    """Comprueba que las distintas fuentes del Excel reconcilian entre si.

    finance.xlsx vive en una unidad de Google Drive que el usuario edita en
    vivo; una lectura a medio-sincronizar puede devolver una hoja vacia o
    incompleta sin que openpyxl lance ningun error (simplemente hay menos
    filas). Estas comprobaciones existen para detectar esa situacion -no
    para validar reglas de negocio- asi que el umbral es deliberadamente
    laxo (1 EUR) y solo compara cosas que YA deberian cuadrar al centimo.
    """
    warnings = []
    pat = data["patrimonio"]
    tol = 1.0

    suma_categorias = pat["banca"] + pat["activos_otros"] + pat["renta_variable"] + pat["cripto"]
    if abs(suma_categorias - pat["total"]) > tol:
        warnings.append(
            f"CashFlow: BANCA+ACTIVOS+RV+Cripto ({suma_categorias:,.2f}) no coincide "
            f"con PATRIMONIO ({pat['total']:,.2f})."
        )

    suma_bancos = sum(b["importe"] for b in data["bancos"])
    if abs(suma_bancos - pat["banca"]) > tol:
        warnings.append(
            f"Bancos: la suma de posiciones ({suma_bancos:,.2f}) no coincide con "
            f"CashFlow!BANCA ({pat['banca']:,.2f}) — posible lectura incompleta de Portfolio."
        )

    suma_activos = sum(a["importe"] for a in data["activos_otros"])
    if abs(suma_activos - pat["activos_otros"]) > tol:
        warnings.append(
            f"Activos: la suma de posiciones ({suma_activos:,.2f}) no coincide con "
            f"CashFlow!ACTIVOS ({pat['activos_otros']:,.2f})."
        )

    suma_cripto_cartera = sum(x["actual"] for x in data["cripto_cartera"])
    suma_cripto_pos = sum(p["capital"] for p in data["cripto_posiciones"])
    if abs(suma_cripto_cartera - pat["cripto"]) > tol or abs(suma_cripto_pos - pat["cripto"]) > tol:
        warnings.append(
            f"Cripto: cartera={suma_cripto_cartera:,.2f}, posiciones={suma_cripto_pos:,.2f}, "
            f"CashFlow!Cripto={pat['cripto']:,.2f} — no cuadran, probable lectura parcial "
            "del Excel (reintenta actualizar)."
        )
    if pat["cripto"] < tol and suma_cripto_pos > tol:
        warnings.append(
            "Cripto muestra 0€ en el resumen pero hay posiciones con valor: la hoja "
            "CashFlow probablemente se leyo a medias. Pulsa 'Actualizar datos'."
        )

    suma_rv = sum(p["valor_actual"] for p in data["rv_posiciones"])
    suma_rv_cat = sum(x["importe"] for x in data["rv_categoria"])
    if abs(suma_rv - pat["renta_variable"]) > tol or abs(suma_rv_cat - pat["renta_variable"]) > tol:
        warnings.append(
            f"Renta variable: posiciones={suma_rv:,.2f}, por categoria={suma_rv_cat:,.2f}, "
            f"CashFlow!Renta Variable={pat['renta_variable']:,.2f} — no cuadran."
        )
    categorias_desconocidas = sorted({
        p["categoria"] for p in data["rv_posiciones"]
        if p["categoria"] not in ("LIQUIDED", "FONDOS", "MP", "ETF", "ACC")
    })
    if categorias_desconocidas:
        warnings.append(
            f"Renta variable: hay posiciones con categoria no reconocida: {categorias_desconocidas} "
            "(la app solo sabe colorear LIQUIDED/FONDOS/MP/ETF/ACC)."
        )

    rv_resumen = data["rv_resumen"]
    if rv_resumen["actual"] is None or rv_resumen["invertido"] is None:
        warnings.append(
            "Renta variable: no se pudo leer GUI!RV / GUI!RV_Inv (capital invertido total)."
        )
    elif abs(rv_resumen["actual"] - pat["renta_variable"]) > tol:
        warnings.append(
            f"Renta variable: GUI!RV ({rv_resumen['actual']:,.2f}) no coincide con "
            f"CashFlow!Renta Variable ({pat['renta_variable']:,.2f})."
        )

    return warnings


def _load_patrimonio(g):
    """CashFlow!C:D -> resumen de patrimonio, buscado POR ETIQUETA (columna C)
    en vez de por numero de fila fijo. Las filas de este resumen se han
    desplazado alguna vez al anadir/quitar una sub-linea (p.ej. una divisa
    bajo ACTIVOS), asi que anclar por texto es mas robusto que un rango fijo.
    """
    labels = {}
    for r in range(1, 15):
        label = g.get(f"C{r}")
        if label is None:
            continue
        value = g.get(f"D{r}")
        labels[str(label).strip().upper()] = value

    def pick(*keys):
        for k in keys:
            v = labels.get(k)
            if isinstance(v, (int, float)):
                return v
        return 0

    return {
        "total": pick("PATRIMONIO"),
        "banca": pick("BANCA"),
        "activos_otros": pick("ACTIVOS", "ACTIVO"),
        "inversiones": pick("INVERSIONES"),
        "renta_variable": pick("RENTA VARIABLE"),
        "cripto": pick("CRIPTO"),
    }


def _load_bancos(g):
    """Tabla BANCOS: Portfolio!C9:E~18 (Proveedor / Descripcion / Cantidad).

    Se recorre un rango amplio y se descartan filas sin Descripcion+Cantidad
    simultaneas, para no depender de un numero de filas fijo si el usuario
    anade/quita bancos en el futuro.
    """
    filas = []
    for r in range(9, 23):
        proveedor = g.get(f"C{r}")
        desc = g.get(f"D{r}")
        cantidad = g.get(f"E{r}")
        if desc is None or not isinstance(cantidad, (int, float)):
            continue
        filas.append({
            "proveedor": proveedor if proveedor else desc,
            "detalle": desc,
            "importe": cantidad,
        })
    return filas


def _load_activos_otros(g):
    """Tabla ACTIVOS (moto, divisas...): Portfolio!C24:E~28."""
    filas = []
    for r in range(24, 29):
        tipo = g.get(f"C{r}")
        desc = g.get(f"D{r}")
        cantidad = g.get(f"E{r}")
        if desc is None or not isinstance(cantidad, (int, float)):
            continue
        filas.append({"tipo": tipo, "detalle": desc, "importe": cantidad})
    return filas


def _load_cripto_cartera(g):
    """Portfolio!H12:K16 -> valor actual EN EUR por cartera/estrategia.

    Esta tabla ya viene en EUR en el propio Excel y su suma reconcilia
    exactamente con el total de cripto de CashFlow. No se usa GUI!D8:D12
    porque esa columna resulto estar en USD sin convertir (ver README:
    limitaciones conocidas).
    """
    filas = []
    for r in range(9, 20):
        nombre = g.get(f"H{r}")
        total_eur = g.get(f"K{r}")
        if not nombre or not isinstance(total_eur, (int, float)):
            continue
        filas.append({"cartera": nombre, "actual": total_eur})
    filas.sort(key=lambda x: CARTERA_ORDER.index(x["cartera"]) if x["cartera"] in CARTERA_ORDER else 99)
    return filas


def _aggregate_riesgo(posiciones):
    """Agrega las posiciones cripto (ya en EUR) por nivel de riesgo."""
    agg = {}
    for p in posiciones:
        r = p["riesgo"]
        entry = agg.setdefault(r, {"riesgo": r, "actual": 0.0, "n_posiciones": 0})
        entry["actual"] += p["capital"]
        if p["capital"] > 0.01:
            entry["n_posiciones"] += 1
    filas = list(agg.values())
    filas.sort(key=lambda x: RIESGO_ORDER.index(x["riesgo"]) if x["riesgo"] in RIESGO_ORDER else 99)
    return filas


def _load_cripto_cross(g, fx_usd_eur):
    """GUI!D15:H19 -> cruce cartera x nivel de riesgo.

    Estos valores estan en USD sin convertir en la propia hoja GUI (igual
    que GUI!D8:D12); se multiplican aqui por el tipo de cambio para que
    reconcilien con el total oficial de cripto en EUR.
    """
    carteras_cols = {"D": "Cartera Old", "E": "Cartera New", "F": "RD26", "G": "DEFI", "H": "HYPE"}
    filas = []
    for r in range(15, 20):
        riesgo = _norm_riesgo(g.get(f"C{r}"))
        if not riesgo:
            continue
        for col, cartera in carteras_cols.items():
            valor = g.get(f"{col}{r}")
            if isinstance(valor, (int, float)) and valor > 0:
                filas.append({"riesgo": riesgo, "cartera": cartera, "valor": valor * fx_usd_eur})
    return filas


def _load_cripto_posiciones(g, fx_usd_eur):
    """Portfolio!M9:Q164 -> una fila por posicion cripto individual.

    O (Capital) y Q (Precio) estan en USD en la hoja; se convierten aqui a
    EUR con el tipo de cambio de Portfolio!D3 para que sumen exactamente el
    total de cripto de CashFlow/GUI (en EUR).
    """
    filas = []
    for r in range(9, 165):
        riesgo = _norm_riesgo(g.get(f"M{r}"))
        coin = g.get(f"N{r}")
        if riesgo is None or coin is None:
            continue
        capital_usd = g.get(f"O{r}") or 0
        cantidad = g.get(f"P{r}")
        precio_usd = g.get(f"Q{r}")
        filas.append({
            "coin": coin,
            "riesgo": riesgo,
            "capital": capital_usd * fx_usd_eur,
            "cantidad": cantidad,
            "precio": precio_usd * fx_usd_eur if isinstance(precio_usd, (int, float)) else None,
        })
    return filas


def _load_rv_posiciones(g):
    """Portfolio!A35:K58 -> una fila por posicion de renta variable / MP / liquidez broker."""
    filas = []
    for r in range(35, 60):
        campo = g.get(f"H{r}")
        activo = g.get(f"I{r}")
        total = g.get(f"J{r}")
        if campo is None or activo is None or total is None:
            continue
        k_val = g.get(f"K{r}")
        filas.append({
            "categoria": g.get(f"A{r}"),
            "campo": campo,
            "activo": activo,
            "invertido": g.get(f"F{r}"),
            "valor_actual": total,
            "rendimiento_pct": k_val * 100 if isinstance(k_val, (int, float)) else None,
        })
    return filas


def aggregate_bancos(bancos):
    """Agrupa las filas de bancos por 'proveedor' (columna C de Portfolio,
    o la descripcion cuando C viene vacia) para las graficas de composicion.
    """
    agg = {}
    for b in bancos:
        key = b["proveedor"]
        agg[key] = agg.get(key, 0) + b["importe"]
    return [{"proveedor": k, "importe": v} for k, v in agg.items()]


def _load_rv_resumen(g):
    """GUI!C91:D92 -> capital invertido TOTAL oficial de renta variable.

    OJO: NO se usa la suma de 'invertido' por posicion (Portfolio!F35:F58)
    como total, porque esa suma (~6.512€) NO coincide con esta celda
    dedicada 'RV_Inv' (~7.928€) que el propio Excel usa para calcular su
    rendimiento global de RV (GUI!F91 = 1,95%, coincide con estos dos
    numeros). La celda 'RV_Inv' es la que el usuario trata como oficial,
    asi que es la que mostramos como KPI; el detalle por posicion sigue
    usando el invertido individual de Portfolio, que es mas fino pero no
    suma exactamente igual (puede haber posiciones ya vendidas u otros
    ajustes que 'RV_Inv' arrastra y que no estan en la tabla de posiciones
    actuales).
    """
    actual = None
    invertido = None
    for r in range(1, 93):
        label = g.get(f"C{r}")
        if label is None:
            continue
        label_up = str(label).strip().upper()
        if label_up == "RV":
            actual = g.get(f"D{r}")
        elif label_up == "RV_INV":
            invertido = g.get(f"D{r}")
    rendimiento_pct = None
    if isinstance(actual, (int, float)) and isinstance(invertido, (int, float)) and invertido:
        rendimiento_pct = (actual - invertido) / invertido * 100
    return {"actual": actual, "invertido": invertido, "rendimiento_pct": rendimiento_pct}


def _load_rv_categoria(g):
    """Portfolio!C60:E64 -> desglose RV por tipo de producto (Liquidez/Fondos/MP/ETF/Acciones)."""
    filas = []
    for r in range(60, 65):
        cat = g.get(f"C{r}")
        importe = g.get(f"D{r}")
        if cat is None or importe is None:
            continue
        filas.append({"categoria": cat, "importe": importe})
    return filas


def load_data_checked(max_attempts=3, delay_seconds=0.5):
    """Como load_data(), pero reintenta si detecta que las cifras no
    reconcilian (probable lectura a medio-sincronizar de finance.xlsx en
    Google Drive). Devuelve el mejor intento (el ultimo) con su lista de
    'warnings' incluida, aunque siga sin cuadrar tras los reintentos.
    """
    import time
    result = None
    for attempt in range(max_attempts):
        result = load_data()
        if not result["warnings"]:
            return result
        if attempt < max_attempts - 1:
            time.sleep(delay_seconds)
    return result


if __name__ == "__main__":
    import json
    import time
    t0 = time.time()
    d = load_data_checked()
    print(f"tiempo de carga: {time.time() - t0:.2f}s")
    warnings = d.pop("warnings")
    print(json.dumps({k: (v if not isinstance(v, list) else f"{len(v)} filas") for k, v in d.items()}, indent=2, ensure_ascii=False, default=str))
    print()
    print("warnings:", warnings if warnings else "(ninguno, todo reconcilia)")
