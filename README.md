# Dashboard de Patrimonio (v1 — snapshot por categoría)

Panel local (Streamlit) que lee `finance.xlsx` y muestra una **foto actual**
de tus posiciones, organizadas en 4 categorías: **Bancos, Activos, Renta
Variable y Cripto**. Es deliberadamente gráfico (barras, no tablas) y sin
clasificación adicional dentro de cada categoría todavía — eso se irá
detallando en próximas iteraciones. No incluye evolución temporal.

## Cómo ejecutarlo

Desde esta carpeta (`Finance Dashboard App`):

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abrirá en el navegador (normalmente `http://localhost:8501`). Todo corre
en local; no se sube nada a ningún servicio externo.

## Cómo refrescar los datos

El dashboard lee `finance.xlsx` al arrancar y guarda el resultado en caché de
sesión. Pulsa **"🔄 Actualizar datos desde finance.xlsx"** en la barra
lateral para releer el archivo. Si lo tienes abierto en Excel, guarda los
cambios antes de refrescar.

## Qué muestra

- **Cabecera**: patrimonio total y su reparto entre las 4 categorías (donut).
- **Pestaña Bancos**: cada cuenta/proveedor como una barra (importe).
- **Pestaña Activos**: el resto de activos (moto, divisas en efectivo) como barras.
- **Pestaña Renta variable**: capital invertido, valor actual y rendimiento %
  totales, y un gráfico por posición — barra verde/roja según gane o pierda,
  con una marca del capital invertido y el % de rendimiento en la etiqueta.
- **Pestaña Cripto**: valor actual por posición (barras). El Excel no
  registra el capital invertido por posición de forma fiable, así que aquí
  todavía no se muestra rendimiento (ver limitaciones más abajo).

## Origen de los datos

Lee **en modo solo lectura** `G:\Mi unidad\Marmol\3-Finances\finance.xlsx`
(un nivel por encima de esta carpeta). La app nunca escribe en ese archivo ni
en ningún otro fuera de `Finance Dashboard App`.

Hojas usadas: `CashFlow` (totales) y `Portfolio` (bancos, activos, posiciones
cripto y renta variable).

## Estructura del código

- `data_loader.py` — toda la lectura/parseo del Excel (sin dependencias de
  Streamlit; se puede ejecutar suelto con `python data_loader.py` para
  depurar). Lee cada hoja UNA vez de forma secuencial (importante: el acceso
  aleatorio a celdas en hojas `read_only` de openpyxl puede ser muy lento).
- `app.py` — interfaz Streamlit.
- `legacy/` — intento anterior (Flask + `Estado financiero_v6.1.xlsx`,
  centrado solo en cripto). Se conserva como referencia; no lo usa la app
  nueva.

## Limitaciones conocidas (para la siguiente iteración)

- **Cripto sin capital invertido por posición**: el Excel solo registra un
  invertido agregado para todo el bloque cripto, y ese agregado no cuadra de
  forma fiable con nada más — así que no se muestra rendimiento en cripto
  todavía. Si se empieza a registrar coste de entrada por posición, se puede
  añadir igual que en renta variable.
- **Conversión de divisa en cripto**: las posiciones individuales (hoja
  `Portfolio`, columnas O/Q) están en USD; se convierten a EUR con el tipo de
  cambio de la propia hoja (`Portfolio!D3`) para que cuadren con el total
  oficial de `CashFlow`.
- **Préstamo a "Abuelo" (2.000)**: aparece suelto en la hoja `Portfolio` pero
  no forma parte del total oficial de bancos/patrimonio, así que se deja fuera.
- Sin clasificación todavía por cartera/estrategia, nivel de riesgo (cripto)
  ni campo (renta variable) — se añadirá cuando la definamos en detalle.
- Sin evolución temporal (`Evolucion`, `Movimientos`, `Forecast` fuera de
  esta versión) ni cruce con la watchlist de `ETFs`.
