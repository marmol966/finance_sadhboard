# Dashboard de Patrimonio

Panel local (Streamlit) que lee `finance.xlsx` y muestra tanto una foto
actual de las posiciones (Bancos, Activos, Renta Variable y Cripto) como su
evolución semanal histórica, allí donde el Excel la registra.

## Cómo ejecutarlo

Desde esta carpeta (`Finance Dashboard App`):

```bash
pip install -r requirements.txt
streamlit run app.py
```

o haciendo doble clic en `Iniciar Dashboard.bat`. Se abre en el navegador
(normalmente `http://localhost:8501`); todo corre en local, no se sube nada
a ningún servicio externo.

## Estructura del código

- **`data_loader.py`** — toda la lectura/parseo del Excel. Sin dependencias
  de Streamlit, se puede ejecutar suelto (`python data_loader.py`) para
  depurar: imprime cuánto tarda la carga y un resumen de cada bloque de
  datos. Es la pieza que documenta este README.
- **`app.py`** — interfaz Streamlit: toma el diccionario que devuelve
  `data_loader.load_data_checked()` y lo pinta.
- **`legacy/`** — intento anterior (Flask), se conserva como referencia; no
  lo usa la app actual.

## Cómo funciona el algoritmo

### 1. Lectura del Excel: una pasada secuencial por hoja, no acceso aleatorio

`finance.xlsx` se abre con `openpyxl` en modo `read_only` (más rápido y con
menos memoria para un libro de este tamaño). El problema de `read_only` es
que **acceder a una celda suelta por coordenada** (`ws["H12"]`) obliga a
openpyxl a rebobinar el XML desde el principio cada vez que la siguiente
celda pedida está "más atrás" que la anterior — en una hoja como `Portfolio`
(varias tablas independientes, saltos de fila constantes) esto puede tardar
minutos.

Por eso cada hoja se vuelca **una sola vez**, de forma secuencial
(`iter_rows`), a una lista de tuplas en memoria (clase `_Grid`), y *toda* la
lógica de parseo trabaja después sobre esa copia con acceso O(1) por
coordenada (`_Grid.get("H12")`). `load_data()` hace esto para las 4 hojas
que necesita (`CashFlow`, `Portfolio`, `GUI`, `Evolucion`) al principio y
cierra el workbook; el resto de funciones ya no tocan el Excel.

### 2. Localización dinámica de filas: por contenido, no por número fijo

Dos problemas recurrentes de leer un Excel que el usuario edita a mano:

- **Las tablas crecen** (la hoja `Evolucion` añade una fila nueva cada
  semana) — un rango de filas fijo (`range(3, 317)`) se quedaría corto al
  cabo de unos meses.
- **Las tablas se desplazan** (añadir una sub-línea en `CashFlow` mueve
  todo lo de debajo una fila).

Para lo primero, `_iter_fechas(g, start_row)` itera hacia abajo desde una
fila de arranque leyendo la columna de fecha, y **para cuando encuentra
varias filas en blanco seguidas** (`max_blank=4`) en vez de en un número de
fila fijo — así sigue funcionando aunque la hoja crezca, y no confunde la
tabla semanal densa con el resumen a intervalos crecientes que hay más
abajo en la misma hoja (mismas columnas, otro propósito).

Para lo segundo, `_load_patrimonio` no lee `CashFlow!D5` a pelo: busca la
fila cuya columna C dice literalmente "BANCA", "CRIPTO", etc., y lee el
valor de al lado. Más lento de escribir, pero no se rompe si el usuario
inserta una fila por encima.

### 3. Mapeo de columnas por letra: frágil por diseño, documentado por eso

Bancos/Activos/RV se localizan por rango de filas + columna. Cripto es
distinto: varias tablas de `Evolucion` (carteras, riesgo, uso, categorías
de RV) **no tienen cabecera fiable pegada a cada columna** — se localizaron
a mano comparando qué combinación de columnas reconcilia exactamente con un
total oficial conocido (ver `CARTERA_EVOLUCION_COLS`, `CARTERA_CAPITAL_COL`,
`RIESGO_EVOLUCION_COLS`, etc. en `data_loader.py`, cada una con el
razonamiento de cómo se dedujo en su comentario).

Esto es inherentemente frágil: como el usuario edita `Evolucion` a mano
(añadir una categoría, insertar una columna...), estas letras pueden dejar
de ser válidas de un día para otro **sin que openpyxl lance ningún error**
— simplemente se leería la columna equivocada. La red de seguridad es la
validación del punto 5: si las letras dejan de cuadrar, aparece un warning
en vez de un número silenciosamente incorrecto.

### 4. Conversión de divisa: un tipo puntual para el snapshot, uno por semana para el histórico

Las posiciones cripto individuales están en USD en el Excel; el resto
(bancos, RV, totales oficiales) ya está en EUR. Para el snapshot actual
basta un tipo de cambio puntual (`Portfolio!D3`, `fx_usd_eur`), multiplicado
en el momento de leer cada posición.

Para las series históricas de `Evolucion` esto no vale — el tipo de cambio
de *esa semana* no es el de hoy. `_fx_semanal_por_fila` construye un
diccionario `{fila: tipo_de_cambio}` leyendo `Evolucion!BD` fila a fila,
con **relleno hacia adelante** (la última semana conocida se arrastra) y,
para las semanas anteriores al primer dato de `BD` (esa columna no siempre
existió), **relleno hacia atrás** con el primer valor disponible — para no
mezclar USD y EUR en la misma serie ni introducir un salto de escala
artificial en el punto donde `BD` empieza a existir.

### 5. Reconciliación: detectar una lectura a medias, no validar reglas de negocio

`finance.xlsx` vive en Google Drive y el usuario lo edita en vivo; una
lectura mientras se está sincronizando puede devolver una hoja vacía o
incompleta **sin que openpyxl lance ningún error** (simplemente hay menos
filas de las que debería). `_validate()` compara entre sí varias cifras que
el propio Excel ya calcula por caminos distintos y que *deberían* cuadrar
al céntimo (p.ej. suma de posiciones de banco vs `CashFlow!BANCA`); si no
cuadran dentro de una tolerancia laxa (1€), genera un aviso.

`load_data_checked()` es la función que realmente usa `app.py`: si la
primera lectura sale con avisos, reintenta hasta 3 veces con una pequeña
pausa entre intentos (probablemente el archivo terminó de sincronizarse) y
se queda con el último intento, avisos incluidos.

### 6. La capa Streamlit: cache en sesión + helpers de render reutilizables

`app.py` no vuelve a leer el Excel en cada interacción: `load()` guarda el
resultado en `st.session_state` y solo se refresca al pulsar el botón de la
barra lateral. Sobre ese diccionario de datos, el resto del fichero son
básicamente cuatro helpers de render reutilizados en todas las pestañas:

- `kpi_card(...)` — la tarjeta de cifra + etiqueta que se ve por todas partes.
- `render_table(...)` / `render_cartera_perf_table(...)` — tablas HTML a medida
  (no `st.dataframe`, para controlar formato/color por celda).
- `render_evolucion_apilada(series_dict, order, color_keys, ...)` — el área
  apilada semana a semana; recibe el dict `{clave: [{"fecha":, "valor":}]}`
  que devuelven las funciones `*_evolucion` de `data_loader.py` y pinta
  cualquiera de las 4 evoluciones por categoría (cartera/riesgo/uso de
  cripto, categoría de RV) con la misma lógica.

El color de cada categoría (cartera, nivel de riesgo, tipo de uso...) es
fijo y se define una vez en un diccionario cerca del principio del fichero
(`CARTERA_COLOR_KEYS`, `RIESGO_COLOR_KEYS`, ...) — nunca se reasigna según
el orden o el filtro activo, para que un color siempre signifique lo mismo
en todos los gráficos.

## Limitaciones conocidas

- **Letras de columna frágiles en `Evolucion`**: ver punto 3. Si tras editar
  esa hoja aparecen avisos de reconciliación o un gráfico con valores
  extraños, es la primera sospechosa.
- **Cripto sin capital invertido por posición individual**: el Excel
  registra invertido a nivel de cartera, no por posición suelta, así que la
  tabla de posiciones cripto no muestra rendimiento por posición (sí a
  nivel de cartera).
- **Préstamo a "Abuelo" (Portfolio)**: aparece suelto en la hoja pero no
  forma parte del total oficial de bancos/patrimonio, así que se ignora.
