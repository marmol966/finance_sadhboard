# API de Gestión de Portfolio

Esta API proporciona endpoints para acceder a información detallada sobre diferentes carteras de inversión y sus resúmenes.

## Base URL
```
http://localhost:5000
```

## Clasificación de Endpoints

### 1. Endpoints de Resumen
Estos endpoints proporcionan vistas agregadas y resumidas de los datos:

#### 1.1 Resumen General de Carteras
```
GET /portfolio/resumen-carteras
```
Proporciona una vista general de todas las carteras con sus métricas principales:
- Capital actual
- Capital invertido
- Número de posiciones

**Ejemplo de respuesta:**
```json
{
    "Cripto Old": {
        "Capital Actual": 25000.0,
        "Capital Invertido": 20000.0,
        "N Pos": 5
    },
    "Cripto New": {
        "Capital Actual": 15000.0,
        "Capital Invertido": 12000.0,
        "N Pos": 3
    },
    "DEFI": {
        "Capital Actual": 8000.0,
        "Capital Invertido": 7000.0,
        "N Pos": 4
    },
    "HYPE": {
        "Capital Actual": 5000.0,
        "Capital Invertido": 4000.0,
        "N Pos": 2
    }
}
```

#### 1.2 Resumen por Tipos de Activo
```
GET /portfolio/resumen-tipos
```
Agrupa los activos por tipo y muestra métricas agregadas:
- Total en USD
- Porcentaje del portfolio
- Número de posiciones

**Ejemplo de respuesta:**
```json
{
    "L1": {
        "$ Total": 30000.0,
        "Porcentaje": 45.5,
        "N Pos": 2
    },
    "L2": {
        "$ Total": 15000.0,
        "Porcentaje": 22.7,
        "N Pos": 3
    }
}
```

### 2. Endpoints de Carteras Específicas
Estos endpoints proporcionan información detallada de cada cartera individual:

#### 2.1 Cartera Cripto Old (Activos Establecidos)
```
GET /portfolio/cripto-old
```
Muestra los activos crypto más establecidos y de mayor capitalización.

**Campos específicos:**
- Tipo de activo
- Precio actual
- Unidades (COld_Ud)
- Valor en USD (COld_$)
- Porcentaje en la cartera

#### 2.2 Cartera Cripto New (Activos Emergentes)
```
GET /portfolio/cripto-new
```
Contiene activos crypto más nuevos o en desarrollo.

**Campos específicos:**
- Tipo de activo
- Precio actual
- Unidades (CNew_Ud)
- Valor en USD (CNew_$)
- Porcentaje en la cartera

#### 2.3 Cartera DEFI (Finanzas Descentralizadas)
```
GET /portfolio/defi
```
Activos relacionados con protocolos DeFi.

**Campos específicos:**
- Tipo de activo
- Precio actual
- Unidades (DEFI_Ud)
- Valor en USD (DEFI_$)
- Porcentaje en la cartera

#### 2.4 Cartera HYPE (Activos Especulativos)
```
GET /portfolio/hype
```
Activos con alto potencial especulativo.

**Campos específicos:**
- Tipo de activo
- Precio actual
- Unidades (HYPE_Ud)
- Valor en USD (HYPE_$)
- Porcentaje en la cartera

### 3. Endpoint Completo
```
GET /portfolio/all
```
Este endpoint devuelve todos los datos disponibles, incluyendo:
- Dataset completo
- Información de inversión
- Todas las carteras
- Todos los resúmenes

**Nota:** La respuesta es extensa y contiene todos los datos de los endpoints anteriores combinados.

## Estructura de Datos

### Tipos de Respuesta
1. **Resúmenes Agregados**
   - Formato tabular con métricas calculadas
   - Incluye totales y porcentajes

2. **Datos de Cartera**
   - Detalle a nivel de activo
   - Incluye precios, cantidades y valores
   - Métricas específicas por tipo de cartera

3. **Datos Completos**
   - Conjunto completo de datos
   - Incluye todas las métricas y cálculos

## Notas Técnicas

- Todos los endpoints devuelven datos en formato JSON
- La API implementa un sistema de caché que actualiza los datos cada 5 minutos
- CORS está habilitado para permitir peticiones desde cualquier origen
- Los valores monetarios están en dólares estadounidenses (USD)
- Los porcentajes están calculados sobre el total de cada cartera

## Códigos de Estado

- 200: Petición exitosa
- 404: Endpoint no encontrado
- 500: Error interno del servidor

## Headers de Respuesta

```
Content-Type: application/json
Access-Control-Allow-Origin: *
``` 