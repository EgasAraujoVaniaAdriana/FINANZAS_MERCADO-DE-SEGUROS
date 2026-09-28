# Proyecto de Finanzas I

## Datos de la estudiante

**Nombres y apellidos:** Vania Adriana Egas Araujo  
**Código de matrícula:** 2024200498D

## Tema

**N.º 15 — Determinantes de las primas cedidas, el ingreso formal y la tasa de referencia en las primas netas del mercado asegurador peruano (2020-2025).**

## Periodo de análisis

- FECHA_INICIO: `2020-01`
- FECHA_CORTE: `2025-12`
- Frecuencia: mensual
- Cobertura: 72 meses

## Fuentes de datos

### Banco Central de Reserva del Perú — BCRPData

Extracción mediante API con `codigo/01_extraccion_api.py`.

Series utilizadas:

- `PN31883GM`: ingreso promedio del sector formal total nominal.
- `PD04722MM`: tasa de referencia de la política monetaria.
- `PN38705PM`: IPC de Lima Metropolitana, base diciembre de 2021 = 100.

Endpoints detectados en el script:

- `https://estadisticas.bcrp.gob.pe/estadisticas/series/api`
- `https://estadisticas.bcrp.gob.pe/estadisticas/series/api/{codigo}/json/{inicio}/{fin}`

### Superintendencia de Banca, Seguros y AFP — SBS

Extracción mediante descarga programática con:

`codigo/02_scraping_web.py`

Se utilizan archivos mensuales del cuadro S-401:

- `P009`: primas netas.
- `P010`: primas cedidas.

URL oficial detectada en el script:

- `https://intranet2.sbs.gob.pe/estadistica/financiera`

## Estructura principal del proyecto

- `codigo/01_extraccion_api.py`
- `codigo/02_scraping_web.py`
- `codigo/03_limpieza_datos.py`
- `codigo/04_analisis.py`
- `datos_crudos/`
- `datos_procesados/`
- `salidas/`
- `diccionario_variables.md`
- `requirements.txt`
- `.env.example`
- `log_ejecucion.txt`
- `README.md`

## Orden de ejecución

Desde la carpeta raíz del proyecto ejecutar, en este orden:

1. `python codigo/01_extraccion_api.py`
2. `python codigo/02_scraping_web.py`
3. `python codigo/03_limpieza_datos.py`
4. `python codigo/04_analisis.py`

## Función de cada script

### 01_extraccion_api.py

Extrae las series mensuales del BCRP mediante API y conserva los datos crudos.

### 02_scraping_web.py

Realiza la descarga programática de los archivos mensuales de la SBS y conserva los archivos originales en `datos_crudos/`.

### 03_limpieza_datos.py

Depura e integra las fuentes, homologa las identidades empresariales, transforma las cifras acumuladas de la SBS en flujos mensuales, deflacta las variables monetarias, calcula los logaritmos y genera la base procesada.

### 04_analisis.py

Lee exclusivamente la base procesada y genera las estimaciones, tablas y figuras del artículo en la carpeta `salidas/`.

## Base procesada

Archivo:

`datos_procesados/datos_procesados_2024200498D.csv`

Características:

- 1,277 observaciones.
- 18 columnas.
- 20 identidades empresariales.
- 72 meses.
- 1,254 observaciones en la muestra del modelo.

## Variables principales del modelo

- Variable dependiente: `log_primas_reales`
- X1: `log_primas_cedidas`
- X2: `log_ingreso_real`
- X3: `tasa_referencia`

El detalle completo de las 18 columnas está documentado en:

`diccionario_variables.md`

## Transformaciones principales

Las cifras de primas de la SBS están expresadas originalmente como acumulados dentro de cada año.

- Enero se utiliza como flujo mensual de enero.
- Desde febrero hasta diciembre se calcula la diferencia entre el acumulado del mes actual y el acumulado del mes anterior.

Las variables monetarias reales se obtienen mediante:

`valor_real = valor_nominal / (IPC / 100)`

El IPC utilizado corresponde a la serie BCRP `PN38705PM`, con base diciembre de 2021 = 100.

Los logaritmos se calculan únicamente cuando los valores son positivos.

Los ceros, valores negativos, faltantes y outliers no se eliminan automáticamente de la base procesada.

## Resultados del Script 04

El análisis genera:

- 5 tablas.
- 7 figuras.
- estimaciones econométricas completas.
- archivos de resultados en `salidas/`.

## Python y dependencias

Versión de Python utilizada:

`Python 3.13.15`

Las versiones exactas de las librerías utilizadas se encuentran en:

`requirements.txt`

## Variables de entorno

El proyecto incluye el archivo:

`.env.example`

No se publican claves ni tokens personales dentro del código.

## SHA-256 de la base procesada entregada

`1ec0fcb5bdc6d16695b94973b1de459b02c2420ea79549a8b8bfc4128911beb8`

El hash corresponde específicamente al archivo:

`datos_procesados/datos_procesados_2024200498D.csv`

## Reproducibilidad

- Se utilizan rutas relativas.
- Los datos crudos se conservan sin edición manual.
- La ventana de estudio está congelada mediante FECHA_INICIO y FECHA_CORTE.
- La base procesada se genera mediante los scripts del proyecto.
- Las tablas y figuras se generan desde la base procesada.
- El orden de ejecución está documentado en este archivo.

## Repositorio GitHub

https://github.com/EgasAraujoVaniaAdriana/FINANZAS_MERCADO-DE-SEGUROS
