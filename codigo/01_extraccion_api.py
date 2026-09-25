# Nombres y apellidos: Vania Adriana Egas Araujo
# Código de matrícula: 2024200498D
# Tema: N.º 15 del temario — Determinantes de las primas cedidas, el ingreso formal y la tasa de referencia en las primas netas del mercado asegurador peruano (2020-2025)
# Fecha de extracción: 2026-09-24
"""
SCRIPT 01 — EXTRACCIÓN VÍA API (BCRPData, Banco Central de Reserva del Perú)

Qué hace:
    Descarga las tres series mensuales del BCRP utilizadas en el estudio.
    Conserva cada respuesta JSON original en /datos_crudos y genera además
    un CSV tabular con los valores sin modificar. Registra fecha, hora,
    código HTTP y número de filas en log_ejecucion.txt.

Series:
    PN31883GM — Ingresos promedio del sector formal total nominal (S/)
    PD04722MM — Tasa de Referencia de la Política Monetaria
    PN38705PM — Índice de Precios al Consumidor de Lima Metropolitana
                 (índice Dic.2021 = 100)

Endpoint:
    https://estadisticas.bcrp.gob.pe/estadisticas/series/api/{codigo}/json/{inicio}/{fin}

Cómo ejecutar desde la raíz de la Carpeta N.º 3:
    python codigo/01_extraccion_api.py
"""

# =============================================================================
# BLOQUE 0. LIBRERÍAS
# =============================================================================
import json
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None


# =============================================================================
# BLOQUE 1. PARÁMETROS CONGELADOS
# =============================================================================
FECHA_INICIO = "2020-01"
FECHA_CORTE = "2025-12"
MESES_ESPERADOS = 72
MATRICULA = "2024200498D"

URL_BASE = "https://estadisticas.bcrp.gob.pe/estadisticas/series/api"

SERIES_MENSUALES = {
    "PN31883GM": "ingreso_formal_nominal",
    "PD04722MM": "tasa_referencia",
    "PN38705PM": "ipc_lima",
}


# =============================================================================
# BLOQUE 2. RUTAS RELATIVAS
# =============================================================================
# El script está dentro de /codigo. Subir un nivel lleva a la raíz
# de la Carpeta N.º 3, sin depender de Google Drive ni de una PC específica.
RAIZ = Path(__file__).resolve().parent.parent
CARPETA_CRUDOS = RAIZ / "datos_crudos"
ARCHIVO_LOG = RAIZ / "log_ejecucion.txt"

CARPETA_CRUDOS.mkdir(parents=True, exist_ok=True)

# Cargar .env desde la raíz del proyecto, si existe.
if load_dotenv is not None:
    load_dotenv(RAIZ / ".env")

CONTACTO = os.getenv("CONTACTO_USER_AGENT", "e_2024200498D@uncp.edu.pe")
CABECERAS = {
    "User-Agent": f"Investigacion-Academica-UNCP-FinanzasI ({CONTACTO})"
}


# =============================================================================
# BLOQUE 3. FUNCIONES AUXILIARES
# =============================================================================
def registrar_log(texto):
    """Añade una línea al log con fecha y hora."""
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linea = f"{ahora} | {texto}"
    with open(ARCHIVO_LOG, "a", encoding="utf-8") as archivo:
        archivo.write(linea + "\n")
    print(linea)


def formato_bcrp(periodo_aaaa_mm):
    """Convierte '2020-01' en '2020-1', formato usado por BCRPData."""
    anio, mes = periodo_aaaa_mm.split("-")
    return f"{anio}-{int(mes)}"


def consultar_serie(codigo, inicio, fin):
    """Consulta una serie del BCRP y devuelve JSON, código HTTP y URL."""
    url = f"{URL_BASE}/{codigo}/json/{inicio}/{fin}"

    try:
        respuesta = requests.get(
            url,
            headers=CABECERAS,
            timeout=60,
        )
    except requests.RequestException as error:
        registrar_log(
            f"BCRP {codigo} | HTTP SIN_CONEXION | filas=0 | {url} | {error}"
        )
        return None, "SIN_CONEXION", url

    if respuesta.status_code != 200:
        registrar_log(
            f"BCRP {codigo} | HTTP {respuesta.status_code} | filas=0 | {url}"
        )
        return None, respuesta.status_code, url

    try:
        datos = respuesta.json()
    except ValueError:
        registrar_log(
            f"BCRP {codigo} | HTTP 200 | respuesta_no_json | filas=0 | {url}"
        )
        return None, "NO_JSON", url

    return datos, respuesta.status_code, url


def nombre_oficial(datos_json):
    """Obtiene el nombre oficial de la serie desde el JSON del BCRP."""
    try:
        return datos_json["config"]["series"][0]["name"]
    except (KeyError, IndexError, TypeError):
        return "nombre no informado por la API"


def json_a_tabla(datos_json, codigo):
    """
    Convierte la sección 'periods' a formato tabular sin transformar valores.
    El periodo y el valor se conservan tal como los entrega la API.
    """
    filas = []

    for periodo in datos_json.get("periods", []):
        valores = periodo.get("values", [])
        valor = valores[0] if valores else None

        filas.append(
            {
                "codigo_serie": codigo,
                "periodo_bcrp": periodo.get("name"),
                "valor": valor,
            }
        )

    return pd.DataFrame(
        filas,
        columns=["codigo_serie", "periodo_bcrp", "valor"],
    )


def validar_tabla(tabla, codigo):
    """Comprueba que la serie tenga 72 meses y ningún valor vacío."""
    if len(tabla) != MESES_ESPERADOS:
        raise RuntimeError(
            f"{codigo}: se esperaban {MESES_ESPERADOS} meses y llegaron {len(tabla)}."
        )

    if tabla["periodo_bcrp"].isna().any():
        raise RuntimeError(f"{codigo}: existen periodos vacíos.")

    vacios = tabla["valor"].isna() | tabla["valor"].astype(str).str.strip().eq("")
    if vacios.any():
        raise RuntimeError(
            f"{codigo}: existen {int(vacios.sum())} valores vacíos en la respuesta."
        )


# =============================================================================
# BLOQUE 4. DESCARGA
# =============================================================================
def main():
    print("=" * 72)
    print(f"EXTRACCIÓN BCRP | {FECHA_INICIO} a {FECHA_CORTE}")
    print("=" * 72)

    inicio = formato_bcrp(FECHA_INICIO)
    fin = formato_bcrp(FECHA_CORTE)

    tablas = []

    for codigo, nombre_corto in SERIES_MENSUALES.items():
        datos, http, url = consultar_serie(codigo, inicio, fin)

        if datos is None:
            raise RuntimeError(
                f"No se pudo descargar {codigo} ({nombre_corto}). "
                "Revise log_ejecucion.txt."
            )

        # 4.1 Guardar la respuesta JSON original como evidencia primaria.
        ruta_json = CARPETA_CRUDOS / f"bcrp_{codigo}_{MATRICULA}.json"
        with open(ruta_json, "w", encoding="utf-8") as archivo:
            json.dump(datos, archivo, ensure_ascii=False, indent=2)

        # 4.2 Crear representación tabular sin modificar los valores.
        tabla = json_a_tabla(datos, codigo)

        registrar_log(
            f"BCRP {codigo} ({nombre_corto}) | HTTP {http} | "
            f"filas={len(tabla)} | {url}"
        )
        registrar_log(f"BCRP {codigo} | Nombre oficial: {nombre_oficial(datos)}")

        # 4.3 Control obligatorio antes de continuar.
        validar_tabla(tabla, codigo)
        tablas.append(tabla)

    # =============================================================================
    # BLOQUE 5. CONSOLIDAR Y GUARDAR CSV CRUDO
    # =============================================================================
    crudo = pd.concat(tablas, ignore_index=True)

    conteos = crudo.groupby("codigo_serie").size()
    if not (conteos == MESES_ESPERADOS).all():
        raise RuntimeError("El control final de 72 meses por serie no se cumplió.")

    total_esperado = MESES_ESPERADOS * len(SERIES_MENSUALES)
    if len(crudo) != total_esperado:
        raise RuntimeError(
            f"Se esperaban {total_esperado} filas totales y se obtuvieron {len(crudo)}."
        )

    ruta_csv = CARPETA_CRUDOS / f"datos_crudos_{MATRICULA}_bcrp.csv"
    crudo.to_csv(ruta_csv, index=False, encoding="utf-8")

    registrar_log(
        f"BCRP CONSOLIDADO | HTTP 200 | filas={len(crudo)} | archivo={ruta_csv.name}"
    )

    print("\nCONTROL DE FILAS:")
    print(conteos.to_string())
    print(f"\n¿72 filas en cada serie? {(conteos == MESES_ESPERADOS).all()}")
    print(f"Total: {len(crudo)} filas")
    print(f"CSV crudo: {ruta_csv.relative_to(RAIZ)}")
    print("\nExtracción BCRP finalizada correctamente.")


if __name__ == "__main__":
    main()
