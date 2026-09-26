# Nombres y apellidos: Vania Adriana Egas Araujo
# Código de matrícula: 2024200498D
# Tema: N.º 15 del temario — Determinantes de las primas cedidas, el ingreso formal y la tasa de referencia en las primas netas del mercado asegurador peruano (2020-2025)
# Fecha de extracción: 2026-09-25
"""
SCRIPT 02 — DESCARGA PROGRAMÁTICA SBS

Descarga los archivos oficiales S-401 de la SBS para 2020-01 a 2025-12.
Los XLS se guardan completos e intactos en datos_crudos.

De esta fuente, el proyecto utilizará únicamente:
- primas de seguros netas
- primas cedidas

La selección de esas dos variables y las transformaciones se harán en
03_limpieza_datos.py. No se edita el crudo en este script.
"""

# =============================================================================
# BLOQUE 0. LIBRERÍAS
# =============================================================================
import hashlib
import os
import time
from datetime import datetime
from pathlib import Path
from urllib import robotparser

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

CODIGO_TABLA = "S-401"
URL_BASE = "https://intranet2.sbs.gob.pe/estadistica/financiera"
PAUSA_SEGUNDOS = 1.5
TIMEOUT_SEGUNDOS = 60
MAX_INTENTOS = 3

MESES_SBS = {
    1: ("Enero", "en"),
    2: ("Febrero", "fe"),
    3: ("Marzo", "ma"),
    4: ("Abril", "ab"),
    5: ("Mayo", "my"),
    6: ("Junio", "jn"),
    7: ("Julio", "jl"),
    8: ("Agosto", "ag"),
    9: ("Setiembre", "se"),
    10: ("Octubre", "oc"),
    11: ("Noviembre", "no"),
    12: ("Diciembre", "di"),
}


# =============================================================================
# BLOQUE 2. RUTAS RELATIVAS
# =============================================================================
RAIZ = Path(__file__).resolve().parent.parent
CARPETA_CRUDOS = RAIZ / "datos_crudos"
ARCHIVO_LOG = RAIZ / "log_ejecucion.txt"
CARPETA_CRUDOS.mkdir(parents=True, exist_ok=True)

if load_dotenv is not None:
    load_dotenv(RAIZ / ".env")

CONTACTO = os.getenv(
    "CONTACTO_USER_AGENT",
    "Estudiante-UNCP-FinanzasI-2024200498D",
)

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


def lista_de_meses():
    """Devuelve exactamente los meses comprendidos entre inicio y corte."""
    meses = pd.period_range(
        start=FECHA_INICIO,
        end=FECHA_CORTE,
        freq="M",
    )
    return [(periodo.year, periodo.month) for periodo in meses]


def url_mensual(anio, mes):
    """Construye la URL oficial del archivo S-401 para un mes."""
    carpeta_mes, abreviatura = MESES_SBS[mes]
    return (
        f"{URL_BASE}/{anio}/{carpeta_mes}/"
        f"{CODIGO_TABLA}-{abreviatura}{anio}.XLS"
    )


def ruta_local(anio, mes):
    """Nombre local del XLS original dentro de datos_crudos."""
    return (
        CARPETA_CRUDOS
        / f"sbs_{CODIGO_TABLA}_{anio}-{mes:02d}_{MATRICULA}.xls"
    )


def es_excel(contenido):
    """Valida la firma binaria de un XLS clásico o un XLSX."""
    firma_xls = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    return (
        contenido.startswith(firma_xls)
        or contenido.startswith(b"PK")
    )


def huella_sha256(contenido):
    """Calcula el SHA-256 completo del archivo crudo."""
    return hashlib.sha256(contenido).hexdigest()


def contar_filas(ruta):
    """Cuenta las filas de la primera hoja para registrarlas en el log."""
    tabla = pd.read_excel(
        ruta,
        sheet_name=0,
        header=None,
        engine="xlrd",
    )
    return len(tabla)


def revisar_robots(sesion):
    """
    Revisa robots.txt del servidor que aloja los XLS.
    Si existe una prohibición explícita, el script se detiene.
    """
    robots_url = "https://intranet2.sbs.gob.pe/robots.txt"
    objetivo = url_mensual(2025, 12)

    try:
        respuesta = sesion.get(robots_url, timeout=TIMEOUT_SEGUNDOS)
        time.sleep(PAUSA_SEGUNDOS)
    except requests.RequestException as error:
        registrar_log(
            f"ROBOTS {robots_url} | NO_CONCLUYENTE | error={error}"
        )
        return

    registrar_log(
        f"ROBOTS {robots_url} | HTTP {respuesta.status_code}"
    )

    if respuesta.status_code == 200:
        analizador = robotparser.RobotFileParser()
        analizador.set_url(robots_url)
        analizador.parse(respuesta.text.splitlines())
        permitido = analizador.can_fetch(
            CABECERAS["User-Agent"],
            objetivo,
        )
        registrar_log(
            f"ROBOTS descarga SBS | permitido={permitido} | objetivo={objetivo}"
        )
        if not permitido:
            raise RuntimeError(
                "robots.txt prohíbe la ruta de descarga SBS. "
                "No se continuará con el scraping."
            )
    elif respuesta.status_code in {404, 410}:
        registrar_log(
            "ROBOTS SBS | archivo robots.txt no publicado en ese host"
        )
    else:
        registrar_log(
            f"ROBOTS SBS | revisión no concluyente | "
            f"HTTP {respuesta.status_code}"
        )


def descargar_con_reintentos(sesion, url):
    """Descarga una URL con hasta tres intentos y manejo de errores."""
    ultimo_error = None

    for intento in range(1, MAX_INTENTOS + 1):
        try:
            respuesta = sesion.get(url, timeout=TIMEOUT_SEGUNDOS)
            time.sleep(PAUSA_SEGUNDOS)

            if respuesta.status_code in {429, 500, 502, 503, 504}:
                ultimo_error = f"HTTP {respuesta.status_code}"
                continue

            return respuesta

        except requests.RequestException as error:
            ultimo_error = error
            time.sleep(PAUSA_SEGUNDOS)

    raise RuntimeError(
        f"No se pudo descargar {url} después de {MAX_INTENTOS} intentos. "
        f"Último error: {ultimo_error}"
    )


# =============================================================================
# BLOQUE 4. DESCARGA MES POR MES
# =============================================================================
def main():
    print("=" * 78)
    print(
        f"DESCARGA SBS {CODIGO_TABLA} | "
        f"{FECHA_INICIO} a {FECHA_CORTE}"
    )
    print("=" * 78)

    meses = lista_de_meses()

    if len(meses) != MESES_ESPERADOS:
        raise RuntimeError(
            f"La ventana produce {len(meses)} meses; "
            f"se esperaban {MESES_ESPERADOS}."
        )

    sesion = requests.Session()
    sesion.headers.update(CABECERAS)

    # La consigna exige revisar robots.txt antes de automatizar.
    revisar_robots(sesion)

    faltantes = []

    for anio, mes in meses:
        periodo = f"{anio}-{mes:02d}"
        url = url_mensual(anio, mes)
        destino = ruta_local(anio, mes)

        # Si ya existe un XLS válido y legible, no se vuelve a solicitar.
        if destino.exists():
            contenido_existente = destino.read_bytes()
            if es_excel(contenido_existente):
                try:
                    filas = contar_filas(destino)
                    print(
                        f"{periodo} | ya descargado | "
                        f"filas_primera_hoja={filas} | {destino.name}"
                    )
                    continue
                except Exception:
                    pass

        try:
            respuesta = descargar_con_reintentos(sesion, url)
            http = respuesta.status_code
            contenido = respuesta.content
        except Exception as error:
            registrar_log(
                f"SBS {CODIGO_TABLA} {periodo} | SIN_CONEXION | "
                f"error={error} | {url}"
            )
            faltantes.append(periodo)
            continue

        # Solo se guarda si la respuesta es 200 y contiene un Excel real.
        if http == 200 and es_excel(contenido):
            destino.write_bytes(contenido)

            try:
                filas = contar_filas(destino)
            except Exception as error:
                registrar_log(
                    f"SBS {CODIGO_TABLA} {periodo} | HTTP {http} | "
                    f"XLS NO LEGIBLE | error={error} | {url}"
                )
                destino.unlink(missing_ok=True)
                faltantes.append(periodo)
                continue

            registrar_log(
                f"SBS {CODIGO_TABLA} {periodo} | HTTP {http} | "
                f"filas_primera_hoja={filas} | bytes={len(contenido)} | "
                f"sha256={huella_sha256(contenido)} | "
                f"archivo={destino.name} | {url}"
            )
        else:
            registrar_log(
                f"SBS {CODIGO_TABLA} {periodo} | HTTP {http} | "
                f"NO DESCARGADO | content_type="
                f"{respuesta.headers.get('Content-Type', '')} | {url}"
            )
            faltantes.append(periodo)

    # =============================================================================
    # BLOQUE 5. CONTROL FINAL
    # =============================================================================
    archivos = sorted(
        CARPETA_CRUDOS.glob(
            f"sbs_{CODIGO_TABLA}_*_{MATRICULA}.xls"
        )
    )

    registrar_log(
        f"SBS {CODIGO_TABLA} CONTROL | "
        f"archivos={len(archivos)} de {MESES_ESPERADOS}"
    )

    print(f"\nArchivos SBS: {len(archivos)} de {MESES_ESPERADOS}")
    if archivos:
        print(f"Primero: {archivos[0].name}")
        print(f"Último:  {archivos[-1].name}")

    if faltantes or len(archivos) != MESES_ESPERADOS:
        raise RuntimeError(
            "La descarga no quedó completa. "
            f"Meses con problema: {sorted(set(faltantes))}. "
            "Revise log_ejecucion.txt."
        )

    print("\nDescarga SBS finalizada correctamente.")


if __name__ == "__main__":
    main()
