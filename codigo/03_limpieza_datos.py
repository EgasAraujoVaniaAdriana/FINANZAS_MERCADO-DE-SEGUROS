# Nombres y apellidos: Vania Adriana Egas Araujo
# Código de matrícula: 2024200498D
# Tema: N.º 15 del temario — Determinantes de las primas cedidas, el ingreso formal y la tasa de referencia en las primas netas del mercado asegurador peruano (2020-2025)
# Fecha de extracción: 2026-09-25

"""
SCRIPT 03 — LIMPIEZA, TRANSFORMACIÓN E INTEGRACIÓN DE DATOS

Unidad de análisis:
    Empresa × mes, desde 2020-01 hasta 2025-12.

Fuentes:
    - SBS: primas netas y primas cedidas del formato S-401.
    - BCRP: ingreso formal, tasa de referencia e IPC.

Tratamiento SBS:
    - P009 corresponde a primas netas acumuladas por empresa.
    - P010 corresponde a primas cedidas acumuladas por empresa.
    - Se utiliza únicamente la fila TOTAL de cada empresa.
    - No se trabaja por ramo.
    - Los archivos crudos no se modifican.
    - Enero mensual equivale al acumulado de enero.
    - De febrero a diciembre:
          mensual = acumulado actual - acumulado del mes anterior.
    - Si falta el mes anterior, el flujo mensual no se inventa.

Identidad empresarial:
    - Se construye un id_empresa estable para cada trayectoria.
    - La normalización de nombres permite detectar variantes de escritura,
      pero no determina por sí sola la identidad de una empresa.
    - Las variantes de nombre se contrastan con continuidad temporal y
      evidencia de la SBS.
    - Los cambios simples de denominación mantienen una misma trayectoria
      cuando existe continuidad verificable.
    - Las fusiones y absorciones se tratan separadamente y no se confunden
      con simples cambios de nombre.
    - La columna empresa contiene una etiqueta analítica estable para
      tablas y gráficos.
    - El nombre original observado en SBS se conserva durante el proceso
      para realizar controles de trazabilidad.

Transformaciones:
    - Las primas se expresan en miles de soles reales de diciembre de 2021.
    - El ingreso formal se expresa en soles reales de diciembre de 2021.
    - Los logaritmos se calculan solo para valores estrictamente positivos.
    - Ceros, negativos y faltantes se conservan.
    - Los outliers se diagnostican, pero no se eliminan ni winsorizan.

Salida:
    - El Script 03 genera:
      datos_procesados/datos_procesados_2024200498D.csv
    - Los controles del proceso se muestran durante la ejecución y se
      registran en log_ejecucion.txt.
    - La carpeta /salidas se reserva para los resultados generados
      posteriormente por el Script 04.
"""


# =============================================================================
# BLOQUE 1. LIBRERÍAS
# =============================================================================
import hashlib
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


# =============================================================================
# BLOQUE 2. PARÁMETROS, VARIABLES Y RUTAS RELATIVAS
# =============================================================================

# -----------------------------------------------------------------------------
# 2.1. Ventana temporal congelada
# -----------------------------------------------------------------------------
# El análisis utiliza exactamente enero de 2020 a diciembre de 2025.
FECHA_INICIO = "2020-01"
FECHA_CORTE = "2025-12"
MESES_ESPERADOS = 72

MATRICULA = "2024200498D"


# -----------------------------------------------------------------------------
# 2.2. Series BCRP utilizadas
# -----------------------------------------------------------------------------
# PN31883GM : ingreso promedio del sector formal
# PD04722MM : tasa de referencia del BCRP
# PN38705PM : IPC, base diciembre 2021 = 100
COD_INGRESO = "PN31883GM"
COD_TASA = "PD04722MM"
COD_IPC = "PN38705PM"


# -----------------------------------------------------------------------------
# 2.3. Hojas del formato SBS S-401
# -----------------------------------------------------------------------------
HOJAS = {
    "P009": "primas_netas",
    "P010": "primas_cedidas",
}

ROTULO_ENCABEZADO = r"RIESGOS\s*/\s*EMPRESAS"

ROTULOS_TOTAL_SISTEMA = {
    "TOTAL",
    "TOTAL GENERAL",
    "TOTAL SISTEMA",
    "TOTAL SISTEMA ASEGURADOR",
}

ROTULOS_TOTAL_FILA = {
    "TOTAL",
    "TOTAL GENERAL",
}


# -----------------------------------------------------------------------------
# 2.4. Parámetros de control
# -----------------------------------------------------------------------------
MINIMO_OBSERVACIONES = 1000

# Miles de soles.
# Se admite esta diferencia únicamente para controles de redondeo de la fuente.
TOLERANCIA = 1.0


# -----------------------------------------------------------------------------
# 2.5. Meses utilizados por las series del BCRP
# -----------------------------------------------------------------------------
MESES_BCRP = {
    "ENE": 1,
    "FEB": 2,
    "MAR": 3,
    "ABR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AGO": 8,
    "SEP": 9,
    "SET": 9,
    "OCT": 10,
    "NOV": 11,
    "DIC": 12,
}


# -----------------------------------------------------------------------------
# 2.6. Columnas de la base procesada final
# -----------------------------------------------------------------------------
# empresa_sbs se conserva durante la limpieza para trazabilidad,
# pero NO se exporta a datos_procesados.
#
# estado_p009_p010 y los flags de huecos son controles internos del
# procedimiento y tampoco forman parte de la base analítica final.
#
# id_observacion se incorpora posteriormente, antes del guardado.
COLUMNAS_FINALES = [
    "periodo",
    "id_empresa",
    "empresa",

    # SBS: primas netas
    "primas_netas_acum_sbs",
    "primas_netas_mensual_nominal",
    "primas_netas_mensual_real",
    "log_primas_reales",

    # SBS: primas cedidas
    "primas_cedidas_acum_sbs",
    "primas_cedidas_mensual_nominal",
    "primas_cedidas_mensual_real",
    "log_primas_cedidas",

    # BCRP
    "ingreso_formal_nominal",
    "ingreso_real",
    "log_ingreso_real",
    "tasa_referencia",
    "ipc",

    # Indicador para seleccionar la muestra en el Script 04
    "muestra_modelo",
]


# -----------------------------------------------------------------------------
# 2.7. Rutas relativas
# -----------------------------------------------------------------------------
# No se incluyen rutas de Google Drive ni de Colab dentro del script.
# RAIZ corresponde automáticamente a la carpeta N.º 3 del proyecto.
RAIZ = Path(__file__).resolve().parent.parent

CARPETA_CRUDOS = RAIZ / "datos_crudos"
CARPETA_PROC = RAIZ / "datos_procesados"

# /salidas se reserva para tablas y figuras del Script 04.
CARPETA_SALIDAS = RAIZ / "salidas"

ARCHIVO_LOG = RAIZ / "log_ejecucion.txt"


# -----------------------------------------------------------------------------
# 2.8. Codificación de archivos CSV
# -----------------------------------------------------------------------------
# UTF-8 con BOM permite abrir los CSV directamente en Excel conservando
# correctamente tildes, eñes y otros caracteres.
CSV_ENCODING = "utf-8-sig"


# -----------------------------------------------------------------------------
# 2.9. Preparación de carpetas
# -----------------------------------------------------------------------------
CARPETA_PROC.mkdir(
    parents=True,
    exist_ok=True
)

CARPETA_SALIDAS.mkdir(
    parents=True,
    exist_ok=True
)


# Acumula los mensajes que posteriormente se registrarán en el log.
REPORTE = []


# =============================================================================
# BLOQUE 3. FUNCIONES AUXILIARES
# =============================================================================
def reportar(texto=""):
    print(texto)
    REPORTE.append(str(texto))


def normalizar(texto):
    """Mayúsculas, sin tildes y con espacios simples."""
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto).strip().upper()


def sin_puntuacion(texto):
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", normalizar(texto))).strip()


def tokens(texto):
    return [t for t in sin_puntuacion(texto).split() if t]


def es_subsecuencia(pequena, grande):
    """True si los tokens de 'pequena' aparecen en el mismo orden dentro de 'grande'."""
    it = iter(grande)
    return all(any(x == y for y in it) for x in pequena)


def periodo_a_numero(periodo):
    anio, mes = str(periodo).split("-")
    return int(anio) * 12 + int(mes) - 1


def periodos_esperados():
    return [str(p) for p in pd.period_range(FECHA_INICIO, FECHA_CORTE, freq="M")]


def letra_columna(indice):
    letras = ""
    indice += 1
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def convertir_valor(celda):
    """
    Conserva el valor interno del XLS sin redondearlo.
    0.09 -> 0.09; 0 -> 0; guion textual -> 0; '(12.3)' -> -12.3; vacío -> NaN.
    """
    if celda is None or (isinstance(celda, float) and np.isnan(celda)):
        return np.nan, "vacio"
    if isinstance(celda, (int, float, np.integer, np.floating)):
        return float(celda), "numero"

    texto = str(celda).strip().replace("\xa0", " ")
    if texto == "":
        return np.nan, "vacio"
    if texto in {"-", "--", "–", "—"}:
        return 0.0, "guion"

    negativo = False
    entre_parentesis = re.fullmatch(r"\(\s*(.+?)\s*\)", texto)
    if entre_parentesis:
        texto, negativo = entre_parentesis.group(1), True

    limpio = texto.replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", limpio):
        limpio = limpio.replace(",", "")
    elif re.fullmatch(r"-?\d+,\d+", limpio):
        limpio = limpio.replace(",", ".")

    try:
        numero = float(limpio)
    except ValueError:
        return np.nan, "texto_no_numerico"
    return (-numero if negativo else numero), ("parentesis" if negativo else "texto_numerico")


def limpiar_encabezado(texto):
    """'Rigel (1)' -> ('Rigel', '1'). Las referencias se conservan para vincular notas."""
    nombre = re.sub(r"\s+", " ", str(texto)).strip()
    referencias = []
    while True:
        llamada = re.search(r"\s*(\(\s*\d{1,2}\s*\)|\*+)\s*$", nombre)
        if not llamada or llamada.start() == 0:
            break
        ref = llamada.group(1).replace(" ", "")
        ref = ref.strip("()") if ref.startswith("(") else ref
        referencias.insert(0, ref)
        nombre = nombre[:llamada.start()].strip()
    return nombre, ";".join(referencias)


def clasificar_columna(nombre_limpio):
    nombre = normalizar(nombre_limpio)
    if nombre in ROTULOS_TOTAL_SISTEMA:
        return "total_sistema"
    if "+" in nombre or nombre.startswith("TOTAL"):
        return "agregada"
    return "empresa"


def referencias_en_texto(texto):
    """Extrae llamadas cortas tipo (1), (2), evitando fechas como (25/04/2020)."""
    refs = re.findall(r"\(\s*(\d{1,2})\s*\)", str(texto))
    return ";".join(dict.fromkeys(refs))


def resolver_mencion_empresa(texto, nombres):
    """Resuelve una mención a una empresa sin usar un diccionario manual."""
    objetivo = sin_puntuacion(texto)
    exactos = [n for n in nombres if sin_puntuacion(n) == objetivo]
    if len(exactos) == 1:
        return exactos[0]
    candidatos = [n for n in nombres if objetivo and (
        sin_puntuacion(n).startswith(objetivo + " ") or objetivo.startswith(sin_puntuacion(n) + " ")
    )]
    if len(candidatos) == 1:
        return candidatos[0]
    if not candidatos:
        return None
    return None


# =============================================================================
# BLOQUE 4. BCRP -> una fila por mes
# =============================================================================
def cargar_bcrp():
    ruta = CARPETA_CRUDOS / f"datos_crudos_{MATRICULA}_bcrp.csv"
    if not ruta.exists():
        raise FileNotFoundError("Falta el crudo del BCRP. Ejecute primero 01_extraccion_api.py")

    crudo = pd.read_csv(ruta, dtype=str)
    requeridas = {"periodo_bcrp", "codigo_serie", "valor"}
    faltan = requeridas - set(crudo.columns)
    if faltan:
        raise RuntimeError(f"BCRP: faltan columnas requeridas: {sorted(faltan)}")

    def a_periodo(texto):
        partes = str(texto).strip().split(".")
        if len(partes) != 2:
            raise ValueError(f"Periodo BCRP no reconocido: {texto}")
        mes, anio = partes
        clave = normalizar(mes)[:3]
        if clave not in MESES_BCRP:
            raise ValueError(f"Mes BCRP no reconocido: {texto}")
        return f"{anio}-{MESES_BCRP[clave]:02d}"

    crudo["periodo"] = crudo["periodo_bcrp"].apply(a_periodo)
    crudo["valor"] = pd.to_numeric(crudo["valor"], errors="coerce")

    bcrp = (crudo.pivot(index="periodo", columns="codigo_serie", values="valor")
                 .rename(columns={COD_INGRESO: "ingreso_formal_nominal",
                                  COD_TASA: "tasa_referencia",
                                  COD_IPC: "ipc"})
                 .reset_index())
    bcrp.columns.name = None

    requeridas_final = ["periodo", "ingreso_formal_nominal", "tasa_referencia", "ipc"]
    if not set(requeridas_final).issubset(bcrp.columns):
        raise RuntimeError("BCRP: falta alguna de las tres series requeridas.")
    bcrp = bcrp[requeridas_final]
    bcrp = bcrp[(bcrp["periodo"] >= FECHA_INICIO) & (bcrp["periodo"] <= FECHA_CORTE)].copy()

    encontrados = sorted(bcrp["periodo"].astype(str).tolist())
    esperados = periodos_esperados()
    if bcrp["periodo"].duplicated().any() or encontrados != esperados:
        faltantes = sorted(set(esperados) - set(encontrados))
        sobrantes = sorted(set(encontrados) - set(esperados))
        raise RuntimeError(f"BCRP: ventana mensual incorrecta. Faltantes={faltantes}; sobrantes={sobrantes}")
    if bcrp.isna().any().any():
        raise RuntimeError("BCRP: hay valores vacíos en ingreso, tasa o IPC.")
    if (bcrp["ipc"] <= 0).any():
        raise RuntimeError("BCRP: el IPC debe ser estrictamente positivo en los 72 meses.")

    reportar(f"BCRP: {len(bcrp)} meses exactos ({FECHA_INICIO} a {FECHA_CORTE}), sin vacíos e IPC > 0.")
    return bcrp


# =============================================================================
# BLOQUE 5. LECTURA ESTRUCTURAL DE LA SBS (fila TOTAL por empresa)
# =============================================================================
def ubicar_filas(df, nombre_archivo, hoja):
    rotulos = df.iloc[:, 0].apply(lambda v: normalizar(v) if isinstance(v, str) else "")
    encabezado = rotulos[rotulos.str.contains(ROTULO_ENCABEZADO, regex=True, na=False)]
    if encabezado.empty:
        raise RuntimeError(f"{nombre_archivo} {hoja}: no se encontró 'Riesgos / Empresas'.")
    fila_enc = encabezado.index[0]

    total = rotulos[rotulos.isin(ROTULOS_TOTAL_FILA) & (rotulos.index > fila_enc)]
    if total.empty:
        raise RuntimeError(f"{nombre_archivo} {hoja}: no se encontró la fila TOTAL.")
    return fila_enc, total.index[-1]


def componentes_agregada(nombre_agregado, nombres_empresas):
    """'TOTAL + RIGEL + COFACE' -> nombres reales mencionados en esa columna."""
    piezas = [p.strip() for p in normalizar(nombre_agregado).split("+")]
    componentes = []
    for pieza in piezas[1:]:
        if not pieza:
            continue
        empresa = resolver_mencion_empresa(pieza, nombres_empresas)
        if empresa is None:
            raise RuntimeError(f"No se pudo resolver la empresa mencionada en columna agregada: '{pieza}'")
        componentes.append(empresa)
    return list(dict.fromkeys(componentes))


def controlar_total_sistema(empresas, totales, nombre_archivo, hoja):
    """
    Valida cada columna TOTAL / TOTAL + ... de manera individual.
    Nunca usa np.nansum para ocultar faltantes en empresas reales.
    """
    if not totales:
        return {"estado": "sin_total", "detalle": "sin columna TOTAL/AGREGADA"}

    faltantes = [e["empresa_sbs"] for e in empresas if pd.isna(e["valor_acumulado"])]
    if faltantes:
        return {"estado": "no_validable", "detalle": "totales faltantes: " + "; ".join(faltantes)}

    valores = {e["empresa_sbs"]: float(e["valor_acumulado"]) for e in empresas}
    nombres = list(valores)
    suma_todas = sum(valores.values())

    agregadas = [t for t in totales if t["tipo_columna"] == "agregada" and pd.notna(t["valor"])]
    componentes_por_agregada = {}
    union_especiales = set()
    for t in agregadas:
        comps = componentes_agregada(t["nombre_limpio"], nombres)
        componentes_por_agregada[t["nombre_limpio"]] = comps
        union_especiales.update(comps)

    suma_base_esperada = sum(v for n, v in valores.items() if n not in union_especiales)
    totales_sistema = [t for t in totales if t["tipo_columna"] == "total_sistema" and pd.notna(t["valor"])]

    detalles = []
    valor_total_base = None
    if totales_sistema:
        # Si hay más de uno, todos deben ser coherentes con alguna interpretación válida.
        for t in totales_sistema:
            valor = float(t["valor"])
            if abs(valor - suma_todas) <= TOLERANCIA:
                detalles.append(f"{t['encabezado_original']} = suma de todas las empresas")
            elif union_especiales and abs(valor - suma_base_esperada) <= TOLERANCIA:
                detalles.append(f"{t['encabezado_original']} = total base sin {', '.join(sorted(union_especiales))}")
                valor_total_base = valor
            else:
                raise RuntimeError(
                    f"{nombre_archivo} {hoja}: '{t['encabezado_original']}'={valor:,.2f} no cuadra "
                    f"ni con suma total={suma_todas:,.2f} ni con base={suma_base_esperada:,.2f}."
                )

    # Cada TOTAL + ... se valida contra SUS componentes, no contra la unión global.
    base_referencia = valor_total_base if valor_total_base is not None else suma_base_esperada
    for t in agregadas:
        valor = float(t["valor"])
        comps = componentes_por_agregada[t["nombre_limpio"]]
        esperado = base_referencia + sum(valores[c] for c in comps)
        alternativas = [esperado]
        if set(comps) == union_especiales:
            alternativas.append(suma_todas)
        if not any(abs(valor - e) <= TOLERANCIA for e in alternativas):
            raise RuntimeError(
                f"{nombre_archivo} {hoja}: '{t['encabezado_original']}'={valor:,.2f} no cuadra con "
                f"base + sus componentes ({esperado:,.2f}). Componentes={comps}"
            )
        detalles.append(f"{t['encabezado_original']} validada con {', '.join(comps)}")

    return {"estado": "ok", "detalle": "; ".join(detalles) if detalles else "sin valores comparables"}


def leer_hoja(ruta, hoja, periodo):
    df = pd.read_excel(ruta, sheet_name=hoja, header=None, engine="xlrd")
    fila_enc, fila_tot = ubicar_filas(df, ruta.name, hoja)

    empresas, columnas, totales = [], [], []
    for c in df.columns[1:]:
        encabezado = df.at[fila_enc, c]
        if not isinstance(encabezado, str) or not encabezado.strip():
            continue
        nombre, referencia = limpiar_encabezado(encabezado)
        tipo = clasificar_columna(nombre)
        valor, tipo_valor = convertir_valor(df.at[fila_tot, c])

        columnas.append({
            "periodo": periodo, "hoja": hoja, "columna_excel": letra_columna(c),
            "encabezado_original": encabezado.strip(), "nombre_limpio": nombre,
            "referencia_nota": referencia, "tipo_columna": tipo,
        })

        if tipo == "empresa":
            if tipo_valor == "texto_no_numerico":
                raise RuntimeError(f"{ruta.name} {hoja}: TOTAL no numérico para '{nombre}'.")
            empresas.append({
                "periodo": periodo, "hoja": hoja, "empresa_sbs": nombre,
                "referencia_nota": referencia, "valor_acumulado": valor,
                "tipo_valor": tipo_valor,
            })
        else:
            totales.append({
                "encabezado_original": encabezado.strip(), "nombre_limpio": nombre,
                "tipo_columna": tipo, "valor": valor, "tipo_valor": tipo_valor,
            })

    nombres = [e["empresa_sbs"] for e in empresas]
    if len(nombres) != len(set(nombres)):
        repetidos = sorted({n for n in nombres if nombres.count(n) > 1})
        raise RuntimeError(f"{ruta.name} {hoja}: nombres de empresa duplicados tras limpiar: {repetidos}")

    control = controlar_total_sistema(empresas, totales, ruta.name, hoja)
    notas = extraer_notas(df, fila_tot, periodo, hoja, ruta.name)
    return empresas, columnas, notas, control


def cargar_sbs(guardar_auditoria=True):
    archivos = sorted(CARPETA_CRUDOS.glob(f"sbs_S-401_*_{MATRICULA}.xls"))
    periodos_archivo = []
    por_periodo = {}
    for ruta in archivos:
        m = re.search(r"(\d{4}-\d{2})", ruta.name)
        if not m:
            raise RuntimeError(f"No se pudo reconocer el periodo de {ruta.name}")
        periodo = m.group(1)
        periodos_archivo.append(periodo)
        por_periodo.setdefault(periodo, []).append(ruta.name)

    esperados = periodos_esperados()
    if sorted(periodos_archivo) != esperados or any(len(v) != 1 for v in por_periodo.values()):
        faltantes = sorted(set(esperados) - set(periodos_archivo))
        sobrantes = sorted(set(periodos_archivo) - set(esperados))
        duplicados = {k: v for k, v in por_periodo.items() if len(v) > 1}
        raise RuntimeError(
            f"SBS: periodos incorrectos. Faltantes={faltantes}; sobrantes={sobrantes}; duplicados={duplicados}"
        )

    empresas, columnas, notas, controles = [], [], [], []
    for ruta in archivos:
        periodo = re.search(r"(\d{4}-\d{2})", ruta.name).group(1)
        for hoja in HOJAS:
            e, c, n, control = leer_hoja(ruta, hoja, periodo)
            empresas.extend(e)
            columnas.extend(c)
            notas.extend(n)
            controles.append({"periodo": periodo, "hoja": hoja, **control})

    largo = pd.DataFrame(empresas)
    largo["variable"] = largo["hoja"].map(HOJAS)
    columnas_df = pd.DataFrame(columnas)
    notas_df = pd.DataFrame(notas)
    controles_df = pd.DataFrame(controles)

    claves_largo = ["periodo", "empresa_sbs", "variable"]
    if largo.duplicated(claves_largo).any():
        ejemplo = largo.loc[largo.duplicated(claves_largo, keep=False), claves_largo].head(20)
        raise RuntimeError("SBS: duplicados periodo × empresa × variable:\n" + ejemplo.to_string(index=False))

    # Presencia comparada P009 vs P010: ausencia no se convierte en cero.
    presencia = (largo.assign(presente=1)
                      .pivot_table(index=["periodo", "empresa_sbs"], columns="variable",
                                   values="presente", aggfunc="max", fill_value=0)
                      .reset_index())
    presencia.columns.name = None
    for col in ["primas_netas", "primas_cedidas"]:
        if col not in presencia.columns:
            presencia[col] = 0
    presencia["estado_p009_p010"] = np.select(
        [presencia["primas_netas"].eq(1) & presencia["primas_cedidas"].eq(1),
         presencia["primas_netas"].eq(1) & presencia["primas_cedidas"].eq(0),
         presencia["primas_netas"].eq(0) & presencia["primas_cedidas"].eq(1)],
        ["ambas", "solo_P009", "solo_P010"], default="ninguna"
    )
    comparacion = presencia[["periodo", "empresa_sbs", "estado_p009_p010"]].copy()

    sbs = (largo.set_index(["periodo", "empresa_sbs", "variable"])["valor_acumulado"]
                .unstack("variable")
                .reset_index())
    sbs.columns.name = None
    sbs = sbs.rename(columns={"primas_netas": "primas_netas_acum_sbs",
                              "primas_cedidas": "primas_cedidas_acum_sbs"})
    sbs = sbs.merge(comparacion, on=["periodo", "empresa_sbs"], how="left", validate="one_to_one")

    tabla_col = (columnas_df.groupby(
                    ["encabezado_original", "nombre_limpio", "referencia_nota", "tipo_columna"], dropna=False)
                    ["periodo"].agg(primer_periodo="min", ultimo_periodo="max", apariciones="count")
                    .reset_index())

    if guardar_auditoria:
        tabla_col.to_csv(CARPETA_SALIDAS / "columnas_encabezado_sbs.csv", index=False, encoding="utf-8")
        comparacion.to_csv(CARPETA_SALIDAS / "comparacion_p009_p010.csv", index=False, encoding="utf-8")
        controles_df.to_csv(CARPETA_SALIDAS / "controles_total_sistema.csv", index=False, encoding="utf-8")

    reportar(f"SBS: {len(archivos)} archivos exactos × {len(HOJAS)} hojas; fila TOTAL por empresa.")
    discrepancias = int((comparacion["estado_p009_p010"] != "ambas").sum())
    reportar(f"  Comparación P009/P010: {discrepancias} empresa-mes no aparecen en ambas hojas.")
    if discrepancias:
        reportar("  Esas ausencias se conservan como NaN; nunca se convierten automáticamente en cero.")
    no_validables = int((controles_df["estado"] != "ok").sum()) if not controles_df.empty else 0
    reportar(f"  Controles TOTAL: {len(controles_df) - no_validables} validaciones OK; {no_validables} no validables/sin total.")

    return sbs, largo, notas_df, columnas_df, comparacion


# =============================================================================
# BLOQUE 6. EXTRACCIÓN DE PIES DE PÁGINA DE LA SBS
# =============================================================================
# Los pies de página se extraen porque contienen información útil para
# interpretar cambios de denominación, fusiones, absorciones y otros eventos.
#
# Se utilizan durante el procesamiento del Script 03, pero NO se exportan
# como archivos independientes a /salidas.
# =============================================================================


def extraer_notas(
    df,
    fila_tot,
    periodo,
    hoja,
    nombre_archivo
):
    """
    Extrae las notas ubicadas debajo de la fila TOTAL de cada hoja SBS.

    Cada fila de nota se conserva completa para no separar una misma
    explicación entre varias celdas.

    La información se mantiene en memoria y sirve posteriormente para
    verificar eventos empresariales y registros especiales.
    """

    notas = []

    for i in range(
        fila_tot + 1,
        len(df)
    ):

        partes = []
        columnas = []

        for c in df.columns:

            valor = df.at[i, c]

            if (
                isinstance(valor, str)
                and valor.strip()
            ):

                texto = re.sub(
                    r"\s+",
                    " ",
                    valor.strip()
                )

                partes.append(
                    texto
                )

                columnas.append(
                    letra_columna(c)
                )

        # Si la fila no contiene texto, se ignora.
        if not partes:
            continue

        # Evitar duplicar exactamente la misma pieza
        # cuando aparece repetida en una fila.
        partes = list(
            dict.fromkeys(partes)
        )

        nota = " | ".join(
            partes
        )

        notas.append({
            "periodo": periodo,
            "hoja": hoja,
            "fila_excel": i + 1,
            "columnas_texto": ";".join(columnas),
            "nota_original": nota,
            "nota_normalizada": normalizar(nota),
            "referencias_detectadas": referencias_en_texto(nota),
            "archivo_origen": nombre_archivo,
        })

    return notas


def resumir_notas(notas):
    """
    Resume en consola las notas encontradas.

    No crea archivos en /salidas.
    Las notas continúan disponibles en memoria para los controles
    de identidad empresarial y eventos societarios.
    """

    if notas is None or notas.empty:

        reportar(
            "Pies de página SBS: "
            "no se encontraron textos debajo de TOTAL."
        )

        return

    total_notas = len(
        notas
    )

    textos_distintos = (
        notas[
            "nota_normalizada"
        ]
        .nunique()
    )

    con_referencia = (
        notas[
            "referencias_detectadas"
        ]
        .fillna("")
        .astype(str)
        .str.strip()
        .ne("")
        .sum()
    )

    reportar(
        "Pies de página SBS: "
        f"{total_notas:,} filas extraídas; "
        f"{textos_distintos:,} textos normalizados distintos; "
        f"{con_referencia:,} filas con referencia detectada."
    )


# =============================================================================
# BLOQUE 6B. DEPURACIÓN DE REGISTROS ESPECIALES SBS
# =============================================================================
# Algunas columnas de los XLS tienen apariencia de empresa, pero las notas
# al pie pueden indicar que representan flujos históricos/transitorios de una
# empresa absorbida y no una empresa activa independiente del periodo.
#
# Regla:
# - No se codifican manualmente nombres de empresas.
# - La reclasificación exige evidencia conjunta del encabezado y la nota SBS.
# - El registro especial se conserva en memoria para control y trazabilidad.
# - No entra en el panel principal empresa × mes.
# - Este bloque NO genera archivos en /salidas.
# =============================================================================


def _refs_a_set(valor):
    """Convierte '1;2' en {'1', '2'} y maneja valores vacíos."""

    if pd.isna(valor):
        return set()

    return {
        x.strip()
        for x in str(valor).split(";")
        if x.strip()
    }


def _es_registro_absorbida_transitorio(
    nombre_limpio,
    nota_original
):
    """
    Identifica una columna especial correspondiente a flujos históricos
    de una empresa absorbida.

    No utiliza nombres específicos de compañías.

    Exige conjuntamente:
    1. encabezado con marca "(abs)";
    2. nota que hable de flujos;
    3. nota que mencione una fusión;
    4. nota que indique que los flujos corresponden al periodo
       anterior a la fusión.
    """

    nombre = str(nombre_limpio)
    nota = normalizar(nota_original)

    marca_abs = bool(
        re.search(
            r"\(\s*ABS\.?\s*\)",
            nombre,
            flags=re.IGNORECASE
        )
    )

    habla_de_flujos = (
        "FLUJOS" in nota
    )

    habla_de_fusion = (
        "FUSION" in nota
    )

    habla_periodo_previo = any(
        expresion in nota
        for expresion in [
            "ANTES DE SU FUSION",
            "ANTES DE LA FUSION",
            "PREVIO A LA FUSION",
            "PERIODO PREVIO A LA FUSION",
        ]
    )

    return (
        marca_abs
        and habla_de_flujos
        and habla_de_fusion
        and habla_periodo_previo
    )


def depurar_registros_especiales(
    sbs,
    largo,
    notas,
    columnas_df,
    guardar_auditoria=True
):
    """
    Reclasifica columnas inicialmente leídas como empresas cuando la
    evidencia de la SBS demuestra que son registros transitorios de una
    empresa absorbida.

    El parámetro guardar_auditoria se conserva para compatibilidad con el
    flujo general del Script 03, pero este bloque ya no escribe archivos
    en /salidas.

    Devuelve:
        sbs_filtrado
        largo_filtrado
        comparacion_filtrada
        especiales
    """

    columnas_empresa = columnas_df[
        columnas_df[
            "tipo_columna"
        ].eq("empresa")
    ].copy()

    registros = []

    # -------------------------------------------------------------------------
    # 1. Vincular cada encabezado empresarial con su nota mediante:
    #    periodo + hoja + referencia.
    # -------------------------------------------------------------------------
    for _, col in columnas_empresa.iterrows():

        refs_columna = _refs_a_set(
            col["referencia_nota"]
        )

        if not refs_columna:
            continue

        notas_periodo = notas[
            notas["periodo"].eq(
                col["periodo"]
            )
            & notas["hoja"].eq(
                col["hoja"]
            )
        ]

        for _, nota in notas_periodo.iterrows():

            refs_nota = _refs_a_set(
                nota[
                    "referencias_detectadas"
                ]
            )

            if not (
                refs_columna
                & refs_nota
            ):
                continue

            if _es_registro_absorbida_transitorio(
                col["nombre_limpio"],
                nota["nota_original"]
            ):

                registros.append({
                    "periodo":
                        col["periodo"],

                    "hoja":
                        col["hoja"],

                    "columna_excel":
                        col["columna_excel"],

                    "encabezado_original":
                        col["encabezado_original"],

                    "empresa_sbs":
                        col["nombre_limpio"],

                    "referencia_nota":
                        col["referencia_nota"],

                    "tipo_inicial":
                        "empresa",

                    "tipo_definitivo":
                        "registro_especial_absorcion",

                    "entra_panel":
                        False,

                    "motivo": (
                        "La nota SBS identifica flujos "
                        "de una empresa absorbida "
                        "correspondientes al periodo "
                        "previo a la fusión."
                    ),

                    "nota_sbs":
                        nota["nota_original"],

                    "fila_nota_excel":
                        nota["fila_excel"],

                    "archivo_origen":
                        nota["archivo_origen"],
                })

    columnas_salida = [
        "periodo",
        "hoja",
        "columna_excel",
        "encabezado_original",
        "empresa_sbs",
        "referencia_nota",
        "tipo_inicial",
        "tipo_definitivo",
        "entra_panel",
        "motivo",
        "nota_sbs",
        "fila_nota_excel",
        "archivo_origen",
    ]

    especiales = pd.DataFrame(
        registros,
        columns=columnas_salida
    )

    # -------------------------------------------------------------------------
    # 2. Si se encontraron registros especiales, comprobar consistencia.
    # -------------------------------------------------------------------------
    if not especiales.empty:

        especiales = (
            especiales
            .drop_duplicates(
                [
                    "periodo",
                    "hoja",
                    "empresa_sbs",
                ]
            )
            .sort_values(
                [
                    "periodo",
                    "empresa_sbs",
                    "hoja",
                ]
            )
            .reset_index(
                drop=True
            )
        )

        claves = (
            especiales[
                [
                    "periodo",
                    "empresa_sbs",
                ]
            ]
            .drop_duplicates()
        )

        for _, clave in claves.iterrows():

            periodo = clave[
                "periodo"
            ]

            empresa = clave[
                "empresa_sbs"
            ]

            hojas_presentes = set(
                largo.loc[
                    largo[
                        "periodo"
                    ].eq(periodo)
                    & largo[
                        "empresa_sbs"
                    ].eq(empresa),
                    "hoja"
                ]
            )

            hojas_especiales = set(
                especiales.loc[
                    especiales[
                        "periodo"
                    ].eq(periodo)
                    & especiales[
                        "empresa_sbs"
                    ].eq(empresa),
                    "hoja"
                ]
            )

            if (
                hojas_presentes
                != hojas_especiales
            ):

                raise RuntimeError(
                    "Clasificación especial inconsistente "
                    f"para {empresa} en {periodo}. "
                    f"Hojas presentes="
                    f"{sorted(hojas_presentes)}; "
                    f"hojas clasificadas como especiales="
                    f"{sorted(hojas_especiales)}."
                )

        # ---------------------------------------------------------------------
        # 3. Incorporar los valores originales únicamente al DataFrame
        #    de control que permanece en memoria.
        # ---------------------------------------------------------------------
        valores = largo[
            [
                "periodo",
                "hoja",
                "empresa_sbs",
                "variable",
                "valor_acumulado",
                "tipo_valor",
            ]
        ].copy()

        especiales = especiales.merge(
            valores,
            on=[
                "periodo",
                "hoja",
                "empresa_sbs",
            ],
            how="left",
            validate="one_to_one"
        )

        # ---------------------------------------------------------------------
        # 4. Excluir registros especiales del panel empresa × mes.
        # ---------------------------------------------------------------------
        claves_excluir = set(
            zip(
                especiales["periodo"],
                especiales["empresa_sbs"]
            )
        )

        mascara_largo = largo.apply(
            lambda f: (
                f["periodo"],
                f["empresa_sbs"]
            ) in claves_excluir,
            axis=1
        )

        largo_filtrado = (
            largo.loc[
                ~mascara_largo
            ]
            .copy()
            .reset_index(
                drop=True
            )
        )

        mascara_sbs = sbs.apply(
            lambda f: (
                f["periodo"],
                f["empresa_sbs"]
            ) in claves_excluir,
            axis=1
        )

        sbs_filtrado = (
            sbs.loc[
                ~mascara_sbs
            ]
            .copy()
            .reset_index(
                drop=True
            )
        )

    else:

        largo_filtrado = (
            largo.copy()
        )

        sbs_filtrado = (
            sbs.copy()
        )

    # -------------------------------------------------------------------------
    # 5. Reconstruir comparación P009-P010 después de quitar especiales.
    # -------------------------------------------------------------------------
    presencia = (
        largo_filtrado
        .assign(
            presente=1
        )
        .pivot_table(
            index=[
                "periodo",
                "empresa_sbs",
            ],
            columns="variable",
            values="presente",
            aggfunc="max",
            fill_value=0
        )
        .reset_index()
    )

    presencia.columns.name = None

    for variable in [
        "primas_netas",
        "primas_cedidas",
    ]:

        if (
            variable
            not in presencia.columns
        ):
            presencia[
                variable
            ] = 0

    presencia[
        "estado_p009_p010"
    ] = np.select(
        [
            presencia[
                "primas_netas"
            ].eq(1)
            & presencia[
                "primas_cedidas"
            ].eq(1),

            presencia[
                "primas_netas"
            ].eq(1)
            & presencia[
                "primas_cedidas"
            ].eq(0),

            presencia[
                "primas_netas"
            ].eq(0)
            & presencia[
                "primas_cedidas"
            ].eq(1),
        ],
        [
            "ambas",
            "solo_P009",
            "solo_P010",
        ],
        default="ninguna"
    )

    comparacion_filtrada = presencia[
        [
            "periodo",
            "empresa_sbs",
            "estado_p009_p010",
        ]
    ].copy()

    # Reemplazar el estado anterior por el estado después
    # de la depuración de registros especiales.
    sbs_filtrado = sbs_filtrado.drop(
        columns=[
            "estado_p009_p010"
        ],
        errors="ignore"
    )

    sbs_filtrado = sbs_filtrado.merge(
        comparacion_filtrada,
        on=[
            "periodo",
            "empresa_sbs",
        ],
        how="left",
        validate="one_to_one"
    )

    # -------------------------------------------------------------------------
    # 6. Reporte en consola/log.
    #    NO se crea ningún CSV dentro de /salidas.
    # -------------------------------------------------------------------------
    if especiales.empty:

        reportar(
            "\nRegistros especiales SBS: "
            "no se detectaron columnas "
            "empresariales transitorias."
        )

    else:

        empresas_especiales = (
            especiales[
                "empresa_sbs"
            ]
            .drop_duplicates()
            .tolist()
        )

        observaciones_excluidas = len(
            especiales[
                [
                    "periodo",
                    "empresa_sbs",
                ]
            ]
            .drop_duplicates()
        )

        reportar(
            "\nRegistros especiales SBS detectados: "
            f"{len(especiales):,} registros hoja-periodo."
        )

        reportar(
            "  Empresas/columnas afectadas: "
            + "; ".join(
                empresas_especiales
            )
        )

        reportar(
            "  Observaciones empresa-mes excluidas "
            f"del panel: {observaciones_excluidas:,}."
        )

        reportar(
            "  La evidencia permanece disponible "
            "en memoria durante la ejecución; "
            "no se genera archivo en /salidas."
        )

    return (
        sbs_filtrado,
        largo_filtrado,
        comparacion_filtrada,
        especiales
    )


# =============================================================================
# BLOQUE 7. PRESENCIA Y TRAYECTORIA EMPRESARIAL
# =============================================================================
# Este bloque resume la presencia temporal de las empresas después de que
# el Bloque 8 haya construido id_empresa.
#
# IMPORTANTE:
# - Aquí NO se construye la identidad empresarial.
# - Aquí NO se guardan CSV en /salidas.
# - Los controles se mantienen en memoria y se reportan en consola/log.
# =============================================================================


def calcular_presencia(sbs, columna):
    """
    Resume la presencia temporal de cada valor de una columna.

    Calcula:
    - primer periodo;
    - último periodo;
    - meses observados;
    - meses comprendidos entre inicio y fin;
    - meses faltantes dentro de ese rango.
    """

    g = sbs.groupby(
        columna
    )["periodo"]

    presencia = pd.DataFrame({
        "primer_periodo": g.min(),
        "ultimo_periodo": g.max(),
        "meses_presente": g.nunique(),
    })

    presencia[
        "meses_en_su_rango"
    ] = (
        presencia[
            "ultimo_periodo"
        ].apply(periodo_a_numero)
        -
        presencia[
            "primer_periodo"
        ].apply(periodo_a_numero)
        + 1
    )

    presencia[
        "meses_faltantes_en_su_rango"
    ] = (
        presencia[
            "meses_en_su_rango"
        ]
        -
        presencia[
            "meses_presente"
        ]
    )

    return presencia.reset_index()


def _nombres_en_orden(grupo):
    """
    Devuelve los nombres SBS utilizados por una trayectoria,
    respetando el orden temporal y evitando repeticiones consecutivas.
    """

    grupo = grupo.sort_values(
        "periodo"
    )

    nombres = []

    for nombre in grupo[
        "empresa_sbs"
    ]:

        if (
            not nombres
            or nombres[-1] != nombre
        ):
            nombres.append(
                nombre
            )

    return " -> ".join(
        nombres
    )


def guardar_presencia_y_trayectoria(
    sbs,
    eventos
):
    """
    Realiza controles de presencia y trayectoria empresarial.

    El nombre de la función se conserva para mantener compatibilidad
    con el flujo ya construido del Script 03, pero ya NO escribe
    archivos en /salidas.

    Requiere que el Bloque 8 haya creado id_empresa.
    """

    if "id_empresa" not in sbs.columns:

        raise RuntimeError(
            "No existe id_empresa. "
            "La identidad empresarial debe construirse "
            "primero en el Bloque 8."
        )

    # -------------------------------------------------------------------------
    # 1. PRESENCIA DE LOS NOMBRES TAL COMO APARECEN EN SBS
    # -------------------------------------------------------------------------
    presencia = calcular_presencia(
        sbs,
        "empresa_sbs"
    )

    # -------------------------------------------------------------------------
    # 2. TRAYECTORIAS SEGÚN id_empresa
    # -------------------------------------------------------------------------
    filas_trayectoria = []

    for id_empresa, grupo in sbs.groupby(
        "id_empresa",
        sort=True
    ):

        grupo = grupo.sort_values(
            "periodo"
        )

        primer_periodo = (
            grupo["periodo"].min()
        )

        ultimo_periodo = (
            grupo["periodo"].max()
        )

        meses_presente = (
            grupo["periodo"].nunique()
        )

        meses_en_rango = (
            periodo_a_numero(
                ultimo_periodo
            )
            -
            periodo_a_numero(
                primer_periodo
            )
            + 1
        )

        filas_trayectoria.append({
            "id_empresa":
                id_empresa,

            "nombre_inicial_sbs":
                grupo.iloc[0][
                    "empresa_sbs"
                ],

            "nombre_final_sbs":
                grupo.iloc[-1][
                    "empresa_sbs"
                ],

            "nombres_sbs_en_trayectoria":
                _nombres_en_orden(
                    grupo
                ),

            "primer_periodo":
                primer_periodo,

            "ultimo_periodo":
                ultimo_periodo,

            "meses_presente":
                meses_presente,

            "meses_en_su_rango":
                meses_en_rango,

            "meses_faltantes_en_su_rango":
                (
                    meses_en_rango
                    - meses_presente
                ),
        })

    trayectoria = pd.DataFrame(
        filas_trayectoria
    )

    # -------------------------------------------------------------------------
    # 3. EMPRESAS ACTIVAS POR MES
    # -------------------------------------------------------------------------
    por_mes = (
        sbs.groupby(
            "periodo"
        )["id_empresa"]
        .nunique()
        .rename(
            "n_empresas"
        )
        .reset_index()
    )

    # -------------------------------------------------------------------------
    # 4. CONTROLES
    # -------------------------------------------------------------------------
    if sbs[
        [
            "periodo",
            "id_empresa"
        ]
    ].duplicated().any():

        duplicados = sbs.loc[
            sbs[
                [
                    "periodo",
                    "id_empresa"
                ]
            ].duplicated(
                keep=False
            ),
            [
                "periodo",
                "id_empresa",
                "empresa_sbs"
            ]
        ]

        raise RuntimeError(
            "Se detectaron dos observaciones para "
            "la misma identidad empresarial en un mes:\n"
            + duplicados.to_string(
                index=False
            )
        )

    reportar(
        "\nPresencia empresarial:"
    )

    reportar(
        f"  Nombres distintos observados en SBS: "
        f"{sbs['empresa_sbs'].nunique():,}."
    )

    reportar(
        f"  Identidades empresariales: "
        f"{sbs['id_empresa'].nunique():,}."
    )

    reportar(
        f"  Empresas por mes: mínimo "
        f"{por_mes['n_empresas'].min()}, "
        f"máximo "
        f"{por_mes['n_empresas'].max()}."
    )

    trayectorias_con_huecos = trayectoria[
        trayectoria[
            "meses_faltantes_en_su_rango"
        ] > 0
    ]

    reportar(
        "  Trayectorias con meses faltantes "
        f"dentro de su rango: "
        f"{len(trayectorias_con_huecos):,}."
    )

    reportar(
        "  Las tablas auxiliares permanecen "
        "en memoria; no se generan CSV en /salidas."
    )

    return (
        presencia,
        trayectoria,
        por_mes
    )



# =============================================================================
# BLOQUE 8. IDENTIDAD EMPRESARIAL Y HOMOLOGACIÓN DESDE EVIDENCIA SBS
# =============================================================================
# Objetivo:
# Detectar eventos societarios contenidos en los pies de página SBS y
# determinar cuáles implican continuidad de una trayectoria empresarial.
#
# IMPORTANTE:
# - No se inventan equivalencias empresariales.
# - Los cambios de denominación requieren continuidad temporal.
# - Las fusiones se analizan mediante la evidencia SBS.
# - Los registros "(abs)" funcionan como evidencia.
# - La construcción definitiva de id_empresa será reforzada en el Bloque 8A.
# =============================================================================


# =============================================================================
# 1. CLASIFICACIÓN DE EVENTOS
# =============================================================================

def clasificar_evento(nota_normalizada):
    """Clasifica una nota SBS según el evento societario descrito."""

    n = nota_normalizada

    if "CAMBIO DE DENOMINACION" in n:
        return "cambio_de_denominacion"

    if "FUSION" in n or "ABSORCION" in n:
        return "fusion_absorcion"

    if "LIQUIDACION" in n or "DISOLUCION" in n:
        return "liquidacion_disolucion"

    if (
        "INICIO DE OPERACIONES" in n
        or "INICIO SUS OPERACIONES" in n
    ):
        return "inicio_operaciones"

    if (
        "CESE DE OPERACIONES" in n
        or "CESE SUS OPERACIONES" in n
    ):
        return "cese_operaciones"

    return "otro"


# =============================================================================
# 2. FUNCIONES AUXILIARES
# =============================================================================

def periodo_resolucion(texto):
    """
    Extrae una fecha dd/mm/aaaa o dd-mm-aaaa de una nota SBS
    y devuelve YYYY-MM.
    """

    fecha = re.search(
        r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b",
        str(texto)
    )

    if not fecha:
        return None

    return (
        f"{fecha.group(3)}-"
        f"{int(fecha.group(2)):02d}"
    )


def periodo_anterior(periodo):
    """Devuelve el mes inmediatamente anterior."""

    return str(
        pd.Period(
            str(periodo),
            freq="M"
        ) - 1
    )


def refs_a_set(valor):
    """Convierte '1;2' en {'1', '2'}."""

    if pd.isna(valor):
        return set()

    return {
        x.strip()
        for x in str(valor).split(";")
        if x.strip()
    }


# =============================================================================
# 3. IDENTIFICACIÓN DE EMPRESAS EN LAS NOTAS
# =============================================================================

def nombres_mencionados(texto, nombres):
    """
    Busca dentro de una nota los nombres que realmente existen
    en los encabezados SBS.
    """

    texto_tokens = tokens(texto)

    encontrados = []

    for nombre in nombres:

        nt = tokens(nombre)

        if (
            nt
            and es_subsecuencia(
                nt,
                texto_tokens
            )
        ):
            encontrados.append(nombre)

    return sorted(
        dict.fromkeys(encontrados)
    )


def resolver_nombre_capturado(
    texto,
    nombres
):
    """
    Relaciona un nombre leído dentro de una nota con un nombre
    realmente observado en los encabezados SBS.
    """

    objetivo = sin_puntuacion(texto)

    exactos = [
        nombre
        for nombre in nombres
        if sin_puntuacion(nombre) == objetivo
    ]

    if len(exactos) == 1:
        return exactos[0]

    objetivo_tokens = tokens(texto)

    if not objetivo_tokens:
        return None

    candidatos = []

    for nombre in nombres:

        nt = tokens(nombre)

        if not nt:
            continue

        if (
            es_subsecuencia(
                nt,
                objetivo_tokens
            )
            or es_subsecuencia(
                objetivo_tokens,
                nt
            )
        ):
            candidatos.append(nombre)

    candidatos = sorted(
        dict.fromkeys(candidatos),
        key=lambda n: (
            len(tokens(n)),
            len(sin_puntuacion(n))
        ),
        reverse=True
    )

    if not candidatos:
        return None

    if len(candidatos) == 1:
        return candidatos[0]

    puntaje_1 = (
        len(tokens(candidatos[0])),
        len(sin_puntuacion(candidatos[0]))
    )

    puntaje_2 = (
        len(tokens(candidatos[1])),
        len(sin_puntuacion(candidatos[1]))
    )

    if puntaje_1 == puntaje_2:
        return None

    return candidatos[0]


# =============================================================================
# 4. VINCULACIÓN DE NOTAS Y ENCABEZADOS
# =============================================================================

def encabezados_referenciados(
    grupo_notas,
    columnas_df
):
    """
    Usa referencias (1), (2), etc. para identificar qué encabezados
    están relacionados con una nota.
    """

    encontrados = set()

    for _, nota in grupo_notas.iterrows():

        refs_nota = refs_a_set(
            nota["referencias_detectadas"]
        )

        if not refs_nota:
            continue

        columnas = columnas_df[
            columnas_df["periodo"].eq(
                nota["periodo"]
            )
            & columnas_df["hoja"].eq(
                nota["hoja"]
            )
        ]

        for _, columna in columnas.iterrows():

            refs_columna = refs_a_set(
                columna["referencia_nota"]
            )

            if refs_nota & refs_columna:

                encontrados.add(
                    columna["nombre_limpio"]
                )

    return sorted(encontrados)


# =============================================================================
# 5. REGISTROS ESPECIALES DE ABSORCIÓN
# =============================================================================

def nombre_base_abs(nombre):
    """
    Elimina únicamente la marca '(abs)' para comparar nombres.

    Ejemplo:
    Mapfre Perú (abs) -> MAPFRE PERU
    """

    limpio = re.sub(
        r"\s*\(\s*ABS\.?\s*\)\s*$",
        "",
        normalizar(nombre),
        flags=re.IGNORECASE
    )

    return sin_puntuacion(limpio)


def hay_absorbida_mismo_nombre(
    especiales,
    nombre_actual,
    periodo_transicion
):
    """
    Comprueba si existe un registro especial '(abs)' asociado a la
    empresa vigente durante el mes de una fusión.
    """

    if (
        especiales is None
        or especiales.empty
    ):
        return False

    candidatos = especiales[
        especiales["periodo"].eq(
            periodo_transicion
        )
        & especiales["tipo_definitivo"].eq(
            "registro_especial_absorcion"
        )
    ]

    objetivo = sin_puntuacion(
        nombre_actual
    )

    for nombre in candidatos[
        "empresa_sbs"
    ].dropna().unique():

        if (
            nombre_base_abs(nombre)
            == objetivo
        ):
            return True

    return False


# =============================================================================
# 6. EVALUACIÓN DE EVENTOS SOCIETARIOS
# =============================================================================

def evaluar_eventos(
    sbs,
    notas,
    columnas_df,
    especiales=None
):
    """
    Decide qué eventos solo se documentan y cuáles pueden modificar
    la continuidad de una identidad empresarial.
    """

    columnas_eventos = [
        "nombre_original",
        "nombre_homologado",
        "continuidad_desde",
        "continuidad_hacia",
        "tipo_evento",
        "periodo_resolucion",
        "periodo_transicion",
        "primer_periodo_nota",
        "ultimo_periodo_nota",
        "apariciones_nota",
        "hojas",
        "referencias_detectadas",
        "encabezados_referenciados",
        "aplicado",
        "rompe_mismo_nombre",
        "motivo",
        "evidencia_sbs",
        "archivo_origen",
        "fila_excel",
    ]

    if notas.empty:

        return pd.DataFrame(
            columns=columnas_eventos
        )

    nombres = sorted(
        sbs[
            "empresa_sbs"
        ].dropna().unique()
    )

    presencia = calcular_presencia(
        sbs,
        "empresa_sbs"
    ).set_index(
        "empresa_sbs"
    )

    eventos = []

    for _, grupo in notas.groupby(
        "nota_normalizada",
        sort=False
    ):

        texto = grupo[
            "nota_original"
        ].iloc[0]

        nota_norm = normalizar(
            texto
        )

        tipo = clasificar_evento(
            nota_norm
        )

        if tipo == "otro":
            continue

        periodos_nota = sorted(
            grupo[
                "periodo"
            ].dropna().unique()
        )

        refs = sorted({
            ref
            for valor in grupo[
                "referencias_detectadas"
            ].fillna("")
            for ref in refs_a_set(valor)
        })

        refs_enc = encabezados_referenciados(
            grupo,
            columnas_df
        )

        registro = {
            "nombre_original": None,
            "nombre_homologado": None,
            "continuidad_desde": None,
            "continuidad_hacia": None,
            "tipo_evento": tipo,
            "periodo_resolucion":
                periodo_resolucion(texto),
            "periodo_transicion": None,

            "primer_periodo_nota":
                min(periodos_nota)
                if periodos_nota
                else None,

            "ultimo_periodo_nota":
                max(periodos_nota)
                if periodos_nota
                else None,

            "apariciones_nota":
                len(grupo),

            "hojas":
                ";".join(
                    sorted(
                        grupo[
                            "hoja"
                        ].dropna().unique()
                    )
                ),

            "referencias_detectadas":
                ";".join(refs),

            "encabezados_referenciados":
                ";".join(refs_enc),

            "aplicado": False,
            "rompe_mismo_nombre": False,
            "motivo": "",
            "evidencia_sbs": texto,

            "archivo_origen":
                ";".join(
                    sorted(
                        grupo[
                            "archivo_origen"
                        ].dropna().unique()
                    )
                ),

            "fila_excel":
                ";".join(
                    map(
                        str,
                        sorted(
                            grupo[
                                "fila_excel"
                            ].dropna().unique()
                        )
                    )
                ),
        }

        # ---------------------------------------------------------------------
        # CAMBIO DE DENOMINACIÓN
        # ---------------------------------------------------------------------
        if tipo == "cambio_de_denominacion":

            mencionados = nombres_mencionados(
                texto,
                nombres
            )

            pares = []

            for anterior in mencionados:

                for nuevo in mencionados:

                    if anterior == nuevo:
                        continue

                    fin_anterior = presencia.loc[
                        anterior,
                        "ultimo_periodo"
                    ]

                    inicio_nuevo = presencia.loc[
                        nuevo,
                        "primer_periodo"
                    ]

                    if (
                        periodo_a_numero(
                            inicio_nuevo
                        )
                        -
                        periodo_a_numero(
                            fin_anterior
                        )
                        == 1
                    ):
                        pares.append(
                            (
                                anterior,
                                nuevo
                            )
                        )

            pares = list(
                dict.fromkeys(pares)
            )

            if len(pares) == 1:

                anterior, nuevo = pares[0]

                registro[
                    "nombre_original"
                ] = anterior

                registro[
                    "nombre_homologado"
                ] = nuevo

                registro[
                    "continuidad_desde"
                ] = anterior

                registro[
                    "continuidad_hacia"
                ] = nuevo

                registro[
                    "periodo_transicion"
                ] = presencia.loc[
                    nuevo,
                    "primer_periodo"
                ]

                registro[
                    "aplicado"
                ] = True

                registro[
                    "motivo"
                ] = (
                    "Cambio de denominación documentado "
                    "por SBS y continuidad temporal exacta."
                )

            elif len(pares) == 0:

                registro[
                    "motivo"
                ] = (
                    "La nota indica cambio de denominación, "
                    "pero no existe una pareja única de nombres "
                    "con continuidad temporal exacta."
                )

            else:

                registro[
                    "motivo"
                ] = (
                    "Cambio de denominación ambiguo: "
                    "más de una pareja temporal posible."
                )

            eventos.append(
                registro
            )

            continue

        # ---------------------------------------------------------------------
        # FUSIÓN / ABSORCIÓN
        # ---------------------------------------------------------------------
        if tipo == "fusion_absorcion":

            texto_limpio = sin_puntuacion(
                texto
            )

            patron = re.search(
                r"LAS CIFRAS DE (.+?) "
                r"CONSIDERAN LOS FLUJOS DE (.+?) "
                r"PARA EL PERIODO PREVIO A LA FUSION",
                texto_limpio
            )

            if not patron:

                registro[
                    "motivo"
                ] = (
                    "Fusión/absorción documentada por SBS; "
                    "se registra, pero no se fuerza continuidad."
                )

                eventos.append(
                    registro
                )

                continue

            actual = resolver_nombre_capturado(
                patron.group(1),
                nombres
            )

            previo = resolver_nombre_capturado(
                patron.group(2),
                nombres
            )

            transicion = (
                min(periodos_nota)
                if periodos_nota
                else None
            )

            if (
                not actual
                or not previo
                or not transicion
            ):

                registro[
                    "motivo"
                ] = (
                    "No fue posible resolver de forma única "
                    "las empresas mencionadas en la nota de fusión."
                )

                eventos.append(
                    registro
                )

                continue

            mes_previo = periodo_anterior(
                transicion
            )

            existe_actual = (
                sbs["periodo"].eq(
                    transicion
                )
                & sbs[
                    "empresa_sbs"
                ].eq(actual)
            ).any()

            existe_previo = (
                sbs["periodo"].eq(
                    mes_previo
                )
                & sbs[
                    "empresa_sbs"
                ].eq(previo)
            ).any()

            existe_actual_mes_previo = (
                sbs["periodo"].eq(
                    mes_previo
                )
                & sbs[
                    "empresa_sbs"
                ].eq(actual)
            ).any()

            evidencia_abs = (
                hay_absorbida_mismo_nombre(
                    especiales,
                    actual,
                    transicion
                )
            )

            if (
                existe_previo
                and existe_actual
                and evidencia_abs
            ):

                registro[
                    "nombre_original"
                ] = previo

                registro[
                    "nombre_homologado"
                ] = actual

                registro[
                    "continuidad_desde"
                ] = previo

                registro[
                    "continuidad_hacia"
                ] = actual

                registro[
                    "periodo_transicion"
                ] = transicion

                registro[
                    "aplicado"
                ] = True

                registro[
                    "rompe_mismo_nombre"
                ] = bool(
                    existe_actual_mes_previo
                )

                registro[
                    "motivo"
                ] = (
                    "Fusión con continuidad explícita respaldada "
                    "por SBS: la empresa vigente incorpora los "
                    "flujos de la entidad previa y existe registro "
                    "especial '(abs)' de la absorbida."
                )

            else:

                registro[
                    "motivo"
                ] = (
                    "La nota sugiere continuidad por fusión, "
                    "pero faltan condiciones de verificación "
                    "temporal o evidencia '(abs)'. No se aplica."
                )

            eventos.append(
                registro
            )

            continue

        # ---------------------------------------------------------------------
        # OTROS EVENTOS
        # ---------------------------------------------------------------------
        registro[
            "motivo"
        ] = (
            "Evento societario documentado por SBS; "
            "no modifica automáticamente la identidad."
        )

        eventos.append(
            registro
        )

    return pd.DataFrame(
        eventos,
        columns=columnas_eventos
    )


# =============================================================================
# 7. CONSTRUCCIÓN INICIAL DEL id_empresa
# =============================================================================
# Esta versión será reemplazada/reforzada por el Bloque 8A.
# =============================================================================

def construir_id_empresa(
    sbs,
    eventos
):
    """Construye una identidad estable para cada trayectoria empresarial."""

    p = (
        sbs.copy()
        .sort_values(
            [
                "periodo",
                "empresa_sbs"
            ]
        )
        .reset_index(drop=True)
    )

    claves = list(
        zip(
            p["periodo"],
            p["empresa_sbs"]
        )
    )

    if len(claves) != len(set(claves)):

        raise RuntimeError(
            "Hay duplicados periodo × empresa_sbs "
            "antes de construir id_empresa."
        )

    padre = {
        clave: clave
        for clave in claves
    }

    def buscar(x):

        while padre[x] != x:

            padre[x] = padre[
                padre[x]
            ]

            x = padre[x]

        return x

    def unir(a, b):

        raiz_a = buscar(a)
        raiz_b = buscar(b)

        if raiz_a != raiz_b:
            padre[raiz_b] = raiz_a

    # -------------------------------------------------------------------------
    # Rupturas documentadas del mismo nombre
    # -------------------------------------------------------------------------
    rupturas = set()

    if (
        eventos is not None
        and not eventos.empty
    ):

        aplicados = eventos[
            eventos[
                "aplicado"
            ].eq(True)
        ]

        for _, evento in aplicados.iterrows():

            if bool(
                evento[
                    "rompe_mismo_nombre"
                ]
            ):

                rupturas.add(
                    (
                        evento[
                            "continuidad_hacia"
                        ],
                        evento[
                            "periodo_transicion"
                        ]
                    )
                )

    # -------------------------------------------------------------------------
    # Continuidad del mismo nombre
    # -------------------------------------------------------------------------
    for nombre, grupo in p.groupby(
        "empresa_sbs",
        sort=True
    ):

        periodos = (
            grupo
            .sort_values(
                "periodo"
            )[
                "periodo"
            ]
            .tolist()
        )

        for anterior, actual in zip(
            periodos[:-1],
            periodos[1:]
        ):

            if (
                nombre,
                actual
            ) in rupturas:

                continue

            unir(
                (
                    anterior,
                    nombre
                ),
                (
                    actual,
                    nombre
                )
            )

    # -------------------------------------------------------------------------
    # Continuidades diferentes respaldadas por SBS
    # -------------------------------------------------------------------------
    if (
        eventos is not None
        and not eventos.empty
    ):

        aplicados = eventos[
            eventos[
                "aplicado"
            ].eq(True)
        ]

        for _, evento in aplicados.iterrows():

            origen = evento[
                "continuidad_desde"
            ]

            destino = evento[
                "continuidad_hacia"
            ]

            transicion = evento[
                "periodo_transicion"
            ]

            if (
                pd.isna(origen)
                or pd.isna(destino)
                or pd.isna(transicion)
            ):
                continue

            mes_previo = periodo_anterior(
                transicion
            )

            clave_origen = (
                mes_previo,
                origen
            )

            clave_destino = (
                transicion,
                destino
            )

            if clave_origen not in padre:

                raise RuntimeError(
                    "La continuidad documentada no encuentra "
                    f"{origen} en {mes_previo}."
                )

            if clave_destino not in padre:

                raise RuntimeError(
                    "La continuidad documentada no encuentra "
                    f"{destino} en {transicion}."
                )

            unir(
                clave_origen,
                clave_destino
            )

    # -------------------------------------------------------------------------
    # Componentes
    # -------------------------------------------------------------------------
    componentes = {}

    for clave in claves:

        raiz = buscar(
            clave
        )

        componentes.setdefault(
            raiz,
            []
        ).append(
            clave
        )

    def criterio_raiz(raiz):

        nodos = sorted(
            componentes[raiz],
            key=lambda x: (
                periodo_a_numero(
                    x[0]
                ),
                sin_puntuacion(
                    x[1]
                )
            )
        )

        return (
            periodo_a_numero(
                nodos[0][0]
            ),
            sin_puntuacion(
                nodos[0][1]
            ),
            sin_puntuacion(
                nodos[-1][1]
            )
        )

    raices = sorted(
        componentes,
        key=criterio_raiz
    )

    id_por_raiz = {
        raiz: f"E{i:03d}"
        for i, raiz in enumerate(
            raices,
            1
        )
    }

    p["id_empresa"] = [
        id_por_raiz[
            buscar(
                (
                    fila.periodo,
                    fila.empresa_sbs
                )
            )
        ]
        for fila in p.itertuples()
    ]

    # Nombre observado del mes.
    # El Bloque 8A lo convertirá en nombre analítico estable.
    p["empresa"] = p[
        "empresa_sbs"
    ]

    return p


# =============================================================================
# 8. CONTROLES DE IDENTIDAD
# =============================================================================

def controlar_identidades(
    sbs,
    eventos
):
    """Comprueba que los IDs empresariales sean coherentes."""

    if sbs[
        "id_empresa"
    ].isna().any():

        raise RuntimeError(
            "Hay filas sin id_empresa."
        )

    duplicados = sbs.duplicated(
        [
            "periodo",
            "id_empresa"
        ]
    )

    if duplicados.any():

        ejemplo = sbs.loc[
            duplicados,
            [
                "periodo",
                "id_empresa",
                "empresa_sbs",
                "empresa"
            ]
        ].head(20)

        raise RuntimeError(
            "Una identidad empresarial aparece "
            "más de una vez en el mismo mes:\n"
            + ejemplo.to_string(
                index=False
            )
        )

    if (
        eventos is None
        or eventos.empty
    ):
        return

    aplicados = eventos[
        eventos[
            "aplicado"
        ].eq(True)
    ]

    for _, evento in aplicados.iterrows():

        transicion = evento[
            "periodo_transicion"
        ]

        origen = evento[
            "continuidad_desde"
        ]

        destino = evento[
            "continuidad_hacia"
        ]

        mes_previo = periodo_anterior(
            transicion
        )

        id_origen = sbs.loc[
            sbs["periodo"].eq(
                mes_previo
            )
            & sbs[
                "empresa_sbs"
            ].eq(origen),
            "id_empresa"
        ].unique()

        id_destino = sbs.loc[
            sbs["periodo"].eq(
                transicion
            )
            & sbs[
                "empresa_sbs"
            ].eq(destino),
            "id_empresa"
        ].unique()

        if (
            len(id_origen) != 1
            or len(id_destino) != 1
            or id_origen[0] != id_destino[0]
        ):

            raise RuntimeError(
                "Una continuidad SBS aplicada "
                "no conserva el mismo id_empresa."
            )

        # ---------------------------------------------------------------------
        # Una fusión puede obligar a separar la trayectoria anterior
        # de una empresa con el mismo nombre.
        # ---------------------------------------------------------------------
        if bool(
            evento[
                "rompe_mismo_nombre"
            ]
        ):

            id_nombre_anterior = (
                sbs.loc[
                    sbs[
                        "periodo"
                    ].eq(mes_previo)
                    & sbs[
                        "empresa_sbs"
                    ].eq(destino),
                    "id_empresa"
                ]
                .unique()
            )

            if (
                len(
                    id_nombre_anterior
                ) == 1
                and id_nombre_anterior[0]
                == id_destino[0]
            ):

                raise RuntimeError(
                    "Una fusión que debía romper "
                    "la trayectoria previa del mismo nombre "
                    "terminó usando el mismo ID."
                )


# =============================================================================
# 9. AUDITORÍA DE EVENTOS
# =============================================================================
# Esta función será reemplazada por el Bloque 8A para que no genere
# archivos auxiliares dentro de /salidas.
# =============================================================================

def guardar_eventos_identidad(eventos):
    """Guarda temporalmente la auditoría de eventos societarios."""

    eventos.to_csv(
        CARPETA_SALIDAS
        / "homologacion_empresas.csv",
        index=False,
        encoding="utf-8"
    )


# =============================================================================
# 10. FUNCIÓN PRINCIPAL
# =============================================================================

def homologar(
    sbs,
    notas,
    columnas_df,
    especiales=None,
    guardar_auditoria=True
):
    """Ejecuta el proceso inicial de identidad empresarial."""

    eventos = evaluar_eventos(
        sbs,
        notas,
        columnas_df,
        especiales=especiales
    )

    sbs = construir_id_empresa(
        sbs,
        eventos
    )

    controlar_identidades(
        sbs,
        eventos
    )

    if guardar_auditoria:

        guardar_eventos_identidad(
            eventos
        )

    aplicados = (
        eventos[
            eventos[
                "aplicado"
            ].eq(True)
        ]
        if not eventos.empty
        else eventos
    )

    reportar(
        "\nEventos societarios detectados: "
        f"{len(eventos):,}; "
        "continuidades de identidad aplicadas: "
        f"{len(aplicados):,}."
    )

    if not eventos.empty:

        for tipo, cantidad in (
            eventos[
                "tipo_evento"
            ]
            .value_counts()
            .items()
        ):

            reportar(
                f"  {tipo}: {cantidad}"
            )

    if not aplicados.empty:

        reportar(
            "\nContinuidades empresariales "
            "respaldadas por SBS:"
        )

        for _, evento in aplicados.iterrows():

            reportar(
                f"  {evento['continuidad_desde']} "
                "-> "
                f"{evento['continuidad_hacia']} "
                "| transición "
                f"{evento['periodo_transicion']} "
                "| "
                f"{evento['motivo']}"
            )

    reportar(
        "\nNombres SBS distintos: "
        f"{sbs['empresa_sbs'].nunique()} "
        "| identidades empresariales: "
        f"{sbs['id_empresa'].nunique()}."
    )

    return (
        sbs,
        eventos
    )



# =============================================================================
# BLOQUE 8A. IDENTIDAD EMPRESARIAL DEFINITIVA
# =============================================================================
# Complementa al Bloque 8.
#
# Objetivos:
# - mantener una identidad estable a través del tiempo;
# - reconocer variantes puramente ortográficas del mismo nombre;
# - respetar rupturas documentadas por fusiones/absorciones;
# - aplicar continuidades societarias respaldadas por SBS;
# - crear un nombre analítico estable para cada trayectoria.
#
# empresa_sbs:
#     nombre observado originalmente en la fuente SBS.
#
# empresa:
#     nombre analítico estable que se utilizará posteriormente.
#
# Este bloque NO genera archivos en /salidas.
# =============================================================================


def construir_id_empresa(
    sbs,
    eventos
):
    """
    Construye un id_empresa estable para cada trayectoria empresarial.

    Se consideran conjuntamente:
    1. continuidad temporal;
    2. equivalencia ortográfica del nombre;
    3. eventos societarios validados por SBS;
    4. rupturas documentadas de identidad.
    """

    p = (
        sbs.copy()
        .sort_values(
            [
                "periodo",
                "empresa_sbs"
            ]
        )
        .reset_index(
            drop=True
        )
    )

    # =========================================================================
    # 1. CONTROL INICIAL
    # =========================================================================

    claves = list(
        zip(
            p["periodo"],
            p["empresa_sbs"]
        )
    )

    if len(claves) != len(set(claves)):

        duplicados = p.loc[
            p.duplicated(
                [
                    "periodo",
                    "empresa_sbs"
                ],
                keep=False
            ),
            [
                "periodo",
                "empresa_sbs"
            ]
        ]

        raise RuntimeError(
            "Hay duplicados periodo × empresa_sbs "
            "antes de construir id_empresa:\n"
            + duplicados.to_string(
                index=False
            )
        )

    # =========================================================================
    # 2. CLAVE ORTOGRÁFICA
    # =========================================================================
    # Permite reconocer diferencias que no representan una empresa distinta.
    #
    # Ejemplo:
    #
    # Qualitas
    # Quálitas
    #
    # ambas producen la misma clave normalizada.
    # =========================================================================

    p["_clave_ortografica"] = (
        p["empresa_sbs"]
        .apply(
            sin_puntuacion
        )
    )

    # -------------------------------------------------------------------------
    # Control:
    # dos variantes de una misma clave no pueden coexistir simultáneamente.
    # -------------------------------------------------------------------------

    colisiones = (
        p.groupby(
            [
                "periodo",
                "_clave_ortografica"
            ]
        )[
            "empresa_sbs"
        ]
        .nunique()
        .reset_index(
            name="n_nombres"
        )
    )

    colisiones = colisiones[
        colisiones[
            "n_nombres"
        ] > 1
    ]

    if not colisiones.empty:

        detalle = p.merge(
            colisiones[
                [
                    "periodo",
                    "_clave_ortografica"
                ]
            ],
            on=[
                "periodo",
                "_clave_ortografica"
            ],
            how="inner"
        )[
            [
                "periodo",
                "empresa_sbs",
                "_clave_ortografica"
            ]
        ]

        raise RuntimeError(
            "Se encontraron variantes ortográficas "
            "simultáneas de una misma empresa:\n"
            + detalle.to_string(
                index=False
            )
        )

    # =========================================================================
    # 3. ESTRUCTURA UNION-FIND
    # =========================================================================

    padre = {
        clave: clave
        for clave in claves
    }

    def buscar(x):

        while padre[x] != x:

            padre[x] = padre[
                padre[x]
            ]

            x = padre[x]

        return x

    def unir(a, b):

        raiz_a = buscar(a)
        raiz_b = buscar(b)

        if raiz_a != raiz_b:

            padre[
                raiz_b
            ] = raiz_a

    # =========================================================================
    # 4. RUPTURAS DOCUMENTADAS
    # =========================================================================
    # Caso fundamental:
    #
    # MAPFRE Perú antes de la fusión
    #
    # no debe confundirse con:
    #
    # MAPFRE Perú después de la fusión,
    #
    # cuando la propia evidencia SBS establece que la continuidad posterior
    # proviene de MAPFRE Perú Vida.
    # =========================================================================

    rupturas = set()

    if (
        eventos is not None
        and not eventos.empty
        and "aplicado" in eventos.columns
    ):

        aplicados = eventos[
            eventos[
                "aplicado"
            ].eq(True)
        ]

        for _, evento in aplicados.iterrows():

            rompe = evento.get(
                "rompe_mismo_nombre",
                False
            )

            if pd.isna(rompe):
                rompe = False

            if not bool(rompe):
                continue

            destino = evento.get(
                "continuidad_hacia"
            )

            transicion = evento.get(
                "periodo_transicion"
            )

            if (
                pd.isna(destino)
                or pd.isna(transicion)
            ):
                continue

            rupturas.add(
                (
                    sin_puntuacion(
                        destino
                    ),
                    str(
                        transicion
                    )
                )
            )

    # =========================================================================
    # 5. CONTINUIDAD POR MISMA CLAVE ORTOGRÁFICA
    # =========================================================================

    for clave_ortografica, grupo in p.groupby(
        "_clave_ortografica",
        sort=True
    ):

        grupo = grupo.sort_values(
            "periodo"
        )

        filas = list(
            grupo[
                [
                    "periodo",
                    "empresa_sbs"
                ]
            ].itertuples(
                index=False,
                name=None
            )
        )

        for anterior, actual in zip(
            filas[:-1],
            filas[1:]
        ):

            periodo_actual = (
                actual[0]
            )

            # Si SBS documentó una ruptura,
            # las dos trayectorias no se unen.
            if (
                clave_ortografica,
                periodo_actual
            ) in rupturas:

                continue

            unir(
                anterior,
                actual
            )

    # =========================================================================
    # 6. CONTINUIDADES SOCIETARIAS VALIDADAS
    # =========================================================================

    if (
        eventos is not None
        and not eventos.empty
        and "aplicado" in eventos.columns
    ):

        aplicados = eventos[
            eventos[
                "aplicado"
            ].eq(True)
        ]

        for _, evento in aplicados.iterrows():

            origen = evento.get(
                "continuidad_desde"
            )

            destino = evento.get(
                "continuidad_hacia"
            )

            transicion = evento.get(
                "periodo_transicion"
            )

            if (
                pd.isna(origen)
                or pd.isna(destino)
                or pd.isna(transicion)
            ):
                continue

            transicion = str(
                transicion
            )

            mes_previo = periodo_anterior(
                transicion
            )

            clave_origen = (
                mes_previo,
                origen
            )

            clave_destino = (
                transicion,
                destino
            )

            if clave_origen not in padre:

                raise RuntimeError(
                    "La continuidad societaria validada "
                    f"no encuentra {origen} "
                    f"en {mes_previo}."
                )

            if clave_destino not in padre:

                raise RuntimeError(
                    "La continuidad societaria validada "
                    f"no encuentra {destino} "
                    f"en {transicion}."
                )

            unir(
                clave_origen,
                clave_destino
            )

    # =========================================================================
    # 7. FORMAR COMPONENTES DE IDENTIDAD
    # =========================================================================

    componentes = {}

    for clave in claves:

        raiz = buscar(
            clave
        )

        componentes.setdefault(
            raiz,
            []
        ).append(
            clave
        )

    # =========================================================================
    # 8. IDs REPRODUCIBLES
    # =========================================================================

    def criterio_raiz(raiz):

        nodos = sorted(
            componentes[
                raiz
            ],
            key=lambda x: (
                periodo_a_numero(
                    x[0]
                ),
                sin_puntuacion(
                    x[1]
                )
            )
        )

        return (
            periodo_a_numero(
                nodos[0][0]
            ),
            sin_puntuacion(
                nodos[0][1]
            ),
            sin_puntuacion(
                nodos[-1][1]
            )
        )

    raices = sorted(
        componentes,
        key=criterio_raiz
    )

    id_por_raiz = {
        raiz: f"E{i:03d}"
        for i, raiz in enumerate(
            raices,
            1
        )
    }

    p["id_empresa"] = [
        id_por_raiz[
            buscar(
                (
                    fila.periodo,
                    fila.empresa_sbs
                )
            )
        ]
        for fila in p.itertuples()
    ]

    # =========================================================================
    # 9. NOMBRE ANALÍTICO ESTABLE
    # =========================================================================
    # Por defecto se utiliza la última denominación observada de cada
    # trayectoria.
    #
    # Ejemplos:
    #
    # Vida Cámara -> Vivir Seguros
    # Secrex -> Cesce Perú
    # Ohio National Vida -> AuguStar
    # Qualitas -> Quálitas
    # =========================================================================

    nombre_estable = {}

    for id_empresa, grupo in p.groupby(
        "id_empresa",
        sort=False
    ):

        grupo = grupo.sort_values(
            "periodo"
        )

        nombre_estable[
            id_empresa
        ] = grupo.iloc[-1][
            "empresa_sbs"
        ]

    p["empresa"] = (
        p["id_empresa"]
        .map(
            nombre_estable
        )
    )

    # =========================================================================
    # 9B. ETIQUETAR CAMBIOS DE DENOMINACIÓN
    # =========================================================================
    # Si SBS confirma que la misma entidad cambió de denominación,
    # conserva el mismo id_empresa y utiliza como etiqueta:
    #
    # Nombre actual (anteriormente Nombre anterior)
    # =========================================================================

    if (
        eventos is not None
        and not eventos.empty
        and "aplicado" in eventos.columns
    ):

        cambios_nombre = eventos[
            eventos["aplicado"].eq(True)
            &
            eventos["tipo_evento"].eq(
                "cambio_de_denominacion"
            )
        ]

        for _, evento in cambios_nombre.iterrows():

            origen = evento.get(
                "continuidad_desde"
            )

            destino = evento.get(
                "continuidad_hacia"
            )

            transicion = evento.get(
                "periodo_transicion"
            )

            if (
                pd.isna(origen)
                or pd.isna(destino)
                or pd.isna(transicion)
            ):
                continue

            ids = p.loc[
                p["periodo"].eq(
                    str(transicion)
                )
                &
                p["empresa_sbs"].eq(
                    destino
                ),
                "id_empresa"
            ].unique()

            if len(ids) == 1:

                p.loc[
                    p["id_empresa"].eq(
                        ids[0]
                    ),
                    "empresa"
                ] = (
                    f"{destino} "
                    f"(anteriormente {origen})"
                )
    # =========================================================================
    # 10. ETIQUETAS ESPECIALES DE FUSIÓN
    # =========================================================================
    # Cuando existe ruptura del mismo nombre:
    #
    # trayectoria antigua:
    #     Mapfre Perú (pre-fusión)
    #
    # trayectoria que continúa:
    #     Mapfre Perú (antes Mapfre Perú Vida)
    # =========================================================================

    if (
        eventos is not None
        and not eventos.empty
        and "aplicado" in eventos.columns
    ):

        aplicados = eventos[
            eventos[
                "aplicado"
            ].eq(True)
        ]

        for _, evento in aplicados.iterrows():

            rompe = evento.get(
                "rompe_mismo_nombre",
                False
            )

            if pd.isna(rompe):
                rompe = False

            if not bool(rompe):
                continue

            origen = evento.get(
                "continuidad_desde"
            )

            destino = evento.get(
                "continuidad_hacia"
            )

            transicion = evento.get(
                "periodo_transicion"
            )

            if (
                pd.isna(origen)
                or pd.isna(destino)
                or pd.isna(transicion)
            ):
                continue

            transicion = str(
                transicion
            )

            mes_previo = periodo_anterior(
                transicion
            )

            # -----------------------------------------------------------------
            # Identidad que continúa después de la fusión.
            # -----------------------------------------------------------------

            id_continua = p.loc[
                p["periodo"].eq(
                    transicion
                )
                & p[
                    "empresa_sbs"
                ].eq(
                    destino
                ),
                "id_empresa"
            ].unique()

            if len(id_continua) == 1:

                p.loc[
                    p[
                        "id_empresa"
                    ].eq(
                        id_continua[0]
                    ),
                    "empresa"
                ] = (
                    f"{destino} "
                    f"(antes {origen})"
                )

            # -----------------------------------------------------------------
            # Trayectoria previa que ya tenía el nombre del destino.
            # -----------------------------------------------------------------

            id_previo = p.loc[
                p["periodo"].eq(
                    mes_previo
                )
                & p[
                    "empresa_sbs"
                ].eq(
                    destino
                ),
                "id_empresa"
            ].unique()

            if len(id_previo) == 1:

                p.loc[
                    p[
                        "id_empresa"
                    ].eq(
                        id_previo[0]
                    ),
                    "empresa"
                ] = (
                    f"{destino} "
                    "(pre-fusión)"
                )

    # =========================================================================
    # 11. CONTROLES FINALES
    # =========================================================================

    if p[
        "id_empresa"
    ].isna().any():

        raise RuntimeError(
            "Hay observaciones sin id_empresa."
        )

    if p[
        "empresa"
    ].isna().any():

        raise RuntimeError(
            "Hay observaciones sin nombre analítico estable."
        )

    duplicados_id = p.duplicated(
        [
            "periodo",
            "id_empresa"
        ]
    )

    if duplicados_id.any():

        detalle = p.loc[
            duplicados_id,
            [
                "periodo",
                "id_empresa",
                "empresa_sbs",
                "empresa"
            ]
        ]

        raise RuntimeError(
            "Una identidad empresarial aparece "
            "más de una vez en el mismo periodo:\n"
            + detalle.to_string(
                index=False
            )
        )

    p = p.drop(
        columns=[
            "_clave_ortografica"
        ]
    )

    return p


# =============================================================================
# 12. AUDITORÍA EN MEMORIA
# =============================================================================

def guardar_eventos_identidad(
    eventos
):
    """
    Conserva el control de eventos dentro de la ejecución.

    No genera homologacion_empresas.csv ni otros archivos en /salidas.
    """

    if eventos is None:

        reportar(
            "Eventos societarios: "
            "sin información para controlar."
        )

        return

    reportar(
        "Eventos societarios: "
        f"{len(eventos):,} registros "
        "conservados en memoria para control."
    )



# =============================================================================
# BLOQUE 8B. VALIDACIÓN Y CONSOLIDACIÓN DE EVENTOS SOCIETARIOS
# =============================================================================
# Complementa los Bloques 8 y 8A.
#
# Funciones:
# 1. Validar las continuidades propuestas.
# 2. Rechazar continuidades inválidas A -> A.
# 3. Exigir presencia real de origen y destino.
# 4. Exigir continuidad temporal t -> t+1.
# 5. Consolidar notas duplicadas del mismo evento.
# 6. Construir finalmente id_empresa mediante la función del Bloque 8A.
#
# No genera archivos en /salidas.
# =============================================================================


# =============================================================================
# 1. VALIDAR UNA CONTINUIDAD PROPUESTA
# =============================================================================

def validar_continuidad_evento(evento, sbs):
    """
    Revisa si una continuidad propuesta puede aplicarse realmente
    dentro de la ventana 2020-01 a 2025-12.
    """

    evento = evento.copy()

    if not bool(
        evento.get(
            "aplicado",
            False
        )
    ):
        return evento

    origen = evento.get(
        "continuidad_desde"
    )

    destino = evento.get(
        "continuidad_hacia"
    )

    transicion = evento.get(
        "periodo_transicion"
    )

    # -------------------------------------------------------------------------
    # 1A. Origen, destino y periodo deben existir.
    # -------------------------------------------------------------------------

    if (
        pd.isna(origen)
        or pd.isna(destino)
        or pd.isna(transicion)
        or not str(origen).strip()
        or not str(destino).strip()
        or not str(transicion).strip()
    ):

        evento[
            "aplicado"
        ] = False

        evento[
            "motivo"
        ] = (
            "Evento no aplicado: faltan origen, destino "
            "o periodo de transición verificables."
        )

        return evento

    # -------------------------------------------------------------------------
    # 1B. No aceptar A -> A.
    #
    # También evita tratar como evento societario una diferencia
    # meramente ortográfica como Qualitas -> Quálitas.
    # Esa continuidad se resuelve mediante la clave ortográfica del 8A.
    # -------------------------------------------------------------------------

    if (
        sin_puntuacion(origen)
        ==
        sin_puntuacion(destino)
    ):

        evento[
            "aplicado"
        ] = False

        evento[
            "motivo"
        ] = (
            "Evento documentado, pero no aplicado como transición "
            "societaria: origen y destino tienen la misma clave "
            "normalizada. La continuidad ortográfica se trata por separado."
        )

        return evento

    # -------------------------------------------------------------------------
    # 1C. Ambos nombres deben estar realmente observados.
    # -------------------------------------------------------------------------

    existe_origen = (
        sbs[
            "empresa_sbs"
        ]
        .eq(origen)
        .any()
    )

    existe_destino = (
        sbs[
            "empresa_sbs"
        ]
        .eq(destino)
        .any()
    )

    if (
        not existe_origen
        or not existe_destino
    ):

        evento[
            "aplicado"
        ] = False

        evento[
            "motivo"
        ] = (
            "Evento societario documentado, pero no aplicado: "
            "el nombre anterior o posterior no está observado "
            "dentro de la ventana 2020-2025."
        )

        return evento

    # -------------------------------------------------------------------------
    # 1D. Verificar continuidad t -> t+1.
    # -------------------------------------------------------------------------

    transicion = str(
        transicion
    )

    mes_previo = periodo_anterior(
        transicion
    )

    origen_previo = (
        sbs[
            "periodo"
        ].eq(mes_previo)
        &
        sbs[
            "empresa_sbs"
        ].eq(origen)
    ).any()

    destino_actual = (
        sbs[
            "periodo"
        ].eq(transicion)
        &
        sbs[
            "empresa_sbs"
        ].eq(destino)
    ).any()

    if (
        not origen_previo
        or not destino_actual
    ):

        evento[
            "aplicado"
        ] = False

        evento[
            "motivo"
        ] = (
            "Evento societario documentado, pero no aplicado: "
            "no se verifica origen en el mes anterior "
            "y destino en el mes de transición."
        )

        return evento

    return evento


# =============================================================================
# 2. VALIDAR TODOS LOS EVENTOS
# =============================================================================

def validar_eventos_aplicados(
    eventos,
    sbs
):
    """Aplica los controles anteriores a todos los eventos detectados."""

    if (
        eventos is None
        or eventos.empty
    ):

        return eventos.copy()

    revisados = []

    for _, evento in eventos.iterrows():

        revisados.append(
            validar_continuidad_evento(
                evento,
                sbs
            )
        )

    return pd.DataFrame(
        revisados,
        columns=eventos.columns
    )


# =============================================================================
# 3. CONSOLIDAR EVENTOS DUPLICADOS
# =============================================================================

def consolidar_eventos(eventos):
    """
    Consolida varias notas SBS que describen una misma transición.

    Ejemplo:
    Secrex -> Cesce Perú puede aparecer en más de una nota, pero
    representa una sola continuidad empresarial.
    """

    if (
        eventos is None
        or eventos.empty
    ):

        return eventos.copy()

    no_aplicados = eventos[
        ~eventos[
            "aplicado"
        ].eq(True)
    ].copy()

    aplicados = eventos[
        eventos[
            "aplicado"
        ].eq(True)
    ].copy()

    if aplicados.empty:

        return eventos.copy()

    clave = [
        "tipo_evento",
        "continuidad_desde",
        "continuidad_hacia",
        "periodo_transicion",
    ]

    consolidados = []

    for _, grupo in aplicados.groupby(
        clave,
        dropna=False,
        sort=False
    ):

        fila = (
            grupo.iloc[0]
            .copy()
        )

        # ---------------------------------------------------------------------
        # Combinar evidencia SBS.
        # ---------------------------------------------------------------------

        evidencias = list(
            dict.fromkeys(
                grupo[
                    "evidencia_sbs"
                ]
                .dropna()
                .astype(str)
                .tolist()
            )
        )

        fila[
            "evidencia_sbs"
        ] = " || ".join(
            evidencias
        )

        # ---------------------------------------------------------------------
        # Combinar archivos de origen.
        # ---------------------------------------------------------------------

        archivos = []

        for valor in grupo[
            "archivo_origen"
        ].dropna():

            for archivo in str(
                valor
            ).split(";"):

                archivo = (
                    archivo.strip()
                )

                if (
                    archivo
                    and archivo not in archivos
                ):
                    archivos.append(
                        archivo
                    )

        fila[
            "archivo_origen"
        ] = ";".join(
            archivos
        )

        # ---------------------------------------------------------------------
        # Combinar referencias.
        # ---------------------------------------------------------------------

        referencias = []

        for valor in grupo[
            "referencias_detectadas"
        ].fillna(""):

            for ref in str(
                valor
            ).split(";"):

                ref = ref.strip()

                if (
                    ref
                    and ref not in referencias
                ):
                    referencias.append(
                        ref
                    )

        fila[
            "referencias_detectadas"
        ] = ";".join(
            referencias
        )

        # Número total de notas que respaldan el evento.
        fila[
            "apariciones_nota"
        ] = int(
            grupo[
                "apariciones_nota"
            ]
            .fillna(0)
            .sum()
        )

        consolidados.append(
            fila
        )

    aplicados_consolidados = pd.DataFrame(
        consolidados,
        columns=eventos.columns
    )

    salida = pd.concat(
        [
            no_aplicados,
            aplicados_consolidados
        ],
        ignore_index=True
    )

    return salida


# =============================================================================
# 4. CONTROL FINAL DE EVENTOS
# =============================================================================

def control_final_eventos(eventos):
    """
    Impide que sobrevivan errores lógicos entre las continuidades aplicadas.
    """

    if (
        eventos is None
        or eventos.empty
    ):
        return

    aplicados = eventos[
        eventos[
            "aplicado"
        ].eq(True)
    ]

    if aplicados.empty:
        return

    # -------------------------------------------------------------------------
    # 4A. Nunca puede sobrevivir una continuidad A -> A.
    # -------------------------------------------------------------------------

    mismo_nombre = aplicados.apply(
        lambda fila:
        sin_puntuacion(
            fila[
                "continuidad_desde"
            ]
        )
        ==
        sin_puntuacion(
            fila[
                "continuidad_hacia"
            ]
        ),
        axis=1
    )

    if mismo_nombre.any():

        raise RuntimeError(
            "Existe una continuidad inválida "
            "del tipo empresa -> misma empresa."
        )

    # -------------------------------------------------------------------------
    # 4B. No puede quedar duplicada una misma transición.
    # -------------------------------------------------------------------------

    clave = [
        "tipo_evento",
        "continuidad_desde",
        "continuidad_hacia",
        "periodo_transicion",
    ]

    if aplicados.duplicated(
        clave
    ).any():

        raise RuntimeError(
            "Quedaron continuidades societarias duplicadas "
            "después de la consolidación."
        )


# =============================================================================
# 5. HOMOLOGACIÓN CONTROLADA DEFINITIVA
# =============================================================================

def homologar_controlado(
    sbs,
    notas,
    columnas_df,
    especiales=None,
    guardar_auditoria=False
):
    """
    Proceso definitivo:

    detectar
        ->
    validar
        ->
    consolidar
        ->
    construir id_empresa
        ->
    controlar
    """

    # -------------------------------------------------------------------------
    # 5.1. Detectar eventos con el Bloque 8.
    # -------------------------------------------------------------------------

    eventos = evaluar_eventos(
        sbs,
        notas,
        columnas_df,
        especiales=especiales
    )

    # -------------------------------------------------------------------------
    # 5.2. Validar propuestas.
    # -------------------------------------------------------------------------

    eventos = validar_eventos_aplicados(
        eventos,
        sbs
    )

    # -------------------------------------------------------------------------
    # 5.3. Consolidar evidencia duplicada.
    # -------------------------------------------------------------------------

    eventos = consolidar_eventos(
        eventos
    )

    # -------------------------------------------------------------------------
    # 5.4. Control lógico.
    # -------------------------------------------------------------------------

    control_final_eventos(
        eventos
    )

    # -------------------------------------------------------------------------
    # 5.5. Construir identidades.
    #
    # IMPORTANTE:
    # en este punto se utiliza construir_id_empresa() del Bloque 8A,
    # porque 8A se carga después de 8 y antes de 8B.
    # -------------------------------------------------------------------------

    sbs = construir_id_empresa(
        sbs,
        eventos
    )

    # -------------------------------------------------------------------------
    # 5.6. Verificar identidades.
    # -------------------------------------------------------------------------

    controlar_identidades(
        sbs,
        eventos
    )

    # -------------------------------------------------------------------------
    # 5.7. Mantener auditoría únicamente en memoria.
    # -------------------------------------------------------------------------

    if guardar_auditoria:

        guardar_eventos_identidad(
            eventos
        )

    # -------------------------------------------------------------------------
    # 5.8. Reportar.
    # -------------------------------------------------------------------------

    aplicados = eventos[
        eventos[
            "aplicado"
        ].eq(True)
    ]

    reportar(
        "\nEventos societarios después de validación: "
        f"{len(eventos):,}."
    )

    reportar(
        "Continuidades empresariales aplicadas: "
        f"{len(aplicados):,}."
    )

    for _, evento in aplicados.iterrows():

        reportar(
            f"  {evento['continuidad_desde']} "
            "-> "
            f"{evento['continuidad_hacia']} "
            "| transición "
            f"{evento['periodo_transicion']}"
        )

    reportar(
        "\nNombres SBS distintos: "
        f"{sbs['empresa_sbs'].nunique()} "
        "| identidades empresariales: "
        f"{sbs['id_empresa'].nunique()}."
    )

    return (
        sbs,
        eventos
    )


# =============================================================================
# BLOQUE 9. ACUMULADO -> MENSUAL (NOMINAL) Y CONTROLES
# =============================================================================

def desacumular(sbs):
    """
    Convierte acumulados SBS en flujos mensuales por id_empresa y año.

    Enero:
        mensual = acumulado de enero.

    Febrero-diciembre:
        mensual = acumulado actual - acumulado del mes inmediatamente anterior.

    Si falta el valor actual o el mes anterior no es consecutivo/no tiene dato,
    el flujo mensual queda NaN y se activa la bandera correspondiente.
    """

    p = sbs.copy()

    p["anio"] = p["periodo"].str[:4].astype(int)
    p["mes"] = p["periodo"].str[5:7].astype(int)

    p = (
        p.sort_values(
            ["id_empresa", "anio", "mes"]
        )
        .reset_index(drop=True)
    )

    grupo = p.groupby(
        ["id_empresa", "anio"],
        dropna=False
    )

    mes_anterior = grupo["mes"].shift(1)

    es_enero = p["mes"].eq(1)

    anterior_consecutivo = (
        mes_anterior.eq(
            p["mes"] - 1
        )
    )

    for variable, flag in [
        (
            "primas_netas",
            "flag_hueco_netas"
        ),
        (
            "primas_cedidas",
            "flag_hueco_cedidas"
        ),
    ]:

        col_acum = (
            f"{variable}_acum_sbs"
        )

        col_mensual = (
            f"{variable}_mensual_nominal"
        )

        acum = p[col_acum]

        anterior = grupo[
            col_acum
        ].shift(1)

        p[col_mensual] = np.nan

        # ---------------------------------------------------------------------
        # PASO 1. Enero.
        # ---------------------------------------------------------------------
        enero_valido = (
            es_enero
            & acum.notna()
        )

        p.loc[
            enero_valido,
            col_mensual
        ] = acum.loc[
            enero_valido
        ]

        # ---------------------------------------------------------------------
        # PASO 2. Febrero-diciembre.
        # ---------------------------------------------------------------------
        resto_valido = (
            (~es_enero)
            & anterior_consecutivo
            & acum.notna()
            & anterior.notna()
        )

        p.loc[
            resto_valido,
            col_mensual
        ] = (
            acum.loc[resto_valido]
            - anterior.loc[resto_valido]
        )

        # ---------------------------------------------------------------------
        # PASO 3. Bandera de observación no calculable.
        # ---------------------------------------------------------------------
        calculable = (
            enero_valido
            | resto_valido
        )

        p[flag] = ~calculable

        reportar(
            f"  {variable}: "
            f"{int(p[flag].sum()):,} "
            "filas no calculables por faltante/hueco."
        )

    return p.drop(
        columns=["anio", "mes"]
    )


def controlar_desacumulacion(p):
    """
    Controles:
    1. En empresa-año con 12 flujos mensuales válidos:
       suma mensual = acumulado de diciembre.
    2. Diagnóstico del reinicio anual diciembre -> enero.
    """

    q = p.copy()

    q["anio"] = (
        q["periodo"].str[:4]
    )

    reportar(
        "\nControles de desacumulación:"
    )

    for variable in [
        "primas_netas",
        "primas_cedidas"
    ]:

        mensual = (
            f"{variable}_mensual_nominal"
        )

        acum = (
            f"{variable}_acum_sbs"
        )

        grupos = q.groupby(
            ["id_empresa", "anio"],
            dropna=False
        )

        n_meses = (
            grupos["periodo"]
            .nunique()
        )

        n_validos = (
            grupos[mensual]
            .count()
        )

        indices = n_meses[
            (n_meses == 12)
            & (n_validos == 12)
        ].index

        if len(indices) == 0:

            reportar(
                f"  {variable}: "
                "no hay empresa-año con "
                "12 flujos mensuales completos "
                "para validar."
            )

            continue

        suma = (
            grupos[mensual]
            .sum(min_count=12)
            .reindex(indices)
        )

        diciembre = (
            q[
                q["periodo"]
                .str.endswith("-12")
            ]
            .set_index(
                ["id_empresa", "anio"]
            )[acum]
            .reindex(indices)
        )

        dif = (
            suma - diciembre
        ).abs()

        fallas = int(
            (dif > TOLERANCIA)
            .fillna(True)
            .sum()
        )

        max_dif = (
            dif.max(
                skipna=True
            )
        )

        reportar(
            f"  {variable}: "
            f"{len(indices)} empresa-año completos | "
            f"diferencia máxima="
            f"{max_dif:,.6f} miles S/ | "
            f"fallas={fallas}"
        )

        if fallas:

            raise RuntimeError(
                f"{variable}: la suma mensual "
                f"no reproduce diciembre en "
                f"{fallas} empresa-año."
            )

    # -------------------------------------------------------------------------
    # PASO 3. Reinicio diciembre -> enero.
    # Es diagnóstico, no condición obligatoria de 100 %.
    # -------------------------------------------------------------------------
    dic = q[
        q["periodo"]
        .str.endswith("-12")
    ].copy()

    dic["anio_siguiente"] = (
        dic["anio"].astype(int)
        + 1
    ).astype(str)

    ene = q[
        q["periodo"]
        .str.endswith("-01")
    ].copy()

    pares = ene.merge(
        dic,
        left_on=[
            "id_empresa",
            "anio"
        ],
        right_on=[
            "id_empresa",
            "anio_siguiente"
        ],
        suffixes=(
            "_ene",
            "_dic"
        )
    )

    for variable in [
        "primas_netas",
        "primas_cedidas"
    ]:

        c_ene = (
            f"{variable}_acum_sbs_ene"
        )

        c_dic = (
            f"{variable}_acum_sbs_dic"
        )

        comp = pares[
            [c_ene, c_dic]
        ].dropna()

        if comp.empty:

            reportar(
                f"  Reinicio dic->ene "
                f"{variable}: "
                "sin pares comparables."
            )

            continue

        porcentaje = (
            comp[c_ene]
            < comp[c_dic]
        ).mean() * 100

        reportar(
            f"  Reinicio dic->ene "
            f"{variable}: "
            f"{porcentaje:.1f}% de "
            f"{len(comp)} pares comparables."
        )


# =============================================================================
# BLOQUE 10. UNIÓN SBS + BCRP POR PERIODO
# =============================================================================

def integrar(sbs, bcrp):
    """
    Une la base empresa × mes de SBS con las series mensuales del BCRP
    mediante una unión many_to_one por periodo.
    """

    filas_antes = len(sbs)

    # -------------------------------------------------------------------------
    # PASO 1. Controlar unicidad mensual del BCRP.
    # -------------------------------------------------------------------------
    if bcrp["periodo"].duplicated().any():

        raise RuntimeError(
            "BCRP contiene periodos duplicados "
            "antes de la integración."
        )

    # -------------------------------------------------------------------------
    # PASO 2. Unión empresa-mes con BCRP.
    # -------------------------------------------------------------------------
    base = sbs.merge(
        bcrp,
        on="periodo",
        how="left",
        validate="many_to_one"
    )

    # -------------------------------------------------------------------------
    # PASO 3. Verificar que no cambió el número de observaciones.
    # -------------------------------------------------------------------------
    if len(base) != filas_antes:

        raise RuntimeError(
            "La unión SBS-BCRP "
            "cambió el número de filas."
        )

    requeridas_bcrp = [
        "ingreso_formal_nominal",
        "tasa_referencia",
        "ipc",
    ]

    # -------------------------------------------------------------------------
    # PASO 4. Verificar cobertura BCRP.
    # -------------------------------------------------------------------------
    if (
        base[
            requeridas_bcrp
        ]
        .isna()
        .any()
        .any()
    ):

        raise RuntimeError(
            "Hay meses SBS sin dato BCRP "
            "después de la unión."
        )

    # -------------------------------------------------------------------------
    # PASO 5. Control del IPC.
    # -------------------------------------------------------------------------
    if (
        base["ipc"] <= 0
    ).any():

        raise RuntimeError(
            "Hay IPC no positivo "
            "después de integrar SBS y BCRP."
        )

    reportar(
        "\nUnión SBS-BCRP por periodo "
        f"(many_to_one): "
        f"{len(base):,} filas, sin cambio."
    )

    return base


# =============================================================================
# BLOQUE 11. DEFLACTACIÓN Y LOGARITMOS
# =============================================================================

def log_positivo(serie):
    """
    Calcula logaritmo natural solo para valores positivos.
    Ceros, negativos y faltantes permanecen en la base,
    pero su logaritmo queda como NaN.
    """

    return np.log(
        serie.where(
            serie > 0
        )
    )


def transformar(base):
    """
    Deflacta variables monetarias con IPC base dic-2021 = 100
    y genera las variables logarítmicas.
    """

    base = base.copy()

    # -------------------------------------------------------------------------
    # PASO 1. Verificar IPC.
    # -------------------------------------------------------------------------
    if (
        base["ipc"] <= 0
    ).any():

        raise RuntimeError(
            "No se puede deflactar: "
            "existe IPC no positivo."
        )

    deflactor = (
        base["ipc"] / 100
    )

    # -------------------------------------------------------------------------
    # PASO 2. Primas netas reales.
    # SBS permanece expresada en MILES de soles.
    # -------------------------------------------------------------------------
    base[
        "primas_netas_mensual_real"
    ] = (
        base[
            "primas_netas_mensual_nominal"
        ]
        / deflactor
    )

    # -------------------------------------------------------------------------
    # PASO 3. Primas cedidas reales.
    # -------------------------------------------------------------------------
    base[
        "primas_cedidas_mensual_real"
    ] = (
        base[
            "primas_cedidas_mensual_nominal"
        ]
        / deflactor
    )

    # -------------------------------------------------------------------------
    # PASO 4. Ingreso formal real.
    # -------------------------------------------------------------------------
    base[
        "ingreso_real"
    ] = (
        base[
            "ingreso_formal_nominal"
        ]
        / deflactor
    )

    # -------------------------------------------------------------------------
    # PASO 5. Logaritmos naturales.
    # -------------------------------------------------------------------------
    base[
        "log_primas_reales"
    ] = log_positivo(
        base[
            "primas_netas_mensual_real"
        ]
    )

    base[
        "log_primas_cedidas"
    ] = log_positivo(
        base[
            "primas_cedidas_mensual_real"
        ]
    )

    base[
        "log_ingreso_real"
    ] = log_positivo(
        base[
            "ingreso_real"
        ]
    )

    # -------------------------------------------------------------------------
    # PASO 6. Muestra disponible para el modelo.
    # La tasa de referencia se conserva en nivel.
    # -------------------------------------------------------------------------
    base[
        "muestra_modelo"
    ] = base[
        [
            "log_primas_reales",
            "log_primas_cedidas",
            "log_ingreso_real",
            "tasa_referencia",
        ]
    ].notna().all(
        axis=1
    )

    return base


# =============================================================================
# BLOQUE 12. DIAGNÓSTICOS DE VALORES Y OUTLIERS
# =============================================================================
# Los diagnósticos NO eliminan, reemplazan ni winsorizan observaciones.
# Los resultados se muestran en consola/log.
# No se crean archivos en /salidas.
# =============================================================================


def categorias(serie):
    """Clasifica valores para diagnóstico."""

    return {
        "positivos_>=1": int((serie >= 1).sum()),
        "entre_0_y_1": int(((serie > 0) & (serie < 1)).sum()),
        "ceros": int((serie == 0).sum()),
        "negativos": int((serie < 0).sum()),
        "faltantes": int(serie.isna().sum()),
    }


def detectar_outliers_iqr(base, columnas):
    """
    Detecta valores atípicos mediante 1.5 × IQR.

    Es únicamente diagnóstico:
    NO elimina ni modifica observaciones.
    """

    filas = []

    for col in columnas:

        serie = base[col].dropna()

        if len(serie) < 4:
            continue

        q1, q3 = serie.quantile([0.25, 0.75])

        iqr = q3 - q1

        inferior = q1 - 1.5 * iqr
        superior = q3 + 1.5 * iqr

        mask = (
            base[col].notna()
            &
            (
                (base[col] < inferior)
                |
                (base[col] > superior)
            )
        )

        for _, fila in base.loc[
            mask,
            ["periodo", "id_empresa", "empresa", col]
        ].iterrows():

            filas.append({
                "periodo": fila["periodo"],
                "id_empresa": fila["id_empresa"],
                "empresa": fila["empresa"],
                "variable": col,
                "valor": fila[col],
                "q1": q1,
                "q3": q3,
                "limite_inferior": inferior,
                "limite_superior": superior,
            })

    return pd.DataFrame(filas)


def diagnosticar(
    base,
    largo,
    guardar_detalle=False
):
    """
    Reporta ceros, negativos, faltantes y outliers.

    guardar_detalle se conserva por compatibilidad,
    pero no se generan archivos auxiliares.
    """

    reportar(
        "\nCómo venían las celdas TOTAL en los XLS:"
    )

    for tipo, n in (
        largo["tipo_valor"]
        .value_counts(dropna=False)
        .items()
    ):

        reportar(
            f"  {tipo}: {n}"
        )

    cols = [
        "primas_netas_acum_sbs",
        "primas_netas_mensual_nominal",
        "primas_cedidas_acum_sbs",
        "primas_cedidas_mensual_nominal",
    ]

    reportar(
        "\nDiagnóstico por variable "
        "(miles de S/ corrientes):"
    )

    for col in cols:

        resultado = categorias(
            base[col]
        )

        reportar(
            f"  {col:<34} "
            + " | ".join(
                f"{k}={v}"
                for k, v in resultado.items()
            )
        )

    # -------------------------------------------------------------------------
    # Negativos
    # -------------------------------------------------------------------------

    negativos = base[
        (base["primas_netas_mensual_nominal"] < 0)
        |
        (base["primas_cedidas_mensual_nominal"] < 0)
    ].copy()

    reportar(
        "\nFlujos mensuales negativos: "
        f"{len(negativos):,}. "
        "Se conservan como ajustes observados; "
        "no se reemplazan por cero."
    )

    if not negativos.empty:

        for _, fila in (
            negativos
            .sort_values(["empresa", "periodo"])
            .head(10)
            .iterrows()
        ):

            reportar(
                f"  {fila['periodo']} | "
                f"{fila['empresa']} | "
                f"netas={fila['primas_netas_mensual_nominal']} | "
                f"cedidas={fila['primas_cedidas_mensual_nominal']}"
            )

    # -------------------------------------------------------------------------
    # Outliers
    # -------------------------------------------------------------------------

    outliers = detectar_outliers_iqr(
        base,
        [
            "primas_netas_mensual_nominal",
            "primas_cedidas_mensual_nominal",
        ]
    )

    reportar(
        f"\nOutliers señalados por IQR: "
        f"{len(outliers):,}. "
        "No se eliminan ni winsorizan."
    )

    # -------------------------------------------------------------------------
    # Muestra
    # -------------------------------------------------------------------------

    muestra = int(
        base["muestra_modelo"].sum()
    )

    reportar(
        f"\nObservaciones totales: {len(base):,}"
    )

    reportar(
        f"Observaciones utilizables en el modelo: {muestra:,}"
    )

    reportar(
        f"Observaciones fuera de muestra del modelo: "
        f"{len(base) - muestra:,}"
    )

    return negativos, outliers


# =============================================================================
# BLOQUE 13. CONTROLES FINALES
# =============================================================================


def controles_finales(base):
    """
    Verifica estructura, identificadores, cobertura temporal,
    tamaño de la base y variables sustantivas.
    """

    reportar(
        "\nControles finales:"
    )

    # -------------------------------------------------------------------------
    # 1. Cobertura temporal
    # -------------------------------------------------------------------------

    encontrados = sorted(
        base["periodo"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    esperados = periodos_esperados()

    if encontrados != esperados:

        raise RuntimeError(
            "La base final no contiene exactamente los 72 meses. "
            f"Faltan={sorted(set(esperados) - set(encontrados))}; "
            f"sobran={sorted(set(encontrados) - set(esperados))}"
        )

    reportar(
        f"  Periodo: {FECHA_INICIO} a {FECHA_CORTE} | "
        f"meses={len(encontrados)}"
    )

    # -------------------------------------------------------------------------
    # 2. Columnas finales exactas
    # -------------------------------------------------------------------------

    columnas_esperadas = (
        ["id_observacion"]
        + COLUMNAS_FINALES
    )

    if base.columns.tolist() != columnas_esperadas:

        raise RuntimeError(
            "Las columnas finales no coinciden con la estructura definida.\n"
            f"Esperadas={columnas_esperadas}\n"
            f"Obtenidas={base.columns.tolist()}"
        )

    # -------------------------------------------------------------------------
    # 3. Identificadores
    # -------------------------------------------------------------------------

    identificadores = [
        "id_observacion",
        "periodo",
        "id_empresa",
        "empresa",
    ]

    if base[
        identificadores
    ].isna().any().any():

        raise RuntimeError(
            "Hay identificadores vacíos en la base final."
        )

    if base["id_observacion"].duplicated().any():

        raise RuntimeError(
            "Hay id_observacion duplicados."
        )

    if base["id_observacion"].tolist() != list(
        range(1, len(base) + 1)
    ):

        raise RuntimeError(
            "id_observacion no es correlativo de 1 a N."
        )

    # -------------------------------------------------------------------------
    # 4. Una observación por empresa y mes
    # -------------------------------------------------------------------------

    duplicados = int(
        base.duplicated(
            ["periodo", "id_empresa"]
        ).sum()
    )

    if duplicados:

        raise RuntimeError(
            "Hay duplicados id_empresa × periodo."
        )

    reportar(
        "  Duplicados id_empresa × periodo: 0"
    )

    # -------------------------------------------------------------------------
    # 5. Nombre analítico estable
    # -------------------------------------------------------------------------

    nombres_por_id = (
        base.groupby("id_empresa")["empresa"]
        .nunique()
    )

    if (nombres_por_id != 1).any():

        raise RuntimeError(
            "Hay id_empresa con más de un nombre analítico."
        )

    # -------------------------------------------------------------------------
    # 6. Identidades
    # -------------------------------------------------------------------------

    n_identidades = (
        base["id_empresa"].nunique()
    )

    if n_identidades != 20:

        raise RuntimeError(
            f"Se esperaban 20 identidades empresariales "
            f"y se obtuvieron {n_identidades}."
        )

    meses_por_empresa = (
        base.groupby("id_empresa")["periodo"]
        .nunique()
    )

    balanceado = bool(
        (
            meses_por_empresa
            == MESES_ESPERADOS
        ).all()
    )

    reportar(
        f"  Identidades empresariales: {n_identidades} | "
        f"panel "
        f"{'balanceado' if balanceado else 'NO balanceado'} | "
        f"mínimo={meses_por_empresa.min()} meses | "
        f"máximo={meses_por_empresa.max()} meses"
    )

    # -------------------------------------------------------------------------
    # 7. Tamaño mínimo
    # -------------------------------------------------------------------------

    if len(base) < MINIMO_OBSERVACIONES:

        raise RuntimeError(
            f"La base procesada tiene menos de "
            f"{MINIMO_OBSERVACIONES:,} observaciones."
        )

    if base.shape[1] < 6:

        raise RuntimeError(
            "La base final tiene menos de 6 columnas."
        )

    reportar(
        f"  Observaciones: {len(base):,} "
        f"(mínimo exigido: {MINIMO_OBSERVACIONES:,})"
    )

    # -------------------------------------------------------------------------
    # 8. Variables sustantivas
    # -------------------------------------------------------------------------

    sustantivas = [
        "log_primas_reales",
        "log_primas_cedidas",
        "log_ingreso_real",
        "tasa_referencia",
    ]

    faltan = [
        col
        for col in sustantivas
        if col not in base.columns
    ]

    if faltan:

        raise RuntimeError(
            "Faltan variables sustantivas: "
            + ", ".join(faltan)
        )

    # -------------------------------------------------------------------------
    # 9. Presencia de ambas fuentes
    # -------------------------------------------------------------------------

    if not {
        "primas_netas_acum_sbs",
        "primas_cedidas_acum_sbs",
    }.issubset(base.columns):

        raise RuntimeError(
            "No están presentes las variables fuente SBS."
        )

    if not {
        "ingreso_formal_nominal",
        "tasa_referencia",
        "ipc",
    }.issubset(base.columns):

        raise RuntimeError(
            "No están presentes las variables fuente BCRP."
        )

    # -------------------------------------------------------------------------
    # 10. Muestra del modelo
    # -------------------------------------------------------------------------

    muestra = int(
        base["muestra_modelo"].sum()
    )

    reportar(
        f"  Muestra disponible para el modelo: "
        f"{muestra:,} de {len(base):,}"
    )

    # -------------------------------------------------------------------------
    # 11. Faltantes
    # -------------------------------------------------------------------------

    reportar(
        "  Faltantes por variable:"
    )

    for col in base.columns:

        reportar(
            f"    {col:<34} "
            f"no vacíos={int(base[col].notna().sum()):>6,} | "
            f"faltantes={int(base[col].isna().sum()):>5,}"
        )


# =============================================================================
# BLOQUE 14. GUARDADO, LOG Y SHA-256
# =============================================================================


def escribir_log(texto):
    """Añade al log una línea con fecha y hora."""

    ahora = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    limpio = (
        str(texto)
        .replace("\n", " ")
        .replace("\r", " ")
    )

    with open(
        ARCHIVO_LOG,
        "a",
        encoding="utf-8"
    ) as archivo:

        archivo.write(
            f"{ahora} | {limpio}\n"
        )


def guardar(base):
    """
    Guarda la base procesada y calcula su SHA-256.

    /salidas se reserva para las tablas y figuras del Script 04.
    """

    ruta = (
        CARPETA_PROC
        / f"datos_procesados_{MATRICULA}.csv"
    )

    base.to_csv(
        ruta,
        index=False,
        encoding=CSV_ENCODING,
        float_format="%.6f"
    )

    huella = hashlib.sha256(
        ruta.read_bytes()
    ).hexdigest()

    reportar(
        "\nArchivo procesado: "
        f"{ruta.relative_to(RAIZ).as_posix()} "
        f"({len(base):,} filas × "
        f"{base.shape[1]} columnas)"
    )

    reportar(
        f"SHA-256: {huella}"
    )

    escribir_log(
        f"PROCESADO {ruta.name} | "
        f"filas={len(base)} | "
        f"columnas={base.shape[1]} | "
        f"muestra_modelo="
        f"{int(base['muestra_modelo'].sum())} | "
        f"sha256={huella}"
    )

    return ruta, huella


# =============================================================================
# BLOQUE 15. EJECUCIÓN PRINCIPAL
# =============================================================================


def preparar_base():
    """
    Ejecuta todo el pipeline del Script 03.

    Los controles se muestran en consola/log.
    No se generan archivos auxiliares en /salidas.
    """

    # -------------------------------------------------------------------------
    # 1. BCRP
    # -------------------------------------------------------------------------

    bcrp = cargar_bcrp()

    # -------------------------------------------------------------------------
    # 2. SBS
    # -------------------------------------------------------------------------

    (
        sbs,
        largo,
        notas,
        columnas_df,
        comparacion
    ) = cargar_sbs(
        guardar_auditoria=False
    )

    resumir_notas(
        notas
    )

    # -------------------------------------------------------------------------
    # 3. Depuración de registros especiales
    # -------------------------------------------------------------------------

    (
        sbs,
        largo,
        comparacion,
        registros_especiales
    ) = depurar_registros_especiales(
        sbs,
        largo,
        notas,
        columnas_df,
        guardar_auditoria=False
    )

    # -------------------------------------------------------------------------
    # 4. Identidad empresarial
    # -------------------------------------------------------------------------

    (
        sbs,
        eventos
    ) = homologar_controlado(
        sbs,
        notas,
        columnas_df,
        especiales=registros_especiales,
        guardar_auditoria=False
    )

    # Control descriptivo de trayectorias.
    guardar_presencia_y_trayectoria(
        sbs,
        eventos
    )

    # -------------------------------------------------------------------------
    # 5. Desacumulación
    # -------------------------------------------------------------------------

    reportar(
        "\nDesacumulación:"
    )

    sbs = desacumular(
        sbs
    )

    controlar_desacumulacion(
        sbs
    )

    # -------------------------------------------------------------------------
    # 6. Integración SBS + BCRP
    # -------------------------------------------------------------------------

    base = integrar(
        sbs,
        bcrp
    )

    # -------------------------------------------------------------------------
    # 7. Deflactación y logaritmos
    # -------------------------------------------------------------------------

    base = transformar(
        base
    )

    # -------------------------------------------------------------------------
    # 8. Selección de columnas finales
    # -------------------------------------------------------------------------

    base = (
        base[
            COLUMNAS_FINALES
        ]
        .sort_values(
            [
                "id_empresa",
                "periodo"
            ]
        )
        .reset_index(
            drop=True
        )
    )

    # -------------------------------------------------------------------------
    # 9. ID único de observación
    # -------------------------------------------------------------------------

    base.insert(
        0,
        "id_observacion",
        range(
            1,
            len(base) + 1
        )
    )

    # -------------------------------------------------------------------------
    # 10. Diagnósticos
    # -------------------------------------------------------------------------

    diagnosticar(
        base,
        largo,
        guardar_detalle=False
    )

    # -------------------------------------------------------------------------
    # 11. Controles finales
    # -------------------------------------------------------------------------

    controles_finales(
        base
    )

    # -------------------------------------------------------------------------
    # 12. Codificación final de muestra_modelo
    # -------------------------------------------------------------------------
    # Internamente fue booleana.
    # En el CSV final se exporta como:
    # 1 = observación utilizable
    # 0 = observación no utilizable
    # -------------------------------------------------------------------------

    base[
        "muestra_modelo"
    ] = (
        base[
            "muestra_modelo"
        ]
        .astype("int8")
    )

    return base


def pruebas_previas():
    """
    Ejecuta todo el pipeline en memoria,
    pero no guarda el CSV procesado.
    """

    reportar(
        "=" * 78
    )

    reportar(
        "PRUEBA COMPLETA DEL SCRIPT 03"
    )

    reportar(
        "=" * 78
    )

    base = preparar_base()

    reportar(
        "\nPRUEBA COMPLETA SUPERADA."
    )

    reportar(
        f"Resultado: {len(base):,} filas × "
        f"{base.shape[1]} columnas."
    )

    return base


def main():
    """Ejecución definitiva."""

    reportar(
        "=" * 78
    )

    reportar(
        "SCRIPT 03 — LIMPIEZA E INTEGRACIÓN "
        f"| empresa × mes | "
        f"{FECHA_INICIO} a {FECHA_CORTE}"
    )

    reportar(
        "=" * 78
    )

    base = preparar_base()

    guardar(
        base
    )

    print(
        "\nScript 03 finalizado correctamente."
    )


if __name__ == "__main__":

    try:

        if "--pruebas" in sys.argv:

            pruebas_previas()

        else:

            main()

    except Exception as error:

        mensaje = (
            f"{type(error).__name__}: "
            f"{error}"
        )

        print(
            "\n[ERROR] "
            "El Script 03 se detuvo: "
            f"{mensaje}"
        )

        if "--pruebas" not in sys.argv:

            escribir_log(
                "ERROR SCRIPT 03 | "
                f"{mensaje}"
            )

        raise
