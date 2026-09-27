# Nombres y apellidos: Vania Adriana Egas Araujo
# Código de matrícula: 2024200498D
# Tema: N.º 15 del temario — Determinantes de las primas cedidas, el ingreso formal y la tasa de referencia en las primas netas del mercado asegurador peruano (2020-2025)
# Fecha de extracción: 2026-09-25
"""
SCRIPT 03 — LIMPIEZA, TRANSFORMACIÓN E INTEGRACIÓN DE DATOS

Unidad de análisis: empresa × mes, 2020-01 a 2025-12.
- P009 = primas netas acumuladas por empresa.
- P010 = primas cedidas acumuladas por empresa.
- Se toma únicamente la fila TOTAL de cada empresa; no se trabaja por ramo.
- Los crudos no se modifican.
- Los cambios de denominación solo se homologan cuando existe evidencia explícita
  en los pies de página SBS y continuidad temporal verificable.
- Fusiones/absorciones se documentan, pero no se convierten automáticamente en
  una continuidad histórica.
- Enero mensual = enero acumulado; feb-dic = acumulado actual - acumulado previo.
- Los montos reales de primas se conservan en MILES de soles de dic-2021.
- Ceros, negativos y faltantes se conservan; los outliers solo se diagnostican.
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
# BLOQUE 2. PARÁMETROS Y RUTAS
# =============================================================================
FECHA_INICIO = "2020-01"
FECHA_CORTE = "2025-12"
MESES_ESPERADOS = 72
MATRICULA = "2024200498D"

COD_INGRESO = "PN31883GM"
COD_TASA = "PD04722MM"
COD_IPC = "PN38705PM"

HOJAS = {"P009": "primas_netas", "P010": "primas_cedidas"}
ROTULO_ENCABEZADO = r"RIESGOS\s*/\s*EMPRESAS"
ROTULOS_TOTAL_SISTEMA = {"TOTAL", "TOTAL GENERAL", "TOTAL SISTEMA", "TOTAL SISTEMA ASEGURADOR"}
ROTULOS_TOTAL_FILA = {"TOTAL", "TOTAL GENERAL"}

MINIMO_OBSERVACIONES = 1000
TOLERANCIA = 1.0  # miles de S/; tolerancia para redondeos de la fuente

MESES_BCRP = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "SET": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}

COLUMNAS_FINALES = [
    "periodo", "id_empresa", "empresa", "empresa_sbs", "estado_p009_p010",
    "primas_netas_acum_sbs", "primas_netas_mensual_nominal",
    "primas_netas_mensual_real", "log_primas_reales",
    "primas_cedidas_acum_sbs", "primas_cedidas_mensual_nominal",
    "primas_cedidas_mensual_real", "log_primas_cedidas",
    "ingreso_formal_nominal", "ingreso_real", "log_ingreso_real",
    "tasa_referencia", "ipc",
    "flag_hueco_netas", "flag_hueco_cedidas", "muestra_modelo",
]

RAIZ = Path(__file__).resolve().parent.parent
CARPETA_CRUDOS = RAIZ / "datos_crudos"
CARPETA_PROC = RAIZ / "datos_procesados"
CARPETA_SALIDAS = RAIZ / "salidas"
ARCHIVO_LOG = RAIZ / "log_ejecucion.txt"
ARCHIVO_CONTROL = CARPETA_SALIDAS / "control_03_limpieza.txt"

CARPETA_PROC.mkdir(parents=True, exist_ok=True)
CARPETA_SALIDAS.mkdir(parents=True, exist_ok=True)

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
# BLOQUE 6. PIES DE PÁGINA DE LA SBS
# =============================================================================
def extraer_notas(df, fila_tot, periodo, hoja, nombre_archivo):
    """
    Conserva una nota COMPLETA por fila debajo de TOTAL y registra las columnas
    que aportaron texto. Esto evita fragmentar una misma nota en varias celdas.
    """
    notas = []
    for i in range(fila_tot + 1, len(df)):
        partes = []
        columnas = []
        for c in df.columns:
            valor = df.at[i, c]
            if isinstance(valor, str) and valor.strip():
                texto = re.sub(r"\s+", " ", valor.strip())
                partes.append(texto)
                columnas.append(letra_columna(c))
        if not partes:
            continue
        # Evita repetir exactamente la misma pieza dentro de la fila.
        partes = list(dict.fromkeys(partes))
        nota = " | ".join(partes)
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


def guardar_notas(notas):
    ruta = CARPETA_SALIDAS / "notas_sbs_s401.csv"
    if notas.empty:
        pd.DataFrame(columns=["periodo", "hoja", "fila_excel", "columnas_texto",
                              "nota_original", "nota_normalizada", "referencias_detectadas",
                              "archivo_origen"]).to_csv(ruta, index=False, encoding="utf-8")
        reportar("Pies de página SBS: no se encontraron textos debajo de TOTAL.")
        return
    notas.to_csv(ruta, index=False, encoding="utf-8")
    reportar(f"Pies de página SBS: {len(notas):,} filas de nota; "
             f"{notas['nota_normalizada'].nunique():,} textos normalizados distintos.")


# =============================================================================
# BLOQUE 6B. DEPURACIÓN DE REGISTROS ESPECIALES SBS
# =============================================================================
# Algunas columnas de los XLS parecen empresas por su estructura, pero las
# notas al pie pueden indicar que son registros transitorios de una empresa
# absorbida y no una empresa activa del periodo.
#
# Regla:
# - NO se codifican nombres de empresas manualmente.
# - La reclasificación exige evidencia conjunta del encabezado y de la nota SBS.
# - El registro especial se conserva para auditoría.
# - No entra en la base principal empresa × mes.
# =============================================================================


def _refs_a_set(valor):
    """Convierte '1;2' en {'1', '2'} y maneja vacíos."""
    if pd.isna(valor):
        return set()

    return {
        x.strip()
        for x in str(valor).split(";")
        if x.strip()
    }


def _es_registro_absorbida_transitorio(nombre_limpio, nota_original):
    """
    Identifica una columna especial correspondiente a flujos históricos
    de una empresa absorbida.

    No usa nombres específicos de compañías.

    Exige:
    1. que el encabezado tenga la marca '(abs)';
    2. que la nota hable de FLUJOS;
    3. que la nota mencione una FUSIÓN;
    4. que indique que esos flujos pertenecen al periodo anterior a la fusión.
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

    habla_de_flujos = "FLUJOS" in nota
    habla_de_fusion = "FUSION" in nota

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
    Reclasifica columnas inicialmente leídas como empresa cuando la propia
    evidencia SBS demuestra que son registros transitorios de una absorbida.

    Devuelve:
        sbs_filtrado
        largo_filtrado
        comparacion_filtrada
        especiales
    """

    columnas_empresa = columnas_df[
        columnas_df["tipo_columna"].eq("empresa")
    ].copy()

    registros = []

    # -------------------------------------------------------------------------
    # 1. Vincular cada encabezado empresarial con su nota mediante
    #    periodo + hoja + referencia.
    # -------------------------------------------------------------------------
    for _, col in columnas_empresa.iterrows():

        refs_columna = _refs_a_set(
            col["referencia_nota"]
        )

        if not refs_columna:
            continue

        notas_periodo = notas[
            notas["periodo"].eq(col["periodo"])
            & notas["hoja"].eq(col["hoja"])
        ]

        for _, nota in notas_periodo.iterrows():

            refs_nota = _refs_a_set(
                nota["referencias_detectadas"]
            )

            if not (refs_columna & refs_nota):
                continue

            if _es_registro_absorbida_transitorio(
                col["nombre_limpio"],
                nota["nota_original"]
            ):

                registros.append({
                    "periodo": col["periodo"],
                    "hoja": col["hoja"],
                    "columna_excel": col["columna_excel"],
                    "encabezado_original": col["encabezado_original"],
                    "empresa_sbs": col["nombre_limpio"],
                    "referencia_nota": col["referencia_nota"],
                    "tipo_inicial": "empresa",
                    "tipo_definitivo": "registro_especial_absorcion",
                    "entra_panel": False,
                    "motivo": (
                        "La nota SBS identifica flujos de una empresa "
                        "absorbida correspondientes al periodo previo "
                        "a la fusión."
                    ),
                    "nota_sbs": nota["nota_original"],
                    "fila_nota_excel": nota["fila_excel"],
                    "archivo_origen": nota["archivo_origen"],
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

    if not especiales.empty:

        especiales = (
            especiales
            .drop_duplicates(
                ["periodo", "hoja", "empresa_sbs"]
            )
            .sort_values(
                ["periodo", "empresa_sbs", "hoja"]
            )
            .reset_index(drop=True)
        )

        # ---------------------------------------------------------------------
        # 2. Control: si un registro se identifica como especial, todas las
        #    hojas en las que esa misma columna empresarial esté presente
        #    deben concordar.
        # ---------------------------------------------------------------------
        claves = especiales[
            ["periodo", "empresa_sbs"]
        ].drop_duplicates()

        for _, clave in claves.iterrows():

            periodo = clave["periodo"]
            empresa = clave["empresa_sbs"]

            hojas_presentes = set(
                largo.loc[
                    largo["periodo"].eq(periodo)
                    & largo["empresa_sbs"].eq(empresa),
                    "hoja"
                ]
            )

            hojas_especiales = set(
                especiales.loc[
                    especiales["periodo"].eq(periodo)
                    & especiales["empresa_sbs"].eq(empresa),
                    "hoja"
                ]
            )

            if hojas_presentes != hojas_especiales:
                raise RuntimeError(
                    "Clasificación especial inconsistente para "
                    f"{empresa} en {periodo}. "
                    f"Hojas presentes={sorted(hojas_presentes)}; "
                    f"hojas clasificadas como especiales="
                    f"{sorted(hojas_especiales)}."
                )

        # ---------------------------------------------------------------------
        # 3. Incorporar los valores originales a la tabla de auditoría.
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
            on=["periodo", "hoja", "empresa_sbs"],
            how="left",
            validate="one_to_one"
        )

        # ---------------------------------------------------------------------
        # 4. Excluirlos únicamente del panel empresarial.
        #    NO se borran de la auditoría.
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
            largo.loc[~mascara_largo]
            .copy()
            .reset_index(drop=True)
        )

        mascara_sbs = sbs.apply(
            lambda f: (
                f["periodo"],
                f["empresa_sbs"]
            ) in claves_excluir,
            axis=1
        )

        sbs_filtrado = (
            sbs.loc[~mascara_sbs]
            .copy()
            .reset_index(drop=True)
        )

    else:

        largo_filtrado = largo.copy()
        sbs_filtrado = sbs.copy()

    # -------------------------------------------------------------------------
    # 5. Reconstruir comparación P009-P010 después de quitar especiales.
    # -------------------------------------------------------------------------
    presencia = (
        largo_filtrado
        .assign(presente=1)
        .pivot_table(
            index=["periodo", "empresa_sbs"],
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
        "primas_cedidas"
    ]:
        if variable not in presencia.columns:
            presencia[variable] = 0

    presencia["estado_p009_p010"] = np.select(
        [
            presencia["primas_netas"].eq(1)
            & presencia["primas_cedidas"].eq(1),

            presencia["primas_netas"].eq(1)
            & presencia["primas_cedidas"].eq(0),

            presencia["primas_netas"].eq(0)
            & presencia["primas_cedidas"].eq(1),
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

    # Reemplazar el estado antiguo por el estado ya depurado.
    sbs_filtrado = sbs_filtrado.drop(
        columns=["estado_p009_p010"],
        errors="ignore"
    )

    sbs_filtrado = sbs_filtrado.merge(
        comparacion_filtrada,
        on=["periodo", "empresa_sbs"],
        how="left",
        validate="one_to_one"
    )

    # -------------------------------------------------------------------------
    # 6. Auditoría.
    # -------------------------------------------------------------------------
    if guardar_auditoria:

        ruta = (
            CARPETA_SALIDAS
            / "registros_especiales_sbs.csv"
        )

        especiales.to_csv(
            ruta,
            index=False,
            encoding="utf-8"
        )

    if especiales.empty:

        reportar(
            "\nRegistros especiales SBS: "
            "no se detectaron columnas empresariales transitorias."
        )

    else:

        empresas_especiales = (
            especiales["empresa_sbs"]
            .drop_duplicates()
            .tolist()
        )

        reportar(
            "\nRegistros especiales SBS detectados: "
            f"{len(especiales):,} registros hoja-periodo."
        )

        reportar(
            "  Empresas/columnas afectadas: "
            + "; ".join(empresas_especiales)
        )

        reportar(
            "  Se conservan en auditoría y se excluyen "
            "del panel empresa × mes."
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
# IMPORTANTE:
# El id_empresa NO se asigna aquí únicamente por nombre.
# Se asignará en el Bloque 8 utilizando:
# - continuidad temporal;
# - cambios de denominación respaldados por SBS;
# - fusiones con continuidad explícitamente documentada;
# - posibles reutilizaciones de un mismo nombre.
# =============================================================================


def calcular_presencia(sbs, columna):
    """
    Resume la presencia temporal de cada valor de 'columna'.
    No supone que un mismo nombre equivalga necesariamente a una sola
    identidad jurídica.
    """
    g = sbs.groupby(columna)["periodo"]

    presencia = pd.DataFrame({
        "primer_periodo": g.min(),
        "ultimo_periodo": g.max(),
        "meses_presente": g.nunique(),
    })

    presencia["meses_en_su_rango"] = (
        presencia["ultimo_periodo"].apply(periodo_a_numero)
        - presencia["primer_periodo"].apply(periodo_a_numero)
        + 1
    )

    presencia["meses_faltantes_en_su_rango"] = (
        presencia["meses_en_su_rango"]
        - presencia["meses_presente"]
    )

    return presencia.reset_index()


def _nombres_en_orden(grupo):
    """
    Devuelve los nombres SBS utilizados por una trayectoria,
    respetando su orden temporal y evitando repeticiones consecutivas.
    """
    grupo = grupo.sort_values("periodo")

    nombres = []

    for nombre in grupo["empresa_sbs"]:
        if not nombres or nombres[-1] != nombre:
            nombres.append(nombre)

    return " -> ".join(nombres)


def guardar_presencia_y_trayectoria(sbs, eventos):
    """
    Genera las tablas auxiliares de presencia y trayectoria.

    Requiere que el Bloque 8 ya haya creado id_empresa.
    """

    if "id_empresa" not in sbs.columns:
        raise RuntimeError(
            "No existe id_empresa. La identidad empresarial debe "
            "construirse primero en el Bloque 8."
        )

    # -------------------------------------------------------------------------
    # 1. PRESENCIA POR NOMBRE TAL COMO APARECE EN SBS
    # -------------------------------------------------------------------------
    presencia = calcular_presencia(
        sbs,
        "empresa_sbs"
    ).rename(
        columns={
            "empresa_sbs": "empresa_original"
        }
    )

    ids_por_nombre = (
        sbs.groupby("empresa_sbs")["id_empresa"]
        .agg(
            lambda x: ";".join(
                sorted(pd.unique(x.astype(str)))
            )
        )
    )

    presencia["ids_asociados"] = (
        presencia["empresa_original"]
        .map(ids_por_nombre)
    )

    presencia.to_csv(
        CARPETA_SALIDAS / "presencia_empresas.csv",
        index=False,
        encoding="utf-8"
    )

    # -------------------------------------------------------------------------
    # 2. TRAYECTORIA POR IDENTIDAD EMPRESARIAL
    # -------------------------------------------------------------------------
    filas_trayectoria = []

    for id_empresa, grupo in sbs.groupby(
        "id_empresa",
        sort=True
    ):

        grupo = grupo.sort_values("periodo")

        primer_periodo = grupo["periodo"].min()
        ultimo_periodo = grupo["periodo"].max()
        meses_presente = grupo["periodo"].nunique()

        meses_en_rango = (
            periodo_a_numero(ultimo_periodo)
            - periodo_a_numero(primer_periodo)
            + 1
        )

        filas_trayectoria.append({
            "id_empresa": id_empresa,
            "nombre_inicial_sbs": (
                grupo.iloc[0]["empresa_sbs"]
            ),
            "nombre_final_sbs": (
                grupo.iloc[-1]["empresa_sbs"]
            ),
            "nombres_sbs_en_trayectoria": (
                _nombres_en_orden(grupo)
            ),
            "primer_periodo": primer_periodo,
            "ultimo_periodo": ultimo_periodo,
            "meses_presente": meses_presente,
            "meses_en_su_rango": meses_en_rango,
            "meses_faltantes_en_su_rango": (
                meses_en_rango - meses_presente
            ),
            "tipo_evento_identidad": "",
            "periodo_evento": "",
            "evidencia_sbs": "",
        })

    trayectoria = pd.DataFrame(
        filas_trayectoria
    )

    # -------------------------------------------------------------------------
    # 3. Incorporar eventos que realmente afectan la continuidad de identidad
    # -------------------------------------------------------------------------
    if eventos is not None and not eventos.empty:

        if "aplicado" in eventos.columns:

            aplicados = eventos[
                eventos["aplicado"].eq(True)
            ]

            for _, evento in aplicados.iterrows():

                periodo = evento.get(
                    "periodo_transicion",
                    None
                )

                destino = evento.get(
                    "continuidad_hacia",
                    None
                )

                if (
                    not periodo
                    or pd.isna(periodo)
                    or not destino
                    or pd.isna(destino)
                ):
                    continue

                candidatos = sbs[
                    sbs["periodo"].eq(periodo)
                    & sbs["empresa_sbs"].eq(destino)
                ]

                ids = candidatos[
                    "id_empresa"
                ].dropna().unique()

                if len(ids) != 1:
                    continue

                id_afectado = ids[0]

                mask = trayectoria[
                    "id_empresa"
                ].eq(id_afectado)

                trayectoria.loc[
                    mask,
                    "tipo_evento_identidad"
                ] = evento.get(
                    "tipo_evento",
                    ""
                )

                trayectoria.loc[
                    mask,
                    "periodo_evento"
                ] = periodo

                trayectoria.loc[
                    mask,
                    "evidencia_sbs"
                ] = evento.get(
                    "evidencia_sbs",
                    ""
                )

    trayectoria.to_csv(
        CARPETA_SALIDAS / "trayectoria_empresarial.csv",
        index=False,
        encoding="utf-8"
    )

    # -------------------------------------------------------------------------
    # 4. EMPRESAS ACTIVAS POR MES
    # -------------------------------------------------------------------------
    filas_mes = []

    for periodo, grupo in sbs.groupby(
        "periodo",
        sort=True
    ):

        grupo = grupo.sort_values(
            ["id_empresa", "empresa_sbs"]
        )

        detalle = "; ".join(
            f"{fila.id_empresa}: {fila.empresa_sbs}"
            for fila in grupo.itertuples()
        )

        filas_mes.append({
            "periodo": periodo,
            "n_empresas": (
                grupo["id_empresa"].nunique()
            ),
            "empresas": detalle,
        })

    por_mes = pd.DataFrame(filas_mes)

    por_mes.to_csv(
        CARPETA_SALIDAS / "empresas_por_mes.csv",
        index=False,
        encoding="utf-8"
    )

    reportar(
        f"\nEmpresas por mes: mínimo "
        f"{por_mes['n_empresas'].min()}, "
        f"máximo {por_mes['n_empresas'].max()}."
    )

    reportar(
        "  Salidas: presencia_empresas.csv, "
        "empresas_por_mes.csv y "
        "trayectoria_empresarial.csv"
    )


# =============================================================================
# BLOQUE 8. IDENTIDAD EMPRESARIAL Y HOMOLOGACIÓN DESDE EVIDENCIA SBS
# =============================================================================
# OBJETIVO:
# Construir un id_empresa estable utilizando evidencia de la SBS,
# sin homologaciones manuales ni equivalencias inventadas.
#
# PRINCIPIOS:
# - Un mismo nombre no implica automáticamente la misma identidad jurídica.
# - Un cambio de denominación puede conservar la identidad.
# - Una fusión/absorción NO une trayectorias retrospectivamente.
# - Los registros especiales "(abs)" se usan como evidencia, no como empresas.
# =============================================================================


# =============================================================================
# PASO 1. CLASIFICAR LOS EVENTOS SOCIETARIOS
# =============================================================================
def clasificar_evento(nota_normalizada):
    """
    Clasifica una nota SBS según el evento societario descrito.
    """

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
# PASO 2. FUNCIONES AUXILIARES PARA FECHAS Y REFERENCIAS
# =============================================================================
def periodo_resolucion(texto):
    """
    Extrae una fecha dd/mm/aaaa o dd-mm-aaaa de una nota SBS.
    Devuelve YYYY-MM.
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
    """
    Devuelve el mes inmediatamente anterior.
    Ejemplo: 2022-06 -> 2022-05.
    """

    return str(
        pd.Period(
            str(periodo),
            freq="M"
        ) - 1
    )


def refs_a_set(valor):
    """
    Convierte una referencia como '1;2' en {'1', '2'}.
    """

    if pd.isna(valor):
        return set()

    return {
        x.strip()
        for x in str(valor).split(";")
        if x.strip()
    }


# =============================================================================
# PASO 3. IDENTIFICAR EMPRESAS MENCIONADAS EN LAS NOTAS SBS
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
    que realmente aparece en los encabezados SBS.
    """

    objetivo = sin_puntuacion(texto)

    # 3.1. Primero se intenta coincidencia exacta.
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

    # 3.2. Como respaldo, se comparan secuencias de palabras.
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

    # Si los dos candidatos principales son igual de fuertes,
    # se considera ambiguo y no se decide automáticamente.
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
# PASO 4. VINCULAR NOTAS SBS CON SUS ENCABEZADOS
# =============================================================================
def encabezados_referenciados(
    grupo_notas,
    columnas_df
):
    """
    Usa referencias (1), (2), etc. para saber qué encabezados
    están relacionados con una nota.
    """

    encontrados = set()

    for _, nota in grupo_notas.iterrows():

        refs_nota = refs_a_set(
            nota[
                "referencias_detectadas"
            ]
        )

        if not refs_nota:
            continue

        columnas = columnas_df[
            columnas_df[
                "periodo"
            ].eq(nota["periodo"])
            & columnas_df[
                "hoja"
            ].eq(nota["hoja"])
        ]

        for _, columna in columnas.iterrows():

            refs_columna = refs_a_set(
                columna[
                    "referencia_nota"
                ]
            )

            if refs_nota & refs_columna:

                encontrados.add(
                    columna[
                        "nombre_limpio"
                    ]
                )

    return sorted(encontrados)


# =============================================================================
# PASO 5. COMPROBAR REGISTROS ESPECIALES DE EMPRESAS ABSORBIDAS
# =============================================================================
def nombre_base_abs(nombre):
    """
    Elimina solo la marca '(abs)' para comparar nombres.

    Ejemplo:
    'Mapfre Perú (abs)' -> 'MAPFRE PERU'
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
    Comprueba si existe un registro especial '(abs)'
    relacionado con la empresa vigente en el mes de la fusión.
    """

    if (
        especiales is None
        or especiales.empty
    ):
        return False

    candidatos = especiales[
        especiales[
            "periodo"
        ].eq(periodo_transicion)
        & especiales[
            "tipo_definitivo"
        ].eq(
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
# PASO 6. EVALUAR LOS EVENTOS SOCIETARIOS
# =============================================================================
def evaluar_eventos(
    sbs,
    notas,
    columnas_df,
    especiales=None
):
    """
    Decide qué eventos solamente se documentan y cuáles realmente
    modifican la continuidad de una identidad empresarial.
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

        nota_norm = normalizar(texto)

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

        refs_enc = (
            encabezados_referenciados(
                grupo,
                columnas_df
            )
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

            "primer_periodo_nota": (
                min(periodos_nota)
                if periodos_nota
                else None
            ),

            "ultimo_periodo_nota": (
                max(periodos_nota)
                if periodos_nota
                else None
            ),

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

        # =====================================================================
        # PASO 6A. CAMBIO DE DENOMINACIÓN
        # =====================================================================
        if tipo == "cambio_de_denominacion":

            mencionados = (
                nombres_mencionados(
                    texto,
                    nombres
                )
            )

            pares = []

            for anterior in mencionados:

                for nuevo in mencionados:

                    if anterior == nuevo:
                        continue

                    fin_anterior = (
                        presencia.loc[
                            anterior,
                            "ultimo_periodo"
                        ]
                    )

                    inicio_nuevo = (
                        presencia.loc[
                            nuevo,
                            "primer_periodo"
                        ]
                    )

                    # Solo se acepta continuidad exacta:
                    # antiguo termina en t y nuevo empieza en t+1.
                    if (
                        periodo_a_numero(
                            inicio_nuevo
                        )
                        - periodo_a_numero(
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
                    "Cambio de denominación "
                    "documentado por SBS y "
                    "continuidad temporal exacta."
                )

            elif len(pares) == 0:

                registro["motivo"] = (
                    "La nota indica cambio de "
                    "denominación, pero no existe "
                    "una pareja única de nombres "
                    "con continuidad temporal exacta."
                )

            else:

                registro["motivo"] = (
                    "Cambio de denominación ambiguo: "
                    "más de una pareja temporal posible."
                )

            eventos.append(registro)

            continue

        # =====================================================================
        # PASO 6B. FUSIÓN O ABSORCIÓN
        # =====================================================================
        if tipo == "fusion_absorcion":

            texto_limpio = (
                sin_puntuacion(texto)
            )

            # Se exige una formulación fuerte y explícita de la SBS.
            #
            # Ejemplo real:
            # "Las cifras de X consideran los flujos de Y
            # para el periodo previo a la fusión."
            patron = re.search(
                r"LAS CIFRAS DE (.+?) "
                r"CONSIDERAN LOS FLUJOS DE "
                r"(.+?) "
                r"PARA EL PERIODO PREVIO "
                r"A LA FUSION",
                texto_limpio
            )

            # Si no existe esta formulación fuerte,
            # la fusión se documenta pero no se aplica.
            if not patron:

                registro["motivo"] = (
                    "Fusión/absorción documentada "
                    "por SBS; se registra, pero "
                    "no se fuerza continuidad."
                )

                eventos.append(registro)

                continue

            actual = (
                resolver_nombre_capturado(
                    patron.group(1),
                    nombres
                )
            )

            previo = (
                resolver_nombre_capturado(
                    patron.group(2),
                    nombres
                )
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

                registro["motivo"] = (
                    "No fue posible resolver "
                    "de forma única las empresas "
                    "mencionadas en la nota de fusión."
                )

                eventos.append(registro)

                continue

            mes_previo = periodo_anterior(
                transicion
            )

            # Empresa que continúa debe existir en el mes de transición.
            existe_actual = (
                sbs["periodo"].eq(
                    transicion
                )
                & sbs["empresa_sbs"].eq(
                    actual
                )
            ).any()

            # Empresa previa debe existir en el mes inmediatamente anterior.
            existe_previo = (
                sbs["periodo"].eq(
                    mes_previo
                )
                & sbs["empresa_sbs"].eq(
                    previo
                )
            ).any()

            # Se comprueba si el mismo nombre de la empresa actual
            # ya existía antes de la fusión.
            existe_actual_mes_previo = (
                sbs["periodo"].eq(
                    mes_previo
                )
                & sbs["empresa_sbs"].eq(
                    actual
                )
            ).any()

            # También debe existir la evidencia especial "(abs)".
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

                # Solo se rompe la trayectoria del mismo nombre
                # si ese nombre ya existía antes de la fusión.
                registro[
                    "rompe_mismo_nombre"
                ] = bool(
                    existe_actual_mes_previo
                )

                registro[
                    "motivo"
                ] = (
                    "Fusión con continuidad explícita "
                    "respaldada por SBS: la empresa "
                    "vigente incorpora los flujos de "
                    "la entidad previa y existe registro "
                    "especial '(abs)' de la absorbida."
                )

            else:

                registro[
                    "motivo"
                ] = (
                    "La nota sugiere continuidad "
                    "por fusión, pero faltan "
                    "condiciones de verificación "
                    "temporal o evidencia '(abs)'. "
                    "No se aplica."
                )

            eventos.append(registro)

            continue

        # =====================================================================
        # PASO 6C. OTROS EVENTOS
        # =====================================================================
        registro[
            "motivo"
        ] = (
            "Evento societario documentado "
            "por SBS; no modifica "
            "automáticamente la identidad."
        )

        eventos.append(registro)

    return pd.DataFrame(
        eventos,
        columns=columnas_eventos
    )


# =============================================================================
# PASO 7. CONSTRUIR EL id_empresa
# =============================================================================
def construir_id_empresa(
    sbs,
    eventos
):
    """
    Construye una identidad estable para cada trayectoria empresarial.
    """

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

    if (
        len(claves)
        != len(set(claves))
    ):

        raise RuntimeError(
            "Hay duplicados periodo × empresa_sbs "
            "antes de construir id_empresa."
        )

    # -------------------------------------------------------------------------
    # PASO 7A. Preparar estructura de componentes
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # PASO 7B. Identificar rupturas documentadas del mismo nombre
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
    # PASO 7C. Mantener continuidad del mismo nombre
    # -------------------------------------------------------------------------
    # Si una empresa tiene exactamente el mismo nombre y existe un hueco
    # de uno o más meses, el hueco se audita por separado.
    #
    # NO se crea automáticamente una empresa nueva por ese hueco,
    # salvo que exista evidencia SBS de ruptura de identidad.
    # -------------------------------------------------------------------------
    for nombre, grupo in p.groupby(
        "empresa_sbs",
        sort=True
    ):

        periodos = (
            grupo
            .sort_values("periodo")[
                "periodo"
            ]
            .tolist()
        )

        for anterior, actual in zip(
            periodos[:-1],
            periodos[1:]
        ):

            # Si SBS documentó una ruptura societaria,
            # NO se unen ambos lados.
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
    # PASO 7D. Unir nombres diferentes cuando SBS documenta continuidad
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

            mes_previo = (
                periodo_anterior(
                    transicion
                )
            )

            clave_origen = (
                mes_previo,
                origen
            )

            clave_destino = (
                transicion,
                destino
            )

            if (
                clave_origen
                not in padre
            ):

                raise RuntimeError(
                    "La continuidad documentada "
                    "no encuentra "
                    f"{origen} en {mes_previo}."
                )

            if (
                clave_destino
                not in padre
            ):

                raise RuntimeError(
                    "La continuidad documentada "
                    "no encuentra "
                    f"{destino} en {transicion}."
                )

            unir(
                clave_origen,
                clave_destino
            )

    # -------------------------------------------------------------------------
    # PASO 7E. Construir los componentes de identidad
    # -------------------------------------------------------------------------
    componentes = {}

    for clave in claves:

        raiz = buscar(clave)

        componentes.setdefault(
            raiz,
            []
        ).append(clave)

    # -------------------------------------------------------------------------
    # PASO 7F. Ordenar los IDs de forma reproducible
    # -------------------------------------------------------------------------
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

    p[
        "id_empresa"
    ] = [
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

    # -------------------------------------------------------------------------
    # PASO 7G. Mantener el nombre histórico observado en cada mes
    # -------------------------------------------------------------------------
    # MUY IMPORTANTE:
    # No reemplazamos retrospectivamente todos los nombres por el nombre final.
    #
    # Ejemplo:
    # antes de una fusión puede decir MAPFRE PERÚ VIDA,
    # y después MAPFRE PERÚ.
    #
    # El vínculo histórico lo da id_empresa.
    # -------------------------------------------------------------------------
    p[
        "empresa"
    ] = p[
        "empresa_sbs"
    ]

    return p


# =============================================================================
# PASO 8. CONTROLAR LA CONSISTENCIA DE LOS id_empresa
# =============================================================================
def controlar_identidades(
    sbs,
    eventos
):
    """
    Comprueba que los IDs empresariales sean coherentes.
    """

    # 8.1. Ninguna fila puede quedarse sin ID.
    if (
        sbs[
            "id_empresa"
        ].isna().any()
    ):

        raise RuntimeError(
            "Hay filas sin id_empresa."
        )

    # 8.2. Una identidad no puede aparecer dos veces en el mismo mes.
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

    # 8.3. Toda continuidad aplicada debe conservar el mismo ID.
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

        mes_previo = (
            periodo_anterior(
                transicion
            )
        )

        id_origen = sbs.loc[
            sbs["periodo"].eq(
                mes_previo
            )
            & sbs["empresa_sbs"].eq(
                origen
            ),
            "id_empresa"
        ].unique()

        id_destino = sbs.loc[
            sbs["periodo"].eq(
                transicion
            )
            & sbs["empresa_sbs"].eq(
                destino
            ),
            "id_empresa"
        ].unique()

        if (
            len(id_origen) != 1
            or len(id_destino) != 1
            or id_origen[0]
            != id_destino[0]
        ):

            raise RuntimeError(
                "Una continuidad SBS aplicada "
                "no conserva el mismo id_empresa."
            )

        # 8.4. Si la fusión exige romper el mismo nombre,
        # la empresa anterior con ese nombre debe tener otro ID.
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
                    "la trayectoria previa del mismo "
                    "nombre terminó usando el mismo ID."
                )


# =============================================================================
# PASO 9. GUARDAR LA AUDITORÍA DE EVENTOS
# =============================================================================
def guardar_eventos_identidad(
    eventos
):
    """
    Guarda los eventos societarios detectados y la evidencia SBS.
    """

    eventos.to_csv(
        CARPETA_SALIDAS
        / "homologacion_empresas.csv",
        index=False,
        encoding="utf-8"
    )


# =============================================================================
# PASO 10. FUNCIÓN PRINCIPAL DE HOMOLOGACIÓN
# =============================================================================
def homologar(
    sbs,
    notas,
    columnas_df,
    especiales=None,
    guardar_auditoria=True
):
    """
    Ejecuta todo el proceso de identidad empresarial.
    """

    # 10.1. Detectar eventos.
    eventos = evaluar_eventos(
        sbs,
        notas,
        columnas_df,
        especiales=especiales
    )

    # 10.2. Construir los id_empresa.
    sbs = construir_id_empresa(
        sbs,
        eventos
    )

    # 10.3. Ejecutar controles.
    controlar_identidades(
        sbs,
        eventos
    )

    # 10.4. Guardar auditoría únicamente en ejecución definitiva.
    if guardar_auditoria:

        guardar_eventos_identidad(
            eventos
        )

    # 10.5. Reportar resultados.
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
                "  "
                f"{evento['continuidad_desde']} "
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
# RECORDATORIO PARA LA BASE FINAL
# =============================================================================
# id_empresa:
#   identifica la trayectoria de una empresa y puede repetirse durante meses.
#
# id_observacion:
#   NO se crea en este bloque.
#   Se agregará más adelante cuando ordenemos la base final:
#
#       1, 2, 3, ..., N
#
#   con un valor único para cada observación empresa × mes.
# =============================================================================


# =============================================================================
# BLOQUE 8B. VALIDACIÓN Y CONSOLIDACIÓN DE EVENTOS SOCIETARIOS
# =============================================================================
# Este bloque añade controles finales al resultado del Bloque 8.
#
# PASO 1. Rechaza continuidades absurdas del tipo A -> A.
# PASO 2. Exige que origen y destino estén realmente observados en la ventana.
# PASO 3. Exige continuidad temporal entre t y t+1.
# PASO 4. Consolida notas duplicadas del mismo evento societario.
# PASO 5. Conserva toda la evidencia SBS para auditoría.
# =============================================================================


# =============================================================================
# PASO 1. VALIDAR UNA CONTINUIDAD PROPUESTA
# =============================================================================
def validar_continuidad_evento(evento, sbs):
    """
    Revisa si una continuidad propuesta puede aplicarse realmente
    dentro de la ventana 2020-01 a 2025-12.
    """

    evento = evento.copy()

    if not bool(evento.get("aplicado", False)):
        return evento

    origen = evento.get("continuidad_desde")
    destino = evento.get("continuidad_hacia")
    transicion = evento.get("periodo_transicion")

    # -------------------------------------------------------------------------
    # 1A. Deben existir origen, destino y periodo.
    # -------------------------------------------------------------------------
    if (
        pd.isna(origen)
        or pd.isna(destino)
        or pd.isna(transicion)
        or not str(origen).strip()
        or not str(destino).strip()
        or not str(transicion).strip()
    ):
        evento["aplicado"] = False
        evento["motivo"] = (
            "Evento no aplicado: faltan origen, destino "
            "o periodo de transición verificables."
        )
        return evento

    # -------------------------------------------------------------------------
    # 1B. Una empresa no puede homologarse consigo misma.
    # -------------------------------------------------------------------------
    if (
        sin_puntuacion(origen)
        == sin_puntuacion(destino)
    ):
        evento["aplicado"] = False
        evento["motivo"] = (
            "Evento documentado, pero no aplicado: "
            "el origen y el destino identificados son el mismo nombre. "
            "No genera una nueva continuidad dentro de la ventana."
        )
        return evento

    # -------------------------------------------------------------------------
    # 1C. El origen debe existir realmente en la base.
    # -------------------------------------------------------------------------
    existe_origen = (
        sbs["empresa_sbs"]
        .eq(origen)
        .any()
    )

    # -------------------------------------------------------------------------
    # 1D. El destino debe existir realmente en la base.
    # -------------------------------------------------------------------------
    existe_destino = (
        sbs["empresa_sbs"]
        .eq(destino)
        .any()
    )

    if not existe_origen or not existe_destino:
        evento["aplicado"] = False
        evento["motivo"] = (
            "Evento societario documentado, pero no aplicado: "
            "el nombre anterior o el posterior no está observado "
            "dentro de la ventana 2020-2025."
        )
        return evento

    # -------------------------------------------------------------------------
    # 1E. Verificar continuidad exacta alrededor de la transición.
    # -------------------------------------------------------------------------
    mes_previo = periodo_anterior(
        transicion
    )

    origen_previo = (
        sbs["periodo"].eq(mes_previo)
        & sbs["empresa_sbs"].eq(origen)
    ).any()

    destino_actual = (
        sbs["periodo"].eq(transicion)
        & sbs["empresa_sbs"].eq(destino)
    ).any()

    if not origen_previo or not destino_actual:
        evento["aplicado"] = False
        evento["motivo"] = (
            "Evento societario documentado, pero no aplicado: "
            "no se verifica que el origen exista en el mes anterior "
            "y el destino en el mes de transición."
        )
        return evento

    return evento


# =============================================================================
# PASO 2. VALIDAR TODOS LOS EVENTOS
# =============================================================================
def validar_eventos_aplicados(eventos, sbs):
    """
    Aplica los controles anteriores a todos los eventos detectados.
    """

    if eventos is None or eventos.empty:
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
# PASO 3. CONSOLIDAR EVENTOS DUPLICADOS
# =============================================================================
def consolidar_eventos(eventos):
    """
    Si dos o más notas SBS describen exactamente la misma continuidad,
    se conserva una sola fila de evento y se acumula su evidencia.

    Ejemplo:
    SECREX -> CESCE PERÚ
    puede aparecer en varias notas, pero es una sola transición societaria.
    """

    if eventos is None or eventos.empty:
        return eventos.copy()

    no_aplicados = eventos[
        ~eventos["aplicado"].eq(True)
    ].copy()

    aplicados = eventos[
        eventos["aplicado"].eq(True)
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

        fila = grupo.iloc[0].copy()

        # ---------------------------------------------------------------------
        # 3A. Combinar evidencia textual distinta.
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

        fila["evidencia_sbs"] = (
            " || ".join(evidencias)
        )

        # ---------------------------------------------------------------------
        # 3B. Combinar archivos de origen.
        # ---------------------------------------------------------------------
        archivos = []

        for valor in grupo[
            "archivo_origen"
        ].dropna():

            for archivo in str(valor).split(";"):

                archivo = archivo.strip()

                if (
                    archivo
                    and archivo not in archivos
                ):
                    archivos.append(archivo)

        fila["archivo_origen"] = (
            ";".join(archivos)
        )

        # ---------------------------------------------------------------------
        # 3C. Combinar referencias de encabezado.
        # ---------------------------------------------------------------------
        referencias = []

        for valor in grupo[
            "referencias_detectadas"
        ].fillna(""):

            for ref in str(valor).split(";"):

                ref = ref.strip()

                if (
                    ref
                    and ref not in referencias
                ):
                    referencias.append(ref)

        fila[
            "referencias_detectadas"
        ] = ";".join(referencias)

        # ---------------------------------------------------------------------
        # 3D. Registrar cuántas notas respaldan el mismo evento.
        # ---------------------------------------------------------------------
        fila["apariciones_nota"] = int(
            grupo[
                "apariciones_nota"
            ].fillna(0).sum()
        )

        consolidados.append(fila)

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
# PASO 4. CONTROL FINAL DE EVENTOS
# =============================================================================
def control_final_eventos(eventos):
    """
    Impide que sobrevivan errores lógicos en las continuidades aplicadas.
    """

    if eventos is None or eventos.empty:
        return

    aplicados = eventos[
        eventos["aplicado"].eq(True)
    ]

    # -------------------------------------------------------------------------
    # 4A. Nunca A -> A.
    # -------------------------------------------------------------------------
    mismo_nombre = aplicados.apply(
        lambda fila:
        sin_puntuacion(
            fila["continuidad_desde"]
        )
        == sin_puntuacion(
            fila["continuidad_hacia"]
        ),
        axis=1
    )

    if mismo_nombre.any():

        raise RuntimeError(
            "Existe una continuidad inválida "
            "del tipo empresa -> misma empresa."
        )

    # -------------------------------------------------------------------------
    # 4B. No puede haber duplicados de la misma transición.
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
# PASO 5. REEMPLAZAR HOMOLOGAR POR LA VERSIÓN CONTROLADA
# =============================================================================
def homologar_controlado(
    sbs,
    notas,
    columnas_df,
    especiales=None,
    guardar_auditoria=True
):
    """
    Versión definitiva de la homologación:
    detecta -> valida -> consolida -> crea id_empresa -> controla.
    """

    # 5.1. Detectar eventos con el Bloque 8.
    eventos = evaluar_eventos(
        sbs,
        notas,
        columnas_df,
        especiales=especiales
    )

    # 5.2. Invalidar continuidades que no son verificables.
    eventos = validar_eventos_aplicados(
        eventos,
        sbs
    )

    # 5.3. Consolidar duplicados documentales.
    eventos = consolidar_eventos(
        eventos
    )

    # 5.4. Control lógico previo.
    control_final_eventos(
        eventos
    )

    # 5.5. Crear las identidades empresariales.
    sbs = construir_id_empresa(
        sbs,
        eventos
    )

    # 5.6. Verificar las identidades resultantes.
    controlar_identidades(
        sbs,
        eventos
    )

    # 5.7. Guardar auditoría solo en ejecución definitiva.
    if guardar_auditoria:

        guardar_eventos_identidad(
            eventos
        )

    aplicados = eventos[
        eventos["aplicado"].eq(True)
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
            "  "
            f"{evento['continuidad_desde']} "
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

    return sbs, eventos


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
# BLOQUE 12. DIAGNÓSTICOS
# No elimina ni modifica observaciones.
# =============================================================================

def categorias(serie):
    """
    Clasifica valores para diagnóstico.
    """

    return {
        "positivos_>=1":
            int(
                (serie >= 1).sum()
            ),

        "entre_0_y_1":
            int(
                (
                    (serie > 0)
                    & (serie < 1)
                ).sum()
            ),

        "ceros":
            int(
                (serie == 0).sum()
            ),

        "negativos":
            int(
                (serie < 0).sum()
            ),

        "faltantes":
            int(
                serie.isna().sum()
            ),
    }


def detectar_outliers_iqr(
    base,
    columnas
):
    """
    Señala outliers mediante 1.5 × IQR.
    No elimina ni winsoriza observaciones.
    """

    filas = []

    for col in columnas:

        serie = (
            base[col]
            .dropna()
        )

        if len(serie) < 4:
            continue

        q1, q3 = (
            serie.quantile(
                [0.25, 0.75]
            )
        )

        iqr = q3 - q1

        inferior = (
            q1 - 1.5 * iqr
        )

        superior = (
            q3 + 1.5 * iqr
        )

        mask = (
            base[col].notna()
            & (
                (base[col] < inferior)
                | (base[col] > superior)
            )
        )

        columnas_id = [
            c
            for c in [
                "id_observacion",
                "periodo",
                "id_empresa",
                "empresa",
                col,
            ]
            if c in base.columns
        ]

        for _, fila in base.loc[
            mask,
            columnas_id
        ].iterrows():

            registro = {
                "periodo":
                    fila["periodo"],

                "id_empresa":
                    fila["id_empresa"],

                "empresa":
                    fila["empresa"],

                "variable":
                    col,

                "valor":
                    fila[col],

                "q1":
                    q1,

                "q3":
                    q3,

                "limite_inferior":
                    inferior,

                "limite_superior":
                    superior,
            }

            if (
                "id_observacion"
                in fila.index
            ):

                registro[
                    "id_observacion"
                ] = fila[
                    "id_observacion"
                ]

            filas.append(
                registro
            )

    return pd.DataFrame(
        filas
    )


def diagnosticar(
    base,
    largo,
    guardar_detalle=True
):
    """
    Reporta formatos, ceros, negativos,
    faltantes y outliers.
    """

    # -------------------------------------------------------------------------
    # PASO 1. Formato original SBS.
    # -------------------------------------------------------------------------
    reportar(
        "\nCómo venían las celdas TOTAL en los XLS:"
    )

    for tipo, n in (
        largo[
            "tipo_valor"
        ]
        .value_counts(
            dropna=False
        )
        .items()
    ):

        reportar(
            f"  {tipo}: {n}"
        )

    # -------------------------------------------------------------------------
    # PASO 2. Diagnóstico cuantitativo.
    # -------------------------------------------------------------------------
    reportar(
        "\nDiagnóstico por variable "
        "(miles de S/ corrientes):"
    )

    cols = [
        "primas_netas_acum_sbs",
        "primas_netas_mensual_nominal",
        "primas_cedidas_acum_sbs",
        "primas_cedidas_mensual_nominal",
    ]

    for col in cols:

        resultado = categorias(
            base[col]
        )

        reportar(
            f"  {col:<34} "
            + " | ".join(
                f"{k}={v}"
                for k, v
                in resultado.items()
            )
        )

    # -------------------------------------------------------------------------
    # PASO 3. Negativos.
    # -------------------------------------------------------------------------
    negativos = base[
        (
            base[
                "primas_netas_mensual_nominal"
            ] < 0
        )
        |
        (
            base[
                "primas_cedidas_mensual_nominal"
            ] < 0
        )
    ].copy()

    reportar(
        "\nFlujos mensuales negativos: "
        f"{len(negativos):,} "
        "(se conservan; pueden representar "
        "ajustes contables)."
    )

    if not negativos.empty:

        for _, fila in (
            negativos
            .sort_values(
                ["empresa", "periodo"]
            )
            .head(10)
            .iterrows()
        ):

            reportar(
                f"  Ejemplo: "
                f"{fila['periodo']} | "
                f"{fila['empresa']} | "
                f"netas="
                f"{fila['primas_netas_mensual_nominal']} | "
                f"cedidas="
                f"{fila['primas_cedidas_mensual_nominal']}"
            )

    # -------------------------------------------------------------------------
    # PASO 4. Outliers IQR.
    # -------------------------------------------------------------------------
    outliers = detectar_outliers_iqr(
        base,
        [
            "primas_netas_mensual_nominal",
            "primas_cedidas_mensual_nominal",
        ]
    )

    reportar(
        f"Outliers IQR señalados: "
        f"{len(outliers):,}; "
        "no se eliminan ni winsorizan."
    )

    # -------------------------------------------------------------------------
    # PASO 5. Guardar diagnósticos.
    # -------------------------------------------------------------------------
    if guardar_detalle:

        negativos.to_csv(
            CARPETA_SALIDAS
            / "diagnostico_negativos.csv",
            index=False,
            encoding="utf-8"
        )

        if outliers.empty:

            outliers = pd.DataFrame(
                columns=[
                    "id_observacion",
                    "periodo",
                    "id_empresa",
                    "empresa",
                    "variable",
                    "valor",
                    "q1",
                    "q3",
                    "limite_inferior",
                    "limite_superior",
                ]
            )

        outliers.to_csv(
            CARPETA_SALIDAS
            / "diagnostico_outliers.csv",
            index=False,
            encoding="utf-8"
        )

    # -------------------------------------------------------------------------
    # PASO 6. Tamaño de muestra.
    # -------------------------------------------------------------------------
    reportar(
        f"\nObservaciones en la base: "
        f"{len(base):,}"
    )

    reportar(
        "muestra_modelo=True: "
        f"{int(base['muestra_modelo'].sum()):,}"
    )

    reportar(
        "muestra_modelo=False: "
        f"{int((~base['muestra_modelo']).sum()):,}"
    )


# =============================================================================
# BLOQUE 13. CONTROLES FINALES
# =============================================================================

def controles_finales(base):
    """
    Verifica estructura, llaves, cobertura temporal
    y tamaño mínimo de la base procesada.
    """

    reportar(
        "\nControles finales:"
    )

    # -------------------------------------------------------------------------
    # PASO 1. Exactamente 72 meses.
    # -------------------------------------------------------------------------
    encontrados = sorted(
        base[
            "periodo"
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    esperados = (
        periodos_esperados()
    )

    if encontrados != esperados:

        raise RuntimeError(
            "La base final no contiene "
            "exactamente los 72 meses. "
            f"Faltan="
            f"{sorted(set(esperados) - set(encontrados))}; "
            f"sobran="
            f"{sorted(set(encontrados) - set(esperados))}"
        )

    # -------------------------------------------------------------------------
    # PASO 2. Identificadores.
    # -------------------------------------------------------------------------
    identificadores = [
        "id_observacion",
        "periodo",
        "id_empresa",
        "empresa",
        "empresa_sbs",
    ]

    faltan_columnas = [
        col
        for col in identificadores
        if col not in base.columns
    ]

    if faltan_columnas:

        raise RuntimeError(
            "Faltan columnas identificadoras: "
            + ", ".join(
                faltan_columnas
            )
        )

    if (
        base[
            identificadores
        ]
        .isna()
        .any()
        .any()
    ):

        raise RuntimeError(
            "Hay identificadores vacíos "
            "en la base final."
        )

    # -------------------------------------------------------------------------
    # PASO 3. id_observacion único y correlativo.
    # -------------------------------------------------------------------------
    if (
        base[
            "id_observacion"
        ].duplicated().any()
    ):

        raise RuntimeError(
            "Hay id_observacion duplicados."
        )

    esperado_id = list(
        range(
            1,
            len(base) + 1
        )
    )

    if (
        base[
            "id_observacion"
        ].tolist()
        != esperado_id
    ):

        raise RuntimeError(
            "id_observacion no es "
            "correlativo de 1 a N."
        )

    reportar(
        f"  id_observacion: "
        f"1 a {len(base):,}, "
        "sin duplicados."
    )

    # -------------------------------------------------------------------------
    # PASO 4. Un id_empresa por periodo.
    # -------------------------------------------------------------------------
    duplicados = int(
        base.duplicated(
            [
                "periodo",
                "id_empresa"
            ]
        ).sum()
    )

    reportar(
        "  Duplicados "
        "id_empresa × periodo: "
        f"{duplicados}"
    )

    if duplicados:

        raise RuntimeError(
            "Hay duplicados "
            "id_empresa × periodo."
        )

    # -------------------------------------------------------------------------
    # PASO 5. Balance del panel.
    # -------------------------------------------------------------------------
    meses_por_empresa = (
        base.groupby(
            "id_empresa"
        )["periodo"]
        .nunique()
    )

    balanceado = bool(
        (
            meses_por_empresa
            == MESES_ESPERADOS
        ).all()
    )

    reportar(
        f"  Periodo: "
        f"{FECHA_INICIO} a "
        f"{FECHA_CORTE} | meses=72"
    )

    reportar(
        "  Identidades empresariales: "
        f"{base['id_empresa'].nunique()} | "
        f"panel "
        f"{'balanceado' if balanceado else 'NO balanceado'} "
        f"(mínimo="
        f"{meses_por_empresa.min()} meses; "
        f"máximo="
        f"{meses_por_empresa.max()})"
    )

    # -------------------------------------------------------------------------
    # PASO 6. Mínimo de observaciones.
    # -------------------------------------------------------------------------
    reportar(
        f"  Observaciones: "
        f"{len(base):,} "
        f"(mínimo exigido: "
        f"{MINIMO_OBSERVACIONES:,})"
    )

    if (
        len(base)
        < MINIMO_OBSERVACIONES
    ):

        raise RuntimeError(
            "La base procesada tiene "
            f"menos de "
            f"{MINIMO_OBSERVACIONES} "
            "observaciones."
        )

    if base.shape[1] < 6:

        raise RuntimeError(
            "La base final tiene "
            "menos de 6 columnas."
        )

    # -------------------------------------------------------------------------
    # PASO 7. Variables sustantivas.
    # -------------------------------------------------------------------------
    sustantivas = [
        "log_primas_reales",
        "log_primas_cedidas",
        "log_ingreso_real",
        "tasa_referencia",
    ]

    faltan_sustantivas = [
        col
        for col in sustantivas
        if col not in base.columns
    ]

    if faltan_sustantivas:

        raise RuntimeError(
            "Faltan variables sustantivas: "
            + ", ".join(
                faltan_sustantivas
            )
        )

    # -------------------------------------------------------------------------
    # PASO 8. Verificar presencia de SBS y BCRP.
    # -------------------------------------------------------------------------
    if not {
        "primas_netas_acum_sbs",
        "primas_cedidas_acum_sbs",
    }.issubset(
        base.columns
    ):

        raise RuntimeError(
            "No están presentes "
            "las variables fuente SBS."
        )

    if not {
        "ingreso_formal_nominal",
        "tasa_referencia",
        "ipc",
    }.issubset(
        base.columns
    ):

        raise RuntimeError(
            "No están presentes "
            "las variables fuente BCRP."
        )

    # -------------------------------------------------------------------------
    # PASO 9. Muestra del modelo.
    # -------------------------------------------------------------------------
    muestra = int(
        base[
            "muestra_modelo"
        ].sum()
    )

    reportar(
        f"  muestra_modelo=True: "
        f"{muestra:,}"
    )

    if (
        muestra
        < MINIMO_OBSERVACIONES
    ):

        reportar(
            "  AVISO: muestra_modelo "
            f"tiene {muestra:,} observaciones; "
            "el requisito de 1,000 "
            "se evalúa sobre la base procesada."
        )

    # -------------------------------------------------------------------------
    # PASO 10. Faltantes.
    # -------------------------------------------------------------------------
    reportar(
        "  No vacíos / faltantes por variable:"
    )

    for col in base.columns:

        reportar(
            f"    {col:<34} "
            f"no vacíos="
            f"{int(base[col].notna().sum()):>6,} | "
            f"faltantes="
            f"{int(base[col].isna().sum()):>5,}"
        )


# =============================================================================
# BLOQUE 14. GUARDADO, LOG Y SHA-256
# =============================================================================

def escribir_log(texto):
    """
    Añade una línea al log con fecha y hora.
    """

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
    Guarda la base procesada,
    calcula SHA-256 y registra la ejecución.
    """

    # -------------------------------------------------------------------------
    # PASO 1. Guardar CSV.
    # -------------------------------------------------------------------------
    ruta = (
        CARPETA_PROC
        / f"datos_procesados_{MATRICULA}.csv"
    )

    base.to_csv(
        ruta,
        index=False,
        encoding="utf-8",
        float_format="%.6f"
    )

    # -------------------------------------------------------------------------
    # PASO 2. Calcular SHA-256.
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # PASO 3. Registrar log.
    # -------------------------------------------------------------------------
    escribir_log(
        f"PROCESADO {ruta.name} | "
        f"filas={len(base)} | "
        f"columnas={base.shape[1]} | "
        f"muestra_modelo="
        f"{int(base['muestra_modelo'].sum())} | "
        f"sha256={huella}"
    )

    # -------------------------------------------------------------------------
    # PASO 4. Guardar informe de controles.
    # -------------------------------------------------------------------------
    ARCHIVO_CONTROL.write_text(
        "\n".join(REPORTE) + "\n",
        encoding="utf-8"
    )

    print(
        "Informe de control: "
        "salidas/control_03_limpieza.txt"
    )


# =============================================================================
# BLOQUE 15. PRUEBAS PREVIAS Y EJECUCIÓN PRINCIPAL
# =============================================================================

def preparar_base(guardar_auditoria):
    """
    Ejecuta el pipeline completo del Script 03.
    """

    # -------------------------------------------------------------------------
    # PASO 1. BCRP.
    # -------------------------------------------------------------------------
    bcrp = cargar_bcrp()

    # -------------------------------------------------------------------------
    # PASO 2. SBS.
    # -------------------------------------------------------------------------
    (
        sbs,
        largo,
        notas,
        columnas_df,
        comparacion
    ) = cargar_sbs(
        guardar_auditoria=
        guardar_auditoria
    )

    if guardar_auditoria:

        guardar_notas(
            notas
        )

    # -------------------------------------------------------------------------
    # PASO 3. Registros especiales SBS.
    # Ejemplo: MAPFRE Perú (abs).
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
        guardar_auditoria=
        guardar_auditoria
    )

    # Sobrescribe la comparación inicial
    # con la comparación después de depurar especiales.
    if guardar_auditoria:

        comparacion.to_csv(
            CARPETA_SALIDAS
            / "comparacion_p009_p010.csv",
            index=False,
            encoding="utf-8"
        )

    # -------------------------------------------------------------------------
    # PASO 4. Identidad empresarial.
    # Usa Bloque 8 + control Bloque 8B.
    # -------------------------------------------------------------------------
    (
        sbs,
        eventos
    ) = homologar_controlado(
        sbs,
        notas,
        columnas_df,
        especiales=
        registros_especiales,
        guardar_auditoria=
        guardar_auditoria
    )

    if guardar_auditoria:

        guardar_presencia_y_trayectoria(
            sbs,
            eventos
        )

    # -------------------------------------------------------------------------
    # PASO 5. Desacumulación.
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
    # PASO 6. Integrar SBS + BCRP.
    # -------------------------------------------------------------------------
    base = integrar(
        sbs,
        bcrp
    )

    # -------------------------------------------------------------------------
    # PASO 7. Deflactar y calcular logaritmos.
    # -------------------------------------------------------------------------
    base = transformar(
        base
    )

    # -------------------------------------------------------------------------
    # PASO 8. Seleccionar y ordenar columnas.
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
    # PASO 9. Crear id_observacion.
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
    # PASO 10. Diagnósticos.
    # -------------------------------------------------------------------------
    diagnosticar(
        base,
        largo,
        guardar_detalle=
        guardar_auditoria
    )

    # -------------------------------------------------------------------------
    # PASO 11. Controles finales.
    # -------------------------------------------------------------------------
    controles_finales(
        base
    )

    return base


def pruebas_previas():
    """
    Ejecuta los 72 meses completamente en memoria.

    No guarda:
    - datos procesados;
    - auditorías;
    - log definitivo.
    """

    reportar(
        "=" * 78
    )

    reportar(
        "PRUEBAS PREVIAS DEL SCRIPT 03 "
        "— SIN GUARDAR BASE DEFINITIVA"
    )

    reportar(
        "=" * 78
    )

    base = preparar_base(
        guardar_auditoria=False
    )

    reportar(
        "\nPRUEBAS PREVIAS SUPERADAS."
    )

    reportar(
        "Resultado en memoria: "
        f"{len(base):,} filas × "
        f"{base.shape[1]} columnas."
    )

    return base


def main():
    """
    Ejecución definitiva.
    """

    reportar(
        "=" * 78
    )

    reportar(
        "SCRIPT 03 — LIMPIEZA E INTEGRACIÓN "
        "| empresa × mes "
        f"| {FECHA_INICIO} a {FECHA_CORTE}"
    )

    reportar(
        "El balance del panel se determina "
        "a partir de la presencia efectiva "
        "de empresas."
    )

    reportar(
        "=" * 78
    )

    base = preparar_base(
        guardar_auditoria=True
    )

    guardar(
        base
    )

    print(
        "\nScript 03 finalizado correctamente."
    )


if __name__ == "__main__":

    try:

        if (
            "--pruebas"
            in sys.argv
        ):

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
            "El script 03 se detuvo: "
            f"{mensaje}"
        )

        if (
            "--pruebas"
            not in sys.argv
        ):

            REPORTE.append(
                f"\n[ERROR] {mensaje}"
            )

            ARCHIVO_CONTROL.write_text(
                "\n".join(
                    REPORTE
                ) + "\n",
                encoding="utf-8"
            )

            escribir_log(
                "ERROR SCRIPT 03 | "
                f"{mensaje}"
            )

        raise
