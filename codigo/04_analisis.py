# Nombres y apellidos: Vania Adriana Egas Araujo
# Código de matrícula: 2024200498D
# Tema: N.º 15 del temario — Determinantes de las primas cedidas, el ingreso formal y la tasa de referencia en las primas netas del mercado asegurador peruano (2020-2025)
# Fecha de extracción: 2026-09-25
"""
SCRIPT 04 — ANÁLISIS: TABLAS, FIGURAS Y ESTIMACIONES DEL ARTÍCULO

Entrada única (solo lectura, no se modifica):
    datos_procesados/datos_procesados_2024200498D.csv

Modelo:
    Y  = log_primas_reales
    X1 = log_primas_cedidas
    X2 = log_ingreso_real
    X3 = tasa_referencia
    Muestra: muestra_modelo == 1 (definida en el script 03).
    (1) MCO agrupado                                -> referencia
    (2) Efectos fijos de empresa                   -> especificación base
    (3) Efectos fijos de empresa + mes del año     -> especificación PRINCIPAL
    Errores estándar agrupados por empresa (id_empresa), con inferencia t.
    Para el modelo (3) se reporta además Driscoll-Kraay como robustez.
    No se usan efectos fijos completos de periodo: absorberían X2 y X3,
    que solo varían en el tiempo.

Salidas (carpeta salidas/):
    tabla_1_descriptivos.csv / .tex
    tabla_2_empresas.csv / .tex
    tabla_3_correlaciones.csv / .tex
    tabla_4_regresion.csv / .tex
    tabla_5_robustez.csv / .tex
    estimaciones_modelos.txt
    figura_1_primas_mercado.png
    figura_2_variables_bcrp.png
    figura_3_estacionalidad.png
    figura_4_primas_promedio_empresa.png
    figura_5_y_x1_primas_cedidas.png
    figura_6_y_x2_ingreso.png
    figura_7_y_x3_tasa_referencia.png

Cómo se ejecuta (desde la carpeta raíz del proyecto):
    python codigo/04_analisis.py
"""


# =============================================================================
# BLOQUE 1. LIBRERÍAS
# =============================================================================
import hashlib
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf


# =============================================================================
# BLOQUE 2. PARÁMETROS Y RUTAS
# =============================================================================
FECHA_INICIO = "2020-01"
FECHA_CORTE = "2025-12"
MESES_ESPERADOS = 72
OBS_ESPERADAS = 1277
COLUMNAS_ESPERADAS = 18
EMPRESAS_ESPERADAS = 20
OBS_MODELO_ESPERADAS = 1254
MATRICULA = "2024200498D"

# Variables del modelo (definitivas, no se modifican)
Y = "log_primas_reales"
X1 = "log_primas_cedidas"
X2 = "log_ingreso_real"
X3 = "tasa_referencia"
REGRESORAS = [X1, X2, X3]

# Columnas exactas que debe traer la base procesada final del script 03
COLUMNAS_REQUERIDAS = [
    "id_observacion", "periodo", "id_empresa", "empresa",
    "primas_netas_acum_sbs", "primas_netas_mensual_nominal", "primas_netas_mensual_real", "log_primas_reales",
    "primas_cedidas_acum_sbs", "primas_cedidas_mensual_nominal", "primas_cedidas_mensual_real", "log_primas_cedidas",
    "ingreso_formal_nominal", "ingreso_real", "log_ingreso_real",
    "tasa_referencia", "ipc", "muestra_modelo",
]

# Variables que deben estar completas en niveles
VARIABLES_NIVEL_COMPLETAS = [
    "primas_netas_mensual_real",
    "primas_cedidas_mensual_real",
    "ingreso_real",
    "tasa_referencia",
    "ipc",
]

# Variables BCRP: para un mismo mes deben ser idénticas entre empresas
VARIABLES_BCRP_MENSUALES = [
    "ingreso_formal_nominal",
    "ingreso_real",
    "log_ingreso_real",
    "tasa_referencia",
    "ipc",
]

# Etiquetas para tablas y figuras (con unidad)
ETIQUETAS = {
    "primas_netas_mensual_real": "Primas netas mensuales reales (miles S/ dic-2021)",
    "primas_cedidas_mensual_real": "Primas cedidas mensuales reales (miles S/ dic-2021)",
    "ingreso_real": "Ingreso formal real (S/ dic-2021)",
    "tasa_referencia": "Tasa de referencia (%)",
    "log_primas_reales": "ln primas netas reales (Y)",
    "log_primas_cedidas": "ln primas cedidas reales (X1)",
    "log_ingreso_real": "ln ingreso formal real (X2)",
}
NOMBRES_REGRESORAS = {
    X1: "ln primas cedidas reales (X1)",
    X2: "ln ingreso formal real (X2)",
    X3: "Tasa de referencia (X3)",
}

# Driscoll-Kraay: rezagos = piso(4 · (T/100)^(2/9)) con T = 72 meses -> 3
REZAGOS_DK = int(np.floor(4 * (MESES_ESPERADOS / 100) ** (2 / 9)))

# Colores de las figuras
AZUL, NARANJA = "#2a78d6", "#eb6834"
TINTA, TINTA_SECUNDARIA, GRILLA = "#0b0b0b", "#52514e", "#e4e3df"
FUENTE_SBS = "Fuente: SBS, cuadro S-401. Elaboración propia."
FUENTE_BCRP = "Fuente: BCRP, BCRPData (PN31883GM, PD04722MM, PN38705PM). Elaboración propia."
FUENTE_AMBAS = "Fuente: SBS (S-401) y BCRP (BCRPData). Elaboración propia."

# Rutas relativas: el script está en /codigo y sube un nivel a la raíz del proyecto
RAIZ = Path(__file__).resolve().parent.parent
ARCHIVO_BASE = RAIZ / "datos_procesados" / f"datos_procesados_{MATRICULA}.csv"
CARPETA_SALIDAS = RAIZ / "salidas"
ARCHIVO_LOG = RAIZ / "log_ejecucion.txt"
CSV_ENCODING = "utf-8-sig"

CARPETA_SALIDAS.mkdir(parents=True, exist_ok=True)
SALIDAS_GENERADAS = []


# =============================================================================
# BLOQUE 3. FUNCIONES AUXILIARES
# =============================================================================
def escribir_log(texto):
    """Añade una línea con fecha y hora a log_ejecucion.txt."""
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ARCHIVO_LOG, "a", encoding="utf-8") as archivo:
        archivo.write(f"{ahora} | {texto}\n")


def huella(ruta):
    """SHA-256 de un archivo (sirve para comprobar que la base no se modificó)."""
    return hashlib.sha256(Path(ruta).read_bytes()).hexdigest()


def escapar_latex_texto(texto):
    """Escapa caracteres especiales de LaTeX en las notas de tablas."""
    mapa = {
        "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
        "_": r"\_", "{": r"\{", "}": r"\}",
    }
    return "".join(mapa.get(caracter, caracter) for caracter in str(texto))


def guardar_tabla(tabla, nombre, titulo, nota):
    """Guarda una tabla en CSV (datos) y en LaTeX (para Overleaf), con título y nota."""
    ruta_csv = CARPETA_SALIDAS / f"{nombre}.csv"
    tabla.to_csv(ruta_csv, index=False, encoding=CSV_ENCODING)

    latex = tabla.to_latex(
        index=False,
        escape=True,
        na_rep="",
        float_format=lambda v: f"{v:.2f}",
        caption=titulo,
        label=f"tab:{nombre}",
        position="htbp",
    )
    nota_tex = escapar_latex_texto(nota)
    latex = latex.replace(
        "\\end{table}",
        f"\\par\\vspace{{2pt}}{{\\footnotesize \\textit{{Nota.}} {nota_tex}}}\n\\end{{table}}",
    )
    ruta_tex = CARPETA_SALIDAS / f"{nombre}.tex"
    ruta_tex.write_text(latex, encoding="utf-8")

    SALIDAS_GENERADAS.extend([ruta_csv.name, ruta_tex.name])
    print(f"  Tabla guardada: {ruta_csv.name} y {ruta_tex.name}")


def estrellas(p):
    """*** p < 0.01, ** p < 0.05, * p < 0.10."""
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.10 else ""


def a_fecha(periodo):
    """'2020-01' -> fecha del primer día del mes (para el eje de las figuras)."""
    return pd.to_datetime(periodo + "-01")


def limitar_ventana(ax, fechas):
    """El eje de tiempo termina en la ventana del estudio (sin marca de 2026)."""
    ax.set_xlim(fechas.min() - pd.Timedelta(days=20), fechas.max() + pd.Timedelta(days=20))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))


def estilo_ejes(ax):
    """Estilo sobrio: sin marco superior/derecho, grilla tenue y texto en tinta neutra."""
    for lado in ["top", "right"]:
        ax.spines[lado].set_visible(False)
    for lado in ["left", "bottom"]:
        ax.spines[lado].set_color(TINTA_SECUNDARIA)
    ax.grid(axis="y", color=GRILLA, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(colors=TINTA_SECUNDARIA, labelsize=9)


def guardar_figura(fig, nombre, fuente):
    """Agrega la nota de fuente y guarda la figura en PNG a 300 dpi."""
    fig.text(0.01, -0.01, fuente, fontsize=8, color=TINTA_SECUNDARIA, ha="left", va="top")
    ruta = CARPETA_SALIDAS / f"{nombre}.png"
    fig.savefig(ruta, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    SALIDAS_GENERADAS.append(ruta.name)
    print(f"  Figura guardada: {ruta.name}")


# =============================================================================
# BLOQUE 4. CARGA Y VALIDACIÓN DE LA BASE PROCESADA (solo lectura)
# =============================================================================
def cargar_base():
    if not ARCHIVO_BASE.exists():
        raise FileNotFoundError("No existe la base procesada. Ejecute primero 03_limpieza_datos.py")

    base = pd.read_csv(ARCHIVO_BASE, encoding=CSV_ENCODING, dtype={"periodo": str})

    # Si el archivo se volvió a guardar desde Excel con ';', no se usa: cambiaría el SHA-256
    if base.shape[1] == 1:
        raise RuntimeError(
            "La base parece guardada con ';' (probablemente desde Excel). "
            "Vuelva a generarla con 03_limpieza_datos.py; no la edite en Excel."
        )

    # La entrada debe ser exactamente la base final cerrada del Script 03
    if base.shape != (OBS_ESPERADAS, COLUMNAS_ESPERADAS):
        raise RuntimeError(
            f"Dimensiones inesperadas: {base.shape}. "
            f"Se esperaba ({OBS_ESPERADAS}, {COLUMNAS_ESPERADAS})."
        )

    if list(base.columns) != COLUMNAS_REQUERIDAS:
        faltan = [c for c in COLUMNAS_REQUERIDAS if c not in base.columns]
        extras = [c for c in base.columns if c not in COLUMNAS_REQUERIDAS]
        raise RuntimeError(
            "La estructura de columnas no coincide con la base final del Script 03. "
            f"Faltan={faltan}; extras={extras}."
        )

    if base["periodo"].min() != FECHA_INICIO or base["periodo"].max() != FECHA_CORTE:
        raise RuntimeError(
            f"Ventana inesperada: {base['periodo'].min()} a {base['periodo'].max()}."
        )
    if base["periodo"].nunique() != MESES_ESPERADOS:
        raise RuntimeError(f"La base no tiene los {MESES_ESPERADOS} meses esperados.")

    if base["id_empresa"].nunique() != EMPRESAS_ESPERADAS:
        raise RuntimeError(
            f"Número inesperado de empresas: {base['id_empresa'].nunique()}. "
            f"Se esperaban {EMPRESAS_ESPERADAS}."
        )

    if base["id_observacion"].isna().any() or not base["id_observacion"].is_unique:
        raise RuntimeError("id_observacion debe existir sin faltantes y ser único.")

    if base.duplicated(["id_empresa", "periodo"]).any():
        raise RuntimeError("Hay duplicados id_empresa × periodo.")

    # muestra_modelo puede venir como True/False o 1/0: se interpreta, no se recalcula
    valores = set(base["muestra_modelo"].astype(str).str.strip().str.upper().unique())
    if not valores <= {"TRUE", "FALSE", "1", "0", "1.0", "0.0"}:
        raise RuntimeError(f"muestra_modelo tiene valores no válidos: {valores}")
    base["muestra_modelo"] = (
        base["muestra_modelo"].astype(str).str.strip().str.upper().isin(["TRUE", "1", "1.0"])
    )

    n_modelo = int(base["muestra_modelo"].sum())
    if n_modelo != OBS_MODELO_ESPERADAS:
        raise RuntimeError(
            f"muestra_modelo tiene {n_modelo} observaciones; "
            f"se esperaban {OBS_MODELO_ESPERADAS}."
        )

    # No deben existir huecos reales en las variables de nivel
    faltantes_nivel = base[VARIABLES_NIVEL_COMPLETAS].isna().sum()
    faltantes_nivel = faltantes_nivel[faltantes_nivel > 0]
    if not faltantes_nivel.empty:
        raise RuntimeError(
            "Se detectaron faltantes inesperados en variables de nivel: "
            + faltantes_nivel.to_dict().__str__()
        )

    # BCRP es mensual: dentro de cada periodo debe existir un único valor para todas las empresas
    for columna in VARIABLES_BCRP_MENSUALES:
        max_distintos = int(base.groupby("periodo")[columna].nunique(dropna=False).max())
        if max_distintos != 1:
            raise RuntimeError(
                f"{columna} presenta más de un valor dentro del mismo mes; revise la integración BCRP."
            )

    # Las filas marcadas para el modelo deben tener completas Y y las tres X
    if base.loc[base["muestra_modelo"], [Y] + REGRESORAS].isna().any().any():
        raise RuntimeError(
            "muestra_modelo = 1 contiene faltantes en Y o en alguna regresora."
        )

    print(
        f"Base procesada validada: {len(base):,} filas × {base.shape[1]} columnas | "
        f"{base['id_empresa'].nunique()} empresas | "
        f"{base['periodo'].min()} a {base['periodo'].max()} | "
        f"muestra_modelo = 1: {n_modelo:,} | fuera: {len(base) - n_modelo:,}"
    )
    print("  ✓ id_observacion único")
    print("  ✓ 0 duplicados id_empresa × periodo")
    print("  ✓ variables de nivel sin faltantes")
    print("  ✓ variables BCRP constantes dentro de cada mes")
    return base


# =============================================================================
# BLOQUE 5. TABLA 1 — ESTADÍSTICOS DESCRIPTIVOS
# =============================================================================
def tabla_descriptivos(base):
    filas = []
    for columna, etiqueta in ETIQUETAS.items():
        serie = base[columna].dropna()
        filas.append({
            "Variable": etiqueta,
            "N": len(serie),
            "Media": round(serie.mean(), 3),
            "DE": round(serie.std(), 3),
            "Mediana": round(serie.median(), 3),
            "Mín.": round(serie.min(), 3),
            "Máx.": round(serie.max(), 3),
        })
    tabla = pd.DataFrame(filas)

    empresas_mes = base.groupby("periodo")["id_empresa"].nunique()
    nota = (f"Panel empresa-mes, {base['periodo'].min()} a {base['periodo'].max()} "
            f"({base['periodo'].nunique()} meses): {len(base):,} observaciones de "
            f"{base['id_empresa'].nunique()} empresas (entre {empresas_mes.min()} y {empresas_mes.max()} por mes). "
            f"Muestra del modelo: {int(base['muestra_modelo'].sum()):,} observaciones. "
            "Los logaritmos no existen cuando el valor es cero o negativo, por eso su N es menor. "
            + FUENTE_AMBAS)
    guardar_tabla(tabla, "tabla_1_descriptivos", "Estadísticos descriptivos de las variables del estudio", nota)
    return nota


# =============================================================================
# BLOQUE 6. TABLA 2 — ESTRUCTURA DEL MERCADO POR EMPRESA
# =============================================================================
def tabla_empresas(base):
    g = base.groupby("id_empresa")
    tabla = pd.DataFrame({
        "Empresa": g["empresa"].first(),
        "Desde": g["periodo"].min(),
        "Hasta": g["periodo"].max(),
        "Meses": g["periodo"].nunique(),
        "Obs. modelo": g["muestra_modelo"].sum().astype(int),
        "Primas netas (miles S/ reales)": g["primas_netas_mensual_real"].sum(),
        "Primas cedidas (miles S/ reales)": g["primas_cedidas_mensual_real"].sum(),
    })

    total_netas = tabla["Primas netas (miles S/ reales)"].sum()
    total_cedidas = tabla["Primas cedidas (miles S/ reales)"].sum()

    tabla["Participación (%)"] = (
        100 * tabla["Primas netas (miles S/ reales)"] / total_netas
    )

    # La cesión solo se calcula si las primas netas acumuladas
    # de la empresa durante el periodo completo son positivas.
    netas_positivas = tabla["Primas netas (miles S/ reales)"].where(
        tabla["Primas netas (miles S/ reales)"] > 0
    )
    tabla["Cesión (%)"] = (
        100 * tabla["Primas cedidas (miles S/ reales)"] / netas_positivas
    )

    tabla = tabla.sort_values(
        "Participación (%)", ascending=False
    ).reset_index(drop=True)

    total = {
        "Empresa": "Total mercado",
        "Desde": base["periodo"].min(),
        "Hasta": base["periodo"].max(),
        "Meses": base["periodo"].nunique(),
        "Obs. modelo": int(base["muestra_modelo"].sum()),
        "Primas netas (miles S/ reales)": total_netas,
        "Primas cedidas (miles S/ reales)": total_cedidas,
        "Participación (%)": 100.0,
        "Cesión (%)": 100 * total_cedidas / total_netas,
    }
    tabla = pd.concat([tabla, pd.DataFrame([total])], ignore_index=True)

    # Se mantienen dos decimales porque la base conserva la unidad
    # original en miles de soles con decimales.
    for col in [
        "Primas netas (miles S/ reales)",
        "Primas cedidas (miles S/ reales)",
    ]:
        tabla[col] = tabla[col].round(2) + 0.0

    for col in ["Participación (%)", "Cesión (%)"]:
        tabla[col] = tabla[col].round(2) + 0.0

    nota = (
        "Primas mensuales reales sumadas de 2020-01 a 2025-12, "
        "expresadas en miles de soles reales de diciembre de 2021. "
        "No se realiza conversión a millones. "
        "Participación = primas netas de la empresa / primas netas del mercado. "
        "Cesión = primas cedidas / primas netas y se presenta únicamente cuando "
        "las primas netas acumuladas de la empresa en el periodo son positivas; "
        "es un indicador descriptivo y no entra al modelo. "
        "Panel no balanceado: algunas empresas no están presentes los 72 meses. "
        + FUENTE_AMBAS
    )

    guardar_tabla(
        tabla,
        "tabla_2_empresas",
        "Estructura del mercado asegurador por empresa, 2020-2025",
        nota,
    )


# =============================================================================
# BLOQUE 7. FIGURAS 1 A 3 — EVOLUCIÓN Y ESTACIONALIDAD
# =============================================================================
def series_mercado(base):
    """Suma mensual del mercado y variables BCRP correspondientes al mes."""
    mercado = base.groupby("periodo").agg(
        netas=("primas_netas_mensual_real", "sum"),
        cedidas=("primas_cedidas_mensual_real", "sum"),
        ingreso=("ingreso_real", "first"),
        tasa=("tasa_referencia", "first"),
    ).reset_index()
    mercado["fecha"] = mercado["periodo"].apply(a_fecha)
    return mercado


def figura_primas(mercado):
    """Primas reales del mercado en miles de soles, sin cambiar la unidad de la base."""
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8.4, 6.0), sharex=True)

    ax1.plot(mercado["fecha"], mercado["netas"], color=AZUL, linewidth=1.8)
    ax1.set_title("Primas netas reales del mercado asegurador", loc="left", fontsize=10.5, color=TINTA)
    ax1.set_ylabel("Miles de S/ reales\n(IPC base dic. 2021 = 100)", color=TINTA_SECUNDARIA, fontsize=9)

    ax2.plot(mercado["fecha"], mercado["cedidas"], color=NARANJA, linewidth=1.8)
    ax2.set_title("Primas cedidas reales del mercado asegurador", loc="left", fontsize=10.5, color=TINTA)
    ax2.set_ylabel("Miles de S/ reales\n(IPC base dic. 2021 = 100)", color=TINTA_SECUNDARIA, fontsize=9)

    limitar_ventana(ax2, mercado["fecha"])
    for ax in (ax1, ax2):
        estilo_ejes(ax)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{x:,.0f}"))

    fig.suptitle(
        "Evolución mensual de las primas del mercado asegurador, 2020-2025",
        x=0.125, ha="left", fontsize=11.5, color=TINTA
    )
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    guardar_figura(
        fig,
        "figura_1_primas_mercado",
        FUENTE_AMBAS
        + " Suma mensual de las empresas presentes. "
          "Unidad: miles de soles reales; IPC base dic. 2021 = 100."
    )


def figura_bcrp(mercado):
    """Ingreso formal real y tasa de referencia, en paneles y unidades separadas."""
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8.4, 5.8), sharex=True)

    ax1.plot(mercado["fecha"], mercado["ingreso"], color=AZUL, linewidth=1.8)
    ax1.set_title("Ingreso formal real", loc="left", fontsize=10.5, color=TINTA)
    ax1.set_ylabel("S/ reales\n(IPC base dic. 2021 = 100)", color=TINTA_SECUNDARIA, fontsize=9)

    ax2.step(mercado["fecha"], mercado["tasa"], where="post", color=NARANJA, linewidth=1.8)
    ax2.set_title("Tasa de referencia del BCRP", loc="left", fontsize=10.5, color=TINTA)
    ax2.set_ylabel("Porcentaje (%)", color=TINTA_SECUNDARIA, fontsize=9)

    limitar_ventana(ax2, mercado["fecha"])
    for ax in (ax1, ax2):
        estilo_ejes(ax)

    fig.suptitle(
        "Variables macrofinancieras mensuales, 2020-2025",
        x=0.125, ha="left", fontsize=11.5, color=TINTA
    )
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    guardar_figura(fig, "figura_2_variables_bcrp", FUENTE_BCRP)


def figura_estacionalidad(mercado):
    """
    Promedio observado por mes del año.
    No se crea ningún índice ni normalización adicional.
    """
    datos = mercado.copy()
    datos["mes"] = datos["periodo"].str[5:7].astype(int)
    promedio = datos.groupby("mes").agg(
        netas=("netas", "mean"),
        cedidas=("cedidas", "mean"),
    )

    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
             "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8.4, 6.0), sharex=True)

    ax1.plot(range(1, 13), promedio["netas"].values, color=AZUL,
             linewidth=1.8, marker="o", markersize=4)
    ax1.set_title("Primas netas reales: promedio por mes del año", loc="left", fontsize=10.5, color=TINTA)
    ax1.set_ylabel("Miles de S/ reales\n(IPC base dic. 2021 = 100)", color=TINTA_SECUNDARIA, fontsize=9)

    ax2.plot(range(1, 13), promedio["cedidas"].values, color=NARANJA,
             linewidth=1.8, marker="o", markersize=4)
    ax2.set_title("Primas cedidas reales: promedio por mes del año", loc="left", fontsize=10.5, color=TINTA)
    ax2.set_ylabel("Miles de S/ reales\n(IPC base dic. 2021 = 100)", color=TINTA_SECUNDARIA, fontsize=9)
    ax2.set_xticks(range(1, 13))
    ax2.set_xticklabels(meses)

    for ax in (ax1, ax2):
        estilo_ejes(ax)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{x:,.0f}"))

    fig.suptitle(
        "Estacionalidad mensual de las primas, 2020-2025",
        x=0.125, ha="left", fontsize=11.5, color=TINTA
    )
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    guardar_figura(
        fig,
        "figura_3_estacionalidad",
        FUENTE_AMBAS
        + " Promedio de los valores mensuales observados durante 2020-2025. "
          "Unidad: miles de soles reales; IPC base dic. 2021 = 100. "
          "No se utiliza ningún índice."
    )
    return promedio


# =============================================================================
# BLOQUE 8. ESTIMACIÓN DE LOS MODELOS
# =============================================================================
FORMULAS = {
    "(1) MCO agrupado": f"{Y} ~ {X1} + {X2} + {X3}",
    "(2) EF empresa": f"{Y} ~ {X1} + {X2} + {X3} + C(id_empresa)",
    "(3) EF empresa + mes del año": (
        f"{Y} ~ {X1} + {X2} + {X3} + "
        "C(id_empresa) + C(mes, Treatment(reference='12'))"
    ),
}


def r2_dentro(modelo, datos):
    """R² dentro de empresas: variación temporal de Y explicada por el modelo."""
    y_dentro = datos[Y] - datos.groupby("id_empresa")[Y].transform("mean")
    denominador = (y_dentro ** 2).sum()
    if denominador <= 0:
        return np.nan
    return 1 - (modelo.resid ** 2).sum() / denominador


def estimar(base):
    muestra = base.loc[base["muestra_modelo"]].copy()
    muestra = muestra.sort_values(["periodo", "id_empresa"]).reset_index(drop=True)
    muestra["mes"] = muestra["periodo"].str[5:7]

    if len(muestra) != OBS_MODELO_ESPERADAS:
        raise RuntimeError(
            f"La muestra de estimación tiene {len(muestra)} observaciones; "
            f"se esperaban {OBS_MODELO_ESPERADAS}."
        )

    grupos = pd.factorize(muestra["id_empresa"])[0]

    resultados = {}
    for nombre, formula in FORMULAS.items():
        modelo = smf.ols(formula, data=muestra).fit(
            cov_type="cluster",
            cov_kwds={"groups": grupos, "use_correction": True},
            use_t=True,
        )
        resultados[nombre] = modelo
        print(
            f"  Estimado: {nombre} | N = {int(modelo.nobs):,} | "
            f"empresas = {muestra['id_empresa'].nunique()}"
        )

    # Robustez del modelo principal: Driscoll-Kraay
    # El índice de tiempo es común a todas las empresas dentro de cada mes.
    tiempo = pd.factorize(muestra["periodo"], sort=True)[0]
    principal = FORMULAS["(3) EF empresa + mes del año"]
    modelo_dk = smf.ols(principal, data=muestra).fit(
        cov_type="hac-groupsum",
        cov_kwds={"time": tiempo, "maxlags": REZAGOS_DK},
    )
    return muestra, resultados, modelo_dk


# =============================================================================
# BLOQUE 9. TABLAS 3, 4 Y 5 + ARCHIVO DE ESTIMACIONES
# =============================================================================
def tabla_correlaciones(muestra):
    columnas = [Y] + REGRESORAS
    etiquetas = {
        Y: "ln primas netas reales (Y)",
        X1: "ln primas cedidas reales (X1)",
        X2: "ln ingreso formal real (X2)",
        X3: "Tasa de referencia (X3)",
    }
    correlaciones = muestra[columnas].corr().rename(index=etiquetas, columns=etiquetas).round(3)
    tabla = correlaciones.reset_index().rename(columns={"index": "Variable"})

    nota = (
        f"Correlaciones de Pearson calculadas sobre muestra_modelo = 1 "
        f"(N = {len(muestra):,}). Son asociaciones descriptivas y no implican causalidad. "
        + FUENTE_AMBAS
    )
    guardar_tabla(
        tabla,
        "tabla_3_correlaciones",
        "Matriz de correlaciones de las variables del modelo",
        nota,
    )
    return tabla


def tabla_regresion(muestra, resultados):
    nombres = list(resultados)
    filas = []

    for x in REGRESORAS:
        fila_coef = {"": NOMBRES_REGRESORAS[x]}
        fila_ee = {"": ""}
        for nombre, modelo in resultados.items():
            fila_coef[nombre] = f"{modelo.params[x]:.4f}{estrellas(modelo.pvalues[x])}"
            fila_ee[nombre] = f"({modelo.bse[x]:.4f})"
        filas += [fila_coef, fila_ee]

    filas.append({
        "": "Efectos fijos de empresa",
        **{n: ("Sí" if "EF" in n else "No") for n in nombres},
    })
    filas.append({
        "": "Indicadores de mes del año",
        **{n: ("Sí" if "mes" in n else "No") for n in nombres},
    })
    filas.append({
        "": "Observaciones",
        **{n: f"{int(m.nobs):,}" for n, m in resultados.items()},
    })
    filas.append({
        "": "Empresas",
        **{n: str(muestra["id_empresa"].nunique()) for n in nombres},
    })
    filas.append({
        "": "R²",
        **{
            n: (
                f"{r2_dentro(m, muestra):.3f} (dentro)"
                if "EF" in n
                else f"{m.rsquared:.3f}"
            )
            for n, m in resultados.items()
        },
    })
    tabla = pd.DataFrame(filas)

    nota = (
        "Variable dependiente: ln primas netas reales (Y). Muestra: muestra modelo = 1. "
        "Errores estándar agrupados por empresa entre paréntesis, con inferencia t. "
        "El modelo (3) incluye 11 indicadores de mes del año (diciembre como base); "
        "sus coeficientes completos se conservan en estimaciones_modelos.txt. "
        "No se incluyen efectos fijos completos de periodo porque absorberían X2 y X3. "
        "Los coeficientes se interpretan como asociaciones condicionadas por el modelo, no como efectos causales. "
        "*** p < 0.01, ** p < 0.05, * p < 0.10. "
        + FUENTE_AMBAS
    )
    guardar_tabla(
        tabla,
        "tabla_4_regresion",
        "Determinantes de las primas netas reales: estimaciones de panel",
        nota,
    )
    return tabla


def tabla_robustez(muestra, resultados, modelo_dk):
    principal = resultados["(3) EF empresa + mes del año"]
    filas = []

    for x in REGRESORAS:
        filas.append({
            "Variable": NOMBRES_REGRESORAS[x],
            "Coeficiente": round(float(principal.params[x]), 4),
            "EE clúster empresa": round(float(principal.bse[x]), 4),
            "p clúster empresa": round(float(principal.pvalues[x]), 4),
            "EE Driscoll-Kraay": round(float(modelo_dk.bse[x]), 4),
            "p Driscoll-Kraay": round(float(modelo_dk.pvalues[x]), 4),
        })

    tabla = pd.DataFrame(filas)
    nota = (
        "Misma especificación del modelo principal (3). Las dos columnas de inferencia utilizan los mismos "
        "coeficientes OLS; cambia únicamente la matriz de covarianza. El clúster por empresa es la inferencia "
        f"principal y Driscoll-Kraay con {REZAGOS_DK} rezagos se presenta como robustez. "
        "No se interpreta como prueba causal. "
        + FUENTE_AMBAS
    )
    guardar_tabla(
        tabla,
        "tabla_5_robustez",
        "Robustez de la inferencia del modelo principal",
        nota,
    )
    return tabla


def archivo_estimaciones(muestra, resultados, modelo_dk):
    correlaciones = muestra[[Y] + REGRESORAS].corr().round(3)
    partes = [
        "ESTIMACIONES DEL SCRIPT 04 — salida completa de los modelos",
        (
            f"Muestra: muestra_modelo = 1 | observaciones = {len(muestra):,} | "
            f"empresas = {muestra['id_empresa'].nunique()} | meses = {muestra['periodo'].nunique()}"
        ),
        "",
        "Matriz de correlaciones (muestra del modelo):",
        correlaciones.to_string(),
        "",
    ]

    for nombre, modelo in resultados.items():
        partes += [
            "=" * 90,
            f"{nombre} — errores estándar agrupados por empresa, inferencia t",
            "=" * 90,
            modelo.summary().as_text(),
            "",
        ]

    partes += [
        "=" * 90,
        f"(3) EF empresa + mes del año — errores de Driscoll-Kraay ({REZAGOS_DK} rezagos)",
        "=" * 90,
        modelo_dk.summary().as_text(),
    ]

    ruta = CARPETA_SALIDAS / "estimaciones_modelos.txt"
    ruta.write_text("\n".join(partes) + "\n", encoding="utf-8")
    SALIDAS_GENERADAS.append(ruta.name)
    print(f"  Estimaciones guardadas: {ruta.name}")


# =============================================================================
# BLOQUE 10. FIGURA 4 — PRIMAS PROMEDIO POR EMPRESA
# =============================================================================
def figura_empresas(base):
    """Prima neta real promedio mensual por trayectoria empresarial."""
    resumen = (
        base.groupby("id_empresa")
        .agg(
            empresa=("empresa", "first"),
            prima_neta_promedio=("primas_netas_mensual_real", "mean"),
            meses=("periodo", "nunique"),
        )
        .reset_index()
        .sort_values("prima_neta_promedio", ascending=True)
    )

    fig, ax = plt.subplots(figsize=(8.4, 7.2))
    ax.barh(resumen["empresa"], resumen["prima_neta_promedio"], color=AZUL, alpha=0.85)

    ax.set_title(
        "Primas netas reales promedio mensual por empresa, 2020-2025",
        loc="left", fontsize=11, color=TINTA
    )
    ax.set_xlabel(
        "Miles de S/ reales (IPC base dic. 2021 = 100)",
        color=TINTA_SECUNDARIA, fontsize=9
    )
    ax.set_ylabel("")
    estilo_ejes(ax)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{x:,.0f}"))
    ax.tick_params(axis="y", labelsize=8)
    fig.tight_layout()

    guardar_figura(
        fig,
        "figura_4_primas_promedio_empresa",
        FUENTE_AMBAS
        + " Promedio mensual calculado sobre los meses observados de cada empresa. "
          "Unidad: miles de soles reales; IPC base dic. 2021 = 100."
    )


# =============================================================================
# BLOQUE 10B. FIGURAS 5 A 7 — RELACIONES PARCIALES DEL MODELO PRINCIPAL
# =============================================================================

def _muestra_para_relaciones(base):
    """
    Prepara exactamente la muestra utilizada por el modelo principal.
    No modifica la base procesada.
    """
    datos = base.loc[
        base["muestra_modelo"].astype(bool)
    ].copy()

    datos = datos.sort_values(
        ["periodo", "id_empresa"]
    ).reset_index(drop=True)

    # Mes del año utilizado en la especificación principal
    datos["mes"] = datos["periodo"].str[5:7]

    return datos


def _relacion_parcial(base, x_objetivo):
    """
    Calcula la relación parcial entre Y y una regresora X mediante
    el principio de Frisch-Waugh-Lovell.

    Para cada variable:
    1. Se retira de Y la parte explicada por las otras regresoras,
       los efectos fijos de empresa y los indicadores de mes.
    2. Se retira de X la parte explicada por esos mismos controles.
    3. Se relacionan ambos residuos.

    La pendiente obtenida coincide con el coeficiente de esa X
    en el modelo principal.
    """

    datos = _muestra_para_relaciones(base)

    regresoras = [X1, X2, X3]

    otras_x = [
        variable
        for variable in regresoras
        if variable != x_objetivo
    ]

    controles = (
        otras_x
        + [
            "C(id_empresa)",
            "C(mes, Treatment(reference='12'))"
        ]
    )

    controles_formula = " + ".join(controles)

    # -------------------------------------------------------------------------
    # Residualizar Y respecto de todos los controles excepto X objetivo
    # -------------------------------------------------------------------------
    formula_y = f"{Y} ~ {controles_formula}"

    modelo_y = smf.ols(
        formula_y,
        data=datos
    ).fit()

    residuo_y = modelo_y.resid

    # -------------------------------------------------------------------------
    # Residualizar X objetivo respecto de los mismos controles
    # -------------------------------------------------------------------------
    formula_x = f"{x_objetivo} ~ {controles_formula}"

    modelo_x = smf.ols(
        formula_x,
        data=datos
    ).fit()

    residuo_x = modelo_x.resid

    # -------------------------------------------------------------------------
    # Coeficiente del modelo principal completo
    # -------------------------------------------------------------------------
    formula_principal = FORMULAS[
        "(3) EF empresa + mes del año"
    ]

    modelo_principal = smf.ols(
        formula_principal,
        data=datos
    ).fit()

    beta = float(
        modelo_principal.params[x_objetivo]
    )

    resultado = pd.DataFrame({
        "x_residual": residuo_x,
        "y_residual": residuo_y,
    })

    return resultado, beta


def _dibujar_relacion_parcial(
    base,
    x_objetivo,
    titulo,
    etiqueta_x,
    nombre_archivo
):
    """
    Genera una figura de relación parcial coherente con
    la especificación econométrica principal.
    """

    datos, beta = _relacion_parcial(
        base,
        x_objetivo
    )

    fig, ax = plt.subplots(
        figsize=(7.2, 5.2)
    )

    # Observaciones ajustadas
    ax.scatter(
        datos["x_residual"],
        datos["y_residual"],
        s=18,
        alpha=0.30,
        color=AZUL,
        edgecolors="none",
    )

    # -------------------------------------------------------------------------
    # Línea cuya pendiente es exactamente el coeficiente del modelo principal
    # -------------------------------------------------------------------------
    x_min = datos["x_residual"].min()
    x_max = datos["x_residual"].max()

    x_linea = np.linspace(
        x_min,
        x_max,
        100
    )

    y_linea = beta * x_linea

    ax.plot(
        x_linea,
        y_linea,
        color=NARANJA,
        linewidth=2.0,
        label=f"Pendiente ajustada: β = {beta:.3f}",
    )

    ax.set_title(
        titulo,
        loc="left",
        fontsize=11,
        color=TINTA,
    )

    ax.set_xlabel(
        etiqueta_x,
        fontsize=9,
    )

    ax.set_ylabel(
        "Componente residual de log(primas netas reales)",
        fontsize=9,
    )

    estilo_ejes(ax)

    ax.legend(
        frameon=False,
        fontsize=8,
    )

    fig.tight_layout()

    guardar_figura(
        fig,
        nombre_archivo,
        FUENTE_AMBAS
        + " Relación parcial obtenida controlando las demás regresoras, "
          "efectos fijos de empresa e indicadores de mes del año. "
          "La pendiente corresponde al coeficiente del modelo principal. "
          "La figura muestra asociación condicionada, no causalidad."
    )


# =============================================================================
# FIGURA 5 — Y VS X1
# =============================================================================
def figura_relacion_y_x1(base):
    _dibujar_relacion_parcial(
        base=base,
        x_objetivo=X1,
        titulo=(
            "Relación parcial entre primas netas "
            "y primas cedidas reales"
        ),
        etiqueta_x=(
            "Componente residual de log(primas cedidas reales)"
        ),
        nombre_archivo=(
            "figura_5_y_x1_primas_cedidas"
        ),
    )


# =============================================================================
# FIGURA 6 — Y VS X2
# =============================================================================
def figura_relacion_y_x2(base):
    _dibujar_relacion_parcial(
        base=base,
        x_objetivo=X2,
        titulo=(
            "Relación parcial entre primas netas "
            "e ingreso formal real"
        ),
        etiqueta_x=(
            "Componente residual de log(ingreso formal real)"
        ),
        nombre_archivo=(
            "figura_6_y_x2_ingreso"
        ),
    )


# =============================================================================
# FIGURA 7 — Y VS X3
# =============================================================================
def figura_relacion_y_x3(base):
    _dibujar_relacion_parcial(
        base=base,
        x_objetivo=X3,
        titulo=(
            "Relación parcial entre primas netas "
            "y tasa de referencia"
        ),
        etiqueta_x=(
            "Componente residual de la tasa de referencia (%)"
        ),
        nombre_archivo=(
            "figura_7_y_x3_tasa_referencia"
        ),
    )


# =============================================================================
# BLOQUE 11. EJECUCIÓN PRINCIPAL
# =============================================================================
def main():
    print("=" * 78)
    print(f"SCRIPT 04 — ANÁLISIS | {FECHA_INICIO} a {FECHA_CORTE}")
    print("=" * 78)

    huella_inicial = huella(ARCHIVO_BASE) if ARCHIVO_BASE.exists() else None
    base = cargar_base()

    print("\nTablas descriptivas:")
    tabla_descriptivos(base)
    tabla_empresas(base)

    print("\nFiguras descriptivas:")
    mercado = series_mercado(base)
    figura_primas(mercado)                # Figura 1
    figura_bcrp(mercado)                  # Figura 2
    figura_estacionalidad(mercado)        # Figura 3
    figura_empresas(base)                  # Figura 4

    print("\nModelos:")
    muestra, resultados, modelo_dk = estimar(base)

    print("\nFiguras de relación Y-X:")
    figura_relacion_y_x1(base)             # Figura 5
    figura_relacion_y_x2(base)             # Figura 6
    figura_relacion_y_x3(base)             # Figura 7

    print("\nTablas analíticas:")
    tabla_correlaciones(muestra)                                # Tabla 3
    tabla_principal = tabla_regresion(muestra, resultados)      # Tabla 4
    tabla_robustez(muestra, resultados, modelo_dk)              # Tabla 5
    archivo_estimaciones(muestra, resultados, modelo_dk)

    huella_final = huella(ARCHIVO_BASE)
    if huella_final != huella_inicial:
        raise RuntimeError("La base procesada cambió durante el análisis.")

    print("\nTabla 4 (vista rápida):")
    print(tabla_principal.to_string(index=False))
    print(f"\nSalidas generadas en salidas/ ({len(SALIDAS_GENERADAS)} archivos):")
    for nombre in SALIDAS_GENERADAS:
        print(f"  - {nombre}")

    print(f"\nSHA-256 de la base analizada (sin cambios): {huella_inicial}")
    escribir_log(
        f"ANALISIS 04 | base={ARCHIVO_BASE.name} | sha256={huella_inicial} | "
        f"observaciones_modelo={int(muestra.shape[0])} | salidas={len(SALIDAS_GENERADAS)}"
    )
    print("\nScript 04 finalizado correctamente.")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        escribir_log(f"ERROR SCRIPT 04 | {type(error).__name__}: {error}")
        print(f"\n[ERROR] El script 04 se detuvo: {type(error).__name__}: {error}")
        raise
