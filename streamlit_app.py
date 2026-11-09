"""
================================================================================
COMPARADOR ACADÉMICO 2.0 - SISTEMA WEB DE EVALUACIÓN CURRICULAR
================================================================================
Desarrollado para Streamlit Community Cloud.
Compara sistemáticamente planes de estudio contenidos en documentos Word (.docx),
aislando la sección curricular y contrastando exclusivamente temas principales.
"""

import csv
import io
import os
import re
import unicodedata
from difflib import SequenceMatcher

import altair as alt
import docx
import pandas as pd
import streamlit as st

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Comparador Académico 2.0",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS - Paleta Biovita (Azul Eléctrico + Verde Lima + Blanco + Pizarra)
st.markdown("""
<style>
    /* Ocultar barra lateral izquierda por completo */
    [data-testid="stSidebar"], section[data-testid="stSidebar"] {
        display: none !important;
    }
    [data-testid="collapsedControl"] {
        display: none !important;
    }

    /* Contenedor principal con fondo suave */
    .stApp {
        background-color: #F8FAF7;
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 4rem;
        max-width: 1250px;
    }

    /* Hero Banner estilo Biovita */
    .hero-banner {
        background: #FFFFFF;
        border-radius: 24px;
        padding: 32px 38px;
        border: 1px solid #E6ECE1;
        box-shadow: 0 4px 25px rgba(15, 23, 42, 0.04);
        margin-bottom: 1.8rem;
    }
    .hero-pill {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: #D3F33D;
        color: #0F172A;
        padding: 6px 16px;
        border-radius: 50px;
        font-size: 0.85rem;
        font-weight: 800;
        letter-spacing: 0.3px;
        margin-bottom: 14px;
    }
    .hero-title {
        font-size: 2.35rem;
        font-weight: 900;
        color: #0F172A;
        letter-spacing: -0.8px;
        line-height: 1.18;
        margin-bottom: 10px;
    }
    .hero-title span.accent-blue {
        color: #2251FF;
    }
    .hero-subtitle {
        font-size: 1.05rem;
        color: #64748B;
        line-height: 1.55;
        max-width: 920px;
        margin: 0;
    }

    /* Paneles de carga de archivos */
    .upload-box {
        background: #FFFFFF;
        border-radius: 20px;
        padding: 22px;
        border: 1px solid #E6ECE1;
        box-shadow: 0 2px 10px rgba(15, 23, 42, 0.03);
        margin-bottom: 1rem;
    }
    .upload-label {
        font-size: 1.05rem;
        font-weight: 800;
        color: #0F172A;
        margin-bottom: 4px;
    }
    .upload-desc {
        font-size: 0.88rem;
        color: #64748B;
        margin-bottom: 12px;
    }

    /* Botón Primario: Verde Lima Biovita */
    div[data-testid="stButton"] button[kind="primary"] {
        background-color: #D3F33D !important;
        color: #0F172A !important;
        font-weight: 800 !important;
        font-size: 1.15rem !important;
        padding: 16px 36px !important;
        border-radius: 16px !important;
        border: 1px solid #C4E530 !important;
        box-shadow: 0 4px 16px rgba(211, 243, 61, 0.45) !important;
        letter-spacing: 0.3px !important;
        transition: all 0.2s ease !important;
        margin-top: 8px;
    }
    div[data-testid="stButton"] button[kind="primary"]:hover {
        background-color: #C5E731 !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 24px rgba(211, 243, 61, 0.6) !important;
    }

    /* Botón Secundario Grande de Reinicio al fondo */
    div[data-testid="stButton"] button[kind="secondary"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-weight: 800 !important;
        font-size: 1.12rem !important;
        padding: 18px 36px !important;
        border-radius: 18px !important;
        border: 2px solid #E2E8F0 !important;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.04) !important;
        transition: all 0.25s ease !important;
    }
    div[data-testid="stButton"] button[kind="secondary"]:hover {
        background-color: #F0F4FF !important;
        color: #2251FF !important;
        border-color: #2251FF !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 25px rgba(34, 81, 255, 0.15) !important;
    }

    /* Tarjetas de Métricas (KPIs) estilo Biovita */
    .kpi-card-white {
        background: #FFFFFF;
        border-radius: 18px;
        padding: 20px 14px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.03);
        text-align: center;
    }
    .kpi-card-lime {
        background: #D3F33D;
        border-radius: 18px;
        padding: 20px 14px;
        color: #0F172A;
        box-shadow: 0 4px 16px rgba(211, 243, 61, 0.25);
        text-align: center;
    }
    .kpi-card-blue {
        background: #2251FF;
        border-radius: 18px;
        padding: 20px 14px;
        color: #FFFFFF;
        box-shadow: 0 4px 16px rgba(34, 81, 255, 0.25);
        text-align: center;
    }
    .kpi-card-yellow {
        background: #FEF9C3;
        border-radius: 18px;
        padding: 20px 14px;
        border: 1px solid #FEF08A;
        text-align: center;
    }
    .kpi-card-gray {
        background: #F1F5F9;
        border-radius: 18px;
        padding: 20px 14px;
        border: 1px solid #E2E8F0;
        text-align: center;
    }
    .kpi-num {
        font-size: 2.1rem;
        font-weight: 900;
        line-height: 1.1;
        margin-bottom: 4px;
    }
    .kpi-title {
        font-size: 0.78rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* Pestañas (Tabs) estilo Biovita */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        border-bottom: 2px solid #E6ECE1;
        margin-bottom: 1.5rem;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 1.02rem;
        font-weight: 700;
        color: #64748B;
        border-radius: 12px 12px 0 0;
        padding: 10px 20px;
    }
    .stTabs [aria-selected="true"] {
        color: #2251FF !important;
        border-bottom-color: #2251FF !important;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. PARÁMETROS Y UMBRALES DE COMPARACIÓN
# ==============================================================================
UMBRAL_LITERAL = 95.0      # Coincidencia casi idéntica o literal
UMBRAL_ALTA = 72.0         # Similitud textual alta
UMBRAL_TEMATICA = 42.0     # Relación temática o paralela

PALABRAS_PARADA_ES = {
    "de", "del", "la", "las", "el", "los", "un", "una", "unos", "unas",
    "y", "e", "o", "u", "en", "con", "por", "para", "sobre", "a", "al",
    "ante", "bajo", "cabe", "contra", "desde", "hacia", "hasta", "mediante",
    "sin", "so", "tras", "durante", "segun", "que", "su", "sus", "se", "como",
    "lo", "les", "nos", "te", "mi", "mis", "tu", "tus", "muy", "mas", "pero"
}


# ==============================================================================
# 2. MOTOR DE EXTRACCIÓN CURRICULAR (.DOCX)
# ==============================================================================
def leer_docx(fuente_archivo):
    """
    Extrae texto de documentos Word (.docx), procesando párrafos y tablas.
    Soporta rutas de archivo y buffers en memoria (Streamlit UploadedFile).
    """
    if isinstance(fuente_archivo, str):
        if not os.path.exists(fuente_archivo):
            raise FileNotFoundError(f"No se encontró el archivo: {fuente_archivo}")
        target = fuente_archivo
    elif hasattr(fuente_archivo, "read"):
        contenido_bytes = fuente_archivo.read()
        if hasattr(fuente_archivo, "seek"):
            fuente_archivo.seek(0)
        target = io.BytesIO(contenido_bytes)
    else:
        target = fuente_archivo

    doc = docx.Document(target)
    lineas = []

    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto:
            lineas.append(texto)

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto_celda = celda.text.strip()
                if texto_celda:
                    for sublinea in texto_celda.splitlines():
                        sub_limpia = sublinea.strip()
                        if sub_limpia and sub_limpia not in lineas:
                            lineas.append(sub_limpia)

    if not lineas:
        raise ValueError("El archivo Word no contiene texto legible.")
    return lineas


def cargar_documento(fuente_archivo):
    """Valida que el archivo sea Word (.docx) y extrae sus líneas."""
    if not fuente_archivo:
        raise ValueError("Archivo no especificado.")
    if hasattr(fuente_archivo, "name"):
        if not fuente_archivo.name.lower().endswith(".docx"):
            raise ValueError(f"Formato no admitido en '{fuente_archivo.name}'. Se requiere Word (.docx).")
    return leer_docx(fuente_archivo)


def _es_descarte_o_descripcion(linea):
    """Identifica descripciones, objetivos o metadatos que no son temas."""
    linea_lower = linea.lower().strip()
    palabras_prohibidas = [
        "objetivo", "objetivos", "competencia", "competencias",
        "descripción", "descripcion", "bibliografía", "bibliografia",
        "requisito", "requisitos", "evaluación", "evaluacion",
        "metodología", "metodologia", "profesor", "profesorado",
        "duración", "duracion", "créditos", "creditos", "horas",
        "introducción", "introduccion", "presentación", "presentacion"
    ]
    for p in palabras_prohibidas:
        if linea_lower.startswith(p):
            return True

    palabras = linea.split()
    if len(palabras) > 22 or len(linea) > 140:
        return True

    if linea.endswith(".") and len(palabras) >= 8:
        if re.search(r"\b(el alumno|los estudiantes|aprenderá|analizará|comprenderá)\b", linea_lower):
            return True

    if re.match(r"^\s*(?:\d+\.\d+\.\d+|[a-zA-Z]\))\s*", linea):
        return True

    return False


def _limpiar_titulo_tema(titulo):
    """Retira viñetas o prefijos numéricos para aislar el título limpio."""
    titulo = titulo.strip()
    titulo = re.sub(r"^(?:Tema|TEMA|Unidad|UNIDAD|Capítulo|CAPÍTULO)\s+\d+[\s\.\:\-]+", "", titulo, flags=re.IGNORECASE)
    titulo = re.sub(r"^\d+(?:\.\d+)?[\s\.\:\-]+", "", titulo)
    titulo = re.sub(r"^[\-\•\*\–\—\>]\s*", "", titulo)
    titulo = re.sub(r"\.+$", "", titulo)
    return titulo.strip()


def delimitar_seccion_plan_estudios(lineas):
    """
    Aísla la sección del Plan de Estudios en documentos extensos (44+ páginas),
    descartando admisiones, rankings, claustro y metodologías.
    """
    if not lineas:
        return lineas

    patron_inicio = re.compile(
        r"^(?:(?:\d{1,2}[\.\:\-\s]+)?(?:plan\s+de\s+estudios?|estructura\s+(?:curricular|del\s+programa|modular)|contenido[s]?\s+formativo[s]?|temario))\b",
        re.IGNORECASE
    )
    patron_fin = re.compile(
        r"^(?:(?:\d{1,2}[\.\:\-\s]+)?(?:metodolog[ií]a|claustro|profesorado|admisi[oó]n|requisitos\s+de\s+acceso|salidas\s+profesionales|tasas|contacto))\b",
        re.IGNORECASE
    )

    inicio_idx, fin_idx = None, None
    for i, l in enumerate(lineas):
        l_strip = l.strip()
        if re.search(r"\.{3,}\s*\d+$", l_strip):
            continue
        if inicio_idx is None:
            if patron_inicio.search(l_strip) and len(l_strip.split()) <= 8:
                inicio_idx = i
        else:
            if (i - inicio_idx >= 5) and patron_fin.search(l_strip) and len(l_strip.split()) <= 8:
                fin_idx = i
                break

    if inicio_idx is not None:
        return lineas[inicio_idx:fin_idx] if fin_idx is not None else lineas[inicio_idx:]
    return lineas


def extraer_temas(lineas_o_texto, nombre_doc="Documento"):
    """
    Extrae exclusivamente temas principales (viñetas 1.1) y módulos (1.),
    descartando subtemas (1.1.1) y texto accesorio.
    """
    if isinstance(lineas_o_texto, str):
        lineas_raw = [l.strip() for l in lineas_o_texto.splitlines() if l.strip()]
    elif isinstance(lineas_o_texto, list):
        lineas_raw = [l.strip() for l in lineas_o_texto if isinstance(l, str) and l.strip()]
    else:
        raise ValueError("Entrada no válida para extraer_temas.")

    lineas = delimitar_seccion_plan_estudios(lineas_raw)
    temas_extraidos = []
    modulo_actual_num = "Módulo General"
    modulo_actual_nombre = "Contenido General"
    contador_tema = 0

    patron_modulo_texto = re.compile(
        r"^(?:Módulo|MODULO|MÓDULO|Bloque|BLOQUE|Unidad\s+Didáctica)\s*(\d+|[IVXLCDM]+)?\s*[\:\.\-\–]?\s*(.*)$",
        re.IGNORECASE
    )
    patron_tema_etiquetado = re.compile(
        r"^(?:Tema|TEMA|Unidad|UNIDAD|Capítulo|CAPÍTULO)\s*(\d+|[IVXLCDM]+)?\s*[\:\.\-\–]\s*(.+)$",
        re.IGNORECASE
    )

    i = 0
    while i < len(lineas):
        linea = lineas[i].strip()
        i += 1

        if _es_descarte_o_descripcion(linea):
            continue

        # Excluir subtemas 1.1.1
        if re.match(r"^\s*\d{1,2}\.\d{1,2}\.\d{1,2}", linea):
            continue

        # Tema 1.1
        match_tema_num = re.match(r"^\s*(?:Tema\s+)?(\d{1,2}\.\d{1,2})\s*[\.\:\-\–\)]?\s*(.*)$", linea, re.IGNORECASE)
        if match_tema_num:
            t_num = match_tema_num.group(1)
            t_tit_raw = match_tema_num.group(2).strip()

            if not t_tit_raw and i < len(lineas):
                sig = lineas[i].strip()
                if not re.match(r"^\s*\d", sig) and not _es_descarte_o_descripcion(sig):
                    t_tit_raw = sig
                    i += 1

            if t_tit_raw:
                t_tit_limpio = _limpiar_titulo_tema(t_tit_raw)
                if t_tit_limpio and len(t_tit_limpio) >= 3:
                    contador_tema += 1
                    temas_extraidos.append({
                        "modulo_num": modulo_actual_num,
                        "modulo_nombre": modulo_actual_nombre,
                        "tema_num": str(t_num),
                        "tema_titulo": t_tit_limpio,
                        "tema_completo": f"{t_num} {t_tit_limpio}",
                        "origen": nombre_doc
                    })
            continue

        # Módulo 1.
        match_mod_num = re.match(r"^\s*(?:M[oó]dulo\s+|Bloque\s+)?(\d{1,2})\s*[\.\:\-\–\)]\s*(.*)$", linea, re.IGNORECASE)
        if match_mod_num:
            m_num = match_mod_num.group(1)
            m_nom = match_mod_num.group(2).strip()
            if not m_nom and i < len(lineas):
                sig = lineas[i].strip()
                if not re.match(r"^\s*\d", sig) and not _es_descarte_o_descripcion(sig):
                    m_nom = sig
                    i += 1
            modulo_actual_num = f"Módulo {m_num}"
            modulo_actual_nombre = m_nom if m_nom else f"Módulo {m_num}"
            continue

        # Módulo textual
        match_mod_txt = patron_modulo_texto.match(linea)
        if match_mod_txt:
            num = match_mod_txt.group(1)
            nom = match_mod_txt.group(2).strip()
            modulo_actual_num = f"Módulo {num}" if num else "Módulo"
            modulo_actual_nombre = nom if nom else "Sin nombre"
            continue

        # Tema etiquetado
        match_tema_etiq = patron_tema_etiquetado.match(linea)
        if match_tema_etiq:
            t_num = match_tema_etiq.group(1)
            t_tit_raw = match_tema_etiq.group(2).strip()
            t_tit_limpio = _limpiar_titulo_tema(t_tit_raw)
            if t_tit_limpio and len(t_tit_limpio) >= 3:
                contador_tema += 1
                temas_extraidos.append({
                    "modulo_num": modulo_actual_num,
                    "modulo_nombre": modulo_actual_nombre,
                    "tema_num": f"Tema {t_num}" if t_num else f"Tema {contador_tema}",
                    "tema_titulo": t_tit_limpio,
                    "tema_completo": linea,
                    "origen": nombre_doc
                })
            continue

    return temas_extraidos


# ==============================================================================
# 3. MOTOR DE COMPARACIÓN TEXTUAL Y CLASIFICACIÓN
# ==============================================================================
def normalizar_texto(texto):
    """Normaliza texto: minúsculas, sin tildes ni signos especiales."""
    if not texto or not isinstance(texto, str):
        return ""
    texto = texto.lower().strip()
    texto_nfd = unicodedata.normalize("NFD", texto)
    texto_sin_tildes = "".join(c for c in texto_nfd if unicodedata.category(c) != "Mn")
    texto_limpio = re.sub(r"[^\w\s]", " ", texto_sin_tildes)
    return " ".join(texto_limpio.split())


def extraer_palabras_clave(texto_normalizado):
    """Filtra stopwords y conserva palabras significativas."""
    return {p for p in texto_normalizado.split() if p not in PALABRAS_PARADA_ES and len(p) >= 3}


def _obtener_raiz(palabra):
    """Extrae raíz morfológica en español."""
    p = palabra.lower()
    sufijos = [
        "izaciones", "izacion", "amiento", "amientos", "acionales", "acional",
        "ariales", "arial", "idades", "idad", "mente", "triz", "viles",
        "eras", "eros", "era", "ero", "icas", "icos", "ica", "ico",
        "ales", "al", "ista", "istas", "ivas", "ivos", "iva", "ivo",
        "ante", "entes", "ente", "ores", "oras", "ora", "or", "ando",
        "iendo", "ado", "ada", "ados", "adas", "as", "os", "a", "o", "e"
    ]
    for s in sufijos:
        if p.endswith(s) and len(p) - len(s) >= 4:
            return p[:-len(s)]
    return p


def calcular_similitud(titulo1, titulo2):
    """Calcula similitud combinando SequenceMatcher, palabras clave y raíces."""
    norm1 = normalizar_texto(titulo1)
    norm2 = normalizar_texto(titulo2)

    if norm1 == norm2 and norm1:
        return 100.0, "Coincidencia literal exacta entre ambos títulos."
    if not norm1 or not norm2:
        return 0.0, "Uno de los títulos carece de términos válidos."

    ratio_difflib = SequenceMatcher(None, norm1, norm2).ratio() * 100.0
    palabras1 = extraer_palabras_clave(norm1)
    palabras2 = extraer_palabras_clave(norm2)

    interseccion_exacta = palabras1.intersection(palabras2)
    terminos_relacionados = set(interseccion_exacta)

    raices1 = {p: _obtener_raiz(p) for p in palabras1}
    raices2 = {p: _obtener_raiz(p) for p in palabras2}

    for p1, r1 in raices1.items():
        for p2, r2 in raices2.items():
            if p1 in terminos_relacionados and p2 in terminos_relacionados:
                continue
            if len(r1) >= 4 and len(r2) >= 4 and (r1 == r2 or r1.startswith(r2) or r2.startswith(r1)):
                terminos_relacionados.add(f"{p1} ~ {p2}")
            elif (len(p1) >= 5 and p1 in p2) or (len(p2) >= 5 and p2 in p1):
                terminos_relacionados.add(f"{p1} ~ {p2}")

    total_claves = max(len(palabras1), len(palabras2)) if (palabras1 or palabras2) else 1
    coincidencias_ponderadas = len(interseccion_exacta) + (len(terminos_relacionados) - len(interseccion_exacta)) * 0.85
    proporcion_terminos = (coincidencias_ponderadas / total_claves) * 100.0

    if terminos_relacionados:
        lista_terminos = sorted(list(terminos_relacionados))
        terminos_str = ", ".join(f"'{t}'" for t in lista_terminos[:5])
        explicacion = f"Comparten términos o raíces clave: {terminos_str}."
        porcentaje_combinado = (ratio_difflib * 0.35) + (proporcion_terminos * 0.65)
        porcentaje_final = max(ratio_difflib, porcentaje_combinado)
    else:
        explicacion = "No comparten términos clave principales."
        porcentaje_final = ratio_difflib * 0.65

    return round(min(100.0, max(0.0, porcentaje_final)), 1), explicacion


def clasificar_resultado(porcentaje):
    """Asigna la categoría oficial y recomendación según el porcentaje."""
    if porcentaje >= UMBRAL_LITERAL:
        return "COINCIDENCIA LITERAL", "Posible sustitución (Coincidencia exacta de títulos)", 1
    elif porcentaje >= UMBRAL_ALTA:
        return "SIMILITUD ALTA", "Posible sustitución (Alta coincidencia textual de títulos)", 2
    elif porcentaje >= UMBRAL_TEMATICA:
        return "RELACIÓN TEMÁTICA O PARALELA", "Similitud que requiere revisión (Relación temática)", 3
    else:
        return "SIN RELACIÓN IDENTIFICADA", "Sin relación identificada (Coincidencia insuficiente)", 4


def comparar_temas(temas_doc1, temas_doc2):
    """Compara sistemáticamente los temas de ambos documentos."""
    comparaciones = []
    temas_doc1_con_coincidencia = set()
    conteo_categorias = {
        "COINCIDENCIA LITERAL": 0,
        "SIMILITUD ALTA": 0,
        "RELACIÓN TEMÁTICA O PARALELA": 0,
        "SIN RELACIÓN IDENTIFICADA": 0
    }

    for idx1, t1 in enumerate(temas_doc1):
        mejor_similitud_t1 = -1.0
        mejor_match_t1 = None

        for idx2, t2 in enumerate(temas_doc2):
            similitud, explicacion = calcular_similitud(t1["tema_titulo"], t2["tema_titulo"])
            categoria, recomendacion, nivel = clasificar_resultado(similitud)

            registro = {
                "doc1_modulo": t1.get("modulo_num", "Módulo General"),
                "doc1_tema_titulo": t1.get("tema_titulo", ""),
                "doc2_modulo": t2.get("modulo_num", "Módulo General"),
                "doc2_tema_titulo": t2.get("tema_titulo", ""),
                "similitud": similitud,
                "categoria": categoria,
                "recomendacion": recomendacion,
                "nivel": nivel,
                "explicacion": explicacion
            }
            comparaciones.append(registro)

            if similitud > mejor_similitud_t1:
                mejor_similitud_t1 = similitud
                mejor_match_t1 = registro

        if mejor_match_t1 and mejor_match_t1["nivel"] <= 3:
            temas_doc1_con_coincidencia.add(idx1)
            conteo_categorias[mejor_match_t1["categoria"]] += 1
        else:
            conteo_categorias["SIN RELACIÓN IDENTIFICADA"] += 1

    comparaciones.sort(key=lambda x: x["similitud"], reverse=True)

    resumen = {
        "total_temas_doc1": len(temas_doc1),
        "total_temas_doc2": len(temas_doc2),
        "coincidencias_literales": conteo_categorias["COINCIDENCIA LITERAL"],
        "similitudes_altas": conteo_categorias["SIMILITUD ALTA"],
        "relaciones_tematicas": conteo_categorias["RELACIÓN TEMÁTICA O PARALELA"],
        "sin_coincidencias": conteo_categorias["SIN RELACIÓN IDENTIFICADA"]
    }

    return comparaciones, resumen


# ==============================================================================
# 4. INTERFAZ DE USUARIO STREAMLIT (ESTILO BIOVITA)
# ==============================================================================

# Clave de reinicio de cargadores de archivos
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# Hero Banner estilo Biovita
st.markdown("""
<div class="hero-banner">
    <div class="hero-pill">
        ✨ EVALUACIÓN Y HOMOLOGACIÓN CURRICULAR
    </div>
    <div class="hero-title">
        Comparador Académico <span class="accent-blue">2.0</span>
    </div>
    <p class="hero-subtitle">
        Contraste algorítmico y matemático entre títulos temáticos contenidos exclusivamente en el
        <strong>Plan de Estudios</strong> de dosieres en formato Word (.docx).
    </p>
</div>
""", unsafe_allow_html=True)

# Acordeón limpio con reglas de extracción y umbrales (reemplaza a la barra lateral)
with st.expander("ℹ️ Ver Reglas de Extracción Curricular y Criterios Oficiales"):
    col_reg1, col_reg2 = st.columns(2)
    with col_reg1:
        st.markdown("**📌 Reglas de Extracción:**")
        st.markdown("""
        - **Ámbito exclusivo:** Analiza únicamente el **Plan de Estudios** (descarta admisiones, claustrales y metodología).
        - **Módulos:** Numeración de primer nivel (`1.`, `2.`).
        - **Temas principales:** Viñetas de segundo nivel (`1.1`, `1.2`).
        - **Subtemas (`1.1.1`):** **Estrictamente excluidos**.
        """)
    with col_reg2:
        st.markdown("**🏷️ Categorías y Umbrales Oficiales:**")
        st.markdown("""
        - 🟢 **Coincidencia Literal (≥ 95%):** *Posible sustitución (Coincidencia exacta)*
        - 🔵 **Similitud Alta (72% - 94.9%):** *Posible sustitución (Alta coincidencia textual)*
        - 🟡 **Relación Temática (42% - 71.9%):** *Requiere revisión complementaria*
        - ⚪ **Sin Relación (< 42%):** *Coincidencia insuficiente*
        """)

st.markdown("")

# Paneles de subida estilizados
col_doc1, col_doc2 = st.columns(2)

with col_doc1:
    st.markdown("""
    <div class="upload-box">
        <div class="upload-label">📄 Documento 1 (Origen / Base)</div>
        <div class="upload-desc">Dosier del programa cuyos temas se desean contrastar:</div>
    </div>
    """, unsafe_allow_html=True)
    archivo1 = st.file_uploader(
        "Cargar archivo Word (.docx)",
        type=["docx"],
        key=f"doc1_uploader_{st.session_state['uploader_key']}",
        label_visibility="collapsed"
    )

with col_doc2:
    st.markdown("""
    <div class="upload-box">
        <div class="upload-label">📄 Documento 2 (Comparación / Destino)</div>
        <div class="upload-desc">Dosier del programa contra el cual se evaluarán las equivalencias:</div>
    </div>
    """, unsafe_allow_html=True)
    archivo2 = st.file_uploader(
        "Cargar archivo Word (.docx)",
        type=["docx"],
        key=f"doc2_uploader_{st.session_state['uploader_key']}",
        label_visibility="collapsed"
    )

# Procesamiento cuando se cargan los dos documentos
if archivo1 and archivo2:
    try:
        lineas1 = cargar_documento(archivo1)
        temas_doc1 = extraer_temas(lineas1, nombre_doc=archivo1.name)

        lineas2 = cargar_documento(archivo2)
        temas_doc2 = extraer_temas(lineas2, nombre_doc=archivo2.name)

        col_prev1, col_prev2 = st.columns(2)
        with col_prev1:
            st.success(f"✅ **{archivo1.name}**: {len(temas_doc1)} temas principales identificados.")
            with st.expander(f"Ver {len(temas_doc1)} temas de Doc 1"):
                for t in temas_doc1:
                    st.write(f"- **[{t['modulo_num']}]** {t['tema_num']} - {t['tema_titulo']}")

        with col_prev2:
            st.success(f"✅ **{archivo2.name}**: {len(temas_doc2)} temas principales identificados.")
            with st.expander(f"Ver {len(temas_doc2)} temas de Doc 2"):
                for t in temas_doc2:
                    st.write(f"- **[{t['modulo_num']}]** {t['tema_num']} - {t['tema_titulo']}")

        st.markdown("")
        if st.button("🚀 EJECUTAR ANÁLISIS COMPARATIVO", type="primary", use_container_width=True):
            with st.spinner("Ejecutando comparación algorítmica y léxica entre títulos..."):
                comparaciones, resumen = comparar_temas(temas_doc1, temas_doc2)
                st.session_state["analisis"] = {
                    "comparaciones": comparaciones,
                    "resumen": resumen,
                    "nombre_doc1": archivo1.name,
                    "nombre_doc2": archivo2.name
                }
            st.rerun()

    except Exception as e:
        st.error(f"Error al procesar los documentos: {str(e)}")

# ==============================================================================
# VISUALIZACIÓN DE RESULTADOS CON PESTAÑAS (TABS)
# ==============================================================================
if "analisis" in st.session_state:
    datos = st.session_state["analisis"]
    comparaciones = datos["comparaciones"]
    resumen = datos["resumen"]
    nombre_doc1 = datos["nombre_doc1"]
    nombre_doc2 = datos["nombre_doc2"]

    st.markdown("---")

    # Organización en Pestañas Temáticas
    tab_dashboard, tab_matriz, tab_inspector, tab_exportar = st.tabs([
        "📊 Tablero & Gráficos",
        "📋 Matriz Comparativa",
        "🔍 Inspector Léxico Detallado",
        "📥 Exportación de Informes"
    ])

    # --------------------------------------------------------------------------
    # TAB 1: TABLERO & GRÁFICOS (ESTILO BIOVITA)
    # --------------------------------------------------------------------------
    with tab_dashboard:
        st.markdown("#### 📈 Indicadores Clave de Convalidación")

        # Tarjetas KPI con la estética Biovita
        kp1, kp2, kp3, kp4, kp5, kp6 = st.columns(6)
        with kp1:
            st.markdown(f"""
            <div class="kpi-card-white">
                <div class="kpi-num" style="color: #2251FF;">{resumen['total_temas_doc1']}</div>
                <div class="kpi-title" style="color: #64748B;">Temas Doc 1</div>
            </div>
            """, unsafe_allow_html=True)
        with kp2:
            st.markdown(f"""
            <div class="kpi-card-white">
                <div class="kpi-num" style="color: #2251FF;">{resumen['total_temas_doc2']}</div>
                <div class="kpi-title" style="color: #64748B;">Temas Doc 2</div>
            </div>
            """, unsafe_allow_html=True)
        with kp3:
            st.markdown(f"""
            <div class="kpi-card-lime">
                <div class="kpi-num">{resumen['coincidencias_literales']}</div>
                <div class="kpi-title">🟢 Literales</div>
            </div>
            """, unsafe_allow_html=True)
        with kp4:
            st.markdown(f"""
            <div class="kpi-card-blue">
                <div class="kpi-num">{resumen['similitudes_altas']}</div>
                <div class="kpi-title">🔵 Altas</div>
            </div>
            """, unsafe_allow_html=True)
        with kp5:
            st.markdown(f"""
            <div class="kpi-card-yellow">
                <div class="kpi-num" style="color: #854D0E;">{resumen['relaciones_tematicas']}</div>
                <div class="kpi-title" style="color: #854D0E;">🟡 Temáticas</div>
            </div>
            """, unsafe_allow_html=True)
        with kp6:
            st.markdown(f"""
            <div class="kpi-card-gray">
                <div class="kpi-num" style="color: #475569;">{resumen['sin_coincidencias']}</div>
                <div class="kpi-title" style="color: #64748B;">⚪ Sin Relación</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("")

        # Índice de homologación directa
        total_sustituibles = resumen['coincidencias_literales'] + resumen['similitudes_altas']
        tasa_sustitucion = round((total_sustituibles / max(1, resumen['total_temas_doc1'])) * 100, 1)

        st.info(f"""
        💡 **Resumen Ejecutivo:** Un **{tasa_sustitucion}%** de los temas de **{nombre_doc1}** ({total_sustituibles} temas)
        clasifican con **Coincidencia Literal o Similitud Alta**, lo cual sugiere viabilidad de sustitución/convalidación directa.
        """)

        st.markdown("#### 📊 Distribución Visual de Coincidencias")
        col_chart1, col_chart2 = st.columns([1, 1.3])

        df_distribucion = pd.DataFrame({
            "Categoría": ["Literal (≥95%)", "Alta (72-94%)", "Temática (42-71%)", "Sin Relación (<42%)"],
            "Cantidad": [
                resumen["coincidencias_literales"],
                resumen["similitudes_altas"],
                resumen["relaciones_tematicas"],
                resumen["sin_coincidencias"]
            ]
        })

        with col_chart1:
            st.markdown("##### 🍩 Proporción por Categorías")
            donut_chart = alt.Chart(df_distribucion).mark_arc(innerRadius=65, stroke="#FFFFFF", strokeWidth=3).encode(
                theta=alt.Theta(field="Cantidad", type="quantitative"),
                color=alt.Color(
                    field="Categoría",
                    type="nominal",
                    scale=alt.Scale(
                        domain=["Literal (≥95%)", "Alta (72-94%)", "Temática (42-71%)", "Sin Relación (<42%)"],
                        range=["#A3E635", "#2251FF", "#FACC15", "#CBD5E1"]
                    ),
                    legend=alt.Legend(orient="bottom", title=None, columns=2)
                ),
                tooltip=[alt.Tooltip("Categoría:N"), alt.Tooltip("Cantidad:Q", title="Temas")]
            ).properties(height=320)
            st.altair_chart(donut_chart, use_container_width=True)

        with col_chart2:
            st.markdown("##### 📊 Cantidad de Temas por Nivel")
            bar_chart = alt.Chart(df_distribucion).mark_bar(cornerRadiusTopLeft=10, cornerRadiusTopRight=10).encode(
                x=alt.X("Categoría:N", title=None, axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Cantidad:Q", title="Número de Temas"),
                color=alt.Color(
                    field="Categoría",
                    type="nominal",
                    scale=alt.Scale(
                        domain=["Literal (≥95%)", "Alta (72-94%)", "Temática (42-71%)", "Sin Relación (<42%)"],
                        range=["#A3E635", "#2251FF", "#FACC15", "#CBD5E1"]
                    ),
                    legend=None
                ),
                tooltip=[alt.Tooltip("Categoría:N"), alt.Tooltip("Cantidad:Q")]
            ).properties(height=320)
            st.altair_chart(bar_chart, use_container_width=True)

    # --------------------------------------------------------------------------
    # TAB 2: MATRIZ COMPARATIVA INTERACTIVA
    # --------------------------------------------------------------------------
    with tab_matriz:
        st.markdown("#### 📋 Explorador de Parejas Curriculares")

        f_col1, f_col2, f_col3 = st.columns([1.2, 1.8, 1])
        with f_col1:
            categoria_seleccionada = st.selectbox(
                "Filtrar por Categoría:",
                ["Todas las categorías", "COINCIDENCIA LITERAL", "SIMILITUD ALTA", "RELACIÓN TEMÁTICA O PARALELA", "SIN RELACIÓN IDENTIFICADA"]
            )
        with f_col2:
            termino_busqueda = st.text_input("🔍 Buscar término o palabra clave:", "").lower().strip()
        with f_col3:
            min_similitud = st.slider("Similitud mínima:", min_value=0, max_value=100, value=0, step=5)

        items_filtrados = []
        for r in comparaciones:
            if r["similitud"] < min_similitud:
                continue
            if categoria_seleccionada != "Todas las categorías" and r["categoria"] != categoria_seleccionada:
                continue
            if termino_busqueda:
                t1 = r["doc1_tema_titulo"].lower()
                t2 = r["doc2_tema_titulo"].lower()
                if termino_busqueda not in t1 and termino_busqueda not in t2:
                    continue
            items_filtrados.append(r)

        st.caption(f"Mostrando **{len(items_filtrados)}** de **{len(comparaciones)}** comparaciones evaluadas.")

        filas_tabla = []
        for r in items_filtrados:
            icono = "🟢" if r["categoria"] == "COINCIDENCIA LITERAL" else \
                    "🔵" if r["categoria"] == "SIMILITUD ALTA" else \
                    "🟡" if r["categoria"] == "RELACIÓN TEMÁTICA O PARALELA" else "⚪"

            filas_tabla.append({
                "Estado": icono,
                "Similitud": f"{r['similitud']}%",
                "Módulo Doc 1": r["doc1_modulo"],
                "Tema Doc 1": r["doc1_tema_titulo"],
                "Módulo Doc 2": r["doc2_modulo"],
                "Tema Doc 2": r["doc2_tema_titulo"],
                "Categoría Oficial": r["categoria"],
                "Recomendación": r["recomendacion"]
            })

        df_tabla = pd.DataFrame(filas_tabla)

        if not df_tabla.empty:
            st.dataframe(
                df_tabla,
                use_container_width=True,
                height=480,
                column_config={
                    "Estado": st.column_config.TextColumn(" ", width="small"),
                    "Similitud": st.column_config.TextColumn("Similitud", width="small"),
                    "Tema Doc 1": st.column_config.TextColumn("Tema Doc 1", width="large"),
                    "Tema Doc 2": st.column_config.TextColumn("Tema Doc 2", width="large"),
                    "Categoría Oficial": st.column_config.TextColumn("Categoría", width="medium"),
                    "Recomendación": st.column_config.TextColumn("Recomendación Orientativa", width="large")
                }
            )
        else:
            st.warning("No se encontraron temas con los criterios de filtro aplicados.")

    # --------------------------------------------------------------------------
    # TAB 3: INSPECTOR LÉXICO DETALLADO
    # --------------------------------------------------------------------------
    with tab_inspector:
        st.markdown("#### 🔍 Análisis Profundo de Justificación Textual")
        st.caption("Selecciona cualquier pareja de temas para inspeccionar los términos compartidos y la justificación algorítmica:")

        lista_opciones = [
            f"[{r['similitud']}%] {r['doc1_tema_titulo']}  ⟷  {r['doc2_tema_titulo']}"
            for r in items_filtrados[:60]
        ] if items_filtrados else []

        if lista_opciones:
            opcion_sel = st.selectbox("Seleccione tema a inspeccionar:", lista_opciones)
            if opcion_sel:
                idx_sel = lista_opciones.index(opcion_sel)
                item = items_filtrados[idx_sel]

                st.progress(item['similitud'] / 100.0)

                c_insp1, c_insp2 = st.columns(2)
                with c_insp1:
                    st.markdown(f"**📄 Documento 1 ({nombre_doc1}):**")
                    st.info(f"**Módulo:** {item['doc1_modulo']}\n\n**Título:** {item['doc1_tema_titulo']}")

                with c_insp2:
                    st.markdown(f"**📄 Documento 2 ({nombre_doc2}):**")
                    st.info(f"**Módulo:** {item['doc2_modulo']}\n\n**Título:** {item['doc2_tema_titulo']}")

                st.markdown("##### 💬 Justificación Léxica y Semántica")
                st.write(f"🔎 **Términos compartidos:** {item['explicacion']}")
                st.write(f"🎯 **Recomendación Académica:** *{item['recomendacion']}*")

        else:
            st.info("Ajusta los filtros en la pestaña 'Matriz Comparativa' para inspeccionar resultados.")

    # --------------------------------------------------------------------------
    # TAB 4: EXPORTACIÓN DE INFORMES
    # --------------------------------------------------------------------------
    with tab_exportar:
        st.markdown("#### 📥 Descargar Resultados del Análisis")
        st.caption("Exporta la matriz completa de comparaciones y estadísticas para su uso en Excel o archivo de texto:")

        exp_col1, exp_col2 = st.columns(2)

        salida_csv = io.StringIO()
        escritor = csv.writer(salida_csv)
        escritor.writerow(["COMPARADOR ACADÉMICO 2.0 - INFORME COMPARATIVO"])
        escritor.writerow(["Documento 1", nombre_doc1])
        escritor.writerow(["Documento 2", nombre_doc2])
        escritor.writerow([])
        escritor.writerow(["Módulo Doc 1", "Tema Doc 1", "Módulo Doc 2", "Tema Doc 2", "Similitud (%)", "Clasificación", "Recomendación", "Explicación"])
        for r in comparaciones:
            escritor.writerow([
                r["doc1_modulo"], r["doc1_tema_titulo"],
                r["doc2_modulo"], r["doc2_tema_titulo"],
                f"{r['similitud']}%", r["categoria"], r["recomendacion"], r["explicacion"]
            ])
        csv_datos = salida_csv.getvalue().encode("utf-8-sig")

        with exp_col1:
            st.markdown("""
            <div style="background: #FFFFFF; border-radius: 18px; padding: 22px; border: 1px solid #E2E8F0;">
                <h5 style="margin-bottom: 8px;">📊 Formato CSV (Microsoft Excel)</h5>
                <p style="font-size: 0.88rem; color: #64748B;">
                    Ideal para abrir directamente en Excel, realizar tablas dinámicas o integrar con otros sistemas de gestión académica.
                </p>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("")
            st.download_button(
                label="📥 Descargar Informe CSV (Excel)",
                data=csv_datos,
                file_name="informe_comparador_academico.csv",
                mime="text/csv",
                type="primary",
                use_container_width=True
            )

        salida_txt = io.StringIO()
        salida_txt.write("=" * 80 + "\n")
        salida_txt.write("          COMPARADOR ACADÉMICO 2.0 - INFORME COMPARATIVO\n")
        salida_txt.write("=" * 80 + "\n\n")
        salida_txt.write(f"Documento 1: {nombre_doc1}\n")
        salida_txt.write(f"Documento 2: {nombre_doc2}\n")
        salida_txt.write(f"Coincidencias Literales: {resumen['coincidencias_literales']}\n")
        salida_txt.write(f"Similitudes Altas:       {resumen['similitudes_altas']}\n")
        salida_txt.write(f"Relaciones Temáticas:    {resumen['relaciones_tematicas']}\n")
        salida_txt.write(f"Sin Relación:            {resumen['sin_coincidencias']}\n\n")
        salida_txt.write("-" * 80 + "\nDETALLE DE COMPARACIONES\n" + "-" * 80 + "\n\n")
        for i, r in enumerate(comparaciones, 1):
            salida_txt.write(f"[{i}] {r['similitud']}% | {r['categoria']}\n")
            salida_txt.write(f"    Doc 1: [{r['doc1_modulo']}] {r['doc1_tema_titulo']}\n")
            salida_txt.write(f"    Doc 2: [{r['doc2_modulo']}] {r['doc2_tema_titulo']}\n")
            salida_txt.write(f"    Recomendación: {r['recomendacion']}\n")
            salida_txt.write(f"    Explicación:   {r['explicacion']}\n\n")
        txt_datos = salida_txt.getvalue().encode("utf-8")

        with exp_col2:
            st.markdown("""
            <div style="background: #FFFFFF; border-radius: 18px; padding: 22px; border: 1px solid #E2E8F0;">
                <h5 style="margin-bottom: 8px;">📄 Formato TXT Estructurado</h5>
                <p style="font-size: 0.88rem; color: #64748B;">
                    Documento de texto plano legible, perfecto para anexar como informe técnico o evidencia documental en expedientes.
                </p>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("")
            st.download_button(
                label="📥 Descargar Informe TXT",
                data=txt_datos,
                file_name="informe_comparador_academico.txt",
                mime="text/plain",
                use_container_width=True
            )

# ==============================================================================
# BOTÓN DE REINICIO GRANDE AL PIE DE PÁGINA (SIEMPRE VISIBLE CUANDO HAY ARCHIVOS)
# ==============================================================================
if archivo1 or archivo2 or "analisis" in st.session_state:
    st.markdown("<br><hr style='border-top: 1px solid #E2E8F0; margin: 2.5rem 0 2rem 0;'>", unsafe_allow_html=True)
    col_reset_l, col_reset_c, col_reset_r = st.columns([1, 2.5, 1])
    with col_reset_c:
        if st.button("🔄 REINICIAR Y CARGAR NUEVOS DOCUMENTOS", type="secondary", use_container_width=True, key="btn_reiniciar_grande"):
            nuevo_id = st.session_state.get("uploader_key", 0) + 1
            st.session_state.clear()
            st.session_state["uploader_key"] = nuevo_id
            st.rerun()
