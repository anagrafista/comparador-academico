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
    initial_sidebar_state="expanded"
)

# Estilos CSS profesionales
st.markdown("""
<style>
    /* Tipografía y contenedor principal */
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2.5rem;
    }
    
    /* Hero Header */
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        border-radius: 16px;
        padding: 24px 30px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.25);
        margin-bottom: 1.8rem;
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hero-badge {
        font-size: 0.75rem;
        background: rgba(255, 255, 255, 0.2);
        padding: 3px 10px;
        border-radius: 20px;
        font-weight: 600;
        vertical-align: middle;
    }
    .hero-subtitle {
        font-size: 1.02rem;
        color: #E2E8F0;
        margin-bottom: 0px;
        max-width: 900px;
        line-height: 1.45;
    }

    /* Tarjetas de Métricas (KPIs) */
    .kpi-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 16px 14px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        text-align: center;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px rgba(0, 0, 0, 0.08);
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 800;
        line-height: 1.1;
        margin-bottom: 4px;
    }
    .kpi-label {
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #64748B;
    }

    /* Badges de categorías */
    .badge-pill {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.3px;
    }
    .badge-literal { background-color: #DEF7EC; color: #03543F; border: 1px solid #84E1BC; }
    .badge-alta { background-color: #E1EFFE; color: #1E429F; border: 1px solid #A4CAFE; }
    .badge-tematica { background-color: #FEF08A; color: #713F12; border: 1px solid #FDE047; }
    .badge-sin { background-color: #F3F4F6; color: #374151; border: 1px solid #D1D5DB; }

    /* Tarjeta de inspección léxica */
    .inspection-card {
        background: #F8FAFC;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        padding: 20px;
        margin-top: 10px;
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

    # Extraer de párrafos
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto:
            lineas.append(texto)

    # Extraer de tablas
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

    # Descartar subtemas de tercer nivel (1.1.1, a), b))
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

        # Excluir subtemas de tercer nivel (1.1.1, 1.1.2)
        if re.match(r"^\s*\d{1,2}\.\d{1,2}\.\d{1,2}", linea):
            continue

        # Tema numérico de segundo nivel (1.1, 1.2)
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

        # Módulo numérico (1., 2.)
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
    """Filtra palabras de parada (stopwords) y conserva términos significativos."""
    return {p for p in texto_normalizado.split() if p not in PALABRAS_PARADA_ES and len(p) >= 3}


def _obtener_raiz(palabra):
    """Extrae la raíz morfológica aproximada de un término en español."""
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
# 4. INTERFAZ DE USUARIO STREAMLIT
# ==============================================================================

# Control de reinicio de cargadores de archivos
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/diploma.png", width=65)
    st.markdown("### Comparador Académico 2.0")
    st.caption("Sistema de Homologación Curricular")
    st.markdown("---")

    st.markdown("#### 📌 Reglas de Extracción")
    st.markdown("""
    - **Ámbito exclusivo:** Analiza únicamente el **Plan de Estudios**.
    - **Módulos:** Numeración de primer nivel (`1.`, `2.`).
    - **Temas principales:** Viñetas de segundo nivel (`1.1`, `1.2`).
    - **Subtemas (`1.1.1`):** **Estrictamente excluidos**.
    """)

    st.markdown("---")
    st.markdown("#### 🏷️ Categorías y Umbrales")
    st.markdown("""
    - 🟢 **Coincidencia Literal** (≥ 95%)
    - 🔵 **Similitud Alta** (72% - 94.9%)
    - 🟡 **Relación Temática** (42% - 71.9%)
    - ⚪ **Sin Relación** (< 42%)
    """)

    st.markdown("---")
    if st.button("🔄 Reiniciar / Limpiar sesión", use_container_width=True):
        nuevo_id = st.session_state.get("uploader_key", 0) + 1
        st.session_state.clear()
        st.session_state["uploader_key"] = nuevo_id
        st.rerun()

# Hero Banner
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">
        <span>🎓 COMPARADOR ACADÉMICO 2.0</span>
        <span class="hero-badge">Web Edition</span>
    </div>
    <p class="hero-subtitle">
        Contraste algorítmico y matemático entre títulos temáticos contenidos exclusivamente en el
        <strong>Plan de Estudios</strong> de dosieres y programas académicos en formato Word (.docx).
    </p>
</div>
""", unsafe_allow_html=True)

# Paneles de subida estilizados
col_doc1, col_doc2 = st.columns(2)

with col_doc1:
    st.markdown("##### 📄 Documento 1 (Origen / Base)")
    st.caption("Dosier del programa cuyos temas se desean contrastar:")
    archivo1 = st.file_uploader(
        "Cargar archivo Word (.docx)",
        type=["docx"],
        key=f"doc1_uploader_{st.session_state['uploader_key']}",
        label_visibility="collapsed"
    )

with col_doc2:
    st.markdown("##### 📄 Documento 2 (Comparación / Destino)")
    st.caption("Dosier del programa contra el cual se evaluarán las equivalencias:")
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
    # TAB 1: TABLERO & GRÁFICOS
    # --------------------------------------------------------------------------
    with tab_dashboard:
        st.markdown("#### 📈 Indicadores Clave de Convalidación")

        # Tarjetas KPI visuales
        kp1, kp2, kp3, kp4, kp5, kp6 = st.columns(6)
        with kp1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value" style="color: #1E3A8A;">{resumen['total_temas_doc1']}</div>
                <div class="kpi-label">Temas Doc 1</div>
            </div>
            """, unsafe_allow_html=True)
        with kp2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value" style="color: #1E3A8A;">{resumen['total_temas_doc2']}</div>
                <div class="kpi-label">Temas Doc 2</div>
            </div>
            """, unsafe_allow_html=True)
        with kp3:
            st.markdown(f"""
            <div class="kpi-card" style="border-top: 3px solid #10B981;">
                <div class="kpi-value" style="color: #059669;">{resumen['coincidencias_literales']}</div>
                <div class="kpi-label">🟢 Literales</div>
            </div>
            """, unsafe_allow_html=True)
        with kp4:
            st.markdown(f"""
            <div class="kpi-card" style="border-top: 3px solid #3B82F6;">
                <div class="kpi-value" style="color: #2563EB;">{resumen['similitudes_altas']}</div>
                <div class="kpi-label">🔵 Altas</div>
            </div>
            """, unsafe_allow_html=True)
        with kp5:
            st.markdown(f"""
            <div class="kpi-card" style="border-top: 3px solid #F59E0B;">
                <div class="kpi-value" style="color: #D97706;">{resumen['relaciones_tematicas']}</div>
                <div class="kpi-label">🟡 Temáticas</div>
            </div>
            """, unsafe_allow_html=True)
        with kp6:
            st.markdown(f"""
            <div class="kpi-card" style="border-top: 3px solid #94A3B8;">
                <div class="kpi-value" style="color: #64748B;">{resumen['sin_coincidencias']}</div>
                <div class="kpi-label">⚪ Sin Relación</div>
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

        # Datos para gráficos
        df_distribucion = pd.DataFrame({
            "Categoría": ["Literal (≥95%)", "Alta (72-94%)", "Temática (42-71%)", "Sin Relación (<42%)"],
            "Cantidad": [
                resumen["coincidencias_literales"],
                resumen["similitudes_altas"],
                resumen["relaciones_tematicas"],
                resumen["sin_coincidencias"]
            ],
            "Color": ["#10B981", "#3B82F6", "#F59E0B", "#94A3B8"]
        })

        with col_chart1:
            st.markdown("##### 🍩 Proporción por Categorías")
            # Gráfico de Dona con Altair
            donut_chart = alt.Chart(df_distribucion).mark_arc(innerRadius=65, stroke="#FFFFFF", strokeWidth=2).encode(
                theta=alt.Theta(field="Cantidad", type="quantitative"),
                color=alt.Color(
                    field="Categoría",
                    type="nominal",
                    scale=alt.Scale(
                        domain=["Literal (≥95%)", "Alta (72-94%)", "Temática (42-71%)", "Sin Relación (<42%)"],
                        range=["#10B981", "#3B82F6", "#F59E0B", "#94A3B8"]
                    ),
                    legend=alt.Legend(orient="bottom", title=None, columns=2)
                ),
                tooltip=[alt.Tooltip("Categoría:N"), alt.Tooltip("Cantidad:Q", title="Temas")]
            ).properties(height=320)
            st.altair_chart(donut_chart, use_container_width=True)

        with col_chart2:
            st.markdown("##### 📊 Cantidad de Temas por Rango")
            # Gráfico de Barras con Altair
            bar_chart = alt.Chart(df_distribucion).mark_bar(cornerRadiusTopLeft=8, cornerRadiusTopRight=8).encode(
                x=alt.X("Categoría:N", title=None, axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Cantidad:Q", title="Número de Temas"),
                color=alt.Color(
                    field="Categoría",
                    type="nominal",
                    scale=alt.Scale(
                        domain=["Literal (≥95%)", "Alta (72-94%)", "Temática (42-71%)", "Sin Relación (<42%)"],
                        range=["#10B981", "#3B82F6", "#F59E0B", "#94A3B8"]
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

        # Filtros de búsqueda y categoría
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

        # Aplicar filtros
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

        # Preparar DataFrame con iconos de estado
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

                # Card de inspección
                st.markdown(f"""
                <div class="inspection-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                        <span style="font-size: 1.1rem; font-weight: 700; color: #1E293B;">Puntuación de Coincidencia: {item['similitud']}%</span>
                        <span class="badge-pill {'badge-literal' if item['categoria']=='COINCIDENCIA LITERAL' else 'badge-alta' if item['categoria']=='SIMILITUD ALTA' else 'badge-tematica' if item['categoria']=='RELACIÓN TEMÁTICA O PARALELA' else 'badge-sin'}">
                            {item['categoria']}
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                st.progress(item['similitud'] / 100.0)

                # Detalle en dos columnas
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

        # Generar CSV en memoria con UTF-8-SIG (compatible con Excel)
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
            <div class="kpi-card" style="text-align: left; padding: 20px;">
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

        # Generar TXT estructurado
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
            <div class="kpi-card" style="text-align: left; padding: 20px;">
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
