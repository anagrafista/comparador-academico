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
    .main-header {
        font-size: 2.2rem;
        font-weight: 800;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #475569;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
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
    st.image("https://img.icons8.com/fluency/96/diploma.png", width=70)
    st.title("Comparador 2.0")
    st.markdown("**Evaluación Curricular de Posgrados**")
    st.markdown("---")

    st.markdown("### 📌 Reglas de Extracción")
    st.markdown("""
    - **Ámbito exclusivo:** Analiza únicamente el **Plan de Estudios** (omite admisiones, claustrales y metodologías).
    - **Módulos:** Numeración de primer nivel (`1.`, `2.`).
    - **Temas principales:** Viñetas de dos niveles (`1.1`, `1.2`).
    - **Subtemas (`1.1.1`):** Estrictamente excluidos.
    """)

    st.markdown("---")
    st.markdown("### 🏷️ Categorías")
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

# Encabezado
st.markdown('<div class="main-header">🎓 COMPARADOR ACADÉMICO 2.0</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Análisis comparativo exclusivo entre temas principales contenidos en el Plan de Estudios de archivos Word (.docx).</div>', unsafe_allow_html=True)

# Paneles de subida
col_doc1, col_doc2 = st.columns(2)

with col_doc1:
    st.subheader("📄 Documento 1 (Origen)")
    archivo1 = st.file_uploader(
        "Cargar archivo Word (.docx)",
        type=["docx"],
        key=f"doc1_uploader_{st.session_state['uploader_key']}"
    )

with col_doc2:
    st.subheader("📄 Documento 2 (Comparación)")
    archivo2 = st.file_uploader(
        "Cargar archivo Word (.docx)",
        type=["docx"],
        key=f"doc2_uploader_{st.session_state['uploader_key']}"
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
            st.success(f"**{archivo1.name}**: {len(temas_doc1)} temas principales.")
            with st.expander("Ver temas extraídos"):
                for t in temas_doc1:
                    st.write(f"- **[{t['modulo_num']}]** {t['tema_num']} - {t['tema_titulo']}")

        with col_prev2:
            st.success(f"**{archivo2.name}**: {len(temas_doc2)} temas principales.")
            with st.expander("Ver temas extraídos"):
                for t in temas_doc2:
                    st.write(f"- **[{t['modulo_num']}]** {t['tema_num']} - {t['tema_titulo']}")

        st.markdown("---")
        if st.button("🔍 EJECUTAR ANÁLISIS COMPARATIVO", type="primary", use_container_width=True):
            with st.spinner("Comparando sistemáticamente títulos curriculares..."):
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

# Visualización de resultados
if "analisis" in st.session_state:
    datos = st.session_state["analisis"]
    comparaciones = datos["comparaciones"]
    resumen = datos["resumen"]
    nombre_doc1 = datos["nombre_doc1"]
    nombre_doc2 = datos["nombre_doc2"]

    st.markdown("### 📊 Métricas del Análisis")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Temas Doc 1", resumen["total_temas_doc1"])
    k2.metric("Temas Doc 2", resumen["total_temas_doc2"])
    k3.metric("🟢 Literales", resumen["coincidencias_literales"])
    k4.metric("🔵 Altas", resumen["similitudes_altas"])
    k5.metric("🟡 Temáticas", resumen["relaciones_tematicas"])
    k6.metric("⚪ Sin Relación", resumen["sin_coincidencias"])

    st.markdown("---")
    st.markdown("### 📋 Tabla Comparativa de Temas")

    c_filtro1, c_filtro2 = st.columns([1, 2])
    with c_filtro1:
        cat_filtro = st.selectbox(
            "Filtrar por categoría:",
            ["Todas las categorías", "COINCIDENCIA LITERAL", "SIMILITUD ALTA", "RELACIÓN TEMÁTICA O PARALELA", "SIN RELACIÓN IDENTIFICADA"]
        )
    with c_filtro2:
        txt_buscar = st.text_input("Buscar por término en los títulos:", "").lower().strip()

    # Filtrado
    filtrados = []
    for r in comparaciones:
        if cat_filtro != "Todas las categorías" and r["categoria"] != cat_filtro:
            continue
        if txt_buscar:
            if txt_buscar not in r["doc1_tema_titulo"].lower() and txt_buscar not in r["doc2_tema_titulo"].lower():
                continue
        filtrados.append(r)

    st.caption(f"Mostrando **{len(filtrados)}** de **{len(comparaciones)}** comparaciones evaluadas.")

    filas = [{
        "Módulo Doc 1": r["doc1_modulo"],
        "Tema Doc 1": r["doc1_tema_titulo"],
        "Módulo Doc 2": r["doc2_modulo"],
        "Tema Doc 2": r["doc2_tema_titulo"],
        "Similitud": f"{r['similitud']}%",
        "Clasificación": r["categoria"],
        "Recomendación": r["recomendacion"]
    } for r in filtrados]

    df_vista = pd.DataFrame(filas)
    if not df_vista.empty:
        st.dataframe(df_vista, use_container_width=True, height=450)
    else:
        st.info("No hay temas que coincidan con los filtros aplicados.")

    # Detalle técnico
    st.markdown("### 🔍 Justificación Textual de Coincidencias")
    opciones = [
        f"[{r['similitud']}%] {r['doc1_tema_titulo']}  ⟷  {r['doc2_tema_titulo']}"
        for r in filtrados[:50]
    ]
    if opciones:
        sel = st.selectbox("Seleccione una relación para examinar:", opciones)
        if sel:
            idx = opciones.index(sel)
            item = filtrados[idx]
            st.info(f"""
            **Doc 1:** [{item['doc1_modulo']}] {item['doc1_tema_titulo']}  
            **Doc 2:** [{item['doc2_modulo']}] {item['doc2_tema_titulo']}  
            **Similitud:** `{item['similitud']}%` ({item['categoria']})  
            **Recomendación:** *{item['recomendacion']}*  
            **Explicación técnica:** {item['explicacion']}
            """)

    # Exportación
    st.markdown("---")
    st.markdown("### 📥 Exportar Resultados")
    col_e1, col_e2 = st.columns(2)

    # 1. CSV
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

    with col_e1:
        st.download_button(
            label="📊 Descargar Informe CSV (Excel)",
            data=csv_datos,
            file_name="informe_comparador_academico.csv",
            mime="text/csv",
            use_container_width=True
        )

    # 2. TXT
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

    with col_e2:
        st.download_button(
            label="📄 Descargar Informe TXT",
            data=txt_datos,
            file_name="informe_comparador_academico.txt",
            mime="text/plain",
            use_container_width=True
        )
