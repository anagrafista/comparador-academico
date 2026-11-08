#!/usr/bin/env python3
"""
================================================================================
PROYECTO: COMPARADOR ACADÉMICO 2.0
Aplicación de escritorio en Python para análisis y posible equivalencia de temas
================================================================================

Este archivo contiene el CÓDIGO COMPLETO Y AUTÓNOMO de la aplicación.
No depende de submódulos externos, integrando en un solo archivo:
1. Lectura y extracción exclusiva de temas desde documentos Word (.docx).
2. Algoritmos de normalización lingüística, raíces léxicas y cálculo de similitud.
3. Clasificación en 4 categorías oficiales y recomendaciones académicas.
4. Exportación de informes en CSV y TXT.
5. Interfaz gráfica interactiva desarrollada con Tkinter.

CUMPLIMIENTO DE ESTRUCTURAS OBLIGATORIAS:
- if: Evaluación de condiciones y validaciones.
- elif: Clasificación en 4 categorías según umbrales de similitud.
- else: Gestión de casos residuales (sin relación identificada).
- for: Ciclos matriciales para contrastar todos los temas de ambos documentos.
- while: Ciclo de despacho continuo y reinicio de análisis sin cerrar la app.
================================================================================
"""

import os
import re
import csv
import queue
import threading
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import docx
except ImportError:
    docx = None

# ==============================================================================
# CONFIGURACIÓN Y CONSTANTES DEL SISTEMA
# ==============================================================================

# Umbrales oficiales de clasificación (fáciles de modificar)
UMBRAL_LITERAL = 95.0      # Categoría 1: Coincidencia Literal (>= 95%)
UMBRAL_ALTA = 72.0         # Categoría 2: Similitud Alta (72% - 94.9%)
UMBRAL_TEMATICA = 42.0     # Categoría 3: Relación Temática o Paralela (42% - 71.9%)
                           # Categoría 4: Sin Relación Identificada (< 42%)

# Palabras vacías en español (stopwords) para filtrar términos irrelevantes
PALABRAS_PARADA_ES = {
    "de", "del", "la", "las", "el", "los", "un", "una", "unos", "unas",
    "y", "e", "o", "u", "en", "con", "por", "para", "sobre", "a", "al",
    "ante", "bajo", "cabe", "contra", "desde", "hacia", "hasta", "mediante",
    "sin", "so", "tras", "durante", "segun", "que", "su", "sus", "se", "como",
    "lo", "les", "nos", "te", "mi", "mis", "tu", "tus", "muy", "mas", "pero"
}


# ==============================================================================
# 1. FUNCIONES DE LECTURA Y EXTRACCIÓN DE DOCUMENTOS WORD (.DOCX)
# ==============================================================================

def leer_docx(ruta_archivo):
    """
    Lee y extrae las líneas de texto de un archivo Microsoft Word (.docx)
    procesando párrafos y tablas.
    """
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")

    if docx is None:
        raise ImportError("La biblioteca python-docx no está instalada. Ejecute: pip install python-docx")

    try:
        doc = docx.Document(ruta_archivo)
        lineas = []

        # Extraer párrafos
        for p in doc.paragraphs:
            texto = p.text.strip()
            if texto:
                lineas.append(texto)

        # Extraer contenido de tablas
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

    except Exception as e:
        if "no contiene texto" in str(e):
            raise
        raise RuntimeError(f"Error al leer el archivo Word: {str(e)}")


def cargar_documento(ruta_archivo):
    """
    Función obligatoria para cargar y validar un documento Word (.docx).
    """
    if not ruta_archivo or not isinstance(ruta_archivo, str):
        raise ValueError("Ruta de archivo no especificada o inválida.")

    extension = os.path.splitext(ruta_archivo)[1].lower()
    if extension != ".docx":
        raise ValueError(
            f"Formato no admitido: '{extension}'. La aplicación opera exclusivamente con documentos Word (.docx)."
        )

    if not os.path.isfile(ruta_archivo):
        raise FileNotFoundError(f"El archivo especificado no existe: {ruta_archivo}")

    return leer_docx(ruta_archivo)


def _es_descarte_o_descripcion(linea):
    """
    Determina si una línea es un párrafo descriptivo, objetivo, competencia,
    bibliografía u otro elemento que NO debe considerarse título.
    """
    linea_lower = linea.lower().strip()

    palabras_prohibidas = [
        "objetivo", "objetivos", "competencia", "competencias",
        "descripción", "descripcion", "bibliografía", "bibliografia",
        "requisito", "requisitos", "evaluación", "evaluacion",
        "metodología", "metodologia", "profesor", "profesorado",
        "duración", "duracion", "créditos", "creditos", "horas",
        "introducción", "introduccion", "presentación", "presentacion",
        "criterios de evaluación", "nota:", "observaciones:"
    ]

    for p in palabras_prohibidas:
        if linea_lower.startswith(p):
            return True

    # Párrafos explicativos extensos
    palabras = linea.split()
    if len(palabras) > 22 or len(linea) > 140:
        return True

    # Oraciones largas terminadas en punto
    if linea.endswith(".") and len(palabras) >= 8:
        if re.search(r"\b(el alumno|los estudiantes|aprenderá|analizará|comprenderá|desarrollará)\b", linea_lower):
            return True

    # Subtemas profundos de tercer nivel (1.1.1, a), b))
    if re.match(r"^\s*(?:\d+\.\d+\.\d+|[a-zA-Z]\)|\([a-zA-Z]\))\s*", linea):
        return True

    return False


def _limpiar_titulo_tema(titulo):
    """
    Limpia el título de un tema retirando viñetas o numeraciones iniciales.
    """
    titulo = titulo.strip()
    titulo = re.sub(r"^(?:Tema|TEMA|Unidad|UNIDAD|Capítulo|CAPÍTULO)\s+\d+[\s\.\:\-]+", "", titulo, flags=re.IGNORECASE)
    titulo = re.sub(r"^\d+(?:\.\d+)?[\s\.\:\-]+", "", titulo)
    titulo = re.sub(r"^[\-\•\*\–\—\>]\s*", "", titulo)
    titulo = re.sub(r"\.+$", "", titulo)
    return titulo.strip()


def delimitar_seccion_plan_estudios(lineas):
    """
    Aísla exclusivamente las páginas correspondientes al 'Plan de estudios'
    en documentos extensos (ej. 44+ páginas), descartando la introducción,
    rankings, metodologías, admisiones y claustro de profesores.
    """
    if not lineas:
        return lineas

    patron_inicio = re.compile(
        r"^(?:(?:\d{1,2}[\.\:\-\s]+)?(?:plan\s+de\s+estudios?|estructura\s+(?:curricular|del\s+programa|modular|del\s+plan\s+de\s+estudios?)|contenido[s]?\s+(?:acad[eé]mico[s]?|formativo[s]?|del\s+programa)|temario(?:\s+(?:general|del\s+m[aá]ster|del\s+programa))?|programa\s+acad[eé]mico|malla\s+curricular))\b",
        re.IGNORECASE
    )

    patron_fin = re.compile(
        r"^(?:(?:\d{1,2}[\.\:\-\s]+)?(?:metodolog[ií]a|claustro|profesorado|equipo\s+docente|admisi[oó]n|proceso\s+de\s+admisi[oó]n|requisitos\s+de\s+acceso|matr[ií]cula|salidas\s+profesionales|bolsa\s+de\s+empleo|sistema\s+de\s+evaluaci[oó]n|evaluaci[oó]n\s+y\s+titulaci[oó]n|financiaci[oó]n|tasas|precios\s+y\s+becas|preguntas\s+frecuentes|contacto))\b",
        re.IGNORECASE
    )

    inicio_idx = None
    fin_idx = None

    for i, l in enumerate(lineas):
        l_strip = l.strip()
        # Descartar líneas del índice que contienen puntos sucesivos hacia números de página
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
        if fin_idx is not None:
            return lineas[inicio_idx:fin_idx]
        else:
            return lineas[inicio_idx:]

    return lineas


def extraer_temas(lineas_o_texto, nombre_doc="Documento"):
    """
    Función obligatoria para extraer exclusivamente los títulos de los temas.
    
    Reconoce:
    - Módulo: numeración 1., 2., Módulo 1, etc.
    - Tema principal: numeración de dos niveles 1.1, 1.2, 2.1, etc.
    - Subtemas (1.1.1, 1.1.2, X.Y.Z): EXCLUIDOS POR COMPLETO.
    """
    if isinstance(lineas_o_texto, str):
        lineas_raw = [l.strip() for l in lineas_o_texto.splitlines() if l.strip()]
    elif isinstance(lineas_o_texto, list):
        lineas_raw = [l.strip() for l in lineas_o_texto if isinstance(l, str) and l.strip()]
    else:
        raise ValueError("Entrada no válida para extraer_temas. Se requiere lista o texto.")

    # 1. Delimitar la sección del Plan de Estudios
    lineas = delimitar_seccion_plan_estudios(lineas_raw)

    temas_extraidos = []
    modulo_actual_num = "Módulo General"
    modulo_actual_nombre = "Contenido General"
    contador_tema = 0

    patron_modulo_texto = re.compile(
        r"^(?:Módulo|MODULO|MÓDULO|Bloque|BLOQUE|Unidad\s+Didáctica|UNIDAD\s+DIDÁCTICA|Área|AREA)\s*(\d+|[IVXLCDM]+)?\s*[\:\.\-\–]?\s*(.*)$",
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

        # A. Descartar párrafos descriptivos y objetivos
        if _es_descarte_o_descripcion(linea):
            continue

        # B. DESCARTE ESTRICTO DE SUBTEMAS: 1.1.1, 1.1.2, X.Y.Z
        if re.match(r"^\s*\d{1,2}\.\d{1,2}\.\d{1,2}", linea):
            continue

        # C. DETECCIÓN DE TEMA NUMÉRICO PRINCIPAL: 1.1, 1.2, 2.1 (X.Y)
        match_tema_num = re.match(
            r"^\s*(?:Tema\s+)?(\d{1,2}\.\d{1,2})\s*[\.\:\-\–\)]?\s*(.*)$",
            linea,
            re.IGNORECASE
        )
        if match_tema_num:
            t_num = match_tema_num.group(1)
            t_tit_raw = match_tema_num.group(2).strip()

            # Si el título viene en la línea siguiente
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

        # D. DETECCIÓN DE MÓDULO NUMÉRICO: '1.', '2.', '1. Nombre'
        match_mod_num = re.match(
            r"^\s*(?:M[oó]dulo\s+|Bloque\s+|Unidad\s+)?(\d{1,2})\s*[\.\:\-\–\)]\s*(.*)$",
            linea,
            re.IGNORECASE
        )
        if match_mod_num:
            m_num = match_mod_num.group(1)
            m_nom = match_mod_num.group(2).strip()

            # Si el nombre viene en la siguiente línea
            if not m_nom and i < len(lineas):
                sig = lineas[i].strip()
                if not re.match(r"^\s*\d", sig) and not _es_descarte_o_descripcion(sig):
                    m_nom = sig
                    i += 1

            modulo_actual_num = f"Módulo {m_num}"
            modulo_actual_nombre = m_nom if m_nom else f"Módulo {m_num}"
            continue

        # E. DETECCIÓN DE MÓDULO TEXTUAL (ej. "Módulo I: Finanzas")
        match_mod_txt = patron_modulo_texto.match(linea)
        if match_mod_txt:
            num = match_mod_txt.group(1)
            nom = match_mod_txt.group(2).strip()
            modulo_actual_num = f"Módulo {num}" if num else "Módulo"
            modulo_actual_nombre = nom if nom else "Sin nombre específico"
            continue

        # F. DETECCIÓN DE TEMA ETIQUETADO EXPLÍCITO (ej. "Tema 1: Planificación")
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

    # Modo dosier simple (si no se detectaron temas con las reglas numéricas)
    if not temas_extraidos:
        for linea in lineas:
            if _es_descarte_o_descripcion(linea):
                continue
            if re.match(r"^\s*\d{1,2}\.\d{1,2}\.\d{1,2}", linea):
                continue
            palabras = linea.split()
            if 2 <= len(palabras) <= 12 and not linea.endswith("."):
                contador_tema += 1
                t_tit_limpio = _limpiar_titulo_tema(linea)
                temas_extraidos.append({
                    "modulo_num": "Módulo General",
                    "modulo_nombre": "Temario General",
                    "tema_num": f"Tema {contador_tema}",
                    "tema_titulo": t_tit_limpio,
                    "tema_completo": linea,
                    "origen": nombre_doc
                })

    return temas_extraidos


# ==============================================================================
# 2. FUNCIONES DE NORMALIZACIÓN, SIMILITUD Y COMPARACIÓN DE TEMAS
# ==============================================================================

def normalizar_texto(texto):
    """
    Función obligatoria para normalizar texto:
    - Conversión a minúsculas.
    - Supresión de tildes y diacríticos (Unicode NFD).
    - Eliminación de signos de puntuación y espacios duplicados.
    """
    if not texto or not isinstance(texto, str):
        return ""

    texto = texto.lower().strip()
    texto_nfd = unicodedata.normalize("NFD", texto)
    texto_sin_tildes = "".join(c for c in texto_nfd if unicodedata.category(c) != "Mn")
    texto_limpio = re.sub(r"[^\w\s]", " ", texto_sin_tildes)
    return " ".join(texto_limpio.split())


def _obtener_raiz(palabra):
    """
    Obtiene la raíz morfológica aproximada de una palabra en español
    para reconocer términos afines (empresa/empresarial, automóvil/automotriz).
    """
    p = palabra.lower()
    sufijos = [
        "izaciones", "izacion", "amiento", "amientos",
        "acionales", "acional", "ariales", "arial",
        "idades", "idad", "mente", "triz", "viles",
        "eras", "eros", "era", "ero", "icas", "icos",
        "ica", "ico", "ales", "al", "ista", "istas",
        "ivas", "ivos", "iva", "ivo", "ante", "entes",
        "ente", "ores", "oras", "ora", "or", "ando",
        "iendo", "ado", "ada", "ados", "adas", "as", "os", "a", "o", "e"
    ]
    for s in sufijos:
        if p.endswith(s) and len(p) - len(s) >= 4:
            return p[:-len(s)]
    return p


def extraer_palabras_clave(texto_normalizado):
    """Extrae palabras clave significativas excluyendo stopwords."""
    palabras = texto_normalizado.split()
    clave = set()
    for p in palabras:
        if p not in PALABRAS_PARADA_ES and len(p) >= 3:
            clave.add(p)
    return clave


def clasificar_resultado(porcentaje):
    """
    Función obligatoria para clasificar el resultado en una de las 4 categorías.
    
    ESTRUCTURA OBLIGATORIA: if / elif / else
    """
    if porcentaje >= UMBRAL_LITERAL:
        categoria = "COINCIDENCIA LITERAL"
        recomendacion = "Posible sustitución (Coincidencia exacta de títulos)"
        nivel = 1
    elif porcentaje >= UMBRAL_ALTA:
        categoria = "SIMILITUD ALTA"
        recomendacion = "Posible sustitución (Alta coincidencia textual de títulos)"
        nivel = 2
    elif porcentaje >= UMBRAL_TEMATICA:
        categoria = "RELACIÓN TEMÁTICA O PARALELA"
        recomendacion = "Similitud que requiere revisión (Relación temática, examinar contenido)"
        nivel = 3
    else:
        categoria = "SIN RELACIÓN IDENTIFICADA"
        recomendacion = "Sin relación identificada (Coincidencia insuficiente)"
        nivel = 4

    return categoria, recomendacion, nivel


def calcular_similitud(titulo1, titulo2):
    """
    Función obligatoria para calcular la similitud matemática real entre títulos.
    Combina difflib, coincidencia de raíces morfológicas y términos clave.
    """
    norm1 = normalizar_texto(titulo1)
    norm2 = normalizar_texto(titulo2)

    # Coincidencia literal exacta tras normalización
    if norm1 == norm2 and norm1:
        return 100.0, "Coincidencia literal exacta entre ambos títulos."

    if not norm1 or not norm2:
        return 0.0, "Uno de los títulos está vacío o carece de términos válidos."

    # Similitud de secuencias con difflib
    ratio_difflib = SequenceMatcher(None, norm1, norm2).ratio() * 100.0

    # Palabras clave y raíces
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
    else:
        explicacion = "No comparten términos clave principales."

    if palabras1 and palabras2 and palabras1 == palabras2:
        return 98.0, f"Mismos términos clave en distinta disposición: {terminos_str}."

    if terminos_relacionados:
        porcentaje_combinado = (ratio_difflib * 0.35) + (proporcion_terminos * 0.65)
        porcentaje_final = max(ratio_difflib, porcentaje_combinado)
    else:
        porcentaje_final = ratio_difflib * 0.65

    porcentaje_final = min(100.0, max(0.0, porcentaje_final))
    return round(porcentaje_final, 1), explicacion


def comparar_temas(temas_doc1, temas_doc2):
    """
    Función obligatoria para contrastar sistemáticamente los temas de ambos documentos.
    
    ESTRUCTURA OBLIGATORIA: Ciclos 'for' anidados.
    """
    comparaciones = []
    temas_doc1_con_coincidencia = set()

    conteo_categorias = {
        "COINCIDENCIA LITERAL": 0,
        "SIMILITUD ALTA": 0,
        "RELACIÓN TEMÁTICA O PARALELA": 0,
        "SIN RELACIÓN IDENTIFICADA": 0
    }

    # ==========================================================================
    # CICLOS 'FOR' OBLIGATORIOS
    # ==========================================================================
    for idx1, t1 in enumerate(temas_doc1):
        mejor_similitud_t1 = -1.0
        mejor_match_t1 = None

        for idx2, t2 in enumerate(temas_doc2):
            similitud, explicacion = calcular_similitud(
                t1["tema_titulo"],
                t2["tema_titulo"]
            )

            categoria, recomendacion, nivel = clasificar_resultado(similitud)

            registro = {
                "doc1_modulo": t1.get("modulo_num", "Módulo General"),
                "doc1_modulo_nom": t1.get("modulo_nombre", ""),
                "doc1_tema_num": t1.get("tema_num", ""),
                "doc1_tema_titulo": t1.get("tema_titulo", ""),
                "doc2_modulo": t2.get("modulo_num", "Módulo General"),
                "doc2_modulo_nom": t2.get("modulo_nombre", ""),
                "doc2_tema_num": t2.get("tema_num", ""),
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

    temas_sin_coincidencia = [
        t1 for idx1, t1 in enumerate(temas_doc1) if idx1 not in temas_doc1_con_coincidencia
    ]

    resumen = {
        "total_temas_doc1": len(temas_doc1),
        "total_temas_doc2": len(temas_doc2),
        "coincidencias_literales": conteo_categorias["COINCIDENCIA LITERAL"],
        "similitudes_altas": conteo_categorias["SIMILITUD ALTA"],
        "relaciones_tematicas": conteo_categorias["RELACIÓN TEMÁTICA O PARALELA"],
        "sin_coincidencias": conteo_categorias["SIN RELACIÓN IDENTIFICADA"],
        "total_comparaciones_realizadas": len(comparaciones)
    }

    return comparaciones, resumen, temas_sin_coincidencia


# ==============================================================================
# 3. FUNCIÓN DE EXPORTACIÓN DE INFORMES
# ==============================================================================

def generar_informe(resultados, resumen, info_doc1, info_doc2, ruta_salida, formato="csv"):
    """
    Función obligatoria para exportar el informe comparativo en CSV o TXT.
    """
    if not ruta_salida:
        raise ValueError("Ruta de archivo de salida no especificada.")

    fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if formato.lower() == "csv" or ruta_salida.lower().endswith(".csv"):
        with open(ruta_salida, mode="w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["COMPARADOR ACADÉMICO 2.0 - INFORME COMPARATIVO DE TEMAS"])
            writer.writerow(["Fecha del análisis", fecha_hora])
            writer.writerow(["Documento 1 (Origen)", info_doc1.get("nombre", "Desconocido")])
            writer.writerow(["Documento 2 (Comparación)", info_doc2.get("nombre", "Desconocido")])
            writer.writerow([])
            writer.writerow(["RESUMEN GENERAL"])
            writer.writerow(["Temas Documento 1", resumen.get("total_temas_doc1", 0)])
            writer.writerow(["Temas Documento 2", resumen.get("total_temas_doc2", 0)])
            writer.writerow(["Coincidencias literales", resumen.get("coincidencias_literales", 0)])
            writer.writerow(["Similitudes altas", resumen.get("similitudes_altas", 0)])
            writer.writerow(["Relaciones temáticas", resumen.get("relaciones_tematicas", 0)])
            writer.writerow(["Temas sin coincidencias", resumen.get("sin_coincidencias", 0)])
            writer.writerow([])
            writer.writerow([
                "Módulo Doc 1", "Tema Doc 1",
                "Módulo Doc 2", "Tema Doc 2",
                "Similitud (%)", "Clasificación", "Recomendación", "Explicación Técnica"
            ])
            for r in resultados:
                writer.writerow([
                    f"{r['doc1_modulo']} - {r['doc1_modulo_nom']}".strip(" -"),
                    r["doc1_tema_titulo"],
                    f"{r['doc2_modulo']} - {r['doc2_modulo_nom']}".strip(" -"),
                    r["doc2_tema_titulo"],
                    f"{r['similitud']}%",
                    r["categoria"],
                    r["recomendacion"],
                    r["explicacion"]
                ])
    else:
        with open(ruta_salida, mode="w", encoding="utf-8") as f:
            sep = "=" * 90
            f.write(f"{sep}\n          COMPARADOR ACADÉMICO 2.0 - INFORME COMPARATIVO\n{sep}\n\n")
            f.write(f"Fecha: {fecha_hora}\n")
            f.write(f"Doc 1: {info_doc1.get('nombre')}\nDoc 2: {info_doc2.get('nombre')}\n\n")
            f.write(f"Total Temas Doc 1: {resumen.get('total_temas_doc1')}\n")
            f.write(f"Total Temas Doc 2: {resumen.get('total_temas_doc2')}\n")
            f.write(f"Coincidencias Literales: {resumen.get('coincidencias_literales')}\n")
            f.write(f"Similitudes Altas: {resumen.get('similitudes_altas')}\n")
            f.write(f"Relaciones Temáticas: {resumen.get('relaciones_tematicas')}\n\n")
            f.write(f"{'-' * 90}\nDETALLE DE TEMAS COMPARADOS\n{'-' * 90}\n\n")
            for i, r in enumerate(resultados, 1):
                f.write(f"[{i}] {r['similitud']}% | {r['categoria']}\n")
                f.write(f"    Doc 1: [{r['doc1_modulo']}] {r['doc1_tema_titulo']}\n")
                f.write(f"    Doc 2: [{r['doc2_modulo']}] {r['doc2_tema_titulo']}\n")
                f.write(f"    Recomendación: {r['recomendacion']}\n")
                f.write(f"    Explicación:   {r['explicacion']}\n\n")

    return ruta_salida


# ==============================================================================
# 4. INTERFAZ GRÁFICA DE USUARIO (TKINTER) CON CICLO WHILE
# ==============================================================================

class ComparadorAcademicoGUI:
    """Clase de la interfaz gráfica del Comparador Académico 2.0."""

    def __init__(self, root):
        self.root = root
        self.root.title("COMPARADOR ACADÉMICO 2.0")
        self.root.geometry("1180x820")
        self.root.minsize(980, 680)

        self.doc1_ruta = None
        self.doc1_datos = None
        self.doc1_nombre = ""

        self.doc2_ruta = None
        self.doc2_datos = None
        self.doc2_nombre = ""

        self.resultados_actuales = []
        self.resumen_actual = {}
        self.temas_sin_match = []

        # Cola de control para ciclo while
        self.cola_solicitudes = queue.Queue()
        self.app_activa = True
        self.analisis_en_curso = False

        self._configurar_estilos()
        self._construir_interfaz()
        self._iniciar_ciclo_control()

    def _configurar_estilos(self):
        self.colores = {
            "fondo": "#F8FAFC",
            "panel": "#FFFFFF",
            "primario": "#2563EB",
            "texto_oscuro": "#0F172A",
            "texto_mutado": "#64748B",
            "exito": "#059669",
            "advertencia": "#D97706",
            "peligro": "#DC2626"
        }
        self.root.configure(bg=self.colores["fondo"])
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=("Helvetica", 11), background=self.colores["fondo"])
        style.configure("Treeview", font=("Helvetica", 10), rowheight=28)
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"), background="#F1F5F9")

    def _construir_interfaz(self):
        contenedor = tk.Frame(self.root, bg=self.colores["fondo"], padx=18, pady=14)
        contenedor.pack(fill=tk.BOTH, expand=True)

        # Encabezado
        frame_cab = tk.Frame(contenedor, bg=self.colores["fondo"])
        frame_cab.pack(fill=tk.X, pady=(0, 10))
        tk.Label(frame_cab, text="COMPARADOR ACADÉMICO 2.0", font=("Helvetica", 20, "bold"), bg=self.colores["fondo"]).pack(anchor="w")
        tk.Label(frame_cab, text="Comparación exclusiva entre títulos de temas contenidos en el Plan de Estudios de documentos Word (.docx).", font=("Helvetica", 11), bg=self.colores["fondo"], fg=self.colores["texto_mutado"]).pack(anchor="w", pady=(2, 0))

        # Paneles de carga
        frame_carga = tk.Frame(contenedor, bg=self.colores["fondo"])
        frame_carga.pack(fill=tk.X, pady=(0, 10))
        frame_carga.columnconfigure(0, weight=1, uniform="docs")
        frame_carga.columnconfigure(1, weight=1, uniform="docs")

        self.panel_doc1 = self._crear_panel_doc(frame_carga, 0, 1, "DOCUMENTO 1: PROGRAMA DE ORIGEN")
        self.panel_doc2 = self._crear_panel_doc(frame_carga, 1, 2, "DOCUMENTO 2: PROGRAMA DE COMPARACIÓN")

        # Barra de acción
        frame_acc = tk.Frame(contenedor, bg=self.colores["panel"], relief="solid", bd=1)
        frame_acc.pack(fill=tk.X, pady=(0, 10), ipady=6, padx=2)

        self.btn_analizar = tk.Button(
            frame_acc, text="🔍  ANALIZAR DOCUMENTOS", font=("Helvetica", 13, "bold"),
            bg="#2563EB", fg="white", activebackground="#1D4ED8", relief="flat",
            cursor="hand2", padx=20, pady=8, command=self.solicitar_analisis
        )
        self.btn_analizar.pack(side=tk.LEFT, padx=14)

        self.lbl_estado = tk.Label(frame_acc, text="Seleccione ambos documentos Word (.docx) para iniciar el análisis.", font=("Helvetica", 11), bg=self.colores["panel"], fg=self.colores["texto_mutado"])
        self.lbl_estado.pack(side=tk.LEFT, padx=10)

        # Tarjetas de resumen
        frame_res = tk.Frame(contenedor, bg=self.colores["fondo"])
        frame_res.pack(fill=tk.X, pady=(0, 8))

        self.kpis = {}
        items_kpi = [
            ("temas_doc1", "Temas Doc 1", "0", self.colores["texto_oscuro"]),
            ("temas_doc2", "Temas Doc 2", "0", self.colores["texto_oscuro"]),
            ("literales", "Coincidencia Literal", "0", self.colores["exito"]),
            ("altas", "Similitud Alta", "0", self.colores["primario"]),
            ("tematicas", "Relación Temática", "0", self.colores["advertencia"]),
            ("sin_match", "Sin Coincidencia", "0", self.colores["peligro"])
        ]
        for i, (key, label, val_def, color) in enumerate(items_kpi):
            frame_res.columnconfigure(i, weight=1, uniform="kpi")
            c = tk.Frame(frame_res, bg=self.colores["panel"], relief="solid", bd=1, padx=6, pady=4)
            c.grid(row=0, column=i, padx=3, sticky="nsew")
            tk.Label(c, text=label, font=("Helvetica", 9), bg=self.colores["panel"], fg=self.colores["texto_mutado"]).pack()
            val_lbl = tk.Label(c, text=val_def, font=("Helvetica", 14, "bold"), bg=self.colores["panel"], fg=color)
            val_lbl.pack()
            self.kpis[key] = val_lbl

        # Barra de filtros
        frame_filtro = tk.Frame(contenedor, bg=self.colores["fondo"])
        frame_filtro.pack(fill=tk.X, pady=(0, 6))
        tk.Label(frame_filtro, text="Filtrar por categoría:", font=("Helvetica", 10, "bold"), bg=self.colores["fondo"]).pack(side=tk.LEFT, padx=(2, 6))

        self.combo_filtro = ttk.Combobox(frame_filtro, values=["Todas las categorías", "Coincidencia Literal", "Similitud Alta", "Relación Temática", "Sin Relación"], state="readonly", width=22)
        self.combo_filtro.current(0)
        self.combo_filtro.pack(side=tk.LEFT, padx=(0, 15))
        self.combo_filtro.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        tk.Label(frame_filtro, text="Buscar tema:", font=("Helvetica", 10, "bold"), bg=self.colores["fondo"]).pack(side=tk.LEFT, padx=(0, 6))
        self.entry_buscar = tk.Entry(frame_filtro, font=("Helvetica", 10), width=26)
        self.entry_buscar.pack(side=tk.LEFT, padx=(0, 6))
        self.entry_buscar.bind("<KeyRelease>", lambda e: self._aplicar_filtro())

        tk.Button(frame_filtro, text="Limpiar filtro", font=("Helvetica", 9), relief="flat", padx=8, pady=2, command=self._limpiar_filtro).pack(side=tk.LEFT)

        # Tabla de resultados
        frame_tab = tk.Frame(contenedor, bg=self.colores["panel"], relief="solid", bd=1)
        frame_tab.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        columnas = ("doc1_mod", "doc1_tema", "doc2_mod", "doc2_tema", "similitud", "categoria", "recomendacion")
        self.tabla = ttk.Treeview(frame_tab, columns=columnas, show="headings", selectmode="browse")

        self.tabla.heading("doc1_mod", text="Módulo Doc 1", anchor="w")
        self.tabla.heading("doc1_tema", text="Tema Doc 1", anchor="w")
        self.tabla.heading("doc2_mod", text="Módulo Doc 2", anchor="w")
        self.tabla.heading("doc2_tema", text="Tema Doc 2", anchor="w")
        self.tabla.heading("similitud", text="Similitud", anchor="center")
        self.tabla.heading("categoria", text="Clasificación", anchor="center")
        self.tabla.heading("recomendacion", text="Recomendación", anchor="w")

        self.tabla.column("doc1_mod", width=120)
        self.tabla.column("doc1_tema", width=220)
        self.tabla.column("doc2_mod", width=120)
        self.tabla.column("doc2_tema", width=220)
        self.tabla.column("similitud", width=80, anchor="center")
        self.tabla.column("categoria", width=160, anchor="center")
        self.tabla.column("recomendacion", width=240)

        scroll_y = ttk.Scrollbar(frame_tab, orient=tk.VERTICAL, command=self.tabla.yview)
        scroll_x = ttk.Scrollbar(frame_tab, orient=tk.HORIZONTAL, command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame_tab.grid_rowconfigure(0, weight=1)
        frame_tab.grid_columnconfigure(0, weight=1)

        self.tabla.bind("<<TreeviewSelect>>", self._al_seleccionar_fila)

        # Panel detalle
        frame_det = tk.Frame(contenedor, bg="#F1F5F9", relief="solid", bd=1, padx=12, pady=6)
        frame_det.pack(fill=tk.X, pady=(0, 8))
        self.lbl_detalle = tk.Label(frame_det, text="Detalle: Seleccione una fila para ver la justificación de similitud y términos compartidos.", font=("Helvetica", 10, "italic"), bg="#F1F5F9", fg=self.colores["texto_oscuro"], wraplength=1100, justify="left")
        self.lbl_detalle.pack(anchor="w")

        # Botones inferiores
        frame_inf = tk.Frame(contenedor, bg=self.colores["fondo"])
        frame_inf.pack(fill=tk.X)
        tk.Button(frame_inf, text="🔄  NUEVO ANÁLISIS", font=("Helvetica", 11, "bold"), bg="#64748B", fg="white", relief="flat", cursor="hand2", padx=16, pady=7, command=self.reiniciar_analisis).pack(side=tk.LEFT)
        tk.Button(frame_inf, text="📥  EXPORTAR RESULTADOS (CSV / TXT)", font=("Helvetica", 11, "bold"), bg="#059669", fg="white", relief="flat", cursor="hand2", padx=16, pady=7, command=self.solicitar_exportacion).pack(side=tk.RIGHT)

    def _crear_panel_doc(self, parent, col, num, titulo):
        panel = tk.Frame(parent, bg=self.colores["panel"], relief="solid", bd=1, padx=12, pady=10)
        panel.grid(row=0, column=col, padx=4, sticky="nsew")
        tk.Label(panel, text=titulo, font=("Helvetica", 11, "bold"), bg=self.colores["panel"]).pack(anchor="w")
        lbl_arch = tk.Label(panel, text="Archivo: Ningún archivo seleccionado", font=("Helvetica", 10), bg=self.colores["panel"])
        lbl_arch.pack(anchor="w", pady=(2, 0))
        lbl_est = tk.Label(panel, text="Estado: Pendiente de carga", font=("Helvetica", 9, "bold"), bg=self.colores["panel"], fg=self.colores["advertencia"])
        lbl_est.pack(anchor="w", pady=(2, 6))
        btn = tk.Button(panel, text=f"Seleccionar Documento {num}...", font=("Helvetica", 10), bg="#E2E8F0", relief="flat", cursor="hand2", padx=10, pady=4, command=lambda: self.seleccionar_documento(num))
        btn.pack(anchor="w")

        panel.lbl_archivo = lbl_arch
        panel.lbl_estado = lbl_est
        panel.btn = btn
        return panel

    def _iniciar_ciclo_control(self):
        self._procesar_cola_control()

    def _procesar_cola_control(self):
        """
        ESTRUCTURA OBLIGATORIA: while
        Procesa de forma continua las solicitudes de análisis y eventos sin congelar la GUI.
        """
        while self.app_activa and not self.cola_solicitudes.empty():
            accion, datos = self.cola_solicitudes.get_nowait()
            if accion == "EJECUTAR_ANALISIS":
                self._ejecutar_analisis_en_hilo(datos)
            elif accion == "FINALIZAR_ANALISIS":
                self._mostrar_resultados_en_ui(datos)
            elif accion == "ERROR_ANALISIS":
                self._manejar_error_analisis(datos)
            elif accion == "REINICIAR":
                self.reiniciar_analisis()

        if self.app_activa:
            self.root.after(100, self._procesar_cola_control)

    def seleccionar_documento(self, num_doc):
        """Abre el explorador de archivos para seleccionar un documento Word (.docx)."""
        formatos = [("Documentos Word (*.docx)", "*.docx"), ("Todos los archivos (*.*)", "*.*")]
        ruta = filedialog.askopenfilename(title=f"Seleccionar Documento {num_doc} (.docx)", filetypes=formatos)
        if not ruta:
            return

        panel = self.panel_doc1 if num_doc == 1 else self.panel_doc2
        nombre = os.path.basename(ruta)

        try:
            lineas = cargar_documento(ruta)
            temas = extraer_temas(lineas, nombre_doc=f"Documento {num_doc}")

            if not temas:
                panel.lbl_archivo.config(text=f"Archivo: {nombre}")
                panel.lbl_estado.config(text="Aviso: No se identificaron títulos de temas.", fg=self.colores["peligro"])
                messagebox.showwarning("Estructura no identificada", f"No se identificaron temas válidos en '{nombre}'.")
                return

            if num_doc == 1:
                self.doc1_ruta = ruta
                self.doc1_nombre = nombre
                self.doc1_datos = temas
                self.kpis["temas_doc1"].config(text=str(len(temas)))
            else:
                self.doc2_ruta = ruta
                self.doc2_nombre = nombre
                self.doc2_datos = temas
                self.kpis["temas_doc2"].config(text=str(len(temas)))

            panel.lbl_archivo.config(text=f"Archivo: {nombre}")
            panel.lbl_estado.config(text=f"Cargado con éxito: {len(temas)} tema(s) identificados", fg=self.colores["exito"])
            panel.btn.config(text=f"Reemplazar Documento {num_doc}...")

            if self.doc1_datos and self.doc2_datos:
                self.lbl_estado.config(text=f"Listos para analizar: Doc 1 ({len(self.doc1_datos)} temas) y Doc 2 ({len(self.doc2_datos)} temas).", fg=self.colores["primario"])
            else:
                self.lbl_estado.config(text=f"Documento {num_doc} cargado. Seleccione el segundo documento.", fg=self.colores["texto_mutado"])

        except Exception as e:
            panel.lbl_estado.config(text=f"Error al leer: {str(e)[:45]}...", fg=self.colores["peligro"])
            messagebox.showerror("Error en Documento", f"Ocurrió un error al procesar el archivo:\n\n{str(e)}")

    def solicitar_analisis(self):
        """Valida que ambos documentos estén listos y encola la solicitud."""
        if not self.doc1_datos or not self.doc2_datos:
            messagebox.showinfo("Documentos incompletos", "Debe cargar ambos documentos Word (.docx) antes de analizar.")
            return

        if self.analisis_en_curso:
            messagebox.showinfo("Procesando", "Ya hay un análisis en curso.")
            return

        self.analisis_en_curso = True
        self.btn_analizar.config(state=tk.DISABLED, text="⏳  ANALIZANDO TEMAS...")
        self.lbl_estado.config(text="Comparando títulos de temas entre ambos documentos...", fg=self.colores["primario"])

        solicitud = {"temas_doc1": self.doc1_datos, "temas_doc2": self.doc2_datos}
        self.cola_solicitudes.put(("EJECUTAR_ANALISIS", solicitud))

    def _ejecutar_analisis_en_hilo(self, datos):
        def tarea():
            try:
                comparaciones, resumen, sin_match = comparar_temas(datos["temas_doc1"], datos["temas_doc2"])
                self.cola_solicitudes.put(("FINALIZAR_ANALISIS", (comparaciones, resumen, sin_match)))
            except Exception as err:
                self.cola_solicitudes.put(("ERROR_ANALISIS", str(err)))

        threading.Thread(target=tarea, daemon=True).start()

    def _mostrar_resultados_en_ui(self, datos):
        self.analisis_en_curso = False
        self.btn_analizar.config(state=tk.NORMAL, text="🔍  ANALIZAR DOCUMENTOS")

        comparaciones, resumen, sin_match = datos
        self.resultados_actuales = comparaciones
        self.resumen_actual = resumen
        self.temas_sin_match = sin_match

        self.kpis["literales"].config(text=str(resumen["coincidencias_literales"]))
        self.kpis["altas"].config(text=str(resumen["similitudes_altas"]))
        self.kpis["tematicas"].config(text=str(resumen["relaciones_tematicas"]))
        self.kpis["sin_match"].config(text=str(resumen["sin_coincidencias"]))

        self.lbl_estado.config(text=f"Análisis completado: {len(comparaciones)} comparaciones evaluadas.", fg=self.colores["exito"])
        self._aplicar_filtro()

        messagebox.showinfo(
            "Análisis Completado",
            f"Se ha completado el análisis comparativo con éxito.\n\n"
            f"• Temas en Documento 1: {resumen['total_temas_doc1']}\n"
            f"• Temas en Documento 2: {resumen['total_temas_doc2']}\n"
            f"• Coincidencias literales: {resumen['coincidencias_literales']}\n"
            f"• Similitudes altas: {resumen['similitudes_altas']}\n"
            f"• Relaciones temáticas: {resumen['relaciones_tematicas']}\n"
            f"• Temas sin coincidencias: {resumen['sin_coincidencias']}"
        )

    def _manejar_error_analisis(self, mensaje):
        self.analisis_en_curso = False
        self.btn_analizar.config(state=tk.NORMAL, text="🔍  ANALIZAR DOCUMENTOS")
        self.lbl_estado.config(text=f"Error: {mensaje}", fg=self.colores["peligro"])
        messagebox.showerror("Error en Análisis", str(mensaje))

    def _aplicar_filtro(self):
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        categoria_sel = self.combo_filtro.get()
        busqueda = self.entry_buscar.get().lower().strip()

        self.tabla.tag_configure("tag_literal", foreground="#047857", background="#ECFDF5")
        self.tabla.tag_configure("tag_alta", foreground="#1D4ED8", background="#EFF6FF")
        self.tabla.tag_configure("tag_tematica", foreground="#B45309", background="#FFFBEB")
        self.tabla.tag_configure("tag_sin_relacion", foreground="#64748B", background="#F8FAFC")

        for r in self.resultados_actuales:
            if categoria_sel == "Coincidencia Literal" and r["categoria"] != "COINCIDENCIA LITERAL":
                continue
            elif categoria_sel == "Similitud Alta" and r["categoria"] != "SIMILITUD ALTA":
                continue
            elif categoria_sel == "Relación Temática" and r["categoria"] != "RELACIÓN TEMÁTICA O PARALELA":
                continue
            elif categoria_sel == "Sin Relación" and r["categoria"] != "SIN RELACIÓN IDENTIFICADA":
                continue

            if busqueda:
                t1 = r["doc1_tema_titulo"].lower()
                t2 = r["doc2_tema_titulo"].lower()
                if busqueda not in t1 and busqueda not in t2:
                    continue

            tag = "tag_literal" if r["nivel"] == 1 else "tag_alta" if r["nivel"] == 2 else "tag_tematica" if r["nivel"] == 3 else "tag_sin_relacion"
            self.tabla.insert("", tk.END, values=(r["doc1_modulo"], r["doc1_tema_titulo"], r["doc2_modulo"], r["doc2_tema_titulo"], f"{r['similitud']}%", r["categoria"], r["recomendacion"]), tags=(tag,))

    def _limpiar_filtro(self):
        self.combo_filtro.current(0)
        self.entry_buscar.delete(0, tk.END)
        self._aplicar_filtro()

    def _al_seleccionar_fila(self, event):
        sel = self.tabla.selection()
        if not sel:
            return
        item = self.tabla.item(sel[0])
        val = item["values"]
        if not val:
            return

        t1, t2 = val[1], val[3]
        explicacion = "Términos compartidos identificados."
        for r in self.resultados_actuales:
            if r["doc1_tema_titulo"] == t1 and r["doc2_tema_titulo"] == t2:
                explicacion = r["explicacion"]
                break

        self.lbl_detalle.config(text=f"📌 Doc 1: {t1}   ⟷   📌 Doc 2: {t2}\n• Similitud: {val[4]} ({val[5]})   |   • Recomendación: {val[6]}\n• Justificación: {explicacion}", font=("Helvetica", 10, "normal"))

    def reiniciar_analisis(self):
        """
        Función obligatoria para reiniciar el estado de la aplicación.
        
        ESTRUCTURA OBLIGATORIA: while
        """
        # ======================================================================
        # CICLOS 'WHILE' OBLIGATORIOS PARA VACIAR RESULTADOS Y TABLA
        # ======================================================================
        while len(self.resultados_actuales) > 0:
            self.resultados_actuales.pop()

        while len(self.temas_sin_match) > 0:
            self.temas_sin_match.pop()

        self.resumen_actual.clear()

        items = self.tabla.get_children()
        idx = 0
        while idx < len(items):
            self.tabla.delete(items[idx])
            idx += 1

        self.kpis["literales"].config(text="0")
        self.kpis["altas"].config(text="0")
        self.kpis["tematicas"].config(text="0")
        self.kpis["sin_match"].config(text="0")
        self.entry_buscar.delete(0, tk.END)
        self.combo_filtro.current(0)
        self.lbl_detalle.config(text="Detalle: Seleccione una fila para ver la justificación de similitud y términos compartidos.", font=("Helvetica", 10, "italic"))

        if self.doc1_datos and self.doc2_datos:
            self.lbl_estado.config(text="Resultados reiniciados. Puede volver a pulsar 'ANALIZAR DOCUMENTOS'.", fg=self.colores["primario"])
        else:
            self.lbl_estado.config(text="Análisis reiniciado. Cargue los documentos para iniciar.", fg=self.colores["texto_mutado"])

    def solicitar_exportacion(self):
        """Exporta los resultados a formato CSV o TXT."""
        if not self.resultados_actuales:
            messagebox.showinfo("Sin resultados", "No hay resultados para exportar. Ejecute primero el análisis.")
            return

        ruta = filedialog.asksaveasfilename(
            title="Guardar Informe Comparativo",
            defaultextension=".csv",
            filetypes=[("Archivo CSV (*.csv)", "*.csv"), ("Archivo de Texto (*.txt)", "*.txt")]
        )
        if not ruta:
            return

        try:
            formato = "txt" if ruta.lower().endswith(".txt") else "csv"
            generar_informe(self.resultados_actuales, self.resumen_actual, {"nombre": self.doc1_nombre}, {"nombre": self.doc2_nombre}, ruta, formato=formato)
            messagebox.showinfo("Exportación Exitosa", f"Informe guardado correctamente en:\n\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el informe:\n\n{str(e)}")

    def cerrar_aplicacion(self):
        self.app_activa = False
        self.root.destroy()


# ==============================================================================
# 5. PUNTO DE ENTRADA PRINCIPAL
# ==============================================================================

def main():
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.25)
    except Exception:
        pass
    app = ComparadorAcademicoGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.cerrar_aplicacion)
    root.mainloop()


if __name__ == "__main__":
    main()
