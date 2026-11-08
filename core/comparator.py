"""
Módulo de comparación de temas académicos.
Implementa normalización, cálculo de similitud textual,
clasificación de coincidencias y ciclo sistemático de comparación.

Estructuras de programación obligatorias implementadas:
- if, elif, else (evaluación y clasificación)
- for (ciclos de comparación entre documentos)
"""

import re
import unicodedata
from difflib import SequenceMatcher

# ==============================================================================
# UMBRALES DE CLASIFICACIÓN CONFIGURABLES
# ==============================================================================
UMBRAL_LITERAL = 95.0      # Coincidencia casi idéntica o literal
UMBRAL_ALTA = 72.0         # Similitud textual alta
UMBRAL_TEMATICA = 42.0     # Relación temática o paralela

# Palabras vacías en español (stopwords) para aislar términos clave académicos
PALABRAS_PARADA_ES = {
    "de", "del", "la", "las", "el", "los", "un", "una", "unos", "unas",
    "y", "e", "o", "u", "en", "con", "por", "para", "sobre", "a", "al",
    "ante", "bajo", "cabe", "contra", "desde", "hacia", "hasta", "mediante",
    "sin", "so", "tras", "durante", "segun", "que", "su", "sus", "se", "como",
    "lo", "les", "nos", "te", "mi", "mis", "tu", "tus", "muy", "mas", "pero"
}


def normalizar_texto(texto):
    """
    Normaliza una cadena de texto para facilitar comparaciones precisas:
    - Conversión a minúsculas.
    - Eliminación de acentos y diacríticos.
    - Eliminación de signos de puntuación y caracteres especiales.
    - Reducción de espacios múltiples.
    """
    if not texto or not isinstance(texto, str):
        return ""

    # 1. Convertir a minúsculas
    texto = texto.lower().strip()

    # 2. Descomponer caracteres Unicode y eliminar tildes (NFD)
    texto_nfd = unicodedata.normalize("NFD", texto)
    texto_sin_tildes = "".join(c for c in texto_nfd if unicodedata.category(c) != "Mn")

    # 3. Reemplazar caracteres de puntuación por espacios
    texto_limpio = re.sub(r"[^\w\s]", " ", texto_sin_tildes)

    # 4. Eliminar espacios múltiples
    texto_normalizado = " ".join(texto_limpio.split())

    return texto_normalizado


def extraer_palabras_clave(texto_normalizado):
    """
    Extrae un conjunto de palabras clave significativas (sin stopwords y con longitud >= 3).
    """
    palabras = texto_normalizado.split()
    clave = set()
    for p in palabras:
        if p not in PALABRAS_PARADA_ES and len(p) >= 3:
            clave.add(p)
    return clave


def clasificar_resultado(porcentaje):
    """
    Clasifica el porcentaje de similitud en una de las cuatro categorías oficiales
    y asigna la recomendación académica orientativa correspondiente.
    
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


def _obtener_raiz(palabra):
    """
    Obtiene la raíz léxica aproximada de una palabra en español mediante
    reglas morfológicas estándar (sin requerir bibliotecas externas ni IA).
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


def calcular_similitud(titulo1, titulo2):
    """
    Calcula el porcentaje de similitud real entre dos títulos académicos
    mediante algoritmos textuales combinados (difflib, análisis de raíces léxicas
    y coincidencias de términos clave).
    
    Retorna:
    - porcentaje (float entre 0.0 y 100.0)
    - explicacion (str detallando los términos compartidos o el tipo de coincidencia)
    """
    norm1 = normalizar_texto(titulo1)
    norm2 = normalizar_texto(titulo2)

    # Caso 1: Coincidencia literal exacta tras normalización
    if norm1 == norm2 and norm1:
        return 100.0, "Coincidencia literal exacta entre ambos títulos."

    if not norm1 or not norm2:
        return 0.0, "Uno de los títulos está vacío o carece de términos válidos."

    # 1. Similitud de secuencias con difflib
    ratio_difflib = SequenceMatcher(None, norm1, norm2).ratio() * 100.0

    # 2. Análisis de palabras clave significativas y raíces léxicas
    palabras1 = extraer_palabras_clave(norm1)
    palabras2 = extraer_palabras_clave(norm2)

    interseccion_exacta = palabras1.intersection(palabras2)
    terminos_relacionados = set(interseccion_exacta)

    # Buscar términos emparentados por raíz o inclusión (ej. seguridad/ciberseguridad, empresa/empresarial, automovil/automotriz)
    raices1 = {p: _obtener_raiz(p) for p in palabras1}
    raices2 = {p: _obtener_raiz(p) for p in palabras2}

    for p1, r1 in raices1.items():
        for p2, r2 in raices2.items():
            if p1 in terminos_relacionados and p2 in terminos_relacionados:
                continue
            # Coincidencia por raíz morfológica común
            if len(r1) >= 4 and len(r2) >= 4 and (r1 == r2 or r1.startswith(r2) or r2.startswith(r1)):
                terminos_relacionados.add(f"{p1} ~ {p2}")
            # Coincidencia por subcadena interna (ej. "seguridad" dentro de "ciberseguridad")
            elif (len(p1) >= 5 and p1 in p2) or (len(p2) >= 5 and p2 in p1):
                terminos_relacionados.add(f"{p1} ~ {p2}")

    total_claves = max(len(palabras1), len(palabras2)) if (palabras1 or palabras2) else 1
    coincidencias_ponderadas = len(interseccion_exacta) + (len(terminos_relacionados) - len(interseccion_exacta)) * 0.85

    proporcion_terminos = (coincidencias_ponderadas / total_claves) * 100.0

    # Construcción de la explicación de términos compartidos
    if terminos_relacionados:
        lista_terminos = sorted(list(terminos_relacionados))
        terminos_str = ", ".join(f"'{t}'" for t in lista_terminos[:5])
        explicacion = f"Comparten términos o raíces clave: {terminos_str}."
    else:
        explicacion = "No comparten términos clave principales."

    # Si todas las palabras clave son idénticas
    if palabras1 and palabras2 and palabras1 == palabras2:
        return 98.0, f"Mismos términos clave en distinta disposición: {terminos_str}."

    # Ponderación del porcentaje de similitud
    if terminos_relacionados:
        porcentaje_combinado = (ratio_difflib * 0.35) + (proporcion_terminos * 0.65)
        porcentaje_final = max(ratio_difflib, porcentaje_combinado)
    else:
        porcentaje_final = ratio_difflib * 0.65  # Penalizar si no hay conceptos afines

    porcentaje_final = min(100.0, max(0.0, porcentaje_final))
    return round(porcentaje_final, 1), explicacion


def comparar_temas(temas_doc1, temas_doc2):
    """
    Recorre sistemáticamente todos los títulos de temas del Documento 1
    y los contrasta con todos los temas del Documento 2.
    
    ESTRUCTURA OBLIGATORIA: Ciclos 'for' anidados.
    
    Retorna:
    - comparaciones: lista de todas las coincidencias calculadas, ordenadas de mayor a menor similitud.
    - mejores_coincidencias: para cada tema del Doc 1, su mejor relación en Doc 2.
    - resumen: conteos estadísticos de categorías encontradas.
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
    # CICLOS 'FOR' OBLIGATORIOS PARA COMPARACIÓN SISTEMÁTICA
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

            # Guardamos la comparación
            comparaciones.append(registro)

            # Rastrear la mejor coincidencia para este tema de Doc 1
            if similitud > mejor_similitud_t1:
                mejor_similitud_t1 = similitud
                mejor_match_t1 = registro

        # Evaluar si el tema del Doc 1 tuvo alguna relación válida (nivel 1, 2 o 3)
        if mejor_match_t1 and mejor_match_t1["nivel"] <= 3:
            temas_doc1_con_coincidencia.add(idx1)
            conteo_categorias[mejor_match_t1["categoria"]] += 1
        else:
            conteo_categorias["SIN RELACIÓN IDENTIFICADA"] += 1

    # Ordenar comparaciones de mayor a menor porcentaje de similitud
    comparaciones.sort(key=lambda x: x["similitud"], reverse=True)

    # Identificar temas de Documento 1 sin ninguna coincidencia relevante
    temas_sin_coincidencia = []
    for idx1, t1 in enumerate(temas_doc1):
        if idx1 not in temas_doc1_con_coincidencia:
            temas_sin_coincidencia.append(t1)

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
