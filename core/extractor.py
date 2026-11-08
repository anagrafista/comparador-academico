"""
Módulo de lectura y extracción exclusiva de temas académicos.
Diseñado para operar con documentos en formato Microsoft Word (.docx).
Aísla la sección del Plan de Estudios, extrae módulos y temas principales (viñetas 1. y 1.1)
y descarta subtemas (1.1.1), descripciones, objetivos y contenidos adicionales.
"""

import os
import re

try:
    import docx
except ImportError:
    docx = None


def leer_docx(fuente_archivo):
    """
    Lee y extrae las líneas de texto de un archivo Microsoft Word (.docx)
    utilizando python-docx, procesando párrafos y tablas.
    
    Acepta tanto rutas en disco (str) como flujos en memoria (io.BytesIO / UploadedFile de Streamlit).
    """
    if docx is None:
        raise ImportError("La biblioteca python-docx no está instalada.")

    # Si es una ruta en disco de tipo string
    if isinstance(fuente_archivo, str):
        if not os.path.exists(fuente_archivo):
            raise FileNotFoundError(f"No se encontró el archivo: {fuente_archivo}")
        target = fuente_archivo
    elif hasattr(fuente_archivo, "read"):
        # Flujo en memoria (Streamlit UploadedFile o BytesIO)
        import io
        contenido_bytes = fuente_archivo.read()
        if hasattr(fuente_archivo, "seek"):
            fuente_archivo.seek(0)
        target = io.BytesIO(contenido_bytes)
    else:
        target = fuente_archivo

    try:
        doc = docx.Document(target)
        lineas = []

        # Extraer texto de párrafos
        for p in doc.paragraphs:
            texto = p.text.strip()
            if texto:
                lineas.append(texto)

        # Extraer texto de tablas si el plan de estudios está estructurado en celdas
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


def cargar_documento(fuente_archivo):
    """
    Carga y valida un documento académico en formato Word (.docx).
    Soporta rutas de archivo locales y archivos subidos en Streamlit.
    """
    if not fuente_archivo:
        raise ValueError("Archivo no especificado o inválido.")

    # Caso 1: Archivo subido en Streamlit o buffer en memoria
    if hasattr(fuente_archivo, "name"):
        nombre = fuente_archivo.name.lower()
        if not nombre.endswith(".docx"):
            raise ValueError(f"Formato no admitido en '{fuente_archivo.name}'. Se requiere formato Word (.docx).")
        return leer_docx(fuente_archivo)

    # Caso 2: Ruta en disco (string)
    if isinstance(fuente_archivo, str):
        extension = os.path.splitext(fuente_archivo)[1].lower()
        if extension != ".docx":
            raise ValueError(
                f"Formato no admitido: '{extension}'. La aplicación opera exclusivamente con documentos Word (.docx)."
            )

        if not os.path.isfile(fuente_archivo):
            raise FileNotFoundError(f"El archivo especificado no existe: {fuente_archivo}")

        return leer_docx(fuente_archivo)

    return leer_docx(fuente_archivo)


def _es_descarte_o_descripcion(linea):
    """
    Determina si una línea corresponde a descripciones, objetivos,
    competencias, bibliografía o subtemas profundos que no deben analizarse.
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

    # Párrafos explicativos extensos (un título académico rara vez supera 22 palabras o 140 caracteres)
    palabras = linea.split()
    if len(palabras) > 22 or len(linea) > 140:
        return True

    # Oraciones largas con punto final
    if linea.endswith(".") and len(palabras) >= 8:
        if re.search(r"\b(el alumno|los estudiantes|aprenderá|analizará|comprenderá|desarrollará)\b", linea_lower):
            return True

    # Subtemas profundos de tercer nivel (1.1.1, a), b))
    if re.match(r"^\s*(?:\d+\.\d+\.\d+|[a-zA-Z]\)|\([a-zA-Z]\))\s*", linea):
        return True

    return False


def _limpiar_titulo_tema(titulo):
    """
    Limpia el título de un tema retirando viñetas o prefijos numéricos iniciales.
    """
    titulo = titulo.strip()
    titulo = re.sub(r"^(?:Tema|TEMA|Unidad|UNIDAD|Capítulo|CAPÍTULO)\s+\d+[\s\.\:\-]+", "", titulo, flags=re.IGNORECASE)
    titulo = re.sub(r"^\d+(?:\.\d+)?[\s\.\:\-]+", "", titulo)
    titulo = re.sub(r"^[\-\•\*\–\—\>]\s*", "", titulo)
    titulo = re.sub(r"\.+$", "", titulo)
    return titulo.strip()


def delimitar_seccion_plan_estudios(lineas):
    """
    Localiza y aísla exclusivamente la sección del 'Plan de estudios'
    dentro de documentos extensos (ej. dosieres de 44 o más páginas),
    descartando páginas previas (introducción, bienvenida, rankings)
    y páginas posteriores (metodología, claustro, admisión, becas, etc.).
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
        # Ignorar líneas del índice con puntos sucesivos hacia números de página
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
    Extrae exclusivamente los títulos de los temas contenidos en el documento.
    
    Aísla la sección del Plan de estudios y extrae:
    - Módulo: numeración 1., 2., Módulo 1, etc.
    - Temas principales: numeración 1.1, 1.2, 2.1, etc.
    - SUBTEMAS (1.1.1, 1.1.2, X.Y.Z): EXCLUIDOS POR COMPLETO.
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

        # A. Descartar descripciones, objetivos, competencias
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

        # D. DETECCIÓN DE MÓDULO NUMÉRICO: '1.', '2.', '1. Nombre', 'Módulo 1: Nombre'
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
