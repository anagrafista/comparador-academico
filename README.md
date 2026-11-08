# COMPARADOR ACADÉMICO 2.0
### Aplicación de escritorio en Python para análisis y posible equivalencia de programas académicos

---

## 1. DESCRIPCIÓN GENERAL

**COMPARADOR ACADÉMICO 2.0** es una aplicación de escritorio desarrollada en Python con interfaz gráfica Tkinter, diseñada para contrastar de manera automatizada los **títulos de temas** contenidos en documentos académicos en formato **Microsoft Word (.docx)** (dosieres, índices de máster, programas de estudio o planes curriculares).

El objetivo es identificar posibles coincidencias y relaciones temáticas entre dos programas formativos para evaluar posibles equivalencias académicas o sustituciones de asignaturas.

> **Regla de oro del sistema:** 
> 1. En documentos extensos (ej. dosieres de 44 o más páginas), la aplicación localiza y procesa **exclusivamente las páginas del Plan de Estudios** (descartando índices, presentaciones, rankings, metodologías, claustros, admisiones y becas).
> 2. Dentro del Plan de Estudios, extrae:
>    - **Módulos:** numeración `1.`, `2.`, `Módulo 1`, etc.
>    - **Temas principales:** viñetas numéricas de dos niveles `1.1`, `1.2`, `2.1`, etc.
>    - **Subtemas (`1.1.1`, `1.1.2`, `X.Y.Z`):** quedan **estrictamente excluidos**, al igual que descripciones, objetivos y competencias.

---

## 2. CARACTERÍSTICAS Y MODALIDADES DE USO

- **Lenguaje:** Python 3 (ejecución 100% nativa).
- **Modalidad 1 (Web en la Nube con Streamlit):** Funciona como aplicación web interactiva desplegable gratuitamente en **Streamlit Community Cloud** para acceder desde cualquier navegador o teléfono sin instalar nada.
- **Modalidad 2 (Escritorio Local con Tkinter):** Funciona con interfaz nativa de ventanas en macOS haciendo doble clic en `Iniciar.command`.
- **Formato Exclusivo:** **Microsoft Word (.docx)**. Lectura directa, veloz y sin problemas de maquetación.
- **Exportación:** Informes completos en **CSV** (compatible con Excel) y **TXT** estructurado.

---

## 3. INSTALACIÓN Y EJECUCIÓN

### Paso 1: Instalar dependencias
Abra la Terminal en la carpeta del proyecto y ejecute:
```bash
pip3 install -r requirements.txt
```

### Paso 2: Ejecutar la aplicación

**A. Ejecutar la versión Web (Streamlit):**
```bash
streamlit run streamlit_app.py
```
*(Se abrirá automáticamente en su navegador web).*

**B. Desplegar en Streamlit Community Cloud (Gratis en la Nube):**
1. Suba este repositorio a su cuenta de **GitHub**.
2. Vaya a [share.streamlit.io](https://share.streamlit.io) e inicie sesión con su GitHub.
3. Pulse **New app**, seleccione el repositorio y deje `streamlit_app.py` como archivo principal.
4. Pulse **Deploy** y obtendrá un enlace público web permanente.

**C. Ejecutar la versión de Escritorio (Mac Tkinter):**
- Haga doble clic en el archivo **`Iniciar.command`** en Finder, o ejecute:
```bash
python3 main.py
```

---

## 4. EXPERIENCIA DE USO PASO A PASO

1. **Abrir la aplicación:** Se desplegará la ventana principal con título, descripción y los dos paneles de carga.
2. **Cargar Documento 1 (Origen):** Seleccione el primer dosier académico (.docx). El sistema informará el total de temas identificados.
3. **Cargar Documento 2 (Comparación):** Seleccione el segundo dosier o plan curricular de contraste (.docx).
4. **Analizar Documentos:** Pulse el botón azul **"🔍 ANALIZAR DOCUMENTOS"**. El sistema ejecutará la comparación sistemática sin congelar la interfaz.
5. **Examinar Resultados:**
   - Revise las **tarjetas resumen** (Temas identificados, Coincidencias literales, Similitudes altas, Relaciones temáticas, Sin coincidencia).
   - Use la **tabla interactiva** para ordenar y explorar cada comparación.
   - Haga clic en cualquier fila para ver en el **panel de detalle** la explicación textual exacta y los términos compartidos.
   - Utilice el **filtro por categoría** o la **caja de búsqueda** para localizar temas específicos.
6. **Exportar Resultados:** Pulse el botón verde **"📥 EXPORTAR RESULTADOS"** para guardar el informe en CSV o TXT en la ubicación que elija.
7. **Nuevo Análisis:** Pulse el botón gris **"🔄 NUEVO ANÁLISIS"** para reiniciar la tabla y métricas y realizar una nueva comparación sin cerrar la aplicación.

---

## 5. ALGORITMOS DE COMPARACIÓN Y CLASIFICACIÓN

El sistema compara cada título del Documento 1 contra cada título del Documento 2 calculando un porcentaje matemático real (0% a 100%) mediante:
1. **Normalización lingüística:** Conversión a minúsculas, descomposición Unicode NFD para supresión de tildes y diacríticos, y eliminación de puntuación.
2. **Filtrado de palabras vacías (stopwords):** Aislamiento de sustantivos y términos clave significativos en español.
3. **Análisis morfológico y de raíces:** Detección de familias de palabras y afijos comunes (ej. *seguridad* en *ciberseguridad*, *empresa* en *empresarial*, *automóvil* en *automotriz*).
4. **Comparación de secuencias (`difflib`):** Medición de similitud estructural de cadenas.
5. **Clasificación en 4 Categorías Oficiales:**

| Categoría | Rango de Similitud | Recomendación Orientativa | Ejemplo |
|---|---|---|---|
| **CATEGORÍA 1: Coincidencia Literal** | $\ge 95\%$ | *Posible sustitución (Coincidencia exacta de títulos)* | *Marketing digital* ⟷ *Marketing digital* |
| **CATEGORÍA 2: Similitud Alta** | $72\% - 94.9\%$ | *Posible sustitución (Alta coincidencia textual)* | *Gestión de proyectos empresariales* ⟷ *Administración de proyectos de empresa* |
| **CATEGORÍA 3: Relación Temática o Paralela** | $42\% - 71.9\%$ | *Similitud que requiere revisión (Relación temática, examinar contenido)* | *Historia del automóvil* ⟷ *Evolución de la industria automotriz* |
| **CATEGORÍA 4: Sin Relación Identificada** | $< 42\%$ | *Sin relación identificada (Coincidencia insuficiente)* | *Derecho fiscal* ⟷ *Robótica espacial* |

---

## 6. VERIFICACIÓN ACADÉMICA DEL CÓDIGO (REQUISITO OBLIGATORIO)

A continuación se detalla el cumplimiento estricto y funcional de las cinco estructuras de control en el código fuente:

| Estructura | Función dentro del proyecto | Archivo y Función |
|:---|:---|:---|
| **`if`** | Validación de archivos seleccionados, control de formatos (.docx) y evaluación de umbrales. | `core/extractor.py` (`cargar_documento`, `leer_docx`), `gui/app.py` (`solicitar_analisis`). |
| **`elif`** | Clasificación sistemática de las cuatro categorías de resultados y asignación de recomendaciones según porcentajes. | `core/comparator.py` (`clasificar_resultado`). |
| **`else`** | Gestión de casos residuales (sin coincidencia suficiente, formatos no válidos o fallos de lectura). | `core/comparator.py` (`clasificar_resultado`), `core/extractor.py` (`cargar_documento`). |
| **`for`** | Recorrido sistemático y anidado de todos los títulos del Documento 1 contrastados contra todos los del Documento 2. | `core/comparator.py` (`comparar_temas`). |
| **`while`** | Ciclo de control no bloqueante para procesar solicitudes y repetición de análisis, así como reinicio de resultados sin cerrar la app. | `gui/app.py` (`_procesar_cola_control`, `reiniciar_analisis`). |

### Fragmentos Reales del Código Fuente

#### 1. Estructura `if / elif / else` (Clasificación de Resultados)
Ubicada en `core/comparator.py` en la función `clasificar_resultado()`:
```python
def clasificar_resultado(porcentaje):
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
```

#### 2. Estructura `for` (Ciclo Sistemático de Comparación)
Ubicada en `core/comparator.py` en la función `comparar_temas()`:
```python
def comparar_temas(temas_doc1, temas_doc2):
    comparaciones = []
    # Recorrido sistemático mediante ciclos for anidados
    for idx1, t1 in enumerate(temas_doc1):
        mejor_similitud_t1 = -1.0
        mejor_match_t1 = None

        for idx2, t2 in enumerate(temas_doc2):
            similitud, explicacion = calcular_similitud(
                t1["tema_titulo"],
                t2["tema_titulo"]
            )
            categoria, recomendacion, nivel = clasificar_resultado(similitud)
            ...
```

#### 3. Estructura `while` (Control de Ciclo de Repetición y Reinicio)
Ubicada en `gui/app.py` en las funciones `_procesar_cola_control()` y `reiniciar_analisis()`:
```python
def _procesar_cola_control(self):
    # Ciclo while de control de eventos compatible con el mainloop de Tkinter
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
```
Y en `reiniciar_analisis()`:
```python
def reiniciar_analisis(self):
    # Ciclo while para vaciar la lista de resultados y resetear el estado
    while len(self.resultados_actuales) > 0:
        self.resultados_actuales.pop()

    while len(self.temas_sin_match) > 0:
        self.temas_sin_match.pop()

    items_tabla = self.tabla_resultados.get_children()
    idx = 0
    while idx < len(items_tabla):
        self.tabla_resultados.delete(items_tabla[idx])
        idx += 1
```

---

## 7. ARCHIVOS DE PRUEBA INCLUIDOS

En la carpeta `test_data/` se incluyen dos programas académicos en Word (.docx) listos para contrastar:
- `programa_master_direccion_doc1.docx` (Documento 1 en Word)
- `programa_master_gestion_doc2.docx` (Documento 2 en Word)

---

## 8. EJECUCIÓN DE PRUEBAS AUTOMATIZADAS

Para validar la integridad de todas las funciones y algoritmos, ejecute:
```bash
python3 -m unittest discover -s tests -v
```
Resultado: **13/13 pruebas unitarias superadas exitosamente.**
