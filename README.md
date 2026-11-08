# 🎓 Comparador Académico 2.0

Sistema web interactivo desarrollado en **Python** y **Streamlit** para la evaluación y contraste sistemático de planes de estudio contenidos en documentos Word (`.docx`).

---

## 🚀 Acceso en la Nube

La aplicación está desplegada y lista para usarse sin necesidad de instalación local:  
👉 **[Abrir Comparador Académico en Streamlit Cloud](https://comparador-academico.streamlit.app)**

---

## 🎯 Objetivo y Funcionalidades

- **Ámbito exclusivo:** Localiza y analiza automáticamente las páginas correspondientes al **Plan de Estudios**, omitiendo secciones de admisiones, claustrales, rankings y metodologías.
- **Extracción precisa de temas:** Extrae módulos (`1.`, `2.`) y temas principales (`1.1`, `1.2`), descartando estrictamente subtemas (`1.1.1`) y texto descriptivo.
- **Análisis algorítmico combinado:** Evalúa similitud textual mediante secuencias `difflib`, raíces morfológicas del español y filtrado de palabras clave.
- **Clasificación y recomendaciones oficiales:**
  - 🟢 **Coincidencia Literal ($\ge 95\%$):** *Posible sustitución (Coincidencia exacta de títulos)*
  - 🔵 **Similitud Alta ($72\% - 94.9\%$):** *Posible sustitución (Alta coincidencia textual)*
  - 🟡 **Relación Temática ($42\% - 71.9\%$):** *Similitud que requiere revisión (Relación temática)*
  - ⚪ **Sin Relación ($< 42\%$):** *Sin relación identificada (Coincidencia insuficiente)*
- **Exportación de informes:** Descarga inmediata de resultados detallados en formatos **CSV (compatible con Excel)** y **TXT estructurado**.

---

## 📊 Diagrama de Flujo del Sistema

A continuación se representa la arquitectura de control y procesamiento del sistema:

```mermaid
flowchart TD
    A(["INICIO"]) --> B["Definir Umbrales: 95%, 72%, 42%"]
    B --> C[/"Cargar Documento Word 1 y 2"/]
    C --> D{"¿Tienen texto válido?"}

    D -- No --> E["Mostrar Error de Lectura"]
    E --> Z(["FIN"])
    D -- Sí --> F["Extraer todas las líneas y celdas de tablas"]

    F --> G{"¿Es una línea<br>descriptiva o larga?"}
    G -- Sí --> H["Ignorar línea"]
    H --> F1{"¿Hay más líneas?"}
    G -- No --> I{"¿Empieza con<br>Módulo o Tema?"}
    I -- Sí --> J["Limpiar título y guardarlo en la lista"]
    I -- No --> H
    J --> F1

    F1 -- Sí --> G
    F1 -- No --> K["Preparar listas de temas de Doc 1 y Doc 2"]

    K --> L{"¿Quedan temas en Doc 1<br>por evaluar?"}
    L -- Sí --> M["Tomar Tema 1 del Doc 1"]
    M --> N{"¿Quedan temas en Doc 2<br>por evaluar?"}

    N -- Sí --> O["Tomar Tema 2 del Doc 2"]
    O --> P["Normalizar texto, extraer raíces y palabras clave"]
    P --> Q["Calcular porcentaje matemático de similitud"]

    Q --> R{"¿Porcentaje ≥ 95%?"}
    R -- Sí --> S["Categoría: COINCIDENCIA LITERAL"]
    R -- No --> T{"¿Porcentaje ≥ 72%?"}
    T -- Sí --> U["Categoría: SIMILITUD ALTA"]
    T -- No --> V{"¿Porcentaje ≥ 42%?"}
    V -- Sí --> W["Categoría: RELACIÓN TEMÁTICA"]
    V -- No --> X["Categoría: SIN RELACIÓN"]

    S --> Y["Añadir a la lista de comparaciones"]
    U --> Y
    W --> Y
    X --> Y
    Y --> N

    N -- No --> M1["Registrar la mejor similitud encontrada para el Tema 1"]
    M1 --> L

    L -- No --> O1["Ordenar todas las comparaciones de mayor a menor"]
    O1 --> P1[/"Mostrar tabla de resultados en pantalla"/]
    P1 --> Q1{"¿El usuario pide exportar?"}
    Q1 -- Sí --> R1[/"Guardar en archivo CSV / TXT"/]
    R1 --> S1{"¿El usuario pide<br>Nuevo Análisis?"}
    Q1 -- No --> S1

    S1 -- Sí --> T1["Ciclo WHILE: Borrar listas de memoria una por una"]
    T1 --> U1["Ciclo WHILE: Borrar filas de la pantalla una por una"]
    U1 --> C
    S1 -- No --> Z

    classDef inicioFin fill:#eef2ff,stroke:#818cf8,stroke-width:2px,color:#000;
    classDef proceso fill:#f0fdf4,stroke:#4ade80,stroke-width:2px,color:#000;
    classDef decision fill:#fff7ed,stroke:#fb923c,stroke-width:2px,color:#000;
    classDef inputOutput fill:#ecfeff,stroke:#22d3ee,stroke-width:2px,color:#000;

    class A,Z inicioFin;
    class B,F,H,J,K,M,O,P,Q,S,U,W,X,Y,M1,O1,T1,U1 proceso;
    class D,G,I,F1,L,N,R,T,V,Q1,S1 decision;
    class C,P1,R1 inputOutput;
```

---

## 📝 Pseudocódigo del Algoritmo

```text
ALGORITMO ComparadorAcademico
    // 1. Constantes y Umbrales Oficiales de Clasificación
    DEFINIR UMBRAL_LITERAL <- 95.0
    DEFINIR UMBRAL_ALTA <- 72.0
    DEFINIR UMBRAL_TEMATICA <- 42.0

    // 2. Control Principal de Ejecución (Estructura WHILE)
    DEFINIR sesion_activa <- VERDADERO

    MIENTRAS sesion_activa = VERDADERO HACER
        // Carga y validación de documentos Word (.docx)
        LEER archivo_doc1, archivo_doc2
        
        SI NO es_docx(archivo_doc1) O NO es_docx(archivo_doc2) ENTONCES
            MOSTRAR_ERROR "Formato no válido: únicamente se admiten archivos Word (.docx)"
        SINO
            // Extracción y delimitación exclusiva del Plan de Estudios
            lineas_doc1 <- EXTRAER_LINEAS_Y_TABLAS(archivo_doc1)
            lineas_doc2 <- EXTRAER_LINEAS_Y_TABLAS(archivo_doc2)

            seccion_doc1 <- DELIMITAR_PLAN_ESTUDIOS(lineas_doc1)
            seccion_doc2 <- DELIMITAR_PLAN_ESTUDIOS(lineas_doc2)

            temas_doc1 <- EXTRAER_TEMAS_PRINCIPALES(seccion_doc1) // Filtra 1.1 y descarta 1.1.1
            temas_doc2 <- EXTRAER_TEMAS_PRINCIPALES(seccion_doc2)

            // Comparación Sistemática entre Documentos (Estructura FOR anidada)
            DEFINIR lista_comparaciones <- LISTA_VACIA()

            PARA CADA tema1 EN temas_doc1 HACER
                DEFINIR mejor_similitud <- 0.0
                DEFINIR mejor_match <- NULO

                PARA CADA tema2 EN temas_doc2 HACER
                    // Normalización y cálculo léxico/morfológico
                    texto1_norm <- NORMALIZAR(tema1.titulo)
                    texto2_norm <- NORMALIZAR(tema2.titulo)

                    similitud, explicacion <- CALCULAR_SIMILITUD(texto1_norm, texto2_norm)

                    // Clasificación Oficial y Recomendación (Estructura IF / ELIF / ELSE)
                    SI similitud >= UMBRAL_LITERAL ENTONCES
                        categoria <- "COINCIDENCIA LITERAL"
                        recomendacion <- "Posible sustitución (Coincidencia exacta de títulos)"
                    SINO_SI similitud >= UMBRAL_ALTA ENTONCES
                        categoria <- "SIMILITUD ALTA"
                        recomendacion <- "Posible sustitución (Alta coincidencia textual)"
                    SINO_SI similitud >= UMBRAL_TEMATICA ENTONCES
                        categoria <- "RELACIÓN TEMÁTICA O PARALELA"
                        recomendacion <- "Similitud que requiere revisión (Relación temática)"
                    SINO
                        categoria <- "SIN RELACIÓN IDENTIFICADA"
                        recomendacion <- "Sin relación identificada (Coincidencia insuficiente)"
                    FIN_SI

                    REGISTRAR(lista_comparaciones, tema1, tema2, similitud, categoria, recomendacion, explicacion)
                FIN_PARA
            FIN_PARA

            // Ordenamiento y Despliegue de Resultados
            ORDENAR_POR_SIMILITUD_DESCENDENTE(lista_comparaciones)
            MOSTRAR_METRICAS_Y_TABLA(lista_comparaciones)

            // Exportación opcional
            SI usuario_solicita_exportar() ENTONCES
                EXPORTAR_CSV_Y_TXT(lista_comparaciones)
            FIN_SI

            // Reinicio de sesión interactivo
            SI usuario_solicita_reiniciar() ENTONCES
                // Limpieza de memoria (Estructura WHILE)
                MIENTRAS LONGITUD(lista_comparaciones) > 0 HACER
                    ELIMINAR_PRIMER_ELEMENTO(lista_comparaciones)
                FIN_MIENTRAS
                VACIAR_ARCHIVOS_CARGADOS()
                REINICIAR_INTERFAZ()
            SINO
                sesion_activa <- FALSO
            FIN_SI
        FIN_SI
    FIN_MIENTRAS
FIN_ALGORITMO
```

---

## 🛠️ Estructura del Repositorio

```text
├── streamlit_app.py    # Aplicación web completa y motor algorítmico autónomo
├── requirements.txt    # Dependencias mínimas (streamlit, python-docx, pandas)
└── README.md           # Documentación, diagrama de flujo y pseudocódigo
```

---

## 💻 Ejecución Local (Opcional)

Si deseas ejecutar la aplicación en tu propio ordenador:

1. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```
2. Iniciar la aplicación web:
   ```bash
   streamlit run streamlit_app.py
   ```
