# 🎓 Comparador Académico 2.0

Sistema web interactivo desarrollado en **Python** y **Streamlit** para la evaluación y contraste sistemático de planes de estudio contenidos en documentos Word (`.docx`).

---

## 🚀 Acceso en la Nube

La aplicación está disponible en línea y lista para usarse sin necesidad de instalación local:
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

## 🛠️ Estructura del Repositorio

```text
├── streamlit_app.py    # Aplicación web completa y motor algorítmico autónomo
├── requirements.txt    # Librerías necesarias (streamlit, python-docx, pandas)
└── README.md           # Descripción y documentación del proyecto
```

---

## 💻 Ejecución Local (Opcional)

Si deseas ejecutar la aplicación en tu ordenador:

1. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```
2. Iniciar la aplicación web:
   ```bash
   streamlit run streamlit_app.py
   ```
