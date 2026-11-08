"""
================================================================================
COMPARADOR ACADÉMICO 2.0 - APLICACIÓN WEB STREAMLIT
================================================================================
Desarrollada para despliegue en Streamlit Community Cloud y ejecución en navegador.
Permite contrastar de manera interactiva dos documentos Word (.docx),
aislando la sección del Plan de Estudios y comparando exclusivamente temas principales.
"""

import csv
import io
import os
import pandas as pd
import streamlit as st

# Importar funciones del núcleo analítico
from core.extractor import cargar_documento, extraer_temas
from core.comparator import comparar_temas, clasificar_resultado
from core.reporter import generar_informe

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Comparador Académico 2.0",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para una interfaz moderna y profesional
st.markdown("""
<style>
    /* Estilos generales */
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
    .badge-literal {
        background-color: #DCFCE7;
        color: #166534;
        padding: 4px 8px;
        border-radius: 6px;
        font-weight: 600;
    }
    .badge-alta {
        background-color: #DBEAFE;
        color: #1E40AF;
        padding: 4px 8px;
        border-radius: 6px;
        font-weight: 600;
    }
    .badge-tematica {
        background-color: #FEF3C7;
        color: #92400E;
        padding: 4px 8px;
        border-radius: 6px;
        font-weight: 600;
    }
    .badge-sin-relacion {
        background-color: #F1F5F9;
        color: #475569;
        padding: 4px 8px;
        border-radius: 6px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# BARRA LATERAL (INFORMACIÓN Y REGLAS DEL SISTEMA)
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/diploma.png", width=70)
    st.title("Comparador 2.0")
    st.markdown("**Sistema Académico de Posgrados**")
    st.markdown("---")
    
    st.markdown("### 📌 Reglas de Extracción")
    st.markdown("""
    - **Ámbito exclusivo:** Localiza y analiza únicamente las páginas del **Plan de Estudios** (descartando admisiones, rankings, claustro y metodologías).
    - **Módulos:** Numeración de primer nivel (`1.`, `2.`, `Módulo 1`).
    - **Temas principales:** Viñetas de dos niveles (`1.1`, `1.2`, `2.1`).
    - **Subtemas (`1.1.1`, `X.Y.Z`):** **Estrictamente excluidos**, al igual que descripciones y objetivos.
    """)
    
    st.markdown("---")
    st.markdown("### 🏷️ Categorías Oficiales")
    st.markdown("""
    - 🟢 **Coincidencia Literal** (≥ 95%)
    - 🔵 **Similitud Alta** (72% - 94.9%)
    - 🟡 **Relación Temática** (42% - 71.9%)
    - ⚪ **Sin Relación** (< 42%)
    """)
    
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0

    st.markdown("---")
    if st.button("🔄 Reiniciar / Limpiar sesión"):
        nuevo_id = st.session_state.get("uploader_key", 0) + 1
        st.session_state.clear()
        st.session_state["uploader_key"] = nuevo_id
        st.rerun()

# ==============================================================================
# ENCABEZADO PRINCIPAL
# ==============================================================================
st.markdown('<div class="main-header">🎓 COMPARADOR ACADÉMICO 2.0</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Análisis algorítmico exclusivo entre títulos de temas contenidos en el Plan de Estudios de documentos Word (.docx).</div>', unsafe_allow_html=True)

# ==============================================================================
# PANELES DE CARGA DE DOCUMENTOS
# ==============================================================================
col_doc1, col_doc2 = st.columns(2)

with col_doc1:
    st.subheader("📄 Documento 1 (Origen)")
    st.caption("Dosier o programa cuyos temas se desean comparar.")
    archivo1 = st.file_uploader(
        "Cargar archivo Word (.docx)",
        type=["docx"],
        key=f"doc1_uploader_{st.session_state['uploader_key']}",
        help="Seleccione el primer documento Word (.docx)"
    )

with col_doc2:
    st.subheader("📄 Documento 2 (Comparación)")
    st.caption("Segundo dosier o plan curricular de contraste.")
    archivo2 = st.file_uploader(
        "Cargar archivo Word (.docx)",
        type=["docx"],
        key=f"doc2_uploader_{st.session_state['uploader_key']}",
        help="Seleccione el segundo documento Word (.docx)"
    )

# ==============================================================================
# VALIDACIÓN Y EXTRACCIÓN PRELIMINAR
# ==============================================================================
temas_doc1 = None
temas_doc2 = None

if archivo1 and archivo2:
    try:
        lineas1 = cargar_documento(archivo1)
        temas_doc1 = extraer_temas(lineas1, nombre_doc=archivo1.name)

        lineas2 = cargar_documento(archivo2)
        temas_doc2 = extraer_temas(lineas2, nombre_doc=archivo2.name)

        col_prev1, col_prev2 = st.columns(2)
        with col_prev1:
            st.success(f"**{archivo1.name}**: {len(temas_doc1)} temas principales identificados.")
            with st.expander("Ver temas extraídos de Documento 1"):
                for t in temas_doc1:
                    st.write(f"- **[{t['modulo_num']}]** {t['tema_num']} - {t['tema_titulo']}")

        with col_prev2:
            st.success(f"**{archivo2.name}**: {len(temas_doc2)} temas principales identificados.")
            with st.expander("Ver temas extraídos de Documento 2"):
                for t in temas_doc2:
                    st.write(f"- **[{t['modulo_num']}]** {t['tema_num']} - {t['tema_titulo']}")

        st.markdown("---")
        
        # Botón para ejecutar el análisis
        if st.button("🔍 EJECUTAR ANÁLISIS COMPARATIVO", type="primary", use_container_width=True):
            with st.spinner("Comparando sistemáticamente títulos de temas entre ambos documentos..."):
                comparaciones, resumen, sin_match = comparar_temas(temas_doc1, temas_doc2)
                st.session_state["analisis"] = {
                    "comparaciones": comparaciones,
                    "resumen": resumen,
                    "sin_match": sin_match,
                    "nombre_doc1": archivo1.name,
                    "nombre_doc2": archivo2.name
                }
            st.rerun()

    except Exception as e:
        st.error(f"Error al procesar los documentos: {str(e)}")

# ==============================================================================
# VISUALIZACIÓN DE RESULTADOS
# ==============================================================================
if "analisis" in st.session_state:
    datos_analisis = st.session_state["analisis"]
    comparaciones = datos_analisis["comparaciones"]
    resumen = datos_analisis["resumen"]
    sin_match = datos_analisis["sin_match"]
    nombre_doc1 = datos_analisis["nombre_doc1"]
    nombre_doc2 = datos_analisis["nombre_doc2"]

    st.markdown("### 📊 Resumen Estadístico del Análisis")
    
    # Tarjetas KPI de métricas
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("Temas Doc 1", resumen["total_temas_doc1"])
    kpi2.metric("Temas Doc 2", resumen["total_temas_doc2"])
    kpi3.metric("🟢 Literales", resumen["coincidencias_literales"])
    kpi4.metric("🔵 Altas", resumen["similitudes_altas"])
    kpi5.metric("🟡 Temáticas", resumen["relaciones_tematicas"])
    kpi6.metric("⚪ Sin Relación", resumen["sin_coincidencias"])

    st.markdown("---")
    
    # Barra de filtros
    st.markdown("### 📋 Tabla Comparativa de Temas")
    col_filtro1, col_filtro2 = st.columns([1, 2])
    
    with col_filtro1:
        categoria_filtro = st.selectbox(
            "Filtrar por categoría:",
            ["Todas las categorías", "COINCIDENCIA LITERAL", "SIMILITUD ALTA", "RELACIÓN TEMÁTICA O PARALELA", "SIN RELACIÓN IDENTIFICADA"]
        )

    with col_filtro2:
        texto_busqueda = st.text_input("Buscar por palabra o término en los títulos:", "").lower().strip()

    # Filtrar resultados
    resultados_filtrados = []
    for r in comparaciones:
        if categoria_filtro != "Todas las categorías" and r["categoria"] != categoria_filtro:
            continue
        if texto_busqueda:
            t1 = r["doc1_tema_titulo"].lower()
            t2 = r["doc2_tema_titulo"].lower()
            if texto_busqueda not in t1 and texto_busqueda not in t2:
                continue
        resultados_filtrados.append(r)

    st.caption(f"Mostrando **{len(resultados_filtrados)}** de **{len(comparaciones)}** comparaciones evaluadas.")

    # Construir DataFrame para visualización limpia
    filas_tabla = []
    for r in resultados_filtrados:
        filas_tabla.append({
            "Módulo Doc 1": r["doc1_modulo"],
            "Tema Doc 1": r["doc1_tema_titulo"],
            "Módulo Doc 2": r["doc2_modulo"],
            "Tema Doc 2": r["doc2_tema_titulo"],
            "Similitud": f"{r['similitud']}%",
            "Clasificación": r["categoria"],
            "Recomendación": r["recomendacion"]
        })

    df_tabla = pd.DataFrame(filas_tabla)

    if not df_tabla.empty:
        st.dataframe(df_tabla, use_container_width=True, height=450)
    else:
        st.info("No se encontraron temas con los criterios de filtro aplicados.")

    # ==============================================================================
    # INSPECTOR DE DETALLE TÉCNICO (JUSTIFICACIÓN LÉXICA)
    # ==============================================================================
    st.markdown("### 🔍 Justificación Textual de Coincidencias")
    st.caption("Seleccione una coincidencia para examinar los términos y raíces compartidos:")
    
    opciones_inspeccion = [
        f"[{r['similitud']}%] {r['doc1_tema_titulo']}  ⟷  {r['doc2_tema_titulo']}"
        for r in resultados_filtrados[:50]
    ]
    
    if opciones_inspeccion:
        seleccion = st.selectbox("Seleccione tema a inspeccionar:", opciones_inspeccion)
        if seleccion:
            idx_sel = opciones_inspeccion.index(seleccion)
            reg = resultados_filtrados[idx_sel]
            
            with st.container():
                st.info(f"""
                **Doc 1:** [{reg['doc1_modulo']}] {reg['doc1_tema_titulo']}  
                **Doc 2:** [{reg['doc2_modulo']}] {reg['doc2_tema_titulo']}  
                **Porcentaje de similitud:** `{reg['similitud']}%` ({reg['categoria']})  
                **Recomendación:** *{reg['recomendacion']}*  
                **Explicación técnica:** {reg['explicacion']}
                """)

    # ==============================================================================
    # EXPORTACIÓN DE RESULTADOS
    # ==============================================================================
    st.markdown("---")
    st.markdown("### 📥 Exportar Resultados del Análisis")
    
    col_exp1, col_exp2 = st.columns(2)
    
    # 1. Generar CSV en memoria con UTF-8-SIG
    output_csv = io.StringIO()
    writer = csv.writer(output_csv)
    writer.writerow(["COMPARADOR ACADÉMICO 2.0 - INFORME COMPARATIVO"])
    writer.writerow(["Documento 1", nombre_doc1])
    writer.writerow(["Documento 2", nombre_doc2])
    writer.writerow([])
    writer.writerow(["Módulo Doc 1", "Tema Doc 1", "Módulo Doc 2", "Tema Doc 2", "Similitud (%)", "Clasificación", "Recomendación", "Explicación"])
    for r in comparaciones:
        writer.writerow([
            r["doc1_modulo"], r["doc1_tema_titulo"],
            r["doc2_modulo"], r["doc2_tema_titulo"],
            f"{r['similitud']}%", r["categoria"], r["recomendacion"], r["explicacion"]
        ])
    csv_bytes = output_csv.getvalue().encode("utf-8-sig")

    with col_exp1:
        st.download_button(
            label="📊 Descargar Informe CSV (Excel)",
            data=csv_bytes,
            file_name="informe_comparador_academico.csv",
            mime="text/csv",
            use_container_width=True
        )

    # 2. Generar TXT estructurado en memoria
    output_txt = io.StringIO()
    output_txt.write("=" * 80 + "\n")
    output_txt.write("       COMPARADOR ACADÉMICO 2.0 - INFORME COMPARATIVO DE TEMAS\n")
    output_txt.write("=" * 80 + "\n\n")
    output_txt.write(f"Documento 1: {nombre_doc1}\n")
    output_txt.write(f"Documento 2: {nombre_doc2}\n")
    output_txt.write(f"Coincidencias Literales: {resumen['coincidencias_literales']}\n")
    output_txt.write(f"Similitudes Altas:       {resumen['similitudes_altas']}\n")
    output_txt.write(f"Relaciones Temáticas:    {resumen['relaciones_tematicas']}\n\n")
    output_txt.write("-" * 80 + "\nDETALLE DE COMPARACIONES\n" + "-" * 80 + "\n\n")
    for i, r in enumerate(comparaciones, 1):
        output_txt.write(f"[{i}] {r['similitud']}% | {r['categoria']}\n")
        output_txt.write(f"    Doc 1: [{r['doc1_modulo']}] {r['doc1_tema_titulo']}\n")
        output_txt.write(f"    Doc 2: [{r['doc2_modulo']}] {r['doc2_tema_titulo']}\n")
        output_txt.write(f"    Recomendación: {r['recomendacion']}\n")
        output_txt.write(f"    Explicación:   {r['explicacion']}\n\n")
    txt_bytes = output_txt.getvalue().encode("utf-8")

    with col_exp2:
        st.download_button(
            label="📄 Descargar Informe TXT",
            data=txt_bytes,
            file_name="informe_comparador_academico.txt",
            mime="text/plain",
            use_container_width=True
        )
        
