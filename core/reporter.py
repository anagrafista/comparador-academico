"""
Módulo de exportación y generación de informes de comparación académica.
Permite exportar los resultados a formatos CSV y TXT con resumen estadístico,
metadatos de análisis y tabla detallada de relaciones de temas.
"""

import os
import csv
from datetime import datetime


def generar_informe(resultados, resumen, info_doc1, info_doc2, ruta_salida, formato="csv"):
    """
    Función obligatoria para generar el informe comparativo en CSV o TXT.
    
    Parámetros:
    - resultados: lista de comparaciones ordenadas por similitud.
    - resumen: diccionario con las estadísticas del análisis.
    - info_doc1: metadatos del Documento 1 (nombre, ruta, total temas).
    - info_doc2: metadatos del Documento 2 (nombre, ruta, total temas).
    - ruta_salida: archivo de destino seleccionado por el usuario.
    - formato: 'csv' o 'txt'.
    """
    if not ruta_salida:
        raise ValueError("Ruta de archivo de salida no especificada.")

    fecha_hora_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formato = formato.lower().strip()

    if formato == "csv" or ruta_salida.lower().endswith(".csv"):
        _exportar_csv(resultados, resumen, info_doc1, info_doc2, ruta_salida, fecha_hora_actual)
    elif formato == "txt" or ruta_salida.lower().endswith(".txt"):
        _exportar_txt(resultados, resumen, info_doc1, info_doc2, ruta_salida, fecha_hora_actual)
    else:
        # Por defecto según extensión o CSV
        if ruta_salida.lower().endswith(".txt"):
            _exportar_txt(resultados, resumen, info_doc1, info_doc2, ruta_salida, fecha_hora_actual)
        else:
            _exportar_csv(resultados, resumen, info_doc1, info_doc2, ruta_salida, fecha_hora_actual)

    return ruta_salida


def _exportar_csv(resultados, resumen, info_doc1, info_doc2, ruta_salida, fecha_hora):
    """
    Exporta el informe a formato CSV con codificación UTF-8-SIG (compatible con Excel).
    """
    with open(ruta_salida, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)

        # Encabezado informativo y metadatos
        writer.writerow(["COMPARADOR ACADÉMICO 2.0 - INFORME DE COMPARACIÓN DE TEMAS"])
        writer.writerow(["Fecha del análisis", fecha_hora])
        writer.writerow(["Documento 1 (Origen)", info_doc1.get("nombre", "Desconocido")])
        writer.writerow(["Documento 2 (Comparación)", info_doc2.get("nombre", "Desconocido")])
        writer.writerow([])

        # Resumen general
        writer.writerow(["RESUMEN GENERAL DEL ANÁLISIS"])
        writer.writerow(["Métrica", "Cantidad"])
        writer.writerow(["Temas identificados en Documento 1", resumen.get("total_temas_doc1", 0)])
        writer.writerow(["Temas identificados en Documento 2", resumen.get("total_temas_doc2", 0)])
        writer.writerow(["Coincidencias literales", resumen.get("coincidencias_literales", 0)])
        writer.writerow(["Similitudes altas", resumen.get("similitudes_altas", 0)])
        writer.writerow(["Relaciones temáticas", resumen.get("relaciones_tematicas", 0)])
        writer.writerow(["Temas de Doc 1 sin coincidencias", resumen.get("sin_coincidencias", 0)])
        writer.writerow(["Total de comparaciones procesadas", resumen.get("total_comparaciones_realizadas", 0)])
        writer.writerow([])

        # Tabla de resultados
        writer.writerow([
            "Módulo Doc 1",
            "Tema Doc 1",
            "Módulo Doc 2",
            "Tema Doc 2",
            "Similitud (%)",
            "Clasificación",
            "Recomendación",
            "Explicación de Coincidencia"
        ])

        for r in resultados:
            # Mostrar tanto las coincidencias relevantes como las demás ordenadas
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


def _exportar_txt(resultados, resumen, info_doc1, info_doc2, ruta_salida, fecha_hora):
    """
    Exporta el informe a formato de texto plano estructurado y legible (.txt).
    """
    separador = "=" * 90
    subseparador = "-" * 90

    with open(ruta_salida, mode="w", encoding="utf-8") as f:
        f.write(f"{separador}\n")
        f.write("          COMPARADOR ACADÉMICO 2.0 - INFORME DE ANÁLISIS COMPARATIVO\n")
        f.write(f"{separador}\n\n")

        f.write(f"Fecha y hora del análisis: {fecha_hora}\n")
        f.write(f"Documento 1 (Origen):     {info_doc1.get('nombre', 'Desconocido')}\n")
        f.write(f"Documento 2 (Comparación): {info_doc2.get('nombre', 'Desconocido')}\n\n")

        f.write(f"{subseparador}\n")
        f.write("RESUMEN ESTADÍSTICO\n")
        f.write(f"{subseparador}\n")
        f.write(f"- Total de temas identificados en Documento 1: {resumen.get('total_temas_doc1', 0)}\n")
        f.write(f"- Total de temas identificados en Documento 2: {resumen.get('total_temas_doc2', 0)}\n")
        f.write(f"- Coincidencias literales encontradas:         {resumen.get('coincidencias_literales', 0)}\n")
        f.write(f"- Similitudes altas identificadas:             {resumen.get('similitudes_altas', 0)}\n")
        f.write(f"- Relaciones temáticas / paralelas:            {resumen.get('relaciones_tematicas', 0)}\n")
        f.write(f"- Temas de Doc 1 sin coincidencias:            {resumen.get('sin_coincidencias', 0)}\n\n")

        f.write(f"{subseparador}\n")
        f.write("DETALLE DE COMPARACIÓN ENTRE TÍTULOS DE TEMAS\n")
        f.write(f"{subseparador}\n\n")

        for i, r in enumerate(resultados, 1):
            f.write(f"[{i}] Similitud: {r['similitud']}% | Clasificación: {r['categoria']}\n")
            f.write(f"    Doc 1: [{r['doc1_modulo']}] {r['doc1_tema_titulo']}\n")
            f.write(f"    Doc 2: [{r['doc2_modulo']}] {r['doc2_tema_titulo']}\n")
            f.write(f"    Recomendación: {r['recomendacion']}\n")
            f.write(f"    Detalle: {r['explicacion']}\n")
            f.write(f"{'.' * 70}\n")
