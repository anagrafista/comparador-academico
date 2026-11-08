"""
Batería de pruebas automatizadas para COMPARADOR ACADÉMICO 2.0.
Verifica:
1. Lectura y extracción exclusiva en Word (.docx).
2. Detección y exclusión de subtemas, descripciones y objetivos.
3. Rechazo de formatos no válidos.
4. Coincidencias literales, altas, temáticas y sin relación.
5. Presencia y funcionamiento de las 5 estructuras obligatorias (if, elif, else, for, while).
6. Presencia y funcionamiento de las funciones principales.
7. Exportación a CSV y TXT.
"""

import os
import sys
import unittest
import tempfile
import inspect

# Agregar raíz al path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.extractor import (
    cargar_documento,
    leer_docx,
    extraer_temas
)
from core.comparator import (
    normalizar_texto,
    extraer_palabras_clave,
    calcular_similitud,
    clasificar_resultado,
    comparar_temas
)
from core.reporter import generar_informe


class TestComparadorAcademico(unittest.TestCase):

    def setUp(self):
        self.ruta_base = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        self.dir_test = os.path.join(self.ruta_base, "test_data")
        self.docx_doc1 = os.path.join(self.dir_test, "programa_master_direccion_doc1.docx")
        self.docx_doc2 = os.path.join(self.dir_test, "programa_master_gestion_doc2.docx")

    def test_01_funciones_obligatorias_existen(self):
        """Verifica que las funciones solicitadas existan y sean invocables."""
        funciones = [
            cargar_documento,
            leer_docx,
            extraer_temas,
            normalizar_texto,
            comparar_temas,
            calcular_similitud,
            clasificar_resultado,
            generar_informe
        ]
        for fn in funciones:
            self.assertTrue(callable(fn), f"La función {fn.__name__} no es invocable.")

    def test_02_lectura_formato_docx(self):
        """Verifica la lectura de archivos Word (.docx)."""
        lineas_doc1 = cargar_documento(self.docx_doc1)
        self.assertIsInstance(lineas_doc1, list)
        self.assertGreater(len(lineas_doc1), 5)

        lineas_doc2 = cargar_documento(self.docx_doc2)
        self.assertIsInstance(lineas_doc2, list)
        self.assertGreater(len(lineas_doc2), 5)

    def test_03_formato_invalido_rechazado_con_aviso(self):
        """Verifica que un intento de carga en formato no admitido informe que solo se admite DOCX."""
        with self.assertRaises(ValueError) as ctx:
            cargar_documento("documento_invalido.xyz")
        self.assertIn("docx", str(ctx.exception).lower())

    def test_04_extraccion_exclusiva_de_temas(self):
        """Verifica que SOLO se extraigan títulos de temas, descartando objetivos y subtemas."""
        lineas = cargar_documento(self.docx_doc1)
        temas = extraer_temas(lineas, "Doc1")

        self.assertGreaterEqual(len(temas), 8)

        titulos = [t["tema_titulo"].lower() for t in temas]

        # Temas que DEBEN estar
        self.assertTrue(any("planificación estratégica" in t for t in titulos))
        self.assertTrue(any("marketing digital" in t for t in titulos))
        self.assertTrue(any("finanzas corporativas" in t for t in titulos))

        # Contenidos que NO deben estar (objetivos, descripciones, subtemas, bibliografía)
        for t in titulos:
            self.assertNotIn("competencias", t)
            self.assertNotIn("porter, m", t)
            self.assertNotIn("diagnóstico foda", t)  # Subtema 1.1.1 descartado
            self.assertNotIn("capacitar al estudiante", t)  # Descripción descartada

    def test_05_normalizacion_de_texto(self):
        """Verifica la normalización: minúsculas, sin tildes, sin signos."""
        orig = "¡Gestión Estratégica, Innovación & Marketing 2.0!"
        norm = normalizar_texto(orig)
        self.assertEqual(norm, "gestion estrategica innovacion marketing 2 0")

    def test_06_similitud_y_clasificacion_literal(self):
        """Verifica coincidencias literales (Categoría 1)."""
        t1 = "Marketing digital y redes sociales"
        t2 = "Marketing digital y redes sociales"
        sim, exp = calcular_similitud(t1, t2)
        cat, rec, niv = clasificar_resultado(sim)

        self.assertEqual(sim, 100.0)
        self.assertEqual(cat, "COINCIDENCIA LITERAL")
        self.assertEqual(niv, 1)
        self.assertIn("Posible sustitución", rec)

    def test_07_similitud_y_clasificacion_alta(self):
        """Verifica similitudes altas (Categoría 2)."""
        t1 = "Gestión de proyectos empresariales"
        t2 = "Administración de proyectos de empresa"
        sim, exp = calcular_similitud(t1, t2)
        cat, rec, niv = clasificar_resultado(sim)

        self.assertGreaterEqual(sim, 65.0)
        self.assertIn("proyectos", exp.lower())
        self.assertIn(cat, ["SIMILITUD ALTA", "RELACIÓN TEMÁTICA O PARALELA"])

    def test_08_similitud_y_clasificacion_tematica(self):
        """Verifica relaciones temáticas (Categoría 3)."""
        t1 = "Ciberseguridad y protección de datos"
        t2 = "Seguridad de la información en entornos cloud"
        sim, exp = calcular_similitud(t1, t2)
        cat, rec, niv = clasificar_resultado(sim)

        self.assertGreaterEqual(sim, 35.0)
        self.assertIn("seguridad", exp.lower())

    def test_09_sin_relacion_identificada(self):
        """Verifica temas sin relación (Categoría 4)."""
        t1 = "Derecho laboral internacional"
        t2 = "Robótica industrial y automatización de procesos"
        sim, exp = calcular_similitud(t1, t2)
        cat, rec, niv = clasificar_resultado(sim)

        self.assertLess(sim, 40.0)
        self.assertEqual(cat, "SIN RELACIÓN IDENTIFICADA")
        self.assertEqual(niv, 4)

    def test_10_comparacion_completa_entre_documentos(self):
        """Verifica el ciclo for sistemático comparando Doc 1 y Doc 2."""
        temas1 = extraer_temas(cargar_documento(self.docx_doc1), "Doc 1")
        temas2 = extraer_temas(cargar_documento(self.docx_doc2), "Doc 2")

        comparaciones, resumen, sin_match = comparar_temas(temas1, temas2)

        self.assertEqual(len(comparaciones), len(temas1) * len(temas2))
        self.assertGreater(resumen["coincidencias_literales"], 0)
        self.assertGreater(resumen["total_temas_doc1"], 0)
        self.assertGreater(resumen["total_temas_doc2"], 0)

    def test_11_exportacion_informes(self):
        """Verifica la generación de informes en CSV y TXT."""
        temas1 = extraer_temas(cargar_documento(self.docx_doc1), "Doc 1")
        temas2 = extraer_temas(cargar_documento(self.docx_doc2), "Doc 2")
        comparaciones, resumen, sin_match = comparar_temas(temas1, temas2)

        with tempfile.TemporaryDirectory() as tmpdir:
            # CSV
            ruta_csv = os.path.join(tmpdir, "informe.csv")
            generar_informe(comparaciones, resumen, {"nombre": "Doc1"}, {"nombre": "Doc2"}, ruta_csv, "csv")
            self.assertTrue(os.path.exists(ruta_csv))
            self.assertGreater(os.path.getsize(ruta_csv), 200)

            # TXT
            ruta_txt = os.path.join(tmpdir, "informe.txt")
            generar_informe(comparaciones, resumen, {"nombre": "Doc1"}, {"nombre": "Doc2"}, ruta_txt, "txt")
            self.assertTrue(os.path.exists(ruta_txt))
            self.assertGreater(os.path.getsize(ruta_txt), 200)

    def test_12_verificacion_estructuras_obligatorias(self):
        """
        Verificación de presencia de código fuente de las 5 estructuras obligatorias:
        if, elif, else, for, while.
        """
        from core import comparator, extractor, reporter
        from gui import app

        modulos = [comparator, extractor, reporter, app]
        codigo_total = "".join(inspect.getsource(m) for m in modulos)

        self.assertIn("if ", codigo_total, "Estructura 'if' no encontrada en código.")
        self.assertIn("elif ", codigo_total, "Estructura 'elif' no encontrada en código.")
        self.assertIn("else:", codigo_total, "Estructura 'else' no encontrada en código.")
        self.assertIn("for ", codigo_total, "Estructura 'for' no encontrada en código.")
        self.assertIn("while ", codigo_total, "Estructura 'while' no encontrada en código.")

    def test_13_dosier_extenso_plan_estudios_y_subtemas(self):
        """
        Verifica que en documentos extensos (44+ páginas) se aísle la sección
        del 'Plan de estudios' y se extraigan SOLO temas principales (1.1, 1.2),
        descartando subtemas (1.1.1, 1.1.2) y secciones externas.
        """
        dosier_simulado = """
        PÁGINA 1
        PRESENTACIÓN INSTITUCIONAL
        1. Carta del Rector
        2. Rankings de la Universidad
        
        PÁGINA 14
        PLAN DE ESTUDIOS
        1. GESTIÓN ESTRATÉGICA
        1.1 Planificación y Modelos de Negocio
        1.1.1 Diagnóstico de entorno PESTEL
        1.1.2 Cadena de valor
        1.2 Formulación Estratégica
        1.2.1 Metodologías ágiles
        2. FINANZAS Y CONTROL
        2.1 Análisis Financiero
        2.1.1 Ratios de liquidez
        
        PÁGINA 25
        METODOLOGÍA DOCENTE
        1. Metodología de casos
        2. Claustro de profesores
        """
        temas = extraer_temas(dosier_simulado, "DosierPrueba")
        titulos = [t["tema_titulo"] for t in temas]

        # Temas que DEBEN estar (1.1, 1.2, 2.1)
        self.assertTrue(any("Planificación" in t for t in titulos))
        self.assertTrue(any("Formulación" in t for t in titulos))
        self.assertTrue(any("Análisis Financiero" in t for t in titulos))

        # Contenidos externos o subtemas que NO deben estar
        self.assertFalse(any("Carta del Rector" in t for t in titulos))
        self.assertFalse(any("Metodología de casos" in t for t in titulos))
        self.assertFalse(any("PESTEL" in t for t in titulos))
        self.assertFalse(any("Cadena de valor" in t for t in titulos))
        self.assertEqual(len(temas), 3)


if __name__ == "__main__":
    unittest.main()

