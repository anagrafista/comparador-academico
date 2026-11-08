"""
Interfaz gráfica de usuario (GUI) para COMPARADOR ACADÉMICO 2.0.
Desarrollada en Tkinter para ejecución nativa en macOS y otros sistemas operativos.

Cumple con todos los requisitos de diseño, legibilidad, responsividad
y las estructuras de control obligatorias (if, elif, else, for, while).
"""

import os
import sys
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Importar funciones de los módulos del núcleo
from core.extractor import (
    cargar_documento,
    leer_docx,
    extraer_temas
)
from core.comparator import (
    normalizar_texto,
    calcular_similitud,
    clasificar_resultado,
    comparar_temas
)
from core.reporter import generar_informe


class ComparadorAcademicoGUI:
    """
    Clase principal de la interfaz gráfica del Comparador Académico 2.0.
    """

    def __init__(self, root):
        self.root = root
        self.root.title("COMPARADOR ACADÉMICO 2.0")
        self.root.geometry("1180x820")
        self.root.minsize(980, 680)

        # Configuración de variables de estado
        self.doc1_ruta = None
        self.doc1_datos = None  # Lista de temas extraídos
        self.doc1_formato = ""
        self.doc1_nombre = ""

        self.doc2_ruta = None
        self.doc2_datos = None  # Lista de temas extraídos
        self.doc2_formato = ""
        self.doc2_nombre = ""

        self.resultados_actuales = []
        self.resumen_actual = {}
        self.temas_sin_match = []

        # Cola de control para el ciclo WHILE de repetición y gestión de análisis
        self.cola_solicitudes = queue.Queue()
        self.app_activa = True
        self.analisis_en_curso = False

        # Configurar estilos visuales
        self._configurar_estilos()

        # Construir componentes de la interfaz
        self._construir_interfaz()

        # Iniciar ciclo de vida de la aplicación con estructura WHILE compatible
        self._iniciar_ciclo_control()

    def _configurar_estilos(self):
        """Configura una paleta de colores moderna y neutra adecuada para macOS."""
        self.colores = {
            "fondo": "#F8FAFC",          # Slate 50
            "panel": "#FFFFFF",          # Blanco puro
            "borde": "#E2E8F0",          # Slate 200
            "primario": "#2563EB",       # Azul corporativo
            "primario_hover": "#1D4ED8",
            "texto_oscuro": "#0F172A",   # Slate 900
            "texto_mutado": "#64748B",   # Slate 500
            "exito": "#059669",          # Esmeralda 600
            "advertencia": "#D97706",    # Ámbar 600
            "peligro": "#DC2626",        # Rojo 600
            "info": "#475569"            # Slate 600
        }

        self.root.configure(bg=self.colores["fondo"])

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(".", font=("Helvetica", 11), background=self.colores["fondo"])
        style.configure("TLabel", background=self.colores["fondo"], foreground=self.colores["texto_oscuro"])
        style.configure("Card.TFrame", background=self.colores["panel"], relief="solid", borderwidth=1)
        style.configure("Treeview",
                        font=("Helvetica", 10),
                        rowheight=28,
                        background="#FFFFFF",
                        fieldbackground="#FFFFFF",
                        foreground=self.colores["texto_oscuro"])
        style.configure("Treeview.Heading",
                        font=("Helvetica", 10, "bold"),
                        background="#F1F5F9",
                        foreground=self.colores["texto_oscuro"])
        style.map("Treeview", background=[("selected", "#DBEAFE")], foreground=[("selected", "#1E3A8A")])

    def _construir_interfaz(self):
        """Construye todos los elementos visuales de la ventana principal."""
        contenedor_principal = tk.Frame(self.root, bg=self.colores["fondo"], padx=18, pady=14)
        contenedor_principal.pack(fill=tk.BOTH, expand=True)

        # 1. ENCABEZADO Y DESCRIPCIÓN
        frame_cabecera = tk.Frame(contenedor_principal, bg=self.colores["fondo"])
        frame_cabecera.pack(fill=tk.X, pady=(0, 10))

        lbl_titulo = tk.Label(
            frame_cabecera,
            text="COMPARADOR ACADÉMICO 2.0",
            font=("Helvetica", 20, "bold"),
            bg=self.colores["fondo"],
            fg=self.colores["texto_oscuro"]
        )
        lbl_titulo.pack(anchor="w")

        lbl_subtitulo = tk.Label(
            frame_cabecera,
            text="Comparación analítica exclusiva entre títulos de temas contenidos en dosieres, índices y programas académicos.",
            font=("Helvetica", 11),
            bg=self.colores["fondo"],
            fg=self.colores["texto_mutado"]
        )
        lbl_subtitulo.pack(anchor="w", pady=(2, 0))

        # 2. PANELES DE CARGA DE DOCUMENTOS (Doc 1 y Doc 2)
        frame_carga = tk.Frame(contenedor_principal, bg=self.colores["fondo"])
        frame_carga.pack(fill=tk.X, pady=(0, 10))

        # Configurar dos columnas de igual ancho
        frame_carga.columnconfigure(0, weight=1, uniform="docs")
        frame_carga.columnconfigure(1, weight=1, uniform="docs")

        self.panel_doc1 = self._crear_panel_documento(
            frame_carga,
            col=0,
            num_doc=1,
            titulo="DOCUMENTO 1: PROGRAMA DE ORIGEN",
            sub="Dosier o programa cuyos temas se desean comparar"
        )

        self.panel_doc2 = self._crear_panel_documento(
            frame_carga,
            col=1,
            num_doc=2,
            titulo="DOCUMENTO 2: PROGRAMA DE COMPARACIÓN",
            sub="Segundo dosier o plan curricular de contraste"
        )

        # 3. BARRA DE ACCIÓN PRINCIPAL (Botón Analizar + Estado)
        frame_accion = tk.Frame(contenedor_principal, bg=self.colores["panel"], relief="solid", bd=1)
        frame_accion.pack(fill=tk.X, pady=(0, 10), ipady=6, padx=2)

        self.btn_analizar = tk.Button(
            frame_accion,
            text="🔍  ANALIZAR DOCUMENTOS",
            font=("Helvetica", 13, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=8,
            command=self.solicitar_analisis
        )
        self.btn_analizar.pack(side=tk.LEFT, padx=14)

        self.lbl_estado = tk.Label(
            frame_accion,
            text="Seleccione ambos documentos Word (.docx) para iniciar el análisis.",
            font=("Helvetica", 11),
            bg=self.colores["panel"],
            fg=self.colores["texto_mutado"]
        )
        self.lbl_estado.pack(side=tk.LEFT, padx=10)

        # 4. TARJETAS DE RESUMEN ESTADÍSTICO
        self.frame_resumen = tk.Frame(contenedor_principal, bg=self.colores["fondo"])
        self.frame_resumen.pack(fill=tk.X, pady=(0, 8))

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
            self.frame_resumen.columnconfigure(i, weight=1, uniform="kpi")
            card = tk.Frame(self.frame_resumen, bg=self.colores["panel"], relief="solid", bd=1, padx=6, pady=4)
            card.grid(row=0, column=i, padx=3, sticky="nsew")

            lbl_kpi_tit = tk.Label(card, text=label, font=("Helvetica", 9), bg=self.colores["panel"], fg=self.colores["texto_mutado"])
            lbl_kpi_tit.pack(anchor="center")

            lbl_kpi_val = tk.Label(card, text=val_def, font=("Helvetica", 14, "bold"), bg=self.colores["panel"], fg=color)
            lbl_kpi_val.pack(anchor="center")

            self.kpis[key] = lbl_kpi_val

        # 5. BARRA DE FILTRO Y BÚSQUEDA DE RESULTADOS
        frame_filtro = tk.Frame(contenedor_principal, bg=self.colores["fondo"])
        frame_filtro.pack(fill=tk.X, pady=(0, 6))

        lbl_filtro = tk.Label(frame_filtro, text="Filtrar por categoría:", font=("Helvetica", 10, "bold"), bg=self.colores["fondo"])
        lbl_filtro.pack(side=tk.LEFT, padx=(2, 6))

        self.combo_filtro = ttk.Combobox(
            frame_filtro,
            values=["Todas las categorías", "Coincidencia Literal", "Similitud Alta", "Relación Temática", "Sin Relación"],
            state="readonly",
            width=22
        )
        self.combo_filtro.current(0)
        self.combo_filtro.pack(side=tk.LEFT, padx=(0, 15))
        self.combo_filtro.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        lbl_buscar = tk.Label(frame_filtro, text="Buscar tema:", font=("Helvetica", 10, "bold"), bg=self.colores["fondo"])
        lbl_buscar.pack(side=tk.LEFT, padx=(0, 6))

        self.entry_buscar = tk.Entry(frame_filtro, font=("Helvetica", 10), width=26)
        self.entry_buscar.pack(side=tk.LEFT, padx=(0, 6))
        self.entry_buscar.bind("<KeyRelease>", lambda e: self._aplicar_filtro())

        btn_limpiar_filtro = tk.Button(
            frame_filtro,
            text="Limpiar filtro",
            font=("Helvetica", 9),
            bg="#E2E8F0",
            relief="flat",
            padx=8,
            pady=2,
            command=self._limpiar_filtro
        )
        btn_limpiar_filtro.pack(side=tk.LEFT)

        # 6. TABLA DE RESULTADOS (Treeview con scrollbars)
        frame_tabla = tk.Frame(contenedor_principal, bg=self.colores["panel"], relief="solid", bd=1)
        frame_tabla.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        columnas = (
            "doc1_mod",
            "doc1_tema",
            "doc2_mod",
            "doc2_tema",
            "similitud",
            "categoria",
            "recomendacion"
        )

        self.tabla_resultados = ttk.Treeview(
            frame_tabla,
            columns=columnas,
            show="headings",
            selectmode="browse"
        )

        # Encabezados de columnas según especificación
        self.tabla_resultados.heading("doc1_mod", text="Módulo Doc 1", anchor="w")
        self.tabla_resultados.heading("doc1_tema", text="Tema Doc 1", anchor="w")
        self.tabla_resultados.heading("doc2_mod", text="Módulo Doc 2", anchor="w")
        self.tabla_resultados.heading("doc2_tema", text="Tema Doc 2", anchor="w")
        self.tabla_resultados.heading("similitud", text="Similitud", anchor="center")
        self.tabla_resultados.heading("categoria", text="Clasificación", anchor="center")
        self.tabla_resultados.heading("recomendacion", text="Recomendación", anchor="w")

        # Anchos de columna proporcionados
        self.tabla_resultados.column("doc1_mod", width=120, minwidth=90)
        self.tabla_resultados.column("doc1_tema", width=220, minwidth=140)
        self.tabla_resultados.column("doc2_mod", width=120, minwidth=90)
        self.tabla_resultados.column("doc2_tema", width=220, minwidth=140)
        self.tabla_resultados.column("similitud", width=80, minwidth=70, anchor="center")
        self.tabla_resultados.column("categoria", width=160, minwidth=130, anchor="center")
        self.tabla_resultados.column("recomendacion", width=240, minwidth=180)

        scroll_y = ttk.Scrollbar(frame_tabla, orient=tk.VERTICAL, command=self.tabla_resultados.yview)
        scroll_x = ttk.Scrollbar(frame_tabla, orient=tk.HORIZONTAL, command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)

        # Evento de selección en la tabla para mostrar detalle explicativo
        self.tabla_resultados.bind("<<TreeviewSelect>>", self._al_seleccionar_fila)

        # 7. PANEL DE DETALLE EXPLICATIVO
        self.frame_detalle = tk.Frame(contenedor_principal, bg="#F1F5F9", relief="solid", bd=1, padx=12, pady=6)
        self.frame_detalle.pack(fill=tk.X, pady=(0, 8))

        self.lbl_detalle_titulo = tk.Label(
            self.frame_detalle,
            text="Detalle de coincidencia: Seleccione una fila de la tabla para ver la justificación de similitud.",
            font=("Helvetica", 10, "italic"),
            bg="#F1F5F9",
            fg=self.colores["texto_oscuro"],
            wraplength=1100,
            justify="left"
        )
        self.lbl_detalle_titulo.pack(anchor="w")

        # 8. BARRA DE BOTONES INFERIOR (Nuevo Análisis + Exportar)
        frame_inferior = tk.Frame(contenedor_principal, bg=self.colores["fondo"])
        frame_inferior.pack(fill=tk.X)

        self.btn_nuevo = tk.Button(
            frame_inferior,
            text="🔄  NUEVO ANÁLISIS",
            font=("Helvetica", 11, "bold"),
            bg="#64748B",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=7,
            command=self.reiniciar_analisis
        )
        self.btn_nuevo.pack(side=tk.LEFT)

        self.btn_exportar = tk.Button(
            frame_inferior,
            text="📥  EXPORTAR RESULTADOS (CSV / TXT)",
            font=("Helvetica", 11, "bold"),
            bg="#059669",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=7,
            command=self.solicitar_exportacion
        )
        self.btn_exportar.pack(side=tk.RIGHT)

    def _crear_panel_documento(self, parent, col, num_doc, titulo, sub):
        """Crea un panel visual estandarizado para la carga de cada documento."""
        panel = tk.Frame(parent, bg=self.colores["panel"], relief="solid", bd=1, padx=12, pady=10)
        panel.grid(row=0, column=col, padx=4, sticky="nsew")

        lbl_tit = tk.Label(panel, text=titulo, font=("Helvetica", 11, "bold"), bg=self.colores["panel"], fg=self.colores["texto_oscuro"])
        lbl_tit.pack(anchor="w")

        lbl_sub = tk.Label(panel, text=sub, font=("Helvetica", 9), bg=self.colores["panel"], fg=self.colores["texto_mutado"])
        lbl_sub.pack(anchor="w", pady=(0, 6))

        # Fila de archivo seleccionado
        lbl_arch = tk.Label(panel, text="Archivo: Ningún archivo seleccionado", font=("Helvetica", 10), bg=self.colores["panel"], fg=self.colores["texto_oscuro"])
        lbl_arch.pack(anchor="w")

        # Estado y formato
        lbl_est = tk.Label(panel, text="Estado: Pendiente de carga", font=("Helvetica", 9, "bold"), bg=self.colores["panel"], fg=self.colores["advertencia"])
        lbl_est.pack(anchor="w", pady=(2, 6))

        # Botón para cargar / reemplazar
        btn_sel = tk.Button(
            panel,
            text=f"Seleccionar Documento {num_doc}...",
            font=("Helvetica", 10),
            bg="#E2E8F0",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
            command=lambda: self.seleccionar_documento(num_doc)
        )
        btn_sel.pack(anchor="w")

        # Guardar referencias
        panel.lbl_archivo = lbl_arch
        panel.lbl_estado = lbl_est
        panel.btn_seleccionar = btn_sel

        return panel

    def _iniciar_ciclo_control(self):
        """
        Inicia el ciclo de despacho cooperativo que atiende la cola
        mediante la estructura WHILE obligatoria sin congelar Tkinter.
        """
        self._procesar_cola_control()

    def _procesar_cola_control(self):
        """
        ESTRUCTURA OBLIGATORIA: while
        Procesa de forma continua las solicitudes de análisis, reinicios
        y eventos de la aplicación sin bloquear el ciclo de Tkinter.
        """
        # ======================================================================
        # CICLO 'WHILE' OBLIGATORIO DE CONTROL DE ESTADO Y SOLICITUDES
        # ======================================================================
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

        # Reprogramar la verificación periódica (100 ms) para mantener la reactividad
        if self.app_activa:
            self.root.after(100, self._procesar_cola_control)

    def seleccionar_documento(self, num_doc):
        """
        Abre el explorador de archivos para que el usuario seleccione un documento
        Word (.docx) y ejecuta la carga y extracción preliminar.
        """
        formatos = [
            ("Documentos Word (*.docx)", "*.docx"),
            ("Todos los archivos (*.*)", "*.*")
        ]

        ruta = filedialog.askopenfilename(
            title=f"Seleccionar Documento {num_doc} (Programa Académico)",
            filetypes=formatos
        )

        if not ruta:
            return  # Usuario canceló el diálogo

        panel = self.panel_doc1 if num_doc == 1 else self.panel_doc2
        nombre_archivo = os.path.basename(ruta)
        formato = os.path.splitext(ruta)[1].lower()

        try:
            # 1. Cargar el contenido utilizando la función obligatoria
            lineas = cargar_documento(ruta)

            # 2. Extraer exclusivamente los temas del documento
            temas = extraer_temas(lineas, nombre_doc=f"Documento {num_doc}")

            # 3. Validar si se identificaron temas
            if not temas:
                panel.lbl_archivo.config(text=f"Archivo: {nombre_archivo} ({formato})")
                panel.lbl_estado.config(
                    text="Aviso: No se identificaron títulos de temas con la estructura actual.",
                    fg=self.colores["peligro"]
                )
                messagebox.showwarning(
                    "Estructura no identificada",
                    f"El archivo '{nombre_archivo}' fue leído correctamente pero no se "
                    f"pudieron identificar títulos de temas con certeza.\n\n"
                    f"El programa evita presentar encabezados generales como si fueran temas.\n"
                    f"Verifique que el documento contenga temas o módulos definidos."
                )
                return

            # Asignar a las variables de estado
            if num_doc == 1:
                self.doc1_ruta = ruta
                self.doc1_nombre = nombre_archivo
                self.doc1_formato = formato
                self.doc1_datos = temas
                self.kpis["temas_doc1"].config(text=str(len(temas)))
            else:
                self.doc2_ruta = ruta
                self.doc2_nombre = nombre_archivo
                self.doc2_formato = formato
                self.doc2_datos = temas
                self.kpis["temas_doc2"].config(text=str(len(temas)))

            # Actualizar panel visual
            panel.lbl_archivo.config(text=f"Archivo: {nombre_archivo} ({formato})")
            panel.lbl_estado.config(
                text=f"Cargado con éxito: {len(temas)} tema(s) identificados",
                fg=self.colores["exito"]
            )
            panel.btn_seleccionar.config(text=f"Reemplazar Documento {num_doc}...")

            # Comprobar si ambos documentos están listos
            if self.doc1_datos and self.doc2_datos:
                self.lbl_estado.config(
                    text=f"Listos para analizar: Doc 1 ({len(self.doc1_datos)} temas) y Doc 2 ({len(self.doc2_datos)} temas).",
                    fg=self.colores["primario"]
                )
            else:
                self.lbl_estado.config(
                    text=f"Documento {num_doc} cargado. Seleccione el otro documento para habilitar el análisis.",
                    fg=self.colores["texto_mutado"]
                )

        except Exception as e:
            panel.lbl_estado.config(text=f"Error al leer: {str(e)[:45]}...", fg=self.colores["peligro"])
            messagebox.showerror(
                "Error en Documento",
                f"Ocurrió un inconveniente al procesar el archivo:\n\n{nombre_archivo}\n\nDetalle: {str(e)}"
            )

    def solicitar_analisis(self):
        """
        Valida que ambos documentos estén cargados y encola la solicitud
        para que sea procesada en el ciclo WHILE de control.
        """
        # Validación obligatoria mediante 'if'
        if not self.doc1_datos or not self.doc2_datos:
            messagebox.showinfo(
                "Documentos incompletos",
                "Debe seleccionar y cargar satisfactoriamente ambos documentos antes de ejecutar el análisis."
            )
            return

        if self.analisis_en_curso:
            messagebox.showinfo("Procesando", "Ya hay un análisis en curso. Por favor espere.")
            return

        self.analisis_en_curso = True
        self.btn_analizar.config(state=tk.DISABLED, text="⏳  ANALIZANDO TEMAS...")
        self.lbl_estado.config(text="Comparando títulos de temas entre ambos documentos...", fg=self.colores["primario"])

        # Encolar la solicitud de análisis para el procesador WHILE
        solicitud = {
            "temas_doc1": self.doc1_datos,
            "temas_doc2": self.doc2_datos
        }
        self.cola_solicitudes.put(("EJECUTAR_ANALISIS", solicitud))

    def _ejecutar_analisis_en_hilo(self, datos):
        """
        Ejecuta la comparación en un hilo de fondo para no bloquear la GUI
        y devuelve el resultado a la cola.
        """
        def tarea():
            try:
                # Llamada a la función obligatoria con ciclos 'for'
                comparaciones, resumen, sin_match = comparar_temas(
                    datos["temas_doc1"],
                    datos["temas_doc2"]
                )
                self.cola_solicitudes.put(("FINALIZAR_ANALISIS", (comparaciones, resumen, sin_match)))
            except Exception as err:
                self.cola_solicitudes.put(("ERROR_ANALISIS", str(err)))

        hilo = threading.Thread(target=tarea, daemon=True)
        hilo.start()

    def _mostrar_resultados_en_ui(self, datos):
        """Actualiza la interfaz con los resultados de la comparación recién procesada."""
        self.analisis_en_curso = False
        self.btn_analizar.config(state=tk.NORMAL, text="🔍  ANALIZAR DOCUMENTOS")

        comparaciones, resumen, sin_match = datos
        self.resultados_actuales = comparaciones
        self.resumen_actual = resumen
        self.temas_sin_match = sin_match

        # Actualizar indicadores KPI
        self.kpis["literales"].config(text=str(resumen["coincidencias_literales"]))
        self.kpis["altas"].config(text=str(resumen["similitudes_altas"]))
        self.kpis["tematicas"].config(text=str(resumen["relaciones_tematicas"]))
        self.kpis["sin_match"].config(text=str(resumen["sin_coincidencias"]))

        self.lbl_estado.config(
            text=f"Análisis completado: {len(comparaciones)} comparaciones evaluadas. Mostrando resultados.",
            fg=self.colores["exito"]
        )

        # Poblar la tabla de resultados
        self._aplicar_filtro()

        messagebox.showinfo(
            "Análisis Completado",
            f"Se ha completado el análisis comparativo con éxito.\n\n"
            f"• Temas en Documento 1: {resumen['total_temas_doc1']}\n"
            f"• Temas en Documento 2: {resumen['total_temas_doc2']}\n"
            f"• Coincidencias literales: {resumen['coincidencias_literales']}\n"
            f"• Similitudes altas: {resumen['similitudes_altas']}\n"
            f"• Relaciones temáticas: {resumen['relaciones_tematicas']}\n"
            f"• Temas sin coincidencias: {resumen['sin_coincidencias']}\n\n"
            f"Consulte la tabla comparativa para examinar las recomendaciones orientativas."
        )

    def _manejar_error_analisis(self, mensaje_error):
        """Gestiona cualquier excepción imprevista durante el análisis."""
        self.analisis_en_curso = False
        self.btn_analizar.config(state=tk.NORMAL, text="🔍  ANALIZAR DOCUMENTOS")
        self.lbl_estado.config(text=f"Error durante el análisis: {mensaje_error}", fg=self.colores["peligro"])
        messagebox.showerror("Error en Análisis", f"Ocurrió un error al comparar los documentos:\n\n{mensaje_error}")

    def _aplicar_filtro(self):
        """Filtra y muestra las filas en la tabla Treeview según la categoría o búsqueda textual."""
        # Limpiar elementos actuales del Treeview
        for item in self.tabla_resultados.get_children():
            self.tabla_resultados.delete(item)

        categoria_sel = self.combo_filtro.get()
        busqueda = self.entry_buscar.get().lower().strip()

        # Configurar tags de colores para categorías
        self.tabla_resultados.tag_configure("tag_literal", foreground="#047857", background="#ECFDF5")
        self.tabla_resultados.tag_configure("tag_alta", foreground="#1D4ED8", background="#EFF6FF")
        self.tabla_resultados.tag_configure("tag_tematica", foreground="#B45309", background="#FFFBEB")
        self.tabla_resultados.tag_configure("tag_sin_relacion", foreground="#64748B", background="#F8FAFC")

        for r in self.resultados_actuales:
            # Filtro por categoría
            if categoria_sel == "Coincidencia Literal" and r["categoria"] != "COINCIDENCIA LITERAL":
                continue
            elif categoria_sel == "Similitud Alta" and r["categoria"] != "SIMILITUD ALTA":
                continue
            elif categoria_sel == "Relación Temática" and r["categoria"] != "RELACIÓN TEMÁTICA O PARALELA":
                continue
            elif categoria_sel == "Sin Relación" and r["categoria"] != "SIN RELACIÓN IDENTIFICADA":
                continue

            # Filtro por búsqueda de texto
            if busqueda:
                t1 = r["doc1_tema_titulo"].lower()
                t2 = r["doc2_tema_titulo"].lower()
                if busqueda not in t1 and busqueda not in t2:
                    continue

            # Asignar tag visual según categoría
            if r["nivel"] == 1:
                tag = "tag_literal"
            elif r["nivel"] == 2:
                tag = "tag_alta"
            elif r["nivel"] == 3:
                tag = "tag_tematica"
            else:
                tag = "tag_sin_relacion"

            val_similitud = f"{r['similitud']}%"

            self.tabla_resultados.insert(
                "",
                tk.END,
                values=(
                    r["doc1_modulo"],
                    r["doc1_tema_titulo"],
                    r["doc2_modulo"],
                    r["doc2_tema_titulo"],
                    val_similitud,
                    r["categoria"],
                    r["recomendacion"]
                ),
                tags=(tag,)
            )

    def _limpiar_filtro(self):
        """Restablece los campos de filtro a sus valores predeterminados."""
        self.combo_filtro.current(0)
        self.entry_buscar.delete(0, tk.END)
        self._aplicar_filtro()

    def _al_seleccionar_fila(self, event):
        """Muestra los detalles completos y la justificación textual del tema seleccionado."""
        seleccion = self.tabla_resultados.selection()
        if not seleccion:
            return

        item = self.tabla_resultados.item(seleccion[0])
        valores = item["values"]
        if not valores:
            return

        t1_titulo = valores[1]
        t2_titulo = valores[3]
        similitud = valores[4]
        categoria = valores[5]
        recomendacion = valores[6]

        # Buscar el registro original para obtener la explicación detallada
        explicacion = "Información de coincidencia basada en términos compartidos."
        for r in self.resultados_actuales:
            if r["doc1_tema_titulo"] == t1_titulo and r["doc2_tema_titulo"] == t2_titulo:
                explicacion = r["explicacion"]
                break

        texto_detalle = (
            f"📌 Doc 1: {t1_titulo}   ⟷   📌 Doc 2: {t2_titulo}\n"
            f"• Similitud: {similitud} ({categoria})   |   • Recomendación: {recomendacion}\n"
            f"• Explicación técnica: {explicacion}"
        )

        self.lbl_detalle_titulo.config(
            text=texto_detalle,
            font=("Helvetica", 10, "normal")
        )

    def reiniciar_analisis(self):
        """
        Función obligatoria para reiniciar el proceso de análisis y permitir
        realizar nuevos análisis sin cerrar la aplicación.
        
        ESTRUCTURA OBLIGATORIA: while
        """
        # ======================================================================
        # CICLO 'WHILE' OBLIGATORIO PARA VACIAR RESULTADOS Y RESETEAR ESTADO
        # ======================================================================
        while len(self.resultados_actuales) > 0:
            self.resultados_actuales.pop()

        while len(self.temas_sin_match) > 0:
            self.temas_sin_match.pop()

        self.resumen_actual.clear()

        # Limpiar elementos del Treeview
        items_tabla = self.tabla_resultados.get_children()
        idx = 0
        while idx < len(items_tabla):
            self.tabla_resultados.delete(items_tabla[idx])
            idx += 1

        # Restablecer indicadores visuales
        self.kpis["literales"].config(text="0")
        self.kpis["altas"].config(text="0")
        self.kpis["tematicas"].config(text="0")
        self.kpis["sin_match"].config(text="0")

        self.entry_buscar.delete(0, tk.END)
        self.combo_filtro.current(0)
        self.lbl_detalle_titulo.config(
            text="Detalle de coincidencia: Seleccione una fila de la tabla para ver la justificación de similitud.",
            font=("Helvetica", 10, "italic")
        )

        if self.doc1_datos and self.doc2_datos:
            self.lbl_estado.config(
                text="Resultados reiniciados. Puede volver a pulsar 'ANALIZAR DOCUMENTOS' o reemplazar los archivos.",
                fg=self.colores["primario"]
            )
        else:
            self.lbl_estado.config(
                text="Análisis reiniciado. Cargue los documentos para iniciar una nueva comparación.",
                fg=self.colores["texto_mutado"]
            )

    def solicitar_exportacion(self):
        """Permite exportar el informe comparativo a archivo CSV o TXT."""
        if not self.resultados_actuales:
            messagebox.showinfo(
                "Sin resultados",
                "No hay resultados de análisis disponibles para exportar. Ejecute primero el análisis."
            )
            return

        ruta_guardado = filedialog.asksaveasfilename(
            title="Guardar Informe de Análisis Comparativo",
            defaultextension=".csv",
            filetypes=[
                ("Archivo CSV (*.csv)", "*.csv"),
                ("Archivo de Texto (*.txt)", "*.txt")
            ]
        )

        if not ruta_guardado:
            return  # Cancelado por el usuario

        try:
            info_doc1 = {"nombre": self.doc1_nombre, "ruta": self.doc1_ruta}
            info_doc2 = {"nombre": self.doc2_nombre, "ruta": self.doc2_ruta}

            formato = "txt" if ruta_guardado.lower().endswith(".txt") else "csv"

            # Llamada a la función obligatoria
            generar_informe(
                self.resultados_actuales,
                self.resumen_actual,
                info_doc1,
                info_doc2,
                ruta_guardado,
                formato=formato
            )

            messagebox.showinfo(
                "Exportación Exitosa",
                f"El informe comparativo ha sido generado y guardado correctamente en:\n\n{ruta_guardado}"
            )

        except Exception as err:
            messagebox.showerror(
                "Error en Exportación",
                f"No se pudo guardar el informe:\n\n{str(err)}"
            )

    def cerrar_aplicacion(self):
        """Finaliza el ciclo de vida de la aplicación de manera ordenada."""
        self.app_activa = False
        self.root.destroy()
