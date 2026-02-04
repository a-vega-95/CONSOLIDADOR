import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import ConsolidationService, DOCX2PDF_AVAILABLE

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False


class ConsolidadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consolidador de Documentos - Arrastrar y Soltar")
        self.root.geometry("760x720")
        self.root.resizable(True, True)

        self.archivos_seleccionados = []
        self.carpeta_entrada = None
        self.agregar_pagina_blanca = tk.BooleanVar(value=False)
        self.normalizar_a_pdf = tk.BooleanVar(value=False)
        self.formato_salida = tk.StringVar(value="DOCX")

        self.service = ConsolidationService()

        main_frame = ttk.Frame(root, padding="12")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        titulo = ttk.Label(main_frame, text="Consolidador de Documentos (Word y PDF)",
                           font=("Arial", 14, "bold"))
        titulo.grid(row=0, column=0, columnspan=3, pady=10)

        instrucciones = ttk.Label(
            main_frame,
            text="✨ Arrastra archivos PDF/DOCX aquí o usa los botones ✨",
            font=("Arial", 10, "italic"),
            foreground="blue",
        )
        instrucciones.grid(row=1, column=0, columnspan=3, pady=5)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))

        self.btn_agregar = ttk.Button(btn_frame, text="➕ Agregar Archivos",
                                       command=self.agregar_archivos)
        self.btn_agregar.pack(side=tk.LEFT, padx=5)

        self.btn_carpeta = ttk.Button(btn_frame, text="📁 Cargar Carpeta",
                                       command=self.seleccionar_carpeta)
        self.btn_carpeta.pack(side=tk.LEFT, padx=5)

        self.btn_limpiar = ttk.Button(btn_frame, text="🗑️ Limpiar Lista",
                                      command=self.limpiar_lista)
        self.btn_limpiar.pack(side=tk.LEFT, padx=5)

        ttk.Label(main_frame, text="Archivos en orden de consolidación:",
                  font=("Arial", 10, "bold")).grid(row=3, column=0,
                                                  columnspan=3, pady=(10, 5))

        listbox_frame = ttk.Frame(main_frame)
        listbox_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        scroll_frame = ttk.Frame(listbox_frame)
        scroll_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(scroll_frame, height=12, selectmode=tk.SINGLE,
                                  bg="#f5f5f5", font=("Consolas", 9),
                                  activestyle='dotbox')
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        if DRAG_DROP_AVAILABLE:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind('<<Drop>>', self.drop_archivos)

        orden_frame = ttk.Frame(listbox_frame)
        orden_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10)

        ttk.Label(orden_frame, text="Ordenar:", font=("Arial", 9, "bold")).pack(pady=(0, 5))
        ttk.Button(orden_frame, text="⬆️ Subir", command=self.subir_archivo, width=12).pack(pady=3)
        ttk.Button(orden_frame, text="⬇️ Bajar", command=self.bajar_archivo, width=12).pack(pady=3)
        ttk.Button(orden_frame, text="❌ Quitar", command=self.quitar_archivo, width=12).pack(pady=20)

        salida_frame = ttk.Frame(main_frame)
        salida_frame.grid(row=5, column=0, columnspan=3, pady=(15, 5), sticky=(tk.W, tk.E))

        ttk.Label(salida_frame, text="Nombre base:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)

        self.entry_nombre = ttk.Entry(salida_frame, width=35, font=("Arial", 10))
        self.entry_nombre.insert(0, "consolidado")
        self.entry_nombre.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        checkbox_frame = ttk.Frame(main_frame)
        checkbox_frame.grid(row=6, column=0, columnspan=3, pady=(5, 5), sticky=(tk.W))

        self.check_pagina = ttk.Checkbutton(
            checkbox_frame,
            text="📄 Agregar página en blanco entre cada archivo",
            variable=self.agregar_pagina_blanca,
            onvalue=True,
            offvalue=False,
        )
        self.check_pagina.pack(side=tk.LEFT, padx=5)

        checkbox_pdf_frame = ttk.Frame(main_frame)
        checkbox_pdf_frame.grid(row=7, column=0, columnspan=3, pady=(5, 5), sticky=(tk.W))

        self.check_normalizar = ttk.Checkbutton(
            checkbox_pdf_frame,
            text="🔄 Normalizar todo a PDF antes de consolidar (DOCX → PDF → DOCX final)",
            variable=self.normalizar_a_pdf,
            onvalue=True,
            offvalue=False,
            state='normal' if DOCX2PDF_AVAILABLE else 'disabled',
        )
        self.check_normalizar.pack(side=tk.LEFT, padx=5)

        if not DOCX2PDF_AVAILABLE:
            ttk.Label(
                checkbox_pdf_frame,
                text="⚠️ Instala docx2pdf",
                foreground="orange",
                font=("Arial", 8),
            ).pack(side=tk.LEFT, padx=5)

        formato_frame = ttk.Frame(main_frame)
        formato_frame.grid(row=8, column=0, columnspan=3, pady=(5, 5), sticky=(tk.W))

        ttk.Label(formato_frame, text="Formato de salida:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.combo_formato = ttk.Combobox(
            formato_frame,
            textvariable=self.formato_salida,
            values=["DOCX", "PDF", "DOCX+PDF"],
            state="readonly",
            width=12,
        )
        self.combo_formato.pack(side=tk.LEFT, padx=5)
        self.combo_formato.current(0)

        if not DOCX2PDF_AVAILABLE:
            ttk.Label(
                formato_frame,
                text="⚠️ PDF requiere docx2pdf",
                foreground="orange",
                font=("Arial", 8),
            ).pack(side=tk.LEFT, padx=5)

        ruta_frame = ttk.Frame(main_frame)
        ruta_frame.grid(row=9, column=0, columnspan=3, pady=(5, 5), sticky=(tk.W, tk.E))

        ttk.Label(ruta_frame, text="Guardar en:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)

        self.entry_ruta = ttk.Entry(ruta_frame, width=40, font=("Arial", 9), state='readonly')
        self.entry_ruta.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        ttk.Button(ruta_frame, text="📂 Cambiar",
                   command=self.seleccionar_ruta_salida, width=15).pack(side=tk.LEFT, padx=5)

        self.btn_procesar = ttk.Button(main_frame, text="🚀 CONSOLIDAR DOCUMENTOS",
                                       command=self.procesar_documentos)
        self.btn_procesar.grid(row=10, column=0, columnspan=3, pady=8, sticky=(tk.W, tk.E))

        self.btn_solo_pdf = ttk.Button(main_frame, text="📄 CONSOLIDAR A PDF",
                                       command=self.procesar_solo_pdf)
        self.btn_solo_pdf.grid(row=11, column=0, columnspan=3, pady=8, sticky=(tk.W, tk.E))

        self.progress = ttk.Progressbar(main_frame, length=400, mode='indeterminate')
        self.progress.grid(row=12, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.label_estado = ttk.Label(main_frame, text="Estado: Esperando archivos...",
                                      foreground="blue", font=("Arial", 9))
        self.label_estado.grid(row=13, column=0, columnspan=3, sticky=(tk.W, tk.E))

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

    def agregar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos PDF o DOCX",
            filetypes=[
                ("Documentos", "*.pdf *.docx *.doc"),
                ("PDF", "*.pdf"),
                ("Word", "*.docx *.doc"),
                ("Todos", "*.*"),
            ],
        )

        if archivos:
            for archivo in archivos:
                if archivo not in self.archivos_seleccionados:
                    self.archivos_seleccionados.append(archivo)

            self.actualizar_listbox()
            self.label_estado.config(
                text=f"Estado: {len(self.archivos_seleccionados)} archivo(s) cargado(s)",
                foreground="green",
            )

    def drop_archivos(self, event):
        archivos = self.root.tk.splitlist(event.data)

        for archivo in archivos:
            archivo = archivo.strip("{}")

            if os.path.isfile(archivo):
                ext = archivo.lower()
                if ext.endswith((".pdf", ".docx", ".doc")):
                    if archivo not in self.archivos_seleccionados:
                        self.archivos_seleccionados.append(archivo)

        self.actualizar_listbox()

        if self.archivos_seleccionados:
            self.label_estado.config(
                text=f"Estado: {len(self.archivos_seleccionados)} archivo(s) cargado(s)",
                foreground="green",
            )

    def limpiar_lista(self):
        if self.archivos_seleccionados:
            if messagebox.askyesno("Confirmar", "¿Limpiar toda la lista de archivos?"):
                self.archivos_seleccionados = []
                self.listbox.delete(0, tk.END)
                self.label_estado.config(text="Estado: Lista limpiada", foreground="blue")
        else:
            messagebox.showinfo("Info", "La lista ya está vacía")

    def subir_archivo(self):
        seleccion = self.listbox.curselection()
        if not seleccion:
            messagebox.showinfo("Info", "Selecciona un archivo de la lista primero")
            return

        idx = seleccion[0]
        if idx > 0:
            self.archivos_seleccionados[idx], self.archivos_seleccionados[idx - 1] = (
                self.archivos_seleccionados[idx - 1],
                self.archivos_seleccionados[idx],
            )

            self.actualizar_listbox()
            self.listbox.selection_set(idx - 1)
            self.listbox.see(idx - 1)

    def bajar_archivo(self):
        seleccion = self.listbox.curselection()
        if not seleccion:
            messagebox.showinfo("Info", "Selecciona un archivo de la lista primero")
            return

        idx = seleccion[0]
        if idx < len(self.archivos_seleccionados) - 1:
            self.archivos_seleccionados[idx], self.archivos_seleccionados[idx + 1] = (
                self.archivos_seleccionados[idx + 1],
                self.archivos_seleccionados[idx],
            )

            self.actualizar_listbox()
            self.listbox.selection_set(idx + 1)
            self.listbox.see(idx + 1)

    def quitar_archivo(self):
        seleccion = self.listbox.curselection()
        if not seleccion:
            messagebox.showinfo("Info", "Selecciona un archivo de la lista primero")
            return

        idx = seleccion[0]
        nombre = os.path.basename(self.archivos_seleccionados[idx])

        if messagebox.askyesno("Confirmar", f"¿Quitar '{nombre}' de la lista?"):
            del self.archivos_seleccionados[idx]
            self.actualizar_listbox()

            self.label_estado.config(
                text=f"Estado: {len(self.archivos_seleccionados)} archivo(s) en lista",
                foreground="green",
            )

    def actualizar_listbox(self):
        self.listbox.delete(0, tk.END)
        for idx, archivo in enumerate(self.archivos_seleccionados, 1):
            nombre = os.path.basename(archivo)
            ext = os.path.splitext(nombre)[1].upper()
            self.listbox.insert(tk.END, f"{idx}. [{ext}] {nombre}")

        if not self.carpeta_entrada and self.archivos_seleccionados:
            self.carpeta_entrada = os.path.dirname(self.archivos_seleccionados[0])
            self.entry_ruta.config(state='normal')
            self.entry_ruta.delete(0, tk.END)
            self.entry_ruta.insert(0, self.carpeta_entrada)
            self.entry_ruta.config(state='readonly')

    def seleccionar_ruta_salida(self):
        carpeta = filedialog.askdirectory(
            title="Seleccionar carpeta para guardar el consolidado",
            initialdir=self.carpeta_entrada if self.carpeta_entrada else os.path.expanduser("~"),
        )

        if carpeta:
            self.carpeta_entrada = carpeta
            self.entry_ruta.config(state='normal')
            self.entry_ruta.delete(0, tk.END)
            self.entry_ruta.insert(0, carpeta)
            self.entry_ruta.config(state='readonly')
            self.label_estado.config(text=f"Estado: Ruta de guardado actualizada", foreground="green")

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta con documentos")

        if carpeta:
            self.carpeta_entrada = carpeta
            self.entry_ruta.config(state='normal')
            self.entry_ruta.delete(0, tk.END)
            self.entry_ruta.insert(0, carpeta)
            self.entry_ruta.config(state='readonly')
            self.listar_archivos(carpeta)

    def listar_archivos(self, carpeta):
        archivos = []
        for archivo in os.listdir(carpeta):
            if archivo.lower().endswith((".pdf", ".docx", ".doc")):
                ruta_completa = os.path.join(carpeta, archivo)
                archivos.append((archivo, ruta_completa))

        def extraer_numero(nombre_archivo):
            match = re.match(r"(\d+)", nombre_archivo[0])
            return int(match.group(1)) if match else float("inf")

        archivos.sort(key=extraer_numero)

        agregados = 0
        for archivo, ruta in archivos:
            if ruta not in self.archivos_seleccionados:
                self.archivos_seleccionados.append(ruta)
                agregados += 1

        self.actualizar_listbox()

        if agregados > 0:
            self.label_estado.config(
                text=f"Estado: Se agregaron {agregados} archivo(s). Total: {len(self.archivos_seleccionados)}",
                foreground="green",
            )
        else:
            self.label_estado.config(text="Estado: No se encontraron archivos nuevos en la carpeta", foreground="orange")

    def procesar_documentos(self):
        if not self.archivos_seleccionados:
            messagebox.showwarning(
                "Advertencia",
                "No hay archivos para procesar.\n\n"
                "Arrastra archivos a la lista o usa los botones para agregar.",
            )
            return

        nombre_base = self.entry_nombre.get().strip()
        if not nombre_base:
            nombre_base = "consolidado"

        formato = self.formato_salida.get()
        if formato == "PDF":
            self._procesar_como_pdf(nombre_base)
            return
        if formato == "DOCX+PDF" and not DOCX2PDF_AVAILABLE:
            messagebox.showwarning(
                "PDF no disponible",
                "docx2pdf no está instalado. Selecciona DOCX o instala docx2pdf.",
            )
            return

        nombre_salida = f"{nombre_base}.docx"

        if not self._validar_ruta_salida():
            return

        self._set_processing_state(True)

        thread = threading.Thread(
            target=self._procesar_en_background,
            args=(nombre_salida, formato),
        )
        thread.daemon = True
        thread.start()

    def procesar_solo_pdf(self):
        if not self.archivos_seleccionados:
            messagebox.showwarning(
                "Advertencia",
                "No hay archivos para procesar.\n\n"
                "Arrastra archivos a la lista o usa los botones para agregar.",
            )
            return

        nombre_base = self.entry_nombre.get().strip()
        if not nombre_base:
            nombre_base = "consolidado"

        self._procesar_como_pdf(nombre_base)

    def _procesar_como_pdf(self, nombre_base):
        if not DOCX2PDF_AVAILABLE:
            messagebox.showwarning(
                "PDF no disponible",
                "docx2pdf no está instalado. Instálalo para convertir Word a PDF.",
            )
            return

        if not self._validar_ruta_salida():
            return

        nombre_salida = f"{nombre_base}.pdf"
        self._set_processing_state(True)

        thread = threading.Thread(
            target=self._procesar_solo_pdf_en_background,
            args=(nombre_salida,),
        )
        thread.daemon = True
        thread.start()

    def _procesar_en_background(self, nombre_salida, formato):
        try:
            self.progress.start(10)
            self._set_status("Estado: Iniciando consolidación...", "blue")

            ruta_salida = os.path.join(self.carpeta_entrada, nombre_salida)

            resultado = self.service.consolidate_to_docx(
                self.archivos_seleccionados,
                ruta_salida,
                add_blank_page=self.agregar_pagina_blanca.get(),
                normalize_to_pdf=self.normalizar_a_pdf.get(),
                generate_pdf=(formato == "DOCX+PDF"),
                on_status=self._set_status,
                on_item_error=self._notify_item_error,
            )

            self.progress.stop()
            estado_texto = f"Estado: ✅ Consolidación completada: {os.path.basename(resultado['docx'])}"
            if resultado.get("pdf"):
                estado_texto += " (PDF generado)"

            self._set_status(estado_texto, "green")

            detalles = [
                "✅ Documento consolidado guardado exitosamente:",
                f"DOCX: {resultado['docx']}",
                f"Total de archivos procesados: {len(self.archivos_seleccionados)}",
                f"Tamaño DOCX: {os.path.getsize(resultado['docx']) / 1024:.2f} KB",
            ]
            if resultado.get("pdf") and os.path.exists(resultado["pdf"]):
                detalles.append(f"PDF: {resultado['pdf']}")

            messagebox.showinfo("Éxito", "\n".join(detalles))

        except Exception as e:
            self.progress.stop()
            self._set_status("Estado: ❌ Error crítico", "red")
            messagebox.showerror(
                "Error Crítico",
                f"❌ Error al procesar documentos:\n\n{str(e)}\n\n"
                f"Verifica que los archivos no estén dañados y que Word esté cerrado.",
            )
        finally:
            self._set_processing_state(False)

    def _procesar_solo_pdf_en_background(self, nombre_salida):
        try:
            self.progress.start(10)
            self._set_status("Estado: Iniciando consolidación a PDF...", "blue")

            ruta_salida = os.path.join(self.carpeta_entrada, nombre_salida)
            resultado = self.service.consolidate_to_pdf(
                self.archivos_seleccionados,
                ruta_salida,
                add_blank_page=self.agregar_pagina_blanca.get(),
                on_status=self._set_status,
                on_item_error=self._notify_item_error,
            )

            self.progress.stop()
            self._set_status(
                f"Estado: ✅ Consolidación PDF completada: {os.path.basename(resultado)}",
                "green",
            )

            detalles = [
                "✅ PDF consolidado guardado exitosamente:",
                f"PDF: {resultado}",
                f"Total de archivos procesados: {len(self.archivos_seleccionados)}",
                f"Tamaño PDF: {os.path.getsize(resultado) / 1024:.2f} KB",
            ]
            messagebox.showinfo("Éxito", "\n".join(detalles))

        except Exception as e:
            self.progress.stop()
            self._set_status("Estado: ❌ Error crítico", "red")
            messagebox.showerror(
                "Error Crítico",
                f"❌ Error al procesar PDF:\n\n{str(e)}\n\n"
                f"Verifica que los archivos no estén dañados y que Word esté cerrado.",
            )
        finally:
            self._set_processing_state(False)

    def _validar_ruta_salida(self):
        if not self.carpeta_entrada and self.archivos_seleccionados:
            self.carpeta_entrada = os.path.dirname(self.archivos_seleccionados[0])
            self.entry_ruta.config(state='normal')
            self.entry_ruta.delete(0, tk.END)
            self.entry_ruta.insert(0, self.carpeta_entrada)
            self.entry_ruta.config(state='readonly')

        if not self.carpeta_entrada or not os.path.exists(self.carpeta_entrada):
            messagebox.showwarning(
                "Advertencia",
                "Por favor selecciona una carpeta válida para guardar el consolidado.",
            )
            return False
        return True

    def _set_processing_state(self, processing):
        state = 'disabled' if processing else 'normal'
        self.btn_procesar.config(state=state)
        self.btn_agregar.config(state=state)
        self.btn_carpeta.config(state=state)
        self.btn_solo_pdf.config(state=state)

    def _set_status(self, texto, color):
        self.label_estado.config(text=texto, foreground=color)
        self.root.update()

    def _notify_item_error(self, nombre_archivo, error_msg):
        self._set_status(f"Estado: ⚠️ Error en {nombre_archivo}: {error_msg}", "orange")


def run_app():
    if DRAG_DROP_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        messagebox.showinfo(
            "Información",
            "⚠️ La funcionalidad de arrastrar y soltar no está disponible.\n\n"
            "Para habilitarla, instala: pip install tkinterdnd2\n\n"
            "Puedes usar los botones para agregar archivos.",
        )

    app = ConsolidadorApp(root)
    root.mainloop()
