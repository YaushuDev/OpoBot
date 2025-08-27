# gui/main_window.py
"""
Ventana principal de la aplicación con diseño de 3 secciones.
Maneja la interfaz principal con gestión de perfiles, configuraciones y reportes.
"""

# Archivos relacionados: gui/email_modal.py, gui/profile_manager.py, services/excel_service.py

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import subprocess
import platform
from datetime import datetime
from pathlib import Path


class MainWindow:
    def __init__(self, root):
        """
        Inicializa la ventana principal con diseño de 3 secciones

        Args:
            root (tk.Tk): Ventana raíz de tkinter
        """
        self.root = root
        self.message_count = 0
        self.services_loaded = False

        # Inicializar servicios de manera segura
        self.init_services()

        # Configurar ventana
        self.setup_window()

        # Crear layout
        self.create_layout()

        # Verificar estado inicial
        self.check_initial_state()

    def init_services(self):
        """Inicializa los servicios de manera segura"""
        try:
            from gui.email_modal import EmailModal
            from gui.profile_manager import ProfileManager
            from services.excel_service import ExcelService
            from services.profile_service import ProfileService

            self.excel_service = ExcelService()
            self.profile_service = ProfileService()
            self.services_loaded = True

        except ImportError as e:
            self.services_loaded = False
            error_msg = f"Error cargando servicios: {e}"
            print(error_msg)
            messagebox.showerror("Error de Inicialización",
                                 f"{error_msg}\n\nVerifique que todos los archivos estén presentes.")

    def setup_window(self):
        """Configura las propiedades básicas de la ventana"""
        self.root.title("Bot Python - Gestión de Perfiles y Configuración SMTP")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        self.root.minsize(800, 600)

        # Configurar grid para que se expanda
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        # Configurar icono si existe
        try:
            icon_path = Path("icon.ico")
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            pass

    def create_layout(self):
        """Crea el layout de 3 secciones de la interfaz principal"""
        # Crear la barra de estado primero
        self.create_status_bar()

        # Frame principal con padding
        main_container = ttk.Frame(self.root, padding="10")
        main_container.grid(row=0, column=0, sticky="nsew")
        main_container.grid_rowconfigure(0, weight=0)
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        # Sección superior: Gestión de perfiles
        self.create_profiles_section(main_container)

        # Sección inferior: Frame para configuraciones e información
        bottom_frame = ttk.Frame(main_container)
        bottom_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        bottom_frame.grid_rowconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(1, weight=1)

        # Columna izquierda: Configuraciones
        self.create_left_column(bottom_frame)

        # Columna derecha: Información y reportes
        self.create_right_column(bottom_frame)

    def create_profiles_section(self, parent):
        """
        Crea la sección superior para gestión de perfiles

        Args:
            parent: Frame padre donde se colocará la sección
        """
        if not self.services_loaded:
            # Mostrar mensaje de error si los servicios no se cargaron
            error_frame = ttk.LabelFrame(parent, text="Error del Sistema", padding="10")
            error_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

            error_label = ttk.Label(error_frame,
                                    text="❌ Error: No se pudieron cargar los servicios del sistema.\n"
                                         "Verifique que todos los archivos estén presentes.",
                                    foreground="red", font=("Arial", 10))
            error_label.pack(pady=20)
            return

        try:
            from gui.profile_manager import ProfileManager
            self.profile_manager = ProfileManager(parent, self.update_info)
        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error cargando gestor de perfiles: {error_msg}", "error")

            # Crear frame de error
            error_frame = ttk.LabelFrame(parent, text="Error - Gestor de Perfiles", padding="10")
            error_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

            error_label = ttk.Label(error_frame, text=f"Error: {error_msg}", foreground="red")
            error_label.pack(pady=10)

    def create_left_column(self, parent):
        """
        Crea la columna izquierda con configuraciones y botones principales
        """
        left_frame = ttk.LabelFrame(parent, text="Panel de Control", padding="10")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        left_frame.grid_rowconfigure(3, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)

        # Título
        title_label = ttk.Label(left_frame, text="Configuraciones",
                                font=("Arial", 12, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 15), sticky="w")

        # Sección de configuración SMTP
        smtp_frame = ttk.LabelFrame(left_frame, text="Configuracion SMTP", padding="8")
        smtp_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        smtp_frame.grid_columnconfigure(0, weight=1)

        # Botón para configurar email
        self.email_button = ttk.Button(smtp_frame, text="⚙️ Configurar Email SMTP",
                                       command=self.open_email_modal)
        self.email_button.grid(row=0, column=0, sticky="ew", pady=2)

        # Estado de configuración SMTP
        self.smtp_status_label = ttk.Label(smtp_frame, text="Estado: No configurado",
                                           foreground="red", font=("Arial", 8))
        self.smtp_status_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # Sección de reportes
        reports_frame = ttk.LabelFrame(left_frame, text="Generacion de Reportes", padding="8")
        reports_frame.grid(row=2, column=0, sticky="ew", pady=10)
        reports_frame.grid_columnconfigure(0, weight=1)

        # Botones de reportes
        self.excel_button = ttk.Button(reports_frame, text="📊 Generar Reporte Excel",
                                       command=self.generate_excel_report)
        self.excel_button.grid(row=0, column=0, sticky="ew", pady=2)

        self.csv_button = ttk.Button(reports_frame, text="📋 Generar Reporte CSV",
                                     command=self.generate_csv_report)
        self.csv_button.grid(row=1, column=0, sticky="ew", pady=2)

        self.view_reports_button = ttk.Button(reports_frame, text="📁 Ver Reportes Generados",
                                              command=self.show_available_reports)
        self.view_reports_button.grid(row=2, column=0, sticky="ew", pady=2)

        # Sección de utilidades
        utils_frame = ttk.LabelFrame(left_frame, text="Utilidades", padding="8")
        utils_frame.grid(row=3, column=0, sticky="new", pady=(10, 0))
        utils_frame.grid_columnconfigure(0, weight=1)

        # Botón para limpiar log
        clear_log_button = ttk.Button(utils_frame, text="🗑️ Limpiar Log",
                                      command=self.clear_info_log)
        clear_log_button.grid(row=0, column=0, sticky="ew", pady=2)

    def create_right_column(self, parent):
        """
        Crea la columna derecha para información y estado del sistema
        """
        right_frame = ttk.LabelFrame(parent, text="Monitor de Actividad", padding="10")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        right_frame.grid_rowconfigure(1, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        # Título con contador de mensajes
        self.activity_title = ttk.Label(right_frame, text="Registro de Actividad",
                                        font=("Arial", 12, "bold"))
        self.activity_title.grid(row=0, column=0, pady=(0, 10), sticky="w")

        # Frame para el área de texto con scrollbar
        text_frame = ttk.Frame(right_frame)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)

        # Area de texto para información
        self.info_text = tk.Text(text_frame, height=20, width=50,
                                 state="disabled", wrap="word",
                                 font=("Consolas", 9), bg="#f8f9fa")
        self.info_text.grid(row=0, column=0, sticky="nsew")

        # Scrollbar para el área de texto
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical",
                                  command=self.info_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.info_text.configure(yscrollcommand=scrollbar.set)

        # Configurar estilos para diferentes tipos de mensajes
        self.info_text.tag_configure("info", foreground="#0066cc")
        self.info_text.tag_configure("success", foreground="#28a745")
        self.info_text.tag_configure("error", foreground="#dc3545")
        self.info_text.tag_configure("warning", foreground="#fd7e14")
        self.info_text.tag_configure("timestamp", foreground="#6c757d", font=("Consolas", 8))

    def create_status_bar(self):
        """Crea la barra de estado en la parte inferior"""
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))
        self.status_bar.grid_columnconfigure(1, weight=1)

        # Estado general
        self.status_label = ttk.Label(self.status_bar, text="Listo")
        self.status_label.grid(row=0, column=0, padx=(0, 10))

        # Separador
        ttk.Separator(self.status_bar, orient="vertical").grid(row=0, column=1, sticky="ns", padx=10)

        # Información adicional
        self.info_status = ttk.Label(self.status_bar, text="")
        self.info_status.grid(row=0, column=2)

    def check_initial_state(self):
        """Verifica el estado inicial del sistema"""
        try:
            # Verificar estado SMTP
            self.check_smtp_status()

            # Mostrar mensajes iniciales
            self.update_info("🚀 Sistema iniciado correctamente", "success")

            if self.services_loaded:
                self.update_info("📧 Configure su email SMTP y cree perfiles de búsqueda para comenzar", "info")

                # Verificar si hay perfiles
                if hasattr(self, 'profile_service'):
                    profiles = self.profile_service.load_profiles()
                    if profiles:
                        self.update_info(f"📋 {len(profiles)} perfil(es) cargado(s)", "info")
                    else:
                        self.update_info("💡 Cree su primer perfil de búsqueda", "info")
            else:
                self.update_info("❌ Error: Servicios no disponibles", "error")

        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"⚠️ Advertencia en verificación inicial: {error_msg}", "warning")

    def check_smtp_status(self):
        """Verifica el estado de la configuración SMTP"""
        try:
            from services.config_service import ConfigService
            config_service = ConfigService()

            if config_service.credentials_exist():
                self.smtp_status_label.config(text="Estado: Configurado ✓", foreground="green")
                self.status_label.config(text="SMTP Configurado")
                return True
            else:
                self.smtp_status_label.config(text="Estado: No configurado ✗", foreground="red")
                self.status_label.config(text="SMTP No configurado")
                return False
        except Exception:
            self.smtp_status_label.config(text="Estado: Error ⚠", foreground="orange")
            self.status_label.config(text="Error SMTP")
            return False

    def open_email_modal(self):
        """Abre el modal de configuración de email"""
        try:
            if not self.services_loaded:
                messagebox.showerror("Error", "Los servicios del sistema no están disponibles")
                return

            from gui.email_modal import EmailModal
            EmailModal(self.root, self.on_email_configured)
        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error abriendo configuración de email: {error_msg}", "error")
            messagebox.showerror("Error", f"No se pudo abrir la configuración de email:\n{error_msg}")

    def on_email_configured(self, success, message):
        """
        Callback llamado cuando se configura el email
        """
        try:
            clean_message = self.clean_message(message)
            status = "success" if success else "error"

            if success:
                self.update_info(f"✓ Configuración SMTP: {clean_message}", status)
                self.smtp_status_label.config(text="Estado: Configurado ✓", foreground="green")
                self.status_label.config(text="SMTP Configurado")
            else:
                self.update_info(f"✗ Error SMTP: {clean_message}", status)
                self.smtp_status_label.config(text="Estado: Error ⚠", foreground="orange")

        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error procesando configuración SMTP: {error_msg}", "error")

    def generate_excel_report(self):
        """Genera un reporte en formato Excel"""
        try:
            if not self.services_loaded:
                messagebox.showerror("Error", "Los servicios del sistema no están disponibles")
                return

            # Verificar si openpyxl está disponible
            if not self.excel_service.check_openpyxl_available():
                messagebox.showerror("Dependencia Faltante",
                                     "La librería 'openpyxl' no está instalada.\n\n"
                                     "Para instalar, ejecute en la terminal:\n"
                                     "pip install openpyxl")
                return

            # Verificar que hay perfiles
            profiles_stats = self.profile_service.get_all_profiles_stats()
            if not profiles_stats:
                messagebox.showwarning("Sin Datos", "No hay perfiles creados para generar el reporte")
                return

            self.update_info("🔄 Generando reporte Excel...", "info")
            self.excel_button.config(state="disabled", text="Generando...")

            # Ejecutar en hilo separado para no bloquear la UI
            thread = threading.Thread(target=self._generate_excel_thread, daemon=True)
            thread.start()

        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error generando Excel: {error_msg}", "error")
            self.excel_button.config(state="normal", text="📊 Generar Reporte Excel")

    def _generate_excel_thread(self):
        """Genera el reporte Excel en un hilo separado"""
        try:
            profiles_stats = self.profile_service.get_all_profiles_stats()
            filepath = self.excel_service.generate_profiles_report(profiles_stats)

            self.root.after(0, self._show_excel_result, True, f"Reporte Excel generado exitosamente:\n{filepath}")

        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.root.after(0, self._show_excel_result, False, f"Error generando Excel: {error_msg}")

    def _show_excel_result(self, success, message):
        """Muestra el resultado de la generación de Excel"""
        try:
            self.excel_button.config(state="normal", text="📊 Generar Reporte Excel")

            status = "success" if success else "error"
            symbol = "✓" if success else "❌"
            self.update_info(f"{symbol} {message}", status)

            if success:
                result = messagebox.askyesno("Reporte Generado",
                                             f"{message}\n\n¿Desea abrir la carpeta de reportes?")
                if result:
                    self.open_reports_folder()
            else:
                messagebox.showerror("Error", message)
        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error mostrando resultado Excel: {error_msg}", "error")

    def generate_csv_report(self):
        """Genera un reporte en formato CSV"""
        try:
            if not self.services_loaded:
                messagebox.showerror("Error", "Los servicios del sistema no están disponibles")
                return

            profiles_stats = self.profile_service.get_all_profiles_stats()
            if not profiles_stats:
                messagebox.showwarning("Sin Datos", "No hay perfiles creados para generar el reporte")
                return

            self.update_info("🔄 Generando reporte CSV...", "info")
            self.csv_button.config(state="disabled", text="Generando...")

            thread = threading.Thread(target=self._generate_csv_thread, daemon=True)
            thread.start()

        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error generando CSV: {error_msg}", "error")
            self.csv_button.config(state="normal", text="📋 Generar Reporte CSV")

    def _generate_csv_thread(self):
        """Genera el reporte CSV en un hilo separado"""
        try:
            profiles_stats = self.profile_service.get_all_profiles_stats()
            filepath = self.excel_service.create_csv_report(profiles_stats)

            self.root.after(0, self._show_csv_result, True, f"Reporte CSV generado exitosamente:\n{filepath}")

        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.root.after(0, self._show_csv_result, False, f"Error generando CSV: {error_msg}")

    def _show_csv_result(self, success, message):
        """Muestra el resultado de la generación de CSV"""
        try:
            self.csv_button.config(state="normal", text="📋 Generar Reporte CSV")

            status = "success" if success else "error"
            symbol = "✓" if success else "❌"
            self.update_info(f"{symbol} {message}", status)

            if success:
                result = messagebox.askyesno("Reporte Generado",
                                             f"{message}\n\n¿Desea abrir la carpeta de reportes?")
                if result:
                    self.open_reports_folder()
            else:
                messagebox.showerror("Error", message)
        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error mostrando resultado CSV: {error_msg}", "error")

    def show_available_reports(self):
        """Muestra una ventana con los reportes disponibles"""
        try:
            if not self.services_loaded:
                messagebox.showerror("Error", "Los servicios del sistema no están disponibles")
                return

            reports = self.excel_service.get_available_reports()

            if not reports:
                messagebox.showinfo("Sin Reportes", "No hay reportes generados aún")
                return

            # Crear ventana de reportes (código igual que antes)
            self._create_reports_window(reports)

        except Exception as e:
            error_msg = self.clean_message(str(e))
            messagebox.showerror("Error", f"Error mostrando reportes: {error_msg}")

    def _create_reports_window(self, reports):
        """Crea la ventana de reportes disponibles"""
        reports_window = tk.Toplevel(self.root)
        reports_window.title("Reportes Disponibles")
        reports_window.geometry("700x450")
        reports_window.transient(self.root)
        reports_window.grab_set()

        # Centrar ventana
        reports_window.update_idletasks()
        x = (reports_window.winfo_screenwidth() // 2) - (700 // 2)
        y = (reports_window.winfo_screenheight() // 2) - (450 // 2)
        reports_window.geometry(f"700x450+{x}+{y}")

        # Frame principal
        main_frame = ttk.Frame(reports_window, padding="15")
        main_frame.pack(fill="both", expand=True)

        # Título
        title_label = ttk.Label(main_frame, text="📁 Reportes Generados",
                                font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 15))

        # Lista de reportes
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 15))

        columns = ("name", "size", "modified")
        reports_tree = ttk.Treeview(list_frame, columns=columns, show="headings")

        reports_tree.heading("name", text="Nombre del Archivo")
        reports_tree.heading("size", text="Tamaño")
        reports_tree.heading("modified", text="Última Modificación")

        reports_tree.column("name", width=350, anchor="w")
        reports_tree.column("size", width=100, anchor="center")
        reports_tree.column("modified", width=200, anchor="center")

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=reports_tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=reports_tree.xview)
        reports_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        reports_tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        # Agregar reportes
        for report in reports:
            modified_date = datetime.fromtimestamp(report["modified"]).strftime("%d/%m/%Y %H:%M:%S")
            size_kb = f"{report['size'] / 1024:.1f} KB"
            reports_tree.insert("", "end", values=(report["name"], size_kb, modified_date))

        # Botones
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x")

        ttk.Button(buttons_frame, text="📂 Abrir Carpeta",
                   command=self.open_reports_folder).pack(side="left", padx=(0, 10))

        ttk.Button(buttons_frame, text="❌ Cerrar",
                   command=reports_window.destroy).pack(side="right")

    def open_reports_folder(self):
        """Abre la carpeta de reportes en el explorador"""
        try:
            if not self.services_loaded:
                return

            reports_path = str(self.excel_service.reports_dir)

            if platform.system() == "Windows":
                os.startfile(reports_path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", reports_path])
            else:
                subprocess.run(["xdg-open", reports_path])

            self.update_info(f"📂 Carpeta de reportes abierta: {reports_path}", "info")

        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error abriendo carpeta: {error_msg}", "error")
            messagebox.showerror("Error", f"No se pudo abrir la carpeta de reportes: {error_msg}")

    def clear_info_log(self):
        """Limpia el log de información"""
        try:
            result = messagebox.askyesno("Limpiar Log", "¿Desea limpiar el registro de actividad?")
            if result:
                if hasattr(self, 'info_text'):
                    self.info_text.config(state="normal")
                    self.info_text.delete(1.0, "end")
                    self.info_text.config(state="disabled")
                self.message_count = 0
                self.update_activity_title()
                self.update_info("🧹 Log de actividad limpiado", "info")
        except Exception as e:
            error_msg = self.clean_message(str(e))
            self.update_info(f"❌ Error limpiando log: {error_msg}", "error")

    def update_info(self, message, status="info"):
        """Actualiza el área de información en la columna derecha"""
        try:
            clean_message = self.clean_message(message)
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] {clean_message}"

            if hasattr(self, 'info_text'):
                self.info_text.config(state="normal")
                self.info_text.insert("end", f"{formatted_message}\n", status)
                self.info_text.see("end")
                self.info_text.config(state="disabled")

                self.message_count += 1
                self.update_activity_title()

            if hasattr(self, 'info_status'):
                self.info_status.config(text=clean_message[:50] + ("..." if len(clean_message) > 50 else ""))

        except Exception as e:
            print(f"Error actualizando información: {e}")

    def update_activity_title(self):
        """Actualiza el título del monitor de actividad con el contador"""
        try:
            if hasattr(self, 'activity_title'):
                self.activity_title.config(text=f"Registro de Actividad ({self.message_count})")
        except Exception:
            pass

    def clean_message(self, message):
        """Limpia un mensaje de caracteres problemáticos"""
        if not message:
            return ""

        try:
            message = str(message)

            replacements = {
                '\xa0': ' ',  # Espacio no-rompible
                '\u2019': "'",  # Apostrofe curvo derecho
                '\u2018': "'",  # Apostrofe curvo izquierdo
                '\u201c': '"',  # Comilla curva izquierda
                '\u201d': '"',  # Comilla curva derecha
                '\u2013': '-',  # En dash
                '\u2014': '--',  # Em dash
                '\u2026': '...'  # Ellipsis
            }

            for char, replacement in replacements.items():
                message = message.replace(char, replacement)

            return message.encode('ascii', 'ignore').decode('ascii')

        except Exception:
            return "Error de codificacion en mensaje"