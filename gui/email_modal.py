# gui/email_modal.py
"""
Modal para configuración de credenciales SMTP.
Permite ingresar datos de email, probar conexión y guardar credenciales.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from services.email_service import EmailService
from services.config_service import ConfigService


class EmailModal:
    def __init__(self, parent, callback=None):
        """
        Inicializa el modal de configuración de email

        Args:
            parent: Ventana padre
            callback: Función callback para notificar el resultado
        """
        self.parent = parent
        self.callback = callback
        self.email_service = EmailService()
        self.config_service = ConfigService()

        self.create_modal()
        self.load_saved_credentials()

    def create_modal(self):
        """Crea la ventana modal con todos los campos necesarios"""
        # Crear ventana modal
        self.modal = tk.Toplevel(self.parent)
        self.modal.title("Configuracion SMTP")
        self.modal.geometry("450x400")
        self.modal.resizable(False, False)

        # Hacer modal
        self.modal.transient(self.parent)
        self.modal.grab_set()

        # Centrar modal
        self.center_modal()

        # Crear contenido
        self.create_content()

        # Manejar cierre de ventana
        self.modal.protocol("WM_DELETE_WINDOW", self.close_modal)

    def center_modal(self):
        """Centra el modal en la pantalla"""
        self.modal.update_idletasks()
        x = (self.modal.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.modal.winfo_screenheight() // 2) - (400 // 2)
        self.modal.geometry(f"450x400+{x}+{y}")

    def create_content(self):
        """Crea el contenido del modal"""
        # Frame principal
        main_frame = ttk.Frame(self.modal, padding="20")
        main_frame.pack(fill="both", expand=True)

        # Título
        title_label = ttk.Label(main_frame, text="Configuracion Email SMTP",
                                font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))

        # Crear campos
        self.create_form_fields(main_frame)

        # Crear botones
        self.create_buttons(main_frame)

        # Estado de conexión
        self.create_status_area(main_frame)

    def create_form_fields(self, parent):
        """
        Crea los campos del formulario

        Args:
            parent: Frame padre para los campos
        """
        # Frame para campos
        fields_frame = ttk.LabelFrame(parent, text="Datos SMTP", padding="10")
        fields_frame.pack(fill="x", pady=(0, 15))

        # Variables para los campos
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.server_var = tk.StringVar(value="smtp.gmail.com")
        self.port_var = tk.StringVar(value="587")

        # Email
        ttk.Label(fields_frame, text="Email:").grid(row=0, column=0, sticky="w", pady=5)
        email_entry = ttk.Entry(fields_frame, textvariable=self.email_var, width=35)
        email_entry.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=5)
        email_entry.bind('<FocusOut>', self.on_email_change)

        # Contraseña
        ttk.Label(fields_frame, text="Contrasena:").grid(row=1, column=0, sticky="w", pady=5)
        password_entry = ttk.Entry(fields_frame, textvariable=self.password_var,
                                   show="*", width=35)
        password_entry.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)

        # Servidor SMTP
        ttk.Label(fields_frame, text="Servidor SMTP:").grid(row=2, column=0, sticky="w", pady=5)
        server_entry = ttk.Entry(fields_frame, textvariable=self.server_var, width=35)
        server_entry.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)

        # Puerto
        ttk.Label(fields_frame, text="Puerto:").grid(row=3, column=0, sticky="w", pady=5)
        port_entry = ttk.Entry(fields_frame, textvariable=self.port_var, width=35)
        port_entry.grid(row=3, column=1, sticky="ew", padx=(10, 0), pady=5)

        # Configurar expansión de columnas
        fields_frame.grid_columnconfigure(1, weight=1)

    def create_buttons(self, parent):
        """
        Crea los botones del modal

        Args:
            parent: Frame padre para los botones
        """
        # Frame para botones
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill="x", pady=(0, 15))

        # Botón probar conexión
        self.test_button = ttk.Button(buttons_frame, text="Probar Conexion",
                                      command=self.test_connection)
        self.test_button.pack(side="left", padx=(0, 10))

        # Botón guardar
        self.save_button = ttk.Button(buttons_frame, text="Guardar",
                                      command=self.save_credentials)
        self.save_button.pack(side="left", padx=(0, 10))

        # Botón cancelar
        cancel_button = ttk.Button(buttons_frame, text="Cancelar",
                                   command=self.close_modal)
        cancel_button.pack(side="right")

    def create_status_area(self, parent):
        """
        Crea el área de estado de la conexión

        Args:
            parent: Frame padre para el área de estado
        """
        # Frame de estado
        status_frame = ttk.LabelFrame(parent, text="Estado", padding="10")
        status_frame.pack(fill="both", expand=True)

        # Label de estado
        self.status_label = ttk.Label(status_frame, text="Listo para probar conexion",
                                      foreground="blue")
        self.status_label.pack()

        # Progressbar (oculto inicialmente)
        self.progress_bar = ttk.Progressbar(status_frame, mode="indeterminate")

    def on_email_change(self, event=None):
        """Detecta automáticamente el proveedor cuando cambia el email"""
        try:
            email = self.email_var.get().strip()
            if email and "@" in email:
                provider = self.email_service.detect_provider_from_email(email)
                if provider:
                    settings = self.email_service.get_common_smtp_settings(provider)
                    if settings:
                        self.server_var.set(settings["server"])
                        self.port_var.set(str(settings["port"]))
                        self.update_status(f"Configuracion automatica para {provider.upper()}", "green")
        except Exception:
            pass  # Ignorar errores en la detección automática

    def load_saved_credentials(self):
        """Carga las credenciales guardadas si existen"""
        try:
            credentials = self.config_service.load_credentials()
            if credentials:
                self.email_var.set(credentials.get("email", ""))
                self.password_var.set(credentials.get("password", ""))
                self.server_var.set(credentials.get("server", "smtp.gmail.com"))
                self.port_var.set(str(credentials.get("port", "587")))
                self.update_status("Credenciales cargadas", "green")
        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.update_status(f"Error cargando credenciales: {error_msg}", "red")

    def test_connection(self):
        """Prueba la conexión SMTP en un hilo separado"""
        # Validar campos
        if not self.validate_fields():
            return

        # Deshabilitar botón y mostrar progreso
        self.test_button.config(state="disabled")
        self.progress_bar.pack(pady=10)
        self.progress_bar.start()
        self.update_status("Probando conexion...", "blue")

        # Ejecutar prueba en hilo separado
        thread = threading.Thread(target=self._test_connection_thread)
        thread.daemon = True
        thread.start()

    def _test_connection_thread(self):
        """Ejecuta la prueba de conexión en un hilo separado"""
        try:
            credentials = {
                "email": self.email_var.get().strip(),
                "password": self.password_var.get().strip(),
                "server": self.server_var.get().strip(),
                "port": self.port_var.get().strip()
            }

            success, message = self.email_service.test_connection(credentials)

            # Limpiar mensaje de caracteres problemáticos
            clean_message = self.clean_error_message(message)

            # Actualizar UI en el hilo principal
            self.modal.after(0, self._update_test_result, success, clean_message)

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.modal.after(0, self._update_test_result, False, error_msg)

    def _update_test_result(self, success, message):
        """
        Actualiza la UI con el resultado de la prueba

        Args:
            success (bool): True si la conexión fue exitosa
            message (str): Mensaje de resultado
        """
        # Detener progreso y rehabilitar botón
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.test_button.config(state="normal")

        # Actualizar estado
        color = "green" if success else "red"
        self.update_status(message, color)

    def validate_fields(self):
        """
        Valida que todos los campos requeridos estén llenos

        Returns:
            bool: True si todos los campos son válidos
        """
        if not self.email_var.get().strip():
            messagebox.showerror("Error", "El campo Email es requerido")
            return False

        if not self.password_var.get().strip():
            messagebox.showerror("Error", "El campo Contrasena es requerido")
            return False

        if not self.server_var.get().strip():
            messagebox.showerror("Error", "El campo Servidor SMTP es requerido")
            return False

        try:
            port = int(self.port_var.get().strip())
            if port <= 0 or port > 65535:
                raise ValueError("Puerto fuera de rango")
        except ValueError:
            messagebox.showerror("Error", "El puerto debe ser un numero valido (1-65535)")
            return False

        # Validar formato de email
        if not self.email_service.validate_email_format(self.email_var.get().strip()):
            messagebox.showerror("Error", "El formato del email no es valido")
            return False

        return True

    def save_credentials(self):
        """Guarda las credenciales en formato JSON"""
        if not self.validate_fields():
            return

        try:
            credentials = {
                "email": self.email_var.get().strip(),
                "password": self.password_var.get().strip(),
                "server": self.server_var.get().strip(),
                "port": self.port_var.get().strip()
            }

            self.config_service.save_credentials(credentials)
            self.update_status("Credenciales guardadas exitosamente", "green")

            # Notificar al callback si existe
            if self.callback:
                self.callback(True, "Credenciales guardadas correctamente")

            messagebox.showinfo("Exito", "Credenciales guardadas correctamente")

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            full_error = f"Error guardando credenciales: {error_msg}"
            self.update_status(full_error, "red")
            messagebox.showerror("Error", full_error)

            if self.callback:
                self.callback(False, full_error)

    def update_status(self, message, color="blue"):
        """
        Actualiza el label de estado

        Args:
            message (str): Mensaje a mostrar
            color (str): Color del texto
        """
        clean_message = self.clean_error_message(message)
        self.status_label.config(text=clean_message, foreground=color)

    def clean_error_message(self, message):
        """
        Limpia mensajes de error de caracteres problemáticos

        Args:
            message (str): Mensaje original

        Returns:
            str: Mensaje limpio
        """
        try:
            if not message:
                return ""

            # Reemplazar caracteres problemáticos comunes
            message = str(message)
            message = message.replace('\xa0', ' ')  # Espacio no-rompible
            message = message.replace('\u2019', "'")  # Apostrofe curvo
            message = message.replace('\u2018', "'")  # Apostrofe curvo
            message = message.replace('\u201c', '"')  # Comilla curva
            message = message.replace('\u201d', '"')  # Comilla curva

            # Codificar y decodificar para limpiar caracteres problemáticos
            return message.encode('ascii', 'ignore').decode('ascii')
        except Exception:
            return "Error de codificacion en mensaje"

    def close_modal(self):
        """Cierra el modal"""
        try:
            self.modal.destroy()
        except Exception:
            pass  # Ignorar errores al cerrar