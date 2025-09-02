# gui/email_send_modal.py
"""
Modal para configuración de parámetros de envío de reportes por correo.
Permite configurar destinatario, asunto, CC para envío automático de reportes Excel.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from services.config_service import ConfigService


class EmailSendModal:
    def __init__(self, parent, callback=None):
        """
        Inicializa el modal de configuración de envío de reportes

        Args:
            parent: Ventana padre
            callback: Función callback para notificar el resultado
        """
        self.parent = parent
        self.callback = callback
        self.config_service = ConfigService()

        self.create_modal()
        self.load_saved_config()

    def create_modal(self):
        """Crea la ventana modal con todos los campos necesarios"""
        # Crear ventana modal
        self.modal = tk.Toplevel(self.parent)
        self.modal.title("Configuracion de Envio de Reportes")
        self.modal.geometry("650x500")
        self.modal.resizable(True, True)
        self.modal.minsize(600, 450)

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
        x = (self.modal.winfo_screenwidth() // 2) - (650 // 2)
        y = (self.modal.winfo_screenheight() // 2) - (700 // 2)
        self.modal.geometry(f"650x700+{x}+{y}")

    def create_content(self):
        """Crea el contenido del modal"""
        # Frame principal
        main_frame = ttk.Frame(self.modal, padding="25")
        main_frame.pack(fill="both", expand=True)

        # Título
        title_label = ttk.Label(main_frame, text="Configuracion de Envio de Reportes",
                                font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 25))

        # Crear sección de habilitación
        self.create_enable_section(main_frame)

        # Crear campos del formulario
        self.create_form_fields(main_frame)

        # Crear botones
        self.create_buttons(main_frame)

        # Estado
        self.create_status_area(main_frame)

    def create_enable_section(self, parent):
        """Crea la sección para habilitar/deshabilitar envío automático"""
        # Frame para habilitar
        enable_frame = ttk.LabelFrame(parent, text="Estado del Envio Automatico", padding="15")
        enable_frame.pack(fill="x", pady=(0, 20))

        # Variable y checkbox para habilitar
        self.enabled_var = tk.BooleanVar()
        self.enable_checkbox = ttk.Checkbutton(enable_frame,
                                               text="Habilitar envio automatico de reportes",
                                               variable=self.enabled_var,
                                               command=self.on_enabled_change)
        self.enable_checkbox.pack(anchor="w", pady=(0, 5))

        # Nota informativa
        info_label = ttk.Label(enable_frame,
                               text="💡 Cuando esté habilitado, los reportes se enviarán automáticamente por correo",
                               foreground="gray", font=("Arial", 9))
        info_label.pack(anchor="w", pady=(5, 0))

    def create_form_fields(self, parent):
        """
        Crea los campos del formulario

        Args:
            parent: Frame padre para los campos
        """
        # Frame para campos
        self.fields_frame = ttk.LabelFrame(parent, text="Configuracion de Envio", padding="20")
        self.fields_frame.pack(fill="both", expand=True, pady=(0, 20))

        # Variables para los campos
        self.subject_var = tk.StringVar(value="Reporte de Registros de Bot - OpoBot")
        self.recipient_var = tk.StringVar()
        self.cc_var = tk.StringVar()

        # Asunto del correo
        ttk.Label(self.fields_frame, text="Asunto del correo:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(0, 8))
        subject_entry = ttk.Entry(self.fields_frame, textvariable=self.subject_var, width=70)
        subject_entry.pack(fill="x", pady=(0, 20), ipady=4)

        # Destinatario principal
        ttk.Label(self.fields_frame, text="Destinatario principal:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(0, 8))
        recipient_entry = ttk.Entry(self.fields_frame, textvariable=self.recipient_var, width=70)
        recipient_entry.pack(fill="x", pady=(0, 20), ipady=4)

        # CC (Copia carbón)
        ttk.Label(self.fields_frame, text="CC (Copia carbón - opcional):", font=("Arial", 11, "bold")).pack(anchor="w", pady=(0, 8))
        cc_entry = ttk.Entry(self.fields_frame, textvariable=self.cc_var, width=70)
        cc_entry.pack(fill="x", pady=(0, 15), ipady=4)

        # Nota sobre CC
        cc_note = ttk.Label(self.fields_frame,
                           text="💡 Para múltiples destinatarios en CC, sepárelos con comas: email1@domain.com, email2@domain.com",
                           foreground="gray", font=("Arial", 9), wraplength=500)
        cc_note.pack(anchor="w")

        # Configurar estado inicial
        self.on_enabled_change()

    def create_buttons(self, parent):
        """
        Crea los botones del modal

        Args:
            parent: Frame padre para los botones
        """
        # Frame para botones
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill="x", pady=(20, 20))

        # Botón probar envío
        self.test_button = ttk.Button(buttons_frame, text="Probar Envio",
                                      command=self.test_email_send, width=15)
        self.test_button.pack(side="left", padx=(0, 15))

        # Botón guardar
        self.save_button = ttk.Button(buttons_frame, text="Guardar Configuracion",
                                      command=self.save_config, width=20)
        self.save_button.pack(side="left", padx=(0, 15))

        # Botón cancelar
        cancel_button = ttk.Button(buttons_frame, text="Cancelar",
                                   command=self.close_modal, width=12)
        cancel_button.pack(side="right")

    def create_status_area(self, parent):
        """
        Crea el área de estado

        Args:
            parent: Frame padre para el área de estado
        """
        # Frame de estado
        status_frame = ttk.LabelFrame(parent, text="Estado", padding="15")
        status_frame.pack(fill="x", pady=(0, 10))

        # Label de estado
        self.status_label = ttk.Label(status_frame, text="Listo para configurar",
                                      foreground="blue")
        self.status_label.pack(pady=10)

        # Progressbar (oculto inicialmente)
        self.progress_bar = ttk.Progressbar(status_frame, mode="indeterminate")

    def on_enabled_change(self):
        """Maneja el cambio de estado habilitado/deshabilitado"""
        try:
            enabled = self.enabled_var.get()

            # Habilitar/deshabilitar controles de configuración
            state = "normal" if enabled else "disabled"

            # Obtener todos los widgets dentro del frame de campos
            for widget in self.fields_frame.winfo_children():
                try:
                    if hasattr(widget, 'configure'):
                        widget_config = widget.configure()
                        if 'state' in widget_config:
                            widget.configure(state=state)
                except Exception:
                    continue

            # También controlar botones
            if hasattr(self, 'test_button'):
                self.test_button.config(state=state)

        except Exception as e:
            print(f"Error cambiando estado habilitado: {e}")

    def load_saved_config(self):
        """Carga la configuración guardada si existe"""
        try:
            config = self.config_service.load_email_send_config()
            if config:
                self.enabled_var.set(config.get("enabled", False))
                self.subject_var.set(config.get("subject", "Reporte de Registros de Bot - OpoBot"))
                self.recipient_var.set(config.get("recipient", ""))
                self.cc_var.set(config.get("cc", ""))
                self.update_status("Configuracion cargada", "green")
                self.on_enabled_change()
            else:
                self.update_status("Sin configuracion previa", "blue")
        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.update_status(f"Error cargando configuracion: {error_msg}", "red")

    def validate_fields(self):
        """
        Valida que todos los campos requeridos estén correctos

        Returns:
            bool: True si todos los campos son válidos
        """
        if not self.enabled_var.get():
            return True  # Si está deshabilitado, no necesita validación

        # Validar asunto
        if not self.subject_var.get().strip():
            messagebox.showerror("Error", "El asunto del correo es requerido", parent=self.modal)
            return False

        # Validar destinatario principal
        recipient = self.recipient_var.get().strip()
        if not recipient:
            messagebox.showerror("Error", "El destinatario principal es requerido", parent=self.modal)
            return False

        # Validar formato de email del destinatario
        if not self._validate_email_format(recipient):
            messagebox.showerror("Error", "El formato del destinatario principal no es valido", parent=self.modal)
            return False

        # Validar CC si está presente
        cc_emails = self.cc_var.get().strip()
        if cc_emails:
            # Dividir por comas y validar cada email
            cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
            for email in cc_list:
                if not self._validate_email_format(email):
                    messagebox.showerror("Error", f"El formato del email en CC no es valido: {email}", parent=self.modal)
                    return False

        return True

    def save_config(self):
        """Guarda la configuración de envío"""
        if not self.validate_fields():
            return

        try:
            config = {
                "enabled": self.enabled_var.get(),
                "subject": self.subject_var.get().strip(),
                "recipient": self.recipient_var.get().strip(),
                "cc": self.cc_var.get().strip()
            }

            self.config_service.save_email_send_config(config)
            self.update_status("Configuracion guardada exitosamente", "green")

            # Notificar al callback si existe
            if self.callback:
                self.callback(True, "Configuracion de envio guardada correctamente")

            messagebox.showinfo("Exito", "Configuracion de envio guardada correctamente", parent=self.modal)

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            full_error = f"Error guardando configuracion: {error_msg}"
            self.update_status(full_error, "red")
            messagebox.showerror("Error", full_error, parent=self.modal)

            if self.callback:
                self.callback(False, full_error)

    def test_email_send(self):
        """Prueba el envío de correo con la configuración actual"""
        if not self.validate_fields():
            return

        try:
            # Verificar que existan credenciales SMTP
            if not self.config_service.credentials_exist():
                messagebox.showwarning("Advertencia",
                                     "Configure primero las credenciales SMTP desde 'Configurar Email SMTP'",
                                     parent=self.modal)
                return

            # Deshabilitar botón y mostrar progreso
            self.test_button.config(state="disabled")
            self.progress_bar.pack(pady=10)
            self.progress_bar.start()
            self.update_status("Enviando correo de prueba...", "blue")

            # Ejecutar prueba en hilo separado
            thread = threading.Thread(target=self._test_email_send_thread, daemon=True)
            thread.start()

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.update_status(f"Error iniciando prueba: {error_msg}", "red")
            self._update_test_result(False, error_msg)

    def _test_email_send_thread(self):
        """Ejecuta la prueba de envío en un hilo separado"""
        try:
            from services.email_send_service import EmailSendService

            # Obtener configuración
            config = {
                "subject": self.subject_var.get().strip(),
                "recipient": self.recipient_var.get().strip(),
                "cc": self.cc_var.get().strip()
            }

            # Crear servicio de envío
            email_send_service = EmailSendService()

            # Enviar email de prueba (sin adjunto)
            success, message = email_send_service.send_test_email(config)

            # Limpiar mensaje
            clean_message = self.clean_error_message(message)

            # Actualizar UI en el hilo principal
            self.modal.after(0, self._update_test_result, success, clean_message)

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.modal.after(0, self._update_test_result, False, f"Error en prueba: {error_msg}")

    def _update_test_result(self, success, message):
        """
        Actualiza la UI con el resultado de la prueba

        Args:
            success (bool): True si el envío fue exitoso
            message (str): Mensaje de resultado
        """
        try:
            # Detener progreso y rehabilitar botón
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.test_button.config(state="normal")

            # Actualizar estado
            color = "green" if success else "red"
            self.update_status(message, color)

            # Mostrar resultado
            if success:
                messagebox.showinfo("Prueba Exitosa", message, parent=self.modal)
            else:
                messagebox.showerror("Error en Prueba", message, parent=self.modal)

        except Exception as e:
            print(f"Error actualizando resultado de prueba: {e}")

    def _validate_email_format(self, email):
        """
        Valida el formato básico de un email

        Args:
            email (str): Email a validar

        Returns:
            bool: True si el formato es válido
        """
        if not email:
            return False

        import re
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email.strip()) is not None

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