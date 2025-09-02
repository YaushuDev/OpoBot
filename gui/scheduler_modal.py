# gui/scheduler_modal.py
"""
Modal para configuración de búsquedas automáticas programadas.
Permite configurar horarios, intervalos y días para ejecutar búsquedas automáticamente.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from services.scheduler_service import SchedulerService


class SchedulerModal:
    def __init__(self, parent, callback=None):
        """
        Inicializa el modal de configuración de programación

        Args:
            parent: Ventana padre
            callback: Función callback para notificar el resultado
        """
        self.parent = parent
        self.callback = callback
        self.scheduler_service = SchedulerService()

        self.create_modal()
        self.load_current_configuration()

    def create_modal(self):
        """Crea la ventana modal con todos los campos necesarios"""
        # Crear ventana modal
        self.modal = tk.Toplevel(self.parent)
        self.modal.title("Configuración de Búsquedas Automáticas")
        self.modal.geometry("500x600")
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
        x = (self.modal.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.modal.winfo_screenheight() // 2) - (600 // 2)
        self.modal.geometry(f"500x600+{x}+{y}")

    def create_content(self):
        """Crea el contenido del modal"""
        # Frame principal con scrollbar
        main_canvas = tk.Canvas(self.modal)
        scrollbar = ttk.Scrollbar(self.modal, orient="vertical", command=main_canvas.yview)
        self.scrollable_frame = ttk.Frame(main_canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        # Frame de contenido principal
        main_frame = ttk.Frame(self.scrollable_frame, padding="20")
        main_frame.pack(fill="both", expand=True)

        # Título
        title_label = ttk.Label(main_frame, text="Configuración de Búsquedas Automáticas",
                                font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))

        # Crear secciones
        self.create_enable_section(main_frame)
        self.create_schedule_type_section(main_frame)

        # Contenedor para las configuraciones específicas
        self.config_container = ttk.Frame(main_frame)
        self.config_container.pack(fill="x", pady=(0, 15))

        self.create_daily_section(self.config_container)
        self.create_weekly_section(self.config_container)
        self.create_interval_section(self.config_container)

        self.create_buttons(main_frame)
        self.create_status_area(main_frame)

        # Pack canvas y scrollbar
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Configurar eventos
        self.setup_events()

    def create_enable_section(self, parent):
        """Crea la sección para habilitar/deshabilitar programación"""
        # Frame para habilitar
        enable_frame = ttk.LabelFrame(parent, text="Estado de la Programación", padding="10")
        enable_frame.pack(fill="x", pady=(0, 15))

        # Variable y checkbox para habilitar
        self.enabled_var = tk.BooleanVar()
        self.enable_checkbox = ttk.Checkbutton(enable_frame, text="Habilitar búsquedas automáticas",
                                               variable=self.enabled_var)
        self.enable_checkbox.pack(anchor="w")

    def create_schedule_type_section(self, parent):
        """Crea la sección para seleccionar tipo de programación"""
        # Frame para tipo de programación
        type_frame = ttk.LabelFrame(parent, text="Tipo de Programación", padding="10")
        type_frame.pack(fill="x", pady=(0, 15))

        # Variable para tipo de programación
        self.schedule_type_var = tk.StringVar(value="daily")

        # Opciones de tipo
        ttk.Radiobutton(type_frame, text="Diario - Ejecutar todos los días a una hora específica",
                        variable=self.schedule_type_var, value="daily").pack(anchor="w", pady=2)

        ttk.Radiobutton(type_frame, text="Semanal - Ejecutar días específicos de la semana",
                        variable=self.schedule_type_var, value="weekly").pack(anchor="w", pady=2)

        ttk.Radiobutton(type_frame, text="Por Intervalo - Ejecutar cada cierto tiempo",
                        variable=self.schedule_type_var, value="interval").pack(anchor="w", pady=2)

    def create_daily_section(self, parent):
        """Crea la sección de configuración diaria"""
        self.daily_frame = ttk.LabelFrame(parent, text="Configuración Diaria", padding="10")

        # Campo de hora
        time_label = ttk.Label(self.daily_frame, text="Hora de ejecución (HH:MM):")
        time_label.grid(row=0, column=0, sticky="w", pady=5)

        self.daily_time_var = tk.StringVar(value="09:00")
        time_entry = ttk.Entry(self.daily_frame, textvariable=self.daily_time_var, width=10)
        time_entry.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=5)

        # Ejemplo
        example_label = ttk.Label(self.daily_frame, text="Ejemplo: 14:30 (2:30 PM)",
                                  foreground="gray", font=("Arial", 9))
        example_label.grid(row=1, column=0, columnspan=2, sticky="w")

    def create_weekly_section(self, parent):
        """Crea la sección de configuración semanal"""
        self.weekly_frame = ttk.LabelFrame(parent, text="Configuración Semanal", padding="10")

        # Campo de hora
        time_label = ttk.Label(self.weekly_frame, text="Hora de ejecución (HH:MM):")
        time_label.grid(row=0, column=0, sticky="w", pady=5)

        self.weekly_time_var = tk.StringVar(value="09:00")
        time_entry = ttk.Entry(self.weekly_frame, textvariable=self.weekly_time_var, width=10)
        time_entry.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=5)

        # Días de la semana
        days_label = ttk.Label(self.weekly_frame, text="Días de la semana:")
        days_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(15, 5))

        # Frame para checkboxes de días
        days_frame = ttk.Frame(self.weekly_frame)
        days_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)

        # Variables para días
        self.day_vars = {}
        days = [
            ("monday", "Lunes"),
            ("tuesday", "Martes"),
            ("wednesday", "Miércoles"),
            ("thursday", "Jueves"),
            ("friday", "Viernes"),
            ("saturday", "Sábado"),
            ("sunday", "Domingo")
        ]

        for i, (day_key, day_name) in enumerate(days):
            self.day_vars[day_key] = tk.BooleanVar()
            if day_key in ["monday", "wednesday", "friday"]:  # Días por defecto
                self.day_vars[day_key].set(True)

            checkbox = ttk.Checkbutton(days_frame, text=day_name, variable=self.day_vars[day_key])
            checkbox.grid(row=i // 4, column=i % 4, sticky="w", padx=(0, 15), pady=2)

    def create_interval_section(self, parent):
        """Crea la sección de configuración por intervalo"""
        self.interval_frame = ttk.LabelFrame(parent, text="Configuración por Intervalo", padding="10")

        # Intervalo
        interval_label = ttk.Label(self.interval_frame, text="Ejecutar cada:")
        interval_label.grid(row=0, column=0, sticky="w", pady=5)

        self.interval_var = tk.StringVar(value="60")
        interval_spinbox = ttk.Spinbox(self.interval_frame, from_=1, to=999, textvariable=self.interval_var, width=8)
        interval_spinbox.grid(row=0, column=1, sticky="w", padx=(10, 5), pady=5)

        # Unidad
        self.unit_var = tk.StringVar(value="minutes")
        unit_combo = ttk.Combobox(self.interval_frame, textvariable=self.unit_var,
                                  values=["minutes", "hours"], state="readonly", width=10)
        unit_combo.grid(row=0, column=2, sticky="w", padx=(5, 0), pady=5)

        # Ejemplo
        example_label = ttk.Label(self.interval_frame, text="Ejemplo: 30 minutes = cada 30 minutos",
                                  foreground="gray", font=("Arial", 9))
        example_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def create_buttons(self, parent):
        """Crea los botones del modal"""
        # Frame para botones
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill="x", pady=(15, 15))

        # Botón guardar y aplicar
        self.save_button = ttk.Button(buttons_frame, text="Guardar y Aplicar",
                                      command=self.save_configuration, style="Accent.TButton")
        self.save_button.pack(side="left", padx=(0, 10))

        # Botón probar configuración
        self.test_button = ttk.Button(buttons_frame, text="Probar Configuración",
                                      command=self.test_configuration)
        self.test_button.pack(side="left", padx=(0, 10))

        # Botón cancelar
        cancel_button = ttk.Button(buttons_frame, text="Cancelar",
                                   command=self.close_modal)
        cancel_button.pack(side="right")

    def create_status_area(self, parent):
        """Crea el área de estado de la configuración"""
        # Frame de estado
        status_frame = ttk.LabelFrame(parent, text="Estado del Programador", padding="10")
        status_frame.pack(fill="x", pady=(0, 10))

        # Label de estado
        self.status_label = ttk.Label(status_frame, text="Configuración no guardada",
                                      foreground="blue")
        self.status_label.pack(pady=5)

        # Progressbar (oculto inicialmente)
        self.progress_bar = ttk.Progressbar(status_frame, mode="indeterminate")

        # Actualizar estado inicial
        self.update_scheduler_status()

    def setup_events(self):
        """Configura los eventos de la interfaz"""
        # Cambio en tipo de programación
        self.schedule_type_var.trace('w', self.on_schedule_type_change)

        # Cambio en estado habilitado
        self.enabled_var.trace('w', self.on_enabled_change)

        # Aplicar configuración inicial
        self.on_schedule_type_change()
        self.on_enabled_change()

    def on_schedule_type_change(self, *args):
        """Maneja el cambio de tipo de programación"""
        try:
            schedule_type = self.schedule_type_var.get()

            # Ocultar todos los frames primero
            self.daily_frame.pack_forget()
            self.weekly_frame.pack_forget()
            self.interval_frame.pack_forget()

            # Mostrar el frame correspondiente
            if schedule_type == "daily":
                self.daily_frame.pack(fill="x", pady=(0, 10))
            elif schedule_type == "weekly":
                self.weekly_frame.pack(fill="x", pady=(0, 10))
            elif schedule_type == "interval":
                self.interval_frame.pack(fill="x", pady=(0, 10))

        except Exception as e:
            print(f"Error cambiando tipo de programación: {e}")

    def on_enabled_change(self, *args):
        """Maneja el cambio de estado habilitado/deshabilitado"""
        try:
            enabled = self.enabled_var.get()

            # Habilitar/deshabilitar controles de configuración
            state = "normal" if enabled else "disabled"

            # Obtener todos los widgets dentro de cada frame de configuración
            for frame in [self.daily_frame, self.weekly_frame, self.interval_frame]:
                for widget in frame.winfo_children():
                    try:
                        if hasattr(widget, 'configure'):
                            widget_config = widget.configure()
                            if 'state' in widget_config:
                                widget.configure(state=state)
                    except Exception:
                        continue  # Ignorar widgets que no soportan 'state'

        except Exception as e:
            print(f"Error cambiando estado habilitado: {e}")

    def load_current_configuration(self):
        """Carga la configuración actual si existe"""
        try:
            config = self.scheduler_service.load_configuration()
            if config:
                # Cargar estado habilitado
                self.enabled_var.set(config.get("enabled", False))

                # Cargar tipo de programación
                schedule_type = config.get("type", "daily")
                self.schedule_type_var.set(schedule_type)

                # Cargar configuración específica según el tipo
                if schedule_type == "daily":
                    self.daily_time_var.set(config.get("time", "09:00"))
                elif schedule_type == "weekly":
                    self.weekly_time_var.set(config.get("time", "09:00"))
                    # Cargar días seleccionados
                    selected_days = config.get("days", ["monday", "wednesday", "friday"])
                    for day_key in self.day_vars:
                        self.day_vars[day_key].set(day_key in selected_days)
                elif schedule_type == "interval":
                    self.interval_var.set(str(config.get("interval", 60)))
                    self.unit_var.set(config.get("unit", "minutes"))

                self.update_status("Configuración cargada", "green")
            else:
                self.update_status("Sin configuración previa", "blue")

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.update_status(f"Error cargando configuración: {error_msg}", "red")

    def save_configuration(self):
        """Guarda la configuración del programador"""
        if not self.validate_configuration():
            return

        try:
            # Deshabilitar botón y mostrar progreso
            self.save_button.config(state="disabled")
            self.progress_bar.pack(pady=10)
            self.progress_bar.start()
            self.update_status("Guardando configuración...", "blue")

            # Ejecutar guardado en hilo separado
            thread = threading.Thread(target=self._save_configuration_thread, daemon=True)
            thread.start()

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.update_status(f"Error iniciando guardado: {error_msg}", "red")
            self.save_button.config(state="normal")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    def _save_configuration_thread(self):
        """Guarda la configuración en un hilo separado"""
        try:
            # Construir configuración
            config = {
                "enabled": self.enabled_var.get(),
                "type": self.schedule_type_var.get()
            }

            # Agregar configuración específica según el tipo
            schedule_type = self.schedule_type_var.get()

            if schedule_type == "daily":
                config["time"] = self.daily_time_var.get().strip()

            elif schedule_type == "weekly":
                config["time"] = self.weekly_time_var.get().strip()
                config["days"] = [day for day, var in self.day_vars.items() if var.get()]

            elif schedule_type == "interval":
                config["interval"] = int(self.interval_var.get())
                config["unit"] = self.unit_var.get()

            # Guardar configuración
            self.scheduler_service.save_configuration(config)

            # Si está habilitado, reiniciar el programador
            if config["enabled"]:
                success, message = self.scheduler_service.start_scheduler()
                if success:
                    result_message = "Configuración guardada y programador iniciado"
                else:
                    result_message = f"Configuración guardada pero error iniciando programador: {message}"
            else:
                # Si está deshabilitado, detener el programador
                self.scheduler_service.stop_scheduler()
                result_message = "Configuración guardada y programador detenido"

            # Actualizar UI en el hilo principal
            self.modal.after(0, self._update_save_result, True, result_message)

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.modal.after(0, self._update_save_result, False, f"Error guardando configuración: {error_msg}")

    def _update_save_result(self, success, message):
        """Actualiza la UI con el resultado del guardado"""
        try:
            # Detener progreso y rehabilitar botón
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.save_button.config(state="normal")

            # Actualizar estado
            color = "green" if success else "red"
            self.update_status(message, color)

            # Actualizar estado del programador
            self.update_scheduler_status()

            # Mostrar resultado
            if success:
                messagebox.showinfo("Éxito", message, parent=self.modal)

                # Notificar al callback
                if self.callback:
                    self.callback(True, "Configuración de programación actualizada")
            else:
                messagebox.showerror("Error", message, parent=self.modal)

        except Exception as e:
            print(f"Error actualizando resultado del guardado: {e}")

    def test_configuration(self):
        """Prueba la configuración actual"""
        if not self.validate_configuration():
            return

        try:
            # Construir configuración temporal para probar
            config = {
                "enabled": self.enabled_var.get(),
                "type": self.schedule_type_var.get()
            }

            schedule_type = self.schedule_type_var.get()

            if schedule_type == "daily":
                config["time"] = self.daily_time_var.get().strip()
            elif schedule_type == "weekly":
                config["time"] = self.weekly_time_var.get().strip()
                config["days"] = [day for day, var in self.day_vars.items() if var.get()]
            elif schedule_type == "interval":
                config["interval"] = int(self.interval_var.get())
                config["unit"] = self.unit_var.get()

            # Validar configuración
            if self.scheduler_service._validate_config(config):
                description = self._get_config_description(config)
                messagebox.showinfo("Configuración Válida",
                                    f"La configuración es válida:\n\n{description}",
                                    parent=self.modal)
                self.update_status("Configuración válida", "green")
            else:
                messagebox.showerror("Configuración Inválida",
                                     "La configuración no es válida. Verifique los campos.",
                                     parent=self.modal)
                self.update_status("Configuración inválida", "red")

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            messagebox.showerror("Error", f"Error probando configuración: {error_msg}", parent=self.modal)

    def validate_configuration(self):
        """Valida la configuración antes de guardar"""
        try:
            if not self.enabled_var.get():
                return True  # Si está deshabilitado, no necesita validación adicional

            schedule_type = self.schedule_type_var.get()

            if schedule_type in ["daily", "weekly"]:
                time_var = self.daily_time_var if schedule_type == "daily" else self.weekly_time_var
                time_str = time_var.get().strip()

                if not time_str:
                    messagebox.showerror("Error", "El campo de hora es requerido", parent=self.modal)
                    return False

                # Validar formato de hora
                try:
                    from datetime import datetime
                    datetime.strptime(time_str, "%H:%M")
                except ValueError:
                    messagebox.showerror("Error", "Formato de hora inválido. Use HH:MM (ej: 14:30)", parent=self.modal)
                    return False

                # Para configuración semanal, validar que al menos un día esté seleccionado
                if schedule_type == "weekly":
                    selected_days = [day for day, var in self.day_vars.items() if var.get()]
                    if not selected_days:
                        messagebox.showerror("Error", "Debe seleccionar al menos un día de la semana",
                                             parent=self.modal)
                        return False

            elif schedule_type == "interval":
                try:
                    interval = int(self.interval_var.get())
                    if interval <= 0:
                        messagebox.showerror("Error", "El intervalo debe ser mayor que 0", parent=self.modal)
                        return False

                    # Validar límites razonables
                    unit = self.unit_var.get()
                    if unit == "minutes" and interval < 1:
                        messagebox.showerror("Error", "El intervalo mínimo es 1 minuto", parent=self.modal)
                        return False
                    elif unit == "hours" and interval < 1:
                        messagebox.showerror("Error", "El intervalo mínimo es 1 hora", parent=self.modal)
                        return False

                except ValueError:
                    messagebox.showerror("Error", "El intervalo debe ser un número válido", parent=self.modal)
                    return False

            return True

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            messagebox.showerror("Error", f"Error validando configuración: {error_msg}", parent=self.modal)
            return False

    def update_scheduler_status(self):
        """Actualiza el estado del programador"""
        try:
            status = self.scheduler_service.get_scheduler_status()

            if status["is_running"] and status["enabled"]:
                status_text = f"🟢 Programador ACTIVO - {status.get('schedule_details', 'Sin detalles')}"
                if status.get("next_execution"):
                    status_text += f"\nPróxima ejecución: {status['next_execution']}"
                color = "green"
            elif status["has_config"] and status["enabled"]:
                status_text = f"🟡 Configurado pero DETENIDO - {status.get('schedule_details', 'Sin detalles')}"
                color = "orange"
            elif status["has_config"]:
                status_text = "🔴 Configurado pero DESHABILITADO"
                color = "red"
            else:
                status_text = "⚪ Sin configurar"
                color = "gray"

            self.update_status(status_text, color)

        except Exception as e:
            error_msg = self.clean_error_message(str(e))
            self.update_status(f"Error obteniendo estado: {error_msg}", "red")

    def _get_config_description(self, config):
        """Obtiene descripción legible de la configuración"""
        if not config.get("enabled"):
            return "Programación deshabilitada"

        schedule_type = config.get("type", "")

        if schedule_type == "daily":
            return f"Ejecutar diariamente a las {config.get('time', '09:00')}"
        elif schedule_type == "weekly":
            days = config.get("days", [])
            days_str = ", ".join([day.capitalize() for day in days])
            return f"Ejecutar {days_str} a las {config.get('time', '09:00')}"
        elif schedule_type == "interval":
            interval = config.get("interval", 60)
            unit = config.get("unit", "minutes")
            unit_name = "minutos" if unit == "minutes" else "horas"
            return f"Ejecutar cada {interval} {unit_name}"

        return "Configuración desconocida"

    def update_status(self, message, color="blue"):
        """Actualiza el label de estado"""
        try:
            clean_message = self.clean_error_message(message)
            self.status_label.config(text=clean_message, foreground=color)
        except Exception as e:
            print(f"Error actualizando status: {e}")

    def clean_error_message(self, message):
        """Limpia mensajes de error de caracteres problemáticos"""
        try:
            if not message:
                return ""

            # Reemplazar caracteres problemáticos comunes
            message = str(message)
            message = message.replace('\xa0', ' ')
            message = message.replace('\u2019', "'")
            message = message.replace('\u2018', "'")
            message = message.replace('\u201c', '"')
            message = message.replace('\u201d', '"')

            # Codificar y decodificar para limpiar caracteres problemáticos
            return message.encode('ascii', 'ignore').decode('ascii')
        except Exception:
            return "Error de codificación en mensaje"

    def close_modal(self):
        """Cierra el modal"""
        try:
            self.modal.destroy()
        except Exception:
            pass