# gui/profile_manager.py
"""
Interfaz gráfica para gestión de perfiles de búsqueda con programación automática.
Permite crear, editar, eliminar perfiles y configurar búsquedas automáticas programadas.
Incluye actualizaciones en tiempo real del estado del programador.
"""

# Archivos relacionados: services/profile_service.py, services/email_search_service.py, services/config_service.py, services/scheduler_service.py, gui/scheduler_modal.py

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
from services.profile_service import ProfileService
from services.email_search_service import EmailSearchService
from services.config_service import ConfigService
from services.scheduler_service import SchedulerService


class ProfileManager:
    def __init__(self, parent, info_callback=None):
        """
        Inicializa el gestor de perfiles con programación automática

        Args:
            parent: Widget padre donde se colocará el gestor
            info_callback: Función callback para mostrar información
        """
        self.parent = parent
        self.info_callback = info_callback
        self.profile_service = ProfileService()
        self.email_search_service = EmailSearchService()
        self.config_service = ConfigService()
        self.scheduler_service = SchedulerService()
        self.current_active_profiles = []
        self.profile_items = {}

        # Para actualizaciones en tiempo real
        self.status_update_job = None
        self.is_destroyed = False

        self.create_profile_manager()
        self.refresh_profiles_list()
        self.start_scheduler_if_configured()
        self.start_status_updates()

    def create_profile_manager(self):
        """Crea la interfaz del gestor de perfiles con programación automática"""
        # Frame principal
        self.main_frame = ttk.LabelFrame(self.parent, text="Perfiles de Busqueda", padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        self.main_frame.grid_rowconfigure(2, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Frame de botones superiores
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        buttons_frame.grid_columnconfigure(4, weight=1)

        # Botones de gestión de perfiles
        ttk.Button(buttons_frame, text="Nuevo Perfil",
                   command=self.create_new_profile).grid(row=0, column=0, padx=(0, 5))

        ttk.Button(buttons_frame, text="Editar",
                   command=self.edit_selected_profile).grid(row=0, column=1, padx=5)

        ttk.Button(buttons_frame, text="Eliminar",
                   command=self.delete_selected_profile).grid(row=0, column=2, padx=5)

        # Separador visual
        ttk.Separator(buttons_frame, orient="vertical").grid(row=0, column=3, sticky="ns", padx=15)

        # Botones de búsqueda
        self.manual_search_button = ttk.Button(buttons_frame, text="🔍 Buscar Ahora",
                                               command=self.execute_manual_search)
        self.manual_search_button.grid(row=0, column=5, padx=5)

        self.configure_schedule_button = ttk.Button(buttons_frame, text="⏰ Configurar Programación",
                                                    command=self.open_scheduler_modal, style="Accent.TButton")
        self.configure_schedule_button.grid(row=0, column=6, padx=(5, 0))

        # Área de estado del programador
        self.create_scheduler_status_area()

        # Lista de perfiles
        self.create_profiles_list()

    def create_scheduler_status_area(self):
        """Crea el área de estado del programador automático"""
        # Frame de estado del programador
        scheduler_frame = ttk.LabelFrame(self.main_frame, text="Estado de Búsquedas Automáticas", padding="8")
        scheduler_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        scheduler_frame.grid_columnconfigure(0, weight=1)

        # Frame superior con estado y timestamp
        status_header_frame = ttk.Frame(scheduler_frame)
        status_header_frame.grid(row=0, column=0, sticky="ew")
        status_header_frame.grid_columnconfigure(0, weight=1)

        # Label de estado principal
        self.scheduler_status_label = ttk.Label(status_header_frame, text="⚪ Sin configurar",
                                                foreground="gray", font=("Arial", 10, "bold"))
        self.scheduler_status_label.grid(row=0, column=0, sticky="w")

        # Label de última actualización
        self.last_update_label = ttk.Label(status_header_frame, text="",
                                           foreground="gray", font=("Arial", 8))
        self.last_update_label.grid(row=0, column=1, sticky="e")

        # Frame para controles del programador
        controls_frame = ttk.Frame(scheduler_frame)
        controls_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        controls_frame.grid_columnconfigure(3, weight=1)

        # Botones de control
        self.start_scheduler_button = ttk.Button(controls_frame, text="▶️ Iniciar",
                                                 command=self.start_scheduler, width=10)
        self.start_scheduler_button.grid(row=0, column=0, padx=(0, 5))

        self.stop_scheduler_button = ttk.Button(controls_frame, text="⏸️ Detener",
                                                command=self.stop_scheduler, width=10)
        self.stop_scheduler_button.grid(row=0, column=1, padx=5)

        # Botón de historial
        self.history_button = ttk.Button(controls_frame, text="📜 Historial",
                                         command=self.show_execution_history, width=10)
        self.history_button.grid(row=0, column=2, padx=5)

        # Label de próxima ejecución
        self.next_execution_label = ttk.Label(controls_frame, text="",
                                              foreground="blue", font=("Arial", 9))
        self.next_execution_label.grid(row=0, column=4, sticky="e")

        # Actualizar estado inicial
        self.update_scheduler_status_display()

    def start_status_updates(self):
        """Inicia las actualizaciones periódicas del estado"""
        try:
            if not self.is_destroyed:
                self.update_scheduler_status_display()
                # Programar la próxima actualización en 5 segundos
                self.status_update_job = self.parent.after(5000, self.start_status_updates)
        except Exception as e:
            print(f"Error en actualización de estado: {e}")
            # Intentar reprogramar en caso de error
            if not self.is_destroyed:
                self.status_update_job = self.parent.after(10000, self.start_status_updates)

    def stop_status_updates(self):
        """Detiene las actualizaciones periódicas"""
        self.is_destroyed = True
        if self.status_update_job:
            try:
                self.parent.after_cancel(self.status_update_job)
                self.status_update_job = None
            except Exception:
                pass

    def create_profiles_list(self):
        """Crea la lista de perfiles con Treeview"""
        # Frame para la lista
        list_frame = ttk.Frame(self.main_frame)
        list_frame.grid(row=2, column=0, sticky="nsew")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # Treeview para mostrar perfiles
        columns = ("name", "search_title", "status", "executions", "last_execution")
        self.profiles_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)

        # Configurar columnas
        self.profiles_tree.heading("name", text="Nombre")
        self.profiles_tree.heading("search_title", text="Criterio Busqueda")
        self.profiles_tree.heading("status", text="Estado")
        self.profiles_tree.heading("executions", text="Correos Encontrados")
        self.profiles_tree.heading("last_execution", text="Ultima Ejecucion")

        # Ancho de columnas
        self.profiles_tree.column("name", width=150, anchor="w")
        self.profiles_tree.column("search_title", width=200, anchor="w")
        self.profiles_tree.column("status", width=80, anchor="center")
        self.profiles_tree.column("executions", width=120, anchor="center")
        self.profiles_tree.column("last_execution", width=130, anchor="center")

        # Scrollbars
        scrollbar_v = ttk.Scrollbar(list_frame, orient="vertical", command=self.profiles_tree.yview)
        self.profiles_tree.configure(yscrollcommand=scrollbar_v.set)

        scrollbar_h = ttk.Scrollbar(list_frame, orient="horizontal", command=self.profiles_tree.xview)
        self.profiles_tree.configure(xscrollcommand=scrollbar_h.set)

        # Posicionar widgets
        self.profiles_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_v.grid(row=0, column=1, sticky="ns")
        scrollbar_h.grid(row=1, column=0, sticky="ew")

        # Eventos
        self.profiles_tree.bind("<Double-1>", self.on_double_click)
        self.profiles_tree.bind("<Button-3>", self.show_context_menu)

    def refresh_profiles_list(self):
        """Actualiza la lista de perfiles"""
        try:
            # Limpiar lista actual
            for item in self.profiles_tree.get_children():
                self.profiles_tree.delete(item)

            # Limpiar mapeo
            self.profile_items.clear()

            # Obtener perfiles y estadísticas
            profiles_stats = self.profile_service.get_all_profiles_stats()

            if not profiles_stats:
                self.show_info("No hay perfiles creados", "info")
                return

            # Agregar perfiles a la lista
            for profile_id, data in profiles_stats.items():
                try:
                    profile = data.get("profile", {})
                    stats = data.get("stats", {})

                    name = profile.get("name", "Sin nombre")
                    search_title = profile.get("search_title", "Sin criterio")
                    is_active = profile.get("is_active", True)
                    status = "Activo" if is_active else "Inactivo"

                    executions = stats.get("current_emails_found", stats.get("total_emails_found", 0))
                    last_execution = stats.get("last_execution")

                    # Formatear fecha de última ejecución
                    if last_execution:
                        try:
                            from datetime import datetime
                            last_exec_dt = datetime.fromisoformat(last_execution.replace('Z', '+00:00'))
                            last_exec_str = last_exec_dt.strftime("%d/%m/%Y %H:%M")
                        except Exception:
                            last_exec_str = "Error fecha"
                    else:
                        last_exec_str = "Nunca"

                    # Insertar en el tree
                    item_id = self.profiles_tree.insert("", "end",
                                                        values=(name, search_title, status, executions, last_exec_str))

                    # Mapear item_id a profile_id
                    self.profile_items[item_id] = profile_id

                except Exception as e:
                    error_msg = self.clean_string(str(e))
                    self.show_info(f"Error procesando perfil {profile_id}: {error_msg}", "error")
                    continue

            self.show_info(f"Lista actualizada: {len(profiles_stats)} perfil(es) cargado(s)", "info")

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"Error actualizando lista de perfiles: {error_msg}", "error")

    def start_scheduler_if_configured(self):
        """Inicia el programador si está configurado y habilitado"""
        try:
            config = self.scheduler_service.load_configuration()
            if config and config.get("enabled", False):
                success, message = self.scheduler_service.start_scheduler(self.on_scheduler_status_update)
                if success:
                    self.show_info("🟢 Programador automático iniciado", "success")
                else:
                    self.show_info(f"⚠️ No se pudo iniciar el programador: {message}", "warning")

            # Actualizar display
            self.update_scheduler_status_display()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"⚠️ Error verificando programador: {error_msg}", "warning")

    def open_scheduler_modal(self):
        """Abre el modal de configuración del programador"""
        try:
            from gui.scheduler_modal import SchedulerModal
            SchedulerModal(self.parent, self.on_scheduler_configured)
        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"❌ Error abriendo configuración del programador: {error_msg}", "error")
            messagebox.showerror("Error", f"No se pudo abrir la configuración: {error_msg}")

    def on_scheduler_configured(self, success, message):
        """Callback llamado cuando se configura el programador"""
        try:
            clean_message = self.clean_string(message)
            status = "success" if success else "error"
            symbol = "✓" if success else "✗"

            self.show_info(f"{symbol} Programador: {clean_message}", status)
            self.update_scheduler_status_display()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"❌ Error procesando configuración del programador: {error_msg}", "error")

    def on_scheduler_status_update(self, message, status):
        """Callback para actualizaciones de estado del programador"""
        try:
            self.show_info(f"🤖 {message}", status)
            # Programar actualización en el hilo principal
            if not self.is_destroyed:
                self.parent.after(100, self.update_scheduler_status_display)
                self.parent.after(500, self.refresh_profiles_list)
        except Exception as e:
            print(f"Error en callback del programador: {e}")

    def start_scheduler(self):
        """Inicia el programador manualmente"""
        try:
            success, message = self.scheduler_service.start_scheduler(self.on_scheduler_status_update)

            if success:
                self.show_info(f"✓ {message}", "success")
            else:
                self.show_info(f"✗ {message}", "error")
                messagebox.showerror("Error", message)

            self.update_scheduler_status_display()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"❌ Error iniciando programador: {error_msg}", "error")
            messagebox.showerror("Error", f"Error iniciando programador: {error_msg}")

    def stop_scheduler(self):
        """Detiene el programador manualmente"""
        try:
            success, message = self.scheduler_service.stop_scheduler()

            if success:
                self.show_info(f"✓ {message}", "success")
            else:
                self.show_info(f"✗ {message}", "error")

            self.update_scheduler_status_display()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"❌ Error deteniendo programador: {error_msg}", "error")

    def show_execution_history(self):
        """Muestra el historial de ejecuciones automáticas"""
        try:
            history = self.scheduler_service.get_execution_history(20)

            if not history:
                messagebox.showinfo("Historial", "No hay ejecuciones automáticas registradas")
                return

            # Crear ventana de historial
            history_window = tk.Toplevel(self.parent)
            history_window.title("Historial de Ejecuciones Automáticas")
            history_window.geometry("600x400")
            history_window.transient(self.parent)
            history_window.grab_set()

            # Centrar ventana
            history_window.update_idletasks()
            x = (history_window.winfo_screenwidth() // 2) - (600 // 2)
            y = (history_window.winfo_screenheight() // 2) - (400 // 2)
            history_window.geometry(f"600x400+{x}+{y}")

            # Frame principal
            main_frame = ttk.Frame(history_window, padding="15")
            main_frame.pack(fill="both", expand=True)

            # Título
            title_label = ttk.Label(main_frame, text="📜 Historial de Ejecuciones Automáticas",
                                    font=("Arial", 14, "bold"))
            title_label.pack(pady=(0, 15))

            # Lista de historial
            list_frame = ttk.Frame(main_frame)
            list_frame.pack(fill="both", expand=True, pady=(0, 15))

            columns = ("timestamp", "profiles", "emails", "status")
            history_tree = ttk.Treeview(list_frame, columns=columns, show="headings")

            history_tree.heading("timestamp", text="Fecha y Hora")
            history_tree.heading("profiles", text="Perfiles")
            history_tree.heading("emails", text="Correos Encontrados")
            history_tree.heading("status", text="Estado")

            history_tree.column("timestamp", width=150, anchor="w")
            history_tree.column("profiles", width=100, anchor="center")
            history_tree.column("emails", width=120, anchor="center")
            history_tree.column("status", width=150, anchor="center")

            # Scrollbar
            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=history_tree.yview)
            history_tree.configure(yscrollcommand=scrollbar.set)

            history_tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # Agregar datos del historial (más recientes primero)
            for record in reversed(history):
                timestamp = record.get("timestamp", "")
                successful = record.get("successful_profiles", 0)
                failed = record.get("failed_profiles", 0)
                total_profiles = record.get("total_profiles", successful + failed)
                emails = record.get("total_emails", 0)

                try:
                    from datetime import datetime
                    dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
                    formatted_time = dt.strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    formatted_time = timestamp

                profiles_str = f"{successful}/{total_profiles}"
                status_str = "✓ Éxito" if successful > 0 else "⚠ Sin resultados" if failed == 0 else "✗ Error"

                history_tree.insert("", "end", values=(formatted_time, profiles_str, emails, status_str))

            # Botón cerrar
            ttk.Button(main_frame, text="Cerrar", command=history_window.destroy).pack()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            messagebox.showerror("Error", f"Error mostrando historial: {error_msg}")

    def update_scheduler_status_display(self):
        """Actualiza la visualización del estado del programador"""
        try:
            if self.is_destroyed:
                return

            status = self.scheduler_service.get_scheduler_status()

            # Actualizar timestamp de última actualización
            from datetime import datetime
            current_time = datetime.now().strftime("%H:%M:%S")
            self.last_update_label.config(text=f"Actualizado: {current_time}")

            # Actualizar label de estado principal
            if status["is_running"] and status["enabled"]:
                status_text = f"🟢 ACTIVO - {status.get('schedule_details', 'Sin detalles')}"
                if status.get("jobs_count", 0) > 0:
                    status_text += f" ({status['jobs_count']} jobs)"
                status_color = "green"
                start_state = "disabled"
                stop_state = "normal"
            elif status["has_config"] and status["enabled"]:
                status_text = f"🟡 CONFIGURADO - Detenido"
                status_color = "orange"
                start_state = "normal"
                stop_state = "disabled"
            elif status["has_config"]:
                status_text = "🔴 DESHABILITADO"
                status_color = "red"
                start_state = "disabled"
                stop_state = "disabled"
            else:
                status_text = "⚪ SIN CONFIGURAR"
                status_color = "gray"
                start_state = "disabled"
                stop_state = "disabled"

            self.scheduler_status_label.config(text=status_text, foreground=status_color)

            # Actualizar botones
            self.start_scheduler_button.config(state=start_state)
            self.stop_scheduler_button.config(state=stop_state)

            # Actualizar próxima ejecución
            next_execution = status.get("next_execution")
            if next_execution and status["is_running"]:
                self.next_execution_label.config(text=f"Próxima: {next_execution}")
            else:
                self.next_execution_label.config(text="")

        except Exception as e:
            if not self.is_destroyed:
                error_msg = self.clean_string(str(e))
                self.scheduler_status_label.config(text=f"❌ Error: {error_msg}", foreground="red")

    def execute_manual_search(self):
        """Ejecuta búsquedas manuales inmediatas"""
        try:
            # Verificar credenciales
            if not self.config_service.credentials_exist():
                messagebox.showwarning("Advertencia",
                                       "Configure primero las credenciales de email desde el menú 'Configurar Email SMTP'")
                return

            # Obtener perfiles activos
            active_profiles = self.profile_service.get_active_profiles()
            if not active_profiles:
                messagebox.showinfo("Información", "No hay perfiles activos para ejecutar")
                return

            # Confirmar ejecución
            result = messagebox.askyesno("Confirmar Ejecución Manual",
                                         f"¿Desea ejecutar búsquedas manuales para {len(active_profiles)} perfil(es) activo(s)?")
            if not result:
                return

            # Deshabilitar botón y mostrar estado
            self.manual_search_button.config(state="disabled", text="Ejecutando...")
            self.show_info(f"🔄 Iniciando búsqueda manual para {len(active_profiles)} perfil(es)...", "info")

            # Guardar perfiles para el hilo
            self.current_active_profiles = active_profiles

            # Ejecutar en hilo separado
            thread = threading.Thread(target=self._execute_manual_search_thread, daemon=True)
            thread.start()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"✗ Error iniciando búsqueda manual: {error_msg}", "error")
            self.manual_search_button.config(state="normal", text="🔍 Buscar Ahora")

    def _execute_manual_search_thread(self):
        """Ejecuta las búsquedas manuales en un hilo separado"""
        try:
            # Cargar credenciales
            credentials = self.config_service.load_credentials()
            if not credentials:
                self.parent.after(0, self._update_manual_search_result, False, "No se pudieron cargar las credenciales")
                return

            # Usar perfiles guardados
            active_profiles = getattr(self, 'current_active_profiles', [])
            if not active_profiles:
                self.parent.after(0, self._update_manual_search_result, False, "No hay perfiles activos")
                return

            # Notificar progreso
            self.parent.after(0, lambda: self.show_info("🔍 Conectando a servidor de email...", "info"))

            # Ejecutar búsquedas
            results = self.email_search_service.search_multiple_profiles(active_profiles, credentials)

            # Procesar resultados
            total_emails = 0
            successful_profiles = 0
            failed_profiles = []

            for profile_id, result in results.items():
                profile_name = result.get("profile_name", "Desconocido")
                if result["success"]:
                    emails_found = result["emails_found"]
                    self.profile_service.update_profile_execution(profile_id, emails_found)
                    total_emails += emails_found
                    successful_profiles += 1
                else:
                    failed_profiles.append(f"{profile_name}: {result['message']}")

            # Preparar mensaje de resultado
            success_msg = f"✓ Búsqueda manual completada: {successful_profiles}/{len(results)} perfiles exitosos, {total_emails} correos encontrados"

            if failed_profiles and len(failed_profiles) <= 3:
                success_msg += f"\n\n❌ Errores:\n" + "\n".join(failed_profiles[:3])
            elif failed_profiles:
                success_msg += f"\n\n❌ {len(failed_profiles)} perfiles con errores"

            # Actualizar UI en el hilo principal
            overall_success = successful_profiles > 0
            self.parent.after(0, self._update_manual_search_result, overall_success, success_msg)

            # Actualizar lista de perfiles
            self.parent.after(0, self.refresh_profiles_list)

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Error en búsqueda manual: {error_details}")

            error_msg = self.clean_string(str(e))
            self.parent.after(0, self._update_manual_search_result, False,
                              f"Error crítico en búsqueda manual: {error_msg}")
        finally:
            # Limpiar perfiles temporales
            if hasattr(self, 'current_active_profiles'):
                delattr(self, 'current_active_profiles')

    def _update_manual_search_result(self, success, message):
        """Actualiza la UI con el resultado de la búsqueda manual"""
        try:
            # Rehabilitar botón
            self.manual_search_button.config(state="normal", text="🔍 Buscar Ahora")

            # Mostrar mensaje
            status = "success" if success else "error"
            self.show_info(message, status)

            # Mostrar ventana de resultado
            if success:
                messagebox.showinfo("Búsqueda Manual Completada", message)
            else:
                messagebox.showerror("Error en Búsqueda Manual", message)

        except Exception as e:
            print(f"Error actualizando resultado de búsqueda manual: {e}")
            if hasattr(self, 'manual_search_button'):
                self.manual_search_button.config(state="normal", text="🔍 Buscar Ahora")

    # Métodos para gestión de perfiles (sin cambios significativos)
    def get_selected_profile_id(self):
        """Obtiene el profile_id del item seleccionado"""
        selected = self.profiles_tree.selection()
        if not selected:
            return None

        item_id = selected[0]
        return self.profile_items.get(item_id)

    def create_new_profile(self):
        """Abre el diálogo para crear un nuevo perfil"""
        try:
            dialog = ProfileDialog(self.parent, title="Nuevo Perfil")
            result = dialog.show()

            if result:
                name, search_title = result

                if not name.strip():
                    messagebox.showerror("Error", "El nombre del perfil no puede estar vacío")
                    return

                if not search_title.strip():
                    messagebox.showerror("Error", "El criterio de búsqueda no puede estar vacío")
                    return

                # Crear perfil
                profile_id = self.profile_service.create_profile(name.strip(), search_title.strip())

                # Actualizar lista
                self.refresh_profiles_list()
                self.show_info(f"✓ Perfil '{name}' creado correctamente", "success")

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"✗ Error creando perfil: {error_msg}", "error")
            messagebox.showerror("Error", error_msg)

    def edit_selected_profile(self):
        """Edita el perfil seleccionado"""
        profile_id = self.get_selected_profile_id()
        if not profile_id:
            messagebox.showwarning("Advertencia", "Seleccione un perfil para editar")
            return

        try:
            # Obtener datos actuales del perfil
            profile = self.profile_service.get_profile(profile_id)
            if not profile:
                messagebox.showerror("Error", "Perfil no encontrado")
                return

            # Mostrar diálogo de edición
            dialog = ProfileDialog(self.parent, title="Editar Perfil",
                                   name=profile["name"], search_title=profile["search_title"])
            result = dialog.show()

            if result:
                name, search_title = result

                if not name.strip():
                    messagebox.showerror("Error", "El nombre del perfil no puede estar vacío")
                    return

                if not search_title.strip():
                    messagebox.showerror("Error", "El criterio de búsqueda no puede estar vacío")
                    return

                # Actualizar perfil
                self.profile_service.update_profile(profile_id, name=name.strip(), search_title=search_title.strip())

                # Actualizar lista
                self.refresh_profiles_list()
                self.show_info(f"✓ Perfil '{name}' actualizado correctamente", "success")

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"✗ Error editando perfil: {error_msg}", "error")
            messagebox.showerror("Error", error_msg)

    def delete_selected_profile(self):
        """Elimina el perfil seleccionado"""
        profile_id = self.get_selected_profile_id()
        if not profile_id:
            messagebox.showwarning("Advertencia", "Seleccione un perfil para eliminar")
            return

        try:
            # Obtener nombre del perfil
            profile = self.profile_service.get_profile(profile_id)
            if not profile:
                messagebox.showerror("Error", "Perfil no encontrado")
                return

            profile_name = profile.get("name", "Desconocido")

            # Confirmar eliminación
            result = messagebox.askyesno("Confirmar Eliminación",
                                         f"¿Está seguro de eliminar el perfil '{profile_name}'?\n\n"
                                         f"Esta acción no se puede deshacer.")
            if result:
                self.profile_service.delete_profile(profile_id)
                self.refresh_profiles_list()
                self.show_info(f"✓ Perfil '{profile_name}' eliminado correctamente", "success")

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"✗ Error eliminando perfil: {error_msg}", "error")
            messagebox.showerror("Error", error_msg)

    def on_double_click(self, event):
        """Maneja el doble clic en un perfil"""
        self.edit_selected_profile()

    def show_context_menu(self, event):
        """Muestra menú contextual para los perfiles"""
        # Seleccionar item bajo el cursor
        item = self.profiles_tree.identify_row(event.y)
        if item:
            self.profiles_tree.selection_set(item)

            # Crear menú contextual
            context_menu = tk.Menu(self.parent, tearoff=0)
            context_menu.add_command(label="✏️ Editar", command=self.edit_selected_profile)
            context_menu.add_command(label="🗑️ Eliminar", command=self.delete_selected_profile)
            context_menu.add_separator()
            context_menu.add_command(label="⚡ Activar/Desactivar", command=self.toggle_profile_status)

            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

    def toggle_profile_status(self):
        """Activa o desactiva el perfil seleccionado"""
        profile_id = self.get_selected_profile_id()
        if not profile_id:
            return

        try:
            # Obtener estado actual
            profile = self.profile_service.get_profile(profile_id)
            if not profile:
                return

            new_status = not profile.get("is_active", True)
            self.profile_service.update_profile(profile_id, is_active=new_status)

            status_text = "activado" if new_status else "desactivado"
            self.show_info(f"✓ Perfil '{profile['name']}' {status_text}", "success")
            self.refresh_profiles_list()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"✗ Error cambiando estado: {error_msg}", "error")
            messagebox.showerror("Error", error_msg)

    def show_info(self, message, status="info"):
        """Muestra información usando el callback"""
        if self.info_callback:
            try:
                self.info_callback(message, status)
            except:
                # Si falla, usar after para ejecutar en el hilo principal
                if not self.is_destroyed:
                    self.parent.after(0, lambda: self.info_callback(message, status) if self.info_callback else None)

    def clean_string(self, text):
        """Limpia un string de caracteres problemáticos"""
        if not text:
            return ""

        try:
            # Reemplazar caracteres problemáticos comunes
            replacements = {
                '\xa0': ' ',  # Espacio no-rompible
                '\u2019': "'",  # Apostrofe curvo
                '\u2018': "'",  # Apostrofe curvo
                '\u201c': '"',  # Comilla curva
                '\u201d': '"',  # Comilla curva
                '\u2013': '-',  # En dash
                '\u2014': '--',  # Em dash
                '\u2026': '...'  # Ellipsis
            }

            for char, replacement in replacements.items():
                text = text.replace(char, replacement)

            # Codificar y decodificar para limpiar caracteres problemáticos
            return text.encode('ascii', 'ignore').decode('ascii')
        except Exception:
            return str(text)

    def destroy(self):
        """Método para limpiar recursos al cerrar"""
        self.stop_status_updates()


class ProfileDialog:
    def __init__(self, parent, title="Perfil", name="", search_title=""):
        """
        Inicializa el diálogo de perfil

        Args:
            parent: Ventana padre
            title (str): Título del diálogo
            name (str): Nombre inicial del perfil
            search_title (str): Título de búsqueda inicial
        """
        self.parent = parent
        self.title = title
        self.name = name
        self.search_title = search_title
        self.result = None

    def show(self):
        """
        Muestra el diálogo y retorna el resultado

        Returns:
            tuple: (name, search_title) o None si se canceló
        """
        # Crear ventana modal
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title(self.title)
        self.dialog.geometry("550x400")
        self.dialog.resizable(True, True)
        self.dialog.minsize(500, 350)

        # Hacer modal
        self.dialog.transient(self.parent)
        self.dialog.grab_set()

        # Centrar diálogo
        self.center_dialog()

        # Crear contenido
        self.create_content()

        # Manejar cierre
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)

        # Enfocar primer campo
        self.name_entry.focus_set()

        # Esperar hasta que se cierre
        self.dialog.wait_window()

        return self.result

    def center_dialog(self):
        """Centra el diálogo en la pantalla"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (550 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (400 // 2)
        self.dialog.geometry(f"550x400+{x}+{y}")

    def create_content(self):
        """Crea el contenido del diálogo"""
        # Frame principal
        main_frame = ttk.Frame(self.dialog, padding="25")
        main_frame.pack(fill="both", expand=True)

        # Título
        title_label = ttk.Label(main_frame, text=self.title, font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 25))

        # Variables para los campos
        self.name_var = tk.StringVar(value=self.name)
        self.search_title_var = tk.StringVar(value=self.search_title)

        # Frame para los campos
        fields_frame = ttk.LabelFrame(main_frame, text="Información del Perfil", padding="20")
        fields_frame.pack(fill="both", expand=True, pady=(0, 20))

        # Campo nombre
        ttk.Label(fields_frame, text="Nombre del Perfil:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 8))
        self.name_entry = ttk.Entry(fields_frame, textvariable=self.name_var, font=("Arial", 10))
        self.name_entry.pack(fill="x", pady=(0, 20), ipady=5)

        # Campo criterio de búsqueda
        ttk.Label(fields_frame, text="Criterio de Búsqueda (palabras clave en el asunto):",
                  font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 8))
        self.search_entry = ttk.Entry(fields_frame, textvariable=self.search_title_var, font=("Arial", 10))
        self.search_entry.pack(fill="x", pady=(0, 15), ipady=5)

        # Nota informativa
        info_label = ttk.Label(fields_frame,
                               text="💡 Ejemplo: 'factura', 'pedido', 'confirmación'",
                               foreground="gray", font=("Arial", 9, "italic"))
        info_label.pack(anchor="w")

        # Botones
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x")

        # Frame interno para centrar botones
        center_buttons_frame = ttk.Frame(buttons_frame)
        center_buttons_frame.pack(expand=True)

        accept_btn = ttk.Button(center_buttons_frame, text="✓ Aceptar",
                                command=self.accept, width=12)
        accept_btn.pack(side="left", padx=(0, 15))

        cancel_btn = ttk.Button(center_buttons_frame, text="✗ Cancelar",
                                command=self.cancel, width=12)
        cancel_btn.pack(side="left")

        # Bind Enter para aceptar
        self.dialog.bind("<Return>", lambda e: self.accept())
        self.dialog.bind("<Escape>", lambda e: self.cancel())

    def accept(self):
        """Acepta el diálogo"""
        name = self.name_var.get().strip()
        search_title = self.search_title_var.get().strip()

        # Validaciones
        if not name:
            messagebox.showerror("Error", "El nombre del perfil es requerido", parent=self.dialog)
            self.name_entry.focus_set()
            return

        if len(name) < 2:
            messagebox.showerror("Error", "El nombre debe tener al menos 2 caracteres", parent=self.dialog)
            self.name_entry.focus_set()
            return

        if not search_title:
            messagebox.showerror("Error", "El criterio de búsqueda es requerido", parent=self.dialog)
            self.search_entry.focus_set()
            return

        if len(search_title) < 2:
            messagebox.showerror("Error", "El criterio de búsqueda debe tener al menos 2 caracteres",
                                 parent=self.dialog)
            self.search_entry.focus_set()
            return

        self.result = (name, search_title)
        self.dialog.destroy()

    def cancel(self):
        """Cancela el diálogo"""
        self.result = None
        self.dialog.destroy()