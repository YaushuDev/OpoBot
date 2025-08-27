# gui/profile_manager.py
"""
Interfaz gráfica para gestión de perfiles de búsqueda.
Permite crear, editar, eliminar y ejecutar búsquedas con perfiles.
"""

# Archivos relacionados: services/profile_service.py, services/email_search_service.py, services/config_service.py

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
from services.profile_service import ProfileService
from services.email_search_service import EmailSearchService
from services.config_service import ConfigService


class ProfileManager:
    def __init__(self, parent, info_callback=None):
        """
        Inicializa el gestor de perfiles

        Args:
            parent: Widget padre donde se colocará el gestor
            info_callback: Función callback para mostrar información
        """
        self.parent = parent
        self.info_callback = info_callback
        self.profile_service = ProfileService()
        self.email_search_service = EmailSearchService()
        self.config_service = ConfigService()
        self.current_active_profiles = []
        self.profile_items = {}  # Mapeo de item_id a profile_id

        self.create_profile_manager()
        self.refresh_profiles_list()

    def create_profile_manager(self):
        """Crea la interfaz del gestor de perfiles"""
        # Frame principal
        self.main_frame = ttk.LabelFrame(self.parent, text="Perfiles de Busqueda", padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Frame de botones superiores
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        buttons_frame.grid_columnconfigure(3, weight=1)

        # Botones de acción
        ttk.Button(buttons_frame, text="Nuevo Perfil",
                   command=self.create_new_profile).grid(row=0, column=0, padx=(0, 5))

        ttk.Button(buttons_frame, text="Editar",
                   command=self.edit_selected_profile).grid(row=0, column=1, padx=5)

        ttk.Button(buttons_frame, text="Eliminar",
                   command=self.delete_selected_profile).grid(row=0, column=2, padx=5)

        # Botón de ejecución
        self.execute_button = ttk.Button(buttons_frame, text="Ejecutar Busquedas",
                                         command=self.execute_searches, style="Accent.TButton")
        self.execute_button.grid(row=0, column=4, padx=(10, 0))

        # Lista de perfiles
        self.create_profiles_list()

    def create_profiles_list(self):
        """Crea la lista de perfiles con Treeview"""
        # Frame para la lista
        list_frame = ttk.Frame(self.main_frame)
        list_frame.grid(row=1, column=0, sticky="nsew")
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

        # Scrollbar vertical
        scrollbar_v = ttk.Scrollbar(list_frame, orient="vertical", command=self.profiles_tree.yview)
        self.profiles_tree.configure(yscrollcommand=scrollbar_v.set)

        # Scrollbar horizontal
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

                    executions = stats.get("total_emails_found", 0)
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

    def get_selected_profile_id(self):
        """
        Obtiene el profile_id del item seleccionado

        Returns:
            str: Profile ID o None si no hay selección
        """
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

                # Validar datos antes de crear
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

                # Validar datos
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
            # Obtener nombre del perfil para confirmación
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

    def execute_searches(self):
        """Ejecuta búsquedas para todos los perfiles activos"""
        try:
            # Verificar que existan credenciales
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
            result = messagebox.askyesno("Confirmar Ejecución",
                                         f"¿Desea ejecutar búsquedas para {len(active_profiles)} perfil(es) activo(s)?")
            if not result:
                return

            # Deshabilitar botón y mostrar estado
            self.execute_button.config(state="disabled", text="Ejecutando...")
            self.show_info(f"🔄 Iniciando búsqueda para {len(active_profiles)} perfil(es)...", "info")

            # Guardar perfiles para usar en el hilo
            self.current_active_profiles = active_profiles

            # Ejecutar búsquedas en hilo separado
            thread = threading.Thread(target=self._execute_searches_thread, daemon=True)
            thread.start()

        except Exception as e:
            error_msg = self.clean_string(str(e))
            self.show_info(f"✗ Error iniciando búsquedas: {error_msg}", "error")
            self.execute_button.config(state="normal", text="Ejecutar Busquedas")

    def _execute_searches_thread(self):
        """Ejecuta las búsquedas en un hilo separado"""
        try:
            # Cargar credenciales
            credentials = self.config_service.load_credentials()
            if not credentials:
                self.parent.after(0, self._update_execution_result, False, "No se pudieron cargar las credenciales")
                return

            # Usar los perfiles guardados
            active_profiles = getattr(self, 'current_active_profiles', [])
            if not active_profiles:
                self.parent.after(0, self._update_execution_result, False, "No hay perfiles activos")
                return

            # Notificar inicio de búsqueda
            self.parent.after(0, lambda: self.show_info("🔍 Conectando a servidor de email...", "info"))

            # Ejecutar búsquedas múltiples
            results = self.email_search_service.search_multiple_profiles(active_profiles, credentials)

            # Actualizar estadísticas de perfiles
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
            success_msg = f"✓ Búsquedas completadas: {successful_profiles}/{len(results)} perfiles exitosos, {total_emails} correos encontrados"

            if failed_profiles and len(failed_profiles) <= 3:
                success_msg += f"\n\n❌ Errores:\n" + "\n".join(failed_profiles[:3])
            elif failed_profiles:
                success_msg += f"\n\n❌ {len(failed_profiles)} perfiles con errores"

            # Actualizar UI en el hilo principal
            overall_success = successful_profiles > 0
            self.parent.after(0, self._update_execution_result, overall_success, success_msg)

            # Actualizar lista de perfiles
            self.parent.after(0, self.refresh_profiles_list)

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Error en búsquedas: {error_details}")

            error_msg = self.clean_string(str(e))
            self.parent.after(0, self._update_execution_result, False, f"Error crítico en búsquedas: {error_msg}")
        finally:
            # Limpiar perfiles temporales
            if hasattr(self, 'current_active_profiles'):
                delattr(self, 'current_active_profiles')

    def _update_execution_result(self, success, message):
        """
        Actualiza la UI con el resultado de la ejecución

        Args:
            success (bool): Si la ejecución fue exitosa
            message (str): Mensaje de resultado
        """
        try:
            # Rehabilitar botón
            self.execute_button.config(state="normal", text="Ejecutar Busquedas")

            # Mostrar mensaje
            status = "success" if success else "error"
            self.show_info(message, status)

            # Mostrar ventana de resultado
            if success:
                messagebox.showinfo("Búsquedas Completadas", message)
            else:
                messagebox.showerror("Error en Búsquedas", message)

        except Exception as e:
            print(f"Error actualizando resultado: {e}")
            if hasattr(self, 'execute_button'):
                self.execute_button.config(state="normal", text="Ejecutar Busquedas")

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
        """
        Muestra información usando el callback

        Args:
            message (str): Mensaje a mostrar
            status (str): Tipo de mensaje (info, error, success)
        """
        if self.info_callback:
            try:
                self.info_callback(message, status)
            except:
                # Si falla, usar after para ejecutar en el hilo principal
                self.parent.after(0, lambda: self.info_callback(message, status) if self.info_callback else None)

    def clean_string(self, text):
        """
        Limpia un string de caracteres problemáticos

        Args:
            text (str): Texto a limpiar

        Returns:
            str: Texto limpio
        """
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

        # Nota informativa (más pequeña que antes)
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