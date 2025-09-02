# services/scheduler_service.py
"""
Servicio para programación automática de búsquedas de correos con envío automático de reportes.
Maneja la configuración de horarios, intervalos, ejecución automática y envío por correo de reportes Excel.
"""

import json
import threading
import time
from pathlib import Path
import schedule


class SchedulerService:
    def __init__(self, config_dir="config"):
        """
        Inicializa el servicio de programación

        Args:
            config_dir (str): Directorio donde guardar las configuraciones
        """
        self.config_dir = Path(config_dir)
        self.scheduler_file = self.config_dir / "scheduler_config.json"
        self.ensure_config_directory()

        self.scheduler_thread = None
        self.is_running = False
        self.current_config = None
        self.status_callback = None

        # Cargar configuración existente
        self.load_configuration()

    def ensure_config_directory(self):
        """Asegura que el directorio de configuración existe"""
        self.config_dir.mkdir(exist_ok=True)

    def save_configuration(self, config):
        """
        Guarda la configuración del programador

        Args:
            config (dict): Configuración de programación

        Raises:
            Exception: Si hay error guardando la configuración
        """
        try:
            # Validar configuración
            if not self._validate_config(config):
                raise Exception("Configuración de programación inválida")

            # Limpiar strings de caracteres problemáticos
            clean_config = self._clean_config(config)

            # Guardar en archivo JSON
            with open(self.scheduler_file, "w", encoding="utf-8") as f:
                json.dump(clean_config, f, indent=4, ensure_ascii=True)

            self.current_config = clean_config
            print(f"DEBUG: Configuración guardada: {clean_config}")

        except Exception as e:
            error_msg = self._clean_string(str(e))
            raise Exception(f"Error guardando configuración del programador: {error_msg}")

    def load_configuration(self):
        """
        Carga la configuración del programador

        Returns:
            dict: Configuración o None si no existe
        """
        try:
            if not self.scheduler_file.exists():
                return None

            with open(self.scheduler_file, "r", encoding="utf-8") as f:
                config = json.load(f)

            # Limpiar strings de caracteres problemáticos
            self.current_config = self._clean_config(config)
            print(f"DEBUG: Configuración cargada: {self.current_config}")
            return self.current_config

        except Exception as e:
            error_msg = self._clean_string(str(e))
            print(f"Error cargando configuración del programador: {error_msg}")
            return None

    def start_scheduler(self, status_callback=None):
        """
        Inicia el programador de tareas

        Args:
            status_callback: Función callback para reportar estado
        """
        try:
            if self.is_running:
                return True, "El programador ya está en ejecución"

            if not self.current_config:
                return False, "No hay configuración de programación"

            if not self.current_config.get("enabled", False):
                return False, "Programación deshabilitada"

            self.status_callback = status_callback

            # Configurar schedule según la configuración
            success = self._configure_schedule()
            if not success:
                return False, "Error configurando horarios"

            self.is_running = True

            # Iniciar hilo del programador
            self.scheduler_thread = threading.Thread(target=self._scheduler_loop, daemon=True)
            self.scheduler_thread.start()

            print(f"DEBUG: Programador iniciado. Jobs programados: {len(schedule.jobs)}")

            # Mostrar próximas ejecuciones para debug
            if schedule.jobs:
                next_run = schedule.next_run()
                print(f"DEBUG: Próxima ejecución: {next_run}")

            return True, "Programador iniciado correctamente"

        except Exception as e:
            self.is_running = False
            error_msg = self._clean_string(str(e))
            print(f"ERROR: {error_msg}")
            return False, f"Error iniciando programador: {error_msg}"

    def stop_scheduler(self):
        """
        Detiene el programador de tareas

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            if not self.is_running:
                return True, "El programador no está en ejecución"

            self.is_running = False
            schedule.clear()
            print("DEBUG: Schedule jobs cleared, programador detenido")

            # Esperar a que termine el hilo (máximo 3 segundos)
            if self.scheduler_thread and self.scheduler_thread.is_alive():
                self.scheduler_thread.join(timeout=3)

            return True, "Programador detenido correctamente"

        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error deteniendo programador: {error_msg}"

    def get_scheduler_status(self):
        """
        Obtiene el estado actual del programador

        Returns:
            dict: Estado del programador
        """
        try:
            status = {
                "is_running": self.is_running,
                "has_config": self.current_config is not None,
                "enabled": self.current_config.get("enabled", False) if self.current_config else False,
                "next_execution": None,
                "schedule_type": None,
                "schedule_details": None,
                "jobs_count": len(schedule.jobs)
            }

            if self.current_config and self.current_config.get("enabled", False):
                status["schedule_type"] = self.current_config.get("type", "unknown")
                status["schedule_details"] = self._get_schedule_description()

                if self.is_running and schedule.jobs:
                    next_execution = self._get_next_execution_time()
                    status["next_execution"] = next_execution

            return status
        except Exception as e:
            print(f"Error obteniendo estado del programador: {e}")
            return {
                "is_running": False,
                "has_config": False,
                "enabled": False,
                "next_execution": None,
                "schedule_type": None,
                "schedule_details": None,
                "jobs_count": 0
            }

    def execute_scheduled_search(self):
        """
        Ejecuta una búsqueda programada con generación y envío automático de reporte

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            print("DEBUG: Ejecutando búsqueda programada...")

            from services.profile_service import ProfileService
            from services.email_search_service import EmailSearchService
            from services.config_service import ConfigService
            from services.excel_service import ExcelService
            from services.email_send_service import EmailSendService

            config_service = ConfigService()

            # Verificar credenciales SMTP
            if not config_service.credentials_exist():
                error_msg = "No hay credenciales SMTP configuradas"
                print(f"ERROR: {error_msg}")
                self._notify_callback(f"Error en búsqueda automática: {error_msg}", "error")
                return False, error_msg

            credentials = config_service.load_credentials()
            if not credentials:
                error_msg = "No se pudieron cargar las credenciales SMTP"
                print(f"ERROR: {error_msg}")
                self._notify_callback(f"Error en búsqueda automática: {error_msg}", "error")
                return False, error_msg

            # Obtener perfiles activos
            profile_service = ProfileService()
            active_profiles = profile_service.get_active_profiles()

            if not active_profiles:
                error_msg = "No hay perfiles activos"
                print(f"WARNING: {error_msg}")
                self._notify_callback(f"Búsqueda automática: {error_msg}", "warning")
                return False, error_msg

            print(f"DEBUG: Ejecutando búsqueda para {len(active_profiles)} perfiles activos")
            self._notify_callback(f"Iniciando búsqueda automática para {len(active_profiles)} perfiles...", "info")

            # Ejecutar búsquedas
            search_service = EmailSearchService()
            results = search_service.search_multiple_profiles(active_profiles, credentials)

            # Procesar resultados
            total_emails = 0
            successful_profiles = 0
            failed_profiles = 0

            for profile_id, result in results.items():
                if result["success"]:
                    emails_found = result["emails_found"]
                    profile_service.update_profile_execution(profile_id, emails_found)
                    total_emails += emails_found
                    successful_profiles += 1
                    print(f"DEBUG: Perfil {profile_id} - {emails_found} correos encontrados")
                else:
                    failed_profiles += 1
                    print(f"DEBUG: Perfil {profile_id} - Error: {result.get('message', 'Unknown')}")

            # Preparar mensaje de resultado de búsqueda
            search_message = f"Búsqueda automática completada: {successful_profiles}/{len(results)} perfiles exitosos, {total_emails} correos encontrados"
            print(f"DEBUG: {search_message}")

            if successful_profiles == 0:
                self._notify_callback("Búsqueda automática completada sin resultados", "warning")
                return True, "Búsqueda completada sin resultados"

            # NUEVA FUNCIONALIDAD: Generar y enviar reporte automáticamente
            try:
                self._notify_callback("Generando reporte Excel...", "info")

                # Generar reporte Excel
                excel_service = ExcelService()
                profiles_stats = profile_service.get_all_profiles_stats()

                if not profiles_stats:
                    self._notify_callback("No hay datos para generar reporte", "warning")
                    return True, f"{search_message}. Reporte no generado: sin datos"

                report_path = excel_service.generate_profiles_report(profiles_stats)
                print(f"DEBUG: Reporte generado: {report_path}")
                self._notify_callback(f"Reporte Excel generado: {Path(report_path).name}", "success")

                # Verificar configuración de envío
                email_send_status = config_service.get_email_send_status()
                if not email_send_status["ready"]:
                    final_message = f"{search_message}. Reporte generado pero no enviado (configuración de envío no lista)"
                    self._notify_callback("Reporte no enviado: configuración de envío incompleta", "warning")
                    return True, final_message

                # Enviar reporte por correo
                self._notify_callback("Enviando reporte por correo...", "info")
                email_send_service = EmailSendService()

                send_success, send_message = email_send_service.send_report_email(report_path)

                if send_success:
                    final_message = f"{search_message}. Reporte generado y enviado por correo exitosamente"
                    self._notify_callback(f"Reporte enviado por correo: {send_message}", "success")
                    print(f"DEBUG: {final_message}")
                    return True, final_message
                else:
                    final_message = f"{search_message}. Reporte generado pero error enviando: {send_message}"
                    self._notify_callback(f"Error enviando reporte: {send_message}", "error")
                    print(f"ERROR: {final_message}")
                    return True, final_message  # True porque la búsqueda fue exitosa

            except Exception as report_error:
                report_error_msg = self._clean_string(str(report_error))
                final_message = f"{search_message}. Error generando/enviando reporte: {report_error_msg}"
                self._notify_callback(f"Error con reporte: {report_error_msg}", "error")
                print(f"ERROR generando/enviando reporte: {report_error_msg}")
                return True, final_message  # True porque la búsqueda fue exitosa

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR en búsqueda automática: {error_details}")

            error_msg = self._clean_string(str(e))
            self._notify_callback(f"Error en búsqueda automática: {error_msg}", "error")
            return False, f"Error en ejecución automática: {error_msg}"

    def _configure_schedule(self):
        """Configura el programador según la configuración actual"""
        try:
            schedule.clear()

            if not self.current_config:
                return False

            schedule_type = self.current_config.get("type", "")
            print(f"DEBUG: Configurando schedule tipo: {schedule_type}")

            if schedule_type == "daily":
                time_str = self.current_config.get("time", "09:00")
                schedule.every().day.at(time_str).do(self._job_wrapper)
                print(f"DEBUG: Programado diario a las {time_str}")

            elif schedule_type == "weekly":
                time_str = self.current_config.get("time", "09:00")
                days = self.current_config.get("days", ["monday"])

                for day in days:
                    if day == "monday":
                        schedule.every().monday.at(time_str).do(self._job_wrapper)
                    elif day == "tuesday":
                        schedule.every().tuesday.at(time_str).do(self._job_wrapper)
                    elif day == "wednesday":
                        schedule.every().wednesday.at(time_str).do(self._job_wrapper)
                    elif day == "thursday":
                        schedule.every().thursday.at(time_str).do(self._job_wrapper)
                    elif day == "friday":
                        schedule.every().friday.at(time_str).do(self._job_wrapper)
                    elif day == "saturday":
                        schedule.every().saturday.at(time_str).do(self._job_wrapper)
                    elif day == "sunday":
                        schedule.every().sunday.at(time_str).do(self._job_wrapper)

                print(f"DEBUG: Programado semanal para {days} a las {time_str}")

            elif schedule_type == "interval":
                interval = self.current_config.get("interval", 60)
                unit = self.current_config.get("unit", "minutes")

                if unit == "minutes":
                    schedule.every(interval).minutes.do(self._job_wrapper)
                    print(f"DEBUG: Programado cada {interval} minutos")
                elif unit == "hours":
                    schedule.every(interval).hours.do(self._job_wrapper)
                    print(f"DEBUG: Programado cada {interval} horas")

            print(f"DEBUG: Jobs programados: {len(schedule.jobs)}")
            return True

        except Exception as e:
            print(f"ERROR configurando programador: {e}")
            return False

    def _job_wrapper(self):
        """Wrapper para ejecutar el job en un hilo separado"""
        try:
            print("DEBUG: Job wrapper ejecutado")
            job_thread = threading.Thread(target=self.execute_scheduled_search, daemon=True)
            job_thread.start()
            return schedule.CancelJob  # Para jobs que se ejecutan una vez
        except Exception as e:
            print(f"ERROR en job wrapper: {e}")

    def _scheduler_loop(self):
        """Loop principal del programador"""
        try:
            print("DEBUG: Iniciando loop del programador")
            while self.is_running:
                try:
                    schedule.run_pending()
                    time.sleep(1)
                except Exception as e:
                    print(f"ERROR en loop del programador: {e}")
                    time.sleep(5)  # Esperar un poco más si hay error

            print("DEBUG: Loop del programador terminado")
        except Exception as e:
            print(f"ERROR crítico en loop del programador: {e}")
            self.is_running = False

    def _notify_callback(self, message, status):
        """Notifica al callback de manera segura"""
        try:
            if self.status_callback:
                self.status_callback(message, status)
        except Exception as e:
            print(f"Error notificando callback: {e}")

    def _get_schedule_description(self):
        """Obtiene descripción legible del horario programado"""
        if not self.current_config:
            return "Sin configuración"

        schedule_type = self.current_config.get("type", "")

        if schedule_type == "daily":
            time_str = self.current_config.get("time", "09:00")
            return f"Diario a las {time_str}"

        elif schedule_type == "weekly":
            time_str = self.current_config.get("time", "09:00")
            days = self.current_config.get("days", [])
            days_spanish = {
                "monday": "Lunes",
                "tuesday": "Martes",
                "wednesday": "Miércoles",
                "thursday": "Jueves",
                "friday": "Viernes",
                "saturday": "Sábado",
                "sunday": "Domingo"
            }
            days_str = ", ".join([days_spanish.get(d, d.capitalize()) for d in days])
            return f"{days_str} a las {time_str}"

        elif schedule_type == "interval":
            interval = self.current_config.get("interval", 60)
            unit = self.current_config.get("unit", "minutes")
            unit_name = "minutos" if unit == "minutes" else "horas"
            return f"Cada {interval} {unit_name}"

        return "Configuración desconocida"

    def _get_next_execution_time(self):
        """Obtiene el tiempo de la próxima ejecución"""
        try:
            if not schedule.jobs:
                return None

            next_run = schedule.next_run()
            if next_run:
                return next_run.strftime("%d/%m/%Y %H:%M:%S")
            return None
        except Exception as e:
            print(f"Error obteniendo próxima ejecución: {e}")
            return None

    def _validate_config(self, config):
        """
        Valida la configuración del programador

        Args:
            config (dict): Configuración a validar

        Returns:
            bool: True si la configuración es válida
        """
        if not isinstance(config, dict):
            return False

        required_fields = ["type", "enabled"]
        for field in required_fields:
            if field not in config:
                return False

        schedule_type = config.get("type", "")

        if schedule_type in ["daily", "weekly"]:
            if "time" not in config:
                return False
            # Validar formato de hora (HH:MM)
            time_str = config.get("time", "")
            try:
                from datetime import datetime
                datetime.strptime(time_str, "%H:%M")
            except ValueError:
                return False

        if schedule_type == "weekly":
            if "days" not in config or not isinstance(config["days"], list):
                return False
            valid_days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
            for day in config["days"]:
                if day not in valid_days:
                    return False

        elif schedule_type == "interval":
            if "interval" not in config or "unit" not in config:
                return False
            try:
                interval = int(config["interval"])
                if interval <= 0:
                    return False
            except (ValueError, TypeError):
                return False
            if config["unit"] not in ["minutes", "hours"]:
                return False

        return True

    def _clean_config(self, config):
        """
        Limpia la configuración de caracteres problemáticos

        Args:
            config (dict): Configuración a limpiar

        Returns:
            dict: Configuración limpia
        """
        if not isinstance(config, dict):
            return config

        clean_config = {}
        for key, value in config.items():
            if isinstance(value, str):
                clean_config[key] = self._clean_string(value)
            elif isinstance(value, list):
                clean_config[key] = [self._clean_string(str(item)) if isinstance(item, str) else item for item in value]
            else:
                clean_config[key] = value

        return clean_config

    def _clean_string(self, text):
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
            return str(text) if text else ""