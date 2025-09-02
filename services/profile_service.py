# services/profile_service.py
"""
Servicio para gestión de perfiles de búsqueda de correos con soporte para búsqueda aproximada.
Maneja la creación, edición, eliminación y persistencia de perfiles con criterios flexibles.
"""

import json
import uuid
from pathlib import Path
from datetime import datetime


class ProfileService:
    def __init__(self, config_dir="config"):
        """
        Inicializa el servicio de perfiles

        Args:
            config_dir (str): Directorio donde guardar los perfiles
        """
        self.config_dir = Path(config_dir)
        self.profiles_file = self.config_dir / "profiles.json"
        self.stats_file = self.config_dir / "profile_stats.json"
        self.ensure_config_directory()

    def ensure_config_directory(self):
        """Asegura que el directorio de configuración existe"""
        self.config_dir.mkdir(exist_ok=True)

    def create_profile(self, name, search_title):
        """
        Crea un nuevo perfil de búsqueda con soporte para criterios flexibles

        Args:
            name (str): Nombre del perfil
            search_title (str): Título/criterio de búsqueda (soporta espacios y búsqueda aproximada)

        Returns:
            str: ID del perfil creado

        Raises:
            Exception: Si hay error creando el perfil
        """
        try:
            # Validaciones básicas
            if not name or not name.strip():
                raise Exception("El nombre del perfil no puede estar vacio")

            if not search_title or not search_title.strip():
                raise Exception("El criterio de busqueda no puede estar vacio")

            # Limpiar datos
            clean_name = self._clean_string(name.strip())
            clean_search_title = self._clean_string(search_title.strip())

            # Validaciones mejoradas para criterios flexibles
            if len(clean_name) < 2:
                raise Exception("El nombre del perfil debe tener al menos 2 caracteres")

            if not self._validate_search_criteria(clean_search_title):
                raise Exception(
                    "El criterio de busqueda no es valido.\n\n"
                    "EJEMPLOS VALIDOS:\n"
                    "• Palabra simple: 'Factura'\n"
                    "• Frase completa: 'Reporte Automatico de PDFs'\n"
                    "• Multiples palabras: 'Pedido Confirmacion'\n"
                    "• Con caracteres especiales: 'Re: Importante'\n\n"
                    "NOTA: Ahora soporta busqueda aproximada - encontrara correos que "
                    "contengan TODAS las palabras del criterio, aunque tengan texto adicional."
                )

            # Cargar perfiles existentes
            profiles = self.load_profiles()

            # Verificar que no exista un perfil con el mismo nombre (case-insensitive)
            for profile in profiles:
                existing_name = profile.get("name", "").lower()
                if existing_name == clean_name.lower():
                    raise Exception(f"Ya existe un perfil con el nombre: {clean_name}")

            # Crear nuevo perfil
            profile_id = str(uuid.uuid4())
            new_profile = {
                "id": profile_id,
                "name": clean_name,
                "search_title": clean_search_title,
                "search_type": "flexible",  # Nuevo campo para indicar tipo de búsqueda
                "created_at": datetime.now().isoformat(),
                "last_executed": None,
                "is_active": True,
                "version": "1.1"  # Versión actualizada para búsqueda flexible
            }

            # Agregar a la lista
            profiles.append(new_profile)

            # Guardar perfiles
            success = self._save_profiles(profiles)
            if not success:
                raise Exception("Error guardando el perfil en el archivo")

            # Inicializar estadísticas del perfil
            self._init_profile_stats(profile_id)

            return profile_id

        except Exception as e:
            error_msg = self._clean_string(str(e))
            raise Exception(f"Error creando perfil: {error_msg}")

    def update_profile(self, profile_id, name=None, search_title=None, is_active=None):
        """
        Actualiza un perfil existente con soporte para criterios flexibles

        Args:
            profile_id (str): ID del perfil
            name (str, optional): Nuevo nombre
            search_title (str, optional): Nuevo título de búsqueda
            is_active (bool, optional): Estado activo/inactivo

        Raises:
            Exception: Si hay error actualizando el perfil
        """
        try:
            if not profile_id or not profile_id.strip():
                raise Exception("ID del perfil requerido")

            profiles = self.load_profiles()
            profile_found = False
            profile_index = -1

            # Buscar el perfil
            for i, profile in enumerate(profiles):
                if profile.get("id") == profile_id:
                    profile_found = True
                    profile_index = i
                    break

            if not profile_found:
                raise Exception("Perfil no encontrado")

            # Actualizar campos según sea necesario
            if name is not None:
                clean_name = self._clean_string(name.strip()) if name.strip() else ""

                if not clean_name:
                    raise Exception("El nombre del perfil no puede estar vacio")

                if len(clean_name) < 2:
                    raise Exception("El nombre del perfil debe tener al menos 2 caracteres")

                # Verificar nombres duplicados (excluyendo el perfil actual)
                for other_profile in profiles:
                    if (other_profile.get("id") != profile_id and
                            other_profile.get("name", "").lower() == clean_name.lower()):
                        raise Exception(f"Ya existe un perfil con el nombre: {clean_name}")

                profiles[profile_index]["name"] = clean_name

            if search_title is not None:
                clean_search_title = self._clean_string(search_title.strip()) if search_title.strip() else ""

                if not clean_search_title:
                    raise Exception("El criterio de busqueda no puede estar vacio")

                if not self._validate_search_criteria(clean_search_title):
                    raise Exception(
                        "El criterio de busqueda no es valido.\n\n"
                        "EJEMPLOS VALIDOS:\n"
                        "• Palabra simple: 'Factura'\n"
                        "• Frase completa: 'Reporte Automatico de PDFs'\n"
                        "• Multiples palabras: 'Pedido Confirmacion'\n"
                        "• Con caracteres especiales: 'Re: Importante'\n\n"
                        "NOTA: Soporta busqueda aproximada - encontrara correos que "
                        "contengan TODAS las palabras del criterio."
                    )

                profiles[profile_index]["search_title"] = clean_search_title
                # Actualizar a búsqueda flexible si es perfil antiguo
                profiles[profile_index]["search_type"] = "flexible"
                profiles[profile_index]["version"] = "1.1"

            if is_active is not None:
                profiles[profile_index]["is_active"] = bool(is_active)

            # Actualizar timestamp de modificación
            profiles[profile_index]["updated_at"] = datetime.now().isoformat()

            # Guardar cambios
            success = self._save_profiles(profiles)
            if not success:
                raise Exception("Error guardando cambios en el archivo")

        except Exception as e:
            error_msg = self._clean_string(str(e))
            raise Exception(f"Error actualizando perfil: {error_msg}")

    def _validate_search_criteria(self, search_title):
        """
        Valida criterios de búsqueda para búsqueda flexible

        Args:
            search_title (str): Criterio a validar

        Returns:
            bool: True si es válido
        """
        try:
            if not search_title or len(search_title.strip()) < 2:
                return False

            # Limpiar y analizar el criterio
            clean_title = search_title.strip()

            # Criterios básicos de validación
            if len(clean_title) > 200:  # Límite razonable
                return False

            # Verificar que tenga al menos una palabra útil
            words = [word.strip() for word in clean_title.split() if len(word.strip()) >= 1]

            if not words:
                return False

            # Verificar que no sea solo caracteres especiales
            useful_chars = ''.join(words).replace(' ', '')
            if len(useful_chars.strip('.,;:!?-_()[]{}"\' ')) < 1:
                return False

            # Validaciones adicionales para caracteres problemáticos
            # Permitir la mayoría de caracteres pero excluir algunos problemáticos para IMAP
            forbidden_chars = ['\\', '\n', '\r', '\t']
            for char in forbidden_chars:
                if char in clean_title:
                    return False

            return True

        except Exception:
            return False

    def get_search_criteria_examples(self):
        """
        Obtiene ejemplos de criterios de búsqueda válidos

        Returns:
            dict: Ejemplos organizados por categoría
        """
        return {
            "palabras_simples": [
                "Factura",
                "Pedido",
                "Confirmacion",
                "Reporte"
            ],
            "frases_completas": [
                "Reporte Automatico de PDFs",
                "Factura Mensual",
                "Pedido Confirmacion",
                "Estado de Cuenta"
            ],
            "con_caracteres_especiales": [
                "Re: Importante",
                "[URGENTE] Notificacion",
                "Fwd: Documentos",
                "Auto: Confirmado"
            ],
            "multiples_palabras": [
                "Resumen Ejecutivo Mensual",
                "Backup Completado Exitosamente",
                "Informe Ventas Trimestre",
                "Proceso Automatico Finalizado"
            ]
        }

    def delete_profile(self, profile_id):
        """
        Elimina un perfil

        Args:
            profile_id (str): ID del perfil a eliminar

        Returns:
            bool: True si se eliminó correctamente

        Raises:
            Exception: Si hay error eliminando el perfil
        """
        try:
            if not profile_id or not profile_id.strip():
                raise Exception("ID del perfil requerido")

            profiles = self.load_profiles()
            original_count = len(profiles)

            # Filtrar perfiles (mantener todos excepto el que se elimina)
            profiles = [p for p in profiles if p.get("id") != profile_id]

            if len(profiles) == original_count:
                raise Exception("Perfil no encontrado")

            # Guardar lista actualizada
            success = self._save_profiles(profiles)
            if not success:
                raise Exception("Error guardando cambios en el archivo")

            # Eliminar estadísticas del perfil
            self._delete_profile_stats(profile_id)

            return True

        except Exception as e:
            error_msg = self._clean_string(str(e))
            raise Exception(f"Error eliminando perfil: {error_msg}")

    def get_profile(self, profile_id):
        """
        Obtiene un perfil específico

        Args:
            profile_id (str): ID del perfil

        Returns:
            dict: Datos del perfil o None si no existe
        """
        try:
            if not profile_id:
                return None

            profiles = self.load_profiles()
            for profile in profiles:
                if profile.get("id") == profile_id:
                    # Migrar perfiles antiguos a nueva versión si es necesario
                    if "search_type" not in profile:
                        profile["search_type"] = "flexible"
                        profile["version"] = "1.1"
                    return profile
            return None
        except Exception:
            return None

    def load_profiles(self):
        """
        Carga todos los perfiles desde el archivo

        Returns:
            list: Lista de perfiles
        """
        try:
            if not self.profiles_file.exists():
                return []

            with open(self.profiles_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    return []

                profiles = json.loads(content)

            # Validar que sea una lista
            if not isinstance(profiles, list):
                print(f"Archivo de perfiles corrupto: esperaba lista, encontró {type(profiles)}")
                return []

            # Limpiar strings de caracteres problemáticos y migrar perfiles antiguos
            cleaned_profiles = []
            for profile in profiles:
                if isinstance(profile, dict) and "id" in profile:
                    cleaned_profile = {}
                    for key, value in profile.items():
                        if isinstance(value, str):
                            cleaned_profile[key] = self._clean_string(value)
                        else:
                            cleaned_profile[key] = value

                    # Migrar perfiles antiguos a nueva versión
                    if "search_type" not in cleaned_profile:
                        cleaned_profile["search_type"] = "flexible"
                    if "version" not in cleaned_profile:
                        cleaned_profile["version"] = "1.1"

                    cleaned_profiles.append(cleaned_profile)

            return cleaned_profiles

        except json.JSONDecodeError as e:
            print(f"Error JSON en archivo de perfiles: {e}")
            return []
        except Exception as e:
            print(f"Error cargando perfiles: {e}")
            return []

    def get_active_profiles(self):
        """
        Obtiene solo los perfiles activos

        Returns:
            list: Lista de perfiles activos
        """
        try:
            profiles = self.load_profiles()
            active_profiles = [p for p in profiles if p.get("is_active", True)]
            return active_profiles
        except Exception:
            return []

    def update_profile_execution(self, profile_id, emails_found=0):
        """
        Actualiza las estadísticas de ejecución de un perfil

        Args:
            profile_id (str): ID del perfil
            emails_found (int): Número de emails encontrados en esta ejecución (NO acumulativo)
        """
        try:
            # Actualizar fecha de última ejecución en el perfil
            profiles = self.load_profiles()
            for profile in profiles:
                if profile.get("id") == profile_id:
                    profile["last_executed"] = datetime.now().isoformat()
                    break

            self._save_profiles(profiles)

            # Actualizar estadísticas
            stats = self._load_profile_stats()

            if profile_id not in stats:
                stats[profile_id] = {
                    "total_executions": 0,
                    "current_emails_found": 0,
                    "total_emails_accumulated": 0,
                    "last_execution": None,
                    "execution_history": []
                }

            # Actualizar estadísticas
            stats[profile_id]["total_executions"] += 1
            # CORRECCIÓN: Actualizar el número actual de correos (no sumar)
            stats[profile_id]["current_emails_found"] = emails_found
            # Mantener acumulativo separado para historial
            stats[profile_id]["total_emails_accumulated"] += emails_found
            stats[profile_id]["last_execution"] = datetime.now().isoformat()

            # Agregar al historial (mantener últimas 50 ejecuciones)
            execution_record = {
                "timestamp": datetime.now().isoformat(),
                "emails_found": emails_found
            }

            if "execution_history" not in stats[profile_id]:
                stats[profile_id]["execution_history"] = []

            stats[profile_id]["execution_history"].append(execution_record)
            if len(stats[profile_id]["execution_history"]) > 50:
                stats[profile_id]["execution_history"] = stats[profile_id]["execution_history"][-50:]

            self._save_profile_stats(stats)

        except Exception as e:
            print(f"Error actualizando estadísticas de ejecución: {e}")

    def get_profile_stats(self, profile_id):
        """
        Obtiene las estadísticas de un perfil

        Args:
            profile_id (str): ID del perfil

        Returns:
            dict: Estadísticas del perfil
        """
        try:
            stats = self._load_profile_stats()
            profile_stats = stats.get(profile_id, {
                "total_executions": 0,
                "current_emails_found": 0,
                "total_emails_accumulated": 0,
                "last_execution": None,
                "execution_history": []
            })

            # Mantener compatibilidad con versiones anteriores
            if "current_emails_found" not in profile_stats:
                profile_stats["current_emails_found"] = profile_stats.get("total_emails_found", 0)
            if "total_emails_accumulated" not in profile_stats:
                profile_stats["total_emails_accumulated"] = profile_stats.get("total_emails_found", 0)

            return profile_stats
        except Exception:
            return {
                "total_executions": 0,
                "current_emails_found": 0,
                "total_emails_accumulated": 0,
                "last_execution": None,
                "execution_history": []
            }

    def get_all_profiles_stats(self):
        """
        Obtiene estadísticas de todos los perfiles

        Returns:
            dict: Estadísticas de todos los perfiles
        """
        try:
            profiles = self.load_profiles()
            stats = self._load_profile_stats()

            result = {}
            for profile in profiles:
                profile_id = profile.get("id")
                if profile_id:
                    profile_stats = stats.get(profile_id, {
                        "total_executions": 0,
                        "current_emails_found": 0,
                        "total_emails_accumulated": 0,
                        "last_execution": None,
                        "execution_history": []
                    })

                    # Mantener compatibilidad con versiones anteriores
                    if "current_emails_found" not in profile_stats:
                        profile_stats["current_emails_found"] = profile_stats.get("total_emails_found", 0)
                    if "total_emails_accumulated" not in profile_stats:
                        profile_stats["total_emails_accumulated"] = profile_stats.get("total_emails_found", 0)

                    result[profile_id] = {
                        "profile": profile,
                        "stats": profile_stats
                    }

            return result
        except Exception:
            return {}

    def get_profiles_summary(self):
        """
        Obtiene un resumen de los perfiles

        Returns:
            dict: Resumen con contadores y estadísticas
        """
        try:
            profiles = self.load_profiles()
            stats = self._load_profile_stats()

            active_count = sum(1 for p in profiles if p.get("is_active", True))
            inactive_count = len(profiles) - active_count

            total_emails_current = 0
            total_emails_accumulated = 0
            total_executions = 0

            # Contar perfiles con búsqueda flexible
            flexible_profiles = sum(1 for p in profiles if p.get("search_type") == "flexible")

            for stat in stats.values():
                total_emails_current += stat.get("current_emails_found", stat.get("total_emails_found", 0))
                total_emails_accumulated += stat.get("total_emails_accumulated", stat.get("total_emails_found", 0))
                total_executions += stat.get("total_executions", 0)

            return {
                "total_profiles": len(profiles),
                "active_profiles": active_count,
                "inactive_profiles": inactive_count,
                "flexible_profiles": flexible_profiles,
                "total_executions": total_executions,
                "current_emails_found": total_emails_current,
                "total_emails_accumulated": total_emails_accumulated
            }
        except Exception:
            return {
                "total_profiles": 0,
                "active_profiles": 0,
                "inactive_profiles": 0,
                "flexible_profiles": 0,
                "total_executions": 0,
                "current_emails_found": 0,
                "total_emails_accumulated": 0
            }

    def export_profiles(self, export_path):
        """
        Exporta perfiles a un archivo JSON

        Args:
            export_path (str): Ruta donde exportar

        Raises:
            Exception: Si hay error exportando
        """
        try:
            profiles = self.load_profiles()
            stats = self._load_profile_stats()

            export_data = {
                "profiles": profiles,
                "stats": stats,
                "export_date": datetime.now().isoformat(),
                "version": "1.1",  # Versión con búsqueda flexible
                "features": ["flexible_search", "approximate_matching"]
            }

            with open(export_path, "w", encoding="utf-8") as f:
                json.dump(export_data, f, indent=4, ensure_ascii=False)

        except Exception as e:
            error_msg = self._clean_string(str(e))
            raise Exception(f"Error exportando perfiles: {error_msg}")

    def import_profiles(self, import_path, merge=True):
        """
        Importa perfiles desde un archivo JSON

        Args:
            import_path (str): Ruta del archivo a importar
            merge (bool): Si True, fusiona con perfiles existentes; si False, reemplaza

        Raises:
            Exception: Si hay error importando
        """
        try:
            import_file = Path(import_path)
            if not import_file.exists():
                raise Exception(f"El archivo no existe: {import_path}")

            with open(import_file, "r", encoding="utf-8") as f:
                import_data = json.load(f)

            if "profiles" not in import_data:
                raise Exception("El archivo no contiene perfiles validos")

            imported_profiles = import_data["profiles"]
            imported_stats = import_data.get("stats", {})

            # Migrar perfiles importados a nueva versión si es necesario
            for profile in imported_profiles:
                if "search_type" not in profile:
                    profile["search_type"] = "flexible"
                if "version" not in profile:
                    profile["version"] = "1.1"

            if merge:
                existing_profiles = self.load_profiles()
                existing_stats = self._load_profile_stats()

                # Fusionar perfiles (evitar duplicados por nombre)
                existing_names = {p.get("name", "").lower() for p in existing_profiles}

                for profile in imported_profiles:
                    profile_name = profile.get("name", "").lower()
                    if profile_name and profile_name not in existing_names:
                        existing_profiles.append(profile)

                # Fusionar estadísticas
                existing_stats.update(imported_stats)

                self._save_profiles(existing_profiles)
                self._save_profile_stats(existing_stats)
            else:
                self._save_profiles(imported_profiles)
                self._save_profile_stats(imported_stats)

        except Exception as e:
            error_msg = self._clean_string(str(e))
            raise Exception(f"Error importando perfiles: {error_msg}")

    def _save_profiles(self, profiles):
        """
        Guarda perfiles en el archivo JSON

        Args:
            profiles (list): Lista de perfiles a guardar

        Returns:
            bool: True si se guardó correctamente
        """
        try:
            # Validar que es una lista
            if not isinstance(profiles, list):
                raise Exception("Los perfiles deben ser una lista")

            # Crear backup si existe el archivo
            if self.profiles_file.exists():
                backup_path = self.profiles_file.with_suffix('.json.backup')
                try:
                    import shutil
                    shutil.copy2(self.profiles_file, backup_path)
                except Exception:
                    pass  # Si no se puede hacer backup, continuar

            # Guardar con formato legible
            with open(self.profiles_file, "w", encoding="utf-8") as f:
                json.dump(profiles, f, indent=4, ensure_ascii=False, sort_keys=True)

            return True

        except Exception as e:
            print(f"Error guardando perfiles: {e}")
            return False

    def _init_profile_stats(self, profile_id):
        """
        Inicializa estadísticas para un nuevo perfil

        Args:
            profile_id (str): ID del perfil
        """
        try:
            stats = self._load_profile_stats()
            if profile_id not in stats:
                stats[profile_id] = {
                    "total_executions": 0,
                    "current_emails_found": 0,
                    "total_emails_accumulated": 0,
                    "last_execution": None,
                    "execution_history": []
                }
                self._save_profile_stats(stats)
        except Exception:
            pass

    def _delete_profile_stats(self, profile_id):
        """
        Elimina estadísticas de un perfil

        Args:
            profile_id (str): ID del perfil
        """
        try:
            stats = self._load_profile_stats()
            if profile_id in stats:
                del stats[profile_id]
                self._save_profile_stats(stats)
        except Exception:
            pass

    def _load_profile_stats(self):
        """
        Carga estadísticas desde el archivo

        Returns:
            dict: Estadísticas de perfiles
        """
        try:
            if not self.stats_file.exists():
                return {}

            with open(self.stats_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    return {}

                return json.loads(content)

        except json.JSONDecodeError:
            return {}
        except Exception:
            return {}

    def _save_profile_stats(self, stats):
        """
        Guarda estadísticas en el archivo

        Args:
            stats (dict): Estadísticas a guardar

        Returns:
            bool: True si se guardó correctamente
        """
        try:
            with open(self.stats_file, "w", encoding="utf-8") as f:
                json.dump(stats, f, indent=4, ensure_ascii=False)
            return True
        except Exception:
            return False

    def _clean_string(self, text):
        """
        Limpia un string de caracteres problemáticos para ASCII

        Args:
            text (str): Texto a limpiar

        Returns:
            str: Texto limpio
        """
        if not text:
            return ""

        try:
            # Convertir a string si no lo es
            text = str(text)

            # Reemplazar caracteres problemáticos comunes
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
                text = text.replace(char, replacement)

            # Codificar y decodificar para limpiar caracteres problemáticos restantes
            return text.encode('ascii', 'ignore').decode('ascii')

        except Exception:
            return str(text) if text else ""