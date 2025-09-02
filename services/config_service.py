# services/config_service.py
"""
Servicio para manejo de persistencia de configuraciones en formato JSON.
Guarda y carga credenciales de manera segura con cifrado básico y configuración de envío de reportes.
"""

# Archivos relacionados: Ninguno (servicio independiente)

import json
import base64
from pathlib import Path


class ConfigService:
    def __init__(self, config_dir="config"):
        """
        Inicializa el servicio de configuración

        Args:
            config_dir (str): Directorio donde guardar las configuraciones
        """
        self.config_dir = Path(config_dir)
        self.credentials_file = self.config_dir / "credentials.json"
        self.email_send_file = self.config_dir / "email_send_config.json"
        self.ensure_config_directory()

    def ensure_config_directory(self):
        """Asegura que el directorio de configuración existe"""
        self.config_dir.mkdir(exist_ok=True)

    def save_credentials(self, credentials):
        """
        Guarda las credenciales en formato JSON con cifrado básico

        Args:
            credentials (dict): Diccionario con las credenciales

        Raises:
            Exception: Si hay error guardando el archivo
        """
        try:
            # Cifrar la contraseña antes de guardar
            encrypted_credentials = credentials.copy()
            if "password" in encrypted_credentials:
                encrypted_credentials["password"] = self._encrypt_password(
                    encrypted_credentials["password"]
                )

            # Limpiar strings de caracteres problemáticos
            for key, value in encrypted_credentials.items():
                if isinstance(value, str):
                    encrypted_credentials[key] = self._clean_string(value)

            # Guardar en archivo JSON con codificación segura
            with open(self.credentials_file, "w", encoding="utf-8") as f:
                json.dump(encrypted_credentials, f, indent=4, ensure_ascii=True)

        except Exception as e:
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            raise Exception(f"Error guardando credenciales: {error_msg}")

    def load_credentials(self):
        """
        Carga las credenciales desde el archivo JSON

        Returns:
            dict: Credenciales o None si no existen

        Raises:
            Exception: Si hay error cargando el archivo
        """
        try:
            if not self.credentials_file.exists():
                return None

            with open(self.credentials_file, "r", encoding="utf-8") as f:
                credentials = json.load(f)

            # Descifrar la contraseña
            if "password" in credentials:
                credentials["password"] = self._decrypt_password(
                    credentials["password"]
                )

            # Limpiar strings de caracteres problemáticos
            for key, value in credentials.items():
                if isinstance(value, str):
                    credentials[key] = self._clean_string(value)

            return credentials

        except Exception as e:
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            raise Exception(f"Error cargando credenciales: {error_msg}")

    def credentials_exist(self):
        """
        Verifica si existen credenciales guardadas

        Returns:
            bool: True si existen credenciales
        """
        return self.credentials_file.exists()

    def delete_credentials(self):
        """
        Elimina las credenciales guardadas

        Returns:
            bool: True si se eliminaron exitosamente
        """
        try:
            if self.credentials_file.exists():
                self.credentials_file.unlink()
                return True
            return False
        except Exception:
            return False

    def save_email_send_config(self, config):
        """
        Guarda la configuración de envío de reportes por email

        Args:
            config (dict): Configuración de envío

        Raises:
            Exception: Si hay error guardando la configuración
        """
        try:
            # Limpiar strings de caracteres problemáticos
            clean_config = {}
            for key, value in config.items():
                if isinstance(value, str):
                    clean_config[key] = self._clean_string(value)
                else:
                    clean_config[key] = value

            # Guardar en archivo JSON
            with open(self.email_send_file, "w", encoding="utf-8") as f:
                json.dump(clean_config, f, indent=4, ensure_ascii=True)

        except Exception as e:
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            raise Exception(f"Error guardando configuracion de envio: {error_msg}")

    def load_email_send_config(self):
        """
        Carga la configuración de envío de reportes

        Returns:
            dict: Configuración de envío o None si no existe

        Raises:
            Exception: Si hay error cargando la configuración
        """
        try:
            if not self.email_send_file.exists():
                return None

            with open(self.email_send_file, "r", encoding="utf-8") as f:
                config = json.load(f)

            # Limpiar strings de caracteres problemáticos
            clean_config = {}
            for key, value in config.items():
                if isinstance(value, str):
                    clean_config[key] = self._clean_string(value)
                else:
                    clean_config[key] = value

            return clean_config

        except Exception as e:
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            raise Exception(f"Error cargando configuracion de envio: {error_msg}")

    def email_send_config_exists(self):
        """
        Verifica si existe configuración de envío

        Returns:
            bool: True si existe la configuración
        """
        return self.email_send_file.exists()

    def delete_email_send_config(self):
        """
        Elimina la configuración de envío

        Returns:
            bool: True si se eliminó exitosamente
        """
        try:
            if self.email_send_file.exists():
                self.email_send_file.unlink()
                return True
            return False
        except Exception:
            return False

    def get_email_send_status(self):
        """
        Obtiene el estado de la configuración de envío

        Returns:
            dict: Estado de la configuración de envío
        """
        try:
            config = self.load_email_send_config()

            if not config:
                return {
                    "configured": False,
                    "enabled": False,
                    "has_recipient": False,
                    "ready": False
                }

            has_recipient = bool(config.get("recipient", "").strip())
            is_enabled = config.get("enabled", False)

            return {
                "configured": True,
                "enabled": is_enabled,
                "has_recipient": has_recipient,
                "ready": is_enabled and has_recipient,
                "recipient": config.get("recipient", ""),
                "subject": config.get("subject", ""),
                "cc": config.get("cc", "")
            }

        except Exception:
            return {
                "configured": False,
                "enabled": False,
                "has_recipient": False,
                "ready": False
            }

    def backup_credentials(self, backup_name=None):
        """
        Crea un respaldo de las credenciales

        Args:
            backup_name (str): Nombre del archivo de respaldo

        Returns:
            str: Ruta del archivo de respaldo creado

        Raises:
            Exception: Si hay error creando el respaldo
        """
        try:
            if not self.credentials_file.exists():
                raise Exception("No hay credenciales para respaldar")

            if not backup_name:
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_name = f"credentials_backup_{timestamp}.json"

            backup_path = self.config_dir / backup_name

            # Copiar archivo
            with open(self.credentials_file, "r", encoding="utf-8") as src:
                with open(backup_path, "w", encoding="utf-8") as dst:
                    dst.write(src.read())

            return str(backup_path)

        except Exception as e:
            raise Exception(f"Error creando respaldo: {str(e)}")

    def restore_credentials(self, backup_path):
        """
        Restaura credenciales desde un respaldo

        Args:
            backup_path (str): Ruta del archivo de respaldo

        Raises:
            Exception: Si hay error restaurando el respaldo
        """
        try:
            backup_file = Path(backup_path)
            if not backup_file.exists():
                raise Exception(f"El archivo de respaldo no existe: {backup_path}")

            # Validar que el respaldo es válido
            with open(backup_file, "r", encoding="utf-8") as f:
                json.load(f)  # Esto lanzará excepción si no es JSON válido

            # Copiar archivo
            with open(backup_file, "r", encoding="utf-8") as src:
                with open(self.credentials_file, "w", encoding="utf-8") as dst:
                    dst.write(src.read())

        except Exception as e:
            raise Exception(f"Error restaurando respaldo: {str(e)}")

    def get_config_info(self):
        """
        Obtiene información sobre la configuración actual

        Returns:
            dict: Información de configuración
        """
        info = {
            "config_directory": str(self.config_dir),
            "credentials_file": str(self.credentials_file),
            "email_send_file": str(self.email_send_file),
            "credentials_exist": self.credentials_exist(),
            "email_send_config_exists": self.email_send_config_exists(),
            "credentials_file_size": 0,
            "email_send_file_size": 0,
            "credentials_last_modified": None,
            "email_send_last_modified": None
        }

        try:
            if self.credentials_file.exists():
                stat = self.credentials_file.stat()
                info["credentials_file_size"] = stat.st_size
                info["credentials_last_modified"] = stat.st_mtime

            if self.email_send_file.exists():
                stat = self.email_send_file.stat()
                info["email_send_file_size"] = stat.st_size
                info["email_send_last_modified"] = stat.st_mtime
        except Exception:
            pass

        return info

    def _encrypt_password(self, password):
        """
        Cifra una contraseña usando base64 (cifrado básico)
        NOTA: Este no es un cifrado seguro, solo ofuscación

        Args:
            password (str): Contraseña a cifrar

        Returns:
            str: Contraseña cifrada
        """
        try:
            # Codificar a bytes y luego a base64
            encoded_bytes = password.encode('utf-8')
            encrypted_bytes = base64.b64encode(encoded_bytes)
            return encrypted_bytes.decode('utf-8')
        except Exception:
            return password  # Si hay error, devolver sin cifrar

    def _decrypt_password(self, encrypted_password):
        """
        Descifra una contraseña desde base64

        Args:
            encrypted_password (str): Contraseña cifrada

        Returns:
            str: Contraseña descifrada
        """
        try:
            # Decodificar desde base64 a bytes y luego a string
            encrypted_bytes = encrypted_password.encode('utf-8')
            decoded_bytes = base64.b64decode(encrypted_bytes)
            return decoded_bytes.decode('utf-8')
        except Exception:
            return encrypted_password  # Si hay error, devolver sin descifrar

    def export_config(self, export_path, include_passwords=False):
        """
        Exporta la configuración a un archivo

        Args:
            export_path (str): Ruta donde exportar
            include_passwords (bool): Si incluir contraseñas en el export

        Raises:
            Exception: Si hay error exportando
        """
        try:
            credentials = self.load_credentials()
            email_send_config = self.load_email_send_config()

            export_data = {
                "credentials": credentials,
                "email_send_config": email_send_config,
                "export_date": None,
                "version": "1.1"
            }

            # Obtener fecha actual
            from datetime import datetime
            export_data["export_date"] = datetime.now().isoformat()

            if credentials and not include_passwords and "password" in credentials:
                export_data["credentials"]["password"] = "***OCULTA***"

            with open(export_path, "w", encoding="utf-8") as f:
                json.dump(export_data, f, indent=4, ensure_ascii=False)

        except Exception as e:
            raise Exception(f"Error exportando configuración: {str(e)}")

    def import_config(self, import_path):
        """
        Importa configuración desde un archivo

        Args:
            import_path (str): Ruta del archivo a importar

        Raises:
            Exception: Si hay error importando
        """
        try:
            import_file = Path(import_path)
            if not import_file.exists():
                raise Exception(f"El archivo no existe: {import_path}")

            with open(import_file, "r", encoding="utf-8") as f:
                import_data = json.load(f)

            # Importar credenciales si existen
            if "credentials" in import_data and import_data["credentials"]:
                credentials = import_data["credentials"]
                # Validar que tiene los campos mínimos requeridos
                required_fields = ["email", "server", "port"]
                for field in required_fields:
                    if field not in credentials:
                        raise Exception(f"Campo requerido faltante en credenciales: {field}")

                self.save_credentials(credentials)

            # Importar configuración de envío si existe
            if "email_send_config" in import_data and import_data["email_send_config"]:
                email_config = import_data["email_send_config"]
                self.save_email_send_config(email_config)

        except Exception as e:
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            raise Exception(f"Error importando configuracion: {error_msg}")

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
            # Reemplazar caracteres problemáticos comunes
            text = text.replace('\xa0', ' ')  # Espacio no-rompible
            text = text.replace('\u2019', "'")  # Apostrofe curvo
            text = text.replace('\u2018', "'")  # Apostrofe curvo
            text = text.replace('\u201c', '"')  # Comilla curva
            text = text.replace('\u201d', '"')  # Comilla curva

            # Codificar y decodificar para limpiar caracteres problemáticos
            return text.encode('ascii', 'ignore').decode('ascii')
        except Exception:
            return str(text)