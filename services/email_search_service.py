# services/email_search_service.py
"""
Servicio para búsqueda de correos usando perfiles de búsqueda con soporte para búsqueda aproximada.
Conecta con cuentas de email y busca correos basado en criterios definidos, soportando espacios y títulos parciales.
"""

import imaplib
import email
import socket
from datetime import datetime, timedelta
from email.header import decode_header
import re


class EmailSearchService:
    def __init__(self):
        """Inicializa el servicio de búsqueda de emails"""
        self.connection = None
        self.current_credentials = None

    def search_by_profile(self, profile, credentials, days_back=30):
        """
        Busca correos usando un perfil de búsqueda específico

        Args:
            profile (dict): Perfil de búsqueda con criterios
            credentials (dict): Credenciales SMTP/IMAP
            days_back (int): Días hacia atrás para buscar

        Returns:
            tuple: (success: bool, emails_found: int, message: str)
        """
        try:
            search_title = profile.get("search_title", "").strip()
            if not search_title:
                return False, 0, "Perfil sin criterio de busqueda"

            # Conectar al servidor IMAP
            success, message = self._connect_imap(credentials)
            if not success:
                return False, 0, f"Error conectando: {message}"

            # Buscar correos usando búsqueda aproximada
            emails_found = self._search_emails_by_subject_flexible(search_title, days_back)

            # Desconectar
            self._disconnect_imap()

            return True, emails_found, f"Busqueda completada. {emails_found} correos encontrados"

        except Exception as e:
            self._disconnect_imap()
            error_msg = self._clean_string(str(e))
            return False, 0, f"Error en busqueda: {error_msg}"

    def search_multiple_profiles(self, profiles, credentials, days_back=30):
        """
        Busca correos usando múltiples perfiles

        Args:
            profiles (list): Lista de perfiles de búsqueda
            credentials (dict): Credenciales SMTP/IMAP
            days_back (int): Días hacia atrás para buscar

        Returns:
            dict: Resultados por perfil
        """
        results = {}

        try:
            # Conectar una vez para todos los perfiles
            success, message = self._connect_imap(credentials)
            if not success:
                # Si no puede conectar, devolver error para todos los perfiles
                for profile in profiles:
                    results[profile["id"]] = {
                        "success": False,
                        "emails_found": 0,
                        "message": f"Error conectando: {message}",
                        "profile_name": profile["name"]
                    }
                return results

            # Buscar para cada perfil
            for profile in profiles:
                try:
                    if not profile.get("is_active", True):
                        results[profile["id"]] = {
                            "success": False,
                            "emails_found": 0,
                            "message": "Perfil inactivo",
                            "profile_name": profile["name"]
                        }
                        continue

                    search_title = profile.get("search_title", "").strip()
                    if not search_title:
                        results[profile["id"]] = {
                            "success": False,
                            "emails_found": 0,
                            "message": "Perfil sin criterio de busqueda",
                            "profile_name": profile["name"]
                        }
                        continue

                    # Usar búsqueda flexible
                    emails_found = self._search_emails_by_subject_flexible(search_title, days_back)

                    results[profile["id"]] = {
                        "success": True,
                        "emails_found": emails_found,
                        "message": f"Busqueda completada. {emails_found} correos encontrados",
                        "profile_name": profile["name"]
                    }

                except Exception as e:
                    error_msg = self._clean_string(str(e))
                    results[profile["id"]] = {
                        "success": False,
                        "emails_found": 0,
                        "message": f"Error en busqueda: {error_msg}",
                        "profile_name": profile["name"]
                    }

            # Desconectar
            self._disconnect_imap()

        except Exception as e:
            self._disconnect_imap()
            error_msg = self._clean_string(str(e))

            # Si hay error general, aplicar a todos los perfiles no procesados
            for profile in profiles:
                if profile["id"] not in results:
                    results[profile["id"]] = {
                        "success": False,
                        "emails_found": 0,
                        "message": f"Error general: {error_msg}",
                        "profile_name": profile["name"]
                    }

        return results

    def get_email_details(self, profile, credentials, days_back=30, limit=10):
        """
        Obtiene detalles de correos encontrados por un perfil

        Args:
            profile (dict): Perfil de búsqueda
            credentials (dict): Credenciales SMTP/IMAP
            days_back (int): Días hacia atrás para buscar
            limit (int): Límite de correos a obtener detalles

        Returns:
            tuple: (success: bool, emails: list, message: str)
        """
        try:
            search_title = profile.get("search_title", "").strip()
            if not search_title:
                return False, [], "Perfil sin criterio de busqueda"

            # Conectar al servidor IMAP
            success, message = self._connect_imap(credentials)
            if not success:
                return False, [], f"Error conectando: {message}"

            # Buscar y obtener detalles de correos usando búsqueda flexible
            emails = self._get_email_details_by_subject_flexible(search_title, days_back, limit)

            # Desconectar
            self._disconnect_imap()

            return True, emails, f"Se obtuvieron {len(emails)} correos"

        except Exception as e:
            self._disconnect_imap()
            error_msg = self._clean_string(str(e))
            return False, [], f"Error obteniendo detalles: {error_msg}"

    def test_imap_connection(self, credentials):
        """
        Prueba la conexión IMAP

        Args:
            credentials (dict): Credenciales de email

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            success, message = self._connect_imap(credentials)
            if success:
                self._disconnect_imap()
                return True, "Conexion IMAP exitosa"
            else:
                return False, message
        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error probando conexion IMAP: {error_msg}"

    def _connect_imap(self, credentials):
        """
        Conecta al servidor IMAP

        Args:
            credentials (dict): Credenciales de email

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            email_addr = credentials.get("email", "").strip()
            password = credentials.get("password", "").strip()
            server = credentials.get("server", "").strip()

            if not all([email_addr, password, server]):
                return False, "Credenciales incompletas"

            # Determinar servidor IMAP basado en el servidor SMTP
            imap_server = self._get_imap_server(server)
            if not imap_server:
                return False, f"No se pudo determinar servidor IMAP para: {server}"

            # Conectar
            self.connection = imaplib.IMAP4_SSL(imap_server)
            self.connection.login(email_addr, password)

            # Seleccionar buzón de entrada
            self.connection.select('INBOX')

            self.current_credentials = credentials
            return True, "Conexion IMAP establecida"

        except imaplib.IMAP4.error as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error de autenticacion IMAP: {error_msg}"
        except socket.gaierror:
            return False, "Error de DNS. Verifique el servidor"
        except socket.timeout:
            return False, "Tiempo de conexion agotado"
        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error conectando IMAP: {error_msg}"

    def _disconnect_imap(self):
        """Desconecta del servidor IMAP"""
        try:
            if self.connection:
                self.connection.close()
                self.connection.logout()
        except Exception:
            pass
        finally:
            self.connection = None
            self.current_credentials = None

    def _search_emails_by_subject_flexible(self, search_title, days_back=30):
        """
        Busca correos por título/asunto usando búsqueda aproximada y flexible
        Soporta espacios y encuentra coincidencias parciales

        Args:
            search_title (str): Título a buscar (puede contener espacios)
            days_back (int): Días hacia atrás

        Returns:
            int: Número de correos encontrados
        """
        try:
            if not self.connection:
                return 0

            # Calcular fecha de inicio
            start_date = datetime.now() - timedelta(days=days_back)
            date_str = start_date.strftime("%d-%b-%Y")

            # Construir criterio de búsqueda flexible
            search_criteria = self._build_flexible_search_criteria(search_title, date_str)

            print(f"DEBUG: Criterio de búsqueda construido: {search_criteria}")

            typ, msgnums = self.connection.search(None, search_criteria)

            if typ != 'OK':
                return 0

            # Contar correos encontrados
            if msgnums[0]:
                return len(msgnums[0].split())
            else:
                return 0

        except Exception as e:
            print(f"ERROR en búsqueda flexible: {e}")
            return 0

    def _get_email_details_by_subject_flexible(self, search_title, days_back=30, limit=10):
        """
        Obtiene detalles de correos por título/asunto usando búsqueda aproximada

        Args:
            search_title (str): Título a buscar
            days_back (int): Días hacia atrás
            limit (int): Límite de correos

        Returns:
            list: Lista de correos con detalles
        """
        emails = []

        try:
            if not self.connection:
                return emails

            # Calcular fecha de inicio
            start_date = datetime.now() - timedelta(days=days_back)
            date_str = start_date.strftime("%d-%b-%Y")

            # Usar criterio de búsqueda flexible
            search_criteria = self._build_flexible_search_criteria(search_title, date_str)
            typ, msgnums = self.connection.search(None, search_criteria)

            if typ != 'OK' or not msgnums[0]:
                return emails

            # Obtener números de mensaje
            msg_numbers = msgnums[0].split()

            # Limitar número de correos
            msg_numbers = msg_numbers[-limit:]  # Últimos N correos

            # Obtener detalles de cada correo
            for msg_num in msg_numbers:
                try:
                    typ, msg_data = self.connection.fetch(msg_num, '(RFC822)')

                    if typ != 'OK':
                        continue

                    # Procesar mensaje
                    email_message = email.message_from_bytes(msg_data[0][1])

                    # Extraer información básica
                    subject = self._decode_mime_words(email_message.get("Subject", ""))
                    sender = self._decode_mime_words(email_message.get("From", ""))
                    date_str_email = email_message.get("Date", "")

                    # Limpiar strings
                    subject = self._clean_string(subject)
                    sender = self._clean_string(sender)

                    email_info = {
                        "subject": subject,
                        "sender": sender,
                        "date": date_str_email,
                        "message_id": msg_num.decode() if isinstance(msg_num, bytes) else str(msg_num)
                    }

                    emails.append(email_info)

                except Exception:
                    continue  # Si hay error con un correo, continuar con el siguiente

        except Exception:
            pass  # Si hay error general, devolver lo que se pudo obtener

        return emails

    def _build_flexible_search_criteria(self, search_title, date_str):
        """
        Construye criterio de búsqueda IMAP flexible que soporte espacios, timestamps y texto adicional

        Args:
            search_title (str): Título/criterio de búsqueda
            date_str (str): Fecha desde la cual buscar

        Returns:
            str: Criterio de búsqueda IMAP
        """
        try:
            # Limpiar y preparar el criterio de búsqueda
            clean_title = self._clean_string(search_title.strip())

            if not clean_title:
                return f'SINCE "{date_str}"'

            # ESTRATEGIA MEJORADA: Combinación de múltiples aproximaciones
            if ' ' in clean_title:
                # Dividir en palabras significativas (filtrar palabras muy cortas y comunes)
                all_words = clean_title.split()

                # Filtrar palabras significativas (>=2 chars y no palabras de conexión comunes)
                stop_words = {'de', 'en', 'el', 'la', 'los', 'las', 'un', 'una', 'y', 'o', 'a', 'con', 'por', 'para'}
                significant_words = []

                for word in all_words:
                    clean_word = word.strip().lower()
                    if len(clean_word) >= 2 and clean_word not in stop_words:
                        significant_words.append(word.strip())

                if len(significant_words) == 0:
                    # Si no hay palabras significativas, usar todas las palabras >= 2 chars
                    significant_words = [word.strip() for word in all_words if len(word.strip()) >= 2]

                if len(significant_words) == 1:
                    # Solo una palabra significativa
                    return f'(SUBJECT "{significant_words[0]}" SINCE "{date_str}")'

                elif len(significant_words) > 1:
                    # ESTRATEGIA MÚLTIPLE para encontrar correos con timestamps y texto adicional

                    # Opción 1: Buscar frase completa (más específico)
                    phrase_criteria = f'SUBJECT "{clean_title}"'

                    # Opción 2: Buscar todas las palabras significativas (más flexible)
                    word_criteria = []
                    for word in significant_words:
                        word_criteria.append(f'SUBJECT "{word}"')
                    words_criteria = ' '.join(word_criteria)

                    # Opción 3: Buscar combinaciones de palabras clave (ultra flexible)
                    # Tomar las 2-3 palabras más importantes para casos con muchos timestamps/texto extra
                    key_words = significant_words[:3]  # Máximo 3 palabras clave
                    key_criteria = []
                    for word in key_words:
                        key_criteria.append(f'SUBJECT "{word}"')
                    key_words_criteria = ' '.join(key_criteria)

                    # CRITERIO FINAL: Combinar estrategias con OR para máxima flexibilidad
                    # Esto encontrará correos que:
                    # 1. Contengan la frase exacta, O
                    # 2. Contengan todas las palabras significativas, O
                    # 3. Contengan al menos las palabras clave principales
                    if len(significant_words) <= 3:
                        # Pocas palabras: usar estrategia 1 y 2
                        combined_criteria = f'OR {phrase_criteria} ({words_criteria})'
                    else:
                        # Muchas palabras: usar todas las estrategias
                        combined_criteria = f'OR OR {phrase_criteria} ({words_criteria}) ({key_words_criteria})'

                    return f'({combined_criteria} SINCE "{date_str}")'
                else:
                    # No hay palabras válidas
                    return f'SINCE "{date_str}"'
            else:
                # Palabra única, búsqueda directa
                return f'(SUBJECT "{clean_title}" SINCE "{date_str}")'

        except Exception as e:
            print(f"ERROR construyendo criterio de búsqueda: {e}")
            # Fallback a búsqueda básica por primera palabra significativa
            try:
                first_word = clean_title.split()[0] if clean_title.split() else ""
                if first_word and len(first_word) >= 2:
                    return f'(SUBJECT "{first_word}" SINCE "{date_str}")'
            except:
                pass
            return f'SINCE "{date_str}"'

    def _search_emails_by_subject(self, search_title, days_back=30):
        """
        MÉTODO LEGACY - Busca correos por título/asunto (búsqueda exacta)
        Mantenido por compatibilidad, pero se recomienda usar _search_emails_by_subject_flexible
        """
        try:
            if not self.connection:
                return 0

            # Calcular fecha de inicio
            start_date = datetime.now() - timedelta(days=days_back)
            date_str = start_date.strftime("%d-%b-%Y")

            # Buscar correos con el criterio de asunto y fecha (MÉTODO ORIGINAL)
            search_criteria = f'(SUBJECT "{search_title}" SINCE "{date_str}")'

            typ, msgnums = self.connection.search(None, search_criteria)

            if typ != 'OK':
                return 0

            # Contar correos encontrados
            if msgnums[0]:
                return len(msgnums[0].split())
            else:
                return 0

        except Exception:
            return 0

    def _get_email_details_by_subject(self, search_title, days_back=30, limit=10):
        """
        MÉTODO LEGACY - Obtiene detalles de correos por título/asunto (búsqueda exacta)
        Mantenido por compatibilidad, pero se recomienda usar _get_email_details_by_subject_flexible
        """
        emails = []

        try:
            if not self.connection:
                return emails

            # Calcular fecha de inicio
            start_date = datetime.now() - timedelta(days=days_back)
            date_str = start_date.strftime("%d-%b-%Y")

            # Buscar correos (MÉTODO ORIGINAL)
            search_criteria = f'(SUBJECT "{search_title}" SINCE "{date_str}")'
            typ, msgnums = self.connection.search(None, search_criteria)

            if typ != 'OK' or not msgnums[0]:
                return emails

            # Obtener números de mensaje
            msg_numbers = msgnums[0].split()

            # Limitar número de correos
            msg_numbers = msg_numbers[-limit:]  # Últimos N correos

            # Obtener detalles de cada correo
            for msg_num in msg_numbers:
                try:
                    typ, msg_data = self.connection.fetch(msg_num, '(RFC822)')

                    if typ != 'OK':
                        continue

                    # Procesar mensaje
                    email_message = email.message_from_bytes(msg_data[0][1])

                    # Extraer información básica
                    subject = self._decode_mime_words(email_message.get("Subject", ""))
                    sender = self._decode_mime_words(email_message.get("From", ""))
                    date_str_email = email_message.get("Date", "")

                    # Limpiar strings
                    subject = self._clean_string(subject)
                    sender = self._clean_string(sender)

                    email_info = {
                        "subject": subject,
                        "sender": sender,
                        "date": date_str_email,
                        "message_id": msg_num.decode() if isinstance(msg_num, bytes) else str(msg_num)
                    }

                    emails.append(email_info)

                except Exception:
                    continue  # Si hay error con un correo, continuar con el siguiente

        except Exception:
            pass  # Si hay error general, devolver lo que se pudo obtener

        return emails

    def _get_imap_server(self, smtp_server):
        """
        Determina el servidor IMAP basado en el servidor SMTP

        Args:
            smtp_server (str): Servidor SMTP

        Returns:
            str: Servidor IMAP correspondiente
        """
        imap_mapping = {
            "smtp.gmail.com": "imap.gmail.com",
            "smtp-mail.outlook.com": "outlook.office365.com",
            "smtp.mail.yahoo.com": "imap.mail.yahoo.com",
            "smtp.office365.com": "outlook.office365.com"
        }

        # Buscar mapeo directo
        if smtp_server in imap_mapping:
            return imap_mapping[smtp_server]

        # Intentar mapeo por dominio
        domain_mapping = {
            "gmail": "imap.gmail.com",
            "outlook": "outlook.office365.com",
            "yahoo": "imap.mail.yahoo.com",
            "office365": "outlook.office365.com"
        }

        smtp_lower = smtp_server.lower()
        for domain, imap_server in domain_mapping.items():
            if domain in smtp_lower:
                return imap_server

        # Si no encuentra mapeo, intentar convertir smtp a imap
        if smtp_server.startswith("smtp."):
            return smtp_server.replace("smtp.", "imap.")

        return None

    def _decode_mime_words(self, s):
        """
        Decodifica palabras MIME codificadas

        Args:
            s (str): String a decodificar

        Returns:
            str: String decodificado
        """
        if not s:
            return ""

        try:
            decoded_fragments = decode_header(s)
            decoded_string = ""

            for fragment, encoding in decoded_fragments:
                if isinstance(fragment, bytes):
                    if encoding:
                        try:
                            decoded_string += fragment.decode(encoding)
                        except (UnicodeDecodeError, LookupError):
                            decoded_string += fragment.decode('utf-8', 'ignore')
                    else:
                        decoded_string += fragment.decode('utf-8', 'ignore')
                else:
                    decoded_string += fragment

            return decoded_string
        except Exception:
            return str(s)

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