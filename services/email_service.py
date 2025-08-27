# services/email_service.py
"""
Servicio dedicado a manejar conexiones SMTP y pruebas de conectividad.
Proporciona métodos para probar y validar configuraciones de email.
"""

# Archivos relacionados: Ninguno (servicio independiente)

import smtplib
import socket
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class EmailService:
    def __init__(self):
        """Inicializa el servicio de email"""
        self.connection = None

    def test_connection(self, credentials):
        """
        Prueba la conexión SMTP con las credenciales proporcionadas

        Args:
            credentials (dict): Diccionario con email, password, server, port

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            email = credentials.get("email", "").strip()
            password = credentials.get("password", "").strip()
            server = credentials.get("server", "").strip()
            port = credentials.get("port")

            # Validar que todos los campos estén presentes
            if not all([email, password, server, port]):
                return False, "Faltan campos requeridos"

            # Convertir puerto a entero
            try:
                port = int(port)
            except (ValueError, TypeError):
                return False, "Puerto debe ser un numero valido"

            # Intentar conexión SMTP
            smtp_server = smtplib.SMTP(server, port, timeout=30)
            smtp_server.starttls()  # Habilitar seguridad TLS
            smtp_server.login(email, password)
            smtp_server.quit()

            return True, "Conexion SMTP exitosa"

        except smtplib.SMTPAuthenticationError as e:
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            return False, f"Error de autenticacion: {error_msg}"

        except smtplib.SMTPConnectError as e:
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            return False, f"No se pudo conectar al servidor: {error_msg}"

        except smtplib.SMTPServerDisconnected as e:
            return False, "El servidor se desconecto inesperadamente"

        except socket.timeout as e:
            return False, "Tiempo de conexion agotado"

        except socket.gaierror as e:
            return False, "Error de DNS. Verifique el servidor SMTP"

        except Exception as e:
            # Limpiar caracteres problemáticos del mensaje de error
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            return False, f"Error: {error_msg}"

    def send_test_email(self, credentials, recipient=None):
        """
        Envía un email de prueba

        Args:
            credentials (dict): Credenciales SMTP
            recipient (str): Email del destinatario (opcional, usa sender si no se especifica)

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            email = credentials.get("email", "").strip()
            password = credentials.get("password", "").strip()
            server = credentials.get("server", "").strip()
            port = int(credentials.get("port", 587))

            if not recipient:
                recipient = email  # Enviar a sí mismo si no se especifica destinatario

            # Crear mensaje
            message = MIMEMultipart()
            message["From"] = email
            message["To"] = recipient
            message["Subject"] = "Prueba de conexion SMTP - Bot Python"

            # Cuerpo del mensaje (sin caracteres especiales problemáticos)
            body = """Hola!

Este es un email de prueba enviado desde tu Bot de Python.
La configuracion SMTP esta funcionando correctamente.

Saludos,
Bot Python"""

            message.attach(MIMEText(body, "plain", "utf-8"))

            # Enviar email
            smtp_server = smtplib.SMTP(server, port, timeout=30)
            smtp_server.starttls()
            smtp_server.login(email, password)
            smtp_server.send_message(message)
            smtp_server.quit()

            return True, f"Email de prueba enviado exitosamente a {recipient}"

        except Exception as e:
            # Limpiar caracteres problemáticos del mensaje de error
            error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            return False, f"Error enviando email de prueba: {error_msg}"

    def validate_email_format(self, email):
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

    def get_common_smtp_settings(self, email_provider=None):
        """
        Obtiene configuraciones SMTP comunes para proveedores conocidos

        Args:
            email_provider (str): Proveedor de email (gmail, outlook, yahoo, etc.)

        Returns:
            dict: Configuraciones SMTP recomendadas
        """
        settings = {
            "gmail": {
                "server": "smtp.gmail.com",
                "port": 587,
                "tls": True
            },
            "outlook": {
                "server": "smtp-mail.outlook.com",
                "port": 587,
                "tls": True
            },
            "yahoo": {
                "server": "smtp.mail.yahoo.com",
                "port": 587,
                "tls": True
            },
            "hotmail": {
                "server": "smtp-mail.outlook.com",
                "port": 587,
                "tls": True
            }
        }

        if email_provider and email_provider.lower() in settings:
            return settings[email_provider.lower()]

        return settings.get("gmail")  # Gmail por defecto

    def detect_provider_from_email(self, email):
        """
        Detecta el proveedor de email basado en el dominio

        Args:
            email (str): Dirección de email

        Returns:
            str: Proveedor detectado o None si no se reconoce
        """
        if not email or "@" not in email:
            return None

        try:
            domain = email.strip().split("@")[1].lower()

            provider_domains = {
                "gmail.com": "gmail",
                "outlook.com": "outlook",
                "hotmail.com": "hotmail",
                "yahoo.com": "yahoo",
                "yahoo.es": "yahoo"
            }

            return provider_domains.get(domain)
        except (IndexError, AttributeError):
            return None

    def get_connection_diagnostics(self, credentials):
        """
        Obtiene información de diagnóstico sobre la conexión

        Args:
            credentials (dict): Credenciales SMTP

        Returns:
            dict: Información de diagnóstico
        """
        diagnostics = {
            "server_reachable": False,
            "port_open": False,
            "tls_available": False,
            "auth_valid": False,
            "provider_detected": None
        }

        try:
            email = credentials.get("email", "").strip()
            server = credentials.get("server", "").strip()
            port = int(credentials.get("port", 587))

            # Detectar proveedor
            diagnostics["provider_detected"] = self.detect_provider_from_email(email)

            # Probar alcance del servidor
            try:
                socket.create_connection((server, port), timeout=10)
                diagnostics["server_reachable"] = True
                diagnostics["port_open"] = True
            except (socket.timeout, socket.error):
                return diagnostics

            # Probar TLS y autenticación
            try:
                smtp_server = smtplib.SMTP(server, port, timeout=30)
                diagnostics["tls_available"] = smtp_server.has_extn('STARTTLS')

                if diagnostics["tls_available"]:
                    smtp_server.starttls()

                smtp_server.login(email, credentials.get("password", ""))
                diagnostics["auth_valid"] = True
                smtp_server.quit()

            except smtplib.SMTPAuthenticationError:
                pass  # auth_valid queda en False
            except Exception:
                pass

        except Exception:
            pass

        return diagnostics

    def clean_string(self, text):
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