# services/email_send_service.py
"""
Servicio para envío de correos con adjuntos de reportes Excel.
Maneja el envío automático de reportes generados por correo electrónico.
"""

import smtplib
import socket
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import datetime


class EmailSendService:
    def __init__(self):
        """Inicializa el servicio de envío de emails"""
        pass

    def send_report_email(self, report_file_path, email_config=None):
        """
        Envía un reporte por correo electrónico

        Args:
            report_file_path (str): Ruta del archivo de reporte a enviar
            email_config (dict, optional): Configuración específica de envío

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            from services.config_service import ConfigService
            config_service = ConfigService()

            # Verificar que existan credenciales SMTP
            if not config_service.credentials_exist():
                return False, "No hay credenciales SMTP configuradas"

            # Cargar credenciales SMTP
            smtp_credentials = config_service.load_credentials()
            if not smtp_credentials:
                return False, "No se pudieron cargar las credenciales SMTP"

            # Cargar configuración de envío si no se proporciona
            if not email_config:
                email_config = config_service.load_email_send_config()

            if not email_config or not email_config.get("enabled", False):
                return False, "Configuracion de envio no esta habilitada"

            # Validar archivo de reporte
            report_path = Path(report_file_path)
            if not report_path.exists():
                return False, f"El archivo de reporte no existe: {report_file_path}"

            # Preparar datos del email
            subject = email_config.get("subject", "Reporte de Registros de Bot - OpoBot")
            recipient = email_config.get("recipient", "").strip()
            cc_emails = email_config.get("cc", "").strip()

            if not recipient:
                return False, "No hay destinatario configurado"

            # Preparar lista de destinatarios
            recipients = [recipient]
            cc_list = []
            if cc_emails:
                cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]

            # Crear mensaje
            success, message = self._create_and_send_message(
                smtp_credentials, subject, recipient, cc_list, report_path
            )

            return success, message

        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error enviando reporte: {error_msg}"

    def send_test_email(self, email_config):
        """
        Envía un email de prueba sin adjuntos

        Args:
            email_config (dict): Configuración de envío

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            from services.config_service import ConfigService
            config_service = ConfigService()

            # Verificar credenciales SMTP
            if not config_service.credentials_exist():
                return False, "No hay credenciales SMTP configuradas"

            smtp_credentials = config_service.load_credentials()
            if not smtp_credentials:
                return False, "No se pudieron cargar las credenciales SMTP"

            # Validar configuración de prueba
            subject = email_config.get("subject", "Prueba de Envio - OpoBot")
            recipient = email_config.get("recipient", "").strip()
            cc_emails = email_config.get("cc", "").strip()

            if not recipient:
                return False, "No hay destinatario configurado"

            # Preparar lista CC
            cc_list = []
            if cc_emails:
                cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]

            # Crear y enviar mensaje de prueba
            success, message = self._send_test_message(
                smtp_credentials, subject, recipient, cc_list
            )

            return success, message

        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error enviando prueba: {error_msg}"

    def _create_and_send_message(self, smtp_credentials, subject, recipient, cc_list, report_path):
        """
        Crea y envía el mensaje con adjunto

        Args:
            smtp_credentials (dict): Credenciales SMTP
            subject (str): Asunto del correo
            recipient (str): Destinatario principal
            cc_list (list): Lista de emails CC
            report_path (Path): Ruta del archivo adjunto

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            sender_email = smtp_credentials.get("email", "").strip()
            sender_password = smtp_credentials.get("password", "").strip()
            smtp_server = smtp_credentials.get("server", "").strip()
            smtp_port = int(smtp_credentials.get("port", 587))

            # Crear mensaje
            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = recipient
            message["Subject"] = subject

            if cc_list:
                message["Cc"] = ", ".join(cc_list)

            # Cuerpo del mensaje
            body = self._get_email_body()
            message.attach(MIMEText(body, "plain", "utf-8"))

            # Adjuntar archivo
            self._attach_file(message, report_path)

            # Preparar lista completa de destinatarios
            all_recipients = [recipient] + cc_list

            # Enviar email
            success, send_message = self._send_message(
                sender_email, sender_password, smtp_server, smtp_port,
                message, all_recipients
            )

            if success:
                recipients_info = f"Destinatario: {recipient}"
                if cc_list:
                    recipients_info += f", CC: {', '.join(cc_list)}"

                return True, f"Reporte enviado exitosamente. {recipients_info}"
            else:
                return False, send_message

        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error creando mensaje: {error_msg}"

    def _send_test_message(self, smtp_credentials, subject, recipient, cc_list):
        """
        Envía un mensaje de prueba sin adjuntos

        Args:
            smtp_credentials (dict): Credenciales SMTP
            subject (str): Asunto del correo
            recipient (str): Destinatario principal
            cc_list (list): Lista de emails CC

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            sender_email = smtp_credentials.get("email", "").strip()
            sender_password = smtp_credentials.get("password", "").strip()
            smtp_server = smtp_credentials.get("server", "").strip()
            smtp_port = int(smtp_credentials.get("port", 587))

            # Crear mensaje de prueba
            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = recipient
            message["Subject"] = f"PRUEBA - {subject}"

            if cc_list:
                message["Cc"] = ", ".join(cc_list)

            # Cuerpo de prueba
            test_body = self._get_test_email_body()
            message.attach(MIMEText(test_body, "plain", "utf-8"))

            # Preparar destinatarios
            all_recipients = [recipient] + cc_list

            # Enviar
            success, send_message = self._send_message(
                sender_email, sender_password, smtp_server, smtp_port,
                message, all_recipients
            )

            if success:
                recipients_info = f"Destinatario: {recipient}"
                if cc_list:
                    recipients_info += f", CC: {', '.join(cc_list)}"

                return True, f"Email de prueba enviado exitosamente. {recipients_info}"
            else:
                return False, send_message

        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error enviando prueba: {error_msg}"

    def _attach_file(self, message, file_path):
        """
        Adjunta un archivo al mensaje

        Args:
            message: Objeto MIMEMultipart
            file_path (Path): Ruta del archivo a adjuntar
        """
        try:
            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())

            # Codificar archivo en base64
            encoders.encode_base64(part)

            # Añadir header
            filename = file_path.name
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {filename}'
            )

            # Adjuntar al mensaje
            message.attach(part)

        except Exception as e:
            raise Exception(f"Error adjuntando archivo: {str(e)}")

    def _send_message(self, sender_email, sender_password, smtp_server, smtp_port, message, recipients):
        """
        Envía el mensaje por SMTP

        Args:
            sender_email (str): Email del remitente
            sender_password (str): Contraseña del remitente
            smtp_server (str): Servidor SMTP
            smtp_port (int): Puerto SMTP
            message: Mensaje a enviar
            recipients (list): Lista de destinatarios

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            # Crear conexión SMTP
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            server.starttls()  # Habilitar TLS
            server.login(sender_email, sender_password)

            # Enviar mensaje
            server.send_message(message, to_addrs=recipients)
            server.quit()

            return True, "Mensaje enviado correctamente"

        except smtplib.SMTPAuthenticationError as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error de autenticacion SMTP: {error_msg}"

        except smtplib.SMTPConnectError as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error de conexion SMTP: {error_msg}"

        except smtplib.SMTPServerDisconnected:
            return False, "El servidor SMTP se desconecto inesperadamente"

        except smtplib.SMTPRecipientsRefused as e:
            error_msg = self._clean_string(str(e))
            return False, f"Destinatarios rechazados: {error_msg}"

        except socket.timeout:
            return False, "Tiempo de conexion SMTP agotado"

        except socket.gaierror:
            return False, "Error de DNS. Verifique el servidor SMTP"

        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error enviando mensaje: {error_msg}"

    def _get_email_body(self):
        """
        Obtiene el cuerpo estándar del email para reportes

        Returns:
            str: Cuerpo del email
        """
        return """Estimado/a Usuario,

Adjunto encontrará el reporte de registros de ejecución correspondiente.

Saludos cordiales,
Bot de Gestión y Registros OpoBot"""

    def _get_test_email_body(self):
        """
        Obtiene el cuerpo del email de prueba

        Returns:
            str: Cuerpo del email de prueba
        """
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        return f"""Hola!

Este es un email de PRUEBA enviado desde tu Bot de Python.

La configuracion de envio de reportes esta funcionando correctamente.

Fecha y hora de la prueba: {current_time}

Cuando el programador automatico se ejecute, enviara los reportes Excel usando esta misma configuracion.

Saludos cordiales,
Bot de Gestión y Registros OpoBot"""

    def get_send_config_status(self):
        """
        Obtiene el estado de la configuración de envío

        Returns:
            dict: Estado de la configuración
        """
        try:
            from services.config_service import ConfigService
            config_service = ConfigService()

            # Verificar credenciales SMTP
            has_smtp = config_service.credentials_exist()

            # Verificar configuración de envío
            send_config = config_service.load_email_send_config()
            has_send_config = send_config is not None
            is_enabled = send_config.get("enabled", False) if send_config else False

            status = {
                "has_smtp_credentials": has_smtp,
                "has_send_config": has_send_config,
                "is_enabled": is_enabled,
                "ready_to_send": has_smtp and has_send_config and is_enabled
            }

            if send_config and is_enabled:
                status["recipient"] = send_config.get("recipient", "")
                status["subject"] = send_config.get("subject", "")
                status["cc"] = send_config.get("cc", "")

            return status

        except Exception as e:
            print(f"Error obteniendo estado de configuracion de envio: {e}")
            return {
                "has_smtp_credentials": False,
                "has_send_config": False,
                "is_enabled": False,
                "ready_to_send": False
            }

    def validate_email_config(self, email_config):
        """
        Valida la configuración de email antes de enviar

        Args:
            email_config (dict): Configuración a validar

        Returns:
            tuple: (valid: bool, message: str)
        """
        try:
            if not email_config:
                return False, "No hay configuracion de envio"

            if not email_config.get("enabled", False):
                return False, "Configuracion de envio deshabilitada"

            # Validar campos requeridos
            required_fields = ["subject", "recipient"]
            for field in required_fields:
                if not email_config.get(field, "").strip():
                    return False, f"Campo requerido faltante: {field}"

            # Validar formato de emails
            recipient = email_config.get("recipient", "").strip()
            if not self._validate_email_format(recipient):
                return False, f"Formato de destinatario invalido: {recipient}"

            # Validar CC si está presente
            cc_emails = email_config.get("cc", "").strip()
            if cc_emails:
                cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
                for email in cc_list:
                    if not self._validate_email_format(email):
                        return False, f"Formato de CC invalido: {email}"

            return True, "Configuracion valida"

        except Exception as e:
            error_msg = self._clean_string(str(e))
            return False, f"Error validando configuracion: {error_msg}"

    def _validate_email_format(self, email):
        """
        Valida el formato de un email

        Args:
            email (str): Email a validar

        Returns:
            bool: True si es válido
        """
        if not email or not isinstance(email, str):
            return False

        import re
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email.strip()) is not None

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
            text = str(text)
            text = text.replace('\xa0', ' ')  # Espacio no-rompible
            text = text.replace('\u2019', "'")  # Apostrofe curvo
            text = text.replace('\u2018', "'")  # Apostrofe curvo
            text = text.replace('\u201c', '"')  # Comilla curva
            text = text.replace('\u201d', '"')  # Comilla curva

            # Codificar y decodificar para limpiar caracteres problemáticos
            return text.encode('ascii', 'ignore').decode('ascii')
        except Exception:
            return str(text) if text else ""