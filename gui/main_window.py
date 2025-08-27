# gui/main_window.py
"""
Ventana principal de la aplicación con diseño de 2 columnas.
Maneja la interfaz principal y la apertura del modal de configuración de email.
"""

# Archivos relacionados: gui/email_modal.py

import tkinter as tk
from tkinter import ttk
from gui.email_modal import EmailModal


class MainWindow:
    def __init__(self, root):
        """
        Inicializa la ventana principal con diseño de 2 columnas

        Args:
            root (tk.Tk): Ventana raíz de tkinter
        """
        self.root = root
        self.setup_window()
        self.create_layout()

    def setup_window(self):
        """Configura las propiedades básicas de la ventana"""
        self.root.title("Bot Python - Configuración SMTP")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # Configurar grid para que se expanda
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

    def create_layout(self):
        """Crea el layout de 2 columnas de la interfaz principal"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # Columna izquierda
        self.create_left_column(main_frame)

        # Columna derecha
        self.create_right_column(main_frame)

    def create_left_column(self, parent):
        """
        Crea la columna izquierda con configuraciones y botones principales

        Args:
            parent: Frame padre donde se colocará la columna
        """
        left_frame = ttk.LabelFrame(parent, text="Configuraciones", padding="10")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        left_frame.grid_rowconfigure(1, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)

        # Título
        title_label = ttk.Label(left_frame, text="Panel de Configuración",
                                font=("Arial", 12, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="w")

        # Botón para abrir modal de email
        email_button = ttk.Button(left_frame, text="Configurar Email SMTP",
                                  command=self.open_email_modal, width=25)
        email_button.grid(row=1, column=0, sticky="n")

    def create_right_column(self, parent):
        """
        Crea la columna derecha para futuras funcionalidades

        Args:
            parent: Frame padre donde se colocará la columna
        """
        right_frame = ttk.LabelFrame(parent, text="Información", padding="10")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        right_frame.grid_rowconfigure(1, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        # Título
        title_label = ttk.Label(right_frame, text="Estado del Sistema",
                                font=("Arial", 12, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="w")

        # Area de texto para información
        self.info_text = tk.Text(right_frame, height=20, width=40,
                                 state="disabled", wrap="word")
        self.info_text.grid(row=1, column=0, sticky="nsew")

        # Scrollbar para el área de texto
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical",
                                  command=self.info_text.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.info_text.configure(yscrollcommand=scrollbar.set)

        # Mensaje inicial
        self.update_info("Sistema iniciado.\nPuede configurar su email SMTP desde el panel izquierdo.")

    def open_email_modal(self):
        """Abre el modal de configuración de email"""
        EmailModal(self.root, self.on_email_configured)

    def on_email_configured(self, success, message):
        """
        Callback llamado cuando se configura el email

        Args:
            success (bool): True si la configuración fue exitosa
            message (str): Mensaje de estado
        """
        # Limpiar mensaje de caracteres problemáticos
        clean_message = self.clean_message(message)
        status = "✓ EXITO" if success else "✗ ERROR"
        self.update_info(f"{status}: {clean_message}")

    def update_info(self, message):
        """
        Actualiza el área de información en la columna derecha

        Args:
            message (str): Mensaje a mostrar
        """
        clean_message = self.clean_message(message)
        self.info_text.config(state="normal")
        self.info_text.insert("end", f"\n{clean_message}")
        self.info_text.see("end")
        self.info_text.config(state="disabled")

    def clean_message(self, message):
        """
        Limpia un mensaje de caracteres problemáticos

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