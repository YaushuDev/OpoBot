# main.py
"""
Punto de entrada principal del Bot de Python con interfaz tkinter.
Inicializa y ejecuta la ventana principal de la aplicación.
"""

# Archivos relacionados: gui/main_window.py

import tkinter as tk
from gui.main_window import MainWindow

def main():
    """Función principal que inicializa la aplicación"""
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()