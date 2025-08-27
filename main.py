# main.py
"""
Punto de entrada principal del Bot de Python con interfaz tkinter.
Inicializa y ejecuta la ventana principal de la aplicación con gestión de perfiles.
"""

# Archivos relacionados: gui/main_window.py

import tkinter as tk
import sys
import os
from pathlib import Path


def check_dependencies():
    """
    Verifica que las dependencias básicas estén disponibles

    Returns:
        bool: True si todas las dependencias están disponibles
    """
    missing_deps = []

    # Verificar tkinter (debería estar incluido con Python)
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        missing_deps.append("tkinter")

    # Verificar dependencias opcionales pero recomendadas
    try:
        import openpyxl
    except ImportError:
        print("ADVERTENCIA: openpyxl no está instalado. Los reportes Excel no estarán disponibles.")
        print("Para instalar: pip install openpyxl")

    if missing_deps:
        print(f"ERROR: Dependencias faltantes: {', '.join(missing_deps)}")
        return False

    return True


def setup_directories():
    """Crea los directorios necesarios si no existen"""
    directories = [
        "config",  # Para configuraciones y perfiles
        "reports",  # Para reportes generados
        "logs"  # Para logs (futuro uso)
    ]

    for directory in directories:
        Path(directory).mkdir(exist_ok=True)


def main():
    """Función principal que inicializa la aplicación"""
    try:
        # Verificar dependencias
        if not check_dependencies():
            input("Presione Enter para salir...")
            sys.exit(1)

        # Configurar directorios
        setup_directories()

        # Crear ventana principal
        root = tk.Tk()

        # Configurar tema (si está disponible)
        try:
            root.tk.call("source", "azure.tcl")
            root.tk.call("set_theme", "light")
        except tk.TclError:
            # Si no hay tema personalizado, usar el tema por defecto
            pass

        # Importar y crear la aplicación principal
        from gui.main_window import MainWindow
        app = MainWindow(root)

        # Configurar manejo de cierre de ventana
        def on_closing():
            """Maneja el cierre de la aplicación"""
            if tk.messagebox.askyesno("Salir", "¿Está seguro de que desea cerrar la aplicación?"):
                root.quit()
                root.destroy()

        root.protocol("WM_DELETE_WINDOW", on_closing)

        # Mostrar información de inicio en consola
        print("=" * 50)
        print("🤖 BOT PYTHON - GESTIÓN DE PERFILES")
        print("=" * 50)
        print("✅ Sistema iniciado correctamente")
        print("📧 Configure SMTP desde la interfaz")
        print("📊 Cree perfiles de búsqueda")
        print("📋 Genere reportes Excel/CSV")
        print("=" * 50)

        # Iniciar loop principal
        root.mainloop()

    except ImportError as e:
        print(f"ERROR: No se pudo importar un módulo requerido: {e}")
        print("Asegúrese de que todos los archivos están en su lugar:")
        print("- gui/main_window.py")
        print("- gui/email_modal.py")
        print("- gui/profile_manager.py")
        print("- services/*.py")
        input("Presione Enter para salir...")
        sys.exit(1)

    except Exception as e:
        print(f"ERROR CRÍTICO: {e}")
        import traceback
        traceback.print_exc()
        input("Presione Enter para salir...")
        sys.exit(1)


if __name__ == "__main__":
    main()