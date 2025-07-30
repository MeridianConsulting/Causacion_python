import sys
from pathlib import Path

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--console":
        run_console_mode()
    else:
        run_gui_mode()

def run_gui_mode():
    try:
        from excel_automation.ui_main import run_app
        print(" Iniciando interfaz con drag & drop...")
        sys.exit(run_app())
    except ImportError as e:
        print(f"Error al importar la interfaz gráfica: {e}")
        print("Instala las dependencias con: pip install -r requirements.txt")
        sys.exit(1)

def run_console_mode():
    from excel_automation.excel_processor import ExcelProcessor
    from config import Config
    
    print("=== Automatización de Procesos Excel (Modo Consola) ===")
    processor = ExcelProcessor()
    
    try:
        input_file = Config.INPUT_PATH / "archivo_entrada.xlsx"
        output_file = Config.OUTPUT_PATH / "archivo_procesado.xlsx"
        
        print(f"Procesando archivo: {input_file}")
        
        # Aquí irá tu lógica de automatización
        # processor.process_file(input_file, output_file)
        
        print(f"Archivo procesado guardado en: {output_file}")
        
    except Exception as e:
        print(f"Error durante el procesamiento: {e}")

if __name__ == "__main__":
    main() 