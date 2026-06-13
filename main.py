# main.py
import sys
import os
import cleandataset
from pathlib import Path
import tkinter as tk
from tkinter import filedialog


def get_file_paths_interactive():
    # Hide the main root window
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)  # Bring dialog to front

    # Open File Dialog
    input_path = filedialog.askopenfilename(
        title="Select Input Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not input_path:
        return None, None  # User cancelled

    # Auto-generate output path
    file_obj = Path(input_path)
    directory = file_obj.parent
    stem = file_obj.stem
    suffix = file_obj.suffix

    # Add "_output" before extension
    new_filename = f"{stem}_output{suffix}"

    # Optional: Ask where to save or confirm name (using Save Dialog pre-filled)
    # This allows the user to change the name or location if they want
    output_path = filedialog.asksaveasfilename(
        title="Save Output File",
        initialdir=directory,
        initialfile=new_filename,
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    return input_path, output_path


def main():
    # 1. Check for Command Line Arguments (for .bat usage)
    if len(sys.argv) >= 2:
        input_path = sys.argv[1]
        if len(sys.argv) >= 3:
            output_path = sys.argv[2]
        else:
            # Auto-generate if only input provided
            file_obj = Path(input_path)
            output_path = os.path.join(file_obj.parent, f"{file_obj.stem}_output{file_obj.suffix}")

    else:
        # 2. Interactive Visual Selector
        print("Opening file selector...")
        input_path, output_path = get_file_paths_interactive()

    # Validation
    if not input_path:
        print("❌ No file selected. Exiting.")
        return

    if not output_path:
        print("❌ No output location selected. Exiting.")
        return

    if not os.path.exists(input_path):
        print(f"❌ Error: Input file not found: {input_path}")
        input("Press Enter to exit...")
        return

    # Execution
    try:
        print(f"📁 Processing: {os.path.basename(input_path)}")
        # 1. Limpiar y Formatear
        cleandataset.arrange_and_format_file(input_path, output_path)

        # 2. Analizar Uso de Materia Prima (Nuevo paso)
        # Usamos el mismo archivo output_path como entrada y salida (se modifica in-place)
        import movement_fit
        movement_fit.analyze_material_usage(output_path, output_path)

        print("\n✅ Proceso COMPLETO terminado (Limpieza + Análisis MP)!")
        print(f"💾 Archivo final: {output_path}")

    except Exception as e:
        print(f"❌ Un error ocurrió: {e}")

if __name__ == "__main__":
    main()