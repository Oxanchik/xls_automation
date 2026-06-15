import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def find_column(df, target_name):
    """Busca una columna ignorando mayúsculas y espacios extra."""
    target_clean = target_name.lower().strip()
    for col in df.columns:
        if col.lower().strip() == target_clean:
            return col
    # Búsqueda parcial si no hay exacta
    for col in df.columns:
        if target_clean in col.lower():
            return col
    return None


def analyze_material_usage(input_file, output_file):
    print(f"🔍 Iniciando análisis de uso de materia prima en: {os.path.basename(input_file)}")

    # 1. Cargar datos
    df = pd.read_excel(input_file, engine="openpyxl", dtype={'CodigoArticulo': str, 'Partida': str})

    # --- Limpieza de nombres de columnas ---
    df.columns = df.columns.str.strip()

    # Buscar columna Comentario de forma flexible
    col_comment_real = find_column(df, "Comentario")
    if not col_comment_real:
        print("❌ Error: No se encontró la columna 'Comentario'.")
        return

    # Renombrar a "Comentario" para estandarizar
    if col_comment_real != "Comentario":
        df.rename(columns={col_comment_real: "Comentario"}, inplace=True)

    # Nombres de columnas estandarizados
    col_type = "TipoMovimiento"
    col_origin = "OrigenMovimiento"
    col_code = "CodigoArticulo"
    col_batch = "Partida"
    col_desc = "DescripcionArticulo"
    col_qty = "Unidades"
    col_unit = "UnidadMedida1_"

    # Validar columnas esenciales
    required_cols = ["Comentario", col_type, col_origin, col_code, col_qty, col_desc, col_unit]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        print(f"❌ Faltan columnas esenciales: {missing}")
        return

    # 2. Identificar Productos Terminados (PT) y Materia Prima (MP)
    df_f = df[df[col_origin] == 'F'].copy()

    if df_f.empty:
        print("⚠️ No se encontraron movimientos de tipo 'F' (Fabricación).")
        return

    df_pt = df_f[df_f[col_type] == 1].copy()  # Entrada = PT
    df_mp = df_f[df_f[col_type] == 2].copy()  # Salida = MP

    print(f"   - Filas de Fabricación (F): {len(df_f)}")
    print(f"   - Productos Terminados (Entrada/1): {len(df_pt)}")
    print(f"   - Materia Prima (Salida/2): {len(df_mp)}")

    if df_pt.empty or df_mp.empty:
        print("⚠️ No hay suficientes datos para cruzar.")
        return

    # 3. Agrupar Datos

    # pt_summary = df_pt.groupby([col_code, col_batch, "Comentario"]).agg(
    #     DescripcionArticulo=(col_desc, 'first'),
    #     CantidadProducida=(col_qty, 'sum'),
    #     UnidadMedida=(col_unit, 'first')
    # ).reset_index()
    #
    # Antes de reset_index(): Las columnas col_code, col_batch y "Comentario" son el índice del DataFrame (no columnas regulares).
    # Si omites .reset_index(), ocurrirá lo siguiente:
    # 1.	Las columnas agrupadas permanecen como índice: No podrás acceder a col_code, col_batch o "Comentario" como columnas normales usando df["columna"].
    # 2.	Problemas al exportar: Al guardar el DataFrame a Excel o CSV, el índice se guardará de forma diferente, posiblemente creando columnas no deseadas o perdiendo la estructura esperada.
    # 3.	Dificultad para operaciones posteriores: Si intentas hacer merges, filtros o accesos por nombre de columna, fallará porque esas columnas están en el índice, no en las columnas del DataFrame.
    # 4.	Posible error al escribir a Excel: La función dataframe_to_rows de openpyxl podría no manejar correctamente un DataFrame con MultiIndex, generando un formato inesperado en el archivo de salida.

    # Al usar as_index=False directamente dentro de groupby(), le indicas a pandas que no convierta las columnas de agrupación en el índice,
    # sino que las mantenga como columnas normales desde el principio.

    # Resumen PT: Agrupar por Código, Partida y Comentario
    pt_summary = df_pt.groupby([col_code, col_batch, "Comentario"], as_index=False).agg(
        DescripcionArticulo_PT=(col_desc, "first"),
        Unidades_PT=(col_qty, "sum"),
        UnidadMedida_PT=(col_unit, "first")
    )

    # Resumen MP: Agrupar por Código, Partida y Comentario
    # IMPORTANTE: Incluimos también Descripción y Unidad en la agrupación o las tomamos después
    mp_summary = df_mp.groupby([col_code, col_batch, "Comentario"], as_index=False).agg(
        DescripcionArticulo_MP=(col_desc, "first"),
        Unidades_MP=(col_qty, "sum"),
        UnidadMedida_MP=(col_unit, "first")
    )

    # 4. Cruzar Datos (Merge)
    # how = "inner" -> hace que aparezcan en la hoja "uso_MP" solo PT / MP que coinciden.
    # how = "left" -> hace que aparezcan en la hoja "uso_MP" también PT que no tienen "pareja" en MP.
    merged = pd.merge(
        pt_summary,
        mp_summary,
        on="Comentario",
        how="left",
        suffixes=("_PT", "_MP")
    )

    if merged.empty:
        print("⚠️ El cruce de datos no produjo resultados.")
        return

    # 5. Crear Tabla Normal (Plana)
    final_table = merged[[
        "CodigoArticulo_PT",
        "Partida_PT",
        "DescripcionArticulo_PT",
        "Unidades_PT",
        "UnidadMedida_PT",
        "CodigoArticulo_MP",
        "Partida_MP",
        "DescripcionArticulo_MP",
        "Unidades_MP",
        "UnidadMedida_MP",
        "Comentario"
    ]].copy()

    # 6. Ordenar por Producto Terminado y luego por Materia Prima
    final_table.sort_values(by=['CodigoArticulo_PT', 'Partida_PT', 'CodigoArticulo_MP', 'Partida_MP'], inplace=True)

    # 7. Guardar en Excel
    print("Guardando hoja 'uso_MP' en el Excel...")
    wb = load_workbook(output_file)

    # Eliminar hoja si existe para evitar duplicados
    if "uso_MP" in wb.sheetnames:
        del wb["uso_MP"]

    ws = wb.create_sheet("uso_MP")

    # 8. Escribir cabeceras manualmente (para tener control total)
    ws.append(final_table.columns.tolist())

    # 9. Escribir datos fila por fila directamente desde el DataFrame
    # iterrows() o itertuples() son más eficientes que dataframe_to_rows para escritura directa
    for row in final_table.itertuples(index=False, name=None):
        # name=None evita crear NamedTuples, devuelve tuplas simples, más seguro para openpyxl
        ws.append(row)

    # Al añadir name=None, fuerzas a pandas a devolver tuplas estándar en lugar de NamedTuples.
    # Esto evita errores si tus columnas tienen espacios, guiones o caracteres especiales (ej. "Unidad Medida"), ya que NamedTuple intentaría crear atributos inválidos en Python. ws.append funciona perfecto con tuplas estándar.

    # 10. Estilos
    header_font = Font(bold=True) # Fuente predeterminada (Calibri) en negrita
    gray_fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')

    # 11. Obtener dimensiones reales tras escribir los datos
    max_col = ws.max_column
    max_row = ws.max_row # No lo necesitamos para la cabecera, pero sí para anchos

    # 12. Aplicar formato SOLO a la cabecera (Fila 1)
    # Iteramos solo por las columnas existentes
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = gray_fill

    # 13. Ajuste de anchos
    print("Ajustando anchos de columna...")
    for col_idx in range(1, max_col + 1):
        max_length = 0
        for row_idx in range(1, max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                length = len(str(val))
                if length > max_length:
                    max_length = length
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 50)

    wb.save(output_file)
    print(f"✅ Hoja 'uso_MP' creada exitosamente con {len(final_table)} registros.")


if __name__ == '__main__':
    input_p = r"C:\Users\Labo\Downloads\MovimientoStock_output.xlsx"
    if os.path.exists(input_p):
        analyze_material_usage(input_p, input_p)
    else:
        print("Archivo no encontrado.")
