import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


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

    # Resumen PT: Agrupar por Código, Partida y Comentario
    pt_summary = df_pt.groupby([col_code, col_batch, "Comentario"], as_index=False).agg(
        DescripcionArticulo_PT=(col_desc, "first"),
        Unidades_Producidas_PT=(col_qty, "sum"),
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
    # how = "inner" so only matched PT / MP pairs appear in uso_MP.
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
        "Unidades_Producidas_PT",
        "CodigoArticulo_MP",
        "Partida_MP",
        "DescripcionArticulo_MP",
        "Unidades_MP",
        "UnidadMedida_PT",
        "UnidadMedida_MP",
        "Comentario"
    ]].copy()


    # Ordenar por Producto Terminado y luego por Materia Prima
    final_table.sort_values(by=['CodigoArticulo_PT', 'Partida_PT', 'CodigoArticulo_MP', 'Partida_MP'], inplace=True)

    # 6. Guardar en Excel
    print("Guardando hoja 'uso_MP' en el Excel...")
    wb = load_workbook(output_file)

    if "uso_MP" in wb.sheetnames:
        del wb["uso_MP"]

    ws = wb.create_sheet("uso_MP")

    for r_idx, row in enumerate(dataframe_to_rows(final_table, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Ajuste de anchos
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                val_len = len(str(cell.value))
                if val_len > max_length:
                    max_length = val_len
            except:
                pass
        ws.column_dimensions[column].width = min(max_length + 2, 50)

    wb.save(output_file)
    print(f"✅ Hoja 'uso_MP' creada exitosamente con {len(final_table)} registros.")


if __name__ == '__main__':
    input_p = r"C:\Users\Labo\Downloads\MovimientoStock_output.xlsx"
    if os.path.exists(input_p):
        analyze_material_usage(input_p, input_p)
    else:
        print("Archivo no encontrado.")