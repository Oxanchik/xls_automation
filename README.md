# Processing data automation of the production and sales database selection
 This is a small Excel-processing pipeline with three scripts: one script cleans and formats an input workbook, another builds a manufacturing/material-usage analysis sheet, and the main script ties them together into one workflow.

Clone this repository on your computer.

If you have Python interpreter installed on your computer you can launch this script by double-clicking 'analyze.bat' file.

There is also another script that only merges and cleans the data on several xlsx files. You can launch it by double-clicking 'merge_script\xls-merge.bat' file. Depends on specific column order: applies date format to column A (first column) - dd/mm/yyyy: "Fecha", applies number format with 2 decimals for columns G (7th) and J (10th): "Unidades" and "Uni. stock", applies text format to column E: "Cód. artículo".

## What each file does

`main.py`
This is the entry point. It asks you to pick an Excel file and where to save the output, or it can take paths from the command line. Then it runs the cleaning step first and the material-usage analysis second, producing one final processed workbook.

`cleandataset.py`
This script loads an Excel file into a DataFrame, cleans the CodigoArticulo field by trimming extra spaces, removes some unwanted columns, drops fully empty columns, saves the cleaned data back to Excel, and applies formatting. It also formats the header row, sets date formats for Fecha and FechaRegistro, formats Unidades as a numeric value, and auto-adjusts column widths.

`movement_fit.py`
This script analyzes stock movements related to manufacturing. It searches for a Comentario column flexibly, filters rows where OrigenMovimiento == "F", splits those into finished / semifinished products (TipoMovimiento == 1) and raw materials (TipoMovimiento == 2), aggregates quantities by code/batch/comment, merges the two summaries by "Comentario", aggregates nominal weight in "Peso_nominal" field for each product that is reported in units in "Unidades_PT" field instead of kg, by crossing the current table with the reference table from the sheet "PT_UDS" of the file "Lista_pesos.xlsx" **(IMPORTANT! There should be the "Lista_pesos.xlsx" file in the same folder that the output file!)**, and writes the result into a new Excel sheet called "uso_MP".

`loss_by_batch.py`
This script analyzes the losses per batch of finished / semifinished product. It groups first the products by "CodigoArticulo" and "Partida" fields to avoid duplication of the finished / semifinished product weight. Reads the "uso_MP" sheet, calculates weights and generates the "mermas_lotes_PT" sheet. **(IMPORTANT! In the "use_MP" sheet there should be "Peso_nominal" field.)** Rule: If there is "Peso_nominal" it is multiplied by "Unidades_PT", if there is no nominal weight indicated then the value of "Unidades_PT" is taken as it is (in kg).

## How the workflow works
You select an input Excel file.

The cleaning script reads the workbook, standardizes it, removes useless columns, and formats the output.

The analysis script reopens that cleaned workbook, computes a production-vs-consumption table, and stores it in a new sheet.

The final Excel file contains both the cleaned data, the "uso_MP" and  "mermas_lotes_PT" analysis sheet.

## Important details
The analysis depends on specific column names such as TipoMovimiento, OrigenMovimiento, CodigoArticulo, Partida, DescripcionArticulo, Unidades, and UnidadMedida1_; if they are missing, the script stops with an error message.

The scripts are designed for Spanish-language ERP-style stock movement exports, especially manufacturing data where product output and material input are linked through a shared "Comentario" value.

## Initial Data

- Unique key for each product and raw material: "CodigoArticulo" and "Partida". 
- "Comentario" to find matches between products and raw material.
- "TipoMovimiento": 1 = Input, 2 = Output
- To distinguish the finished product from the raw material: 
  - Finished Product has 1 in the field "TipoMovimiento" and F in the field "OrigenMovimiento"
  - Raw Material has 2 in the field "TipoMovimiento" and F in the field "OrigenMovimiento"
- The field "OrigenMovimiento": 
  - A = Albarán de venta
  - B = Albarán de compra
  - E = Entrada desde otro articulo o una entrada por apertura de nuevo ejercicio
  - F = Fabricación 
  - P = Apertura de ejercicio
  - S = Salida por muestras o correcciones
  - T = Traspasos (de un almacén a otro, o de un artículo a otro)
