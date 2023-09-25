import os
import pandas as pd
from datetime import datetime

# Ruta de la carpeta donde se encuentran los archivos de Excel
input_folder = "files"
# Ruta de la carpeta donde se guardarán los resultados
output_folder = "results"

# Obtener la lista de archivos Excel en la carpeta de entrada
excel_files = [file for file in os.listdir(input_folder) if file.endswith(".xlsx")]

# Crear la carpeta de resultados si no existe
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Iterar a través de los archivos Excel
for excel_file in excel_files:
    # Leer el archivo Excel y la hoja "Control"
    file_path = os.path.join(input_folder, excel_file)

    # Leer los archivos Excel y las hojas de interés
    df1 = pd.read_excel(file_path, sheet_name="Ventas")
    df2 = pd.read_excel(file_path, sheet_name="Control")

    # Realizar el merge basado en la columna "ID"
    merged_df = pd.merge(df1, df2, on="Codigo KA/OGK", how="outer")

    # Obtener la fecha y hora actual para el nombre de la carpeta de resultados
    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S_%f")

    # Crear la carpeta de resultados para este archivo
    result_folder = os.path.join(output_folder, timestamp)
    os.makedirs(result_folder)

    # Guardar los valores en un nuevo archivo Excel en la carpeta de resultados
    result_file_path = os.path.join(result_folder, excel_file)
    merged_df.to_excel(result_file_path, index=False)

    # Imprimir un mensaje de confirmación
    print(f"Archivo procesado: {excel_file}. Resultados guardados en: {result_file_path}")
