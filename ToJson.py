import os
import pandas as pd
import json

# Ruta de la carpeta donde están los archivos de resultados
results_folder = "results"
# Ruta de la carpeta donde se guardarán los archivos JSON
json_folder = "JSON"

# Crear la carpeta de JSON si no existe
if not os.path.exists(json_folder):
    os.makedirs(json_folder)

# Obtener la lista de archivos en la carpeta de resultados
result_files = [file for file in os.listdir(results_folder) if file.endswith(".xlsx")]

sequence = 1

# Iterar a través de los archivos de resultados
for result_file in result_files:
    # Leer el archivo Excel
    df = pd.read_excel(os.path.join(results_folder, result_file), sheet_name="Sheet1")

    # Rellenar NaN con strings vacíos en todo el DataFrame
    df.fillna("", inplace=True)

    # Crear la estructura de datos en formato JSON
    data = {
        "header": [],
        "sales": [],
        "stock": []
    }

    # Llenar la estructura con los datos del DataFrame
    for index, row in df.iterrows():
        if index == 0:
            header = {
                "indicadorDeTiempo": f"{row['Indicador de tiempo']}",
                "codigoKA/OGK": f"{row['Codigo KA/OGK']}",
                "codigoDePeriodicidad": f"{row['Código de periodicidad']}",
                "sequence": f"{sequence}"
            }
            data["header"].append(header)

        sale = {
            "fechaDelDocumento": f"{row['Fecha del documento']}",
            "registroDeTiempo": f"{row['Registro de tiempo']}",
            "codigoKA/OGK": f"{row['Codigo KA/OGK']}",
            "codigoDelPDV": f"{row['Codigo del PDV']}",
            "razonSocial": f"{row['Razón Social']}",
            "calle": f"{row['Calle']}",
            "numero": f"{row['Numero']}",
            "localidad": f"{row['Localidad']}",
            "codigoEAN": f"{row['Código EAN']}",
            "EANDescripcion": f"{row['EAN Descripción']}",
            "nroDeFactura": f"{row['Nro de factura']}",
            "cantidadDePaquetes": f"{row['Cantidad de paquetes']}",
            "totalPacksAmount": '1200.54'
        }
        data["sales"].append(sale)

        if index == 0:
            stock = {
                "fechaStock": '1',
                "registroDeTiempo": '1',
                "codigoKA/OGK": '1',
                "codigoDelPDV": '1',
                "razonSocial": '1',
                "calle": '1',
                "numero": '1',
                "localidad": '1',
                "codigoEAN": '1',
                "EANDescripcion": '1',
                "cantidadDePaquetes": '1'
            }
            data["stock"].append(stock)


    # Guardar la estructura como un archivo JSON en la carpeta JSON
    json_file_path = os.path.join(json_folder, result_file.replace(".xlsx", ".json"))
    with open(json_file_path, "w") as json_file:
        json.dump(data, json_file, indent=4)

    print(f"Archivo JSON guardado: {json_file_path}")

    sequence = sequence + 1
