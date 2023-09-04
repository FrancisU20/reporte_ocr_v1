import os
import json
import base64
from PIL import Image
import pandas as pd

# Ruta de la carpeta que contiene los archivos JSON
json_folder = "MuestraJson"

# Ruta del archivo Excel
excel_file = "MuestraOCR.xlsx"

# Crear la carpeta 'imagenes' si no existe
if not os.path.exists("imagenes"):
    os.mkdir("imagenes")

# Cargar el archivo Excel en un DataFrame
df = pd.read_excel(excel_file)

# Función para decodificar y guardar la imagen
def save_image(base64_data, image_path):
    img_data = base64.b64decode(base64_data)
    with open(image_path, "wb") as f:
        f.write(img_data)

# Iterar a través de los archivos JSON en la carpeta
for json_file in os.listdir(json_folder):
    if json_file.endswith(".json"):
        with open(os.path.join(json_folder, json_file), "r") as f:
            json_data = json.load(f)
            codigo_documental = json_data["CodigoDocumental"]
            base64_image = json_data["Base64"]

            # Decodificar y guardar la imagen en la carpeta 'imagenes'
            image_path = os.path.join("imagenes", f"{codigo_documental}.jpg")
            save_image(base64_image, image_path)

            # Crear el hipervínculo en la columna 'IMAGEN'
            hyperlink = f'=HYPERLINK("{image_path}", "Abrir imagen")'
            df.loc[df['CODIGO DOCUMENTAL'] == codigo_documental, 'IMAGEN'] = hyperlink

# Guardar el DataFrame actualizado en el archivo Excel
df.to_excel(excel_file, index=False)
