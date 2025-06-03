from openai import OpenAI
import os
from PIL import Image
import base64
import io
import docx
import pymupdf
import re


def configurar_directorio_trabajo():
    cwd = os.getcwd()
    target_dir_name = "Files"
    wd = os.path.join(cwd, target_dir_name)
    pattern = r"Files\\Files"

    if re.search(pattern, wd):
        wd = wd.replace(r"\Files\Files", r"\Files")
    if os.path.isdir(wd):
        os.chdir(wd)
        print(f"Directorio de trabajo cambiado a: {wd}")
    else:
        print(f"Advertencia: El directorio '{wd}' no existe o no es válido.")
configurar_directorio_trabajo()


openai_api_key = "AIzaSyAgXYnDNJ6gQmMVUk88G6wGPdXz-qG2Gbw"
client = OpenAI(
    api_key=openai_api_key,
    base_url="https://generativelanguage.googleapis.com/v1beta/openai/"
)

model = "gemini-2.5-flash-preview-04-17"

# Cargar y convertir la imagen a base64

image_name = "CERTIFICADO DE FIEL CUMPLIMIENTO ID 1057480-11-LE25. ELITEC SPA.pdf"



regex = r".*(.pdf)"
formato = re.match(regex, image_name)


# Si es PDF
if formato != None:
    try:
        pdf_document = pymupdf.open(image_name)
        primera_pagina = pdf_document.load_page(0)
        pix = primera_pagina.get_pixmap(alpha=False)
        img_bytes = pix.tobytes("jpeg")
        image_base64 = base64.b64encode(img_bytes).decode('utf-8')
        pdf_document.close()
    except Exception as e:
        print(f"Error al procesar el PDF: {e}")
        exit(1)
# Si es Imagen
else:
    image = Image.open(image_name)
# Convertir la imagen a bytes y luego a base64
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG")
    image_bytes = buffer.getvalue()
    image_base64 = base64.b64encode(image_bytes).decode('utf-8')

# Crear el mensaje con la imagen
response = client.chat.completions.create(
    model=model,
    messages=[
        {  "role": "system",
    "content": "You are a transcriber in a company specializing in standardizing documents for Garntía de Fiel Cumplimiento from various institutions. Your task is to extract and interpret key values into a single, consistent format. Note that fields may vary; for example, use 'Tomador' if not explicit, but interpret 'Afianzado' as equivalent. Similarly, interpret 'Acreedor' as 'Beneficiario' if needed. Output the results in a dictionary format based on the provided example, asegurado suele ser el mismo beneficiado el rut del tomador no puede ser el mismo que el rut del beneficiado, tipicamente 61.602.123-0 el rut del beneficiado será . The files are in Spanish and you must work in spanish"},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": """
Your role is to read the file an complete the fields an generate a dictionary, follow the same formating as this example, use the keys from the example, and the values for those keys should be the corresponding data found in the image. If not found, put NULL. If numero Dias is not found it gotta be calculated as the difference of days betwen the second value of vigencia del seguro and the first one, if "Valor a pagar del Contrato" is not on the file, it should be written in letters takin as a base the ammuont in numbers. Vigencia del seguro is since the document is emited until it expires, and Numero de dias is the difference betwen those two elements  
{
    "Tomador": "Servicios de Bioingenieria Limitada",
    "rut_tomador": "76.644.150-5",
    "Asegurado": "Hospital San José de Melipilla",
    "rut_asegurado": "66.644.150-5",
    "Beneficiario": "Hospital San José de Melipilla",
    "rut_beneficiario": "66.644.150-5",
    "Direccion_Tomador": "Avenida Condell 1680",
    "Ciudad": "Nuñoa",
    "Cobertura": "Fiel Cumplimiento",
    "Vigencia_del_seguro": "01/01/2025–02/01-2025",
    "numero_de_dias": "1",
    "Valor_asegurado": "699,50",
    "Prima_neta": "1633,92",
    "IVA": "3,21",
    "total_a_pagar": "20,13",
    "valor_a_pagar_en_letra": "Veinte coma trece UF"
    "ciudad_y_fecha_de_emision": "Santiago, 18 de marzo de 2025",
    "Poliza_N_ID: 17306798"
}

                """},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_base64}"}}
            ]
        }
    ]
)


print(response.choices[0].message.content)


# Obtener la respuesta del modelo Gemini
respuesta = response.choices[0].message.content


import json
respuesta_str = response.choices[0].message.content

# Eliminar marcas ``json y ```
if respuesta_str.startswith("```json"):
    respuesta_str = respuesta_str.split("\n", 1)[1]
if respuesta_str.endswith("```"):
    respuesta_str = respuesta_str.rsplit("\n", 1)[0]

datos = json.loads(respuesta_str)

# Now im importing my Jinja2template

name = "prototipo_tabla_JINJA2.docx"
from docxtpl import DocxTemplate

doce = DocxTemplate(name)
doce.render(datos)
doce.save("test_table.docx")

