import jinja2
from docxtpl import DocxTemplate
import os
from Formated_Base_PEP8 import configurar_directorio_trabajo

configurar_directorio_trabajo()

# Define paths
template_path = "contrato_automatizado.docx"
output_path = "contrato_automatizado_jinja2.docx"

Datos_Contrato = {
    "numero_contrato" : "4",
    "involucrados" : "CRE/RMG/MMJ/MGL/MES",
    "fecha_contrato" : "02 de enero de 2025",
    "nombre_proveedor" : "MEDCORP S.A",
    "rut_proveedor": "76.131.542-0",
    "representante_legal" : "doña Alejandra Ana Cuesta Nazar",
    "rut_representante_legal":"15.638.432-1",
    "domicilio_representante_legal" : "cedula nacional de identidad N° 15.638.432-1",
    "id_licitacion" : "1057480-81-LE24",
    "numero_resolucion_aprobacion" : "1057480-81-LE24",
    "fecha_resolucion_aprobacion" : "05 de diciembre de 2024",
    "numero_resolucion" : "000596",
    "fecha_resolucion" : "06 de enero de 2024"


}


doc = DocxTemplate(template_path)

# 2. Render the template with the context data
doc.render(Datos_Contrato)

# 3. Save the generated document
doc.save(output_path)

print(f"Report '{output_path}' generated successfully!")



