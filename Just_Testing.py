import docx
import os

wd = r"C:\Users\serco\Downloads"
os.chdir(wd)

doc = docx.Document()

# Agregar título
doc.add_heading("Ejemplos de Estilos de Lista", 0)

# List
doc.add_paragraph("Este es un ejemplo de estilo 'List'", style="List")

# List 2
doc.add_paragraph("Este es un ejemplo de estilo 'List 2'", style="List 2")

# List 3
doc.add_paragraph("Este es un ejemplo de estilo 'List 3'", style="List 3")

# List Bullet
doc.add_paragraph("Este es un ejemplo de estilo 'List Bullet'", style="List Bullet")

# List Bullet 2
doc.add_paragraph("Este es un ejemplo de estilo 'List Bullet 2'", style="List Bullet 2")

# List Bullet 3
doc.add_paragraph("Este es un ejemplo de estilo 'List Bullet 3'", style="List Bullet 3")

# List Continue
doc.add_paragraph("Este es un ejemplo de estilo 'List Continue'", style="List Continue")

# List Continue 2
doc.add_paragraph("Este es un ejemplo de estilo 'List Continue 2'", style="List Continue 2")

# List Continue 3
doc.add_paragraph("Este es un ejemplo de estilo 'List Continue 3'", style="List Continue 3")

# List Number
doc.add_paragraph("Este es un ejemplo de estilo 'List Number'", style="List Number")

# List Number 2
doc.add_paragraph("Este es un ejemplo de estilo 'List Number 2'", style="List Number 2")

# List Number 3
doc.add_paragraph("Este es un ejemplo de estilo 'List Number 3'", style="List Number 3")

# List Paragraph
doc.add_paragraph("Este es un ejemplo de estilo 'List Paragraph'", style="List Paragraph")

# Guardar documento
doc.save("ejemplos_estilos_lista.docx")