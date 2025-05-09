import os
import re
import docx
from docx.shared import Inches, Cm, Pt  # Import Cm and Pt for units
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT # Using the correct import for your environment

def configurar_directorio_trabajo():
    """Configura el directorio de trabajo en la subcarpeta 'Files'."""
    cwd = os.getcwd()
    target_dir_name = "Files"
    wd = os.path.join(cwd, target_dir_name)

    if os.path.isdir(wd):
        os.chdir(wd)
        print(f"Directorio de trabajo cambiado a: {wd}")
    else:
        print(f"Advertencia: El directorio '{wd}' no existe. No se cambió el directorio de trabajo.")

# Call the function to set the working directory
configurar_directorio_trabajo()

# Define image paths
logo_melipilla_name = "logo_melipilla.png" # This image will be on the right
ssmo_alta_name = "SSMOalta.png"         # This image will be on the left

# Create a new Document
doc = docx.Document()

# --- Add images side-by-side using a table ---

# Add a table with 1 row and 3 columns for the logos (left, space, right)
table = doc.add_table(rows=1, cols=3)
table.autofit = False # Disable autofit to control column widths manually

# Set column widths (adjust as needed to control spacing and placement)
table.columns[0].width = Inches(2.5) # Width for the left logo cell (SSMOalta)
table.columns[1].width = Inches(3)   # Width for spacing
table.columns[2].width = Inches(2.5) # Width for the right logo cell (logo_melipilla)


left_logo_cell = table.cell(0, 0)   # This cell holds the left logo (SSMOalta)
right_logo_cell = table.cell(0, 2)  # This cell holds the right logo (logo_melipilla)

# Add the *SSMOalta* logo to the first (left) cell with specified height
try:
    # Get the first paragraph in the cell to add the picture
    left_logo_paragraph = left_logo_cell.paragraphs[0]
    run = left_logo_paragraph.add_run()
    # Use specified height for the SSMOalta image, width will scale proportionally
    run.add_picture(ssmo_alta_name, height=Cm(2.87))
    left_logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT # Align content left within the cell

    # --- Reduce space AFTER the paragraph inside THIS cell ---
    left_logo_paragraph_format = left_logo_paragraph.paragraph_format
    left_logo_paragraph_format.space_after = Pt(0) # Set space after to 0 points

    left_logo_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP # Align content to the top of the cell
    print(f"Imagen '{ssmo_alta_name}' added to table (left) with height {Cm(2.87)}cm.")
except Exception as e:
    print(f"Error adding picture '{ssmo_alta_name}' to table: {e}")

# Add the *Melipilla* logo to the third (right) cell with specified width
try:
    # Get the first paragraph in the cell to add the picture
    right_logo_paragraph = right_logo_cell.paragraphs[0]
    run = right_logo_paragraph.add_run()
    # Use specified width for the Melipilla image, height will scale proportionally
    run.add_picture(logo_melipilla_name, width=Cm(3.74))
    right_logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT # Align content right within the cell

    # --- Reduce space AFTER the paragraph inside THIS cell ---
    right_logo_paragraph_format = right_logo_paragraph.paragraph_format
    right_logo_paragraph_format.space_after = Pt(0) # Set space after to 0 points


    right_logo_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP # Align content to the top of the cell
    print(f"Imagen '{logo_melipilla_name}' added to table (right) with width {Cm(3.74)}cm.")
except Exception as e:
     print(f"Error adding picture '{logo_melipilla_name}' to table: {e}")

# --- Add text below the table with reduced spacing ---
archivo = "base"

if archivo == "contrato":
    text_lines = ("S.S.M.Occ.",
                  "HOSPITAL DE MELIPILLA",
                  "UNIDAD DE ABASTECIMIENTO",
                  "CONVENIOS",
                  "N° 4",
                  "CRE/RMG/MMJ/MGL/MES")
else:
    text_lines = [
        "SERVICIO SALUD OCCIDENTE",
        "HOSPITAL SAN JOSE DE MELIPILLA",
        "UNIDAD DE ABASTECIMIENTO",
        "BASE N°140" # This line will be bold
    ]

# Add each line as a separate paragraph and apply formatting
for line in text_lines:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT # Ensure left alignment

    # --- Ensure space before and after is zero for these paragraphs ---
    # This makes the lines within this block closer together
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)
    # You could also potentially set line_spacing_rule if needed for more precise control
    # from docx.enum.text import WD_LINE_SPACING
    # paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


    # Set font size to 8 points for the text in this paragraph
    if paragraph.runs: # Check if there's a run (which add_paragraph usually creates)
        font = paragraph.runs[0].font
        font.size = Pt(8) # Set size to 8 points

    # Make the "BASE N°140" line bold
    if line == "BASE N°140":
        if paragraph.runs:
            font.bold = True # Set bold to True


# Add space after this text block (optional, using paragraphs with default spacing)
# These paragraphs added *after* the loop will have default spacing
# --- Save the document ---
output_filename = "portada_melipilla.docx"
try:
    doc.save(output_filename)
    print(f"Document saved as '{output_filename}' with reduced spacing between table and text.")
except Exception as e:
    print(f"Error saving document: {e}")