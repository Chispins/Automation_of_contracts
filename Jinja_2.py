import jinja2
import pandas as pd
from docxtpl import DocxTemplate
import os
import re
# Assuming configurar_directorio_trabajo is correctly defined

def configurar_directorio_trabajo():
    """Configura el directorio de trabajo en la subcarpeta 'Files'."""
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
# Set up the working directory
configurar_directorio_trabajo()

excel_name = "Libro1.xlsx"

template_name_1 = "base_automatizada.docx"
template_name_2 = "base_automatizada.docx" # Assuming this is the second template
template_name_3 = "contrato_automatizado_tablas.docx"

# Function to strip whitespace and handle potential NaN/empty values
def strip_dataframe_and_handle_empty(df):
    # Replace NaN with empty string before stripping
    df = df.fillna('')
    # Convert all columns to string type before stripping to avoid errors on non-string data
    df = df.astype(str)
    return df.applymap(lambda x: x.strip())

# Read and clean each sheet
try:
    # Ensure column names are treated as data if the first row is actually data, not headers
    Datos_Base_excel = strip_dataframe_and_handle_empty(pd.read_excel(excel_name, sheet_name="Datos_Base", header=None))
    Datos_Contrato_P1 = strip_dataframe_and_handle_empty(pd.read_excel(excel_name, sheet_name="Datos_Contrato_P1", header=None))
    Datos_Contrato_P2 = strip_dataframe_and_handle_empty(pd.read_excel(excel_name, sheet_name="Datos_Contrato_P2", header=None))
    print("Excel sheets read and cleaned successfully.")
    # print(Datos_Base_excel.head()) # Optional: check the cleaned data
except FileNotFoundError:
    print(f"Error: Excel file '{excel_name}' not found.")
    exit()
except KeyError as e:
    print(f"Error: Sheet '{e}' not found in '{excel_name}'. Check sheet names.")
    exit()

except Exception as e:
    print(f"An error occurred while reading Excel: {e}")
    exit()


# --- Process DataFrames into Dictionaries and Clean Values ---
# Assuming the first column (index 0) is the key, second (index 1) is the value.

def process_and_clean_dict(df, sheet_name):
    """
    Reads data from a DataFrame (assumed key-value columns),
    creates a dictionary, cleans keys (removes colons, strips),
    and cleans values (removes trailing commas, ensures string).
    """
    try:
        keys = df.iloc[:, 0].tolist()
        values = df.iloc[:, 1].tolist()
        raw_dict = dict(zip(keys, values))
        print(f"Raw dictionary from {sheet_name}: {list(raw_dict.items())[:5]}...")

        # Clean keys and values
        cleaned_dict = {}
        for key, value in raw_dict.items():
            # Clean key: remove colons, strip whitespace
            cleaned_key = key.replace(':', '').strip()

            # Clean value:
            cleaned_value = str(value).strip() # Ensure it's a string first
            if cleaned_value.endswith(','):
                cleaned_value = cleaned_value[:-1] # Remove trailing comma

            # --- Add Specific Formatting for known keys here ---
            # Example for 'presupuesto_con_impuestos':
            if cleaned_key == 'presupuesto_con_impuestos':
                 try:
                     num_value = float(cleaned_value.replace('.', '').replace(',', '.')) # Convert to number, handle commas/dots
                     # Format as currency (adjust locale/format as needed)
                     # Using a format string that handles thousands separators and decimal points
                     cleaned_value = f"${num_value:,.0f}".replace(',', '.') # Format with dots as thousands separator
                 except ValueError:
                     # Handle cases where conversion fails
                     pass # Keep original cleaned string if not a valid number

            # Example for 'plazo_meses' or 'dias_vigencia_publicacion' if you want them as strings:
            elif cleaned_key in ['plazo_meses', 'dias_vigencia_publicacion']:
                 cleaned_value = str(cleaned_value) # Ensure it's a string

            # Add more specific formatting rules for other keys as needed

            # Avoid adding empty keys if your Excel has empty cells in the first column
            if cleaned_key:
                cleaned_dict[cleaned_key] = cleaned_value

        print(f"Cleaned dictionary from {sheet_name}: {list(cleaned_dict.items())[:5]}...")
        return cleaned_dict

    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {e}")
        return {}


base_data_dict = process_and_clean_dict(Datos_Base_excel, "Datos_Base")
contrato_p1_data_dict = process_and_clean_dict(Datos_Contrato_P1, "Datos_Contrato_P1")
contrato_p2_data_dict = process_and_clean_dict(Datos_Contrato_P2, "Datos_Contrato_P2")

if base_data_dict["anexo_6"]:
    base_data_dict["anexo_6"] = "Anexo N°6 "+ base_data_dict["anexo_6"]

if base_data_dict["anexo_7"]:
    base_data_dict["anexo_7"] = "Anexo N°7 "+ base_data_dict["anexo_7"]

if base_data_dict["anexo_8"]:
    base_data_dict["anexo_8"] = "Anexo N°8 "+ base_data_dict["anexo_8"]

if base_data_dict["anexo_9"]:
    base_data_dict["anexo_9"] = "Anexo N°9 "+ base_data_dict["anexo_9"]

if (base_data_dict["entrega_muestras"]):
    base_data_dict["muestra"] = "Entregar muestras de los productos solicitados y comodato ofertado."
if (base_data_dict["garantia"]):
    base_data_dict["entreg_garan"] = "Entregar garantías de la oferta"
base_data_dict["labor_coordinador"] = " en el desempeño de su cometido, el coordinador del contrato deberá, a lo menos"
base_data_dict["coordinador"] = " El adjudicatario nombra coordinador del contrato a"


contrato_p1_data_dict["espacio"] = " "
contrato_p1_data_dict["Documentos_Integrantes"] = "Tercero"
contrato_p1_data_dict["Cuarto_ModificacionDelContrato"] = "Cuarto"
contrato_p1_data_dict["Quinto_GastoseImpuestos"] = "Quinto"
contrato_p1_data_dict["Sexto_EfectosDerivadosDeIncumplimiento"] = "Sexto"
contrato_p1_data_dict["Septimo_DeLaGarantíaFielCumplimiento"] = "Séptimo"
contrato_p1_data_dict["Octavo_CobroDeLaGarantiaFielCumplimiento"] = "Octavo"
contrato_p1_data_dict["Noveno_TerminoAnticipadoDelContrato"] = "Noveno"
contrato_p1_data_dict["Decimo_ResciliacionMutuoAcuerdo"] = "Décimo"
contrato_p1_data_dict["DecimoPrimero_ProcedimientoIncumplimiento"] = "Décimo Primero"
contrato_p1_data_dict["DecimoSegundo_EmisionOC"] = "Decimo Segundo"
contrato_p1_data_dict["DecimoTercero_DelPago"] = "Décimo Tercero"
contrato_p1_data_dict["DecimoCuarto_VigenciaContrato"] = "Décimo Cuarto"
contrato_p1_data_dict["DecimoQuinto_AdministradorContrato"] = "Décimo Quinto"
contrato_p1_data_dict["DecimoSexto_PactoDeIntegrida"] = "Décimo Sexto"
contrato_p1_data_dict["DecimoSeptimo_ComportamientoEticoAdjudic"] = "Décimo Séptimo"
contrato_p1_data_dict["DecimoOctavo_Auditorias"] = "Décimo Octavo"
contrato_p1_data_dict["DecimoNoveno_Confidencialidad"] = "Décimo Noveno"
contrato_p1_data_dict["Vigesimo_PropiedadDeLaInformacion"] = "Vigésimo"

contrato_p1_data_dict["VigesimoPrimero_SaldosInsolutos"] = "Vigésimo Primero"
contrato_p1_data_dict["VigesimoSegundo_NormasLaboralesAplicable"] = "Vigésimo Segundo"
contrato_p1_data_dict["VigesimoTercero_CambioPersonalProveedor"] = "Vigésimo Tercero"
contrato_p1_data_dict["VigesimoCuarto_CesionySubcontratacion"] = "Vigésimo Cuarto"
contrato_p1_data_dict["VigesimoQuinto_Discrepancias"] = "Vigésimo Quinto"
contrato_p1_data_dict["VigesimoSexto_Constancia"] = "Vigésimo Sexto"
contrato_p1_data_dict["VigesimoSeptimo_Ejemplares"] = "Vigésimo Séptimo"
contrato_p1_data_dict["VigesimoOctavo_Personeria"] = "Vigésimo Octavo"

contrato_p1_data_dict["coordinador"] = "El adjudicatario nombra coordinador del contrato a"
contrato_p1_data_dict["texto_gar_1"] = ", es decir"
contrato_p1_data_dict["texto_gar_2"] = "de pesos a nombre de “EL HOSPITAL” y consigna la siguiente glosa: Para garantizar el fiel cumplimiento del contrato denominado:"
contrato_p1_data_dict["texto_gar_3"] = "ds"
contrato_p1_data_dict["director"] = " la Resolución Exenta RA 116395/343/2024 de fecha 12/08/2024 del SSMOCC., la cual nombra Director del Hospital San José de Melipilla al suscrito"



# --- Prepare Contexts for Templates (Merged Dictionaries) ---

# Context for the first template (base_automatizada.docx)
# Merge all necessary data into a single dictionary
context_for_template1 = {}
context_for_template1.update(base_data_dict)
#context_for_template1.update(contrato_p1_data_dict)
# Add data from P2 here if needed in the first template:
# context_for_template1.update(contrato_p2_data_dict) # Uncomment if P2 data is needed in template 1

# Context for the second template (contrato_automatizado_tablas.docx)
# Merge necessary data into a single dictionary for template 2
context_for_template2 = {}
context_for_template2.update(base_data_dict)
context_for_template2.update(contrato_p1_data_dict)
context_for_template2.update(contrato_p2_data_dict)
# Add data from Base or P1 here if needed in the second template:
# context_for_template2.update(base_data_dict) # Uncomment if Base data is needed in template 2
# context_for_template2.update(contrato_p1_data_dict) # Uncomment if P1 data is needed in template 2


# --- Render and Save Documents ---

# Render and save the first template
try:
    doc1 = DocxTemplate(template_name_1)
    # Call render() only ONCE with the combined context
    doc1.render(context_for_template1)
    output_filename_1 = "base_automatizada_jinja2.docx" # Clearer filename
    doc1.save(output_filename_1)
    print(f"Successfully rendered and saved '{output_filename_1}'")

except FileNotFoundError:
     print(f"Error: Template file '{template_name_1}' not found.")
except Exception as e:
    print(f"An error occurred during rendering of {template_name_1}: {e}")


# Render and save the second template (if you need to generate a second document)
# Check if template_name_2 is defined and exists
"""if template_name_2:
    try:
        doc2 = DocxTemplate(template_name_2)
        # Call render() only ONCE with the combined context
        doc2.render(context_for_template2)
        output_filename_2 = "contrato_automatizado_rendered.docx" # Clearer filename
        doc2.save(output_filename_2)
        print(f"Successfully rendered and saved '{output_filename_2}'")
    except FileNotFoundError:
         print(f"Error: Template file '{template_name_2}' not found.")
    except Exception as e:
        print(f"An error occurred during rendering of {template_name_2}: {e}")
"""
try:
    doc3 = DocxTemplate(template_name_3)
    doc3.render(context_for_template2)
    output_filename_3 = "contrato_automatizado_rendered.docx"
    doc3.save(output_filename_3)
    print(f"Successfully rendered and saved '{output_filename_3}'")
except FileNotFoundError:
    print(f"Error: Template file '{template_name_3}' not found.")
except Exception as e:
    print(f"An error occurred during rendering of {template_name_3}: {e}")
