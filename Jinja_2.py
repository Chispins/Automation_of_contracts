import jinja2
import pandas as pd
from docxtpl import DocxTemplate
# Assuming configurar_directorio_trabajo is correctly defined
from Formated_Base_PEP8 import configurar_directorio_trabajo

# Set up the working directory
configurar_directorio_trabajo()

excel_name = "Libro1.xlsx"

template_name_1 = "base_automatizada.docx"
template_name_2 = "contrato_automatizado_tablas.docx" # Assuming this is the second template

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

contrato_p1_data_dict["espacio"] = " "

# --- Prepare Contexts for Templates (Merged Dictionaries) ---

# Context for the first template (base_automatizada.docx)
# Merge all necessary data into a single dictionary
context_for_template1 = {}
context_for_template1.update(base_data_dict)
context_for_template1.update(contrato_p1_data_dict)
# Add data from P2 here if needed in the first template:
# context_for_template1.update(contrato_p2_data_dict) # Uncomment if P2 data is needed in template 1

# Context for the second template (contrato_automatizado_tablas.docx)
# Merge necessary data into a single dictionary for template 2
context_for_template2 = {}
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
if template_name_2:
    try:
        doc2 = DocxTemplate(template_name_2)
        # Call render() only ONCE with the combined context
        doc2.render(context_for_template2)
        output_filename_2 = "contrato_automatizado_tablas_rendered.docx" # Clearer filename
        doc2.save(output_filename_2)
        print(f"Successfully rendered and saved '{output_filename_2}'")
    except FileNotFoundError:
         print(f"Error: Template file '{template_name_2}' not found.")
    except Exception as e:
        print(f"An error occurred during rendering of {template_name_2}: {e}")