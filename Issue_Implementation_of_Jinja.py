from Formated_Base_PEP8 import configurar_directorio_trabajo  # Your existing import
import docx
from docx.oxml.ns import qn  # For qualified XML tag names (e.g., w:bookmarkStart)
import os  # For path operations


# --- Function to extract bookmark names and their enclosed text ---
def get_bookmark_text_data(doc_object):
    """
    Extracts bookmark names and their enclosed text from a docx.Document object.

    This function iterates through paragraphs and their underlying XML to find
    bookmark start and end tags, and collects the text runs between them.

    Args:
        doc_object (docx.Document): The python-docx Document object to process.

    Returns:
        dict: A dictionary where keys are bookmark names (str) and
              values are the text (str) enclosed by those bookmarks.
              Example: {'ClienteNombre': 'John Doe', 'FechaContrato': '2023-01-15'}
    """
    bookmark_data = {}  # To store final results: {name: text}

    # Buffer for text being collected for currently "open" (active) bookmarks
    # Key: bookmark_id (str), Value: dict {'name': str, 'text_parts': list[str]}
    active_bookmarks_info = {}

    # doc.paragraphs includes paragraphs from the main document body,
    # including those within table cells.
    for paragraph in doc_object.paragraphs:
        # paragraph._p is the lxml _Element object for the <w:p> tag
        for element in paragraph._p:

            # --- Handle Bookmark Start (<w:bookmarkStart ... />) ---
            if element.tag == qn('w:bookmarkStart'):
                bm_id = element.get(qn('w:id'))  # Get the bookmark ID
                bm_name = element.get(qn('w:name'))  # Get the bookmark name

                # A bookmark starts; prepare to collect its text.
                # We assume bookmarks are well-formed and non-overlapping for simplicity.
                if bm_id not in active_bookmarks_info:
                    active_bookmarks_info[bm_id] = {'name': bm_name, 'text_parts': []}

            # --- Handle Bookmark End (<w:bookmarkEnd ... />) ---
            elif element.tag == qn('w:bookmarkEnd'):
                bm_id = element.get(qn('w:id'))  # Get the ID of the ending bookmark

                if bm_id in active_bookmarks_info:
                    # This bookmark has ended; consolidate its collected text.
                    info = active_bookmarks_info[bm_id]
                    name = info['name']
                    text = "".join(info['text_parts'])  # Join all collected text parts

                    # Store the extracted name and text.
                    # If a bookmark name appears multiple times, this logic prioritizes
                    # the first non-empty text found. Word UI usually enforces unique bookmark names.
                    if name not in bookmark_data or (not bookmark_data.get(name) and text):
                        bookmark_data[name] = text

                    # This bookmark is no longer active for collection.
                    del active_bookmarks_info[bm_id]

            # --- Handle Runs (<w:r> ... </w:r>) where text content resides ---
            elif element.tag == qn('w:r'):
                run_text = ""
                # A run can contain multiple <w:t> (text) elements. Concatenate them.
                for t_element in element.findall(qn('w:t')):
                    if t_element.text:
                        run_text += t_element.text

                if run_text:
                    # If this run contains text, append it to ALL currently active bookmarks.
                    # This handles bookmarks that span across multiple runs or paragraphs.
                    for bm_id_active in active_bookmarks_info:  # Iterate over keys
                        active_bookmarks_info[bm_id_active]['text_parts'].append(run_text)

    # After processing all paragraphs, check for any bookmarks that started but didn't end.
    # This might indicate a malformed document or bookmarks spanning into headers/footers
    # not covered by doc.paragraphs.
    if active_bookmarks_info:
        # print(f"Warning: Unclosed bookmarks detected at end of paragraph processing: {list(active_bookmarks_info.keys())}")
        for bm_id, info in active_bookmarks_info.items():
            name = info['name']
            text = "".join(info['text_parts'])
            # Store if the name hasn't been stored yet, or if it was stored as empty and this has text
            if name not in bookmark_data or (not bookmark_data.get(name) and text):
                # print(f"Warning: Storing text for potentially unclosed bookmark '{name}' (ID: {bm_id})")
                bookmark_data[name] = text

    return bookmark_data


# --- End of get_bookmark_text_data function ---


# --- Main script execution ---
if __name__ == "__main__":
    configurar_directorio_trabajo()  # Your function to set the working directory

    doc_filename_base = "contrato_automatizado_con_marcadores"
    doc_path = doc_filename_base + ".docx"

    # Check if the document exists
    if not os.path.exists(doc_path):
        print(f"Error: El archivo '{doc_path}' no fue encontrado en el directorio de trabajo: {os.getcwd()}")
        # You could create a dummy document here for testing if needed:
        # print("Creando un documento de prueba...")
        # dummy_doc = docx.Document()
        # p = dummy_doc.add_paragraph()
        # # Add bookmarkStart for BM_Prueba
        # bm_start_el = docx.oxml.OxmlElement('w:bookmarkStart')
        # bm_start_el.set(docx.oxml.ns.qn('w:name'), 'BM_Prueba')
        # bm_start_el.set(docx.oxml.ns.qn('w:id'), '0')
        # p._p.append(bm_start_el)
        # # Add a run with text
        # p.add_run("Este es el texto para BM_Prueba.")
        # # Add bookmarkEnd for BM_Prueba
        # bm_end_el = docx.oxml.OxmlElement('w:bookmarkEnd')
        # bm_end_el.set(docx.oxml.ns.qn('w:id'), '0')
        # p._p.append(bm_end_el)
        # dummy_doc.save(doc_path)
        # print(f"Documento de prueba '{doc_path}' creado con un marcador 'BM_Prueba'.")
        exit()

    try:
        doc = docx.Document(doc_path)
    except Exception as e:
        print(f"Error al abrir el documento '{doc_path}': {e}")
        exit()

    # Call the function to extract all bookmark names and their texts
    # This replaces your previous `marcadores = obtener_marcadores(doc)` if that
    # function didn't directly return the text. If it did, you can use it instead.
    all_bookmarks_data = get_bookmark_text_data(doc)

    # Now, 'all_bookmarks_data' is a dictionary.
    # You can iterate through it to access each bookmark's name and text.
    print("\n--- Datos de los Marcadores Extraídos ---")
    if all_bookmarks_data:
        for bookmark_name, bookmark_text in all_bookmarks_data.items():
            # 'bookmark_name' and 'bookmark_text' are the variables for each bookmark's data
            print(f"Nombre del Marcador: {bookmark_name}")
            print(f"Texto Asociado     : '{bookmark_text}'")  # Using '' to show if text is empty
            print("-" * 30)

            # You can now use these variables as needed. For example, store them,
            # print them, or use them in further processing.
            # If you need to assign them to specific, uniquely named Python variables
            # (e.g., cliente_nombre = "John Doe"), you would typically do that
            # by checking the bookmark_name:
            #
            # if bookmark_name == "NombreCliente":
            #     nombre_cliente_variable = bookmark_text
            #     print(f"Variable 'nombre_cliente_variable' asignada con: '{nombre_cliente_variable}'")
            # elif bookmark_name == "FechaDocumento":
            #     fecha_documento_variable = bookmark_text
            #     print(f"Variable 'fecha_documento_variable' asignada con: '{fecha_documento_variable}'")

        # The dictionary 'all_bookmarks_data' itself stores all the data.
        # You can access specific bookmark texts like this:
        # texto_especifico = all_bookmarks_data.get("NombreDelMarcadorBuscado")
        # if texto_especifico is not None:
        #     print(f"\nTexto para 'NombreDelMarcadorBuscado': '{texto_especifico}'")
        # else:
        #     print("\nMarcador 'NombreDelMarcadorBuscado' no encontrado.")

    else:
        print("No se encontraron marcadores en el documento, o no se pudo extraer texto de ellos.")

for nombre in all_bookmarks_data:
    print(f"Nombre del Marcador: {nombre}")

valores = all_bookmarks_data.values()
import pandas as pd
valores = all_bookmarks_data.values()
import pandas as pd
df = pd.DataFrame(
    list(all_bookmarks_data.items()),
    columns=['Marcador', 'Texto']
)

df["Texto_a_ingresar"] = "Primero " + df["Texto"]
# Export to xlsx
df.to_excel('marcadores.xlsx', index=False)

