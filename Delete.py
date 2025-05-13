import docx
from docx.oxml.ns import qn # qn is correct for qualified names of elements/tags
from docx.oxml.shared import OxmlElement
# Removed: from docx.oxml.ns import nsattrs # This caused the ImportError

import os
import re

# Define the XML namespace URI for the 'xml' prefix explicitly
XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"


def configurar_directorio_trabajo():
    """Configura el directorio de trabajo en la subcarpeta 'Files'."""
    cwd = os.getcwd()
    target_dir_name = "Files"
    wd = os.path.join(cwd, target_dir_name)
    # Handle potential double 'Files' if script is run from 'Files' itself
    pattern = re.escape(os.sep + target_dir_name) + re.escape(os.sep + target_dir_name)
    if re.search(pattern, wd):
        wd = wd.replace(os.sep + target_dir_name + os.sep + target_dir_name, os.sep + target_dir_name)

    if os.path.isdir(wd):
        os.chdir(wd)
        print(f"Directorio de trabajo cambiado a: {wd}")
    else:
        # Attempt to create the directory if it doesn't exist and is expected
        try:
            os.makedirs(wd)
            os.chdir(wd)
            print(f"Directorio '{wd}' creado y cambiado a él.")
        except Exception as e:
             print(f"Advertencia: El directorio '{wd}' no existe y no pudo ser creado. Error: {e}. No se cambió el directorio de trabajo.")


def obtener_marcadores(documento):
    """Obtiene todos los marcadores de un documento Word."""
    marcadores = {}
    # Find all bookmarkStart elements
    for elemento in documento._element.xpath('//w:bookmarkStart'):
        nombre_marcador = elemento.get(qn('w:name'))
        id_marcador = elemento.get(qn('w:id'))
        if nombre_marcador:
            marcadores[nombre_marcador] = {
                'id': id_marcador,
                'elemento': elemento
            }
    return marcadores


def leer_texto_marcador(documento, nombre_marcador):
    """Lee el texto contenido en un marcador."""
    marcadores = obtener_marcadores(documento)
    if nombre_marcador not in marcadores:
        # print(f"Marcador '{nombre_marcador}' no encontrado para lectura.") # Optional: uncomment for detailed debug
        return ""

    inicio = marcadores[nombre_marcador]['elemento']
    id_marcador = marcadores[nombre_marcador]['id']

    # Find the corresponding bookmarkEnd element
    fin_elementos = documento._element.xpath(f'//w:bookmarkEnd[@w:id="{id_marcador}"]')
    if not fin_elementos:
        # print(f"Marcador '{nombre_marcador}' incompleto (falta el final) para lectura.") # Optional: uncomment
        return ""
    fin = fin_elementos[0]

    texto = []
    actual = inicio
    # Traverse siblings from start to end (exclusive of end)
    while actual is not fin:
        actual = actual.getnext()
        if actual is None:
            # Reached end of parent element before finding bookmarkEnd - should not happen if fin is found
            # print(f"Reached end of parent element before finding end for '{nombre_marcador}'") # Optional: uncomment
            break
        # Find text elements within runs within this range
        # Check if the element itself is a text element (w:t)
        if actual.tag.endswith('t'):
             texto.append(actual.text or "")
        # Or if it's a run element (w:r) potentially containing text elements
        elif actual.tag.endswith('r'):
             for child in actual:
                 if child.tag.endswith('t'):
                     texto.append(child.text or "")
        # Add checks for paragraphs (w:p) containing runs/text if needed, but this sibling traversal
        # between bookmarkStart and bookmarkEnd should cover many common cases.

    return "".join(texto)


def modificar_texto_marcador(documento, nombre_marcador, texto_a_concatenar):
    """
    Concatenates text_a_concatenar to the existing text within a bookmark.
    It reads existing text, combines it, removes old content, and inserts the new combined text.
    """
    marcadores = obtener_marcadores(documento)
    if nombre_marcador not in marcadores:
        print(f"No se encontró el marcador '{nombre_marcador}'. No se puede modificar.")
        return False

    elemento_inicio = marcadores[nombre_marcador]['elemento']
    id_marcador = marcadores[nombre_marcador]['id']
    elemento_fin = None
    # Find the bookmarkEnd element
    for elem in documento._element.xpath(f'//w:bookmarkEnd[@w:id="{id_marcador}"]'):
        elemento_fin = elem
        break

    if elemento_fin is None:
        print(f"Marcador '{nombre_marcador}' incompleto (falta el final)")
        return False

    print(f"Procesando marcador: '{nombre_marcador}'")

    # 1. Read the existing text within the bookmark
    texto_existente = leer_texto_marcador(documento, nombre_marcador)
    print(f"  - Texto existente: '{texto_existente}'")

    # 2. Combine the existing text with the new text
    texto_combinado = texto_a_concatenar + " " + texto_existente
    print(f"  - Texto a concatenar: '{texto_a_concatenar}'")
    print(f"  - Texto combinado: '{texto_combinado}'")

    # 3. Remove the old content between start and end
    # We need to find the elements *between* the start and end and remove them.
    # Collect run elements (w:r) or paragraph elements (w:p) within the range.
    current_elem_del = elemento_inicio
    elements_to_remove = []
    while current_elem_del is not elemento_fin:
         current_elem_del = current_elem_del.getnext()
         if current_elem_del is None:
             # Should not happen if fin is found and is a sibling
             break
         if current_elem_del is not elemento_fin: # Ensure we don't add the end marker itself
             # Collect elements to remove. Focusing on runs and paragraphs that are direct siblings
             # between the start and end bookmark tags.
             if current_elem_del.tag.endswith('r') or current_elem_del.tag.endswith('p'):
                 elements_to_remove.append(current_elem_del)
             # Note: This might not handle complex nested structures perfectly.
             # For simple cases, removing runs between markers usually works.


    # Remove the collected elements.
    for elem in elements_to_remove:
        parent = elem.getparent()
        if parent is not None:
            try:
                parent.remove(elem)
                # print(f"  - Removed element: {elem.tag}") # Optional: uncomment
            except Exception as e:
                # This warning might occur if an element was already removed as part of a parent deletion,
                # or due to complex document structure. It's often non-critical for simple cases.
                print(f"  - Warning: Could not remove element {elem.tag} (already removed or structural issue?): {e}")


    # 4. Insert the new combined text
    # Create a new run element to hold the combined text.
    nuevo_texto_elem = OxmlElement('w:r')
    texto_elem = OxmlElement('w:t')

    # Handle potential whitespace issues in Word by adding xml:space="preserve"
    # if the text contains leading/trailing spaces or multiple internal spaces.
    # Use the correct qualified attribute name {namespace_uri}attribute_name
    if texto_combinado and (texto_combinado.startswith(' ') or texto_combinado.endswith(' ') or '  ' in texto_combinado):
         texto_elem.set(f'{{{XML_NAMESPACE}}}space', 'preserve') # Corrected line

    texto_elem.text = texto_combinado
    nuevo_texto_elem.append(texto_elem)

    # Insert the new run right after the bookmarkStart element.
    # This places the new content at the position of the bookmark.
    try:
        # Find the parent of the bookmarkStart element
        parent_of_start = elemento_inicio.getparent()
        if parent_of_start is not None:
             # Insert the new element before the bookmarkEnd element.
             # Inserting after the start is also valid, depends slightly on desired structure.
             # Inserting before the end ensures it's definitely inside the bookmark range logically.
             parent_of_start.insert(parent_of_start.index(elemento_fin), nuevo_texto_elem)
             print(f"  - Texto combinado inserted correctly.")
             return True
        else:
             print(f"  - Error: Could not find parent element for '{nombre_marcador}' start tag.")
             return False

    except Exception as e:
        print(f"  - Error al insertar texto en '{nombre_marcador}': {e}")
        return False


def main():
    configurar_directorio_trabajo()
    # Use appropriate file names and paths for your setup
    doc_path = "contrato_automatizado.docx" # Your template document with bookmarks
    output_path = "contrato_automatizado_concatenado.docx" # Output document name
    print(f"Cargando documento: {doc_path}")

    try:
        doc = docx.Document(doc_path)
        print("Documento cargado exitosamente.")
    except docx.exceptions.PackageNotFoundError:
         print(f"Error: El archivo '{doc_path}' no se encontró en el directorio de trabajo '{os.getcwd()}'.")
         return
    except Exception as e:
        print(f"Error al cargar el documento '{doc_path}': {e}")
        return

    # Optional: Print available bookmarks
    print("Marcadores disponibles en el documento:")
    marcadores = obtener_marcadores(doc)
    if marcadores:
        for nombre in marcadores:
            print(f"- {nombre}")
    else:
         print("No se encontraron marcadores.")
    print("-" * 20)

    elements_to_add = {"Tercero_DocumentosIntegrantes": "Tercero", "Cuarto_ModificacionDelContrato": "Cuarto",
                "Quinto_GastoseImpuestos": "Quinto", "Sexto_EfectosDerivadosDeIncumplimientos": "Sexto",
                "Septimo_DeLaGarantíaFielCumplimiento": "Septimo", "Octavo_CobroDeLaGarantiaFielCumplimiento": "Octavo",
                "Noveno_TerminoAnticipadoDelContrato": "Noveno", "Decimo_ResciliacionMutuoAcuerdo": "Decimo",
                "DecimoPrimero_ProcedimientoIncumplimient": "Decimo Primero", "DecimoSegundo_EmisionOC": "Decimo Segundo",
                "DecimoTercero_DelPago": "DecimoTercero", "Decimo Cuarto_VigenciaContrato": "Decimo Cuarto",
                "DecimoQuinto_AdministradorContrato": "Decimo Quinto", "DecimoSexto_PactoDeIntegrida": "Decimo Sexto",
                "DecimoSeptimo_ComportamientoEticoAdjudic": "Decimo Septimo", "DecimoOctavo_Auditorias": "Decimo Octavo",
                "DecimoNoveno_Confidencialidad": "Decimo Noveno", "Vigesimo_PropiedadDeLaInformacion": "Vigesimo",
                "VigesimoPrimero_SaldosInsolutos": "Vigesimo Primero",
                "VigesimoSegundo_NormasLaboralesAplicable": "Vigesimo Segundo",
                "VigesimoTercero_CambioPersonalProveedor": "Vigesimo Tercero",
                "VigesimoCuarto_CesionySubcontratacion": "Vigesimo Cuarto",
                "VigesimoQuinto_Discrepancias": "Vigesimo Quinto",
                "VigesimoSexto_Constancia": "Vigesimo Sexto"}

    print("\nIniciando concatenación de texto en marcadores especificados...")
    modificados = 0
    # Loop through the dictionary and concatenate the value to the bookmark's existing text
    for nombre_marcador, texto_a_concatenar in elements_to_add.items():
        # Only attempt modification if the bookmark exists
        if nombre_marcador in obtener_marcadores(doc): # Check existence before calling modify
            if modificar_texto_marcador(doc, nombre_marcador, texto_a_concatenar):
                modificados += 1
        else:
            print(f"Advertencia: El marcador '{nombre_marcador}' especificado en el diccionario no se encontró en el documento.")


    if modificados:
        try:
            doc.save(output_path)
            print(f"\nDocumento guardado como {output_path} con {modificados} marcador(es) modificado(s).")
        except Exception as e:
            print(f"\nError al guardar documento '{output_path}': {e}")
    else:
        print("\nNo se realizaron modificaciones en los marcadores especificados (o no se encontraron).")

if __name__ == "__main__":
    main()