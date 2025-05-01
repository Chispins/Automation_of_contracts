import docx
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
import os
import re


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
        print(f"Advertencia: El directorio '{wd}' no existe. No se cambió el directorio de trabajo.")


def obtener_marcadores(documento):
    """Obtiene todos los marcadores de un documento Word."""
    marcadores = {}
    for elemento in documento._element.xpath('//w:bookmarkStart'):
        nombre_marcador = elemento.get(qn('w:name'))
        id_marcador = elemento.get(qn('w:id'))
        if nombre_marcador:
            marcadores[nombre_marcador] = {
                'id': id_marcador,
                'elemento': elemento
            }
    return marcadores


def modificar_texto_marcador(documento, nombre_marcador, nuevo_texto):
    """Modifica el texto contenido en un marcador eliminando y reemplazando el contenido."""
    marcadores = obtener_marcadores(documento)
    if nombre_marcador not in marcadores:
        print(f"No se encontró el marcador '{nombre_marcador}'. No se puede modificar.")
        return False

    elemento_inicio = marcadores[nombre_marcador]['elemento']
    id_marcador = marcadores[nombre_marcador]['id']
    elemento_fin = None
    for elem in documento._element.xpath(f'//w:bookmarkEnd[@w:id="{id_marcador}"]'):
        elemento_fin = elem
        break

    if elemento_fin is None:
        print(f"Marcador '{nombre_marcador}' incompleto (falta el final)")
        return False

    current_elem = elemento_inicio
    elementos_a_eliminar = []
    while current_elem is not elemento_fin:
        current_elem = current_elem.getnext()
        if current_elem is None:
            print(f"No se pudo encontrar el contenido entre inicio y fin para '{nombre_marcador}'")
            return False
        if current_elem.tag.endswith('r') or current_elem.tag.endswith('t'):
            elementos_a_eliminar.append(current_elem)

    for elem in elementos_a_eliminar:
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)

    nuevo_texto_elem = OxmlElement('w:r')
    texto_elem = OxmlElement('w:t')
    texto_elem.text = nuevo_texto
    nuevo_texto_elem.append(texto_elem)

    try:
        elemento_inicio.addnext(nuevo_texto_elem)
        print(f"Texto reemplazado correctamente para el marcador '{nombre_marcador}'")
        return True
    except Exception as e:
        print(f"Error al insertar texto en '{nombre_marcador}': {e}")
        return False


def crear_marcador_en_documento(documento, parrafo, nombre_marcador, texto):
    """Crea un marcador en un párrafo con el nombre y texto especificados."""
    run = parrafo.add_run()
    tag = run._r

    # Obtener un ID único para el marcador
    existing_ids = {int(elem.get(qn('w:id'))) for elem in
                    documento._element.xpath('//w:bookmarkStart | //w:bookmarkEnd') if elem.get(qn('w:id'))}
    next_id = max(existing_ids) + 1 if existing_ids else 0

    # Crear elemento de inicio del marcador
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(next_id))
    start.set(qn('w:name'), nombre_marcador)
    tag.append(start)

    # Agregar texto dentro del marcador
    parrafo.add_run(texto)

    # Crear elemento de fin del marcador
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(next_id))
    parrafo._p.append(end)

    return True

# -- Function de prueba, sin usar --
def crear_nuevo_documento_con_marcadores(marcadores_a_modificar, output_path_nuevo):
    """Crea un nuevo documento Word con un título, párrafo y texto con marcadores."""
    nuevo_doc = docx.Document()

    nuevo_doc.add_heading('Documento Nuevo con Marcadores', level=1)
    parrafo = nuevo_doc.add_paragraph('Este es un párrafo introductorio en el nuevo documento creado. ')

    # Agregar "hola " y el primer marcador
    parrafo.add_run('Hola ')
    primer_marcador = list(marcadores_a_modificar.keys())[0] if marcadores_a_modificar else "Marcador1"
    texto_primer_marcador = marcadores_a_modificar.get(primer_marcador, "Texto no encontrado")
    crear_marcador_en_documento(nuevo_doc, parrafo, primer_marcador, texto_primer_marcador)

    # Agregar "encantado de" y el segundo marcador
    parrafo.add_run(' encantado de ')
    segundo_marcador = list(marcadores_a_modificar.keys())[1] if len(marcadores_a_modificar) > 1 else "Marcador2"
    texto_segundo_marcador = marcadores_a_modificar.get(segundo_marcador, "Texto no encontrado")
    crear_marcador_en_documento(nuevo_doc, parrafo, segundo_marcador, texto_segundo_marcador)

    try:
        nuevo_doc.save(output_path_nuevo)
        print(f"\nNuevo documento guardado como '{output_path_nuevo}' con los marcadores insertados.")
    except Exception as e:
        print(f"\nError al guardar el nuevo documento: {e}")


def main():
    configurar_directorio_trabajo()
    doc_path = "contrato_sin_cambios.docx"
    output_path = "contrato_con_cambio_marcadores.docx"
    print(f"Cargando documento: {doc_path}")

    try:
        doc = docx.Document(doc_path)
    except Exception as e:
        print(f"Error al cargar el documento: {e}")
        return

    marcadores_iniciales = obtener_marcadores(doc)
    print("Marcadores disponibles inicialmente:")
    if marcadores_iniciales:
        for nombre in marcadores_iniciales:
            print(f"- {nombre} (ID: {marcadores_iniciales[nombre]['id']})")
    else:
        print("No se encontraron marcadores en el documento.")

    marcadores_a_modificar = {
        "QWERTY": "Ejemplo",
        "asdf": "ahhh"
    }

    print("\nIntentando modificar marcadores...")
    modificados_count = 0
    for nombre_marcador, nuevo_texto in marcadores_a_modificar.items():
        if modificar_texto_marcador(doc, nombre_marcador, nuevo_texto):
            print(f"- Marcador '{nombre_marcador}' modificado correctamente.")
            modificados_count += 1
        else:
            print(f"- Falló la modificación del marcador '{nombre_marcador}'.")

    if modificados_count > 0:
        try:
            doc.save(output_path)
            print(f"\nDocumento guardado como '{output_path}' con {modificados_count} marcador(es) modificado(s).")
        except Exception as e:
            print(f"\nError al guardar el documento: {e}")
    else:
        print("\nNo se realizaron modificaciones en los marcadores especificados.")

    # crear_nuevo_documento_con_marcadores(marcadores_a_modificar, "nuevo_documento_con_marcadores.docx")


if __name__ == "__main__":
    main()
