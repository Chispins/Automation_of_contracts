import os
import random
import docx
import copy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from win32com import client
import pythoncom
import re


def configurar_directorio_trabajo():
    """Configura el directorio de trabajo en la subcarpeta 'Files'."""
    # Construye la ruta normalizada al directorio "Files"
    cwd = os.getcwd()
    target_dir_name = "Files"
    wd = os.path.join(cwd, target_dir_name)

    # Define el patrón específico a buscar (duplicación de \Files)
    # Se asume separador de Windows (\). Se usan dobles barras invertidas en el patrón regex.
    pattern = r"Files\\Files"

    # Busca el patrón en la ruta generada
    if re.search(pattern, wd):
        # Si se encuentra, reemplaza la primera ocurrencia de la duplicación
        wd = wd.replace(r"\Files\Files", r"\Files")

    # Cambia al directorio destino, verificando primero si es un directorio válido
    if os.path.isdir(wd):
        os.chdir(wd)
        print(f"Directorio de trabajo cambiado a: {wd}")  # Opcional: confirmar cambio
    else:
        print(f"Advertencia: El directorio '{wd}' no existe o no es válido. No se cambió el directorio de trabajo.")


# --- Funciones para Numeración de Párrafos ---
def crear_numeracion(doc):
    """Crea un formato de numeración y devuelve su ID único."""
    part = doc._part
    if not hasattr(part, 'numbering_part'):
        part._add_numbering_part()
    num_id = random.randint(1000, 9999)
    return num_id


def aplicar_numeracion(parrafo, num_id, nivel=0):
    """
    Aplica numeración a un párrafo con un ID y nivel específicos.
    Elimina numeración previa si existe y ajusta la sangría.
    """
    p = parrafo._p
    pPr = p.get_or_add_pPr()

    # Eliminar numeración previa
    for child in pPr.iterchildren():
        if child.tag.endswith('numPr'):
            pPr.remove(child)

    # Crear nueva numeración
    numPr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(nivel))
    numPr.append(ilvl)
    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), str(num_id))
    numPr.append(numId)
    pPr.append(numPr)

    # Configurar sangría
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), '720')
    ind.set(qn('w:hanging'), '360')
    pPr.append(ind)

    return parrafo


# --- Funciones para Extracción y Copia de Secciones ---
def extraer_seccion_completa(doc_cargado, titulo_seccion):
    """
    Extrae una sección completa (párrafos y tablas) hasta el siguiente encabezado
    del mismo nivel o superior.
    Retorna el encabezado, los elementos de la sección y el nivel del encabezado.
    """
    seccion_heading = None
    elementos_seccion = []
    nivel_seccion = None
    indice_inicio = -1
    indice_fin = None

    # Buscar el encabezado de la sección
    for i, element in enumerate(doc_cargado.paragraphs):
        if element.style.name.startswith('Heading') and element.text.strip() == titulo_seccion:
            seccion_heading = element
            indice_inicio = i
            try:
                nivel_seccion = int(element.style.name.split()[-1])
            except (ValueError, IndexError):
                nivel_seccion = 2  # Nivel por defecto
            break

    if indice_inicio == -1:
        print(f"No se encontró la sección '{titulo_seccion}'")
        return None

    # Encontrar el fin de la sección
    for i in range(indice_inicio + 1, len(doc_cargado.paragraphs)):
        parrafo = doc_cargado.paragraphs[i]
        if parrafo.style.name.startswith('Heading'):
            try:
                nivel_actual = int(parrafo.style.name.split()[-1])
                if nivel_actual <= nivel_seccion:
                    indice_fin = i
                    break
            except (ValueError, IndexError):
                pass

    if indice_fin is None:
        indice_fin = len(doc_cargado.paragraphs)

    # Agregar párrafos entre inicio y fin (excepto encabezados)
    for i in range(indice_inicio + 1, indice_fin):
        parrafo = doc_cargado.paragraphs[i]
        if not parrafo.style.name.startswith('Heading'):
            elementos_seccion.append(('parrafo', parrafo, i))

    # Mapa de posiciones de párrafos en XML
    parrafos_posicion_xml = {p._p: i for i, p in enumerate(doc_cargado.paragraphs)}

    # Extraer posición de tablas en XML
    tablas_posicion = []
    for i, tabla in enumerate(doc_cargado.tables):
        tabla_p = tabla._tbl
        elemento_anterior = tabla_p.getprevious()
        pos_anterior = -1
        while elemento_anterior is not None:
            if elemento_anterior in parrafos_posicion_xml:
                pos_anterior = parrafos_posicion_xml[elemento_anterior]
                break
            elemento_anterior = elemento_anterior.getprevious()
        if indice_inicio <= pos_anterior < indice_fin:
            tablas_posicion.append((i, tabla, pos_anterior + 0.5))

    # Agregar tablas a la sección
    for i, tabla, pos in tablas_posicion:
        elementos_seccion.append(('tabla', tabla, pos))

    # Ordenar elementos por posición
    elementos_seccion.sort(key=lambda x: x[2])
    elementos_seccion = [(tipo, elem) for tipo, elem, _ in elementos_seccion]

    return seccion_heading, elementos_seccion, nivel_seccion if seccion_heading else None


def copiar_seccion_completa(doc_destino, seccion_heading, elementos_seccion, nivel_seccion):
    """
    Copia una sección completa al documento destino manteniendo numeración única
    para párrafos de lista y preservando marcadores (bookmarks) en encabezados y párrafos.
    No copia tablas, solo añade un marcador de posición.
    """
    # Copiar encabezado
    nuevo_encabezado = doc_destino.add_heading(seccion_heading.text, level=nivel_seccion)

    # Copiar propiedades del encabezado, incluyendo marcadores si los tiene
    if seccion_heading._p.pPr is not None:
        nuevo_encabezado._p.append(copy.deepcopy(seccion_heading._p.pPr))

    # Copiar contenido del encabezado, incluyendo marcadores
    for child in seccion_heading._p:
        if child.tag == qn('w:r'):
            nuevo_encabezado._p.append(copy.deepcopy(child))
        elif child.tag == qn('w:bookmarkStart'):
            new_bm_start = OxmlElement('w:bookmarkStart')
            bm_id = child.get(qn('w:id'))
            bm_name = child.get(qn('w:name'))
            if bm_id:
                new_bm_start.set(qn('w:id'), bm_id)
            if bm_name:
                new_bm_start.set(qn('w:name'), bm_name)
            nuevo_encabezado._p.append(new_bm_start)
        elif child.tag == qn('w:bookmarkEnd'):
            new_bm_end = OxmlElement('w:bookmarkEnd')
            bm_id = child.get(qn('w:id'))
            if bm_id:
                new_bm_end.set(qn('w:id'), bm_id)
            nuevo_encabezado._p.append(new_bm_end)

    # ID único de numeración para la sección
    seccion_num_id = None

    # Copiar elementos de la sección
    for tipo, elemento in elementos_seccion:
        if tipo == 'parrafo':
            parrafo_origen = elemento
            nuevo_parrafo = doc_destino.add_paragraph(style=parrafo_origen.style.name)

            # Copiar propiedades del párrafo (estilo, alineación, etc.)
            if parrafo_origen._p.pPr is not None:
                nuevo_parrafo._p.append(copy.deepcopy(parrafo_origen._p.pPr))

            # Copiar contenido y marcadores del párrafo
            for child in parrafo_origen._p:
                if child.tag == qn('w:r'):
                    nuevo_parrafo._p.append(copy.deepcopy(child))
                elif child.tag == qn('w:bookmarkStart'):
                    new_bm_start = OxmlElement('w:bookmarkStart')
                    bm_id = child.get(qn('w:id'))
                    bm_name = child.get(qn('w:name'))
                    if bm_id:
                        new_bm_start.set(qn('w:id'), bm_id)
                    if bm_name:
                        new_bm_start.set(qn('w:name'), bm_name)
                    nuevo_parrafo._p.append(new_bm_start)
                elif child.tag == qn('w:bookmarkEnd'):
                    new_bm_end = OxmlElement('w:bookmarkEnd')
                    bm_id = child.get(qn('w:id'))
                    if bm_id:
                        new_bm_end.set(qn('w:id'), bm_id)
                    nuevo_parrafo._p.append(new_bm_end)

            # Aplicar numeración si el estilo es de lista
            if parrafo_origen.style.name == "List Number":
                if seccion_num_id is None:
                    seccion_num_id = crear_numeracion(doc_destino)
                aplicar_numeracion(nuevo_parrafo, seccion_num_id)

        elif tipo == 'tabla':
            # No copiar tablas, solo añadir un marcador de posición
            doc_destino.add_paragraph("[[TABLE_PLACEHOLDER]]")


def copiar_tablas_con_win32(source_path, intermediate_path, output_path):
    """
    Usa win32com para copiar tablas del documento original al documento intermedio,
    reemplazando los marcadores de posición.
    """
    pythoncom.CoInitialize()
    try:
        word_app = client.Dispatch("Word.Application")
        word_app.Visible = False

        # Abrir documentos
        source_doc = word_app.Documents.Open(os.path.abspath(source_path))
        intermediate_doc = word_app.Documents.Open(os.path.abspath(intermediate_path))

        # Contar tablas en el documento original
        source_tables = source_doc.Tables
        table_index = 0

        # Buscar marcadores de posición en el documento intermedio
        for para in intermediate_doc.Paragraphs:
            if "[[TABLE_PLACEHOLDER]]" in para.Range.Text:
                if table_index < source_tables.Count:
                    # Copiar la tabla correspondiente del documento original
                    source_table = source_tables.Item(table_index + 1)
                    source_table.Range.Copy()

                    # Pegar la tabla en el lugar del marcador
                    para.Range.Paste()
                    table_index += 1

        # Guardar el documento final
        intermediate_doc.SaveAs(os.path.abspath(output_path))
        intermediate_doc.Close()
        source_doc.Close()
        word_app.Quit()

    finally:
        pythoncom.CoUninitialize()


# --- Flujo Principal del Programa ---
def main():
    """Función principal que ejecuta el proceso de copia de secciones y tablas de un documento."""
    # Configurar directorio de trabajo
    configurar_directorio_trabajo()

    # Configuración del documento original
    input_file = "resolucion_numerada.docx"
    intermediate_file = "onlydocx.docx"
    final_output = "seccion_completa_copiada.docx"

    # Cargar documento original y crear documento destino
    print("Cargando documento original...")
    word = docx.Document(input_file)
    doc = docx.Document()
    print("Documento cargado.")

    # Lista de secciones a procesar
    secciones = [
        "VISTOS",
        "CONSIDERANDO",
        "RESOLUCIÓN",
        "REQUISITOS"
    ]

    # Extraer y copiar cada sección
    for seccion in secciones:
        print(f"Procesando sección {seccion}...")
        resultado = extraer_seccion_completa(word, seccion)
        if resultado:
            encabezado, elementos, nivel = resultado
            copiar_seccion_completa(doc, encabezado, elementos, nivel)
            print(f"Sección {seccion} copiada.")
        else:
            print(f"Sección {seccion} no encontrada.")

    # Guardar documento intermedio con solo texto
    print(f"Guardando documento intermedio como: {intermediate_file}...")
    doc.save(intermediate_file)
    print("Documento intermedio guardado.")

    # Copiar tablas usando win32com
    print("Copiando tablas al documento final...")
    copiar_tablas_con_win32(input_file, intermediate_file, final_output)
    print(f"Documento final guardado como: {final_output}")


if __name__ == "__main__":
    main()
