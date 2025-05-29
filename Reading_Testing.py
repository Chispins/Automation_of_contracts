# python
from win32com.client import Dispatch, constants
import os
from Formated_Base_PEP8 import configurar_directorio_trabajo

configurar_directorio_trabajo()
# Rutas de los archivos
src_path = os.path.abspath(r'prototipo_tabla_rellenado.docx')
dest_path = os.path.abspath(r'destino_con_tablas.docx')

# Iniciar Word
word = Dispatch('Word.Application')
word.Visible = False

# Abrir documento origen
src_doc = word.Documents.Open(src_path)

# Crear documento destino
dest_doc = word.Documents.Add()

# Selección para pegar
selection = word.Selection

# Recorrer y copiar cada tabla del origen
for tbl in src_doc.Tables:
    tbl.Range.Copy()
    # Mover al final del destino
    selection.EndKey(Unit=constants.wdStory)
    selection.Paste()
    # Insertar salto de línea tras la tabla
    selection.TypeParagraph()

# Guardar y cerrar
dest_doc.SaveAs(dest_path)
src_doc.Close(False)
dest_doc.Close(False)
word.Quit()