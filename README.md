# Documentación del Proyecto: Automatización de Contratos

## Descripción General
Este proyecto tiene como objetivo automatizar la generación, monitoreo, documentación y procesamiento de bases y contratos, en los procesos licitatorios del Hospital San José de Melipilla. Utiliza bibliotecas como `watchdog`, `docx`, `pandas`, y `docxtpl` para manejar eventos del sistema de archivos, procesar plantillas de documentos y generar documentos renderizados basados en datos de entrada.

## Estructura del Proyecto

### Archivos Principales

1. **Finished_Sequence_of_Scripts.py**
   - Monitorea directorios para detectar cambios en archivos específicos.
   - Genera documentos automatizados basados en plantillas y datos de entrada.
   - Utiliza `watchdog` para manejar eventos como creación, modificación, eliminación y movimiento de archivos.

2. **Formated_Contrats_PEP8_ignore.py**
   - Proporciona funciones para manipular documentos de Word, como aplicar numeración, extraer secciones y copiar contenido entre documentos.
   - Utiliza `docx` y `win32com` para manejar documentos de Word y realizar operaciones avanzadas como copiar tablas.

3. **Jinja_2.py**
   - Genera contextos para plantillas de documentos a partir de datos en archivos Excel.
   - Limpia y procesa datos utilizando `pandas`.
   - Facilita la integración de datos dinámicos en plantillas de Word mediante `docxtpl`.

4. **Portada.py**
   - Crea portadas personalizadas para documentos de Word.
   - Inserta logotipos y texto formateado en las portadas.
   - Utiliza `docx` para manipular documentos.

5. **Formated_Base_PEP8.py**
   - Contiene funciones para aplicar formatos globales y específicos a documentos de Word.
   - Permite la creación de tablas, numeración y alineación de contenido de manera programática.
   - Genera la plantilla que va a ser utilizada tanto para los contratos como para las bases

### Directorios

- **Files/**: Contiene plantillas, documentos generados y otros recursos necesarios para la automatización.

## Funcionalidades Clave

### Monitoreo de Directorios
El archivo `Finished_Sequence_of_Scripts.py` utiliza `watchdog` para monitorear directorios y ejecutar acciones automáticas cuando se detectan cambios en archivos específicos, como:
- Generar portadas y documentos base.
- Renderizar documentos basados en datos de Excel.
- Copiar archivos necesarios al directorio de trabajo.

### Generación de Documentos
- **Plantillas Base**: Se utilizan plantillas de Word (`.docx`) para generar documentos personalizados.
- **Renderizado con Datos**: Los datos se extraen de archivos Excel y se integran en las plantillas mediante `docxtpl`.
- **Manipulación Avanzada**: Funciones como `copiar_tablas_con_win32` permiten copiar tablas entre documentos de Word utilizando `win32com`.

### Procesamiento de Datos
- Los datos de entrada se procesan y limpian utilizando `pandas`.
- Se generan diccionarios de contexto para facilitar la integración de datos en plantillas.

## Requisitos

### Dependencias
- Python 3.x
- Bibliotecas:
  - `watchdog`
  - `docx`
  - `pandas`
  - `docxtpl`
  - `win32com`
  - `random`
  - `os`
  - `shutil`
  - `time`
  - `requests`
  - `bs4`
  - `re`
  - `tempfile`
  - `copy`
  - `datetime`
  - `pythoncom`


### Configuración
1. Asegúrese de que las rutas de los archivos y directorios estén correctamente configuradas en los scripts.
2. Instale las dependencias necesarias utilizando `pip install`.
3. Verifique que las plantillas y archivos de datos estén disponibles en los directorios especificado.

## Uso

1. **Monitoreo de Directorios**:
   - Ejecute `Finished_Sequence_of_Scripts.py` para iniciar el monitoreo de directorios.
   - Los eventos como la creación o modificación de archivos desencadenarán la generación/procesamiento de documentos automatizados.

2. **Generación de Documentos**:
   - Asegúrese de que los datos de entrada (por ejemplo, `Libro1.xlsx`) estén actualizados.
   - Los documentos generados se guardarán dentro de la misma carpeta que desencadenó el proceso.

3. **Personalización**:
   - Modifique las plantillas de Word según sea necesario.
   - Actualice los scripts para adaptarlos a requisitos específicos.

## Notas
- Todavía existen detalles dentro de los documentos que deben ser ajustados, así como también algunos elementos que no están correctamente condicionados, por lo que es necesario revisar las bases o correjir el código
- Asegúrese de realizar pruebas exhaustivas después de realizar cambios en los scripts o plantillas además de respaldar los archivos antes de hacer cambios.
