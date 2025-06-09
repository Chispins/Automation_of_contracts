# README: Sistema Automatizado de Generación de Documentos

## 📝 Descripción General
Este sistema automatiza la creación de Bases y Contratos para los procesos de Licitaciones del HSJM un código que responde a acciones del usuario en archivos Excel.

## 🔄 Flujo de Trabajo Principal

### 1. Monitoreo Inicial
- **Tecnología**: Utiliza `watchdog` para vigilancia continua
- **Alcance**: 
  - Monitorea carpeta principal y todas las subcarpetas
  - Detecta cambios en tiempo real (creación/modificación de archivos)

### 2. Verificación de Archivos Base
Verifica que existan los
**Archivos requeridos**:
- `Libro1.xlsx` (plantilla de datos)
- `plantilla_original.docx` (documento base)
- `portada_melipilla_base.docx`, `portada_melipilla_contrato.docx`
Plantilla original es un documento que contiene todo el texto que será utilizado tanto en las bases como en los contratos, donde aplica `{{}}` en cada elemento que es variable, y cada uno de esos elementos será remplazado posteriormente con las variables designadas en `Libro1.xlsx`. 

## 3. Generación de documentos Necesarios
En caso de no existir los archivos necesarios, el código pega los siguientes archivos `plantilla_original.docx`, `Libro1.xlsx`, `portada_melipilla_base.docx`, `portada_melipilla_contrato.docx` (output_1) en el work directory(wd), los archivos son copiados desde la carpeta `NO_BORRAR`.
Este evento ocurre cada vez que se:

-Crea una carpeta.
-Mueve una carpeta dentro de la dirección monitoreada.
-Borra uno de los documentos necesarios

### 4. Generación de Documento "Base"
**Condiciones de activación**:
1. Modificación reciente de `Libro1.xlsx`
2. Celda D2 contiene "si" o cualquier otro texto
3. No existe documento final existente en la carpeta

**Proceso**:
1. Lee datos de Excel con `pandas`
2. Genera contexto para plantilla utilizando la información de `Libro1.xlsx`
3. Rellena `plantilla_original.docx` Word.
4. Guarda como `plantilla_original_rendered.docx` (Output_2)

### 5. Generación de Contrato
**Condicionebs de activación**:
1. Modificación reciente de `Libro1.xlsx`
2. Celda específica contiene "si" 
3. No existe contrato final previo

**Proceso**:
1. Verifica existencia de documento base.
2. Genera una plantilla de Contrato, que es en realidad el documento de `plantilla_original_rendered.docx` eliminando los elementos que no se utilizan.
3. Genera el contexto para plantilla con la información de `Libro1.xlsx` 
4. Render en la portada 
5. Render en la plantilla del Contrato
6. Genera el contrato mezclando los dos documentos
7. Coloca los table placeholders en los lugares donde deberían ir tabla, guarda el documento como `contrato_autonmatizado_faltan_tablas.docx`
8. Remplaza los table placeholders, por tablas pegandolas utilizando el paquete win32com y lo guarda como `contrato_automatizado_rendered.docx` (Output_3)


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
