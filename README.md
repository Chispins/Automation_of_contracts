# README: Sistema Automatizado de Generación de Documentos

Este programa permite la creación automática de Bases y Contratos para las licitaciones. Remplaza el trabajo de generación manual de los mismos y evita errores. Además genera un registro con todas las bases y contratos creados utilizando esta herramienta.

La ejecución lo que garantiza es proveer los archivos necesarios cada vez que se crea una carpeta para generar la base/contrato, y también garantiza que cuando se cumplan las condiciones para generar la base/contrato, entonces los genera. A continuación se detalla un poco el proceso acerca de "como" es que se logran estos objetivos, cuales son esos criterios, y cual es el proceso.

El programa sigue la siuiente secuencia


![Image](https://github.com/user-attachments/assets/0beae708-343e-40ae-8286-b9bbeef86e9f)


Lo primero que sucede al activar el programa es que se crea un vigilante que estará siempre mirando las carpetas dentro de la carpeta principal, este vigilante estará observando dos tipos de Eventos la creación de carpetas y la modificación de archivos. Esto es para asegurarse de proveer los archivos necesarios y para que cuando se cumplan las condiciones genere las Bases y Contratos de Licitación y guarda un registro del mismo.




### 1. **Inicio Monitoreo**
El programa revisa cada segundo si hay archivos oh carpetas nuevos o modificados. Para que dependiendo del caso generar una base, contrato, o los archivos.

## 2 **Evento**
El vigilante detecta un cambio y notifica un evento, el cual puede ser una modificación de un archivo o creación de una carpeta.

## 3 **¿Es creación de carpeta?**
¿Es el evento una creación de una carpeta? En caso de ser **NO** se salta al paso 5, en caso de ser **SI** la respuesta entonces se pasa al paso 3.

## 4 **Creación de Carpeta de Licitación**
El evento fue una creación de carpeta nueva dentro de la cual se llevará un proceso de licitación, al crearse la carpeta, se deberán proveer los archivos necesarios para poder llevar un archivo, esos archivos necesarios son los que se generan en el siguiente paso.

## 5 **Generación de archivos necesarios **
En caso de que el evento sea una creación de carpeta entonces el vigilante pegará todos los archivos necesarios para el correcto funcionamiento dentro de la carpeta recién creada.
- portada_melipilla_base.docx Es el archivo que será la portada de la Base
- portada_melipilla_contrato.docx Es el archivo que será la portada del contrato
- plantilla_original.docx Es el word que será la plantilla, sobre este archivo se trabajará para crear una base
- Libro1.xlsx Es un excel de 3 hojas, donde la primera corresponde a información para la base, la segunda y la tercera son información para el contrato

## 6 **Modificación del Excel**
Si el evento no es creación entonces es modificación de un archivo, aunque en realidad nos interesa solamente si se modifica el excel, cada vez que se modifique el excel, procederán los siguientes pasos, verificar si es que se cumplen las condiciones para generar una base o contrato y generarlos en caso de que se cumplan. Esto como se decía para garantizar que en **el momento** en que se quiera generar una base o contrato, se genere, sin tener que manualmente hacer nada mas que escribir confirmar en el mismo excel


## 7 **Verificación de Requerimientos**
En caso de que el *"Evento"** no sea una creación de carpeta, significa que es una modificación de un archivo, por lo que necesitamos verificar que si el archivo del evento, es el excel para la generación de la base y/o contrato. En específico lo que verifica es que se cumplan **todas** las siguientes condiciones.
| Requisito | ¿Qué pasa si falta? | ¿Cómo solucionarlo? |
|-----------|---------------------|---------------------|
| **`CONFIRMAR`** en la columna D4 de la primera hoja de Libro1.xlsx | La Base **NO se genera** | Escribir `CONFIRMAR` en la celda D4 y luego guardar|
| **`Plantilla_original.docx** en la carpeta de la licitación`** | La Base **NO se genera** | Copia el archivo desde otra carpeta, o crear otra carpeta y llevar el proceso de licitación en esa nueva carpeta |
| **`CONFIRMAR`** **ESTA SELECCIONADO** en la columna D4 de la tercera hoja de Libro1.xlsx | La base **NO se genera** y se procederá al paso | Borrar lo que esté escrito en la celda D4 de la tercera hoja y luego guardar |


## 8. Generar Base y otros archivos intermedios 
El programa comenzará el procesamiento, lo que hace es tomar los datos que fueron rellenados en el excel, luego remplazará en el archivo plantilla original con los valores del excel, agregará una portada base, y luego guardará un nuevo archivo que se llamará plantilla_original_rendered, este nuevo archivo será una Base que está finalizada y lista.

Se procede a la generación del archivo de base, utilizando **`plantilla_original.docx`**. El programa crea un nuevo archivo de Base, utilizando la portada de la Base, y escribiendo todos los elementos de plantilla original que se utilizan en una base, remplazando los valores por los Valores que están presentes en Hoja 1, este nuevo archivo guardado es almacenado como plantilla_original_rendered, este nuevo archivo es una Base que está Finalizada y lista.
### 9. Verificación de requerimientos

En caso de que ya exista una base creada en la carpeta se comenzará a verificar las siguientes condiciones

| Requisito | ¿Qué pasa si falta? | ¿Cómo solucionarlo? |
|-----------|---------------------|---------------------|
| **`CONFIRMAR`** en la columna D4 de la tercera hoja Libro1.xlsx **NO ESTÁ SELECCIONADO** | El reporte **NO se genera** | 1. Consigue el archivo de gastos del mes<br>2. Colócalo en la carpeta del mes<br>3. Asegúrate que se llame el nombre comienza con `DEVENGADO` | Escribir `CONFIRMAR` en la celda D4 y luego guardar|
| **`Plantilla_original_rendered.docx** en la carpeta de la licitación`** | La Base **NO se genera** | Copia el archivo desde otra carpeta, o borrar confirmar de la hoja 3, y apretar CONFIRMAR en la celda D4 en la primera hoja para generar la base.|


### 10. Generación Contrato
Se procede a la generación del archivo de contrato, utilizando la misma **`plantilla_original.docx`**  que utiliza la base. El programa crea un nuevo archivo de contrato, utilizando la portada del contrato, y escribiendo todos los elementos de plantilla original que se utilizan en un contrato, solo que ahora remplaza por los valores de la Hoja 1, Hoja 2, y Hoja 3. La diferencia es el resultado de este procesamiento entregará un contrato listo.

## Ejemplo de Uso
Necesitamos llevar una licitación para la compra de examenes, por lo que vamos al compartido y creamos una nueva carpeta en Licitaciones Testing/1057480-15-LR25
# Se generan los Archivos en la carpeta
Dentro de los archivos veremos varios, sin embargo el que nos interesa se llama `Libro1.xlsx`, este archivo es el que debemos de completar, este archivo tendrá 3 hojas, para generar una base Se rellena la primera Hoja de Libro1.xlsx, y luego se escribe CONFIRMAR en D4.
Listo, ya se debería generar la Base para la licitación.
Luego, cuando ya se debe realizar el contrato, se rellena la segunda y tercera hoja del excel y se rellena la celda D4 de la tercera hoja.
Listo, ya se debería generar el contrato para la licitación.


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



# Generador Automático de Bases y Contratos

## ¿Para qué sirve?
Crea automáticamente documentos legales para licitaciones, evitando:
- Errores manuales
- Tiempo de creación
- Formateo inconsistente

## ¿Cómo funciona? (Pasos simples)

1. **Preparación inicial**:
   - Crea una carpeta nueva para cada licitación
   - El programa copiará automáticamente 4 archivos esenciales:
     - `portada_melipilla_base.docx` (Portada Base)
     - `portada_melipilla_contrato.docx` (Portada Contrato)
     - `plantilla_original.docx` (Documento principal)
     - `Libro1.xlsx` (Datos para completar)

2. **Creación de la Base**:
   - Abre `Libro1.xlsx`
   - Completa la **Hoja 1**
   - Escribe `CONFIRMAR` en la celda **D4**
   - Guarda el archivo
   - El programa generará:  
     `plantilla_original_rendered.docx` (Base lista)

3. **Creación del Contrato**:
   - Completa las **Hojas 2 y 3** de `Libro1.xlsx`
   - Escribe `CONFIRMAR` en la celda **D4 de la Hoja 3**
   - Guarda el archivo
   - El programa generará el contrato final

## Solución de problemas

| Problema               | Solución                          |
|------------------------|-----------------------------------|
| Base no se genera      | Verificar que `CONFIRMAR` está escrito en D4 de Hoja 1 |
| Contrato no se genera  | Verificar `CONFIRMAR` en D4 de Hoja 3 |
| Faltan archivos        | Crear nueva carpeta para que el programa los copie |

## Ejemplo completo: Licitación 1057480-15-LR25

1. Crear carpeta:  
   `Licitaciones Testing/1057480-15-LR25`

2. El programa copia automáticamente:
