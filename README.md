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

Para generar los contratos, se sigue un flujo largo, donde se generan varios archivos intermedios, el archivo final que nos interesa es 'contrato_automatizado_tablas_rendered'. El detalle de los archivos intermedios se ve en la siguiente imagen 
![Image](https://github.com/user-attachments/assets/e0b777d6-41bc-415f-a552-646835f37553)

### Detalle del Flujo de Generación de Documentos

Esta tabla detalla el flujo de generación de documentos, explicando el propósito de cada archivo clave en el proceso.


| Nombre del Archivo | Descripción del archivo |
| :--- | :--- |
| `Libro1.xlsx` | Es el archivo excel que **se debe modificar**, posee 3 hojas, la primera es de elementos de la base, la segunda y tercera poseen detalles del contrato que deben ser rellenados, el primer elemento de la cuarta fila **en TODAS LAS HOJAS no puede ser vacio, o sino el codigo no funciona**, se recomienda colocar "1", este excel es el que **sirve como fuente** en el que se basarán todos los archivos posteriores. |
| `plantilla_original_rendered.docx` | Es el archivo de la **Base ya listo y procesado** con los valores remplazados. |
| `portada_melipilla_contrato.docx` | Es la **portada de un contrato**. |
| `portada_melipilla_contrato_renderizado.docx` | Es la portada **solamente con los valores remplazados**. |
| `contrato_automatizado_over.docx` | **Toma el archivo anterior de portada**, y le **agrega los primeros 3 items** para un contrato. |
| `contrato_faltan_tablas.docx` | Toma el archivo anterior, y esa será la primera parte del documento, luego sobre eso vamos a pegar todos los items desde el tercero hasta el Vigésimo Óctavo provenientes desde plantilla contrato, Lo que hace es buscar un título, ese título y todo el contenido que posea un nivel de título inferior será copiado y pegado, notar que esto **hace perder cualquier tipo de formato** que posea el documento de origen. Ademas **remplaza todos los lugares donde habían tablas por "[[ TABLE PLACEHOLDER ]]"**. |
| `contrato_automatizado_tablas.docx` | En esta parte se toma el documento anterior y luego **remplaza todas los espacios donde hay [[TABLE PLACEHOLDER]] por las tablas** de 'plantilla_original.docx', por lo que es importante en este sentido mencionar que habrá que **modificar esas tablas para que cuadren** con los cambios en las adjudiaciones, además este proceso es el **más sensible y propenso a fallos**, porque **requiere que esté instalado word** en el computador que este corriendo el código, además, si los archivos **están abiertos la copia podría fallar**, las tablas **mantendrán todos sus formatos** y propiedades originales. Notar también que las tablas son ingresadas, **en el orden en el que están presentes en las bases**, por lo que si se crean nuevas tablas, entonces **el procedimiento podría fallar**, además que como originalmente estaba pensado también realizar el procesamiento de las garantías de fiel cumplimiento, entonces **será necesario insertar un documento "prototipo_tabla_rellenado.docx"** (que son 2 tablas) en la carpeta. |


## Ejemplo de Uso
Necesitamos llevar una licitación para la compra de examenes, por lo que vamos al compartido y creamos una nueva carpeta en Licitaciones Testing/1057480-15-LR25
-- Se generan los Archivos en la carpeta --
Dentro de los archivos veremos varios, sin embargo el que nos interesa se llama `Libro1.xlsx`, este archivo es el que debemos de completar, este archivo tendrá 3 hojas, para generar una base Se rellena la primera Hoja de Libro1.xlsx, y luego se escribe CONFIRMAR en D4.
Listo, ya se debería generar la Base para la licitación.
Luego, cuando ya se debe realizar el contrato, se rellena la segunda y tercera hoja del excel y se rellena la celda D4 de la tercera hoja.
Listo, ya se debería generar el contrato para la licitación.
