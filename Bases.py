import os
import docx
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re

# Construye la ruta normalizada al directorio "Files"
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
        print(f"Directorio de trabajo cambiado a: {wd}") # Opcional: confirmar cambio
    else:
        print(f"Advertencia: El directorio '{wd}' no existe o no es válido. No se cambió el directorio de trabajo.")
        # Aquí podrías decidir crear el directorio si no existe, o manejar el error.
        # Ejemplo para crearlo:
        # try:
        #     os.makedirs(wd, exist_ok=True) # Crea el directorio si no existe
        #     os.chdir(wd)
        #     print(f"Directorio '{wd}' creado y establecido como directorio de trabajo.")
        # except OSError as e:
        #     print(f"Error al crear o acceder al directorio '{wd}': {e}")

# Llamar a la función para configurar el directorio de trabajo
configurar_directorio_trabajo()


def crear_numeracion(doc):
    """Crea un formato de numeración y devuelve su ID"""
    # Asegurarse de que exista la parte de numeración
    part = doc._part
    if not hasattr(part, 'numbering_part'):
        part._add_numbering_part()

    # Crear un ID único para esta numeración
    import random
    num_id = random.randint(1000, 9999)  # ID aleatorio para evitar conflictos
    return num_id
# Función para aplicar numeración a un párrafo
def aplicar_numeracion(parrafo, num_id, nivel=0):
    """Aplica numeración a un párrafo"""
    p = parrafo._p
    pPr = p.get_or_add_pPr()

    # Eliminar numeración previa si existe
    for child in pPr.iterchildren():
        if child.tag.endswith('numPr'):
            pPr.remove(child)

    # Crear nuevo numPr
    numPr = OxmlElement('w:numPr')

    # Establecer nivel
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(nivel))
    numPr.append(ilvl)

    # Establecer ID de numeración
    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), str(num_id))
    numPr.append(numId)

    # Agregar al párrafo
    pPr.append(numPr)

    # Configurar sangría
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), '720')
    ind.set(qn('w:hanging'), '360')
    pPr.append(ind)

    return parrafo

# Configuración del documento

base_partida = "Base_en_Blanco.docx"
doc = docx.Document()

# Crear un único estilo para listas numeradas (o usar el predefinido)
list_style = 'List Number'
if list_style not in doc.styles:
    doc.styles.add_style(list_style, WD_STYLE_TYPE.PARAGRAPH)

# Crear IDs para numeración
num_id_vistos = crear_numeracion(doc)
num_id_resolucion = crear_numeracion(doc)
num_id_bases_p1 = crear_numeracion(doc)

# Centrar para tablas
def centrar_verticalmente_tabla(tabla):
    """Aplica alineación vertical centrada a todas las celdas de una tabla"""
    for fila in tabla.rows:
        for celda in fila.cells:
            # Acceder al elemento XML de la celda
            tc = celda._tc

            # Buscar o crear el elemento tcPr (propiedades de celda)
            tcPr = tc.get_or_add_tcPr()

            # Eliminar cualquier configuración previa de alineación vertical
            for vAlign in tcPr.findall(qn('w:vAlign')):
                tcPr.remove(vAlign)

            # Crear nuevo elemento de alineación vertical
            vAlign = OxmlElement('w:vAlign')
            vAlign.set(qn('w:val'), 'center')  # Valor 'center' para centrado vertical

            # Agregar el elemento a las propiedades de la celda
            tcPr.append(vAlign)


# Titulos Nivel 0
doc.add_heading("RESOLUCIÓN EXENTA Nº1", level=0)

# SECCIÓN VISTOS
doc.add_heading("VISTOS", level=2)
doc.add_paragraph("Lo dispuesto en la Ley Nº 19.886 de Bases sobre Contratos Administrativos de Suministro y Prestación de Servicios; el Decreto Supremo Nº 250 /04 modificado por los Decretos Supremos Nº 1763/09, 1383/11 y 1410/14 todos del Ministerio de Hacienda; D. S. 38/2005, Reglamento Orgánico de los Establecimientos de Menor Complejidad y de los Establecimientos de Autogestión en Red; en uso de las atribuciones que me confieren el D.F.L. Nº 1/2.005, en virtud del cual se fija el texto refundido, coordinado y sistematizado del D.L. 2.763/79 y de las leyes 18.933 y 18.469; lo establecido en los Decretos Supremos Nos 140/04, Reglamento Orgánico de los Servicios de Salud; la Resolución Exenta RA 116395/343/2024 de fecha 12/08/2024 del SSMOCC., la cual nombra Director del Hospital San José de Melipilla al suscrito; lo dispuesto por las Resoluciones 10/2017, 7/2019 y 8/2019 ambas de la Contraloría General de la República, y,")

# SECCIÓN CONSIDERANDO
doc.add_heading("CONSIDERANDO", level=2)

# Lista numerada para CONSIDERANDO (usando el mismo ID que antes usábamos para VISTOS)
p1 = doc.add_paragraph("Visto: La Ley N° 19.880, de 2003, que establece normas sobre los actos administrativos...",
                    style=list_style)
aplicar_numeracion(p1, num_id_vistos)

p2 = doc.add_paragraph("Que, el Hospital de San José de Melipilla perteneciente a la red de salud...",
                    style=list_style)
aplicar_numeracion(p2, num_id_vistos)

p3 = doc.add_paragraph("Que, dada la naturaleza del Establecimiento...", style=list_style)
aplicar_numeracion(p3, num_id_vistos)

# Párrafo con formato en negrita
vistos_p4 = doc.add_paragraph(style=list_style)
vistos_p4.add_run("Que, existe la necesidad ")
run_bold = vistos_p4.add_run("suministro de insumos y accesorios para terapia de presión negativa con equipos en comodato")
run_bold.bold = True
vistos_p4.add_run(", a fin de entregar una prestación de salud integral y oportuna...")
aplicar_numeracion(vistos_p4, num_id_vistos)

p5 = doc.add_paragraph("Que corresponde asegurar la transparencia en este proceso...", style=list_style)
aplicar_numeracion(p5, num_id_vistos)

p6 = doc.add_paragraph("Que, considerando los montos de la contratación...", style=list_style)
aplicar_numeracion(p6, num_id_vistos)

p7 = doc.add_paragraph("Que revisado el catálogo de bienes y servicios...", style=list_style)
aplicar_numeracion(p7, num_id_vistos)

p8 = doc.add_paragraph("Que, en consecuencia y en mérito de lo expuesto...", style=list_style)
aplicar_numeracion(p8, num_id_vistos)

p9 = doc.add_paragraph("Que, en razón de lo expuesto y la normativa vigente;", style=list_style)
aplicar_numeracion(p9, num_id_vistos)

# SECCIÓN RESOLUCIÓN
doc.add_heading("RESOLUCIÓN", level=2)

# Lista numerada para RESOLUCIÓN (nuevo ID = nueva secuencia)
resolucion_p1 = doc.add_paragraph(style=list_style)
resolucion_p1.add_run("LLÁMASE ").bold = True
resolucion_p1.add_run("a Licitación Pública Nacional a través del Portal Mercado Público, para la compra de ")
resolucion_p1.add_run("Suministro de Insumos y Accesorios para Terapia de Presión Negativa con Equipos en Comodato ").bold = True
resolucion_p1.add_run("para el Hospital San José de Melipilla.")
aplicar_numeracion(resolucion_p1, num_id_resolucion)

resolucion_p2 = doc.add_paragraph(style=list_style)
resolucion_p2.add_run("ACOGIENDOSE ").bold = True
resolucion_p2.add_run("al Art.º 25 del decreto 250 que aprueba el reglamento de la ley Nº 19.886...")
aplicar_numeracion(resolucion_p2, num_id_resolucion)

resolucion_p3 = doc.add_paragraph(style=list_style)
resolucion_p3.add_run("APRUÉBENSE las bases administrativas, técnicas y anexos N.º 1, 2, 3, 4, 5, 6, 7, 8 y 9 ").bold = True
resolucion_p3.add_run("desarrollados para efectuar el llamado a licitación, que se transcriben a continuación:")
aplicar_numeracion(resolucion_p3, num_id_resolucion)

# Nueva Sección
doc.add_heading("BASES ADMINISTRATIVAS PARA EL SUMINISTRO DE INSUMOS Y ACCESORIOS PARA TERAPIA DE PRESIÓN NEGATIVA CON EQUIPOS EN COMODATO PARA EL HOSPITAL SAN JOSÉ DE MELIPILLA",level=1)
doc.add_heading("REQUISITOS", level=2)
doc.add_paragraph("En Santiago, a 1 de enero de 2023, se resuelve lo siguiente:")

# Lista numerada para BASES (nuevo ID = nueva secuencia)
bases_p1 = doc.add_paragraph(style=list_style)
bases_p1.add_run("Antecedentes Básicos de la ENTIDAD LICITANTE").bold = True
aplicar_numeracion(bases_p1, num_id_bases_p1)

# Tabla
tabla_p1 = doc.add_table(6, 2, style="Table Grid")
tabla_p1.cell(0, 0).text = "Razón Social del organismo"
tabla_p1.cell(0, 1).text = "Hospital San José de Melipilla"
tabla_p1.cell(1, 0).text = "Unidad de Compra"
tabla_p1.cell(1, 1).text = "Unidad de Abastecimiento de Bienes y Servicios"
tabla_p1.cell(2, 0).text = "R.U.T. del organismo"
tabla_p1.cell(2, 1).text = "61.602.123-0"
tabla_p1.cell(3, 0).text = "Dirección"
tabla_p1.cell(3, 1).text = "O'Higgins #551"
tabla_p1.cell(4, 0).text = "Comuna"
tabla_p1.cell(4, 1).text = "Melipilla"
tabla_p1.cell(5, 0).text = "Región en que se genera la Adquisición"
tabla_p1.cell(5, 1).text = "Región Metropolitana"
centrar_verticalmente_tabla(tabla_p1)

bases_p2 = doc.add_paragraph(style=list_style)
bases_p2.add_run("Antecedentes Administrativos").bold = True
aplicar_numeracion(bases_p2, num_id_bases_p1)

tabla_p2 = doc.add_table(8, 2, style="Table Grid")
tabla_p2.cell(0, 0).text = "Nombre Adquisición"
tabla_p2.cell(0, 1).text = "Suministro de Insumos y Accesorios para Terapia de Presión Negativa con Equipos en Comodato"
tabla_p2.cell(1, 0).text = "Descripción"
tabla_p2.cell(1, 1).text = "El Hospital requiere generar un convenio por el SUMINISTRO DE INSUMOS Y ACCESORIOS PARA TERAPIA DE PRESIÓN NEGATIVA CON EQUIPOS EN COMODATO PARA EL HOSPITAL SAN JOSÉ DE MELIPILLA, en adelante “EL HOSPITAL”. El convenio tendrá una vigencia de 36 meses."
tabla_p2.cell(2, 0).text = "Tipo de Convocatoria"
tabla_p2.cell(2, 1).text = "Abierta"
tabla_p2.cell(3, 0).text = "Moneda o Unidad reajustable"
tabla_p2.cell(3, 1).text = "Pesos Chilenos"
tabla_p2.cell(4, 0).text = "Presupuesto Referencial"
tabla_p2.cell(4, 1).text = "$350.000.000.- (Impuestos incluidos)"
tabla_p2.cell(5, 0).text = "Etapas del Proceso de Apertura"
tabla_p2.cell(5, 1).text = "Una Etapa (Etapa de Apertura Técnica y Etapa de Apertura Económica en una misma instancia)."
tabla_p2.cell(6, 0).text = "Opciones de pago"
tabla_p2.cell(6, 1).text = "Transferencia Electrónica"
tabla_p2.cell(7, 0).text = "Tipo de Adjudicación"
tabla_p2.cell(7, 1).text = "Adjudicación por la totalidad"
centrar_verticalmente_tabla(tabla_p2)

bases_p2_runs = doc.add_paragraph()
bases_p2_runs.add_run("* Presupuesto referencial:").underline = True
bases_p2_runs.add_run(" " + "El Hospital se reserva el derecho de aumentar, previo acuerdo entre las partes, hasta un 30% el presupuesto referencial estipulado en las presentes bases de licitación.")

# -- Definiciones --
bases_p2_definiciones = doc.add_paragraph()
bases_p2_definiciones.add_run("Definiciones").bold = True

bases_p2_definiciones_def_a = doc.add_paragraph(style = "List Number 3")
bases_p2_definiciones_def_a.add_run("Proponente u oferente:").bold = True
bases_p2_definiciones_def_a.add_run(" " + "El proveedor o prestador que participa en el proceso de licitación mediante la presentación de una propuesta, en la forma y condiciones establecidas en las Bases.")

bases_p2_definiciones_def_b = doc.add_paragraph(style = "List Number 3")
bases_p2_definiciones_def_b.add_run("Administrador o coordinador Externo del Contrato").bold = True
bases_p2_definiciones_def_b.add_run(" " + "Persona designada por el oferente adjudicado, quien actuará como contraparte ante el Hospital.")

bases_p2_definiciones_def_c = doc.add_paragraph(style = "List Number 3")
bases_p2_definiciones_def_c.add_run("Dias Hábiles:").bold = True
bases_p2_definiciones_def_c.add_run(" " + "Son todos los días de la semana, excepto los sábados, domingos y festivos.")

bases_p2_definiciones_def_d = doc.add_paragraph(style = "List Number 3")
bases_p2_definiciones_def_d.add_run("Días Corridos:").bold = True
bases_p2_definiciones_def_d.add_run(" " + "Son los días de la semana que se computan uno a uno en forma correlativa. Salvo que se exprese lo contrario, los plazos de días señalados en las presentes bases de licitación son días corridos. En caso que el plazo expire en días sábados, domingos o festivos se entenderá prorrogados para el día hábil siguiente.")

bases_p2_definiciones_def_e = doc.add_paragraph(style = "List Number 3")
bases_p2_definiciones_def_e.add_run("Administrador del Contrato y/o Referente Técnico: ").bold = True
bases_p2_definiciones_def_e.add_run(" " + " Es el funcionario designado por el Hospital para supervisar la correcta ejecución del contrato, solicitar órdenes de compra, validar prefacturas, gestionar multas y/o toda otra labor que guarde relación con la ejecución del contrato.")

bases_p2_definiciones_def_f = doc.add_paragraph(style = "List Number 3")
bases_p2_definiciones_def_f.add_run("Gestor de Contrato:").bold = True
bases_p2_definiciones_def_f.add_run(" " + "Es el funcionario a cargo de la ejecución del presente proceso de Licitación, desde la publicación de las Bases hasta la generación del contrato en formato documental, como la elaboración de ficha en la plataforma “gestor de contrato” en el portal de mercado público, además de ser el responsable de dar seguimiento y cumplimiento a los procesos y plazos establecidos. ")


# Tabla de definiciones de tipos de licitación
tabla_p2_licitaciones = doc.add_table(6, 3, style="Table Grid")

# Encabezados de la tabla
tabla_p2_licitaciones.cell(0, 0).text = "RANGO (en UTM)"
tabla_p2_licitaciones.cell(0, 1).text = "TIPO LICITACION PUBLICA"
tabla_p2_licitaciones.cell(0, 2).text = "PLAZO PUBLICACION EN DIAS CORRIDOS"

# Filas de datos
tabla_p2_licitaciones.cell(1, 0).text = "<100"
tabla_p2_licitaciones.cell(1, 1).text = "L1"
tabla_p2_licitaciones.cell(1, 2).text = "5"

tabla_p2_licitaciones.cell(2, 0).text = "<=100 y <1000"
tabla_p2_licitaciones.cell(2, 1).text = "LE"
tabla_p2_licitaciones.cell(2, 2).text = "10, rebajable a 5"

tabla_p2_licitaciones.cell(3, 0).text = "<=1000 y <2000"
tabla_p2_licitaciones.cell(3, 1).text = "LP"
tabla_p2_licitaciones.cell(3, 2).text = "20, rebajable a 10"

tabla_p2_licitaciones.cell(4, 0).text = "<=2000 y <5000"
tabla_p2_licitaciones.cell(4, 1).text = "LQ"
tabla_p2_licitaciones.cell(4, 2).text = "20, rebajable a 10"

tabla_p2_licitaciones.cell(5, 0).text = "<=5000"
tabla_p2_licitaciones.cell(5, 1).text = "LR"
tabla_p2_licitaciones.cell(5, 2).text = "30"
centrar_verticalmente_tabla(tabla_p2_licitaciones)

bases_p3 = doc.add_paragraph(style=list_style)
bases_p3.add_run("Etapas y plazos:").bold = True
aplicar_numeracion(bases_p3, num_id_bases_p1)

# Crear tabla de etapas y plazos con una fila adicional para el título
tabla_p3_plazos = doc.add_table(9, 3, style="Table Grid")

# Combinar las celdas de la primera fila para el título
tabla_p3_plazos.cell(0, 0).merge(tabla_p3_plazos.cell(0, 2))

# Agregar el título en la celda combinada
titulo_celda = tabla_p3_plazos.cell(0, 0)
titulo_celda.text = "VIGENCIA DE LA PUBLICACION 10 DIAS CORRIDOS"
# Hacer que el texto del título esté en negrita
for paragraph in titulo_celda.paragraphs:
    for run in paragraph.runs:
        run.bold = True
    # Centrar el título
    paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

# Fila 1 (ahora es la segunda fila en la tabla)
tabla_p3_plazos.cell(1, 0).text = "Consultas"
tabla_p3_plazos.cell(1, 1).text = "Hasta las 15:00 Horas de 4º (cuarto) día corrido de publicada la Licitación."
tabla_p3_plazos.cell(1, 2).text = "Deben ser ingresadas al portal www.mercadopublico.cl"

# Fila 2
tabla_p3_plazos.cell(2, 0).text = "Respuestas a Consultas"
tabla_p3_plazos.cell(2, 1).text = "Hasta las 17:00 Horas de 7º (séptimo) día corrido de publicada la Licitación."
tabla_p3_plazos.cell(2, 2).text = "Deben ser ingresadas al portal www.mercadopublico.cl"

# Fila 3
tabla_p3_plazos.cell(3, 0).text = "Aclaratorias"
tabla_p3_plazos.cell(3, 1).text = "Hasta 1 días corrido antes del cierre de recepción de ofertas."
tabla_p3_plazos.cell(3, 2).text = "Deben ser ingresadas al portal www.mercadopublico.cl"

# Fila 4
tabla_p3_plazos.cell(4, 0).text = "Recepción de ofertas"
tabla_p3_plazos.cell(4, 1).text = "Hasta las 17:00 Horas de 10º (décimo) día corrido de publicada la Licitación."
tabla_p3_plazos.cell(4, 2).text = "Deben ser ingresadas al portal www.mercadopublico.cl"

# Fila 5
tabla_p3_plazos.cell(5, 0).text = "Evaluación de las Ofertas"
tabla_p3_plazos.cell(5, 1).text = "Máximo 40 días corridos a partir del cierre de la Licitación."
tabla_p3_plazos.cell(5, 2).text = ""

# Fila 6
tabla_p3_plazos.cell(6, 0).text = "Plazo Adjudicaciones"
tabla_p3_plazos.cell(6, 1).text = "Máximo 20 días corridos a partir de la fecha del acta de evaluación de las ofertas."
tabla_p3_plazos.cell(6, 2).text = ""

# Fila 7
tabla_p3_plazos.cell(7, 0).text = "Suscripción de Contrato"
tabla_p3_plazos.cell(7, 1).text = "Máximo de 20 días hábiles desde la Adjudicación de la Licitación."
tabla_p3_plazos.cell(7, 2).text = ""

# Fila 8
tabla_p3_plazos.cell(8, 0).text = "Consideración"
tabla_p3_plazos.cell(8, 1).text = "Los plazos de días establecidos en la cláusula 3, Etapas y Plazos, son de días corridos, excepto el plazo para emitir la orden de compra, el que se considerará en días hábiles, entendiéndose que son inhábiles los sábados, domingos y festivos en Chile, sin considerar los feriados regionales."
tabla_p3_plazos.cell(8, 2).text = ""

centrar_verticalmente_tabla(tabla_p3_plazos)

# Guardar documento


doc_path = 'resolucion_numerada.docx'
doc.save(doc_path)