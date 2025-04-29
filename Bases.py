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

doc.add_heading("Consultas, Aclaraciones y modificaciones a las bases.", level=2)

# Párrafo 1
p_consultas1 = doc.add_paragraph()
p_consultas1.add_run("Las consultas de los participantes se deberán realizar únicamente a través del portal")
p_consultas1.add_run(" www.mercadopublico.cl").bold = True
p_consultas1.add_run(" conforme el cronograma de actividades de esta licitación señalado en el punto 3 precedente. A su vez, las respuestas y aclaraciones estarán disponibles a través del portal de Mercado Público, en los plazos indicados en el cronograma señalado precedentemente, información que se entenderá conocida por todos los interesados desde el momento de su publicación.")
p_consultas1.style = 'List Bullet' # Apply the bullet style

# Párrafo 2 (now a list item)
p_consultas2 = doc.add_paragraph()
p_consultas2.add_run("No serán admitidas las consultas formuladas fuera de plazo o por un conducto diferente al señalado.").bold = True
p_consultas2.style = 'List Bullet' # Apply the bullet style

# Párrafo 3 (now a list item)
p_consultas3 = doc.add_paragraph()
p_consultas3.add_run("“EL HOSPITAL” realizará las aclaraciones a las Bases comunicando las respuestas a través del Portal Web de Mercado Público, sitio")
p_consultas3.add_run(" www.mercadopublico.cl").bold = True
p_consultas3.style = 'List Bullet' # Apply the bullet style

# Párrafo 4 (now a list item)
p_consultas4 = doc.add_paragraph()
p_consultas4.add_run("Las aclaraciones, derivadas de este proceso de consultas, formarán parte integrante de las Bases, teniéndose por conocidas y aceptadas por todos los participantes aun cuando el oferente no las hubiere solicitado, por lo que los proponentes no podrán alegar desconocimiento de las mismas.")
p_consultas4.style = 'List Bullet' # Apply the bullet style

# Párrafo 5 (now a list item)
p_consultas5 = doc.add_paragraph()
p_consultas5.add_run("“EL HOSPITAL” podrá modificar las presentes bases y sus anexos previa autorización por acto administrativo, durante el periodo de presentación de las ofertas, hasta antes de fecha de cierre de recepción de ofertas. Estas modificaciones, que se llewen a cabo, serán informadas a través del portal")
p_consultas5.add_run(" www.mercadopublico.cl").bold = True
p_consultas5.style = 'List Bullet' # Apply the bullet style

# Párrafo 6 (now a list item)
p_consultas6 = doc.add_paragraph()
p_consultas6.add_run("Estas consultas, aclaratorias y modificaciones formaran parte integra de las bases y estarán vigentes desde la total tramitación del acto administrativo que las apruebe. Junto con aprobar las modificaciones, deberá establecer un nuevo plazo prudencial cuando lo amerite para el cierre o recepción de las propuestas, a fin de que los potenciales oferentes puedan adecuar sus ofertas.")
p_consultas6.style = 'List Bullet' # Apply the bullet style

# Párrafo 7 (now a list item)
p_consultas7 = doc.add_paragraph()
p_consultas7.add_run("No se aceptarán consultas realizadas por otros medios, tales como correos electrónicos, fax u otros.")
p_consultas7.style = 'List Bullet' # Apply the bullet style

# Requisitos Participación
# Crear IDs para numeración
num_id_vistos = crear_numeracion(doc)
num_id_resolucion = crear_numeracion(doc)
num_id_bases_p1 = crear_numeracion(doc)
# Crear un nuevo ID para la sección de consultas
num_id_consultas = crear_numeracion(doc)
# Crear un nuevo ID para la sección de requisitos mínimos
num_id_requisitos = crear_numeracion(doc)

# ... (código anterior para VISTOS, CONSIDERANDO, RESOLUCIÓN, BASES ADMINISTRATIVAS, CONSULTAS) ...

# Requisitos Participación
doc.add_heading("Requisitos Mínimos para Participar.", level=3)

# Párrafos numerados para Requisitos Mínimos
req_p1 = doc.add_paragraph(
    "No haber sido condenado por prácticas antisindicales, infracción a los derechos fundamentales del trabajador o por delitos concursales establecidos en el Código Penal dentro de los dos últimos años anteriores a la fecha de presentación de la oferta, de conformidad con lo dispuesto en el artículo 4° de la ley N° 19.886."
)
aplicar_numeracion(req_p1, num_id_requisitos) # Aplicar numeración

req_p2 = doc.add_paragraph(
    "No haber sido condenado por el Tribunal de Defensa de la Libre Competencia a la medida dispuesta en la letra d) del artículo 26 del Decreto con Fuerza de Ley N°1, de 2004, del Ministerio de Economía, Fomento y Reconstrucción, que Fija el texto refundido, coordinado y sistematizado del Decreto Ley N° 211, de 1973, que fija normas para la defensa de la libre competencia, hasta por el plazo de cinco años contado desde que la sentencia definitiva quede ejecutoriada."
)
aplicar_numeracion(req_p2, num_id_requisitos) # Aplicar numeración

req_p3 = doc.add_paragraph(
    "No ser funcionario directivo de la respectiva entidad compradora; o una persona unida a aquél por los vínculos de parentesco descritos en la letra b) del artículo 54 de la ley N° 18.575; o una sociedad de personas de las que aquél o esta formen parte; o una sociedad comandita por acciones o anónima cerrada en que aquélla o esta sea accionista; o una sociedad anónima abierta en que aquél o esta sean dueños de acciones que representen el 10% o más del capital; o un gerente, administrador, representante o director de cualquiera de las sociedades antedichas."
)
aplicar_numeracion(req_p3, num_id_requisitos) # Aplicar numeración

req_p4 = doc.add_paragraph(
    "Tratándose exclusivamente de una persona jurídica, no haber sido condenada conforme a la ley N° 20.393 a la pena de prohibición de celebrar actos y contratos con el Estado, mientras esta pena esté vigente."
)
aplicar_numeracion(req_p4, num_id_requisitos) # Aplicar numeración

req_p5 = doc.add_paragraph("A fin de acreditar el cumplimiento de dichos requisitos, los oferentes deberán presentar una “Declaración jurada de requisitos para ofertar”, la cual será generada completamente en línea a través de www.mercadopublico.cl en el módulo de presentación de las ofertas. Sin perjuicio de lo anterior, la entidad licitante podrá verificar la veracidad de la información entregada en la declaración, en cualquier momento, a través de los medios oficiales disponibles.")
aplicar_numeracion(req_p5, num_id_requisitos) # Aplicar numeración

req_p6 = doc.add_paragraph()
req_p6.add_run("En caso de que los antecedentes administrativos solicitados en esta sección no sean entregados y/o completados en forma correcta y oportuna, se desestimará la propuesta, no será evaluada y será declarada ")
req_p6.add_run("inadmisible").bold = True
req_p6.add_run(".")
aplicar_numeracion(req_p6, num_id_requisitos) # Aplicar numeración






from docx import Document

# Assume 'doc' is an existing Document object, like:
# doc = Document()

import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from Bases import configurar_directorio_trabajo, centrar_verticalmente_tabla # Assuming these are defined in Bases.py

# Configure working directory (if needed)
configurar_directorio_trabajo()

# Assume 'doc' is an existing Document object or create a new one
# doc = Document() # Uncomment if you need a new document

# Add the heading before the table
doc.add_heading("Instrucciones para la Presentación de Ofertas.", level=2)

# Create a table with exactly 4 rows and 2 columns
table = doc.add_table(rows=4, cols=2)
table.style = 'Table Grid' # Optional: Add grid lines

# --- Row 1: Header ---
cell_r1_c1 = table.cell(0, 0)
cell_r1_c2 = table.cell(0, 1)
cell_r1_c1.text = "Presentar Ofertas por Sistema."
cell_r1_c2.text = "Obligatorio."
# Optional: Center header text if needed
# cell_r1_c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
# cell_r1_c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# --- Row 2: Anexos Administrativos ---
# --- Row 2: Anexos Administrativos ---
cell_r2_c1 = table.cell(1, 0)
cell_r2_c2 = table.cell(1, 1)
cell_r2_c1.text = "Anexos Administrativos."

# Add content to the right cell (Row 2, Col 2)
# Clear the default paragraph if it exists and is empty
if cell_r2_c2.paragraphs and not cell_r2_c2.paragraphs[0].text:
    p_element = cell_r2_c2.paragraphs[0]._p
    cell_r2_c2._element.remove(p_element)


# Item 1: Anexo N° 1
p1 = cell_r2_c2.add_paragraph()
p1.add_run("Anexo N° 1 Identificación del Oferente.").bold = True

# Item 2: Anexo N° 2
p2 = cell_r2_c2.add_paragraph()
p2.add_run("Anexo N° 2 Declaración Jurada de Habilidad.").bold = True

# Item 3: Anexo N° 3 (CORREGIDO)
p3 = cell_r2_c2.add_paragraph() # Crear el párrafo primero
p3.add_run("Anexo N° 3 Declaración Jurada de Cumplimiento de Obligaciones Laborales y Previsionales.").bold = True # Añadir run y poner en negrita

# Item 4: Declaración jurada online
p4 = cell_r2_c2.add_paragraph()
p4.add_run("Declaración jurada online:").bold = True
p4.add_run(" Los oferentes deberán presentar una ") # Espacio añadido al inicio
p4.add_run("Declaración jurada de requisitos para ofertar").bold = True
p4.add_run(", la cual será generada completamente en línea a través de www.mercadopublico.cl en el módulo de presentación de las ofertas.") # Coma y espacio añadidos

# Item 5: Unión Temporal de Proveedores (UTP)
p5 = cell_r2_c2.add_paragraph()
p5.add_run("Unión Temporal de Proveedores (UTP):").bold = True
p5.add_run(" Solo en el caso de que la oferta sea presentada por una unión temporal de proveedores deberán presentar obligatoriamente la siguiente documentación en su totalidad, en caso contrario, ésta no será sujeta a aclaración y la oferta será declarada ")
p5.add_run("inadmisible").bold = True
p5.add_run(".") # Punto añadido

# Item 6: Anexo N°4 for UTP
p6 = cell_r2_c2.add_paragraph()
p6.add_run("Anexo N°4. Declaración para Uniones Temporales de Proveedores:").bold = True
p6.add_run(" Debe ser presentado por el miembro de la UTP que presente la oferta en el Sistema de Información y quien realiza la declaración a través de la “Declaración jurada de requisitos para ofertar” electrónica presentada junto a la oferta.") # Espacio añadido al inicio

# Item 7: UTP Offers and Apoderado
p7 = cell_r2_c2.add_paragraph()
p7.add_run("Las ofertas presentadas por una Unión Temporal de Proveedores (UTP) deberán contar con un apoderado, el cual debe corresponder a un integrante de la misma, ya sea persona natural o jurídica. En el caso que el apoderado sea una persona jurídica, ésta deberá actuar a través de su representante legal para ejercer sus facultades.")

# Item 8: Inadmissibility condition
p8 = cell_r2_c2.add_paragraph()
p8.add_run("En caso de no presentarse debidamente la declaración jurada online constatando la ausencia de conflictos de interés e inhabilidades por condenas, o no presentarse el Anexo N°4, la oferta será declarada ")
p8.add_run("inadmisible").bold = True
p8.add_run(".") # Punto añadido

# --- Row 3: Anexos Económicos ---
cell_r3_c1 = table.cell(2, 0)
cell_r3_c2 = table.cell(2, 1)
cell_r3_c1.text = "Anexos\nEconómicos." # Using \n for line break in the cell text

# Add content to the right cell (Row 3, Col 2)
# Clear the default paragraph if it exists and is empty
if cell_r3_c2.paragraphs and not cell_r3_c2.paragraphs[0].text:
    p_element = cell_r3_c2.paragraphs[0]._p
    cell_r3_c2._element.remove(p_element)

# Item 1: Anexo N°5
ep1 = cell_r3_c2.add_paragraph()
ep1.add_run("Anexo N°5: Oferta económica").bold = True

# Item 2: Entering via system
ep2 = cell_r3_c2.add_paragraph()
ep2.add_run("El anexo referido debe ser ingresado a través del sistema www.mercadopublico.cl , en la sección Anexos Económicos.")

# Item 3: Inadmissibility condition
ep3 = cell_r3_c2.add_paragraph()
ep3.add_run("En caso de que no se presente debidamente el Anexo N°5 “Oferta económica”, la oferta será declarada ")
ep3.add_run("inadmisible").bold = True

# --- Row 4: Anexos Técnicos ---
cell_r4_c1 = table.cell(3, 0)
cell_r4_c2 = table.cell(3, 1)
cell_r4_c1.text = "Anexos Técnicos." # Using \n for line break

# Add content to the right cell (Row 4, Col 2)
# Clear the default paragraph if it exists and is empty
if cell_r4_c2.paragraphs and not cell_r4_c2.paragraphs[0].text:
    p_element = cell_r4_c2.paragraphs[0]._p
    cell_r4_c2._element.remove(p_element)

# Item 1: Anexo N°6
tp1 = cell_r4_c2.add_paragraph()
tp1.add_run("Anexo N°6: Evaluación Técnica").bold = True

# Item 2: Anexo N°7
tp2 = cell_r4_c2.add_paragraph()
tp2.add_run("Anexo N°7: Ficha Técnica").bold = True

# Item 3: Anexo N°8
tp3 = cell_r4_c2.add_paragraph()
tp3.add_run("Anexo N°8: Plazo de Entrega").bold = True

# Item 4: Anexo N°9
tp4 = cell_r4_c2.add_paragraph()
tp4.add_run("Anexo N°9: Servicio Post-venta").bold = True

# Item 5: Entering via system
tp5 = cell_r4_c2.add_paragraph()
tp5.add_run("Los anexos referidos deben ser ingresados a través del sistema www.mercadopublico.cl. en la sección Anexos Técnicos.")

# Item 6: Inadmissibility condition
tp6 = cell_r4_c2.add_paragraph()
tp6.add_run("En el caso que no se presente debidamente los Anexos N°7, N°8 y N°9 la oferta será declarada ")
tp6.add_run("inadmisible").bold = True

# Apply vertical centering to the entire table (optional, requires your function)
centrar_verticalmente_tabla(table)

# Save the document
doc.add_heading("Observaciones", level = 3)
parrafos_observaciones = doc.add_paragraph()
parrafos_observaciones.add_run("Los oferentes deberán presentar su oferta a través de su cuenta en el Sistema de Información www.mercadopublico.cl. De existir discordancia entre el oferente o los antecedentes de su oferta y la cuenta a través de la cual la presenta, esta no será evaluada, siendo desestimada del proceso y declarada como")
parrafos_observaciones.add_run(" " + "inadmisible").bold = True


p1 = doc.add_paragraph()
p1.add_run("Las únicas ofertas válidas serán las presentadas a través del portal")
p1.add_run(" www.mercadopublico.cl").bold = True
p1.add_run(", en la forma en que se solicita en estas bases. No se aceptarán ofertas que se presenten por un medio distinto al establecido en estas Bases, a menos que se acredite la indisponibilidad técnica del sistema, de conformidad con el artículo 62 del Reglamento de la Ley de Compras. Será responsabilidad de los oferentes adoptar las precauciones necesarias para ingresar oportuna y adecuadamente sus ofertas.")

# Paragraph 2
p2 = doc.add_paragraph()
p2.add_run("Los oferentes deben constatar que el envío de su oferta a través del portal electrónico de compras públicas haya sido realizado con éxito, incluyendo el previo ingreso de todos los formularios y anexos requeridos completados de acuerdo con lo establecido en las presentes bases. Debe verificar que los archivos que se ingresen contengan efectivamente los anexos solicitados.")

# Paragraph 3
p3 = doc.add_paragraph()
p3.add_run("Asimismo, se debe comprobar siempre, luego de que se finalice la última etapa de ingreso de la oferta respectiva, que se produzca el despliegue automático del “Comprobante de Envío de Oferta” que se entrega en dicho Sistema, el cual puede ser impreso por el proponente para su resguardo. En dicho comprobante será posible visualizar los anexos adjuntos, cuyo contenido es de responsabilidad del oferente.")

# Paragraph 4
p4 = doc.add_paragraph()
p4.add_run("El hecho de que el oferente haya obtenido el “Comprobante de envío de ofertas” señalado, únicamente acreditará el envío de ésta a través del Sistema, pero en ningún caso certificará la integridad o la completitud de ésta, lo cual será evaluado por la comisión evaluadora. En caso de que, antes de la fecha de cierre de la licitación, un proponente edite una oferta ya enviada, deberá asegurarse de enviar nuevamente la oferta una vez haya realizado los ajustes que estime, debiendo descargar un nuevo Comprobante.")

# Paragraph 5
p5 = doc.add_paragraph()
p5.add_run("Si la propuesta económica subida al portal, presenta diferencias entre el valor del anexo económico solicitado y el valor indicado en la línea de la plataforma")
p5.add_run(" www.mercadopublico.cl").bold = True
p5.add_run(", prevalecerá la oferta del anexo económico solicitado en bases. Sin embargo, el Hospital San José de Melipilla, podrá solicitar aclaraciones de las ofertas realizadas a través del portal.")


doc.add_heading("Antecedentes legales para poder ser contratado.", level=2)


table = doc.add_table(rows=7, cols=3)
table.style = 'Table Grid' # Apply grid lines style

# Optional: Adjust column widths for better layout
# table.columns[0].width = Inches(1.5)
# table.columns[1].width = Inches(4.0)
# table.columns[2].width = Inches(1.5)

# --- Section: Si el oferente es Persona Natural (Rows 0-3) ---
start_natural_row = 0
end_natural_row = 3 # Rows 0, 1, 2, 3 (4 rows total)

# Cell (0, 0): "Si el oferente es Persona Natural" - Merged
cell_0_0 = table.cell(start_natural_row, 0)
# Use paragraphs and runs for better control over potential future formatting or line breaks
p_0_0 = cell_0_0.paragraphs[0]
p_0_0.add_run("Si el oferente\nes Persona\nNatural")

# Cell (0, 2): "Acreditar en el Registro de Proveedores" - Merged
cell_0_2 = table.cell(start_natural_row, 2)
p_0_2 = cell_0_2.paragraphs[0]
p_0_2.add_run("Acreditar en\nel Registro de\nProveedores")

# Populate middle column (Column 1) for Persona Natural section
# Cell (0, 1): Requirement 1
cell_0_1 = table.cell(start_natural_row, 1)
p_0_1 = cell_0_1.paragraphs[0]
# Text is bold in the image for "Inscripción..." and "Registro..."
p_0_1.add_run("Inscripción (en estado hábil) en el Registro electrónico oficial de contratistas de la Administración, en adelante “").bold = True
p_0_1.add_run("Registro de Proveedores").bold = True
p_0_1.add_run("”.").bold = True # The closing quote and period also appear bold


# Cell (1, 1): Requirement 2 - Anexo N°3
cell_1_1 = table.cell(start_natural_row + 1, 1)
p_1_1 = cell_1_1.paragraphs[0]
p_1_1.add_run("Anexo N°3. Declaración Jurada de Cumplimiento de Obligaciones Laborales y Previsionales.").bold = True # Whole line appears bold

# Cell (2, 1): Requirement 3 - Todos los Anexos...
cell_2_1 = table.cell(start_natural_row + 2, 1)
cell_2_1.text = "Todos los Anexos deben ser firmados por la persona natural respectiva."

# Cell (3, 1): Requirement 4 - Fotocopia...
cell_3_1 = table.cell(start_natural_row + 3, 1)
cell_3_1.text = "Fotocopia de su cédula de identidad."

# Perform merges for the Persona Natural section
cell_0_0.merge(table.cell(end_natural_row, 0)) # Merge column 0 (rows 0 to 3)
cell_0_2.merge(table.cell(end_natural_row, 2)) # Merge column 2 (rows 0 to 3)


# --- Section: Si el oferente no es Persona Natural (Rows 4-6) ---
start_nonatural_row = end_natural_row + 1 # Starts at row 4
end_nonatural_row = start_nonatural_row + 2 # Rows 4, 5, 6 (3 rows total)

# Cell (4, 0): "Si el oferente no es Persona Natural" - Merged
cell_4_0 = table.cell(start_nonatural_row, 0)
p_4_0 = cell_4_0.paragraphs[0]
p_4_0.add_run("Si el oferente\nno es\nPersona\nNatural")

# Cell (4, 2): "Acreditar en el Registro de Proveedores" - Merged
cell_4_2 = table.cell(start_nonatural_row, 2)
p_4_2 = cell_4_2.paragraphs[0]
p_4_2.add_run("Acreditar en\nel Registro de\nProveedores")

# Populate middle column (Column 1) for Persona no Natural section
# Cell (4, 1): Requirement 1 - Inscripción...
cell_4_1 = table.cell(start_nonatural_row, 1)
p_4_1 = cell_4_1.paragraphs[0]
p_4_1.add_run("Inscripción (en estado hábil) en el Registro de Proveedores.").bold = True # Whole line appears bold

# Cell (5, 1): Requirement 2 - Certificado de Vigencia...
cell_5_1 = table.cell(start_nonatural_row + 1, 1)
cell_5_1.text = "Certificado de Vigencia del poder del representante legal, con una antigüedad no superior a 60 días corridos, contados desde la fecha de notificación de la adjudicación, otorgado por el Conservador de Bienes Raíces correspondiente o, en los casos que resulte procedente, cualquier otro antecedente que acredite la vigencia del poder del representante legal del oferente, a la época de presentación de la oferta."

# Cell (6, 1): Requirement 3 - Certificado de Vigencia de la Sociedad...
cell_6_1 = table.cell(start_nonatural_row + 2, 1)
cell_6_1.text = "Certificado de Vigencia de la Sociedad con una antigüedad no superior a 60 días corridos, contados desde la fecha de notificación de la" # Text cut off in image

# Perform merges for the Persona no Natural section
cell_4_0.merge(table.cell(end_nonatural_row, 0)) # Merge column 0 (rows 4 to 6)
cell_4_2.merge(table.cell(end_nonatural_row, 2)) # Merge column 2 (rows 4 to 6)


# Observaciones
doc.add_heading("Observaciones", level = 3)
doc.add_paragraph("Los antecedentes legales para poder ser contratado, sólo se requerirán respecto del adjudicatario y deberán estar disponibles en el Registro de Proveedores.")
doc.add_paragraph("Lo señalado en el párrafo precedente no resultará aplicable a la garantía de fiel cumplimiento de contrato, la cual podrá ser entregada físicamente en los términos que indican las presentes bases en aquellos casos que aplique su entrega.")
doc.add_paragraph("En los casos en que se otorgue de manera electrónica, deberá ajustarse a la ley N° 19.799 sobre documentos electrónicos, firma electrónica y servicios de certificación de dicha firma, y remitirse en la forma señalada en la cláusula 8.2 de estas bases.")
observ_parafo_2 = doc.add_paragraph()
observ_parafo_2.add_run("Si el respectivo proveedor no entrega la totalidad de los antecedentes requeridos para ser contratado, dentro del plazo fatal de 10 días hábiles administrativos contados desde la notificación de la resolución de adjudicación o no suscribe el contrato en los plazos establecidos en estas bases, la entidad licitante podrá readjudicar de conformidad a lo establecido en la")
observ_parafo_2.add_run(" " + "cláusula 9 letra i")
observ_parafo_2.add_run(" de las presentes bases. Además, tales incumplimientos darán origen al cobro de la garantía de seriedad de la oferta, si la hubiere.")

doc.add_heading("Inscripción en el registro de proveedores", level =2 )
doc.add_paragraph("En caso de que el proveedor que resulte adjudicado no se encuentre inscrito en el Registro Electrónico Oficial de Contratistas de la Administración (Registro de Proveedores), deberá inscribirse dentro del plazo de 15 días hábiles, contados desde la notificación de la resolución de adjudicación.")
doc.add_paragraph("Tratándose de los adjudicatarios de una Unión Temporal de Proveedores, cada integrante de ésta deberá inscribirse en el Registro de Proveedores, dentro del plazo de 15 días hábiles, contados desde la notificación de la resolución de adjudicación. ")

doc.add_heading("Naturaleza y monto de las garatías", level = 2)
doc.add_heading("Evaluación y adjudicación de las ofertas", level = 3)


# Nueva Lista 4
num_id_evaluacion = crear_numeracion(doc)
eval_p1 = doc.add_paragraph("test", style='List Number')
aplicar_numeracion(eval_p1, num_id_evaluacion)

# Elemento 2 de la lista
eval_p2 = doc.add_paragraph("La comisión evaluadora verificará el cumplimiento de los requisitos mínimos de participación.", style='List Number')
aplicar_numeracion(eval_p2, num_id_evaluacion)

# Elemento 3 de la lista
eval_p3 = doc.add_paragraph("Se evaluarán los criterios técnicos y económicos según la ponderación definida en las bases.", style='List Number')
aplicar_numeracion(eval_p3, num_id_evaluacion)

# Elemento 4 de la lista con negrita
eval_p4 = doc.add_paragraph(style='List Number')
eval_p4.add_run("La adjudicación se realizará al oferente que obtenga el ")
eval_p4.add_run("mayor puntaje total").bold = True
eval_p4.add_run(".")
aplicar_numeracion(eval_p4, num_id_evaluacion)





doc_path = 'resolucion_numerada.docx' # Or your desired output file name
doc.save(doc_path)
print(f"Documento guardado como: {doc_path}")