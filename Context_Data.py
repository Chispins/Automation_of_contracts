import jinja2
from docxtpl import DocxTemplate
import os
from Formated_Base_PEP8 import configurar_directorio_trabajo

configurar_directorio_trabajo()

# Define paths
template_path = "base_automatizada.docx"
output_path = "base_automatizada_jinja2.docx"

Datos_Javi = {
    "director" : "la Resolución Exenta RA 116395/343/2024 de fecha 12/08/2024 del SSMOCC., la cual nombra Director del Hospital San José de Melipilla al suscrito",
    "nombre_adquisicion" : "SUMINISTRO DE INSUMOS Y ACCESORIOS PARA TERAPIA DE PRESIÓN NEGATIVA CON EQUIPOS EN COMODATO PARA EL HOSPITAL SAN JOSÉ DE MELIPILLA",
    "cantidad_anexos": ", 6, 7, 8 y 9",
    "plazo_meses": "36",
    "presupuesto_con_impuestos": "$350.000.000",
    "tipo_adjudicacion": "Adjudicacion por la totalidad",
    "dias_vigencia_publicacion": "10",
    "plazo_consultas": "4º (cuarto)",
    "plazo_respuesta": "7º (séptimo)",
    "plazo_recepcion_ofertas": "10º (décimo)",
    "plazo_suscripcion": "20 días hábiles",
    "adjudicacion_corrido_habiles": "corridos",
    "atraso_para_multa_grave": "seis(6) días hábiles",
    "opciones_referente_tecnico_adm" : "(la) Enfermera Supervisora(o) del Servicio de Pabellón y al Jefe(a) de Farmacia o su subrogante ",
    "resolucion_empates" : "EVALUACION TECNICA, seguido por PLAZO DE ENTREGA, seguido por SERVICIO POST-VENTA, seguido por CRITERIO ECONOMICO",
    "anexos_tecnicos": "y los Anexos Técnicos N°7, N°8 y N°9",
    "metodo_adjudicacion": "la totalidad",
    "administrador_tecnico_administrativo": "la Enfermera Supervisora de Pabellón y el encargado en aspectos administrativos será el Jefe de Farmacia o quien lo subrogue."
    #"texto_extra" : "La adjudicación se realizará por",
    #"tipo_adjudicacion" : "valor unitario",


}

Datos_Base = {
    "coordinador" : "deberá nombrar un coordinador del contrato, cuya identidad deberá ser informada al Hospital."


}



Datos_Contrato = {
    "espacio" : " ",
    "Documentos_Integrantes" : "Tercero",
    "Cuarto_ModificacionDelContrato" : "Cuarto",
    "Quinto_GastoseImpuestos" : "Quinto",
    "Sexto_EfectosDerivadosDeIncumplimiento" : "Sexto",
    "Septimo_DeLaGarantíaFielCumplimiento": "Séptimo",
    "Octavo_CobroDeLaGarantiaFielCumplimiento": "Octavo",
    "Noveno_TerminoAnticipadoDelContrato": "Noveno",
    "Decimo_ResciliacionMutuoAcuerdo": "Décimo",
    "DecimoPrimero_ProcedimientoIncumplimiento": "Décimo Primero",
    "DecimoSegundo_EmisionOC": "Decimo Segundo",
    "DecimoTercero_DelPago": "Décimo Tercero",
    "DecimoCuarto_VigenciaContrato": "Décimo Cuarto",
    "DecimoQuinto_AdministradorContrato": "Décimo Quinto",
    "DecimoSexto_PactoDeIntegrida": "Décimo Sexto",
    "DecimoSeptimo_ComportamientoEticoAdjudic": "Décimo Séptimo",
    "DecimoOctavo_Auditorias": "Décimo Octavo",
    "DecimoNoveno_Confidencialidad": "DécimoNoveno",
    "Vigesimo_PropiedadDeLaInformacion": "Vigésimo",
    "VigesimoPrimero_SaldosInsolutos": "Vigésimo Primero",
    "VigesimoSegundo_NormasLaboralesAplicable": "Vigésimo Segundo",
    "VigesimoTercero_CambioPersonalProveedor": "Vigésimo Tercero",
    "VigesimoCuarto_CesionySubcontratacion": "Vigésimo Cuarto",
    "VigesimoQuinto_Discrepancias": "Vigésimo Quinto",

    "coordinador" : "El adjudicatario nombra coordinador del contrato a",
    "nombre_coordinador": "doña MARIA GABRIELA CARDENAS en el desempeño de su cometido, el coordinador del contrato deberá, a lo menos:",
    "monto_contrato_garantia" : "$3.250.000",
    "texto_gar_1" : ", es decir",
    "texto_gar_2" :" de pesos a nombre de “EL HOSPITAL” y consigna la siguiente glosa: Para garantizar el fiel cumplimiento del contrato denominado:",
    "texto_gar_3" : "ds",
    "id_licitacion" : "1057480-81-LE24",


}

doc = DocxTemplate(template_path)

# 2. Render the template with the context data
doc.render(Datos_Javi)

doc.render(Datos_Contrato)

# 3. Save the generated document
doc.save(output_path)

print(f"Report '{output_path}' generated successfully!")


# Ahora con el final
Datos_Contrato_Portada = {
    "numero_contrato": "4",
    "involucrados": "CRE/RMG/MMJ/MGL/MES",
    "fecha_contrato": "02 de enero de 2025",
    "nombre_proveedor": "MEDCORP S.A",
    "rut_proveedor": "76.131.542-0",
    "representante_legal": "doña Alejandra Ana Cuesta Nazar",
    "rut_representante_legal": "15.638.432-1",
    "domicilio_representante_legal": "San Fernando 1234, Santiago"
}
contrato_con_tablas = DocxTemplate("contrato_automatizado_tablas.docx")
# Usar la instancia 'contrato_con_tablas' para renderizar y guardar el documento
contrato_con_tablas.render(Datos_Contrato_Portada)
contrato_con_tablas.save("contrato_automatizado_portada_jinja2_segunda.docx")



Datos_Contrato_Base = {
    "Documentos Integrantes": "Documentos Integrantes",
    "Cuarto_ModificacionDelContrato": "Cuarto_ModificacionDelContrato",
    "Quinto_GastoseImpuestos": "Quinto_GastoseImpuestos",
    "Sexto_EfectosDerivadosDeIncumplimiento": "Sexto_EfectosDerivadosDeIncumplimiento",
    "Septimo_DeLaGarantíaFielCumplimiento": "Septimo_DeLaGarantíaFielCumplimiento",
    "Octavo_CobroDeLaGarantiaFielCumplimiento": "Octavo_CobroDeLaGarantiaFielCumplimiento",
    "Noveno_TerminoAnticipadoDelContrato": "Noveno_TerminoAnticipadoDelContrato",
    "Decimo_ResciliacionMutuoAcuerdo": "Decimo_ResciliacionMutuoAcuerdo",
    "DecimoPrimero_ProcedimientoIncumplimiento": "DecimoPrimero_ProcedimientoIncumplimiento",
    "DecimoSegundo_EmisionOC": "DecimoSegundo_EmisionOC",
    "DecimoTercero_DelPago": "DecimoTercero_DelPago",
    "DecimoCuarto_VigenciaContrato": "DecimoCuarto_VigenciaContrato",
    "DecimoQuinto_AdministradorContrato": "DecimoQuinto_AdministradorContrato",
    "DecimoSexto_PactoDeIntegrida": "DecimoSexto_PactoDeIntegrida",
    "DecimoSeptimo_ComportamientoEticoAdjudic": "DecimoSeptimo_ComportamientoEticoAdjudic",
    "DecimoOctavo_Auditorias": "DecimoOctavo_Auditorias",
    "DecimoNoveno_Confidencialidad": "DecimoNoveno_Confidencialidad",
    "Vigesimo_PropiedadDeLaInformacion": "Vigesimo_PropiedadDeLaInformacion",
    "VigesimoPrimero_SaldosInsolutos": "VigesimoPrimero_SaldosInsolutos",
    "VigesimoSegundo_NormasLaboralesAplicable": "VigesimoSegundo_NormasLaboralesAplicable",
    "VigesimoTercero_CambioPersonalProveedor": "VigesimoTercero_CambioPersonalProveedor",
    "VigesimoCuarto_CesionySubcontratacion": "VigesimoCuarto_CesionySubcontratacion",
    "VigesimoQuinto_Discrepancias": "VigesimoQuinto_Discrepancias"
}


import pandas as pd
excel_name = "Libro1.xlsx"

Datos_Base_excel = pd.read_excel(excel_name, sheet_name="Datos_Base")
base_nomb = Datos_Base_excel.iloc[:, 0]
base_base = Datos_Base_excel.iloc[:, 1]
diccion = [f"{nombre}: {base}" for nombre, base in zip(base_nomb, base_base)]



Datos_Contrato_P1 = pd.read_excel(excel_name, sheet_name="Datos_Contrato_P1")
# Now the same
contrato_nomb = Datos_Contrato_P1.iloc[:, 0]
contrato_base = Datos_Contrato_P1.iloc[:, 1]
diccion = [f"{nombre}: {base}" for nombre, base in zip(contrato_nomb, contrato_base)]




Datos_Contrato_P2 = pd.read_excel(excel_name, sheet_name="Datos_Contrato_P2")
# Now the same
contrato_nomb_p2 = Datos_Contrato_P2.iloc[:, 0]
contrato_base_p2 = Datos_Contrato_P2.iloc[:, 1]
diccion = [f"{nombre}: {base}" for nombre, base in zip(contrato_nomb_p2, contrato_base_p2)]


# Now the jinja with the the diccionaries of the first two things

templ = "base_automatizada.docx"
templ_output_1 = ""