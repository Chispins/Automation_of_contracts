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


}

Datos_Contrato = {
    "espacio" : " ",
    "Tercero_DocumentosIntegrantes" : "Tercero",
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
    "VigesimoQuinto_Discrepancias": "Vigésimo Quinto"

}

doc = DocxTemplate(template_path)

# 2. Render the template with the context data
doc.render(Datos_Javi)
doc.render(Datos_Contrato)

# 3. Save the generated document
doc.save(output_path)

print(f"Report '{output_path}' generated successfully!")


