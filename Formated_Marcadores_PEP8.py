import docx
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
import os
import re
from Formated_Base_PEP8 import configurar_directorio_trabajo


def obtener_marcadores(documento):
    """Obtiene todos los marcadores de un documento Word."""
    marcadores = {}
    for elemento in documento._element.xpath('//w:bookmarkStart'):
        nombre_marcador = elemento.get(qn('w:name'))
        id_marcador = elemento.get(qn('w:id'))
        if nombre_marcador:
            marcadores[nombre_marcador] = {
                'id': id_marcador,
                'elemento': elemento
            }
    return marcadores


def modificar_texto_marcador(documento, nombre_marcador, nuevo_texto):
    """Modifica el texto contenido en un marcador eliminando y reemplazando el contenido."""
    marcadores = obtener_marcadores(documento)
    if nombre_marcador not in marcadores:
        print(f"No se encontró el marcador '{nombre_marcador}'. No se puede modificar.")
        return False

    elemento_inicio = marcadores[nombre_marcador]['elemento']
    id_marcador = marcadores[nombre_marcador]['id']
    elemento_fin = None
    for elem in documento._element.xpath(f'//w:bookmarkEnd[@w:id="{id_marcador}"]'):
        elemento_fin = elem
        break

    if elemento_fin is None:
        print(f"Marcador '{nombre_marcador}' incompleto (falta el final)")
        return False

    current_elem = elemento_inicio
    elementos_a_eliminar = []
    while current_elem is not elemento_fin:
        current_elem = current_elem.getnext()
        if current_elem is None:
            print(f"No se pudo encontrar el contenido entre inicio y fin para '{nombre_marcador}'")
            return False
        if current_elem.tag.endswith('r') or current_elem.tag.endswith('t'):
            elementos_a_eliminar.append(current_elem)

    for elem in elementos_a_eliminar:
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)

    nuevo_texto_elem = OxmlElement('w:r')
    texto_elem = OxmlElement('w:t')
    texto_elem.text = nuevo_texto
    nuevo_texto_elem.append(texto_elem)

    try:
        elemento_inicio.addnext(nuevo_texto_elem)
        print(f"Texto reemplazado correctamente para el marcador '{nombre_marcador}'")
        return True
    except Exception as e:
        print(f"Error al insertar texto en '{nombre_marcador}': {e}")
        return False

def leer_texto_marcador(documento, nombre_marcador):
    marcadores = obtener_marcadores(documento)
    if nombre_marcador not in marcadores:
        return ""
    inicio = marcadores[nombre_marcador]['elemento']
    id_marcador = marcadores[nombre_marcador]['id']
    fin = documento._element.xpath(f'//w:bookmarkEnd[@w:id="{id_marcador}"]')
    if not fin:
        return ""
    fin = fin[0]
    texto = []
    actual = inicio
    while actual is not fin:
        actual = actual.getnext()
        if actual is None:
            break
        if actual.tag.endswith('t'):
            texto.append(actual.text or "")
    return "".join(texto)


def main():
    configurar_directorio_trabajo()
    doc_path = "contrato_automatizado_con_marcadores.docx"
    output_path = "contrato_automatizado_con_marcadores_testing_v1.docx"
    print(f"Cargando documento: {doc_path}")

    try:
        doc = docx.Document(doc_path)
    except Exception as e:
        print(f"Error al cargar el documento: {e}")
        return

    marcadores = obtener_marcadores(doc)
    print("Marcadores disponibles:")
    for nombre in marcadores:
        print(f"- {nombre}")

    elements = {
        "Tercero_DocumentosIntegrantes": "Tercero",
        "Cuarto_ModificacionDelContrato": "Cuarto",
        # ... resto de elementos ...
        "VigesimoQuinto_Discrepancias": "VigesimoQuinto"
    }

    print("\nIniciando modificaciones de marcadores...")
    modificados = 0
    for nombre, valor in elements.items():
        # Reemplaza completamente el texto del marcador por el valor del diccionario
        if modificar_texto_marcador(doc, nombre, valor):
            modificados += 1

    if modificados:
        try:
            doc.save(output_path)
            print(f"\nDocumento guardado como {output_path} con {modificados} marcador(es) modificado(s).")
        except Exception as e:
            print(f"\nError al guardar documento: {e}")
    else:
        print("\nNo se realizaron modificaciones.")

if __name__ == "__main__":
    main()



elements = {"Tercero_DocumentosIntegrantes": "Tercero", "Cuarto_ModificacionDelContrato": "Cuarto",
                "Quinto_GastoseImpuestos": "Quinto", "Sexto_EfectosDerivadosDeIncumplimientos": "Sexto",
                "Septimo_DeLaGarantíaFielCumplimiento": "Septimo", "Octavo_CobroDeLaGarantiaFielCumplimiento": "Octavo",
                "Noveno_TerminoAnticipadoDelContrato": "Noveno", "Decimo_ResciliacionMutuoAcuerdo": "Decimo",
                "DecimoPrimero_ProcedimientoIncumplimient": "DecimoPrimero", "DecimoSegundo_EmisionOC": "DecimoSegundo",
                "DecimoTercero_DelPago": "DecimoTercero", "DecimoCuarto_VigenciaContrato": "DecimoCuarto",
                "DecimoQuinto_AdministradorContrato": "DecimoQuinto", "DecimoSexto_PactoDeIntegrida": "DecimoSexto",
                "DecimoSeptimo_ComportamientoEticoAdjudic": "DecimoSeptimo", "DecimoOctavo_Auditorias": "DecimoOctavo",
                "DecimoNoveno_Confidencialidad": "DecimoNoveno", "Vigesimo_PropiedadDeLaInformacion": "Vigesimo",
                "VigesimoPrimero_SaldosInsolutos": "VigesimoPrimero",
                "VigesimoSegundo_NormasLaboralesAplicable": "VigesimoSegundo",
                "VigesimoTercero_CambioPersonalProveedor": "VigesimoTercero",
                "VigesimoCuarto_CesionySubcontratacion": "VigesimoCuarto",
                "VigesimoQuinto_Discrepancias": "VigesimoQuinto"}