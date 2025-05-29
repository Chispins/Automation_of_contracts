import pandas as pd
import os
import re


def strip_dataframe_and_handle_empty(df):
    """Reemplaza NaN con cadenas vacías y elimina espacios en blanco."""
    df = df.fillna('')
    df = df.astype(str)
    return df.applymap(lambda x: x.strip())


def process_and_clean_dict(df, sheet_name):
    """
    Procesa un DataFrame en un diccionario, limpiando claves y valores.

    Args:
        df (pd.DataFrame): DataFrame con datos.
        sheet_name (str): Nombre de la hoja para depuración.

    Returns:
        dict: Diccionario limpio con claves y valores procesados.
    """
    try:
        keys = df.iloc[:, 0].tolist()
        values = df.iloc[:, 1].tolist()
        raw_dict = dict(zip(keys, values))
        print(f"Raw dictionary from {sheet_name}: {list(raw_dict.items())[:5]}...")

        cleaned_dict = {}
        for key, value in raw_dict.items():
            cleaned_key = key.replace(':', '').strip()
            cleaned_value = str(value).strip()
            if cleaned_value.endswith(','):
                cleaned_value = cleaned_value[:-1]

            if cleaned_key == 'presupuesto_con_impuestos':
                try:
                    num_value = float(cleaned_value.replace('.', '').replace(',', '.'))
                    cleaned_value = f"${num_value:,.0f}".replace(',', '.')
                except ValueError:
                    pass
            elif cleaned_key in ['plazo_meses', 'dias_vigencia_publicacion']:
                cleaned_value = str(cleaned_value)

            if cleaned_key:
                cleaned_dict[cleaned_key] = cleaned_value

        print(f"Cleaned dictionary from {sheet_name}: {list(cleaned_dict.items())[:5]}...")
        return cleaned_dict
    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {e}")
        return {}


def generate_contexts(wd):
    """
    Genera los contextos para las plantillas a partir de un archivo Excel en el directorio de trabajo.

    Args:
        wd (str): Directorio de trabajo donde se encuentra el archivo Excel.

    Returns:
        tuple: (context_for_template1, context_for_template2) si tiene éxito, (None, None) si falla.
    """
    try:
        excel_path = os.path.join(wd, "Libro1.xlsx")
        if not os.path.exists(excel_path):
            print(f"Error: No se encuentra el archivo Excel en {excel_path}")
            return None, None

        # Leer y limpiar datos de Excel
        Datos_Base_excel = strip_dataframe_and_handle_empty(
            pd.read_excel(excel_path, sheet_name="Datos_Base", header=None))
        Datos_Contrato_P1 = strip_dataframe_and_handle_empty(
            pd.read_excel(excel_path, sheet_name="Datos_Contrato_P1", header=None))
        Datos_Contrato_P2 = strip_dataframe_and_handle_empty(
            pd.read_excel(excel_path, sheet_name="Datos_Contrato_P2", header=None))
        print("Excel sheets read and cleaned successfully.")

        # Procesar datos en diccionarios
        base_data_dict = process_and_clean_dict(Datos_Base_excel, "Datos_Base")
        contrato_p1_data_dict = process_and_clean_dict(Datos_Contrato_P1, "Datos_Contrato_P1")
        contrato_p2_data_dict = process_and_clean_dict(Datos_Contrato_P2, "Datos_Contrato_P2")

        # Ajustes específicos para base_data_dict
        for anexo in ["anexo_6", "anexo_7", "anexo_8", "anexo_9"]:
            if base_data_dict.get(anexo):
                base_data_dict[anexo] = f"Anexo N°{anexo.split('_')[-1]} {base_data_dict[anexo]}"
        if base_data_dict.get("entrega_muestras"):
            base_data_dict["muestra"] = "Entregar muestras de los productos solicitados y comodato ofertado."
        if base_data_dict.get("garantia"):
            base_data_dict["entreg_garan"] = "Entregar garantías de la oferta"
        base_data_dict[
            "labor_coordinador"] = " en el desempeño de su cometido, el coordinador del contrato deberá, a lo menos"
        base_data_dict["coordinador"] = " El adjudicatario nombra coordinador del contrato a"

        # Ajustes específicos para contrato_p1_data_dict
        contrato_p1_data_dict["espacio"] = " "
        contrato_p1_data_dict["Documentos_Integrantes"] = "Tercero"
        contrato_p1_data_dict["Cuarto_ModificacionDelContrato"] = "Cuarto"
        contrato_p1_data_dict["Quinto_GastoseImpuestos"] = "Quinto"
        contrato_p1_data_dict["Sexto_EfectosDerivadosDeIncumplimiento"] = "Sexto"
        contrato_p1_data_dict["Septimo_DeLaGarantíaFielCumplimiento"] = "Séptimo"
        contrato_p1_data_dict["Octavo_CobroDeLaGarantiaFielCumplimiento"] = "Octavo"
        contrato_p1_data_dict["Noveno_TerminoAnticipadoDelContrato"] = "Noveno"
        contrato_p1_data_dict["Decimo_ResciliacionMutuoAcuerdo"] = "Décimo"
        contrato_p1_data_dict["DecimoPrimero_ProcedimientoIncumplimiento"] = "Décimo Primero"
        contrato_p1_data_dict["DecimoSegundo_EmisionOC"] = "Decimo Segundo"
        contrato_p1_data_dict["DecimoTercero_DelPago"] = "Décimo Tercero"
        contrato_p1_data_dict["DecimoCuarto_VigenciaContrato"] = "Décimo Cuarto"
        contrato_p1_data_dict["DecimoQuinto_AdministradorContrato"] = "Décimo Quinto"
        contrato_p1_data_dict["DecimoSexto_PactoDeIntegrida"] = "Décimo Sexto"
        contrato_p1_data_dict["DecimoSeptimo_ComportamientoEticoAdjudic"] = "Décimo Séptimo"
        contrato_p1_data_dict["DecimoOctavo_Auditorias"] = "Décimo Octavo"
        contrato_p1_data_dict["DecimoNoveno_Confidencialidad"] = "Décimo Noveno"
        contrato_p1_data_dict["Vigesimo_PropiedadDeLaInformacion"] = "Vigésimo"
        contrato_p1_data_dict["VigesimoPrimero_SaldosInsolutos"] = "Vigésimo Primero"
        contrato_p1_data_dict["VigesimoSegundo_NormasLaboralesAplicable"] = "Vigésimo Segundo"
        contrato_p1_data_dict["VigesimoTercero_CambioPersonalProveedor"] = "Vigésimo Tercero"
        contrato_p1_data_dict["VigesimoCuarto_CesionySubcontratacion"] = "Vigésimo Cuarto"
        contrato_p1_data_dict["VigesimoQuinto_Discrepancias"] = "Vigésimo Quinto"
        contrato_p1_data_dict["VigesimoSexto_Constancia"] = "Vigésimo Sexto"
        contrato_p1_data_dict["VigesimoSeptimo_Ejemplares"] = "Vigésimo Séptimo"
        contrato_p1_data_dict["VigesimoOctavo_Personeria"] = "Vigésimo Octavo"
        contrato_p1_data_dict["coordinador"] = "El adjudicatario nombra coordinador del contrato a"
        contrato_p1_data_dict["texto_gar_1"] = ", es decir"
        contrato_p1_data_dict[
            "texto_gar_2"] = "de pesos a nombre de “EL HOSPITAL” y consigna la siguiente glosa: Para garantizar el fiel cumplimiento del contrato denominado:"
        contrato_p1_data_dict["texto_gar_3"] = "ds"
        contrato_p1_data_dict[
            "director"] = " la Resolución Exenta RA 116395/343/2024 de fecha 12/08/2024 del SSMOCC., la cual nombra Director del Hospital San José de Melipilla al suscrito"

        # Preparar contextos para las plantillas
        context_for_template1 = base_data_dict
        context_for_template2 = {}
        context_for_template2.update(base_data_dict)
        context_for_template2.update(contrato_p1_data_dict)
        context_for_template2.update(contrato_p2_data_dict)

        return context_for_template1, context_for_template2
    except Exception as e:
        print(f"Error al generar contextos en {wd}: {e}")
        return None, None
