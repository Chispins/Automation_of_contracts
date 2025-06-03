import json
import pandas as pd

direction = r"C:\Users\Thinkpad\Desktop\Automation_of_contracts-6b1b2859afb355479285605ff35bcdf1af8b920c\Files\Libro1.xlsx"
data_p1 = pd.read_excel(direction, sheet_name="Datos_Base")
data_p2 = pd.read_excel(direction, sheet_name="Datos_Contrato_P1")
data_p3 = pd.read_excel(direction, sheet_name="Datos_Contrato_P2")

name_tender = "1058078-15-LE25"
data_p1_key, data_p1_value = data_p1.iloc[:,0], data_p1.iloc[:,1]
data_p2_key, data_p2_value = data_p2.iloc[:,0], data_p2.iloc[:,1]
data_p3_key, data_p3_value = data_p3.iloc[:,0], data_p3.iloc[:,1]

base_dict = {}
base_dict["Base_data"] = { key: value for key, value in zip(data_p1_key, data_p1_value) }
base_dict["Datos_Contrato_p1"] = { key: value for key, value in zip(data_p2_key, data_p2_value) }
base_dict["Datos_Contrato_p2"] = { key: value for key, value in zip(data_p3_key, data_p3_value) }

individual_tender_dict = {}
tender_id = "1058078-15-LE25"
individual_tender_dict[tender_id] = base_dict

individual_tender_dict["1058078-67-LR25"] = base_dict

json_individual = json.dumps(individual_tender_dict, indent = 4)

with open("json_practice.json", "w") as f:
    f.write(json_individual)






for key, value in zip(data_p1_key, data_p1_value):
    base_dict[key] = value

contrato_p1 = {}
for key, value in zip(data_p2_key, data_p2_value):
    contrato_p1[key] = value

contrato_p2 = {}
for key, value in zip(data_p3_key, data_p3_value):
    contrato_p2[key] = value

merged_dict = {}
merged_dict[]


data_1 = ["Nombre", "RUT", "Adjudicado", "Representante_Legal"]
data_2 = ["Mario", "20.202.201-2", "Meliplex", "Rocio"]
ziped_data = zip(data_1, data_2)

+# For the first element
base_dict = {}
for key, value in zip(data_1, data_2):
    base_dict[key] = value

# For the second element
data_3, data_4 = ["ew", "32"]
contrato_p1_dict = {}
for key, value in zip(data_3, data_4):
    base_dict[key] = value

data_5, data_6 = ["ew", "32"]
contrato_p2_dict = {}
for key, value in zip(data_5, data_6):
    base_dict[key] = value

merged_dict = {}
merged_dict["Base"] = base_dict
merged_dict["Contrato"] = contrato_p1_dict
merged_dict["Contrato_parte_2"] = contrato_p2_dict

final_agregation = {}
final_agregation["1058078-14-LE25"] = merged_dict
name_lic = "10.58078-14-LE25"

register_dictionary = json.dumps(merged_dict, indent = 4)
# Now save that file
with open("historial_licitaciones.json", "w") as archivo_json:
    json.dumps(merged_dict, indent=4)

historial_licitaciones_dict = json.loads("historial_licitaciones.json")
historial_licitaciones_dict[name_lic] = merged_dict
with open("historial_licitaciones.json", "w") as archivo_json:
    json.dumps(historial_licitaciones_dict, indent=4)
