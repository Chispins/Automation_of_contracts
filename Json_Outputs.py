import json

practice = '''
    {
  "persona": {
    "definicion": "La persona es un producto de la capacidad recursiva del lenguaje humano, no solo una narrativa que contamos sobre nosotros mismos y los demás.",
    "mecanismo_generador": "El lenguaje humano, que permite coordinar acciones y realizar giros recursivos en una progresión abierta.",
    "base_biologica": "La extraordinaria plasticidad del sistema nervioso humano."
  },
  "organizacion": {
    "definicion": "Un espacio de conversaciones declarativas unidas por promesas mutuas.",
    "cultura": {
      "elementos": [
        "Un pasado compartido",
        "Una forma colectiva de hacer las cosas en el presente",
        "Un sentido común de dirección hacia el futuro"
      ],
      "importancia_conversaciones": "Esenciales para trascender formas mecánicas de coordinación y producir lazos de cooperación y colaboración."
    }
  },
  "fuente": {
    "documento": "Echeverria_Rafael_Ontologia_del_Lenguaje.pdf",
    "fecha_publicacion": "26-05-2025"
  }
} 

'''

practice_json = json.loads(practice)
practice_json_legible = json.dumps(practice_json, indent=4)
pre = json.dumps(practice, indent = 4)

for element in practice_json_legible:
    print(element)


# API KEY
# E3PQH2SBJCEVUJG1
# First I should create the elements as lists

Tender = "105798-25-LR25"
Page_1 = {"Responsable": {"Administrador": "Ivon", "Referente":"Juan"}, "Firmante": {"Director":"Ricardo Contreras", "Subdirector":"Manuel"}}
Page_2 = {"Number": 3, "Date of implementation": "03-02-2024"}
Page_3 = {"Detalles": "Vigencia", "Notebook": "Garantía"}

# Now mix them
test_1 = json.dumps([Page_1, Page_2], indent=4)

with open("test_1.json", "w") as f:
    f.write(test_1)

