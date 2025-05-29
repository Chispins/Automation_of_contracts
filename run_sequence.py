"""
Este script orquesta la generación completa de documentos:
1) Portada para Base y Contrato
2) Generación del documento Base
3) Generación del contrato formateado
4) Renderizado final con Jinja2
"""
import os
import sys
from Portada import create_melipilla_document
from Formated_Base_PEP8 import main as formated_base_main
from Formated_Contracts_PEP8 import main as formated_contracts_main
import subprocess

def main():
    project_root = os.path.abspath(os.path.dirname(__file__))
    files_dir = os.path.join(project_root, "Files")
    os.makedirs(files_dir, exist_ok=True)

    # 1) Portadas
    os.chdir(project_root)
    create_melipilla_document(archivo="base")
    create_melipilla_document(archivo="contrato")

    # 2) Documento Base
    os.chdir(project_root)
    formated_base_main()

    # 3) Contrato formateado
    os.chdir(project_root)
    formated_contracts_main()

    # 4) Jinja2
    os.chdir(project_root)
    python_exec = sys.executable
    subprocess.run([python_exec, os.path.join(project_root, "Jinja_2.py")], check=True)

if __name__ == "__main__":
    main()


