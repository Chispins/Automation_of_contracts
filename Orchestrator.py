'''
Este script orquesta la generación de documentos de base y contrato.
Sigue la secuencia:
1. Creación de Portadas (usando Portada.py)
2. Creación del Documento Base (usando Formated_Base_PEP8.py)
3. Creación del Documento de Contrato (usando Formated_Contracts_PEP8.py)
4. Procesamiento final con Jinja2 (ejecutando Jinja_2.py)
'''
import os
import subprocess
import sys

# Importar funciones principales de los otros módulos
from Portada import create_melipilla_document
from Formated_Base_PEP8 import main as formated_base_main
from Formated_Contracts_PEP8 import main as formated_contracts_main

def main_orchestrator():
    '''Función principal para orquestar la generación de documentos.'''
    project_root = os.path.abspath(os.path.dirname(__file__))
    files_dir = os.path.join(project_root, "Files")

    print("--- Iniciando Proceso de Generación de Documentos ---")

    # Asegurar que el directorio 'Files' exista
    if not os.path.exists(files_dir):
        os.makedirs(files_dir)
        print(f"Directorio creado: {files_dir}")

    # Paso 1 & 2: Crear Portadas
    # Cada script llamado (create_melipilla_document, formated_base_main, etc.)
    # se encargará de cambiar su propio directorio de trabajo a 'Files'
    # a través de su función configurar_directorio_trabajo().
    # El orquestador se asegura de que la llamada se haga desde el project_root.
    print("\nPASO 1 & 2: Creando Portadas...")
    current_cwd = os.getcwd()
    if current_cwd != project_root:
        os.chdir(project_root)
    create_melipilla_document(archivo="base")
    print(f"  - portada_melipilla_base.docx creado/actualizado en {files_dir}")

    current_cwd = os.getcwd() # Portada.py cambia el CWD a 'Files', lo reseteamos
    if current_cwd != project_root:
        os.chdir(project_root)
    create_melipilla_document(archivo="contrato")
    print(f"  - portada_melipilla_contrato.docx creado/actualizado en {files_dir}")
    print("Portadas finalizadas.")

    # Paso 3: Crear Documento Base (Formated_Base_PEP8)
    print("\nPASO 3: Creando Documento Base (Formated_Base_PEP8)...")
    current_cwd = os.getcwd()
    if current_cwd != project_root:
        os.chdir(project_root)
    formated_base_main()
    print(f"  - base_automatizada.docx creado/actualizado en {files_dir}")
    print("Creación de documento base finalizada.")

    # Paso 4: Crear Documento de Contrato (Formated_Contracts_PEP8)
    print("\nPASO 4: Creando Documento de Contrato (Formated_Contracts_PEP8)...")
    # Este paso requiere que "contrato_automatizado_rendered.docx" exista en Files/
    # como plantilla de entrada para Formated_Contracts_PEP8.py.
    required_input_for_contract = os.path.join(files_dir, "contrato_automatizado_rendered.docx")
    if not os.path.exists(required_input_for_contract):
        print(f"ADVERTENCIA: Falta el archivo de entrada requerido para Formated_Contracts_PEP8: {required_input_for_contract}")
        print("  Este archivo debe ser una plantilla preexistente si Jinja2 se ejecuta al final.")
        print("  El proceso podría fallar o usar datos obsoletos para este paso.")
    else:
        print(f"  - Usando {required_input_for_contract} como plantilla de entrada para Formated_Contracts_PEP8.")

    current_cwd = os.getcwd()
    if current_cwd != project_root:
        os.chdir(project_root)
    formated_contracts_main()
    print(f"  - contrato_automatizado_tablas.docx creado/actualizado en {files_dir}")
    print("Creación de documento de contrato finalizada.")

    # Paso 5: Ejecutar Procesamiento Jinja2
    print("\nPASO 5: Ejecutando Procesamiento Jinja2...")
    current_cwd = os.getcwd()
    if current_cwd != project_root:
        os.chdir(project_root)

    python_executable = sys.executable # Usa el mismo intérprete de Python que ejecuta el orquestador
    jinja_script_path = os.path.join(project_root, "Jinja_2.py")

    print(f"  - Ejecutando: {python_executable} {jinja_script_path}")
    # Jinja_2.py se ejecutará y usará su propia lógica configurar_directorio_trabajo
    result = subprocess.run([python_executable, jinja_script_path], cwd=project_root, capture_output=True, text=True, check=False)

    if result.returncode == 0:
        print("  - Procesamiento Jinja2 completado exitosamente.")
        if result.stdout:
            print("    Salida de Jinja_2.py:\n", result.stdout)
        print(f"    - base_automatizada_jinja2.docx debería estar en {files_dir}")
        print(f"    - contrato_automatizado_rendered.docx (recién renderizado) debería estar en {files_dir}")
    else:
        print("  - Falló el procesamiento Jinja2.")
        if result.stdout:
            print("    Salida de Jinja_2.py (stdout):\n", result.stdout)
        if result.stderr:
            print("    Errores de Jinja_2.py (stderr):\n", result.stderr)
    print("Procesamiento Jinja2 finalizado.")

    print("\n--- Proceso de Generación de Documentos Finalizado ---")

if __name__ == "__main__":
    main_orchestrator()

