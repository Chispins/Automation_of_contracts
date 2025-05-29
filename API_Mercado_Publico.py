import requests

# --- API Configuration ---
# Base URL for the Licitaciones API (JSON format) - Using HTTPS as seen in docs
BASE_URL = "https://api.mercadopublico.cl/servicios/v1/publico/licitaciones.json"

# --- Test Parameters provided in your documentation snippets ---
# THIS IS A TEST TICKET - USE YOUR OWN PRODUCTION TICKET FOR REAL APPLICATIONS
ACCESS_TICKET = "78240120-F9A8-4FDA-AC92-167A134B5FF1"

# Example Tender Code provided in your documentation snippets
# Using a known working example is the best first step
TENDER_CODE_EXAMPLE = "1509-5-L114"

# You could try your code again here, but the example is confirmed in the docs:
# TENDER_CODE_TO_TRY = "1549-207-LR24"

# --- Prepare the GET parameters ---
# To get detailed info for a specific tender, the documentation shows '?codigo=...&ticket=...'
params = {
    "codigo": TENDER_CODE_EXAMPLE, # Use the documented example first
    # "codigo": TENDER_CODE_TO_TRY, # Uncomment this to try your original code after confirming the example works
    "ticket": ACCESS_TICKET
}

# --- Make the HTTP GET request ---
print(f"Making request to: {BASE_URL}")
print(f"With parameters: {params}")


try:
    # Set a timeout to prevent the request from hanging indefinitely
    response = requests.get(BASE_URL, params=params, timeout=10)

    # --- Check the response status code ---
    print(f"Received status code: {response.status_code}")

    if response.status_code == 200:
        data = response.json()

        print("\nRequest successful! Received data:")

        cantidad = data.get("Cantidad")
        print(f"Cantidad de Licitaciones encontradas: {cantidad}")

        licitaciones_list = data.get("Listado", [])
        if licitaciones_list:
            print(f"Number of tenders in list (should be 1 for code lookup): {len(licitaciones_list)}")
            for licitacion in licitaciones_list:
                print("\nTender Details:")

                # *** ADD THIS LINE TO SEE THE RAW DATA STRUCTURE ***
                import json # Import json to pretty-print
                print(json.dumps(licitacion, indent=4))
                # *************************************************

                try:
                    print(f"  - Código Externo (API ID): {licitacion.get('CodigoExterno', 'N/A')}")
                    print(f"  - Nombre: {licitacion.get('Nombre', 'N/A')}")

                    # Accessing nested 'Estado' object - Let's check its type
                    estado_obj = licitacion.get('Estado')
                    print(f"  - Estado (Raw value): {estado_obj} (Type: {type(estado_obj)})") # Print type

                    if isinstance(estado_obj, dict):
                         print(f"  - Estado (Code): {estado_obj.get('Codigo', 'N/A')}")
                         print(f"  - Estado (Nombre): {estado_obj.get('Nombre', 'N/A')}")
                    else:
                         print(f"  - Estado: Unexpected type, cannot access Code/Nombre.")


                    print(f"  - Fecha Cierre: {licitacion.get('FechaCierre', 'N/A')}")

                    # Accessing nested 'Fechas' object - Let's check its type
                    fechas_obj = licitacion.get('Fechas')
                    print(f"  - Fechas (Raw value): {fechas_obj} (Type: {type(fechas_obj)})") # Print type

                    if isinstance(fechas_obj, dict):
                        print(f"  - Fecha Publicación: {fechas_obj.get('FechaCreacion', 'N/A')}")
                        # Add other dates if needed, checking for their existence
                        # print(f"  - Fecha Inicio Foro: {fechas_obj.get('FechaInicio', 'N/A')}")
                        # etc.
                    else:
                         print(f"  - Fechas: Unexpected type, cannot access date fields.")

                    # Add more fields as needed, applying similar checks for nested structures

                except AttributeError as ae:
                    print(f"  - An AttributeError occurred while processing tender details: {ae}")
                    print("    This likely means a value was expected to be a dictionary but was another type (like string).")
                    print(f"    Inspect the raw data printed above to find the problematic field.")


        else:
             print("No 'Listado' of tenders found in the response for this code.")
             print(f"Full response data received: {data}")

except requests.exceptions.Timeout:
    print("\nError: The request timed out.")
except requests.exceptions.ConnectionError:
     print("\nError: Could not connect to the API server.")
except requests.exceptions.RequestException as e:
    # Handle other requests library errors
    print(f"\nAn error occurred during the request: {e}")
except Exception as e:
    # Handle any other unexpected errors
    print(f"\nAn unexpected error occurred: {e}")


import requests
from bs4 import BeautifulSoup
import time # Optional: Add a small delay between requests for politeness

# Get


def get_mercadopublico_sales(rut):
    """
    Fetches the sales value for a given RUT from Mercado Público.

    Args:
        rut (str): The RUT (identifier number) of the provider.

    Returns:
        int or None: The sales value as an integer if found and parsed,
                     None otherwise (e.g., page not found, element not found, parse error).
    """
    base_url = "https://proveedor.mercadopublico.cl/ficha/comportamiento/"
    url = f"{base_url}{rut}"
    print(f"Fetching data for RUT: {rut} from {url}")

    try:
        # Make the HTTP GET request
        # Add a timeout to prevent hanging indefinitely
        response = requests.get(url, timeout=15)

        # Raise an exception for bad status codes (4xx or 5xx)
        response.raise_for_status()

        # Parse the page content
        soup = BeautifulSoup(response.content, 'html.parser')

        # Find the element containing the sales figure
        # Based on inspecting the HTML, the sales figure is in an <h3> tag
        # within a div with class 'box-content-compact'.
        # We can target the h3 specifically within that container.
        sales_element = soup.select_one('.box-content-compact h3')

        if sales_element:
            # Extract the text content
            sales_text = sales_element.get_text()

            # Clean the text: remove '$', '.', ',', and whitespace
            # Example: "$ 454.386.081" becomes "454386081"
            cleaned_text = sales_text.replace('$', '').replace('.', '').replace(',', '').strip()

            try:
                # Convert the cleaned text to an integer
                sales_value = int(cleaned_text)
                print(f"Successfully extracted sales: {sales_value}")
                return sales_value
            except ValueError:
                print(f"Could not convert extracted text '{cleaned_text}' to an integer for RUT {rut}")
                return None # Handle cases where the text format is unexpected
        else:
            print(f"Could not find the sales element (h3 within .box-content-compact) for RUT {rut}.")
            # Check if the page indicates the provider wasn't found or is invalid
            not_found_message = soup.select_one('.box-content-compact p') # Look for potential messages
            if not_found_message and "no existe" in not_found_message.get_text():
                 print(f"Provider with RUT {rut} not found.")
            return None # Element not found on the page

    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for RUT {rut}: {e}")
        return None # Handle request errors (network issues, timeout, bad status code)
    except Exception as e:
        print(f"An unexpected error occurred for RUT {rut}: {e}")
        return None # Handle any other unexpected errors during parsing/processing




# --- Main execution ---

# Your list of RUT numbers
"""rut_numbers = [
    "78.615.850-8"
]
# Dictionary to store the results
# Key: RUT, Value: Sales value (int) or None
sales_results = {}

# Iterate through the list and scrape data for each RUT
for rut in rut_numbers:
    sales = get_mercadopublico_sales(rut)
    sales_results[rut] = sales
    # Optional: Add a small delay to be polite to the server
    # time.sleep(1) # Sleep for 1 second between requests

# --- Output the results ---
print("\n--- Scraping Results ---")
for rut, sales in sales_results.items():
    if sales is not None:
        # Format the sales number with commas for readability
        print(f"RUT: {rut}, VENTAS EN MERCADO PÚBLICO: ${sales:,}")
    else:
        print(f"RUT: {rut}, VENTAS EN MERCADO PÚBLICO: N/A (Could not retrieve)")

# You could also save this data to a file (e.g., CSV)
# import csv
# with open('mercadopublico_sales.csv', 'w', newline='', encoding='utf-8') as csvfile:
#     fieldnames = ['RUT', 'VENTAS_EN_MERCADO_PUBLICO']
#     writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
#
#     writer.writeheader()
#     for rut, sales in sales_results.items():
#         writer.writerow({'RUT': rut, 'VENTAS_EN_MERCADO_PUBLICO': sales if sales is not None else ''})
# print("\nResults saved to mercadopublico_sales.csv")"""


def obtener_detalles_licitacion(tender_id):
    """
    Obtiene todos los detalles de una licitación desde la API de Mercado Público usando su ID.

    Args:
        tender_id (str): ID de la licitación (ej. "1509-5-L114")

    Returns:
        dict: Datos completos de la licitación o None en caso de error
    """
    import requests

    # Configuración de la API
    BASE_URL = "https://api.mercadopublico.cl/servicios/v1/publico/licitaciones.json"
    ACCESS_TICKET = "78240120-F9A8-4FDA-AC92-167A134B5FF1"  # TICKET DE PRUEBA

    # Parámetros de la solicitud
    params = {
        "codigo": tender_id,
        "ticket": ACCESS_TICKET
    }

    try:
        # Realizar la solicitud GET
        response = requests.get(BASE_URL, params=params, timeout=10)

        # Verificar si la solicitud fue exitosa
        if response.status_code == 200:
            data = response.json()

            # Verificar si se encontraron licitaciones
            licitaciones = data.get("Listado", [])
            if licitaciones and len(licitaciones) > 0:
                # Retornar la primera (y probablemente única) licitación encontrada
                return licitaciones[0]
            else:
                print(f"No se encontraron datos para la licitación con ID: {tender_id}")
                return None
        else:
            print(f"Error en la solicitud: Código de estado {response.status_code}")
            return None

    except requests.exceptions.Timeout:
        print("Error: La solicitud excedió el tiempo de espera.")
    except requests.exceptions.ConnectionError:
        print("Error: No se pudo conectar al servidor de la API.")
    except requests.exceptions.RequestException as e:
        print(f"Error durante la solicitud: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")

    return None

obtener_detalles_licitacion("1057480-15-LR24")