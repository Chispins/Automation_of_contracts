import os # Línea añadida
from Bases import configurar_directorio_trabajo # Línea añadida

configurar_directorio_trabajo() # Línea añadida

# Inside your Reading_file.py (or wherever the original functions are)
def try_one(numero, multiplicador): # Renamed
    result = numero * multiplicador
    return result

def es_par(numero): # Renamed
    return numero % 2 == 0

# Add these lines to your test file (e.g., test_reading_file.py)
# Make sure to import the functions if they are in a different file
# from Reading_file import try_one, es_par

def test_try_one_multiplies(): # This is the actual pytest test
    assert try_one(2, 3) == 6
    assert try_one(5, -2) == -10
    assert try_one(0, 100) == 0

def test_es_par_checks_even(): # This is another pytest test
    assert es_par(7) is False
    assert es_par(4) is True
    assert es_par(0) is True