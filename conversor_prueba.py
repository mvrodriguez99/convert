import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

# Definir los alcances para que pueda acceder a la API de Google Sheets y Google Drive
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Ruta al archivo JSON con las credenciales de la clave de servicio
# Tuve que crearlo y luego compartir el Google Sheets con este "nuevo correo"
creds = ServiceAccountCredentials.from_json_keyfile_name('effortless-edge-456111-k0-035a45594eb7.json', scopes)

# Autorizar la conexión con Google Sheets
client = gspread.authorize(creds)

# Abre el Google Sheet usando su ID
spreadsheet = client.open_by_key('1VG-r3P67R4NqAEzpvP0-uVnB_m7m-m7W0UOiNHV9ekQ') #este es el de copia, cambiar para el definitivo

# Accede a la primera hoja (si quisiéramos acceder a otra, cambiamos el índice)
worksheet = spreadsheet.get_worksheet(0)

# Obtiene todos los datos de la hoja
data = worksheet.get_all_values()

# Diccionario con conversiones comunes o más complejas
word_to_num = {
    'uno': 1, 'una': 1, 'dos': 2, 'tres': 3, 'cuatro': 4, 'cinco': 5,
    'seis': 6, 'siete': 7, 'ocho': 8, 'nueve': 9, 'diez': 10, 'once': 11, 
    'doce': 12, 'trece': 13, 'catorce': 14, 'quince': 15, 'dieciséis': 16, 
    'diecisiete': 17, 'dieciocho': 18, 'diecinueve': 19, 'veinte': 20, 
    'veintiuno': 21, 'veintidos': 22, 'veintitres': 23, 'veinticuatro': 24,
    'veinticinco': 25, 'veintiseis': 26, 'veintisiete': 27, 'veintiocho': 28,
    'veintinueve': 29, 'treinta': 30, 'cuarenta': 40, 'cincuenta': 50, 'sesenta': 60,
    'setenta': 70, 'ochenta': 80, 'noventa': 90, 'cien': 100, 'ciento': 100, 'doscientos': 200, 
    'trescientos': 300, 'cuatrocientos': 400, 'quinientos': 500, 'seiscientos': 600,
    'setecientos': 700, 'ochocientos': 800, 'novecientos': 900, 'mil': 1000
}

# Diccionario de fracciones comunes
fractions = {
    'medio': 0.5,
    'media': 0.5,
    'un cuarto': 0.25,
    'tres cuartos': 0.75
}

# Función para convertir palabras a números
def cantidad_a_numero(cantidad_letra):
    cantidad_letra = cantidad_letra.strip().lower()

    # Intentar convertir directamente a número (entero o flotante)
    try:
        # Reemplazar la coma por punto en caso de decimales con coma
        cantidad_letra = cantidad_letra.replace(',', '.')
        return float(cantidad_letra) if '.' in cantidad_letra else int(cantidad_letra)
    except ValueError:
        pass  # Continuamos si no se puede convertir directamente

    # Caso simple: Si la cantidad es una palabra exacta (por ejemplo, "dos", "mil")
    if cantidad_letra in word_to_num:
        return word_to_num[cantidad_letra]

    # Manejo de fracciones como "uno y medio" o "dos y tres cuartos"
    match_fraction = re.match(r'(\w+)\s*y\s*(\w+\s\w+|\w+)', cantidad_letra)
    if match_fraction:
        num1 = word_to_num.get(match_fraction.group(1))  # El número entero (por ejemplo "dos")
        fraction = match_fraction.group(2)  # La fracción (por ejemplo "medio" o "un cuarto")

        if num1 is not None and fraction in fractions:
            return num1 + fractions[fraction]    

    # Manejo de números decimales como "dos coma cinco" o "dos punto cinco"
    match_decimal = re.match(r'(\w+)\s*(punto|coma)\s*(\w+)', cantidad_letra)
    if match_decimal:
        parte_entera = match_decimal.group(1)  # "dos"
        separador = match_decimal.group(2)  # "punto" o "coma"
        parte_decimal = match_decimal.group(3)  # "cinco"
        
        if parte_entera in word_to_num and parte_decimal in word_to_num:
            entero = word_to_num[parte_entera]
            decimal = word_to_num[parte_decimal]
            return float(f"{entero}.{decimal}")  # Crear el número decimal

    # Caso "mil"
    match_large_number = re.match(r'(\w+)\s+mil(?:\s+(.*))?', cantidad_letra)
    if match_large_number:
        num_mil = word_to_num.get(match_large_number.group(1))
        resto = match_large_number.group(2)
        if num_mil is not None:
            resultado = num_mil * 1000
            if resto:
                resto_num = cantidad_a_numero(resto)
                if resto_num is not None:
                    resultado += resto_num
            return resultado
        
    # Partes múltiples separadas por espacio
    partes = cantidad_letra.split()
    total = 0
    for palabra in partes:
        if palabra in word_to_num:
            total += word_to_num[palabra]
    if total > 0:
        return total
    
    # Si no se pudo convertir
    return None

''' # Caso redundante una vez que tenemos el código de partes múltiples separadas por espacio 
    # Caso "decenas y unidades" → veinte y tres
    match_decenas_unidades = re.match(r'(\w+)\s+y\s+(\w+)', cantidad_letra)
    if match_decenas_unidades:
        decena = word_to_num.get(match_decenas_unidades.group(1))
        unidad = word_to_num.get(match_decenas_unidades.group(2))
        if decena is not None and unidad is not None:
            return decena + unidad
'''

# Establecer el título de la nueva columna (columna G)
worksheet.update_cell(1, 7, "Cantidad")  # Título en la fila 1, columna G

# Recorre todas las filas (desde la segunda fila en adelante)
for i, row in enumerate(data[1:], start=2): 
    cantidad_letra = row[2]  # columna C (índice 2)
    
    # Si la celda está vacía o contiene solo espacios, la ignoramos
    if not cantidad_letra.strip():
        continue
    
    # Llamamos a la función para onvertir la cantidad de letra a número
    cantidad_numero = cantidad_a_numero(cantidad_letra)
    
    if cantidad_numero is not None:
        # Actualiza la columna G (índice 7) con el número convertido
        worksheet.update_cell(i, 7, cantidad_numero)
    else:
        # Si no se puede convertir, dejar la celda en blanco o con un valor específico
        worksheet.update_cell(i, 7, ' ')