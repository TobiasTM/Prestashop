from bs4 import BeautifulSoup
import pygetwindow as gw
import pyautogui
import time
import csv
import tkinter as tk
from tkinter import simpledialog
import pyperclip  # Importa pyperclip
import re
import csv
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Tus credenciales aquí como un diccionario de Python
credentials_json = {
    "type": "service_account",
    "project_id": "api-prueba-387019",
    "private_key_id": "c60b4b271d135538c9c9240b4989f26266315b8d",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCQcYeTrcbML6oX\n+NWal4bnx07IyY5+hbmFFjfsQD7bUaLFBMTZ3bt/oldMLFQbUOpRjwxiR7Sym5V4\nv+2467ERK2tMBSsbi9PpTxHVeaYAaipQPn0IJIeraLXfMflpH4/Cvx0/hzQ2d4vb\n+8adz/eIbMG1jn0jzG61ByAW0jn164jZxA5mx+CBowu7Y+f2PsoeOkDGKvpO0cY1\nhlFP1vbEb8ffX7rM3xjJ7soFTP662j1rg7AHtbcjePcVGR6VCK+0oPSKFoiiB/2H\n13iLvIW3kcyJjHiDWS+0lST2+4UbtgmVsEdlMMHWc3VyQ4rgB80QVbIHUbzorH0n\nBD4MNNYVAgMBAAECggEACvmZWH2lI9lroOfUYRXqLJ0ZTcp6HtUYessso3E2kBxZ\n6Vg6qybSIDlMopohUjRZzX8jQioadHy7r3b+gCURE+cllLCK1+G4ulrlQGq+aScn\n67zHwXE3HsLKw433mlweef4NTJ5Vt7K/KCSaBjPUUM5vVyp1Cd3hZooHTzyY0Kgq\n1kr0WkJK/rS70t2lm4neP534XtvdktVyBIqezOIsx/46u1ZroLmxYlJO1RLBXgjP\n240QdSiDGIdpPaNwkRqWRnxhkqzWr9ay2xYit0DkPfMyic4PhRitetznkkXFLr61\nnMrKrApIZb4xz/Yi9QnqV0FRwT3tZoWbpLbIMXoQUQKBgQDJzHz8O7jFbpnQ6kT+\n4ZJgTkECSWlq49wlRp6j3ewB083GSurMbx32PceDT8nBk5Nv0JZSecdUB/98Fxoc\nMc3GhvkDXmwKOe+mGWjyONU7oj10LI98hUPKq2sK7hgSrQmPZwxn0qCTcYIsYeQ5\nGp7xGaJL+HPrJ4C65gA/PFgu0QKBgQC3PVboBKpa3q4wmBMKWb+NlORx9AK9Fu3S\nbK0yw3oEkqNHpNcTDo3GbKesU6zQv9A+m2GNKmuXbXtIdlVMhLugR5R6+wUFWTsL\nzAhSyFOoTwga17LY7nL3Iu1Au6YfHhKj8bClcNlkZRMrluf60Kl98pRqRCkfj9NW\njP6nS3AsBQKBgQCyvSz3PO6r8QrMwLPcDnBYXPe3zs5QnwKfAa4B9s7Tz4az3Cec\na89eC9prtIA/tTciEt8SrkqfY3Yns06tKm/ZKDPnh/qqFCbwOBF8elpkN4+3FsEA\nygkBulNVmw43fIy7N8sFKsqPzjo+lXZQHgQqCUp3f9kssBCVeqM9X3W8AQKBgAfo\nMbPZX7CEI2gdZ9TugoGNhz4TlXqrXp/R6LdkEAPagAk7Z7x+yEdjsOiSw8ZOQKIy\n+kapKfNi2gsKcCvZHm+QJywXYOQWMaIUr9dCpbmBj4v4+tK5l2RqsWo1rrlxBsTk\nTQcWk4rtgaJD5MbB8k5pBVaAknW2MxxtASAe9TwxAoGBAMjd2VlT3i4eZo7KoqyK\niVgqVOtQMMxp2jQrZY9KvzTxDdEVkvFFZfhOt47IEEoNjxKZ2PdALFj7F8gk35uP\nbDW2xEaCUYaMACYxgGhLJFZgJof0KaKyO3ijXZA1C01N6Typ8WycZC0iKXU0Q8oH\nBa5QaYSXuXNpNweCEJu3DV4X\n-----END PRIVATE KEY-----\n",
    "client_email": "api-prueba@api-prueba-387019.iam.gserviceaccount.com",
    "client_id": "105131300813815017345",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/api-prueba%40api-prueba-387019.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}
# Inicializa la API de Google Sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = Credentials.from_service_account_info(credentials_json, scopes=scopes)

service = build('sheets', 'v4', credentials=credentials)
sheet = service.spreadsheets()

def get_window_title():
    def on_ok():
        nonlocal pedido, custom_id
        pedido = entry_pedido.get()
        custom_id = entry_id.get()
        root.destroy()

    def center_window(w=300, h=200):
        # Obtiene las dimensiones de la pantalla
        ws = root.winfo_screenwidth()
        hs = root.winfo_screenheight()
        # Calcula la posición x, y
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    custom_id = None
    pedido = None
    root = tk.Tk()
    center_window(500, 250)  # Centra la ventana en la pantalla
    root.title('Ingreso de pedido')

    tk.Label(root, text="Ingrese el número del pedido:", font=("Arial", 14, 'bold')).pack(pady=10)
    entry_pedido = tk.Entry(root, width=50, justify='center')
    entry_pedido.pack(pady=5)

    tk.Label(root, text="Ingrese el ID personalizado:", font=("Arial", 14, 'bold')).pack(pady=10)
    entry_id = tk.Entry(root, width=50, justify='center')
    entry_id.pack(pady=5)

    tk.Button(root, text="OK", command=on_ok).pack(pady=10)
    
    root.mainloop()

    if pedido:
        return f"Pedidos > {pedido} • Todomicro", custom_id
    else:
        return None, None

title, custom_id = get_window_title()

def obtener_codigo_fuente_chrome(title):
    chrome_windows = gw.getWindowsWithTitle(title)
    if not chrome_windows:
        print(f"No se encontró una ventana de Chrome con el título '{title}'.")
        return None
    chrome_window = chrome_windows[0]
    chrome_window.activate()
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'u')
    time.sleep(3)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'w')
    return pyperclip.paste()  # Usa pyperclip.paste en lugar de pyautogui.paste

html_content = obtener_codigo_fuente_chrome(title)

# Convertir el contenido HTML extraído en un objeto BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

def clean_price(price_text):
    price_text = price_text.replace(" ARS", "")  # Elimina la moneda
    price_text = price_text.replace(".", "")  # Elimina los puntos
    return price_text


def extract_information(soup):
    # Para obtener la URL actual del pedido
    script_content = soup.find('script', string=re.compile('psl_current_url'))
    if script_content is not None:
        url_pattern = re.compile(r'psl_current_url = "(.*?)"')
        match = url_pattern.search(script_content.string)
        order_url = match.group(1) if match else 'N/A'
    else:
        order_url = 'N/A'

    # Extraer detalles de los productos
    product_elements = soup.find_all(lambda tag: tag.name == "a" and "Número de referencia:" in tag.get_text())

    detalle_data = []
    total_quantity = 0  # Inicializar la suma total de la cantidad
    for product_element in product_elements:
        reference_parts = product_element.get_text().split('Número de referencia:')
        product_reference = reference_parts[1].strip() if len(reference_parts) > 1 else 'N/A'

        product_quantity_element = product_element.find_next('span', class_='product_quantity_show')
        product_quantity = product_quantity_element.get_text().strip() if product_quantity_element else 'N/A'

        product_price_element = product_element.find_next('span', class_='product_price_show')
        if product_price_element:
            raw_price = product_price_element.get_text().strip()
            clean_product_price = clean_price(raw_price)  # Llamada a clean_price para limpiar el formato del precio
        else:
            clean_product_price = 'N/A'

        detalle_data.append([product_reference, product_quantity, clean_product_price])
                # Sumar la cantidad al total (asegurándose de que sea un número)
        if product_quantity != 'N/A':
            total_quantity += int(product_quantity)

    # Extraer información del pedido
    order_info_element = soup.select_one('h1.page-title')
    order_info = order_info_element.get_text().strip() if order_info_element else 'N/A'

    # Separar el número de pedido y el nombre del destinatario
    info_parts = re.search(r"Pedido TM00(\d+) de (.+)", order_info)
    if info_parts:
        order_number = info_parts.group(1)  # Guardar solo el número sin el prefijo "TM00"
        recipient_name = info_parts.group(2)  # Guardar el nombre del destinatario
    else:
        order_number = 'N/A'
        recipient_name = 'N/A'

    datos_data = [[custom_id, order_number, total_quantity, order_url, recipient_name]]

    return detalle_data, datos_data


detalle, datos = extract_information(soup)

# Obtenemos el número de pedido de los datos recolectados
order_number = datos[0][0]  # Asumiendo que datos[0][0] contiene el número de pedido

# Nombre del archivo CSV para detalles
detalle_filename = f"{order_number} DETALLES.csv"


# Verifica si hemos obtenido contenido HTML
if html_content:
    # Convertir el contenido HTML extraído en un objeto BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')

    # Busca la información de la dirección
    scripts = soup.find_all('script')
    for script in scripts:
        script_content = script.string
        if script_content and "address:" in script_content:
            # Usa expresión regular para encontrar la dirección
            match = re.search(r"address: '(.*?)'", script_content)
            if match:
                address_info = match.group(1)
                print(f"Se encontró la dirección: {address_info}")

                # Separa la información de la dirección en sus componentes
                domicilio, cp, localidad, ciudad, pais = address_info.split(',')

                # Guarda en CSV
                csv_file = 'datos_de_envio.csv'
                # Verifica si el archivo ya existe para saber si añadir headers
                file_exists = os.path.isfile(csv_file)

                with open(csv_file, mode='a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    if not file_exists:
                        writer.writerow(["Domicilio", "CP", "Localidad", "Ciudad", "País"])
                    writer.writerow([domicilio.strip(), cp.strip(), localidad.strip(), ciudad.strip(), pais.strip()])
else:
    print("No se pudo obtener el contenido HTML.")

# Escribir detalles en el archivo CSV
with open(detalle_filename, 'w', encoding='utf-8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['Número de Referencia', 'Cantidad', 'Precio del Producto'])
    writer.writerows(detalle)

# Nombre del archivo CSV para datos
datos_filename = f"{order_number} DATOS.csv"

# Escribir datos en el archivo CSV
with open(datos_filename, 'w', encoding='utf-8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['ID', 'Información del Pedido', 'Cantidad', 'URL del Pedido', 'Destinatario'])
    writer.writerows(datos)

print(f"La información ha sido guardada en {detalle_filename} y {datos_filename}.")

def extract_information(soup):
    # Para obtener la URL actual del pedido
    script_content = soup.find('script', string=re.compile('psl_current_url'))
    if script_content is not None:
        url_pattern = re.compile(r'psl_current_url = "(.*?)"')
        match = url_pattern.search(script_content.string)
        order_url = match.group(1) if match else 'N/A'
    else:
        order_url = 'N/A'

    # Extraer detalles de los productos
    product_elements = soup.find_all(lambda tag: tag.name == "a" and "Número de referencia:" in tag.get_text())

    detalle_data = []
    total_quantity = 0  # Inicializar la suma total de la cantidad
    for product_element in product_elements:
        reference_parts = product_element.get_text().split('Número de referencia:')
        product_reference = reference_parts[1].strip() if len(reference_parts) > 1 else 'N/A'

        product_quantity_element = product_element.find_next('span', class_='product_quantity_show')
        product_quantity = product_quantity_element.get_text().strip() if product_quantity_element else 'N/A'

        product_price_element = product_element.find_next('span', class_='product_price_show')
        if product_price_element:
            raw_price = product_price_element.get_text().strip()
            clean_product_price = clean_price(raw_price)  # Llamada a clean_price para limpiar el formato del precio
        else:
            clean_product_price = 'N/A'

        detalle_data.append([product_reference, product_quantity, clean_product_price])
                # Sumar la cantidad al total (asegurándose de que sea un número)
        if product_quantity != 'N/A':
            total_quantity += int(product_quantity)

    # Extraer información del pedido
    order_info_element = soup.select_one('h1.page-title')
    order_info = order_info_element.get_text().strip() if order_info_element else 'N/A'

    # Separar el número de pedido y el nombre del destinatario
    info_parts = re.search(r"Pedido TM00(\d+) de (.+)", order_info)
    if info_parts:
        order_number = info_parts.group(1)  # Guardar solo el número sin el prefijo "TM00"
        recipient_name = info_parts.group(2)  # Guardar el nombre del destinatario
    else:
        order_number = 'N/A'
        recipient_name = 'N/A'

    datos_data = [[custom_id, order_number, total_quantity, order_url, recipient_name]]

    return detalle_data, datos_data


detalle, datos = extract_information(soup)

# Obtenemos el número de pedido de los datos recolectados
order_number = datos[0][0]  # Asumiendo que datos[0][0] contiene el número de pedido

# Nombre del archivo CSV para detalles
detalle_filename = f"{order_number} DETALLES.csv"

# Escribir detalles en el archivo CSV
with open(detalle_filename, 'w', encoding='utf-8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['Número de Referencia', 'Cantidad', 'Precio del Producto'])
    writer.writerows(detalle)

# Nombre del archivo CSV para datos
datos_filename = f"{order_number} DATOS.csv"

# Escribir datos en el archivo CSV
with open(datos_filename, 'w', encoding='utf-8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['ID', 'Información del Pedido', 'Cantidad', 'URL del Pedido', 'Destinatario'])
    writer.writerows(datos)

print(f"La información ha sido guardada en {detalle_filename} y {datos_filename}.")

# ID de tus Google Sheets
SPREADSHEET_ID1 = '1df1tKhQZBqJAqL1lrQr-TEKh_FjEwk7NStyGpPu-fUo'
SPREADSHEET_ID2 = '1eoVpyO0IlAvV1DymRq7tMHvkN7s5T2BUoBdVFw1pxUc'

# Obtener los nombres de las hojas
spreadsheet1 = sheet.get(spreadsheetId=SPREADSHEET_ID1).execute()
sheet_name1 = spreadsheet1['sheets'][0]['properties']['title']

spreadsheet2 = sheet.get(spreadsheetId=SPREADSHEET_ID2).execute()
sheet_name2 = spreadsheet2['sheets'][4]['properties']['title']
# Busca archivos CSV en la carpeta actual
for filename in os.listdir('.'):
    if filename.endswith(" DATOS.csv"):
        with open(filename, newline='') as f:
            reader = csv.reader(f)
            data = list(reader)
            
            # Extrae el ID de la columna A del CSV
            file_id = data[1][0]  # Asume que el ID está en la segunda fila de la columna A
            
            # Busca el ID en la columna A de la primera hoja
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID1,
                                        range=f"{sheet_name1}!A:A").execute()
            values = result.get('values', [])
            
            row_number = None
            for i, row in enumerate(values):
                if row and row[0] == file_id:
                    row_number = i + 1
                    break
            
            if row_number is not None:
                # Actualiza las celdas en la misma fila pero en diferentes columnas
                for i, row in enumerate(data[1:]):  # Ignora la primera fila (encabezados)
                    update_data = [
                        row[1],  # Columna B del CSV
                        row[2],  # Columna C del CSV
                        row[3],  # Columna D del CSV
                        row[4]   # Columna E del CSV
                    ]
                    
                    # Actualiza las celdas D, E, F, L en la primera hoja
                    update_range = f"{sheet_name1}!D{row_number+i}:F{row_number+i}"
                    sheet.values().update(
                        spreadsheetId=SPREADSHEET_ID1,
                        range=update_range,
                        body={"values": [update_data[:3]]},
                        valueInputOption="RAW"
                    ).execute()
                    
                    update_range = f"{sheet_name1}!L{row_number+i}"
                    sheet.values().update(
                        spreadsheetId=SPREADSHEET_ID1,
                        range=update_range,
                        body={"values": [[update_data[3]]]},
                        valueInputOption="RAW"
                    ).execute()
                    
                print(f"Datos actualizados para el ID {file_id}.")
# Busca archivos CSV en la carpeta actual que terminan en " detalles.csv"
for filename in os.listdir('.'):
    if filename.endswith("DETALLES.csv"):
        print(f"Procesando archivo de detalles: {filename}")
        
        with open(filename, newline='') as f:
            reader = csv.reader(f)
            data = list(reader)[1:]  # Ignora la primera fila
            
            # Elimina los datos existentes en las columnas A, B, C (excepto la primera fila)
            sheet.values().clear(
                spreadsheetId=SPREADSHEET_ID2,
                range=f"{sheet_name2}!A2:C"
            ).execute()
            
            # Envía los nuevos datos a las columnas A, B, C
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID2,
                range=f"{sheet_name2}!A2",
                body={"values": data},
                valueInputOption="RAW"
            ).execute()
            
            print(f"Datos de detalles actualizados para el archivo {filename}.")