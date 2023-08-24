from bs4 import BeautifulSoup
import pygetwindow as gw
import pyautogui
import time
import csv
import tkinter as tk
from tkinter import simpledialog
import pyperclip  # Importa pyperclip
import re

def get_window_title():
    root = tk.Tk()
    root.withdraw() 
    pedido = simpledialog.askstring("Pedido", "Ingrese el número del pedido:")
    if pedido:
        return f"Pedidos > {pedido} • Todomicro"
    else:
        return None

title = get_window_title()

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

def extract_information(soup):
    # Para obtener la URL actual del pedido
    script_content = soup.find('script', string=re.compile('psl_current_url'))
    url_pattern = re.compile(r'psl_current_url = "(.*?)"')
    match = url_pattern.search(script_content.string)
    order_url = match.group(1) if match else 'N/A'

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
        product_price = product_price_element.get_text().strip() if product_price_element else 'N/A'

        detalle_data.append([product_reference, product_quantity, product_price])
                # Sumar la cantidad al total (asegurándose de que sea un número)
        if product_quantity != 'N/A':
            total_quantity += int(product_quantity)
        
        detalle_data.append([product_reference, product_quantity, product_price])
        
        

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

    datos_data = [[order_number, total_quantity, order_url, recipient_name]]  # Usar total_quantity aquí

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
    writer.writerow(['Información del Pedido', 'Cantidad', 'URL del Pedido', 'Destinatario'])
    writer.writerows(datos)

print(f"La información ha sido guardada en {detalle_filename} y {datos_filename}.")
