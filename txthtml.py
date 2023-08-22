from bs4 import BeautifulSoup
import pygetwindow as gw
import pyautogui
import time
import csv
import tkinter as tk
from tkinter import simpledialog
import pyperclip  # Importa pyperclip

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
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'w')
    return pyperclip.paste()  # Usa pyperclip.paste en lugar de pyautogui.paste

html_content = obtener_codigo_fuente_chrome(title)

def extract_information(soup):
    product_elements = soup.find_all(lambda tag: tag.name == "a" and "Número de referencia:" in tag.get_text())
    data = []

    for product_element in product_elements:
        product_name_element = product_element.find('span', class_='productName')
        product_name = product_name_element.get_text().strip() if product_name_element else 'N/A'

        reference_parts = product_element.get_text().split('Número de referencia:')
        product_reference = reference_parts[1].strip() if len(reference_parts) > 1 else 'N/A'

        product_quantity = soup.select_one('span.product_quantity_show.badge')
        order_info = soup.select_one('h1.page-title')
        product_price = soup.select_one('span.product_price_show')

        data.append([
            product_name,
            product_reference,
            product_quantity.get_text().strip() if product_quantity else 'N/A',
            order_info.get_text().strip() if order_info else 'N/A',
            product_price.get_text().strip() if product_price else 'N/A'
        ])
    return data

if html_content:
    soup = BeautifulSoup(html_content, 'html.parser')
    extracted_data = extract_information(soup)

    with open('productos.csv', 'w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Nombre del Producto', 'Número de Referencia', 'Cantidad', 'Información del Pedido', 'Precio del Producto'])
        writer.writerows(extracted_data)

    print("La información ha sido guardada en productos.csv.")
else:
    print("Hubo un error al intentar obtener el código HTML.")
