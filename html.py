import pygetwindow as gw
import pyautogui
import time
import os

def obtener_html_de_chrome():
    # Buscar ventanas con el título exacto "Pedidos • Todomicro"
    chrome_windows = gw.getWindowsWithTitle('Pedidos • Todomicro')
    
    # Si no hay ninguna ventana de Chrome con ese título, salir
    if not chrome_windows:
        print("No se encontró una ventana de Chrome con el título 'Pedidos • Todomicro'.")
        return None
    
    # Tomar la primera ventana (si hay varias ventanas que coinciden, se toma la primera)
    chrome_window = chrome_windows[0]
    chrome_window.activate()

    # Esperar a que Chrome sea la ventana activa
    time.sleep(2)
    
    # Abrir el código fuente de la página
    pyautogui.hotkey('ctrl', 'u')
    time.sleep(10)

    # Guardar el código fuente
    pyautogui.hotkey('ctrl', 's')
    time.sleep(2)

    # Escribir la ruta
    ruta = 'C:\\Users\\todom\\OneDrive\\Escritorio\\privados\\codigo.html'
    pyautogui.write(ruta)
    time.sleep(2)

    # Seleccionar "Webpage, HTML Only" y guardar. Asumiendo que al presionar 'enter' se guardará correctamente el archivo.
    pyautogui.press('enter')
    time.sleep(3)

    # Cerrar la pestaña de código fuente
    pyautogui.hotkey('ctrl', 'w')

    # Leer el archivo guardado y devolver su contenido
    if os.path.exists(ruta):
        with open(ruta, 'r', encoding='utf-8') as file:
            html_content = file.read()
        return html_content
    else:
        return None

html_content = obtener_html_de_chrome()

# Verificamos que html_content tenga contenido antes de intentar escribirlo en txt
if html_content:
    # Guardamos el contenido en un archivo llamado Codigo.txt
    with open('Codigo.txt', 'w', encoding='utf-8') as file:
        file.write(html_content)
    print("El código HTML ha sido guardado en Codigo.txt.")
else:
    print("Hubo un error al intentar obtener el código HTML.")aaaa
