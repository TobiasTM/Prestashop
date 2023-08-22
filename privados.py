from bs4 import BeautifulSoup
import re
import csv
from fpdf import FPDF

def obtener_referencias_cantidades_precios(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    product_rows = soup.find_all('tr', class_='product-line-row')

    referencias = []
    cantidades = []
    precios = []

    for product_row in product_rows:
        referencia = None
        cantidad = None
        precio = None

        for line in product_row.stripped_strings:
            if "Número de referencia:" in line:
                referencia = line.split(":")[1].strip()

        quantity_td = product_row.find('td', class_='productQuantity text-center')
        if quantity_td:
            cantidad = quantity_td.text.strip()

        precio_td = product_row.find('td', class_='total_product')
        if precio_td:
            precio_bruto = precio_td.text.strip()
            precio = limpiar_precio(precio_bruto)

        referencias.append(referencia)
        cantidades.append(cantidad)
        precios.append(precio)

    return referencias, cantidades, precios



def limpiar_precio(precio):
    # Remover el "ARS" y espacios extra, luego reemplazar el punto por nada
    return precio.replace("ARS", "").strip().replace(".", "")

def calcular_precio_sin_iva(precio, factor_iva):
    # Convertir el precio a float y calcular el precio sin IVA
    precio_float = float(precio.replace(',', '.'))
    return round(precio_float / factor_iva, 2)

def obtener_pedido_nombre(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    h1 = soup.find('h1', class_='page-title')
    pedido = None
    nombre = None
    if h1:
        texto = h1.text.strip()
        partes = texto.split(" de ")
        if len(partes) == 2:
            pedido_partes = partes[0].split()
            if len(pedido_partes) >= 2:
                pedido = pedido_partes[1]
                nombre = partes[1]
    return pedido, nombre

def obtener_url_from_line(html_content):
    match = re.search(r'psl_current_url = "(https://[^"]+)"', html_content)
    return match.group(1) if match else None

with open("codigo.txt", "r", encoding="utf-8") as file:
    html_content = file.read()

referencias, cantidades, precios = obtener_referencias_cantidades_precios(html_content)
pedido, nombre = obtener_pedido_nombre(html_content)
url = obtener_url_from_line(html_content)

if pedido:
    pedido = pedido.replace("TM00", "")

cantidad_total = sum([int(cant) for cant in cantidades if cant.isdigit()])

# Archivo CSV 1
file_name_1 = f'{pedido}_resumen.csv' if pedido else 'resultado_resumen.csv'
with open(file_name_1, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = ['Número de pedido', 'Cantidad', 'URL', 'Nombre del cliente']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    data = {
        'Número de pedido': pedido,
        'Cantidad': cantidad_total,
        'URL': url,
        'Nombre del cliente': nombre
    }
    writer.writerow(data)

# Archivo CSV 2
file_name_2 = f'{pedido}_detalles.csv' if pedido else 'resultado_detalles.csv'
with open(f'detalles_{pedido}.csv', 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = ['Número de referencia', 'Cantidad (otra vez)', 'Precio', 'IVA 1.105', 'IVA 1.21']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()

    for referencia, cantidad_individual, precio in zip(referencias, cantidades, precios):
        iva_1_105 = calcular_precio_sin_iva(precio, 1.105)
        iva_1_21 = calcular_precio_sin_iva(precio, 1.21)
        
        data = {
            'Número de referencia': referencia,
            'Cantidad (otra vez)': cantidad_individual,
            'Precio': precio,
            'IVA 1.105': f"{iva_1_105:.2f}".replace('.', ','),
            'IVA 1.21': f"{iva_1_21:.2f}".replace('.', ',')
        }
        writer.writerow(data)


print(f"Datos de resumen exportados a {file_name_1}")
print(f"Datos de detalles exportados a {file_name_2}")

def generar_etiqueta_envio(pedido, nombre, url):
    class PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, 'Etiqueta de Envío', 0, 1, 'C')

        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

    # Crear instancia de PDF
    pdf = PDF()
    pdf.add_page()

    # Establecer fuente
    pdf.set_font('Arial', '', 12)

    # Agregar información del cliente y pedido
    pdf.ln(10)
    pdf.cell(0, 10, f'Número de pedido: {pedido}', 0, 1)
    pdf.cell(0, 10, f'Nombre del cliente: {nombre}', 0, 1)

    # Guardar el PDF
    filename = f"Etiqueta_{pedido}.pdf"
    pdf.output(filename)

    return filename

# Llamar a la función para generar la etiqueta
nombre_archivo = generar_etiqueta_envio(pedido, nombre, url)
print(f"Etiqueta de envío generada: {nombre_archivo}")