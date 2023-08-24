# Descripción del Proyecto
Este proyecto automatiza la extracción de información de pedido de una página web y la guarda en archivos CSV. Los datos extraídos incluyen detalles del pedido como el número de referencia, cantidad y precio del producto, así como la URL del pedido y el nombre del destinatario.
# Funcionamiento
El código sigue estos pasos:

1-Utiliza una ventana emergente para recibir el número del pedido.

2-Busca y activa una ventana de Chrome con un título específico relacionado con el número del pedido.

3-Ejecuta una serie de comandos de teclado (CTRL+U, CTRL+A, CTRL+C, CTRL+V) para abrir el código fuente de la página y copiar todo su contenido.

4-Cierra la pestaña del código fuente.

5-Utiliza BeautifulSoup para parsear el contenido HTML copiado.

6-Utiliza expresiones regulares y otros métodos de búsqueda para extraer información relevante del HTML parseado.

7-Guarda los detalles y datos extraídos en archivos CSV con nombres que contienen el número del pedido para fácil identificación.

# Cómo utilizar
1-Asegúrese de que todas las bibliotecas requeridas están instaladas.

2-Ejecute el script.

3-Ingrese el número del pedido cuando se le solicite.

4-tener la ventana del pedido abierta.

Los archivos CSV generados estarán en el mismo directorio que el script y llevarán el número del pedido en sus nombres para fácil identificación.
