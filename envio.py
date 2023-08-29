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

# ID de tu Google Sheet
SPREADSHEET_ID = '1df1tKhQZBqJAqL1lrQr-TEKh_FjEwk7NStyGpPu-fUo'

# Obtener el nombre de la primera hoja (índice 0)
spreadsheet = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
sheet_name = spreadsheet['sheets'][0]['properties']['title']

# Busca archivos CSV en la carpeta actual
for filename in os.listdir('.'):
    if filename.endswith(" DATOS.csv"):
        with open(filename, newline='') as f:
            reader = csv.reader(f)
            data = list(reader)
            
            # Extrae el ID de la columna A del CSV
            file_id = data[1][0]  # Asume que el ID está en la segunda fila de la columna A
            
            # Busca el ID en la columna A de la primera hoja
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                        range=f"{sheet_name}!A:A").execute()
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
                    update_range = f"{sheet_name}!D{row_number+i}:F{row_number+i}"
                    sheet.values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=update_range,
                        body={"values": [update_data[:3]]},
                        valueInputOption="RAW"
                    ).execute()
                    
                    update_range = f"{sheet_name}!L{row_number+i}"
                    sheet.values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=update_range,
                        body={"values": [[update_data[3]]]},
                        valueInputOption="RAW"
                    ).execute()
                    
                print(f"Datos actualizados para el ID {file_id}.")
