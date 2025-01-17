import win32com.client
import os
from datetime import datetime

# Crear una instancia de la aplicación Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Seleccionar la carpeta de bandeja de entrada
inbox = outlook.GetDefaultFolder(6)  # 6 es el índice de la bandeja de entrada
messages = inbox.Items

# Filtrar los correos que cumplen con las condiciones
filtered_messages = messages.Restrict("[SenderEmailAddress] = 'daraya@bullcapital.cl'")

# Obtener la fecha de hoy
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Crear una carpeta para guardar los archivos adjuntos
download_folder = os.path.join(r"C:\Users\Lucian\Documents", f"Correosdavid_{fecha_hoy}")
if not os.path.exists(download_folder):
    os.makedirs(download_folder)
    print(f"Se ha creado la carpeta: {download_folder}")
else:
    print(f"Utilizando la carpeta existente: {download_folder}")

# Verificar si se encontraron correos
if len(filtered_messages) == 0:
    print("No se encontraron correos que cumplan con las condiciones especificadas.")
else:
    print(f"Se encontraron {len(filtered_messages)} correos que cumplen con las condiciones especificadas.")

# Procesar los correos filtrados
for message in filtered_messages:
    print(f"Procesando correo con asunto: {message.Subject}")
    attachments = message.Attachments
    for attachment in attachments:
        if attachment.FileName.endswith(".msg"):
            file_path = os.path.join(download_folder, attachment.FileName)
            attachment.SaveAsFile(file_path)
            print(f"Archivo guardado en: {file_path}")

print("Correos procesados exitosamente.")