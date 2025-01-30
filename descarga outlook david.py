import win32com.client
import os
from datetime import datetime

# Crear una instancia de la aplicación Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Seleccionar la carpeta de bandeja de entrada
inbox = outlook.GetDefaultFolder(6)  # 6 es el índice de la bandeja de entrada
messages = inbox.Items

# Filtrar los correos no leídos de las direcciones especificadas
filtered_messages = messages.Restrict("[Unread] = True AND ([SenderEmailAddress] = 'hpavez@bullcapital.cl' OR [SenderEmailAddress] = 'daraya@bullcapital.cl' OR [SenderEmailAddress] = 'nhuerta@bullcapital.cl')")

# Obtener la fecha de hoy
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Crear una carpeta llamada 'david' en el escritorio del usuario
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
david_folder = os.path.join(desktop_path, "david")
hemel_folder = os.path.join(desktop_path, "hemel")
nico_folder = os.path.join(desktop_path, "nico")

# Crear las carpetas si no existen
if not os.path.exists(david_folder):
    os.makedirs(david_folder)
    print(f"Se ha creado la carpeta: {david_folder}")
else:
    print(f"Utilizando la carpeta existente: {david_folder}")

if not os.path.exists(hemel_folder):
    os.makedirs(hemel_folder)
    print(f"Se ha creado la carpeta: {hemel_folder}")
else:
    print(f"Utilizando la carpeta existente: {hemel_folder}")

if not os.path.exists(nico_folder):
    os.makedirs(nico_folder)
    print(f"Se ha creado la carpeta: {nico_folder}")
else:
    print(f"Utilizando la carpeta existente: {nico_folder}")

# Crear subcarpetas 'correos_fecha de hoy' dentro de cada carpeta
david_correos_folder = os.path.join(david_folder, f"correos_{fecha_hoy}")
hemel_correos_folder = os.path.join(hemel_folder, f"correos_{fecha_hoy}")
nico_correos_folder = os.path.join(nico_folder, f"correos_{fecha_hoy}")

if not os.path.exists(david_correos_folder):
    os.makedirs(david_correos_folder)
    print(f"Se ha creado la carpeta: {david_correos_folder}")
else:
    print(f"Utilizando la carpeta existente: {david_correos_folder}")

if not os.path.exists(hemel_correos_folder):
    os.makedirs(hemel_correos_folder)
    print(f"Se ha creado la carpeta: {hemel_correos_folder}")
else:
    print(f"Utilizando la carpeta existente: {hemel_correos_folder}")

if not os.path.exists(nico_correos_folder):
    os.makedirs(nico_correos_folder)
    print(f"Se ha creado la carpeta: {nico_correos_folder}")
else:
    print(f"Utilizando la carpeta existente: {nico_correos_folder}")

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
            if message.SenderEmailAddress == 'daraya@bullcapital.cl':
                file_path = os.path.join(david_correos_folder, attachment.FileName)
            elif message.SenderEmailAddress == 'hpavez@bullcapital.cl':
                file_path = os.path.join(hemel_correos_folder, attachment.FileName)
            elif message.SenderEmailAddress == 'nhuerta@bullcapital.cl':
                file_path = os.path.join(nico_correos_folder, attachment.FileName)
            else:
                continue  # Si el remitente no coincide, saltar al siguiente adjunto
            attachment.SaveAsFile(file_path)
            print(f"Archivo guardado en: {file_path}")

print("Correos procesados exitosamente.")