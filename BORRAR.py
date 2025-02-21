import os
import win32com.client
from datetime import datetime
import sys
import tkinter as tk
from tkinter import simpledialog

# Crear la ventana principal de Tkinter
root = tk.Tk()
root.withdraw()  # Ocultar la ventana principal

# Solicitar la dirección de correo electrónico al usuario
email_address = simpledialog.askstring("Input", "Inserte email:", parent=root)

# Solicitar el nombre del cliente al usuario
nombre_cliente = simpledialog.askstring("Input", "Inserte el nombre del cliente:", parent=root)

# Verificar si se ingresaron valores
if not email_address or not nombre_cliente:
    print("Debe ingresar el email y el nombre del cliente.")
    sys.exit()

# Obtener la fecha de hoy en formato YYYYMMDD
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Directorio base del escritorio
base_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Equipo Com BULLCAPITAL", "HPAVEZ")

# Verificar si existe una carpeta con el nombre del cliente
directorio_cliente_base = os.path.join(base_dir, nombre_cliente)
if not os.path.exists(directorio_cliente_base):
    os.makedirs(directorio_cliente_base)

# Directorio específico para la fecha de hoy
directorio_cliente_fecha = os.path.join(directorio_cliente_base, f'{nombre_cliente}_{fecha_hoy}')
if not os.path.exists(directorio_cliente_fecha):
    os.makedirs(directorio_cliente_fecha)

directorio_facturas = os.path.join(directorio_cliente_fecha, 'facturas')
directorio_confirmacion_mail = os.path.join(directorio_cliente_fecha, 'confirmacion_mail')

# Crear subdirectorios si no existen
os.makedirs(directorio_facturas, exist_ok=True)
os.makedirs(directorio_confirmacion_mail, exist_ok=True)

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Función para encontrar una carpeta de manera recursiva
def find_folder(folder, folder_name):
    for f in folder.Folders:
        if f.Name == folder_name:
            return f
        subfolder = find_folder(f, folder_name)
        if subfolder:
            return subfolder
    return None

# Acceder a la bandeja de entrada
inbox = namespace.GetDefaultFolder(6)  # 6 es para la bandeja de entrada

# Seleccionar la carpeta "Test" dentro de la bandeja de entrada
carpeta = find_folder(inbox, "Test")
if not carpeta:
    raise Exception("No se encontró la carpeta 'Test'.")

# Variable para verificar si se encontraron correos no leídos
correos_no_leidos = False

# Recorrer los correos no leídos en la carpeta
for item in carpeta.Items:
    if item.UnRead and item.SenderEmailAddress == email_address:
        correos_no_leidos = True
        if item.Attachments.Count > 0:
            for attachment in item.Attachments:
                # Guardar archivos .pdf en el directorio de facturas
                if attachment.FileName.endswith('.pdf'):
                    attachment.SaveAsFile(os.path.join(directorio_facturas, attachment.FileName))
                    print(f"Archivo PDF guardado: {attachment.FileName}")
                # Guardar archivos .msg en el directorio de confirmación de mail
                elif attachment.FileName.endswith('.msg'):
                    attachment.SaveAsFile(os.path.join(directorio_confirmacion_mail, attachment.FileName))
                    print(f"Archivo MSG guardado: {attachment.FileName}")

if not correos_no_leidos:
    print("No hay correos no leídos del email indicado.")

# Función para enviar un correo con archivos adjuntos
def enviar_correo(directorio, destinatario):
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.CC = 'Srodriguez@bullcapital.cl;lguillen@bullcapital.cl;drodriguez@bullcapital.cl;atrejo@bullcapital.cl;hpavez@bullcapital.cl;lcollazos@bullcapital.cl;jrojas@bullcapital.cl;evillalobos@bullcapital.cl;fvillalobos@bullcapital.cl'
    mail.Subject = f'confirmacion {nombre_cliente}'
    mail.Body = 'por favor confirmar'
    
    # Adjuntar todos los archivos en el directorio
    for root, dirs, files in os.walk(directorio):
        for file in files:
            mail.Attachments.Add(os.path.join(root, file))
    
    mail.Send()
    print(f"Correo enviado a {destinatario} con los archivos del cliente {nombre_cliente}.")

# Enviar los archivos del cliente por correo
enviar_correo(directorio_cliente_fecha, 'omartinez@bullcapital.cl')

print("Descarga y envío completados.")