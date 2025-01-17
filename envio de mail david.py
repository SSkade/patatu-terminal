import os
import win32com.client as win32
from datetime import datetime

# Obtener el directorio base del usuario actual
base_dir = os.path.expanduser("~/Documents")

# Obtener la fecha de hoy en formato YYYYMMDD
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Ruta del archivo Excel
archivo_excel = os.path.join(base_dir, f'Excel david_{fecha_hoy}', f'reporte_{fecha_hoy}.xlsx')

# Enviar el archivo Excel por correo electrónico usando Outlook
def enviar_correo(archivo):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'daraya@bullcapital.cl'  # Reemplaza con la dirección de correo del destinatario
    mail.Subject = 'Archivo Excel Resultante'
    mail.Body = 'Adjunto encontrarás el archivo Excel resultante de la operación.'
    mail.Attachments.Add(archivo)
    mail.Send()
    print(f"Correo enviado con el archivo {archivo}")

# Verificar si el archivo Excel existe y enviarlo por correo
if os.path.exists(archivo_excel):
    enviar_correo(archivo_excel)
else:
    print(f"El archivo {archivo_excel} no existe.")

print("Operación completada.")