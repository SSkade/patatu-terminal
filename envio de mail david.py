import os
from datetime import datetime
import win32com.client as win32

# Directorio donde se encuentra el archivo Excel
base_dir = r"C:\Users\Lucian\Desktop\david"

# Obtener la fecha de hoy en formato YYYYMMDD
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Ruta del archivo Excel
archivo_excel = os.path.join(base_dir, f'reporte_{fecha_hoy}.xlsx')

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

# Enviar el correo con el archivo Excel
enviar_correo(archivo_excel)