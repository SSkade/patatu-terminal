import subprocess
import os

# Obtener la ruta del directorio actual
base_dir = os.path.dirname(os.path.abspath(__file__))

# Rutas de los scripts
ruta_descarga_outlook = os.path.join(base_dir, 'descarga outlook.py')
ruta_descarga_mails = os.path.join(base_dir, 'descarga de mails.py')
ruta_scraping = os.path.join(base_dir, 'Scraping.py')
ruta_envio_mail = os.path.join(base_dir, 'envio de mail.py')

# Ejecutar el script de descarga outlook
print("Ejecutando descarga outlook...")
subprocess.run(['python', ruta_descarga_outlook], check=True)
print("Descarga outlook completada.")

# Ejecutar el script de descarga de mails
print("Ejecutando descarga de mails...")
subprocess.run(['python', ruta_descarga_mails], check=True)
print("Descarga de mails completada.")

# Ejecutar el script de scraping
print("Ejecutando scraping...")
subprocess.run(['python', ruta_scraping], check=True)
print("Scraping completado.")

# Ejecutar el script de envío de mail
print("Ejecutando envío de mail...")
subprocess.run(['python', ruta_envio_mail], check=True)
print("Envío de mail completado.")
