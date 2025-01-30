import os
from datetime import datetime
import extract_msg

# Obtener la fecha de hoy
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Directorio donde se encuentran los archivos .msg
correos_folder = 'C:\\Users\\Lucian\\Documents\\EMMA\\Correos'

# Directorio donde se guardarán los archivos .pdf
pdfs_folder = 'C:\\Users\\Lucian\\Documents\\EMMA\\pdfs'

# Verificar si el directorio de destino existe
if not os.path.exists(pdfs_folder):
    os.makedirs(pdfs_folder)
    print(f"Se ha creado el directorio de destino: {pdfs_folder}")
else:
    print(f"Utilizando el directorio de destino existente: {pdfs_folder}")

# Recorrer todos los archivos .msg en el directorio de origen
for archivo in os.listdir(correos_folder):
    if archivo.endswith('.msg'):
        file_path = os.path.join(correos_folder, archivo)
        msg = extract_msg.Message(file_path)

        # Variable para verificar si se encontró un archivo PDF adjunto
        pdf_found = False

        # Recorrer los archivos adjuntos
        for attachment in msg.attachments:
            if attachment.longFilename and attachment.longFilename.endswith('.pdf'):
                # Guardar el archivo adjunto en el directorio de destino
                try:
                    attachment.save(customPath=pdfs_folder)
                    print(f"Archivo {attachment.longFilename} guardado en {pdfs_folder}")
                    pdf_found = True
                except Exception as e:
                    print(f"Error al guardar el archivo {attachment.longFilename}: {e}")

        if not pdf_found:
            print(f"El correo {archivo} no tenía archivos PDF adjuntos.")

print("Descarga de archivos .pdf completada.")