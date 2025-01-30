import os
from datetime import datetime
import extract_msg

# Obtener la fecha de hoy
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Obtener el directorio base del usuario actual
base_dir = os.path.expanduser("~/Desktop")

# Configuración de los directorios de origen
david_folder = os.path.join(base_dir, "david", f"correos_{fecha_hoy}")
hemel_folder = os.path.join(base_dir, "hemel", f"correos_{fecha_hoy}")
nico_folder = os.path.join(base_dir, "nico", f"correos_{fecha_hoy}")

# Lista de directorios de origen
directorios_origen = [david_folder, hemel_folder, nico_folder]

# Recorrer todos los archivos .msg en los directorios de origen
for directorio_origen in directorios_origen:
    if os.path.exists(directorio_origen):
        # Determinar la carpeta principal (david, hemel o nico)
        if "david" in directorio_origen:
            directorio_destino = os.path.join(base_dir, "david", f'TXT_{fecha_hoy}')
        elif "hemel" in directorio_origen:
            directorio_destino = os.path.join(base_dir, "hemel", f'TXT_{fecha_hoy}')
        elif "nico" in directorio_origen:
            directorio_destino = os.path.join(base_dir, "nico", f'TXT_{fecha_hoy}')
        
        # Verificar si el directorio de destino existe
        if not os.path.exists(directorio_destino):
            os.makedirs(directorio_destino)
            print(f"Se ha creado el directorio de destino: {directorio_destino}")
        else:
            print(f"Utilizando el directorio de destino existente: {directorio_destino}")

        for archivo in os.listdir(directorio_origen):
            if archivo.endswith('.msg'):
                file_path = os.path.join(directorio_origen, archivo)
                msg = extract_msg.Message(file_path)

                # Recorrer los archivos adjuntos
                for attachment in msg.attachments:
                    if attachment.longFilename.endswith('.txt'):
                        # Guardar el archivo adjunto en el directorio de destino
                        attachment.save(customPath=directorio_destino)
                        print(f"Archivo {attachment.longFilename} guardado en {directorio_destino}")

print("Descarga de archivos .txt completada.")