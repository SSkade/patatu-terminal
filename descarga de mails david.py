import os
from datetime import datetime
import extract_msg

# Obtener la fecha de hoy
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Obtener el directorio base del usuario actual
base_dir = os.path.expanduser("~/Documents")

# Configuración del directorio de origen y destino
directorio_origen = os.path.join(base_dir, f'Correosdavid_{fecha_hoy}')
directorio_destino = os.path.join(base_dir, f'TXT david_{fecha_hoy}')

# Verificar si el directorio de origen existe
if not os.path.exists(directorio_origen):
    os.makedirs(directorio_origen)
    print(f"Se ha creado el directorio de origen: {directorio_origen}")
else:
    print(f"Utilizando el directorio de origen existente: {directorio_origen}")

# Verificar si el directorio de destino existe
if not os.path.exists(directorio_destino):
    os.makedirs(directorio_destino)
    print(f"Se ha creado el directorio de destino: {directorio_destino}")
else:
    print(f"Utilizando el directorio de destino existente: {directorio_destino}")

# Recorrer todos los archivos .msg en el directorio de origen
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