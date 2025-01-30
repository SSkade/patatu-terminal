import os
import pandas as pd
from datetime import datetime

# Obtener el directorio base del usuario actual
base_dir = os.path.expanduser("~/Desktop")

# Obtener la fecha de hoy en formato YYYYMMDD
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Configuración de los directorios de origen
david_folder = os.path.join(base_dir, "david", f'TXT_{fecha_hoy}')
hemel_folder = os.path.join(base_dir, "hemel", f'TXT_{fecha_hoy}')
nico_folder = os.path.join(base_dir, "nico", f'TXT_{fecha_hoy}')

# Lista de directorios de origen
directorios_origen = [david_folder, hemel_folder, nico_folder]

# Función para extraer información de un archivo .txt
def extraer_informacion(file_path):
    with open(file_path, 'r', encoding='latin-1') as file:
        lines = file.readlines()
    
    info = {
        'Emisor': '',
        'Receptor': '',
        'Cedido por': '',
        'Cedido a': '',
        'eMail': ''
    }
    
    for line in lines:
        if 'Emisor  ' in line:
            info['Emisor'] = line.split(':')[1].strip()
        elif 'Receptor' in line:
            info['Receptor'] = line.split(':')[1].strip()
        elif 'Cedido por' in line:
            info['Cedido por'] = line.split(':')[1].strip()
        elif 'Cedido a' in line:
            info['Cedido a'] = line.split(':')[1].strip()
        elif 'eMail:' in line and info['eMail'] == '':
            info['eMail'] = line.split('eMail:')[1].strip().split()[0]
    
    return info

# Recorrer todos los archivos .txt en los directorios de origen
for directorio_origen in directorios_origen:
    datos = []
    if os.path.exists(directorio_origen):
        for archivo in os.listdir(directorio_origen):
            if archivo.endswith('.txt'):
                file_path = os.path.join(directorio_origen, archivo)
                info = extraer_informacion(file_path)
                datos.append(info)
        
        # Crear un DataFrame con la información extraída
        df = pd.DataFrame(datos)

        # Mostrar el DataFrame
        print(df)

        # Determinar la carpeta principal (david, hemel o nico)
        if "david" in directorio_origen:
            directorio_excel = os.path.join(base_dir, "david")
        elif "hemel" in directorio_origen:
            directorio_excel = os.path.join(base_dir, "hemel")
        elif "nico" in directorio_origen:
            directorio_excel = os.path.join(base_dir, "nico")

        # Verificar si el directorio de Excel existe, si no, crearlo
        if not os.path.exists(directorio_excel):
            os.makedirs(directorio_excel)

        # Guardar el DataFrame en un archivo Excel (.xlsx)
        archivo_excel = os.path.join(directorio_excel, f'reporte_{fecha_hoy}.xlsx')
        df.to_excel(archivo_excel, index=False, engine='openpyxl')

        # Guardar el DataFrame en un archivo CSV (opcional)
        # archivo_csv = os.path.join(directorio_excel, f'reporte_{fecha_hoy}.csv')
        # df.to_csv(archivo_csv, index=False)

print("Procesamiento completado.")