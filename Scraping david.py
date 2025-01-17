import os
import pandas as pd
from datetime import datetime

# Obtener el directorio base del usuario actual
base_dir = os.path.expanduser("~/Documents")

# Obtener la fecha de hoy en formato YYYYMMDD
fecha_hoy = datetime.now().strftime('%Y-%m-%d')

# Directorio donde se encuentran los archivos .txt
directorio_txt = os.path.join(base_dir, f'TXT david_{fecha_hoy}')

# Verificar si el directorio de Excel existe, si no, crearlo
directorio_excel = os.path.join(base_dir, f'Excel david_{fecha_hoy}')
if not os.path.exists(directorio_excel):
    os.makedirs(directorio_excel)

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

# Lista para almacenar la información extraída
datos = []

# Recorrer todos los archivos .txt en el directorio
for archivo in os.listdir(directorio_txt):
    if archivo.endswith('.txt'):
        file_path = os.path.join(directorio_txt, archivo)
        info = extraer_informacion(file_path)
        datos.append(info)

# Crear un DataFrame con la información extraída
df = pd.DataFrame(datos)

# Mostrar el DataFrame
print(df)

# Guardar el DataFrame en un archivo Excel (.xlsx)
archivo_excel = os.path.join(directorio_excel, f'reporte_{fecha_hoy}.xlsx')
df.to_excel(archivo_excel, index=False, engine='openpyxl')

# Guardar el DataFrame en un archivo CSV
#archivo_csv = os.path.join(directorio_excel, f'reporte_{fecha_hoy}.csv')
#df.to_csv(archivo_csv, index=False)