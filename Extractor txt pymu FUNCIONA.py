import os
import fitz  # PyMuPDF

# Directorio de los PDFs
pdf_directory = r"C:\Users\Lucian\Documents\EMMA\pdfs"
# Directorio donde se guardarán los archivos de texto
output_directory = r"C:\Users\Lucian\Documents\EMMA\txts"

# Crear el directorio de salida si no existe
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

def open_pdf(file_path):
    """Abrir un archivo PDF y devolver el objeto del documento."""
    try:
        doc = fitz.open(file_path)
        return doc
    except Exception as e:
        print(f"Error al abrir el archivo PDF: {e}")
        return None

def read_pdf(file_path):
    """Leer el contenido de un archivo PDF."""
    doc = open_pdf(file_path)
    if doc:
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text

# Iterar sobre todos los archivos en el directorio de PDFs
for filename in os.listdir(pdf_directory):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(pdf_directory, filename)
        txt_path = os.path.join(output_directory, filename.replace(".pdf", ".txt"))

        # Leer el PDF y extraer el texto
        text = read_pdf(pdf_path)

        # Guardar el texto en un archivo .txt
        with open(txt_path, "w", encoding="latin-1") as txt_file:
            txt_file.write(text)

print("Extracción completada.")