import pandas as pd
import PyPDF2  # Instalar previamente con pip install PyPDF2
import glob
import os
import shutil  # Para mover los archivos
import pytesseract  # Instalar previamente con pip install pytesseract
from pdf2image import convert_from_path  # Instalar previamente con pip install pdf2image
from pathlib import Path

# 1. Automatización Lectura Ficheros Excel (Ventas)
# Cambié la ruta para que sea relativa. Asegúrate de tener el archivo Excel en el mismo directorio que el script o en una subcarpeta.
df_ventas = pd.read_excel(os.path.join(os.getcwd(), 'Inputs', 'Ventas_productos_automóvil.xlsx'))

# Ver las primeras filas del DataFrame
print(df_ventas.head())

# 2. Automatización Lectura Ficheros PDF (Facturas)
# Obtener lista de ficheros PDF en el directorio 'Inputs'
lista_ficheros_pdf = glob.glob(os.path.join(os.getcwd(), 'Inputs', '*.pdf'))

# Crear la ruta de destino (si no existe, se crea) de forma relativa
ruta_destino = os.path.join(os.getcwd(), 'orden')
if not os.path.exists(ruta_destino):
    os.makedirs(ruta_destino)

# Bucle para recorrer todos los ficheros PDF
for fichero_pdf in lista_ficheros_pdf:
    # Abrir archivo PDF en formato binario
    with open(fichero_pdf, 'rb') as pdfFile:
        # Crear objeto lector para el PDF
        lector = PyPDF2.PdfReader(pdfFile)

        # Obtener número de páginas del PDF
        print(f"\nNúmero de páginas del PDF '{fichero_pdf}': {len(lector.pages)}")

        # Intentar extraer texto de la primera página (si es texto accesible)
        pag = lector.pages[0]
        texto = pag.extract_text()
        
        if texto:
            print(f"Texto extraído de la primera página:\n{texto}")
        else:
            print("No se pudo extraer texto directamente del PDF. Procediendo con OCR...")

            # Convertir la página PDF en imagen
            images = convert_from_path(fichero_pdf)
            for i, image in enumerate(images):
                # Utilizar pytesseract para hacer OCR en la imagen
                texto_imagen = pytesseract.image_to_string(image)
                print(f"Texto extraído de la imagen (Página {i + 1}):\n{texto_imagen}")

                # Concatenar el texto extraído de las imágenes
                texto = texto + texto_imagen

        # Identificar donde está el primer salto de línea para obtener el nombre del proveedor
        indice_final = texto.find('\n')
        proveedor_factura = texto[0:indice_final]
        print(f"Proveedor de '{fichero_pdf}': {proveedor_factura}")

        # Limpiar el nombre del proveedor para evitar caracteres no permitidos en nombres de archivo
        nombre_destino = proveedor_factura.strip().replace(" ", "_").replace("é", "e").replace("í", "i").replace("á", "a")  # Limpiar acentos y espacios

        # Crear la carpeta del proveedor si no existe
        carpeta_proveedor = os.path.join(ruta_destino, nombre_destino)
        if not os.path.exists(carpeta_proveedor):
            os.makedirs(carpeta_proveedor)
            print(f"Carpeta creada para el proveedor: {carpeta_proveedor}")

        # Crear el nombre del archivo de destino dentro de la carpeta del proveedor
        nombre_archivo_destino = os.path.join(carpeta_proveedor, f"{nombre_destino}_Factura.pdf")

        # Verificar si el archivo de destino ya existe (para evitar sobreescribirlo)
        contador = 1
        while os.path.exists(nombre_archivo_destino):
            nombre_archivo_destino = os.path.join(carpeta_proveedor, f"{nombre_destino}_Factura_{contador}.pdf")
            contador += 1

        # Copiar el archivo a la nueva ubicación
        shutil.copy(fichero_pdf, nombre_archivo_destino)  # Copiar el archivo al destino
        print(f"Archivo '{fichero_pdf}' copiado a '{nombre_archivo_destino}'")

        # Eliminar el archivo original solo si la copia fue exitosa
        try:
            os.remove(fichero_pdf)
            print(f"Archivo original '{fichero_pdf}' eliminado.")
        except Exception as e:
            print(f"Error al eliminar el archivo original '{fichero_pdf}': {e}")

