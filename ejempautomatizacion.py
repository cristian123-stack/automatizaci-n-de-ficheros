import pandas as pd
import PyPDF2  # Instalar previamente con pip install PyPDF2
import glob
import os
import shutil  # Para mover los archivos

# 1. Automatización Lectura Ficheros Excel (Ventas)
# Define la ruta relativa a los archivos de entrada
ruta_excel = './Inputs/Ventas_productos_automóvil.xlsx'
df_ventas = pd.read_excel(ruta_excel)

# Ver las primeras filas del DataFrame
print(df_ventas.head())

# 2. Automatización Lectura Ficheros PDF (Facturas)
# Define la ruta relativa para los ficheros PDF
ruta_pdfs = './Inputs/*.pdf'

# Obtener lista de ficheros PDF en el directorio 'Inputs'
lista_ficheros_pdf = glob.glob(ruta_pdfs)

# Crear la ruta de destino (si no existe, se crea)
ruta_destino = './orden'
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

        # Extraer texto de la primera página
        pag = lector.pages[0]
        texto = pag.extract_text()
        print(f"Texto extraído de la primera página:\n{texto}")

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
