# LIBRERIAS
import os
import glob
import openpyxl
import pandas as pd
from experimento_2 import lista_claves

# VARIABLES GLOBALES
num_semana = 8
ruta_archivo_origen = f'/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_0{num_semana}/SEMANA_0{num_semana}_REPORTE.xlsx'
ruta_directorio = "/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_08"

# Obtener una lista de todos los archivos en el directorio
ruta_directorio = f"/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_0{num_semana}"
archivos_en_directorio = os.listdir(ruta_directorio)

# Buscar los nombres de archivo que comienzan con cada valor en la lista de claves
valores_celda = []
for valor_clave in lista_claves:
    nombre_archivo = None
    for archivo in archivos_en_directorio:
        if archivo.startswith(valor_clave):
            nombre_archivo = archivo
            break

    # Verificar si se encontró el archivo
    if nombre_archivo is None:
        print(f"No se encontró ningún archivo que comience con '{valor_clave}'.")
    else:
        # Construir la ruta completa del archivo
        ruta_archivo = os.path.join(ruta_directorio, nombre_archivo)

        # Abrir el archivo utilizando openpyxl y leer el valor de la celda B15
        wb = openpyxl.load_workbook(ruta_archivo)
        hoja = wb.active
        valor_celda = hoja.cell(row=15, column=2).value

        # Agregar el valor de la celda a la lista de valores de celda
        valores_celda.append(valor_celda)

# Imprimir la lista de valores de celda
print(valores_celda)