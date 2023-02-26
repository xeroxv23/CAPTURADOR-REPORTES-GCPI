import openpyxl
import numpy as np

# Abre el archivo de Excel original
wb1 = openpyxl.load_workbook('/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_08/SEMANA_8_REPORTE.xlsx')

# Selecciona la hoja de trabajo que contiene los datos que quieres copiar
hoja_origen = wb1.worksheets[0]

# Obtiene el valor actual de la celda A5
celda_origen = hoja_origen['A5']
valor_origen = celda_origen.value

# Define el valor umbral para buscar celdas con valor bajo
valor_umbral = valor_origen

# Crea una lista para almacenar los valores bajos encontrados
valores_bajos = [valor_origen]

# Obtiene todas las celdas debajo de A5 y verifica si su valor es bajo
for fila in hoja_origen.iter_rows(min_row=6, min_col=1, max_col=1):
    celda = fila[0]
    if celda.value is not None and celda.value <= valor_umbral:
        valores_bajos.append(celda.value)
    else:
        break  # rompe el ciclo si no hay un valor debajo de la celda A5

print(valores_bajos)

# Obtiene el valor actual de la celda D5
celda_origen_d5 = hoja_origen['D5']
valor_origen_d5 = celda_origen_d5.value * 7/6

# Define el valor umbral para buscar celdas con valor bajo
valor_umbral_d5 = valor_origen_d5

# Crea una lista para almacenar los valores bajos encontrados
valores_bajos_d5 = [valor_origen_d5]

# Obtiene todas las celdas debajo de D5 y verifica si su valor es bajo
for fila in hoja_origen.iter_rows(min_row=6, min_col=4, max_col=4):
    celda = fila[0]
    if celda.value is not None and celda.value <= valor_umbral:
        valores_bajos_d5.append(celda.value * 7/6)

print(valores_bajos_d5)

# Abrir el archivo Excel
workbook = openpyxl.load_workbook('/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_08/A002 Mantenimientos Av España.xlsm')

# Seleccionar la hoja de trabajo que deseas escribir
worksheet = workbook.active

# Seleccionar el primer valor del array
primer_valor = valores_bajos[0]
segundo_valor = valores_bajos_d5[0]
porcentaje = 1
numero_de_origen = hoja_origen['']

# Escribir el primer valor en la celda A18
worksheet['A18'] = primer_valor
worksheet['L18'] = segundo_valor
worksheet['S18'] = porcentaje

# Guardar el archivo Excel
workbook.save('/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_08/A002 Mantenimientos Av España.xlsm')

print('Se guardo el archivo')



