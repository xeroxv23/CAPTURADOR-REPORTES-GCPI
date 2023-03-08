import openpyxl
import os
import re
from capturador import datos_de_captura

num_semana = 8
print(datos_de_captura)

def prueba_captura(trabajador):
    clave_de_obra = datos_de_captura[trabajador][1]
    # obtener la ruta de búsqueda
    ruta_busqueda = f'/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_0{num_semana}'

    # buscar archivos en la ruta de búsqueda que inicien con el valor de búsqueda
    for archivo in os.listdir(ruta_busqueda):
        if archivo.startswith(clave_de_obra + " ") and os.path.isfile(os.path.join(ruta_busqueda, archivo)):
            # si se encuentra el archivo, regresar la ruta completa
            archivo_para_captura = os.path.join(ruta_busqueda, archivo)
    
    # Cargamos el archivo_para_captura de Excel
    wb = openpyxl.load_workbook(archivo_para_captura)
    # Seleccionamos la hoja en la que queremos buscar
    ws = wb.active

    # Inicializar las variables que almacenarán la celda con el último valor encontrado
    ultima_celda_b = None
    ultima_celda_d = None

    # Recorrer las filas del rango especificado
    for fila in range(14, 301):
        # Obtener el valor de la columna B en la fila actual
        valor_b = ws.cell(row=fila, column=2).value
        # Si el valor es un número menor a 70, lo almacenamos
        if isinstance(valor_b, (int, float)) and valor_b < 70:
            ultima_celda_b = ws.cell(row=fila, column=2)

        # Obtener el valor de la columna D en la fila actual
        valor_d = ws.cell(row=fila, column=4).value
        # Si el valor es un string, lo almacenamos
        if isinstance(valor_d, str):
            ultima_celda_d = ws.cell(row=fila, column=4)

        # Obtener la celda para captura
        if ultima_celda_b is not None and ultima_celda_d is not None:
            # Si se encontró una celda en ambas columnas, seleccionar la que tenga el row mayor
            ultima_celda = ultima_celda_b if ultima_celda_b.row > ultima_celda_d.row else ultima_celda_d
        elif ultima_celda_b is not None:
            # Si solo se encontró una celda en la columna B, usar esa celda
            ultima_celda = ultima_celda_b
        elif ultima_celda_d is not None:
            # Si solo se encontró una celda en la columna D, usar esa celda
            ultima_celda = ultima_celda_d
        else:
            # Si no se encontró ninguna celda, seleccionar la celda en la fila 15, columna 1
            ultima_celda = ws.cell(row=15, column=1)
        
        # Reemplazar a que siempre sea la columna A
        ultima_celda.column = 1
       # Obtener el número de fila y columna de la celda
        fila_actual = ultima_celda.row
        columna_actual = ultima_celda.column
        # Sumar 1 al número de fila
        nueva_fila = fila_actual + 1
        # Crear una nueva instancia de la clase Cell con la misma columna y la nueva fila
        celda_para_captura = ultima_celda.parent.cell(row=nueva_fila, column=columna_actual)
        

    return archivo_para_captura, ultima_celda_b, ultima_celda_d, celda_para_captura
        
# Llamar a la funcion
archivo_para_captura, ultima_celda_b, ultima_celda_d, celda_para_captura = prueba_captura(0)

# Imprimir las coordenadas de las últimas celdas encontradas
print(archivo_para_captura)
print('La última celda en la columna B es:', ultima_celda_b)
print('La última celda en la columna D es:', ultima_celda_d)
print(celda_para_captura.coordinate)

        
