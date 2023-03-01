import openpyxl

def obtener_ultimo_valor(ruta_archivo_destino):
    # Cargamos el archivo Excel
    libro = openpyxl.load_workbook(ruta_archivo_destino)
    
    # Seleccionamos la primera hoja del libro
    hoja = libro.active
    
    # Empezamos a buscar desde la fila 14
    fila_actual = 14
    
    # Inicializamos el valor a devolver con None
    ultimo_valor = None
    
    # Recorremos todas las filas de la hoja hasta encontrar un valor numérico menor a 70
    while fila_actual <= 300:
        celda_b = hoja.cell(row=fila_actual, column=2)
        valor_b = celda_b.value
        
        # Si la celda B de la fila actual tiene un valor numérico, lo guardamos como último valor
        if isinstance(valor_b, (int, float)):
            if valor_b < 70 and (ultimo_valor is None or valor_b > ultimo_valor):
                ultimo_valor = valor_b
        
        fila_actual += 1
    
    # Cerramos el libro de Excel
    libro.close()
    
    # Si no se encontró ningún valor menor a 70, se devuelve 1
    if ultimo_valor is None:
        ultimo_valor = 1
    
    # Devolvemos el último valor encontrado
    return ultimo_valor

ruta_archivo_destino = "/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_08/A-002B Mantenimientos Av España Apoyo Vicky .xlsm"
ultimo_valor = obtener_ultimo_valor(ruta_archivo_destino)
print("El último valor encontrado en la columna B es:", ultimo_valor)








    







