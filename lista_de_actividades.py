# LIBRERIAS
import openpyxl

# VARIABLES
archivo_para_captura = '/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_08/A-002 Mantenimientos Av España.xlsm'

lista_de_actividades = ["Esta es la primera lista que pretende tener mas de 45 caracteres y sera dividida en una sublista", "This is the second array that whants to be a long characters array in order to save the world"]
nueva_lista = [[] for i in range(len(lista_de_actividades))]

for i, actividad in enumerate(lista_de_actividades):
    # Dividir la cadena de texto en subcadenas de máximo 46 caracteres
    subcadenas = []
    while len(actividad) > 0:
        if len(actividad) <= 46:
            subcadenas.append(actividad)
            actividad = ""
        else:
            espacio = actividad.rfind(" ", 0, 46)
            if espacio == -1:
                subcadenas.append(actividad[:46])
                actividad = actividad[46:]
            else:
                subcadenas.append(actividad[:espacio])
                actividad = actividad[espacio+1:]
    
    # Agregar las subcadenas a la nueva lista correspondiente
    nueva_lista[i].extend(subcadenas)


def capturar_actividades(lista):
    # Cargamos el archivo de Excel que contiene la celda que queremos capturar
    wb = openpyxl.load_workbook(archivo_para_captura)
    # Seleccionamos la hoja en la que queremos buscar
    ws = wb.active

    # Crear la celda del codigo y asignarle la clave del trabajador
    nueva_fila = 16
    nueva_columna = 4
    for valor in lista:
        celda_codigo = ws.cell(row=nueva_fila, column=nueva_columna)
        celda_codigo.value = valor
        nueva_fila += 1
    
    # Guardar el archivo de Excel
    print("se capturaron los datos")
    wb.save(archivo_para_captura)

capturar_actividades(nueva_lista[0])