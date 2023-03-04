# LIBRERIAS
""" En esta sección, se importan las librerías necesarias para el funcionamiento del script. openpyxl se utiliza para cargar y procesar archivos de Excel, glob para buscar archivos en una ruta específica y os para manipular rutas de archivos. """

import openpyxl
import glob
import os

# VARIABLES_GLOBALES
""" VARIABLES_GLOBALES: Se definen tres variables globales: num_semana, ruta_proyecto y nombre_archivo. Estas variables se utilizan para construir las rutas de los archivos de origen y destino que se procesarán. También se definen las variables ruta_archivo_origen y ruta_archivo_destino a partir de las variables globales. """

num_semana = 8
ruta_proyecto = '/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/'
nombre_archivo = f'SEMANA_0{num_semana}_REPORTE.xlsx'

ruta_archivo_origen = os.path.join(ruta_proyecto, f'SEMANA_0{num_semana}', nombre_archivo)
ruta_archivo_destino = os.path.join(ruta_proyecto, f'SEMANA_0{num_semana}')

# FUNCIONES


def obtencion_datos_reporte(ruta_de_archivo):
    """  Esta función toma un argumento llamado ruta_de_archivo, que se utiliza para cargar el archivo de Excel correspondiente. La función carga el archivo utilizando openpyxl y extrae ciertos datos de la hoja activa del archivo. En particular, la función extrae los valores de las columnas A, C, D, E, F y G a partir de la fila 5. Luego, se multiplica el valor de la columna D por 7/6 y se agrega a la lista de datos. La función continúa extrayendo datos de las filas siguientes hasta que no haya más datos en la columna A. Por último, la función devuelve la lista de datos extraídos. """

    # Cargamos el archivo de Excel con openpyxl
    libro = openpyxl.load_workbook(ruta_archivo_origen)
    hoja = libro.active

    # Definimos la lista donde vamos a guardar los datos
    datos_de_captura = []

    # Empezamos a leer desde la fila 5
    fila = 5

    # Iteramos mientras haya datos en la columna A
    while hoja.cell(row=fila, column=1).value:
        # Extraemos los valores de las columnas A, C, D, E, F, G, I, J
        valores_fila = [hoja.cell(row=fila, column=columna).value for columna in range(1, 11) if columna in [1, 3, 4, 5, 6, 7, 9, 10]]
        
        # Multiplicamos el valor de la columna D por 7/6
        valores_fila[2] = valores_fila[2] * 7/6

        # Agregamos los valores a la lista de datos
        datos_de_captura.append(valores_fila)

        # Avanzamos a la siguiente fila
        fila += 1

    # Devolvemos la lista de datos
    return datos_de_captura

# Creamos la lista de datos_de_captura, con la cual trabajaremos
datos_de_captura = obtencion_datos_reporte(ruta_archivo_origen)

# Creamos una sublista que representa las actividades de cada trabajador
actividades = []
for sublista in datos_de_captura:
    actividad = sublista[5]
    actividades.append(actividad)

lista_de_actividades = [[] for i in range(len(actividades))]

for i, actividad in enumerate(actividades):
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
    lista_de_actividades[i].extend(subcadenas)


# La lista trabajadores, contendra las claves de cada uno de los trabajadores en datos_de_captura
trabajadores = [lista[0] for lista in datos_de_captura]

# Este ciclo for llenara la lista trabajador, enumerando a trabajadores iniciando desde el 0, para poder usarla como parametro en nuestra variable
trabajador = []
for i in trabajadores:
    count = 0
    for j in trabajadores:
        if j < i:
            count += 1
    trabajador.append(count)
trabajador = sorted(trabajador)


def capturar_reporte_personal(trabajador):
    """ Esta funcion realiza una búsqueda en un archivo de Excel que contiene un registro de reporte personal de un trabajador específico. El objetivo es buscar el último valor numérico menor a 70 en la columna B, guardar la coordenada de la celda y devolverla en la lista datos_capturados. En caso de no encontrar ningún valor numérico menor a 70, se devolverá la coordenada de la celda A15.

    Después de encontrar la coordenada de la celda, la función realiza varias tareas adicionales, como cargar el archivo de Excel, asignar valores a celdas específicas y guardar el archivo con los cambios. """

    datos_capturados = []
    valor = 1

    # obtener el valor de búsqueda
    clave_obra = datos_de_captura[trabajador][valor]

    # obtener la ruta de búsqueda
    ruta_busqueda = f'/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/SEMANA_0{num_semana}'

    # buscar archivos en la ruta de búsqueda que inicien con el valor de búsqueda
    for archivo in os.listdir(ruta_busqueda):
        if archivo.startswith(clave_obra):
            # si se encuentra el archivo, regresar la ruta completa
            archivo_para_captura = os.path.join(ruta_busqueda, archivo)
            datos_capturados.append(archivo_para_captura)
            break
    
     # si no se encontró ningún archivo, regresar None
    if archivo_para_captura is None:
        return None
    
    # Cargamos el archivo_para_captura de Excel
    wb = openpyxl.load_workbook(archivo_para_captura)
    # Seleccionamos la hoja en la que queremos buscar
    ws = wb.active
    # Inicializamos las variables para almacenar la última celda con texto en la columna 4
    ult_celda_con_texto = None
    # Recorremos las filas desde la 14 hasta la 300
    for fila in range(14, 301):
        # Obtenemos el valor de la celda D en la fila actual
        valor_celda = ws.cell(row=fila, column=4).value
        # Si el valor es un string, lo almacenamos
        if isinstance(valor_celda, str):
            ult_celda_con_texto = ws.cell(row=fila, column=4).coordinate
    # Si encontramos un valor con texto, retornamos la coordenada de la última celda encontrada
    if ult_celda_con_texto is not None:
        celda = ws[ult_celda_con_texto]
        nueva_fila = celda.row + 2
        nueva_columna = celda.column - 3
        celda_para_captura = ws.cell(row=nueva_fila, column=nueva_columna)

        datos_capturados.append(celda_para_captura.coordinate)
    # Si no encontramos ningún valor con texto, retornamos la celda A15
    else:
        nueva_fila = 15
        nueva_columna = 1
        celda_para_captura = ws.cell(row=nueva_fila, column=nueva_columna).coordinate
        datos_capturados.append(celda_para_captura)

    # Empezamos a buscar desde la fila 14
    fila_actual = 14
    
    # Inicializamos el valor a devolver con None
    ultimo_valor = None
    
    # Recorremos todas las filas de la hoja hasta encontrar un valor numérico menor a 70
    while fila_actual <= 300:
        celda_b = ws.cell(row=fila_actual, column=2)
        valor_b = celda_b.value
        
        # Si la celda B de la fila actual tiene un valor numérico, lo guardamos como último valor
        if isinstance(valor_b, (int, float)):
            if valor_b < 70 and (ultimo_valor is None or valor_b > ultimo_valor):
                ultimo_valor = valor_b
        
        fila_actual += 1
    
    # Cerramos el libro de Excel
    wb.close()
    
    # Si no se encontró ningún valor menor a 70, se devuelve 1
    if ultimo_valor is None:
        ultimo_valor = 0
    datos_capturados.append(ultimo_valor)

    # Cargamos el archivo de Excel que contiene la celda que queremos capturar
    wb = openpyxl.load_workbook(archivo_para_captura)
    # Seleccionamos la hoja en la que queremos buscar
    ws = wb.active

    # Crear la celda del codigo y asignarle la clave del trabajador
    celda_codigo = ws.cell(row=nueva_fila, column=nueva_columna)
    celda_codigo.value = datos_de_captura[trabajador][valor - 1]

    # Crear la celda de horas extras y domingo, solo si existen y asignarles el valor de hora extra
    if datos_de_captura[trabajador][valor + 2] and datos_de_captura[trabajador][valor + 3] is not None:
        # Código para crear las celdas de horas extras y domingo para asignarles los valores
        celda_horas = ws.cell(row=nueva_fila + 1, column=nueva_columna)
        celda_horas.value = datos_de_captura[trabajador][valor + 6]
        celda_horas2 = ws.cell(row=nueva_fila + 1, column=12)
        celda_horas2.value = datos_de_captura[trabajador][valor + 2]
        celda_domingo = ws.cell(row=nueva_fila + 2, column=nueva_columna)
        celda_domingo.value = datos_de_captura[trabajador][valor + 5]
        celda_domingo2 = ws.cell(row=nueva_fila + 2, column=12)
        celda_domingo2.value = int(datos_de_captura[trabajador][valor + 3]) + 1
        celda_domingo3 = ws.cell(row=nueva_fila + 2, column=4)
        celda_domingo3.value = "Domingo trabajado"

    elif datos_de_captura[trabajador][valor + 2] is not None:
        celda_horas = ws.cell(row=nueva_fila + 1, column=nueva_columna)
        celda_horas.value = datos_de_captura[trabajador][valor + 6]
        celda_horas2 = ws.cell(row=nueva_fila + 1, column=12)
        celda_horas2.value = datos_de_captura[trabajador][valor + 2]

    # Crear las celdas de orden y asignarle el valor ultimo valor +1
    
    if datos_de_captura[trabajador][valor + 2] and datos_de_captura[trabajador][valor + 3] is not None:
        celda_orden1 = ws.cell(row=nueva_fila, column=2)
        celda_orden1.value = ultimo_valor +1
        celda_orden2 = ws.cell(row=nueva_fila + 2, column=16)
        celda_orden2.value = ultimo_valor +1
        celda_ordenh = ws.cell(row=nueva_fila + 1, column=2)
        celda_ordenh.value = ultimo_valor +1
        celda_ordend = ws.cell(row=nueva_fila + 2, column=2)
        celda_ordend.value = ultimo_valor +1

    elif datos_de_captura[trabajador][valor + 2] is not None:
        celda_orden1 = ws.cell(row=nueva_fila, column=2)
        celda_orden1.value = ultimo_valor +1
        celda_orden2 = ws.cell(row=nueva_fila + 1, column=16)
        celda_orden2.value = ultimo_valor +1
        celda_ordenh = ws.cell(row=nueva_fila + 1, column=2)
        celda_ordenh.value = ultimo_valor +1
    else:
        celda_orden1 = ws.cell(row=nueva_fila, column=2)
        celda_orden1.value = ultimo_valor +1
        celda_orden2 = ws.cell(row=nueva_fila, column=16)
        celda_orden2.value = ultimo_valor +1

    # Crear las celdas de dias trabajados y asignarles el valor de datos_de_captura[0][2]
    celda_dias = ws.cell(row=nueva_fila, column=12)
    celda_dias.value = datos_de_captura[trabajador][valor + 1]

    # Crear la celda de porcentaje y asignarles el valor de 1
    celda_porcentaje = ws.cell(row=nueva_fila, column=19)
    celda_porcentaje.value = 1

    # Crear la celda de actividades y asignarles el valor de lista_de_actividades
    if datos_de_captura[trabajador][valor + 2] and datos_de_captura[trabajador][valor + 3] is not None:
        
        for valor in lista_de_actividades[trabajador]:
            celda_actividades = ws.cell(row=nueva_fila +3, column=4)
            celda_actividades.value = valor
            nueva_fila += 1

    elif datos_de_captura[trabajador][valor + 2] is not None:

        for valor in lista_de_actividades[trabajador]:
            celda_actividades = ws.cell(row=nueva_fila +2, column=4)
            celda_actividades.value = valor
            nueva_fila += 1
    
    else:
        for valor in lista_de_actividades[trabajador]:
            celda_actividades = ws.cell(row=nueva_fila +1, column=4)
            celda_actividades.value = valor
            nueva_fila += 1
    
            
    # Guardar los cambios y retornar la coordenada de la nueva celda
    wb.save(archivo_para_captura)
    trabajador +1
    return datos_capturados

# El ciclo for que capturara todos los valores en los archivos de excel, recorriendo en el parametro de la funcion cada uno de los valores del reporte
for clave in trabajador:
    print("Se capturo en los archivos, el trabajador numero : ", clave)
    capturar_reporte_personal(clave)

print(f"Se ha terminado la captura del reporte :  SEMANA_0{num_semana} ")













   






