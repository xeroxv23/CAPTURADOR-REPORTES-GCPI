# LIBRERIAS
import openpyxl
import os

# VARIABLES GLOBALES
num_semana = 8
ruta_proyecto = '/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/'
nombre_archivo = f'SEMANA_0{num_semana}_REPORTE.xlsx'

ruta_archivo_origen = os.path.join(ruta_proyecto, f'SEMANA_0{num_semana}', nombre_archivo)
ruta_archivo_destino = os.path.join(ruta_proyecto, f'SEMANA_0{num_semana}')

# GENERAR LOS DATOS DE CAPTURA
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


# GENERAMOS LA LISTA DE ACTIVIDADES
# Creamos una sublista que representa las actividades de cada trabajador
actividades = []
for sublista in datos_de_captura:
    actividad = sublista[5]
    if actividad is None or actividad == "":
        actividad = "" # Si actividad es None o una cadena vacía, asignamos una cadena vacía
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

# GENERAMOS LA LISTA DE TRABAJADORES
# La lista trabajadores, contendra las claves de cada uno de los trabajadores en datos_de_captura
trabajadores = [lista[0] for lista in datos_de_captura]

# Este ciclo for llenara la lista trabajador, enumerando a trabajadores iniciando desde el 0, para poder usarla como parametro en nuestra variable
# Enumeramos los elementos de la lista y guardamos los índices en una lista
trabajador = [i for i, _ in enumerate(trabajadores)]

# FUNCION PARA CAPTURAR EN EL ARCHIVO EXCEL: 
""" Esta funcion realiza una búsqueda en un archivo de Excel que contiene un registro de reporte personal de un trabajador específico. El objetivo es buscar el último valor numérico menor a 70 en la columna B, guardar la coordenada de la celda y devolverla en la lista datos_capturados. En caso de no encontrar ningún valor numérico menor a 70, se devolverá la coordenada de la celda A15.

Después de encontrar la coordenada de la celda, la función realiza varias tareas adicionales, como cargar el archivo de Excel, asignar valores a celdas específicas y guardar el archivo con los cambios. """

def capturar_reporte_personal(trabajador):

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
        print(f"No se encontro la clave {clave_obra} para el trabajador no: {trabajador}")
        return None
    
    # Cargamos el archivo_para_captura de Excel
    wb = openpyxl.load_workbook(archivo_para_captura)
    # Seleccionamos la hoja en la que queremos buscar
    ws = wb.active

    # Inicializamos las variables para almacenar la última celda con un valor menor a 70
    ult_celda_con_valor = None
    # Recorremos las filas desde la 14 hasta la 300
    for fila in range(14, 301):
        # Obtenemos el valor de la celda B en la fila actual
        valor_celda = ws.cell(row=fila, column=2).value
        # Si el valor es un número menor a 70, lo almacenamos
        if isinstance(valor_celda, (int, float)) and valor_celda < 70:
            ult_celda_con_valor = ws.cell(row=fila, column=2).coordinate

    # Inicializamos las variables para almacenar la última celda con texto
    ult_texto = None
    ult_celda_con_texto = None

    # Recorremos las filas desde la 13 hasta la 301
    for fila in range(14, 301):
        # Obtenemos el valor de la celda D en la fila actual
        valor_celda = ws.cell(row=fila, column=4).value
        # Si el valor es un string, lo almacenamos
        if isinstance(valor_celda, str):
            ult_texto = valor_celda
            ult_celda_con_texto = ws.cell(row=fila, column=4).coordinate

    # Logica para asignar las celdas de captura dependiendo el ultimo valor
    if ult_texto is None and ult_celda_con_valor is None:
        nueva_fila = 15
        nueva_columna = 1
        celda_para_captura = ws.cell(row=nueva_fila, column=nueva_columna).coordinate
        celda_para_captura = ws.cell(row=nueva_fila, column=nueva_columna)

        datos_capturados.append(celda_para_captura)

    elif ult_texto == "Domingo trabajado":
        celda = ws[ult_celda_con_valor]
        nueva_fila = celda.row + 1
        nueva_columna = celda.column - 1
        celda_para_captura = ws.cell(row=nueva_fila, column=nueva_columna)

        datos_capturados.append(celda_para_captura.coordinate)

    elif ult_texto is not None and len(ult_texto) > 2:
        celda = ws[ult_celda_con_texto]
        nueva_fila = celda.row + 2
        nueva_columna = celda.column - 3
        celda_para_captura = ws.cell(row=nueva_fila, column=nueva_columna)

        datos_capturados.append(celda_para_captura.coordinate)
    else:
        celda = ws[ult_celda_con_valor]
        nueva_fila = celda.row +1
        nueva_columna = celda.column -1
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
clave_trabajadores = [sublista[0] for sublista in datos_de_captura]
lista_de_obras = [sublista[1] for sublista in datos_de_captura]

for clave in trabajador:
    print("Se ha capturado en la obra", lista_de_obras[clave], "el trabajador numero : ", clave_trabajadores[clave])
    capturar_reporte_personal(clave)

print(f"Se ha terminado la captura del reporte :  SEMANA_0{num_semana} ")
