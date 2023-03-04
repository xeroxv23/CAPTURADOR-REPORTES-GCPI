# LIBRERIAS
""" En esta sección, se importan las librerías necesarias para el funcionamiento del script. openpyxl se utiliza para cargar y procesar archivos de Excel, os para buscar archivos en una ruta específica y os para manipular rutas de archivos. """

import openpyxl
import glob
import os

# VARIABLES GLOBALES
""" VARIABLES_GLOBALES: Se definen tres variables globales: num_semana, ruta_proyecto y nombre_archivo. Estas variables se utilizan para construir las rutas de los archivos de origen y destino que se procesarán. También se definen las variables ruta_archivo_origen y ruta_archivo_destino a partir de las variables globales. """

num_semana = 8
ruta_proyecto = '/home/xeroxv23/Documents/Proyectos GCPI/reportes_personal_zonaindustrial/'
nombre_archivo = f'SEMANA_0{num_semana}_REPORTE.xlsx'

ruta_archivo_origen = os.path.join(ruta_proyecto, f'SEMANA_0{num_semana}', nombre_archivo)
ruta_archivo_destino = os.path.join(ruta_proyecto, f'SEMANA_0{num_semana}')

# GENERAR LOS DATOS DE CAPTURA
""" Este bloque de código carga un archivo de Excel utilizando la librería "openpyxl", después define una lista vacía llamada "datos_de_captura" donde se almacenarán los datos procesados. El código recorre las filas de la hoja de Excel a partir de la fila 5 hasta encontrar una fila vacía en la columna A.

Para cada fila de datos, el código extrae los valores de las columnas A, C, D, E, F, G, I y J y los almacena en una lista llamada "valores_fila". Luego, el valor en la tercera posición de "valores_fila" (correspondiente a la columna D) se multiplica por 7/6. Los valores procesados se agregan a la lista "datos_de_captura" y el código avanza a la siguiente fila. """

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
""" Este bloque de código genera una lista de actividades de los trabajadores a partir de la lista de datos previamente procesada en el bloque anterior.

En primer lugar, se crea una lista vacía llamada "actividades". Se recorre cada sublista de la lista "datos_de_captura" y se extrae el valor correspondiente a la columna F (es decir, la actividad realizada por cada trabajador) y se agrega a la lista "actividades".

Luego, se crea una lista vacía llamada "lista_de_actividades" que tendrá una sublista para cada actividad encontrada en la lista "actividades". El código recorre cada actividad de la lista "actividades" y divide su contenido en subcadenas de un máximo de 46 caracteres utilizando un ciclo "while". Las subcadenas se almacenan en una lista llamada "subcadenas".

Finalmente, se agrega cada subcadena a la sublista correspondiente de "lista_de_actividades" utilizando el método "extend()".  """

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

# GENERAMOS LA LISTA DE TRABAJADORES
""" Este bloque de código genera una lista de trabajadores a partir de los datos capturados en datos_de_captura. Cada trabajador tiene una clave única que se encuentra en la primera columna de los datos de captura.

La variable "trabajadores" se crea a partir de una comprensión de lista que extrae las claves de los trabajadores de los datos de captura.

Luego, se usa un bucle for para crear una lista "trabajador" donde se enumeran los trabajadores a partir de 0. La variable "count" se usa para contar el número de trabajadores con una clave menor que la clave actual, y esta cuenta se agrega a la lista "trabajador". Finalmente, la lista "trabajador" se ordena. """

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




