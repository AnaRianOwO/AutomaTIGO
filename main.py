# El código ha sido creado por Ana María Riaño Caro 
# El código es de libre uso y modificación, siempre y cuando se reconozca la autoría
# Email: rianoana.901@gmail.com

# Librerías necesarias para funcionamiento
from time import sleep # Para hacer pausas
import tkinter as tk # Para la interfaz gráfica
from pyperclip import copy # Para copiar al portapapeles
import json # Para leer el archivo de las variables
import webbrowser # Para abrir en el navegador el link de Salesforce
import os # Para abrir la carpeta de descargas, eliminar archivos y abrir archivos
from pyautogui import click # Para hacer click en la pantalla y descargar los archivos
from xlwings import Book # Para abrir el archivo de Excel y copiar los datos
import pickle # Para guardar los datos en un archivo pickle
import pandas as pd  # Para crear los dataframes y leer pickles
from datetime import date, timedelta, datetime # Para obtener la fecha de hoy y la de hace 7 días
import matplotlib.pyplot as plt # Para crear las gráficas
import subprocess # Para el tutorial y abrir la carpeta de informes
from tkinter import ttk, filedialog # Para la interfaz gráfica y pedir la carpeta de descarga
import ttkbootstrap as ttb # Para los diseños de la interfaz gráfica
from ttkbootstrap.scrolled import ScrolledFrame # Para crear un frame con scroll
import math # Para redondear los números
from pynput import mouse # Para obtener las coordenadas de la pantalla
import platform # Para saber el sistema operativo


# Variables
if not os.path.exists('data/datos.json'):
    print("No se encontró el archivo de variables")

with open('data/datos.json', 'r', encoding='utf-8') as f:
    jsonData = json.load(f)

if not os.path.exists('data/coordenadas.json'):
    print("No se encontró el archivo de coordenadas")

with open('data/coordenadas.json', 'r') as f:
    jsonCoordenadas = json.load(f)

# Variables de la interfaz, no se pueden cambiar, solamente cambiar lo de data.json
coordenada = jsonCoordenadas['coordenada']
carpeta = jsonData['carpeta']
archivos = jsonData['archivos']
macro = jsonData['macro']
columna = jsonData['columna']
link = jsonData['link']
hoja = jsonData['hoja']
tabla = jsonData['tabla']
saludo = jsonData['saludo']
tiempoEspera = jsonData['tiempoEspera']
so = jsonData['SO']
sistemaOperativo = platform.system()

# Funciones primarias

def abrirLink(url):
    webbrowser.open(url)
    mostrarMensaje("Se ha abierto el link")

def descargarArchivo(option):
    click(coordenada['actualizar'][0], coordenada['actualizar'][1], 1, 1)
    sleep(tiempoEspera["actualizarInforme"])
    click(coordenada['opciones'][0], coordenada['opciones'][1], 1, 1)
    # La opción 1 es para OPP general, la opción 0 es para Account Plan
    # Esta variable cambia ya que en mi cuenta se puede hacer otra opción en guardar como
    if option == 1:
        click(coordenada['exportar'][0], coordenada['exportar'][1], 1, 1)
    elif option == 0:
        click(coordenada['exportar2'][0], coordenada['exportar2'][1], 1, 1)
    click(coordenada['detalles'][0], coordenada['detalles'][1], 1, 1)
    click(coordenada['desplegable'][0], coordenada['desplegable'][1], 1, 1)
    click(coordenada['xlsx'][0], coordenada['xlsx'][1], 1, 1)
    click(coordenada['descargar'][0], coordenada['descargar'][1], 1, 1)
    mostrarMensaje("Se ha descargado el archivo")

def obtenerUltimoArchivo(carpeta):
    if not os.path.exists(carpeta):
        mensaje = f"La carpeta {carpeta} no existe"
        mostrarMensaje(mensaje)
        return None
    archivos = os.listdir(carpeta)
    archivosRuta = [os.path.join(carpeta, archivo) for archivo in archivos]
    ultimoArchivo = max(archivosRuta, key=os.path.getmtime)
    return ultimoArchivo

def moverArchivo(archivoAntiguo, carpetaDescargas):
    if os.path.exists(archivoAntiguo):
        os.remove(archivoAntiguo)
        mensaje = f"Se ha eliminado {archivoAntiguo} exitosamente"
        mostrarMensaje(mensaje)
    else:
        mensaje = f"El archivo {archivoAntiguo} no existe"
        mostrarMensaje(mensaje)

    ultimoArchivo = obtenerUltimoArchivo(carpetaDescargas)

    os.rename(ultimoArchivo, archivoAntiguo)
    mostrarMensaje("Se ha movido el archivo "+ultimoArchivo+" exitosamente")

def ejecutarMacro(rutaInforme, macro):
    if os.path.exists(rutaInforme):
        libro = Book(rutaInforme)
        sleep(tiempoEspera["esperarCargaExcel"])
        libro.macro(macro).run()
        libro.save()
    else:
        mostrarMensaje("No se encontró el archivo "+rutaInforme)

def ejecutarMacroConVariable(rutaInforme, macro, nombreHoja, rutaRAW, nombreTabla):
    libro = Book(rutaInforme)
    if os.path.exists(rutaRAW):
        rutaABS = os.path.abspath(rutaRAW)
    else:
        mostrarMensaje("No se encontró el archivo "+rutaRAW)
        return
    libro.macro(macro).run(nombreHoja, rutaABS, nombreTabla)
    libro.save()

def manipularExcel(rutaInforme, nombreHoja, nombreTablaDinamica):
    libro = Book(rutaInforme)

    if not libro.sheets[nombreHoja]:
        mostrarMensaje("No se encontró la hoja "+nombreHoja+" en el archivo "+rutaInforme)
        return
    
    hoja = libro.sheets[nombreHoja]

    if not hoja.api.PivotTables(nombreTablaDinamica):
        mostrarMensaje("No se encontró la tabla dinámica "+nombreTablaDinamica+" en la hoja "+nombreHoja)
        return
    
    tablaDinamica = hoja.api.PivotTables(nombreTablaDinamica)

    rangoTabla = tablaDinamica.TableRange1.Address
    dataValues = hoja.range(rangoTabla).value

    if dataValues == None:
        mostrarMensaje("No se encontró información en la tabla dinámica "+nombreTablaDinamica+" en la hoja "+nombreHoja)
        return
    data = pd.DataFrame(dataValues[1:], columns=dataValues[0])
    mostrarMensaje("Se ha obtenido la información del archivo "+rutaInforme+" y de la hoja "+nombreHoja+" exitosamente")
    return data

def FastTrack(data):
    FastTrack = {}

    # La información de las oportunidades se guarda en un diccionario
    for index, row in data.iterrows():
        ejecutivo = row[columna['ejecutivo']]
        opp = row[columna['oportunidad']]
        fecha = row[columna['fecha']]
        prob = row[columna['probabilidad']]
        
        if ejecutivo in FastTrack:
            FastTrack[ejecutivo].append({'opp':opp, 'fecha':fecha, 'probabilidad':prob})
            pass
        else:
            FastTrack[ejecutivo] = [{'opp':opp, 'fecha':fecha, 'probabilidad':prob}]
    
    # Se crea un string con la información de las oportunidades, este string se copia al portapapeles por cada ejecutivo que aplique

    ejecutivos = FastTrack.keys()
    pantalla = ""
    for ejecutivo in ejecutivos:
        mensaje = ejecutivo+saludo["FastTrack"]
        pantalla += ejecutivo + "\n\n"
        for opps in FastTrack[ejecutivo]:
            mensaje+=opps['opp']+': vence el '+opps['fecha']+' con probabilidad de '+opps['probabilidad']+'\n'
            pantalla += opps['opp']+' -- '+opps['fecha']+' -- '+opps['probabilidad']+'\n'
        pantalla += "\n\n"
        copy(mensaje)
        sleep(tiempoEspera["copiarPortapapeles"])
    
    return pantalla

def proximoACerraroDRB(data, opc):
    proxi = {}

    for index, row in data.iterrows():
        ejecutivo = row[columna['ejecutivo']]
        opp = row[columna['oportunidad']]
        fecha = row[columna['fecha']]
        if opc == 1:
            fecha+= ' que está en la etapa de '+row[columna['etapa']]
        
        if ejecutivo in proxi:
            proxi[ejecutivo].append({'opp':opp, 'fecha':fecha})
            pass
        else:
            proxi[ejecutivo] = [{'opp':opp, 'fecha':fecha}]
    
    ejecutivos = proxi.keys()
    pantalla = ""
    for ejecutivo in ejecutivos:
        if opc == 1:
            mensaje = ejecutivo+saludo["ProximoVencer"]
            pantalla += ejecutivo + "\n\n"
        else:
            mensaje = ejecutivo+saludo["drb"]
            pantalla += ejecutivo + "\n\n"
        for opps in proxi[ejecutivo]:
            mensaje+=opps['opp']+': vence el '+str(opps['fecha'])+'\n'
            pantalla += opps['opp']+' -- '+str(opps['fecha'])+'\n'
        pantalla += "\n\n"
        copy(mensaje)
        sleep(tiempoEspera["copiarPortapapeles"])
    
    return pantalla

def productosRaros(data, negativo, contexto):
    raro = {}
    for index, row in data.iterrows():
        ejecutivo = row[columna['ejecutivo']]
        opp = row[columna['oportunidad']]
        producto = row[columna['producto']]

        if producto == None:
            producto = 'sin producto'

        if negativo == True:
            ventaNeta = row[columna['ventaNeta']]
            producto+= ' con venta neta negativa de '+str(ventaNeta)

        if ejecutivo in raro:
            if opp in raro[ejecutivo]:
                raro[ejecutivo][opp].append(producto)
            else:  
                raro[ejecutivo][opp] = [producto]
        else:
            raro[ejecutivo] = {opp:[producto]}

    ejecutivos = raro.keys()
    pantalla = ""
    for ejecutivo in ejecutivos:
        mensaje = ejecutivo+saludo["oportunidadesRaras"][contexto]
        pantalla += ejecutivo + "\n\n"
        for opps in raro[ejecutivo]:
            opp = raro[ejecutivo][opps]
            mensaje += opps + ': tiene productos de ' + ', '.join(opp) + ' \n'
            pantalla += opps + ': tiene productos de ' + ', '.join(opp) + ' \n'
        pantalla += "\n\n"
        copy(mensaje)
        sleep(tiempoEspera["copiarPortapapeles"])

    return pantalla

def clientesCompletitud(data):
    clientes = {}

    for index, row in data.iterrows():
        ejecutivo = row[columna['ejecutivo']]
        cliente = row[columna['cliente']]
        completitud = row[columna['porcentajeCompletitud']]

        if math.isnan(completitud):
            completitud = 0
        
        if ejecutivo in clientes:
            clientes[ejecutivo].append({'cliente':cliente, 'completitud':completitud})
            pass
        else:
            clientes[ejecutivo] = [{'cliente':cliente, 'completitud':completitud}]
    
    ejecutivos = clientes.keys()
    pantalla = ""
    for ejecutivo in ejecutivos:
        mensaje = ejecutivo+saludo["clientesCompletitud"]
        pantalla += ejecutivo + "\n\n"
        for opps in clientes[ejecutivo]:
            mensaje+=opps['cliente']+': '+str(opps['completitud'])+'\n'
            pantalla += opps['cliente']+' -- '+str(opps['completitud'])+'\n'
        pantalla += "\n\n"
        copy(mensaje)
        sleep(tiempoEspera["copiarPortapapeles"])
    
    return pantalla

def guardarDataframe(inf, rutaInforme, nombreHoja, nombreTablaDinamica):
    datos = manipularExcel(rutaInforme, nombreHoja, nombreTablaDinamica)

    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return

    fecha = date.today()
    directorio = carpeta["dataframes"]+inf+"/"
    nombreArchivo = f"{directorio}{fecha}.pkl"
    
    # Verificar si el directorio existe, si no, crearlo
    if not os.path.exists(directorio):
        os.makedirs(directorio)

    # Verificar si el archivo ya existe
    if not os.path.exists(nombreArchivo):
        with open(nombreArchivo, "wb") as f:
            pickle.dump(datos, f)
            mensaje = "Archivo creado exitosamente"
            mostrarMensaje(mensaje)
    else:
        mensaje = "El archivo ya existe"
        mostrarMensaje(mensaje)
    
    mensaje = "El archivo se ha guardado exitosamente"
    return mensaje

def traerArchivosParaComparar(inf):
    fechaHoy = date.today()
    file = carpeta["dataframes"]

    archivoHoy = f"{file}{inf}/{fechaHoy}.pkl"
    archivoSP = filedialog.askopenfilename(initialdir=file+inf, title="Selecciona el archivo que desees comparar", filetypes=(("Pickle files", "*.pkl"), ("all files", "*.*")))

    if not os.path.exists(archivoHoy):
        mensaje = "No existe el archivo de hoy"
        mostrarMensaje(mensaje)

        return None, None
    elif not os.path.exists(archivoSP):
        mensaje = "No se seleccionó ningún archivo para comparar"
        mostrarMensaje(mensaje)

        return None, None
    else:
        dataframeHoy = pd.read_pickle(archivoHoy)
        dataframeSP = pd.read_pickle(archivoSP)
        mostrarMensaje("Se han cargado los archivos exitosamente")

        return dataframeHoy, dataframeSP

def grafico(dataframe):
    if dataframe.empty:
        mostrarMensaje("No se encontró el archivo para graficar.")
        return
    dataframe[columna["precioSinIva"]] = dataframe[columna["precioSinIva"]].astype(float)

    dataframeSIMPL = dataframe[[columna["tipoProducto"], columna["precioSinIva"]]]

    df = dataframeSIMPL.groupby(columna["tipoProducto"]).sum()
    df_values = df[columna["precioSinIva"]].astype(float)
    grafico = df.plot(kind='barh', figsize=(10, 10), color='#86bf91', zorder=2, width=0.85)

    for i, v in enumerate(df_values.values):
        formatted_value = '{:,.2f}'.format(v) 
        grafico.text(v, i, formatted_value, ha='left', va='center')
        
    grafico.set_xlabel(columna["precioSinIva"])

    eje_x = grafico.get_xaxis()
    eje_x.set_visible(False)
    
    return grafico

def graficarCOMP(hoy, SP):
    # Aquí se crean ambos gráficos y se les asigna un título
    hoy = grafico(hoy).set_title("Hoy")
    semanaPasada = grafico(SP).set_title("Antes")
    mostrarMensaje("Se han creado los gráficos exitosamente")

    return hoy, semanaPasada

def abrirArchivo(nombreArchivo):
    
    try:
        if nombreArchivo[-4:] == "json":
            subprocess.run(['start', 'notepad', nombreArchivo], shell=True)
        else:
            subprocess.run(['start', nombreArchivo], shell=True)
        mostrarMensaje(f"Se ha abierto el archivo {nombreArchivo} exitosamente")
    except FileNotFoundError:
        mostrarMensaje(f"No se pudo abrir el archivo {nombreArchivo}. Verifica el nombre y la ruta.")

def abrirCarpeta(rutaCarpeta):
    ruta = os.path.realpath(rutaCarpeta)
    try:
        subprocess.run([so[sistemaOperativo]["carpeta"], ruta], shell=True)
        mostrarMensaje("Se ha abierto la carpeta de informes exitosamente")
    except FileNotFoundError:
        mostrarMensaje(f"No se pudo abrir la carpeta {ruta}. Verifica la ruta.")

def definirCoordenadas():
    XandY = []
    def on_click(x, y, button, pressed):
        if pressed:
            # print(f"Se hizo un clic en las coordenadas ({x}, {y}) con el botón {button}")
            XandY.append([x, y])

    abrirLink(link["OPP"])

    listener = mouse.Listener(on_click=on_click)
    listener.start()
    sleep(tiempoEspera["actualizarCoordenadas"])
    listener.stop()

    coord = {
        "coordenada": {
            "actualizar": XandY[0],
            "opciones": XandY[1],
            "exportar": XandY[2],
            "exportar2": [XandY[2][0], XandY[2][1]-45],
            "detalles": XandY[3],
            "desplegable": XandY[4],
            "xlsx": XandY[5],
            "descargar": XandY[6]
        }
    }

    with open('data/coordenadas.json', 'w') as f:
        json.dump(coord, f, indent=4)

    abrirArchivo(archivos["coordenadas"])
    mostrarMensaje("Se han guardado las coordenadas exitosamente")
# Opciones

def actualizarOPPGeneral():
    abrirLink(link["OPP"])
    sleep(tiempoEspera["esperarCargaPagina"])
    descargarArchivo(1)
    sleep(tiempoEspera["esperarDescarga"])
    moverArchivo(archivos["OA"], carpeta["descargas"])
    ejecutarMacroConVariable(archivos["Informe OPP general"], macro["macroOPP"], hoja["DROPP"], archivos["OA"], tabla["DROPP"])
    mostrarMensaje("Informe de opp actualizado")

def candidatosFastTrack():
    datos = manipularExcel(archivos["Informe OPP general"], hoja["FastTrack"], tabla["FastTrack"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    diccionario = FastTrack(datos)
    mensaje = "Revisa el portapapeles (Windows + v), ya se copiaron los datos de Fast Track"
    return diccionario, mensaje

def proximosAVencer():
    datos = manipularExcel(archivos["Informe OPP general"], hoja["ProximoVencer"], tabla["EsteMes"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    diccionario = proximoACerraroDRB(datos, 1)
    mensaje = "Revisa el portapapeles (Windows + v), ya se copiaron los datos de las oportunidades próximas a vencer"
    return diccionario, mensaje

def productosNegativos():
    datos = manipularExcel(archivos["Informe OPP general"], hoja["VentaNeta"], tabla["VentaNeta"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    diccionario = productosRaros(datos, True, "negativos")
    mensaje = "Revisa el portapapeles (Windows + v), ya se copiaron los datos de las oportunidades con productos negativos"
    return diccionario, mensaje

def productosCero():
    datos = manipularExcel(archivos["Informe OPP general"], hoja["Productos0"], tabla["Productos0"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    diccionario = productosRaros(datos, False, "en cero")
    mensaje= "Revisa el portapapeles (Windows + v), ya se copiaron los datos de las oportunidades con productos en 0"
    return diccionario, mensaje

def oportunidadesCero():
    datos = manipularExcel("data/OPP/Informe OPP general.xlsm", hoja["Productos0"], tabla["Oportunidad0"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    datos = datos[datos['Suma de Precio total'] == 0]
    diccionario = ""
    for index, dato in datos.iterrows():
        diccionario += dato['Oportunidad']+'\n'
    copy(diccionario)
    mensaje = "Datos de las oportunidades con valores 0 extraídos"
    return diccionario, mensaje

def actualizarBacklog():
    abrirLink(link["BL"])
    sleep(tiempoEspera["esperarCargaPagina"])
    descargarArchivo(1)
    sleep(tiempoEspera["esperarDescargaLarga"])
    moverArchivo(archivos["NPI"], carpeta["descargas"])
    ejecutarMacroConVariable(archivos["Informe Backlog"], macro["macroBL"], hoja["DRBacklog"], archivos["NPI"], tabla["DRBacklog"])            
    mostrarMensaje("El informe de backlog se ha actualizado")

def AccountPLan():
    abrirLink(link["AC"])
    sleep(tiempoEspera["esperarCargaPagina"])
    descargarArchivo(1)
    sleep(tiempoEspera["esperarDescarga"])
    moverArchivo(archivos["AC"], carpeta["descargas"])
    ejecutarMacroConVariable(archivos["AccountPlan"], macro["AC"], hoja["DRAC"], archivos["AC"], tabla["DRAC"])
    datos = manipularExcel(archivos["AccountPlan"], hoja["AccountPlan"], tabla["AccountPlan"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    diccionario = clientesCompletitud(datos)
    mensaje = f"Revisa el portapapeles (Windows + v), ya se copiaron los datos de los clientes con menos de 70% de completitud"
    return diccionario, mensaje

def SoW():
    ejecutarMacro(archivos["SoW"], macro["SoW"])
    datos = manipularExcel(archivos["SoW"], hoja["SoW"], tabla["SoW"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    diccionario = clientesCompletitud(datos)

    mensaje = f"Revisa el portapapeles (Windows + v), ya se copiaron los datos de los clientes con menos de 70% de completitud"
    return diccionario, mensaje

def comparacionOPP():
    guardarDataframe("OPP", archivos["Informe OPP general"], hoja["General"], tabla["General"])
    hoy, sp = traerArchivosParaComparar("OPP")
    if hoy.empty or sp.empty:
        mostrarMensaje("No se encontraron archivos para comparar. Revisa el mensaje anterior")
        return
    grafico1, grafico2 = graficarCOMP(hoy, sp)
    plt.show()
    
    dfh = hoy.merge(sp, how='outer', indicator='union')
    dfh = dfh[dfh.union=='left_only'].sort_values(by=[columna["precioSinIva"]])
    dfh = dfh[[columna['tipoProducto'], columna['ejecutivo'], columna['cliente'], columna['precioSinIva']]]

    dfsp = sp.merge(hoy, how='outer', indicator='union')
    dfsp = dfsp[dfsp.union=='left_only'].sort_values(by=[columna["precioSinIva"]])
    dfsp = dfsp[[columna['tipoProducto'], columna['ejecutivo'], columna['cliente'], columna['precioSinIva']]]

    mostrarMensaje("Se ha creado la comparación de OPP exitosamente")

    return dfh, dfsp

def DRB():
    ejecutarMacro(archivos["DRB"], macro["DRB"])
    datos = manipularExcel(archivos["DRB"], hoja["DRB"], tabla["DRB"])
    if datos.empty:
        mostrarMensaje("No se ha extraído la información del archivo. Revisa el mensaje anterior")
        return
    diccionario = proximoACerraroDRB(datos, 0)
    mensaje = "Revisa el portapapeles, ya se copiaron los datos de las oportunidades que aplican a DRB"
    return diccionario, mensaje

def comparacionBL():
    guardarDataframe("BL", archivos["Informe Backlog"], hoja["BL"], tabla["BL"])
    hoy, sp = traerArchivosParaComparar("BL")

    if hoy.empty or sp.empty:
        mensaje = "No se encontraron archivos para comparar"
        mostrarMensaje(mensaje)
        return
    
    dfh = hoy.merge(sp, how='outer', indicator='union')
    dfh = dfh[dfh.union=='left_only'].sort_values(by=[columna["fecha"]])
    dfh = dfh[[columna['año'], columna['fecha'], columna['cliente'], columna['oportunidad']]]

    dfsp = sp.merge(hoy, how='outer', indicator='union')
    dfsp = dfsp[dfsp.union=='left_only'].sort_values(by=[columna["fecha"]])
    dfsp = dfsp[[columna['año'], columna['fecha'], columna['cliente'], columna['oportunidad']]]

    mostrarMensaje("Se ha creado la comparación de Backlog exitosamente")

    return dfh, dfsp

def comparacionClientes(tipo, archivo, hojaCli, tablaCli):
    guardarDataframe(tipo, archivo, hojaCli, tablaCli)
    hoy, sp = traerArchivosParaComparar(tipo)

    if hoy.empty or sp.empty:
        mensaje = "No se encontraron archivos para comparar"
        mostrarMensaje(mensaje)
        return
    
    dfh = hoy.merge(sp, how='outer', indicator='union')
    dfh = dfh[dfh.union=='left_only'].sort_values(by=[columna["porcentajeCompletitud"]])
    dfh = dfh[[columna["ejecutivo"], columna['cliente'], columna['porcentajeCompletitud']]]

    dfsp = sp.merge(hoy, how='outer', indicator='union')
    dfsp = dfsp[dfsp.union=='left_only'].sort_values(by=[columna["porcentajeCompletitud"]])
    dfsp = dfsp[[columna["ejecutivo"], columna['cliente'], columna['porcentajeCompletitud']]]

    mostrarMensaje("Se ha creado la comparación de clientes exitosamente")

    return dfh, dfsp

def tutorial():
    nombreArchivo = archivos["tuto"]
    abrirArchivo(nombreArchivo)

def eliminarDataframesAntiguos():
    carpetaOPP = carpeta["dataframes"]+ "OPP/"
    carpetaBL = carpeta["dataframes"]+ "BL/"
    carpetaAC = carpeta["dataframes"]+ "AC/"
    carpetaSOW = carpeta["dataframes"]+ "SOW/"

    fechaHoy = date.today()
    SemanaPasada = fechaHoy - timedelta(month=1)

    def eliminarDataframes(carpeta):
        for archivo in os.listdir(carpeta):
            fechaArchivo = date.fromisoformat(archivo[:-4])
            if fechaArchivo < SemanaPasada:
                os.remove(carpeta+archivo)
                mensaje = "Se ha eliminado "+ carpeta + archivo +" exitosamente\n"
                mostrarMensaje(mensaje)       
    
    eliminarDataframes(carpetaOPP)
    eliminarDataframes(carpetaBL)
    eliminarDataframes(carpetaAC)
    eliminarDataframes(carpetaSOW)

def asignarCarpetaDescargas():
    carpeta["descargas"] = filedialog.askdirectory()
    with open('data/datos.json', 'w') as f:
        json.dump(jsonData, f, indent=4)
    mostrarMensaje("Se ha guardado la ruta de la carpeta de descargas exitosamente")

def antesDeEmpezar():
    tuto = ttb.Window(themename='solar')
    tuto.title("Realiza todas estas funciones antes de empezar a usar el aplicativo")

    sesion = ttk.Button(tuto, text="1. Sesión de Salesforce", command=lambda: abrirLink(link["OPP"]))
    coordenadas = ttk.Button(tuto, text="2. Coordenadas", command=definirCoordenadas)
    descargas = ttk.Button(tuto, text="3. Carpeta de descargas", command=lambda: asignarCarpetaDescargas())
    variables = ttk.Button(tuto, text="4. Variables", command=lambda: abrirArchivo(archivos["variables"]))
    excel = ttk.Button(tuto, text="5. Configurar Excel", command=lambda: abrirArchivo(archivos["DRB"]))
    data = ttk.Button(tuto, text="6. Configurar carpeta de data", command=lambda: abrirCarpeta(carpeta["data"]))

    sesion.pack(pady=10, padx=10)
    coordenadas.pack(pady=10, padx=10)
    descargas.pack(pady=10, padx=10)
    variables.pack(pady=10, padx=10)
    excel.pack(pady=10, padx=10)
    data.pack(pady=10, padx=10)

# Menu

root = ttb.Window(themename='solar')
root.title("automaTIGO")

def mostrarDatos(mensaje, log):
    datosLabel.config(text=mensaje, justify="left")
    mostrarMensaje(log)

def mostrarComparacion(dataframe1, dataframe2, opcion):
    n.add(hoy, text='Hoy')
    hoyText = ""
    for index, row in dataframe1.iterrows():
    # La comparación se hace tanto para OPP como para Backlog, por eso se pregunta por la opción
        if opcion == 1:
            hoyText += f"{row['Tipo Producto']} -- {row['Ejecutivo']} -- {row['Cliente']} -- {row['Suma de Precio sin IVA']}\n"
        elif opcion == 2:
            hoyText += f"{row['Año']} -- {row['Fecha de cierre']} -- {row['Cliente']} -- {row['Oportunidad']}\n"
        elif opcion == 3:
            if row['Promedio de Porcentaje Completitud'] < 1:
                row['Promedio de Porcentaje Completitud'] = row['Promedio de Porcentaje Completitud'] *100
            hoyText += f"{row['Ejecutivo']} -- {row['Cliente']} -- {row['Promedio de Porcentaje Completitud']}\n"
    hoyLabel.config(text="Datos nuevos\n\n" + hoyText, justify="left")
    
    n.add(sp, text='Histórico')
    spText = ""
    for index, row in dataframe2.iterrows():
        if opcion == 1:
            spText += f"{row['Tipo Producto']} -- {row['Ejecutivo']} -- {row['Cliente']} -- {row['Suma de Precio sin IVA']}\n"
        elif opcion == 2:
            spText += f"{row['Año']} -- {row['Fecha de cierre']} -- {row['Cliente']} -- {row['Oportunidad']}\n"
        elif opcion == 3:
            spText += f"{row['Ejecutivo']} -- {row['Cliente']} -- {row['Promedio de Porcentaje Completitud']*100}\n"
    spLabel.config(text="Datos previos\n\n" + spText, justify="left")

def mostrarMensaje(mensaje):
    log = "["+str(datetime.now()) + "]:" + mensaje + "."
    label = tk.Label(logging, text=log, font=("Arial", 10))
    label.pack()

def salir():
    root.quit()

menu = tk.Menu(root)
root.config(menu=menu)
root.geometry("1000x700")
# Crear el menú "Archivo"
oppMenu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="OPP", menu=oppMenu)
oppMenu.add_command(label="Actualizar Opp General", command=actualizarOPPGeneral)
oppMenu.add_command(label="Candidatos Fast Track", command=lambda: mostrarDatos(*candidatosFastTrack()))
oppMenu.add_command(label="Oportunidades próximas a vencer", command=lambda: mostrarDatos(*proximosAVencer()))
oppMenu.add_command(label="Productos en $0", command=lambda: mostrarDatos(*productosCero()))
oppMenu.add_command(label="Oportunidades en $0", command=lambda: mostrarDatos(*oportunidadesCero()))
oppMenu.add_command(label="Productos en -", command=lambda: mostrarDatos(*productosNegativos()))
oppMenu.add_command(label="Mostrar comparación OPP general", command=lambda: mostrarComparacion(*comparacionOPP(), 1))
oppMenu.add_command(label="DRB", command=lambda: mostrarDatos(*DRB()))

backlogMenu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Backlog", menu=backlogMenu)
backlogMenu.add_command(label="Actualizar Backlog", command=actualizarBacklog)
backlogMenu.add_command(label="Mostrar comparación Backlog", command=lambda: mostrarComparacion(*comparacionBL(), 2))

clientesMenu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Clientes", menu=clientesMenu)
clientesMenu.add_command(label="Actualizar Account Plan", command=lambda: mostrarDatos(*AccountPLan()))
clientesMenu.add_command(label="Mostrar comparación Account Plan", command=lambda: mostrarComparacion(*comparacionClientes("AC", archivos["AccountPlan"], hoja["AccountPlan"], tabla["AccountPlan"]), 3))
clientesMenu.add_command(label="Actualizar SoW", command=lambda: mostrarDatos(*SoW()))
clientesMenu.add_command(label="Mostrar comparación SoW", command=lambda: mostrarComparacion(*comparacionClientes("SOW", archivos["SoW"], hoja["SoW"], tabla["SoW"]), 3))

opcionesMenu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Opciones", menu=opcionesMenu)
opcionesMenu.add_command(label="Manual de usuario", command=tutorial)
opcionesMenu.add_command(label="Antes de empezar", command=antesDeEmpezar)
opcionesMenu.add_command(label="Eliminar dataframes antiguos", command=eliminarDataframesAntiguos)
opcionesMenu.add_command(label="Abrir carpeta de informes", command=lambda: abrirCarpeta(carpeta["data"]))
opcionesMenu.add_command(label="Cambiar variables del programa", command=lambda: abrirArchivo(archivos["variables"]))
opcionesMenu.add_command(label="Repositorio", command=lambda: abrirLink(link["repo"]))
opcionesMenu.add_command(label="¿Cómo agregar más opciones?", command=lambda: abrirArchivo(archivos["automatizacion"]))
opcionesMenu.add_separator()
opcionesMenu.add_command(label="Salir", command=salir)

# Crear la ventana principal
ventanaPrincipal = ttb.PanedWindow(root, orient="vertical", height=700, width=1200)
comparacion = ttb.LabelFrame(ventanaPrincipal, text="Comparación de datos", height=300, width=1200, padding=10)
datos = ttb.LabelFrame(ventanaPrincipal, text="Datos solicitados", height=300, width=1000)
registros = ttb.LabelFrame(ventanaPrincipal, text="Log", height=200, width=1000)
ventanaPrincipal.add(comparacion)
ventanaPrincipal.add(datos)
ventanaPrincipal.add(registros)
ventanaPrincipal.pack()

# Crear la ventana de comparación
n = ttk.Notebook(comparacion)
hoy = ttk.Frame(n)
sp = ttk.Frame(n)
spScroll = ScrolledFrame(sp, autohide=True, height=500, width=1200)
spLabel = tk.Label(spScroll, font=("Arial", 10))
hoyScroll = ScrolledFrame(hoy, autohide=True, height=500, width=1200)
hoyLabel = tk.Label(hoyScroll, font=("Arial", 10))
spScroll.pack()
hoyScroll.pack()
hoyLabel.pack()
spLabel.pack()
n.pack()

# Crear la ventana de datos
datosScroll = ScrolledFrame(datos, autohide=True, height=500, width=1200)
datosLabel = tk.Label(datosScroll, text="", wraplength=1000, font=("Arial", 10), justify="left")
datosScroll.pack()
datosLabel.pack()

# Crear la ventana de registros
logging = ScrolledFrame(registros, autohide=True, height=500, width=1200)
logging.pack()

root.mainloop()