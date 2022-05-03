import glob
from unicodedata import name
from ezodf2 import newdoc
import os
import zipfile
import tempfile
import win32print
import win32api
import time
import shutil

# ***********************Paths de textos planos*************************
path = "E:\Demonio_AS400\SMM\FACTURAS-RECIBOS"

# ***********************Nombres de documentos temporales*************************
namef = "prueba_factura-JR.odt"

ifExistFAC = glob.glob(path + '\*.FAC')
# print(ifExistFAC[0])

if len(ifExistFAC) == 0:
    print("No hay archivos en el directorio")
else:
    with open(ifExistFAC[0], "r") as archivo:
        lineas = archivo.read().split('\n')
        # print(lineas)
        # for linea in lineas:
        #     variable = linea.split('=')
        #     print(variable)

n_factura       = lineas[0].split('=')[1]
observaciones   = lineas[1].split('=')[1]
observaciones1  = lineas[2].split('=')[1]
observaciones2  = lineas[3].split('=')[1]
observaciones3  = lineas[4].split('=')[1]
autorizacion    = lineas[5].split('=')[1]
valido          = lineas[6].split('=')[1]
fecha           = lineas[7].split('=')[1]
tramite         = lineas[8].split('=')[1]
tabDatos        = lineas[9].split('=')[1]
subTotal        = lineas[10].split('=')[1]
iva             = lineas[11].split('=')[1]
valorIva        = lineas[12].split('=')[1]
valorIva0       = lineas[13].split('=')[1]
valorRecaudado  = lineas[14].split('=')[1]
docID           = lineas[15].split('=')[1]
cliente         = lineas[16].split('=')[1]
direccion       = lineas[17].split('=')[1]
tel             = lineas[18].split('=')[1]


odt = newdoc(doctype='odt', filename=namef, template='E:/Demonio_AS400/factura-JR.odt')
odt.save()
a = zipfile.ZipFile("E:/Demonio_AS400/factura-JR.odt")
content = a.read("content.xml")
content = str(content.decode(encoding="utf-8"))
content = str.replace(content, "CLIENTE", cliente)
content = str.replace(content, "DOC_ID", docID)
content = str.replace(content, "DIRECCION", direccion)
content = str.replace(content, "TRAMITE", tramite)
content = str.replace(content, "TABDATOS", tabDatos)
content = str.replace(content, "OBSERVACIONES", observaciones)
content = str.replace(content, "OBSERVACIONES1", observaciones1)
content = str.replace(content, "OBSERVACIONES2", observaciones2)
content = str.replace(content, "OBSERVACIONES3", observaciones3)
content = str.replace(content, "FECHA", fecha)
content = str.replace(content, "n_factura", n_factura)
content = str.replace(content, "s_total", subTotal)
content = str.replace(content, "v_recaud", valorRecaudado)

# Moviendo .FAC a Respaldos
shutil.move(ifExistFAC[0], "E:/Demonio_AS400/SMM/Respaldos")
# Limpiando arrays
ifExistFAC.clear()
lineas.clear()

# ***********************Impresora*************************
def updateZip(zipname, filename, data):

    # Definiendo nombre de la impresora (probar cambiar a otro)
    printerName = win32print.GetDefaultPrinter()

    #Generando archivo temporal
    tmpfd, tempname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmpfd)

    #Crear compia temporal del archivo sin filename
    with zipfile.ZipFile(zipname, 'r') as zin:
        with zipfile.ZipFile(tempname, 'w') as zout:
            zout.comment = zin.comment # preserve the comment
            for item in zin.infolist():
                if item.filename != filename:
                    zout.writestr(item, zin.read(item.filename))
    
    #Reemplazando con archivo temporal
    os.remove(zipname)
    os.rename(tempname, zipname)

    #Agragando filename con nueva data
    with zipfile.ZipFile(zipname, mode='a', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(filename, data)
    
    # Impresión de documento
    printer = win32print.OpenPrinter(printerName)
    printerValues = win32print.GetPrinter(printer, 2)
    dir(printerValues['pDevMode'])
    win32api.ShellExecute(0, "print", zipname, None, ".", 1)

    # Borrar archivo temporal creado con edición de .odt
    time.sleep(15)
    os.remove(zipname)

updateZip(namef, "content.xml", content)



