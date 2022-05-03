from unicodedata import name
from ezodf2 import newdoc
import os
import zipfile
import tempfile
import win32print
import win32api
import time

namef = "prueba_factura-JR.odt"

odt = newdoc(doctype='odt', filename=namef, template='E:/Demonio_AS400/SMM/Modelos/factura-JR.odt')
odt.save()
a = zipfile.ZipFile("E:/Demonio_AS400/SMM/Modelos/factura-JR.odt")
content = a.read("content.xml")
content = str(content.decode(encoding="utf-8"))
content = str.replace(content, "CLIENTE", "Aquiles Baeza")
content = str.replace(content, "DOC_ID", "00325320325325")

def updateZip(zipname, filename, data):

    # Definiendo nombre de la impresora
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


    