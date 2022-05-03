import win32print
import win32api

printerName = win32print.GetDefaultPrinter()
printer = win32print.OpenPrinter(printerName)
printerValues = win32print.GetPrinter(printer, 2)
dir(printerValues['pDevMode'])
win32api.ShellExecute(0, "print", "E:/Demonio_AS400/demonioAS400/prueba_factura-JR.odt", None, ".", 1)
