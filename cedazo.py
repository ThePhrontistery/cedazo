#!/bin/env python
# Cedazo - Morph, Meld and Merge data - Copyright © 2023 Iwan van der Kleijn - See LICENSE.txt for conditions
#from openpyxl import Workbook
import openpyxl

from data import change_colour, remove, removeFormatting, unmerge_cells
#Damos la localización del fichero de entrada
ruta_input = "C:\\Users\\debarrei\\Documents\\FORMACION-CREW VERVE\\EJERCICIOS\\CedazoExcel\\in.xlsx"
#Damos la localización del fichero de salida
ruta_output = "C:\\Users\\debarrei\\Documents\\FORMACION-CREW VERVE\\EJERCICIOS\\CedazoExcel\\out.xlsx"
# Creamos objeto wb (libro) de tipo workbook y lo cargamos con lo del excel
wb = openpyxl.load_workbook(ruta_input)
# Creamos objeto ws (hoja), siendo la hoja activa
#ws = wb.active 
ws = wb['Retain Report']
#print('la celda A1 es:' ,ws['A1'].value)

#Metodo que desmergea las celdas de las filas a eliminar
unmerge_cells(ws)

# Metodo que sirve para borrar todas las filas de 1 a 11
#for row in ws: 
remove(ws)

# Metodo que elimina el color amarillo de todas la celdas que sean amarillas 
change_colour(ws)

wb.save(ruta_output)


