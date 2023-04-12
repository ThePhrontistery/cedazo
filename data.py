from copy import copy
from types import CellType
import openpyxl
from openpyxl.styles import PatternFill 


def remove_empty(ws): 
    '''Metodo que sirve para borrar todas las filas que estan vacias (no se utiliza de momento porque no sirve para borrar
       lad filas que hay que eliminar pero tienen contenido)'''
    filas = ws.max_row
    for i in range(filas, 0, -1):
        celdas_vacias = all([cell.value is None for cell in ws[i]])
        if celdas_vacias:
            ws.delete_rows(i, 1)

def unmerge_cells(ws):
    '''Metodo que desmergea las celdas de las filas a eliminar'''
    # Buscamos las celdas que estan mergeadas
    for merge in list(ws.merged_cells):
        # Separamos esas celdas mergeadas
        ws.unmerge_cells(range_string=str(merge))

def remove(ws,ini,end): 
    '''Metodo que sirve para borrar todas las filas de 1 a 11'''
    ws.delete_rows(ini,end)


def change_colour(ws): 
    '''Metodo que a todas la celdas les pone color blanco (quita el amarillo) '''
    max_row = ws.max_row
    #print('El numero de filas rellenas es: ', max_row)
    max_column = ws.max_column
    #print('El numero de columnas rellenas es: ', max_column)

    # Hacemos doble bucle para recoger todos los datos de la tabla que empezaba en fila 12 (ahora empezara en la tabla 
    # destino en la fila 1).
    for i_row in range(1, max_row + 1):
       for i_column in range(1, max_column + 1):
           cell = ws.cell(row = i_row, column = i_column)
           #If ((cell.fill.fgcolor.type == 'indexed' and cell.fill.fgcolor.indexed == 43) or
           #(cell.fill.fgcolor.type == 'rgb' and cell.fill.fgcolor.rgb == 'FFFFFF99')):
           # Cambia para todas las celdas el fondo a blanco
           cell.fill = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type = "solid")


def change_colour_mio(ws): 
#Elimina el color amarillo de todas la celdas.
    filas = ws.max_row
    #style_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type = "solid")
    #style_yellow = "00FFFF00"
    for i_row in range(filas, 0, -1):
       for cell in ws[i_row]:
           #if cell.fill == style_yellow:
        #cells_yellows = all([cell.fill is style_yellow for cell in ws[i]])
        
        #if cells_yellows  == style_yellow:
            cell.fill = PatternFill(start_color="FFFFFF", end_color=None, fill_type = None)
           
def remove_formatting(ws):
    # ws is not the worksheet name, but the worksheet object
    filas = ws.max_row
    for i_row in range(filas, 0, -1):
        for cell in ws[i_row]:
            cell.style = 'Normal'
def copyStyle(newCell, cell): 
    '''Metodo que copia el formato de una celda a una nueva '''
    if cell.has_style: 
        newCell.style = copy(cell.style) 
        newCell.font = copy(cell.font) 
        newCell.border = copy(cell.border) 
        newCell.fill = copy(cell.fill) 
        newCell.number_format = copy(cell.number_format) 
        newCell.protection = copy(cell.protection) 
        newCell.alignment = copy(cell.alignment)
  
def insert_column(ws, colNr ,headerRow , headerVal):
    '''Metodo que inserta una columna nueva detrás de la columna H poniendo la cabecera'''
    ws.insert_cols(colNr)
    #añadimos el título a la columna 
    ws.cell(row=headerRow, column=colNr).value = headerVal

def format_column(ws, colNr):
    '''Metodo que da formato a la columna que se ha creado'''
    max_row = ws.max_row
    #recorremos todas las celdas de la columna H para copiar su formato a la nueva columna mediante el método
    #copyStyle definido más arriba
    for i_row in range(1, max_row + 1):
        cell_new = ws.cell(row=i_row, column= colNr)
        cell_origin = ws.cell(row=i_row, column=colNr-1)
        copyStyle(cell_new, cell_origin)

    #max_col = ws.max_column
    #for col in ws.iter_cols(min_col = 8, max_col = 8):
        #cell_new = col[8]
        #cell_origin = col[7]
        #cell_condition = ws.cell(row=i_row, column= colNr-3)
        

def format_condition(ws, colNr, condition, color):     
    '''Metodo que da formato a la columna que se ha creado según criterio'''   
    #Si la celda de la columna F (Additional Notes) es igual a CRITICA se cambia el fondo de la celda de la
     #columna nueva (I - FTES. Pdtes.) a amarillo
    max_row = ws.max_row    
    for i_row in range(1, max_row + 1):   
        cell_new = ws.cell(row=i_row, column= colNr)
        cell_condition = ws.cell(row=i_row, column= colNr-3)
        if cell_condition.value == condition:
           cell_new.fill = color

      