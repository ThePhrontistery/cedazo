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

def remove(ws): 
    '''Metodo que sirve para borrar todas las filas de 1 a 11'''
    ws.delete_rows(1, 11)


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