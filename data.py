from types import CellType
import openpyxl
from openpyxl.styles import PatternFill

"""def remove(ws):

    # iterate the sheet by rows
    for row in ws.iter_rows():

    # all() return False if all of the row value is None
        if not all(cell.value for cell in row):

    # detele the empty row

            ws.delete_rows(row[0].row, 1)

    # recursively call the remove() with modified sheet data
            remove(ws)
            

            return"""

def remove_empty(ws): 
#Sirve para borrar todas las filas que estan vacias, (no se utiliza de momento)
    filas = ws.max_row
    for i in range(filas, 0, -1):
        celdas_vacias = all([cell.value is None for cell in ws[i]])
        if celdas_vacias:
            ws.delete_rows(i, 1)

def remove(ws): 
#Sirve para borrar todas las filas de 1 a 11
    ws.delete_rows(1, 11)


def change_colour(ws): 
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
           
def remove_ormatting(ws):
    # ws is not the worksheet name, but the worksheet object
    filas = ws.max_row
    for i_row in range(filas, 0, -1):
        for cell in ws[i_row]:
            cell.style = 'Normal'

def unmerge_cells(ws):
    print(ws.merged_cells)
    for merge in list(ws.merged_cells):
        ws.unmerge_cells(range_string=str(merge))