#en este fichero vamos a incluir la ruta de entrada, la ruta de salida
#y todos los parámetros necesarios para ejecutar cedazo.py para cualquier hoja excel
import argparse

#definimos objeto parser de tipo clase ArgumentParser
parser = argparse.ArgumentParser()
#añadimos argumento ruta_input asignándole un valor constante mediante action y const
parser.add_argument('ruta_input', action='store_const', const='C:\\Users\\lgordoga\\L99PYTHON\\SPRINT_0\\in.xlsx')
#parser.add_argument('-ruta_input', '--ruta_input', type=str)
parser.add_argument('ruta_output', action='store_const', const='C:\\Users\\lgordoga\\L99PYTHON\\SPRINT_0\\out.xlsx')
#parser.add_argument('-ruta_output', '--ruta_output', type=str)
#para ejecutarlo desde terminal:
#python cedazo.py -ruta_input 'C:\Users\lgordoga\L99PYTHON\SPRINT_0\in_corto.xlsx' -ruta_output 'C:\Users\lgordoga\L99PYTHON\SPRINT_0\out.xlsx'
#para ejecutar desde powershell, hay que posicionarse en la carpeta donde está el ejecutable cedazo.py y luego igual que en terminal