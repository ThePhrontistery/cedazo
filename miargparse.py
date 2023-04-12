#en este fichero vamos a incluir la ruta de entrada, la ruta de salida
#y todos los parámetros necesarios para ejecutar cedazo.py para cualquier hoja excel
import argparse

#definimos objeto parser de tipo clase ArgumentParser
parser = argparse.ArgumentParser()
#añadimos argumento ruta_input asignándole un valor constante mediante action y const
parser.add_argument('ruta_input', action='store_const', const='C:\\Users\\lgordoga\\L99PYTHON\\SPRINT_0\\in_corto.xlsx')

parser.add_argument('ruta_output', action='store_const', const='C:\\Users\\lgordoga\\L99PYTHON\\SPRINT_0\\out.xlsx')

