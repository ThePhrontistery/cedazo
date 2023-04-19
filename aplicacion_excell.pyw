from tkinter import *

root= Tk()
#le damos título a la interface
root.title("Aplicación Picadora")

miFrame=Frame(root)
#con el método pack se coloca el objeto miFrame en la ventana principal y se ancha a la parte inferior de la ventana
#principal con anchor ="s", anchor me da la posición relativa dentro del contenedor (lleva los valores según las
#coordenadas n,s,e,w,ne,nw,se,sw o center)
miFrame.pack(anchor="s")
#config se utilizará para configurar nuestro miFrame, de momento no pongo nada no haría falta tenerlo
miFrame.config()

#creamos una variable en la que decimos que es una cadena de caracteres (StringVar)
#se codifica aquí porque necesito hacerlo por encima del objeto al que le voy a asociar
infile=StringVar()

#el widgets Label será un cuadro de texto que contendrá el valor del argumento text
textinput=Label(miFrame, text="File IN: ")
#con la siguiente línea lo estamos posicionando con grid en una cuadrícula en la fila 0 y la columna 0, 
#con sticky="e" le indicamos hacia donde debe justificarse hadia el este(derecha de la celda), mismos valores que anchor,
#con padx y pady le estamos indicando el relleno horizontal y vertical alrededor del widget
textinput.grid(row=0,column=0, sticky="e", padx=10, pady=10)
#el widgets Entry será una caja donde el usuario podrá introducir información
#con textvariable le voy a asociar a la cajita del fichero in la variable infile 
input=Entry(miFrame, width=50, textvariable=infile)
input.grid(row=0,column=1, padx=10, pady=10)

textoutput=Label(miFrame, text="File OUT: ")
textoutput.grid(row=1,column=0,sticky="e", padx=10, pady=10)
output=Entry(miFrame, width=50)
output.grid(row=1,column=1,padx=10, pady=10)

#creo un widget Checkbutton que voy a posicionar en root 
#empaqueto el widget en la ventana principal, indicando que se colocará en el lado iaquierdo (side="left"-> 
# valores disponibles top, bottom, left, right)
#los demás widgets se colocarán a su derecha
Checkbutton(root, text="Merge").pack(side="left",padx=10, pady=10)

def codigoBoton():
    infile.set("Introducir fichero de entrada")

botonOk=Button(root, text="OK", command=codigoBoton)
botonOk.pack(side="right", padx=5, pady=20)
botonCancel=Button(root, text="Cancel")
botonCancel.pack(side="right", padx=5, pady=20)

root.mainloop()