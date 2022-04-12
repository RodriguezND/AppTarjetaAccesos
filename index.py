from tkinter import *
from Tarjetas import Tarjeta


if __name__ == "__main__":
    window = Tk()

    window.resizable(height=False, width=False)
    
    application = Tarjeta(window)
    
    #SETEA TAMAÑO Y FIJA VENTANA CENTRO DE LA PANTALLA
    window_altura = 461
    window_anchura = 611

    Pantalla_ancho = window.winfo_screenwidth()
    Pantalla_alto= window.winfo_screenheight()

    y_coordinate = int((Pantalla_alto/2) - (window_altura/2))
    x_coordinate = int((Pantalla_ancho/2) - (window_anchura/2))

    window.geometry("{}x{}+{}+{}".format(window_anchura, window_altura, x_coordinate, y_coordinate))


    window.mainloop()
 

 