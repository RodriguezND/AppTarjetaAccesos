from logging import exception
from tkinter import ttk
from tkinter import *
import sqlite3
from openpyxl import load_workbook
from os import system
from tkcalendar import Calendar, DateEntry
import datetime
from tkinter import messagebox

path = 'Q:\\Support\\AppTarjeta\\remito.xlsx'

class Tarjeta:

    db_name = "dbtarjeta.db"

    def __init__(self, window):
        self.wind = window
        self.wind.title("Aplicacion de Tarjetas - Daniel-san Version 0.0.5")

        # CREAR UN FRAME CONTAINER
        frame = LabelFrame(self.wind, text = "GESTION DE TARJETA")
        frame.grid(row = 0, column = 0, columnspan = 8, pady = 20, padx = 20)
        
        cmd = "NET USE Q: \\\\fsfnprp76c01b\\Groupdir"
        system(cmd) 

        #MENSAJE SALIDA
        self.message = Label(frame, text = "", fg = "red")
        self.message.grid(row=7, column=1,columnspan= 2, sticky= W + E)

        #BOTON AGREGAR TARJETA
        ttk.Button(frame, text = "Listar Tarjeta", command = self.get_tarjetas).grid(row = 8, columnspan=8, sticky= W + E)

        #BOTON BUSCAR NUMERACION
        Label(frame, text = "Buscar por Numeracion: ").grid(row = 1, column = 1)
        self.buscar = Entry(frame)
        self.buscar.grid(row=2, column = 1)
        
        #BOTON BUSCAR NOMBRE
        Label(frame, text = "Buscar por Nombre: ").grid(row = 3, column = 1)
        self.buscarnombre = Entry(frame)
        self.buscarnombre.grid(row=4, column = 1)
        ttk.Button(frame, text = "Buscar", command = self.filtrar_tarjeta).grid(row = 5, column= 1)

        #BOTON REMITO/IMPRIMIR
        ttk.Button(frame, text = "Imprimir", command = self.imprimir_remito).grid(row = 5, column= 3)

        #TABLA
        self.tree = ttk.Treeview(frame, height= 10, columns=("1","2","3","4","5"))
        self.tree.grid(row = 9, column = 0, columnspan=8)
        self.tree.heading("#0", text= "Nombre", anchor = CENTER)
        self.tree.column("#0", minwidth=0, width=140)
        self.tree.heading("#1", text= "DNI", anchor = CENTER)
        self.tree.column("#1", minwidth=0, width=65)
        self.tree.heading("#2", text= "Numeracion", anchor = CENTER)
        self.tree.column("#2", minwidth=0, width=120)
        self.tree.heading("#3", text= "Estado", anchor = CENTER) 
        self.tree.column("#3", minwidth=0, width=70)
        self.tree.heading("#4", text= "Fecha Entrega", anchor = CENTER)   
        self.tree.column("#4", minwidth=0, width=70)
        self.tree.heading("#5", text= "Observacion", anchor = CENTER)   
        self.tree.column("#5", minwidth=0, width=100)
        """ self.tree["displaycolumns"]=("1","2","3") """
        
        #BOTON REGISTRAR, ELIMINAR Y ACTUALIZAR
        Button(frame, text= "REGISTRAR", bg="#CEFF91", command= self.registrar_tarjeta).grid(row=10, column=0, columnspan=3, sticky= W + E)
        
        Button(frame, text = "EDITAR", bg="#FFFF75", command = self.edit_tarjeta).grid(row = 10, column= 3, columnspan=3, sticky= W + E)
        Button(frame, text = "ELIMINAR",bg="#F7A4A4", command = self.delete_tarjeta).grid(row = 10, column= 6, columnspan=2, sticky= W + E)
        

        self.get_tarjetas()


    def registrar_tarjeta(self):
        
        self.agregarTarjeta = Toplevel(padx=20, pady=20)
        self.agregarTarjeta.title = "Agregar Tarjeta"
        self.agregarTarjeta.resizable(height=False, width=False)


        #NOMBRE INPUT
        Label(self.agregarTarjeta, text = "Nombre y Apellido: ").grid(row = 1, column = 0)
        self.nombre = Entry(self.agregarTarjeta)
        self.nombre.grid(row=1, column = 1)

        #DNI INPUT
        Label(self.agregarTarjeta, text = "DNI: ").grid(row = 2, column = 0)
        self.dni = Entry(self.agregarTarjeta)
        self.dni.grid(row=2, column = 1)

        #NUMERACION INPUT
        Label(self.agregarTarjeta, text = "Numeracion: ").grid(row = 3, column = 0)
        self.numeracion = Entry(self.agregarTarjeta)
        self.numeracion.grid(row=3, column = 1)

        #ESTADO INPUT
        Label(self.agregarTarjeta, text = "Estado: ").grid(row = 4, column = 0)
        self.listaDesplegableEstado = ttk.Combobox(self.agregarTarjeta, width=17, state="readonly")
        self.listaDesplegableEstado.grid(row=4, column =1)
        opciones=["Sin Asignar", "En Proceso", "Habilitada", "Deshabilitada","Extraviada/Rota"]
        self.listaDesplegableEstado['values'] = opciones

        #FECHAENTREGA INPUT
        today = datetime.date.today()
        Label(self.agregarTarjeta, text = "Fecha de Entrega: ").grid(row = 5, column = 0)
        self.fecha = DateEntry(self.agregarTarjeta, width=17, locale="es_AR", year= today.year, month = today.month, day = today.day)
        self.fecha.grid(row=5, column=1, padx=15)
        fecha = self.fecha.get_date()
        self.fechaCorregida = fecha.strftime("%d/%m/%Y") 

        #OBSERVACIONES INPUT
        Label(self.agregarTarjeta, text = "Observaciones: ").grid(row = 6, column = 0)
        self.observacion = Entry(self.agregarTarjeta)
        self.observacion.grid(row=6, column = 1) 

        #BOTON AGREGAR TARJETA
        Button(self.agregarTarjeta, text = "Agregar Tarjeta", command = self.add_tarjeta).grid(row = 7, columnspan=8, sticky= W + E)
        

    def run_query(self, query, parameters = ()):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result

    def get_tarjetas(self):
        #Limpiando tabla
        records = self.tree.get_children()
        for element in records:
            self.tree.delete(element)
        #Consultando datos en la tabla
        query = "SELECT * FROM dbtarjeta ORDER BY Numeracion ASC"
        db_rows = self.run_query(query)
        for row in db_rows:
            self.tree.insert("", 0, text = row[1], values = (row[2],row[3],row[4],row[5], row[6])) 

    def cambiar_mes(self, fechaarg):
        fecha = fechaarg.split("/")
        if fecha[1] == "01":
            fecha[1] = "Enero"
        if fecha[1]== "02":
            fecha[1] = "Febrero"
        if fecha[1] == "03":
            fecha[1] = "Marzo"
        if fecha[1] == "04":
            fecha[1] = "Abril"
        if fecha[1] == "05":
            fecha[1] = "Mayo"
        if fecha[1] == "06":
            fecha[1] = "Junio"
        if fecha[1] == "07":
            fecha[1] = "Julio"
        if fecha[1] == "08":
            fecha[1] = "Agosto"
        if fecha[1] == "09":
            fecha[1] = "Septiembre"
        if fecha[1] == "10":
            fecha[1] = "Octubre"
        if fecha[1] == "11":
            fecha[1] = "Noviembre"
        if fecha[1] == "12":
            fecha[1] = "Diciembre"

        return fecha

    #IMPRIMIR TARJETA
    def imprimir_remito(self):
        self.message["text"] = ""
        
        try:
            system('md "C:\\remito"')
            system('copy remito.xlsx C:\\remito')
            
            book = load_workbook(path)
            hoja1=book["Plan1"]
            hoja2=book["Plan2"]
            """ hoja = book.active """
            
            nombre= self.tree.item(self.tree.selection())["text"]
            dni = self.tree.item(self.tree.selection())["values"][0]
            numeracion = self.tree.item(self.tree.selection())["values"][1]
            fecha = self.tree.item(self.tree.selection())["values"][3]

            if(nombre == "" or dni == "" or numeracion == "" or fecha == ""):
                self.message["text"] = "Completa Nombre, DNI, Numeracion o Fecha"
            else:
                hoja1['B34'] = nombre
                hoja1['B38'] = dni
                hoja1['D24'] = numeracion

                fechaMes = self.cambiar_mes(fecha)
                hoja1['D13'] = "{} de {} de {}".format(fechaMes[0],fechaMes[1],fechaMes[2])
                
                hoja2['C5'] = nombre
                hoja2['C6'] = dni
                hoja2['C4'] = numeracion
            
                book.save(path)
                
                system('start {}'.format(path))
        except Exception as e:
            print(type(e).__name__)
            self.message["text"] = "Selecciona una tarjeta para imprimir"

        

    #FILTRAR TARJETA
    def filtrar_tarjeta(self):
        if len(self.buscar.get()) != 0:
        #Limpiando tabla
            records = self.tree.get_children()
            for element in records:
                self.tree.delete(element)
        #Consultando datos en la tabla
            query = "SELECT * FROM dbtarjeta WHERE Numeracion LIKE ?"
            
            valor = self.buscar.get().strip()
            
            db_rows = self.run_query(query,('%'+valor+'%',))
            
            for row in db_rows:
                
                self.tree.insert("", 0, text = row[1], values = (row[2],row[3],row[4],row[5], row[6]))

            self.message["text"] = "Busqueda completa"
        elif len(self.buscarnombre.get()) != 0:
        #Limpiando tabla
            records = self.tree.get_children()
            for element in records:
                self.tree.delete(element)
        #Consultando datos en la tabla
            query = "SELECT * FROM dbtarjeta WHERE NombreApellido LIKE ?"
            
            valor = self.buscarnombre.get().strip()
            
            db_rows = self.run_query(query,('%'+valor+'%',))
            
            for row in db_rows:
                
                self.tree.insert("", 0, text = row[1], values = (row[2],row[3],row[4],row[5], row[6]))

            self.message["text"] = "Busqueda completa"


    def validacion(self):
        return len(self.nombre.get()) != 0 and len(self.numeracion.get()) != 0 and len(self.listaDesplegableEstado.get()) !=0 and len(self.fechaCorregida) != 0

    #FUNCION PARA AGREGAR TARJETA
    def add_tarjeta(self):
        if self.validacion():
            try:
                query = "INSERT INTO dbtarjeta VALUES(NULL, ? ,?, ?, ?, ?, ?)"
                parameters = (self.nombre.get(), self.dni.get(), self.numeracion.get(), self.listaDesplegableEstado.get(), self.fechaCorregida, self.observacion.get())
                self.run_query(query,parameters)
                self.message["text"] = "La tarjeta se ha registrado"
                self.limparCampos()
            except:
                self.message["text"] = "Volve a ingresar la tarjeta"
        else: 
            self.message["text"] = "Completa todos los campos"
        self.get_tarjetas()

    #LIMPIAR CAMPOS
    def limparCampos(self):
        self.nombre.delete(0, END)
        self.dni.delete(0, END)
        self.numeracion.delete(0, END)
        self.observacion.delete(0, END)

    #FUNCION PARA ELIMINAR TARJETA
    def delete_tarjeta(self):
        self.message["text"] = ""
        try:
            self.tree.item(self.tree.selection())["values"][1]
        except IndexError as e:
            self.message["text"] = "Selecciona una tarjeta"
            return
        self.message["text"] = ""
        equipo = self.tree.item(self.tree.selection())["values"][1]

        cartel = messagebox.askyesno("ALERTA", "¿Seguro que queres eliminar?")
        if cartel == True:
            self.delete_confirmation(equipo)
 
        """ self.confirmation = Toplevel()
        self.confirmation.title = "Eliminar Tarjeta"
        Label(self.confirmation, text = "¿Realmente desea eliminar la tarjeta?").grid(row=0, column=0, columnspan=2, pady=20, ipadx=20, ipady=10, padx=20)
        ttk.Button(self.confirmation, text="Si", command= lambda: self.delete_confirmation(equipo)).grid(row=2, column=0, ipady=5)
        ttk.Button(self.confirmation, text="No", command=self.confirmation.destroy).grid(row=2, column=1, ipady=5) """

        self.get_tarjetas()

    def delete_confirmation(self, dato):
        
        query =  "DELETE FROM dbtarjeta WHERE Numeracion = ?"
        self.run_query(query, (dato,))
        self.message["text"] = "La tarjeta {} fue eliminada".format(dato)
        """ self.confirmation.destroy() """
        self.get_tarjetas()

    #FUNCION PARA EDITAR MOCHILA
    def edit_tarjeta(self):
        self.message["text"] = ""
        try:
            self.tree.item(self.tree.selection())["values"][1]
        except IndexError as e:
            self.message["text"] = "Selecciona una tarjeta"
            return
        self.message["text"] = ""

        nombre= self.tree.item(self.tree.selection())["text"]
        dni= self.tree.item(self.tree.selection())["values"][0]
        numeracion= self.tree.item(self.tree.selection())["values"][1]
        estado=self.tree.item(self.tree.selection())["values"][2]
        fecha=self.tree.item(self.tree.selection())["values"][3]
        observacion=self.tree.item(self.tree.selection())["values"][4]

        self.edit_wind = Toplevel(padx=20, pady=20)
        self.edit_wind.title = "Editar tarjeta"
        self.edit_wind.resizable(height=False, width=False)
        
        
        #Anterior nombre
        Label(self.edit_wind, text = "Nombre y Apellido: ").grid(row = 0, column = 1)
        nuevo_nombre = Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value = nombre))
        nuevo_nombre.grid(row=0, column=2)

        #DNI
        Label(self.edit_wind, text = "DNI: ").grid(row = 2, column = 1)
        nuevo_dni = Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value = dni))
        nuevo_dni.grid(row=2, column=2)

        #Numeracion
        Label(self.edit_wind, text = "Numeracion: ").grid(row = 4, column = 1)
        nuevo_numeracion = Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value = numeracion))
        nuevo_numeracion.grid(row=4, column=2)

        #Estado
        Label(self.edit_wind, text = "Estado: ").grid(row = 6, column = 1)
        
        listaDesplegableEstado = ttk.Combobox(self.edit_wind, width=17, state="readonly")
        listaDesplegableEstado.grid(row = 6, column = 2)
        listaDesplegableEstado.set(estado)
        opciones=["Sin Asignar", "En Proceso", "Habilitada", "Deshabilitada","Extraviada/Rota"]
        listaDesplegableEstado['values'] = opciones

        #Fecha
        today = datetime.date.today()
        Label(self.edit_wind, text = "Fecha: ").grid(row = 8, column = 1)
        nuevaFecha = DateEntry(self.edit_wind, width=17, locale="es_AR", year= today.year, month = today.month, day = today.day)
        nuevaFecha.grid(row=8, column=2, padx=15)

        if fecha == None or fecha=="":
            fecha = str(today.day) + "/" + str(today.month) + "/" + str(today.year) 
            
        nuevaFecha.set_date(fecha)

        #Observacion
        Label(self.edit_wind, text = "Observacion: ").grid(row = 10, column = 1)
        nuevo_observacion = Entry(self.edit_wind, textvariable=StringVar(self.edit_wind, value = observacion))
        nuevo_observacion.grid(row=10, column=2)

        #BOTON ACTUALIZAR
        Button(self.edit_wind,text = "Actualizar", command = lambda: self.edit_registro(nuevo_nombre.get(), nuevo_dni.get(), nuevo_numeracion.get(), listaDesplegableEstado.get(), nuevaFecha.get_date(), nuevo_observacion.get(),numeracion)).grid(row = 11, columnspan=8, sticky= W + E)

    def edit_registro(self, nuevo_nombre, nuevo_dni, nuevo_numeracion, nuevo_estado, nuevo_fecha, nuevo_observacion, numeracion):
        
        fechasNueva = nuevo_fecha.strftime("%d/%m/%Y")

        query = "UPDATE dbtarjeta SET NombreApellido = ?, DNI = ?, Numeracion = ?, Estado = ?, FechaEntrega = ?, Observaciones = ? WHERE Numeracion = ?"
        parameters = (nuevo_nombre,nuevo_dni,nuevo_numeracion,nuevo_estado, fechasNueva, nuevo_observacion, numeracion)
        self.run_query(query, parameters)
        self.edit_wind.destroy()
        self.message["text"] = "La tarjeta se actualizo"
        self.get_tarjetas()
    