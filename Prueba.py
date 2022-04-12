from os import system

system('md "C:\\tarjeta"')

ruta = "C:\\tarjeta"

f = open(ruta + "/mapeo.bat", "w")

f.write("NET USE Q: \\\\fsfnprp76c01b\\Groupdir")

f.close()



""" system('copy remito.xlsx C:\\remito') """