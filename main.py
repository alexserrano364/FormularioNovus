import openpyxl
from openpyxl import *
from openpyxl.styles import Font
from datetime import datetime
import os.path
from tkinter import *
from tkinter import messagebox
from tkinter import font

meses = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
         9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

blue = "#686bb0"

def obtenerHora():
    today = datetime.now()
    hour = today.hour
    minute = today.minute
    return [hour, minute]


def obtenerFecha():
    today = datetime.now()
    year = today.year
    month = today.month
    day = today.day
    return [day, month, year]


def prepararEncabezado(sheet):
    # Hace que toda la primera fila sea negritas
    header = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"]
    for cell in header:
        activeCell = sheet[cell]
        activeCell.font = Font(bold=True)
    sheet['A1'] = "Nombre"
    sheet['B1'] = "Celular"
    sheet['C1'] = "Correo Electrónico"
    sheet['D1'] = "Código Postal"
    sheet['E1'] = "Servicio"
    sheet['F1'] = "Importe"
    sheet['G1'] = "Turno"
    sheet['H1'] = "Hora"


def conseguirDatos(botones):
    datos = []
    for boton in botones:           # Recorre todos los botones y los convierte en texto
        dato = boton.get()
        datos.append(dato)
    return datos


def crearExcel(root, botones):
    # Se obtiene fecha y hora
    fecha = obtenerFecha()
    hora = obtenerHora()
    stringFecha = "%02d/%02d/%02d" % (fecha[0], fecha[1], fecha[2])
    stringHora = "%02d:%02d" % (hora[0], hora[1])

    if 0 <= hora[0] < 7:  # Si es turno de madrugada, se añade a la base del día anterior
        fecha[0] -= 1

    # Se crea el Excel del día, si ya existe solo se carga
    title = 'Corte Caja %d de %s del %d.xlsx' % (fecha[0], meses[fecha[1]], fecha[2])
    if os.path.exists(title):
        workbook = load_workbook(title)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        prepararEncabezado(sheet)

    datos = conseguirDatos(botones)

    # Se pide y añade la info a la base del día
    datos.append(stringHora)
    sheet.append(datos)
    workbook.save(title)
    workbook.close()

    # Se añade la info a la base general que ya existe
    title = "Base General.xlsx"
    workbook = load_workbook(title)
    sheet = workbook.active
    datos.append(stringFecha)
    sheet.append(datos)
    workbook.save(title)
    workbook.close()

    messagebox.showinfo("Éxito", "Datos agregados")
    root.quit()


def crearInfoVentana(root):

    fuenteLabels = font.Font(family="Arial", size=12)
    fuenteMenu = font.Font(family="Arial", size=8)
    fuenteBoton = font.Font(family="Arial", size=10)

    labelNombre = Label(root, text="Nombre:", font=fuenteLabels)
    nombre = Entry(root, width=85)

    labelTurno = Label(root, text="Turno:", font=fuenteLabels)
    turno = StringVar()
    turno.set("Seleccionar")
    turnoMenu = OptionMenu(root, turno, "Matutino", "Vespertino", "Nocturno")
    turnoMenu.config(font=fuenteBoton)
    turnoDrop = root.nametowidget(turnoMenu.menuname)
    turnoDrop.config(font=fuenteMenu)                 # Se le empata la fuente de las labels a las opciones

    labelServicio = Label(root, text="Servicio:", font=fuenteLabels)
    servicio = Entry(root, width=85)

    labelImporte = Label(root, text="Importe:", font=fuenteLabels)
    importe = Entry(root, width=15)

    labelCelular = Label(root, text="Número Celular:", font=fuenteLabels)
    celular = Entry(root, width=85)

    labelCP = Label(root, text="Código Postal:", font=fuenteLabels)
    CP = Entry(root, width=15)

    labelCorreo = Label(root, text="Correo Electrónico:", font=fuenteLabels)
    correo = Entry(root, width=85)

    botones = [nombre, celular, correo, CP, servicio, importe, turno]

    boton = Button(root, text="Insertar", padx=10, command=lambda: crearExcel(root, botones),
                   bg=blue, fg="white", font=fuenteBoton)

    # Insertar labels
    labelNombre.grid(row=1, column=1, padx=15, pady=5, sticky="W")
    nombre.grid(row=2, column=1, padx=15, pady=5, sticky="W")
    labelTurno.grid(row=1, column=3, padx=15, pady=5, sticky="W")
    turnoMenu.grid(row=2, column=3, padx=15, pady=5, sticky="W")
    labelServicio.grid(row=4, column=1, padx=15, sticky="W")
    servicio.grid(row=5, column=1, padx=15, pady=10, sticky="W")
    labelImporte.grid(row=4, column=3, padx=15, sticky="W")
    importe.grid(row=5, column=3, padx=15, pady=10, sticky="W")
    labelCelular.grid(row=7, column=1, padx=15, sticky="W")
    celular.grid(row=8, column=1, padx=15, pady=10, sticky="W")
    labelCP.grid(row=7, column=3, padx=15, sticky="W")
    CP.grid(row=8, column=3, padx=15, pady=10, sticky="W")
    labelCorreo.grid(row=10, column=1, padx=15, sticky="W")
    correo.grid(row=11, column=1, padx=15, pady=10, sticky="W")
    boton.grid(row=12, column=3, padx=15, pady=10)




def main():
    # Se crea la ventana
    root = Tk()
    crearInfoVentana(root)
    root.mainloop()


main()
