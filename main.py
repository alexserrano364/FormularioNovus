import openpyxl
from openpyxl import *
from openpyxl.styles import Font
from datetime import datetime
import os.path
from tkinter import *
from tkinter import messagebox
from tkinter import font
from PIL import ImageTk, Image

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
    header = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"]
    for cell in header:
        activeCell = sheet[cell]
        activeCell.font = Font(bold=True)
    sheet['A1'] = "Nombre"
    sheet['B1'] = "Celular"
    sheet['C1'] = "Correo Electrónico"
    sheet['D1'] = "Código Postal"
    sheet['E1'] = "Servicio"
    sheet['F1'] = "Importe"
    sheet['G1'] = "Hora"
    sheet['H1'] = "Folio"


def conseguirDatos(botones):
    datos = []
    for boton in botones:           # Recorre todos los botones y los convierte en texto
        dato = boton.get()
        datos.append(dato)
    return datos


def datosSonValidos(datos):
    datos[0] = datos[0].upper()                     # Pasa el nombre a mayúsculas

    datos[1] = datos[1].replace(" ","")             # Se deshace de espacios y guiones en el núm. tel.
    datos[1] = datos[1].replace("-","")

    datos[4] = datos[4].upper()                     # Pasa el servicio a mayúsculas

    try:                                # Checa que el número telefónico sea un número
        int(datos[1])
    except ValueError:
        messagebox.showinfo("Alerta", "El número telefónico debe ser un número")
        return False

    try:                                # Checa que el código postal sea un número
        int(datos[3])
    except ValueError:
        messagebox.showinfo("Alerta", "El código postal debe ser un número")
        return False

    try:                                # Checa que el importe sea un número
        float(datos[5])
    except ValueError:
        messagebox.showinfo("Alerta", "El importe debe ser un número")
        return False
    else:
        datos[5] = "$" + datos[5]

    if datos[6] == "Seleccionar":       # Checa que el turno haya sido seleccionado
        messagebox.showinfo("Alerta", "Seleccionar Turno")
        return False

    if datos[7] == "Seleccionar":       # Checa que el tipo de pago haya sido seleccionado
        messagebox.showinfo("Alerta", "Seleccionar Tipo de Pago")
        return False

    return True


def crearExcel(root, botones):
    datos = conseguirDatos(botones)
    if datosSonValidos(datos):                # Solo funciona si los datos están bien

        # Se obtienen datos de fecha y hora
        fecha = obtenerFecha()
        hora = obtenerHora()
        if 0 <= hora[0] < 7:  # Si es turno de madrugada, es parte del corte del día anterior
            fecha[0] -= 1

        # Se crea el Excel del turno, si ya existe solo se carga
        turno = datos[6][:1]
        title = 'Corte Caja %d de %s del %d - %s.xlsx' % (fecha[0], meses[fecha[1]], fecha[2], turno)
        if os.path.exists(title):
            workbook = load_workbook(title)
            sheet = workbook.active
        else:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            prepararEncabezado(sheet)

        # Se crea el folio y se da formato a la fecha y hora
        seriado = len(sheet['H'])
        folio = "%03d-%04d%02d%02d%s" % (seriado, fecha[2], fecha[1], fecha[0], turno)
        stringHora = "%02d:%02d" % (hora[0], hora[1])
        stringFecha = "%02d/%02d/%02d" % (fecha[0], fecha[1], fecha[2])

        # Se hace el formato para el corte de caja
        datosCorteDeCaja = [folio, datos[0], datos[4]]

        if datos[7] == "Tarjeta":               # Revisa cómo fue el pago
            datosCorteDeCaja.append("")
            datosCorteDeCaja.append(datos[5])
        else:
            datosCorteDeCaja.append(datos[5])
            datosCorteDeCaja.append("")
        datos.remove(datos[7])

        sheet.append(datosCorteDeCaja)
        workbook.save(title)
        workbook.close()

        # Se añade la info a la base general que ya existe
        title = "Base General.xlsx"
        workbook = load_workbook(title)
        sheet = workbook.active
        datos.remove(datos[6])
        datos.append(stringHora)
        datos.append(folio)
        datos.append(stringFecha)
        sheet.append(datos)
        workbook.save(title)
        workbook.close()

        messagebox.showinfo("Éxito", "Datos agregados")
        root.quit()


def crearInfoVentana(root):

    padx = 65
    pady = 20

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
    turnoDropMenu = root.nametowidget(turnoMenu.menuname)
    turnoDropMenu.config(font=fuenteMenu)                 # Se le empata la fuente de las labels a las opciones

    labelServicio = Label(root, text="Servicio:", font=fuenteLabels)
    servicio = Entry(root, width=85)

    labelImporte = Label(root, text="Importe:", font=fuenteLabels)
    importe = Entry(root, width=15)

    labelPago = Label(root, text="Método de pago:", font=fuenteLabels)
    tipoPago = StringVar()
    tipoPago.set("Seleccionar")
    pagoMenu = OptionMenu(root, tipoPago, "Efectivo", "Tarjeta")
    pagoMenu.config(font=fuenteBoton)
    turnoDropPago = root.nametowidget(pagoMenu.menuname)
    turnoDropPago.config(font=fuenteMenu)


    labelCelular = Label(root, text="Número Celular:", font=fuenteLabels)
    celular = Entry(root, width=85)

    labelCorreo = Label(root, text="Correo Electrónico:", font=fuenteLabels)
    correo = Entry(root, width=85)

    labelCP = Label(root, text="Código Postal:", font=fuenteLabels)
    CP = Entry(root, width=15)

    imgLogo = ImageTk.PhotoImage(Image.open("./img/NovusLogo.jpeg"))
    labelImagen = Label(image=imgLogo)

    botones = [nombre, celular, correo, CP, servicio, importe, turno, tipoPago]

    boton = Button(root, text="Insertar", padx=10,
                   command=lambda: crearExcel(root, botones),
                   bg=blue, fg="white", font=fuenteBoton, width=30, height=5)

    # Insertar labels
    labelNombre.grid(row=1, column=1, padx=padx, pady=pady, sticky="W")
    nombre.grid(row=2, column=1, padx=padx, sticky="W")
    labelTurno.grid(row=1, column=3, padx=padx, pady=pady, sticky="W")
    turnoMenu.grid(row=2, column=3, padx=padx, sticky="W")
    labelServicio.grid(row=4, column=1, padx=padx, pady=pady, sticky="W")
    servicio.grid(row=5, column=1, padx=padx, sticky="W")
    labelImporte.grid(row=4, column=3, padx=padx, pady=pady, sticky="W")
    importe.grid(row=5, column=3, padx=padx, sticky="W")
    labelPago.grid(row=7, column=3,padx=padx, sticky="W")
    pagoMenu.grid(row=8, column=3,padx=padx, sticky="W")
    labelCelular.grid(row=7, column=1, padx=padx, pady=pady, sticky="W")
    celular.grid(row=8, column=1, padx=padx, sticky="W")
    labelCP.grid(row=9, column=3, padx=padx, pady=pady, sticky="W")
    CP.grid(row=10, column=3, padx=padx, sticky="W")
    labelCorreo.grid(row=9, column=1, padx=padx, pady=pady, sticky="W")
    correo.grid(row=10, column=1, padx=padx, sticky="W")
    boton.grid(row=11, column=3, padx=padx, pady=40)
    labelImagen.grid(row=11, column=1)


def main():
    # Se crea la ventana
    root = Tk()
    root.title("Sistema de Corte de Caja Novus")
    root.iconbitmap('./img/caduceus.ico')
    crearInfoVentana(root)
    root.mainloop()


main()
