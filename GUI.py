# -*- coding: utf-8 -*-
import glob
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as font
from openpyxl import load_workbook
from openpyxl.styles import Font
import Diccionari


class startWindow:
    """Primera finestra"""
    carpeta = ""

    def __init__(self):
        """Inicialitza la primera pantalla"""

        # Declara la finestra de l'aplicació
        # Treu la barra de tk
        self.arrel = tk.Tk()
        self.arrel.overrideredirect(True)

        # Defineix dimensions de la finestra ample x alt 300x200
        # que se situarà en la coordenada x=500,y=50
        # Centrem la finestra a la pantalla
        amplada_finestra = 600
        altura_finestra = 160
        amplada_monitor = self.arrel.winfo_screenwidth()
        altura_monitor = self.arrel.winfo_screenheight()
        x = round(amplada_monitor / 2 - amplada_finestra / 2)
        y = round(altura_monitor / 2 - altura_finestra / 2)

        self.arrel.geometry(f'{amplada_finestra}x{altura_finestra}+{x}+{y}')

        # Frames
        # Crea un frame per a la barra nova de títol
        back_ground = '#1d1d1d'
        title_barframe = tk.Frame(self.arrel, width=535, height=20, bg=back_ground, relief='raised', bd=1, pady=3,
                                  highlightcolor=back_ground, highlightthickness=0)
        # crear frame per al botó tancar
        close_frame = tk.Frame(self.arrel, bg=back_ground, width=10, height=10, relief='raised', bd=1,
                               highlightcolor=back_ground, highlightthickness=0)

        # Crea un frame per a la sel·leció de la carpeta
        folder_frame = tk.Frame(self.arrel, bg=back_ground, width=555, height=200)

        # Configurar grid
        self.arrel.columnconfigure(0, weight=1)

        # Grid Frames
        title_barframe.grid(row=0, sticky=tk.EW)
        close_frame.grid(row=0, sticky=tk.NE)
        folder_frame.grid(row=1, sticky=tk.NSEW)

        # Widggets
        # Títol finestra
        title_name = tk.Label(title_barframe, text="Financial Control", bg=back_ground, fg='white')
        # Crea un botó per tancar a la barra de títol
        close_button = tk.Button(close_frame, text='x', command=self.arrel.destroy, bg=back_ground,
                                 activebackground="red", bd=0, font="bold", fg='white', activeforeground="white",
                                 highlightthickness=0)
        # Etiqueta carpeta
        folder_label = tk.Label(folder_frame, text="Sel·lecciona la carpeta d'excels",
                                bg=back_ground, fg='white', padx=15, pady=30)
        # Entrada text carpeta
        self.carpeta = tk.StringVar()
        entry_folder = ttk.Entry(folder_frame, textvariable=self.carpeta, justify=tk.LEFT, width=50,
                                 background=back_ground)
        # Botó per buscar carpeta
        button_font = font.Font(family="Helvetica", size=8, weight="bold")
        search_button = tk.Button(folder_frame, text='Buscar carpeta', bg='#b5b5b5', activebackground="#ffffff", bd=0,
                                  font="bold", fg='black', activeforeground="black", command=self.get_folder_path)
        search_button['font'] = button_font
        # Botó continuar
        continue_button = tk.Button(folder_frame, text='Continuar', bg='#b5b5b5', activebackground="#ffffff", bd=0,
                                    font="bold", fg='black', activeforeground="black", command=self.next_window)
        continue_button['font'] = button_font

        # Grid widgets
        title_name.grid(row=0, column=0, columnspan=7, sticky=tk.NS)
        close_button.grid(sticky=tk.NE)
        folder_label.grid(row=0, column=0, sticky=tk.W)
        entry_folder.grid(row=0, column=1, sticky=tk.W)
        search_button.grid(row=0, column=2, sticky=tk.E, padx=5)
        continue_button.grid(row=1, column=2, sticky=tk.SW, padx=5, pady=20)

        # Esdeveniment amb bind per poder moure la finestra

        def move_window(event):
            self.arrel.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

        # Els botons canvien de color al passar per damunt
        def change_on_hovering(event):
            close_button.configure(bg='red')

        def change_search_on_hovering(event):
            search_button.configure(bg='black', fg='white')

        def change_continue_on_hovering(event):
            continue_button.configure(bg='black', fg='white')

        def return_to_normal_state(event):
            close_button.configure(bg=back_ground)

        def return_search_to_normal_state(event):
            search_button.configure(bg='#b5b5b5', fg='black')

        def return_continue_to_normal_state(event):
            continue_button.configure(bg='#b5b5b5', fg='black')

        title_barframe.bind('<B1-Motion>', move_window)
        close_button.bind('<Enter>', change_on_hovering)
        close_button.bind('<Leave>', return_to_normal_state)
        search_button.bind('<Enter>', change_search_on_hovering)
        search_button.bind('<Leave>', return_search_to_normal_state)
        continue_button.bind('<Enter>', change_continue_on_hovering)
        continue_button.bind('<Leave>', return_continue_to_normal_state)

        self.arrel.mainloop()

    def get_folder_path(self):
        """Obte la ruta de la carpeta i la llista d'excels"""
        self.folder_path = tk.filedialog.askdirectory(initialdir=r"C:\Users\ferra\OneDrive\Tesla\Economia",
                                                      title="Sel·lecciona una carpeta")
        self.carpeta.set(self.folder_path)

    def next_window(self):
        """Tanca la primera finestra i obre la següent"""
        ## -----Implementar error de carpeta----------
        startWindow.carpeta = self.carpeta.get()
        self.arrel.destroy()
        secondWindow(startWindow)


class secondWindow(startWindow):
    """Segona finestra"""

    # Funcio iterar concepte per classificarlo

    def check_concept(self, excel):
        """Classifica els conceptes segons el diccionari i si no estan classificats crea una interfície per
        classificarlos """
        self.ws_act = excel.active

        # iterem per cada una de les files
        for j in range(2, self.ws_act.max_row + 1):
            concept = str(self.ws_act.cell(row=j, column=1).value).lower()
            # iterem pel diccionari de conceptes
            for elem in Diccionari.classificació.values():
                for res in elem:
                    if res in concept:
                        self.ws_act.cell(row=j, column=5).value = list(Diccionari.classificació.keys())[
                            list(Diccionari.classificació.values()).index(elem)]
        # Creem la segona pantalla
        self.sw = tk.Tk()
        self.sw.overrideredirect(True)

        # Configurar grid
        self.sw.rowconfigure(0, weight=1)
        self.sw.columnconfigure(0, weight=1)

        amplada_finestra = 900
        altura_finestra = 700
        amplada_monitor = self.sw.winfo_screenwidth()
        altura_monitor = self.sw.winfo_screenheight()
        x = round(amplada_monitor / 2 - amplada_finestra / 2)
        y = round((altura_monitor - 50) / 2 - altura_finestra / 2)

        self.sw.geometry(f'{amplada_finestra}x{altura_finestra}+{x}+{y}')

        # Frames
        # Creem un frame general
        self.frame_main = tk.Frame(self.sw, bg="gray", width=amplada_finestra, height=altura_finestra)
        self.frame_main.grid(sticky=tk.NSEW)
        # Crea un frame per a la barra nova de títol
        back_ground = '#1d1d1d'
        title_barframe = tk.Frame(self.frame_main, width=amplada_finestra, height=20, bg=back_ground, relief='raised', bd=1,
                                  pady=3, highlightcolor=back_ground, highlightthickness=0)
        # crear frame per al botó tancar
        close_frame = tk.Frame(self.frame_main, bg=back_ground, width=10, height=10, relief='raised', bd=1,
                               highlightcolor=back_ground, highlightthickness=0)

        # Crea un frame gestió excel
        gestio_frame = tk.Frame(self.frame_main, bg=back_ground)

        # Creem un frame per al canvas que allotjarà la taula
        self.canvas_frame =tk.Frame(self.frame_main, bg=back_ground)

        # Grid Frames
        title_barframe.grid(row=0, sticky=tk.EW)
        close_frame.grid(row=0, sticky=tk.NE)
        gestio_frame.grid(row=1, sticky=tk.NSEW)
        self.canvas_frame.grid(row=2, column=0, sticky=tk.NSEW)
        self.canvas_frame.rowconfigure(0, weight=1)
        self.canvas_frame.columnconfigure(0, weight=1)
        self.canvas_frame.grid_propagate(False)

        # Widggets
        # Títol finestra
        title_name = tk.Label(title_barframe, text="Financial Control", bg=back_ground, fg='white')
        # Crea un botó per tancar a la barra de títol
        close_button = tk.Button(close_frame, text='x', command=self.sw.destroy, bg=back_ground,
                                 activebackground="red", bd=0, font="bold", fg='white',
                                 activeforeground="white",
                                 highlightthickness=0)
        # Labels gestió excel
        label_gestionant = tk.Label(gestio_frame, text=f'Gestionant despeses del mes', bg=back_ground,
                                    fg='white', font=font.Font(family="Helvetica", size=15, weight="bold"),
                                    padx=20, pady=15)
        label_mes = tk.Label(gestio_frame, text=self.nom_fulla,
                             font=font.Font(family="Helvetica", size=25, weight="bold"),
                             padx=20, pady=10, bg=back_ground, fg='white')

        # Grid widgets
        title_name.grid(row=0, column=0, columnspan=7, sticky=tk.NS)
        close_button.grid(sticky=tk.NE)
        label_gestionant.pack(fill=tk.X)
        label_mes.pack(fill=tk.X)

        def move_window(event):
            self.sw.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

        # Els botons canvien de color al passar per damunt
        def change_on_hovering(event):
            close_button.configure(bg='red')

        def return_to_normal_state(event):
            close_button.configure(bg=back_ground)

        title_barframe.bind('<B1-Motion>', move_window)
        close_button.bind('<Enter>', change_on_hovering)
        close_button.bind('<Leave>', return_to_normal_state)

        # Llista de conceptes sense classificar
        self.rows = []
        for j in range(2, self.ws_act.max_row + 1):
            if self.ws_act.cell(row=j, column=5).value is None:
                self.rows.append(j)

        # creació taula

        self.taula(Diccionari.classificació)

        self.sw.mainloop()

    def taula(self, diccionari):
        # creem un canvas per allotjar la scrollbar
        # Add a canvas in that frame
        canvas = tk.Canvas(self.canvas_frame, bg="white")
        canvas.grid(row=0, column=0, sticky=tk.NSEW)

        # Link a scrollbar to the canvas
        vsb = tk.Scrollbar(self.canvas_frame, orient="vertical", command=canvas.yview)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        canvas.configure(yscrollcommand=vsb.set)

        # Crea un frame per la taula de conceptes
        self.con_frame = tk.Frame(canvas, bg='#1d1d1d')
        canvas.create_window((0, 0), window=self.con_frame, anchor=tk.NW)

        # Add 9-by-5 buttons to the frame
        cb = '#1d1d1d'
        font_titol = font.Font(family="Helvetica", size=10, weight="bold")
        rows = len(self.rows)
        columns = 4
        self.labels = [[tk.Label() for j in range(columns)] for i in range(rows)]

        self.labels[0][0] = tk.Label(self.con_frame, text="CONCEPTE", font=font_titol, bg=cb, fg='white')
        self.labels[0][0].grid(row=0, column=0, sticky=tk.NSEW, ipadx=50, ipady=10)
        self.labels[0][1] = tk.Label(self.con_frame, text="DATA", font=font_titol, bg=cb, fg='white')
        self.labels[0][1].grid(row=0, column=1, sticky=tk.NSEW, ipadx=50, ipady=10)
        self.labels[0][2] = tk.Label(self.con_frame, text="IMPORT", font=font_titol, bg=cb, fg='white')
        self.labels[0][2].grid(row=0, column=2, sticky=tk.NSEW, ipadx=50, ipady=10)
        self.labels[0][3] = tk.Label(self.con_frame, text="CLASSIFICACIÓ", font=font_titol, bg=cb, fg='white')
        self.labels[0][3].grid(row=0, column=3, sticky=tk.NSEW, ipadx=50, ipady=10)

        for i in range(1, rows):
            # taula
            font_lab = font.Font(family="Helvetica", size=9)
            self.labels[i][0] = tk.Label(self.con_frame, text=str(self.ws_act.cell(row=self.rows[i], column=1).value)[:18],
                                    font=font_lab, bg=cb, fg='white')
            self.labels[i][0].grid(row=i, column=0, sticky=tk.NSEW, ipadx=70, ipady=10)
            self.labels[i][1] = tk.Label(self.con_frame, text=str(self.ws_act.cell(row=self.rows[i], column=2).value),
                                    font=font_lab, bg=cb, fg='white')
            self.labels[i][1].grid(row=i, column=1, sticky=tk.NSEW, ipadx=70, ipady=10)
            self.labels[i][2] = tk.Label(self.con_frame, text=str(self.ws_act.cell(row=self.rows[i], column=3).value),
                                    font=font_lab, bg=cb, fg='white')
            self.labels[i][2].grid(row=i, column=2, sticky=tk.NSEW, ipadx=70, ipady=10)

            llista_clas = list(diccionari.keys())
            variable = tk.StringVar(self.con_frame)
            variable.set("")
            self.opt = tk.OptionMenu(self.con_frame, variable, *llista_clas)
            self.opt.config(font=font_lab, bg=cb, fg='white', padx=50, pady=7, highlightthickness=0, width=1)
            self.opt.grid(row=i, column=3, sticky=tk.NSEW, ipadx=70, ipady=5)

        # Update buttons frames idle tasks to let tkinter calculate buttons sizes
        self.con_frame.update_idletasks()

        # Resize the canvas frame to show exactly 5-by-5 buttons and the scrollbar
        columns_width = sum([self.labels[0][j].winfo_width() for j in range(0, 3)])
        rows_height = sum([self.labels[i][0].winfo_height() for i in range(0, 10)])
        self.canvas_frame.config(width=columns_width + vsb.winfo_width(),
                            height=rows_height)

        # Set the canvas scrolling region
        canvas.config(scrollregion=canvas.bbox("all"))

    # A partir d'una llista d'excels els agrupa en un excel
    def combiexcel(self, llista):
        """Afegeix els excels del banc a un de sol"""
        # Carreguem l'excel de comptes
        ex_comptes = load_workbook('C:/Users/ferra/OneDrive/Tesla/Economia/EstatComptes.xlsx')
        for document in llista:
            # Carreguem l'excel del banc
            ex_caixa = load_workbook(document)
            sheet_caixa = ex_caixa['in']

            # Creació fulla segons el mes
            data_mes = str(sheet_caixa['B4'].value)[0:10].split('-')[1]
            self.nom_fulla = Diccionari.mes.get(data_mes)
            if self.nom_fulla not in ex_comptes.sheetnames:
                ws1 = ex_comptes.create_sheet(self.nom_fulla)
                ws1.title = self.nom_fulla
                ws2 = ex_comptes.active = ex_comptes[self.nom_fulla]

                # calculate total number of rows and
                # columns in source Excel file
                self.mr = sheet_caixa.max_row
                self.mc = sheet_caixa.max_column

                # copying the cell values from source
                # Excel file to destination Excel file
                for i in range(3, self.mr + 1):
                    for j in range(1, self.mc + 1):
                        if sheet_caixa.cell(row=i, column=j).value is not None:
                            # reading cell value from source excel file
                            c = sheet_caixa.cell(row=i, column=j)
                            # writing the read value to destination excel file
                            ws2.cell(row=i - 2, column=j).value = c.value

                for i in range(2, self.mr - 1):
                    # Format data
                    split_data = "/".join(list(reversed(str(ws2.cell(row=i, column=2).value)[0:10].split('-'))))
                    ws2.cell(row=i, column=2).value = split_data
                    # Format import
                    if ws2.cell(row=i, column=3).value is not None:
                        imports = ".".join((str(ws2.cell(row=i, column=3).value).split('.')))
                        ws2.cell(row=i, column=3).value = float(imports)
                    # Format saldo
                    if ws2.cell(row=i, column=4).value is not None:
                        saldo = ".".join("".join(str(ws2.cell(row=i, column=4).value)[0:-3].split('.')).split(','))
                        ws2.cell(row=i, column=4).value = float(saldo)

                # format titols columnes
                a1 = ws2['A1']
                b1 = ws2['B1']
                c1 = ws2['C1']
                d1 = ws2['D1']
                e1 = ws2['E1']
                e1.value = "Classificació"

                a1.font = Font(bold=True, size=15)
                b1.font = Font(bold=True, size=15)
                c1.font = Font(bold=True, size=15)
                d1.font = Font(bold=True, size=15)
                e1.font = Font(bold=True, size=15)

                self.check_concept(ex_comptes)

                # filtres
                maxrow = ws2.max_row
                ws2.auto_filter.ref = f"A1:E{maxrow}"
                ws2.auto_filter.add_filter_column(5, ["Menja", "Compres", "Transport"])
                ws2.auto_filter.add_sort_condition(f"B2:B{maxrow}")

                # Taula resum
                sheet1 = ex_comptes['Gener']
                for i in range(2, 20):
                    for j in range(7, 9):
                        ws2.cell(row=i, column=j).value = sheet1.cell(row=i, column=j).value
                # Cel·la estalvis
                ws2['H19'] = f'=D{maxrow}-D2'

            ex_comptes.save(filename='C:/Users/ferra/OneDrive/Tesla/Economia/EstatComptes.xlsx')

    def __init__(self, finestra1):
        """Inicialitza la segona finestra"""
        llista_exel = glob.glob(finestra1.carpeta + '/*.xlsx')
        self.combiexcel(llista_exel)


def main():
    mi_app = startWindow()


if __name__ == "__main__":
    main()
