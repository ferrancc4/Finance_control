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
        x = round(amplada_monitor/2 - amplada_finestra/2)
        y = round(altura_monitor/2 - altura_finestra/2)

        self.arrel.geometry(f'{amplada_finestra}x{altura_finestra}+{x}+{y}')

        # Frames
        # Crea un frame per a la barra nova de títol
        back_ground = '#1d1d1d'
        title_barframe = tk.Frame(self.arrel, width=535, height=20,  bg=back_ground, relief='raised', bd=1, pady=3,
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
        entry_folder = ttk.Entry(folder_frame, textvariable=self.carpeta, justify=tk.LEFT, width=50, background=back_ground)
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
        self.folder_path = tk.filedialog.askdirectory(initialdir=r"C:\Users\ferra\OneDrive\Tesla\Economia", title="Sel·lecciona una carpeta")
        self.carpeta.set(self.folder_path)

    def next_window(self):
        """Tanca la primera finestra i obre la següent"""
        ## Implementar error de carpeta
        startWindow.carpeta = self.carpeta.get()
        self.arrel.destroy()
        secondWindow(startWindow)


class secondWindow(startWindow):
    """Segona finestra"""

    # Funcio iterar concepte per classificarlo

    def check_concept(self, diccionari, excel):
        """Classifica els conceptes segons el diccionari i si no estan classificats crea una interfície per
        classificarlos """
        ws_act = excel.active

        # iterem per cada una de les files
        for j in range(2, ws_act.max_row + 1):
            concept = str(ws_act.cell(row=j, column=1).value).lower()
            # iterem pel diccionari de conceptes
            for elem in diccionari.values():
                for res in elem:
                    if res in concept:
                        ws_act.cell(row=j, column=5).value = list(diccionari.keys())[
                            list(diccionari.values()).index(elem)]
        # Creem la segona pantalla
        self.sw = tk.Tk()
        self.sw.overrideredirect(True)

        amplada_finestra = 800
        altura_finestra = 700
        amplada_monitor = self.sw.winfo_screenwidth()
        altura_monitor = self.sw.winfo_screenheight()
        x = round(amplada_monitor / 2 - amplada_finestra / 2)
        y = round((altura_monitor - 50) / 2 - altura_finestra / 2)

        self.sw.geometry(f'{amplada_finestra}x{altura_finestra}+{x}+{y}')

        # Frames
        # Crea un frame per a la barra nova de títol
        back_ground = '#1d1d1d'
        title_barframe = tk.Frame(self.sw, width=amplada_finestra, height=20, bg=back_ground, relief='raised', bd=1,
                                  pady=3,
                                  highlightcolor=back_ground, highlightthickness=0)
        # crear frame per al botó tancar
        close_frame = tk.Frame(self.sw, bg=back_ground, width=10, height=10, relief='raised', bd=1,
                               highlightcolor=back_ground, highlightthickness=0)

        # Crea un frame gestió excel
        gestio_frame = tk.Frame(self.sw, bg='purple', width=amplada_finestra, height=100)

        # Crea un frame classificació conceptes
        concepte_frame = tk.Frame(self.sw, )

        # Configurar grid
        self.sw.columnconfigure(0, weight=1)

        # Grid Frames
        title_barframe.grid(row=0, sticky=tk.EW)
        close_frame.grid(row=0, sticky=tk.NE)
        gestio_frame.grid(row=1, sticky=tk.NSEW)

        # Widggets
        # Títol finestra
        title_name = tk.Label(title_barframe, text="Financial Control", bg=back_ground, fg='white')
        # Crea un botó per tancar a la barra de títol
        close_button = tk.Button(close_frame, text='x', command=self.sw.destroy, bg=back_ground,
                                 activebackground="red", bd=0, font="bold", fg='white',
                                 activeforeground="white",
                                 highlightthickness=0)
        # Labels gestió excel
        label_gestionant = tk.Label(gestio_frame, text=f'Gestionant despeses del mes', bg='blue',
                                    fg='white', font=font.Font(family="Helvetica", size=15, weight="bold"),
                                    padx=20, pady=15)
        label_mes = tk.Label(gestio_frame, text=self.nom_fulla,
                             font=font.Font(family="Helvetica", size=25, weight="bold"),
                             padx=20, pady=5, bg='green', fg='white')

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
        rows = []
        for j in range(2, ws_act.max_row + 1):
            if ws_act.cell(row=j, column=5).value is None:
                rows.append(j)

        # Labels titols columnes



        self.sw.mainloop()




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
                mr = sheet_caixa.max_row
                mc = sheet_caixa.max_column

                # copying the cell values from source
                # Excel file to destination Excel file
                for i in range(3, mr + 1):
                    for j in range(1, mc + 1):
                        if sheet_caixa.cell(row=i, column=j).value is not None:
                            # reading cell value from source excel file
                            c = sheet_caixa.cell(row=i, column=j)
                            # writing the read value to destination excel file
                            ws2.cell(row=i - 2, column=j).value = c.value

                for i in range(2, mr - 1):
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

                self.check_concept(Diccionari.classificació, ex_comptes)

                # filtres
                maxrow = ws2.max_row
                ws2.auto_filter.ref = f"A1:E{maxrow}"
                ws2.auto_filter.add_filter_column(5, ["Menja", "Compres", "Transport"])
                ws2.auto_filter.add_sort_condition(f"B2:B{maxrow}")

                #Taula resum
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
