# -*- coding: utf-8 -*-
import glob
import tkinter as tk
from tkinter import ttk, filedialog
import tkinter.font as font
from openpyxl import load_workbook, styles
from openpyxl.styles import Font
from openpyxl.formatting.rule import CellIsRule
from copy import copy
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
            """Dotar de moviment a la finestra"""
            self.arrel.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

        # Els botons canvien de color al passar per damunt
        def change_on_hovering(event):
            """El botó canvia de color al passar per sobre"""
            close_button.configure(bg='red')

        def change_search_on_hovering(event):
            """El botó canvia de color al passar per sobre"""
            search_button.configure(bg='black', fg='white')

        def change_continue_on_hovering(event):
            """El botó canvia de color al passar per sobre"""
            continue_button.configure(bg='black', fg='white')

        def return_to_normal_state(event):
            """El botó torna al seu estat inicial"""
            close_button.configure(bg=back_ground)

        def return_search_to_normal_state(event):
            """El botó torna al seu estat inicial"""
            search_button.configure(bg='#b5b5b5', fg='black')

        def return_continue_to_normal_state(event):
            """El botó torna al seu estat inicial"""
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

    def select_item(self, seleccio):
        """Assigna la classificació a l'element"""
        rows = len(self.rows)
        zip_list = list(zip(self.rows, self.rows_taula))
        for i in range(0, rows):
            if seleccio == self.vconcept[i][3].get():
                self.ws_act.cell(row=zip_list[i][0], column=5).value = seleccio
        self.ex_comptes.save(filename='C:/Users/ferra/OneDrive/Tesla/Economia/EstatComptes.xlsx')
    def taula(self, diccionari):
        """Crea la taula dels conceptes per classificar"""
        # creem un canvas per allotjar la scrollbar
        # Creem un canvas al frame canvas
        canvas = tk.Canvas(self.canvas_frame, bg='#1d1d1d')
        canvas.grid(row=0, column=0, sticky=tk.NSEW)
        # fiquem el scrolbar al frame canvas
        vsb = tk.Scrollbar(self.canvas_frame, orient="vertical", command=canvas.yview)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        canvas.configure(yscrollcommand=vsb.set)
        # Crea un frame per la taula de conceptes
        self.con_frame = tk.Frame(canvas, bg='#1d1d1d')
        canvas.create_window((0, 0), window=self.con_frame, anchor=tk.NW, width=900)
        # Afegim els títols de la taula i les files corresponents a conceptes sense classificar
        cb = '#1d1d1d'
        font_titol = font.Font(family="Helvetica", size=10, weight="bold")
        rows = len(self.rows)

        self.rows_taula = []
        for i in range(1, rows + 1):
            self.rows_taula.append(i)
        columns = 4
        # Per poder crear variables de la taula
        self.labels = [[tk.Label() for j in range(columns)] for i in range(rows)]
        llista_clas = list(diccionari.keys())
        self.vconcept = [[tk.StringVar(self.con_frame) for j in range(columns)] for i in range(rows)]
        self.opt = [[tk.OptionMenu(self.con_frame, self.vconcept[i][3], *llista_clas) for j in range(columns)] for i in
                    range(rows)]

        self.labels[0][0] = tk.Label(self.con_frame, text="CONCEPTE", font=font_titol, bg=cb, fg='white')
        self.labels[0][0].grid(row=0, column=0, ipadx=50, ipady=10)
        self.labels[0][1] = tk.Label(self.con_frame, text="DATA", font=font_titol, bg=cb, fg='white')
        self.labels[0][1].grid(row=0, column=1, ipadx=50, ipady=10)
        self.labels[0][2] = tk.Label(self.con_frame, text="IMPORT", font=font_titol, bg=cb, fg='white')
        self.labels[0][2].grid(row=0, column=2, ipadx=50, ipady=10)
        self.labels[0][3] = tk.Label(self.con_frame, text="CLASSIFICACIÓ", font=font_titol, bg=cb, fg='white')
        self.labels[0][3].grid(row=0, column=3, ipadx=50, ipady=10)

        self.vclist = []
        for i in range(0, rows):
            # taula
            font_lab = font.Font(family="Helvetica", size=9)
            self.labels[i][0] = tk.Label(self.con_frame, text=str(self.ws_act.cell(row=self.rows[i], column=1).value),
                                         font=font_lab, bg=cb, fg='white')
            self.labels[i][0].grid(row=i, column=0, ipadx=70, ipady=10)
            self.labels[i][1] = tk.Label(self.con_frame, text=str(self.ws_act.cell(row=self.rows[i], column=2).value),
                                         font=font_lab, bg=cb, fg='white')
            self.labels[i][1].grid(row=i, column=1, ipadx=70, ipady=10)
            self.labels[i][2] = tk.Label(self.con_frame, text=str(self.ws_act.cell(row=self.rows[i], column=3).value),
                                         font=font_lab, bg=cb, fg='white')
            self.labels[i][2].grid(row=i, column=2, ipadx=70, ipady=10)
            self.vconcept[i][3] = tk.StringVar(self.con_frame)
            self.vconcept[i][3].set(f'SELECT - {i}')
            self.opt[i][3] = tk.OptionMenu(self.con_frame, self.vconcept[i][3], *llista_clas, command=self.select_item)
            self.opt[i][3].config(font=font_lab, bg=cb, fg="white", highlightthickness=0, width=1)
            self.opt[i][3].grid(row=i, column=3, sticky=tk.EW, ipadx=70, ipady=5)

        # Update buttons frames idle tasks to let tkinter calculate buttons sizes
        self.con_frame.update_idletasks()

        # Recalcul del canvas per que mostri totes les files de conceptes
        nfiles = 0
        if len(self.rows) >= 12:
            nfiles = 12
        else:
            nfiles = len(self.rows)
        columns_width = sum([self.labels[0][j].winfo_width() for j in range(0, 4)])
        rows_height = sum([self.labels[i][0].winfo_height() for i in range(0, nfiles)])
        extra_amplada = 0
        extra_altura = 0
        if columns_width < self.amplada_finestra:
            extra_amplada = self.amplada_finestra - columns_width - 15
            columns_width = columns_width + extra_amplada
        if rows_height < self.altura_finestra - 100:
            extra_altura = self.altura_finestra - rows_height - 185
            rows_height = rows_height + extra_altura
        self.canvas_frame.config(width=columns_width + vsb.winfo_width(), height=rows_height)
        # Set the canvas scrolling region
        canvas.config(scrollregion=canvas.bbox("all"))

    def check_concept(self, excel):
        """Classifica els conceptes segons el diccionari i si no estan classificats crea una interfície per
        classificarlos """
        self.ws_act = excel.active

        # iterem per cada una de les files de l'excel i assignem un concepte
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

        self.amplada_finestra = 900
        self.altura_finestra = 700
        amplada_monitor = self.sw.winfo_screenwidth()
        altura_monitor = self.sw.winfo_screenheight()
        x = round(amplada_monitor / 2 - self.amplada_finestra / 2)
        y = round((altura_monitor - 50) / 2 - self.altura_finestra / 2)

        self.sw.geometry(f'{self.amplada_finestra}x{self.altura_finestra}+{x}+{y}')

        # Frames
        # Creem un frame general
        self.frame_main = tk.Frame(self.sw, bg="gray", width=self.amplada_finestra, height=self.altura_finestra)
        self.frame_main.grid(sticky=tk.NSEW)
        # Crea un frame per a la barra nova de títol
        back_ground = '#1d1d1d'
        title_barframe = tk.Frame(self.frame_main, width=self.amplada_finestra, height=20, bg=back_ground,
                                  relief='raised', bd=1,
                                  pady=3, highlightcolor=back_ground, highlightthickness=0)
        # crear frame per al botó tancar
        close_frame = tk.Frame(self.frame_main, bg=back_ground, width=10, height=10, relief='raised', bd=1,
                               highlightcolor=back_ground, highlightthickness=0)
        # Crea un frame gestió excel
        gestio_frame = tk.Frame(self.frame_main, bg=back_ground)
        # Creem un frame per al canvas que allotjarà la taula
        self.canvas_frame = tk.Frame(self.frame_main, bg=back_ground)
        # Crear frame botons part baixa finestra
        self.fbuttons = tk.Frame(self.frame_main, bg=back_ground)

        # Grid Frames
        title_barframe.grid(row=0, sticky=tk.EW)
        close_frame.grid(row=0, sticky=tk.NE)
        gestio_frame.grid(row=1, sticky=tk.EW)
        self.canvas_frame.grid(row=2, column=0, sticky=tk.EW)
        self.canvas_frame.rowconfigure(0, weight=1)
        self.canvas_frame.columnconfigure(0, weight=1)
        self.canvas_frame.grid_propagate(False)
        self.fbuttons.grid(row=3, sticky=tk.EW)
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
        # Barra inferior
        nouconcepte = tk.Button(self.fbuttons, text="Nova classe i/o concepte", bg=back_ground, fg='white')
        tancar_finestra = tk.Button(self.fbuttons, text="Tanca", bg=back_ground, fg='white', command=self.sw.destroy)
        ## ------- Nou concepte
        # Grid widgets
        title_name.grid(row=0, column=0, sticky=tk.NS)
        close_button.grid(sticky=tk.NE)
        label_gestionant.pack(fill=tk.X)
        label_mes.pack(fill=tk.X)
        nouconcepte.pack(side='left')
        tancar_finestra.pack(side='right')

        def move_window(event):
            """Dotar de moviment a la finestra"""
            self.sw.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

        def change_on_hovering(event):
            """El botó canvia de color al passar per sobre"""
            close_button.configure(bg='red')

        def return_to_normal_state(event):
            """El botó torna al seu estat inicial"""
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

    def combiexcel(self, llista):
        """Afegeix els excels del banc a un de sol"""
        # Carreguem l'excel de comptes
        self.ex_comptes = load_workbook('C:/Users/ferra/OneDrive/Tesla/Economia/EstatComptes.xlsx')
        for document in llista:
            # Carreguem l'excel del banc
            ex_caixa = load_workbook(document)
            sheet_caixa = ex_caixa['in']

            # Creació fulla segons el mes
            data_mes = str(sheet_caixa['B4'].value)[0:10].split('-')[1]
            self.nom_fulla = Diccionari.mes.get(data_mes)
            if self.nom_fulla not in self.ex_comptes.sheetnames:
                ws1 = self.ex_comptes.create_sheet(self.nom_fulla)
                ws1.title = self.nom_fulla
                ws2 = self.ex_comptes.active = self.ex_comptes[self.nom_fulla]

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

                #Funcio classificacio conceptes
                self.check_concept(self.ex_comptes)

                # filtres
                maxrow = ws2.max_row
                ws2.auto_filter.ref = f"A1:E{maxrow}"
                ws2.auto_filter.add_filter_column(5, ["Menja", "Compres", "Transport"])
                ws2.auto_filter.add_sort_condition(f"B2:B{maxrow}")

                # Taula resum
                sheet1 = self.ex_comptes['Gener']
                # Crear la taula a partir del diccionari
                ws2['G2'] = "TAULA RESUM"
                ws2['G3'] = "Classificació"
                ws2['H3'] = "€"
                key_list = list(Diccionari.classificació.keys())
                num_files = len(key_list)
                ws2[f'G{num_files + 4}'] = "Estalvis"
                for i in range(4, num_files + 4):
                    ws2.cell(row=i, column=7).value = key_list[i - 4]
                    ws2.cell[f'H{i}'] = f'=SUMAR.SI(E2:E{maxrow};G{i};C2:C{maxrow})'
                # Fica la mateixa amplada de columna
                for idx, rd in sheet1.column_dimensions.items():
                    ws2.column_dimensions[idx] = copy(rd)
                # Mateix format
                for (row, col), source_cell in sheet1._cells.items():
                    cell = ws2.cell(column=col, row=row)
                    cell.font = copy(source_cell.font)
                    cell.fill = copy(source_cell.fill)
                    cell.border = copy(source_cell.border)
                    cell.number_format = copy(source_cell.number_format)
                    cell.alignment = copy(source_cell.alignment)
                # Format condicional
                red_font = styles.Font(size=11, color='9c0006')
                redFill = styles.PatternFill(bgColor='ffc7ce', fill_type='solid')
                greenFill = styles.PatternFill(bgColor='c6efce', fill_type='solid')
                green_font = styles.Font(size=11, color='006100')
                ws2.conditional_formatting.add('H4:H18', CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True,
                                                                    fill=redFill, font=red_font))
                ws2.conditional_formatting.add('H4:H19', CellIsRule(operator='greaterThan', formula=['0'], stopIfTrue=True,
                                                                    fill=greenFill, font=green_font))
                ws2.merge_cells('G2:H2')
                # Cel·la estalvis
                ws2[f'H{num_files + 4}'] = f'=D{maxrow}-D2'

            self.ex_comptes.save(filename='C:/Users/ferra/OneDrive/Tesla/Economia/EstatComptes.xlsx')

    def __init__(self, finestra1):
        """Inicialitza la segona finestra"""
        llista_exel = glob.glob(finestra1.carpeta + '/*.xlsx')
        self.combiexcel(llista_exel)


def main():
    mi_app = startWindow()


if __name__ == "__main__":
    main()
