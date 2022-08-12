# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
import tkinter.font as font


class Aplicacio:
    """Classe Aplicació"""

    # Variable per controlar les finestres
    finestra = 0

    # Variable de clase per usar en el càlcul de la posició de la finestra
    posx_y = 0

    def __init__(self):
        """Construeix una finestra de l'aplicació"""

        # Declara la finestra de l'aplicació
        # Treu la barra de tk
        self.arrel = tk.Tk()
        self.arrel.overrideredirect(True)

        # Defineix dimensions de la finestra ample x alt 300x200
        # que se situarà en la coordenada x=500,y=50
        # Centrem la finestra a la pantalla
        amplada_finestra = 600
        altura_finestra= 200
        amplada_monitor = self.arrel.winfo_screenwidth()
        altura_monitor = self.arrel.winfo_screenheight()
        print(f'Amplada monitor {self.arrel.winfo_screenwidth()}')
        print(f'Altura monitor {self.arrel.winfo_screenheight()}')
        x = round(amplada_monitor/2 - amplada_finestra/2)
        y = round(altura_monitor/2 - altura_finestra/2)

        self.arrel.geometry(f'{amplada_finestra}x{altura_finestra}+{x}+{y}')

        # Frames
        # Crea un frame per a la barra nova de títol
        back_ground = '#1d1d1d'
        title_barframe = tk.Frame(self.arrel, width=535, height=20,  bg=back_ground, relief='raised', bd=1, pady=3,
                                  highlightcolor=back_ground, highlightthickness=0)
        # crear frame per al boto tancar
        close_frame = tk.Frame(self.arrel, bg=back_ground, width=10, height=10, relief='raised', bd=1,
                               highlightcolor=back_ground, highlightthickness=0)

        # Crea un frame per a la sel·leció de la carpeta
        folder_frame = tk.Frame(self.arrel, bg=back_ground, width=555, height=200)

        # Configurar grid
        self.arrel.columnconfigure(0, weight=1)

        # Grid Frames
        title_barframe.grid(row=0, sticky=tk.EW)
        close_frame.grid(row=0, sticky=tk.NE)
        folder_frame.grid(row=1, sticky=tk.EW)

        # Widggets
        # Títol finestra
        title_name = tk.Label(title_barframe, text="Financial Control", bg=back_ground, fg='white')
        # Crea un boto per tancar a la barra de títol
        close_button = tk.Button(close_frame, text='x', command=self.arrel.destroy, bg=back_ground,
                                 activebackground="red", bd=0, font="bold", fg='white', activeforeground="white",
                                 highlightthickness=0)
        # Etiqueta carpeta
        folder_label = tk.Label(folder_frame, text="Sel·lecciona la carpeta d'excels",
                                bg=back_ground, fg='white', padx=15, pady=30)
        entry_folder = ttk.Entry(folder_frame, justify=tk.LEFT, width=50, background=back_ground)
        # Boto per buscar carpeta
        button_font = font.Font(family="Helvetica", size=8, weight="bold")
        search_button = tk.Button(folder_frame, text='Buscar carpeta', bg='#b5b5b5', activebackground="#ffffff", bd=0,
                                  font="bold", fg='black', activeforeground="black")
        search_button['font'] = button_font
        # Boto continuar
        continue_button = tk.Button(folder_frame, text='Continuar', bg='#b5b5b5', activebackground="#ffffff", bd=0,
                                    font="bold", fg='black', activeforeground="black")
        continue_button['font'] = button_font

        # Grid widgets
        title_name.grid(row=0, column=0, columnspan=7, sticky=tk.NS)
        close_button.grid(sticky=tk.NE)
        folder_label.grid(row=0, column=0, sticky=tk.W)
        entry_folder.grid(row=0, column=1, sticky=tk.W)
        search_button.grid(row=0, column=2, sticky=tk.E, padx=5)
        continue_button.grid(row=1, column=2, sticky=tk.SW, padx=5, pady=20)



        x_axis = None
        y_axis = None

        # Events amb bind per poder moure la finestra

        def move_window(event):
            self.arrel.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

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




def main():
    mi_app = Aplicacio()


if __name__ == "__main__":
    main()
