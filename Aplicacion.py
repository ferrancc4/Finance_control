# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk


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

        # Defineix dimensions de la finestra 300x200
        # que se situarà en la coordenada x=500,y=50
        self.arrel.geometry('300x200+500+50')

        # Crea un frame per a la barra nova de títol
        back_ground = '#1d1d1d'
        title_bar = tk.Frame(self.arrel, bg=back_ground, relief='raised', bd=1, highlightcolor=back_ground, highlightthickness=0)

        # Crea un boto per tancar a la barra de títol
        close_button = tk.Button(title_bar, text='x', command=self.arrel.destroy, bg=back_ground, padx=5, pady=2, activebackground="red", bd=0, font="bold", fg='white', activeforeground="white", highlightthickness=0)

        # Títol finestra
        title_window = "Financial Control"
        title_name = tk.Label(title_bar, text=title_window, bg=back_ground, fg='white')

        # Canvas per ficar tots els elements
        window = tk.Canvas(self.arrel, bg='#696969', highlightthickness=0)

        # Pack dels widgets
        title_bar.pack(expand=1, fill=tk.X)
        title_name.pack(side=tk.LEFT)
        close_button.pack(side=tk.RIGHT)
        window.pack(expand=1, fill=tk.BOTH)
        x_axis = None
        y_axis = None

        # Events amb bind per poder moure la finestra

        def move_window(event):
            self.arrel.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

        def change_on_hovering(event):
            close_button.configure(bg='red')

        def return_to_normal_state(event):
            close_button.configure(bg=back_ground)

        title_bar.bind('<B1-Motion>', move_window)
        close_button.bind('<Enter>', change_on_hovering)
        close_button.bind('<Leave>', return_to_normal_state)

        self.arrel.mainloop()




def main():
    mi_app = Aplicacio()


if __name__ == "__main__":
    main()
