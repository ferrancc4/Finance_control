from tkinter import *
from functools import partial

root = Tk()
root.geometry('300x300')

menubar = Menu(root)
root.config(menu=menubar)

###the other function was removed

def get_optionMenu_selection(variable): ###this changed
    print(variable.get())

def upload_months():
    months = ["January", "Feb","Mar"]
    variable = StringVar(root)
    variable.set('Choose')

    w = OptionMenu(root, variable, *months)
    B = Button(root, text ="Send", command = partial(get_optionMenu_selection, variable)) ###this changed
    w.place(x=20, y=10)
    B.place(x= 105, y=11)


filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Upload months", command=upload_months)
filemenu.add_command(label="Open")
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)
root.mainloop()

print('he entrat')
        classe = self.vconcept.get()
        rows_taula = []
        opt_row = self.opt['row']
        rows_taula.append(opt_row)
        corelacio = zip(self.rows, rows_taula)
        print(corelacio)
        self.ws_act.cell(row=opt_row, column=5).value = classe
        self.ex_comptes.save(filename='C:/Users/ferra/OneDrive/Tesla/Economia/EstatComptes.xlsx')