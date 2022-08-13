import tkinter as tk
from tkinter import font

root = tk.Tk()
root.grid_rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

frame_main = tk.Frame(root, bg="gray", height=600, width=500)
frame_main.grid(sticky='news')

label1 = tk.Label(frame_main, text="Label 1", fg="green")
label1.grid(row=0, column=0, pady=(5, 0), sticky='nw')

label2 = tk.Label(frame_main, text="Label 2", fg="blue")
label2.grid(row=1, column=0, pady=(5, 0), sticky='nw')

label3 = tk.Label(frame_main, text="Label 3", fg="red")
label3.grid(row=3, column=0, pady=5, sticky='nw')

# Create a frame for the canvas with non-zero row&column weights
frame_canvas = tk.Frame(frame_main)
frame_canvas.grid(row=2, column=0, pady=(5, 0), sticky='nw')
frame_canvas.grid_rowconfigure(0, weight=1)
frame_canvas.grid_columnconfigure(0, weight=1)
# Set grid_propagate to False to allow 5-by-5 buttons resizing later
frame_canvas.grid_propagate(False)

# Add a canvas in that frame
canvas = tk.Canvas(frame_canvas, bg="yellow")
canvas.grid(row=0, column=0, sticky="news")

# Link a scrollbar to the canvas
vsb = tk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
vsb.grid(row=0, column=1, sticky='ns')
canvas.configure(yscrollcommand=vsb.set)

# Create a frame to contain the buttons
frame_buttons = tk.Frame(canvas, bg="blue")
canvas.create_window((0, 0), window=frame_buttons, anchor='nw')

# Add 9-by-5 buttons to the frame
cb = '#1d1d1d'
font_titol = font.Font(family="Helvetica", size=10, weight="bold")
rows = 9
columns = 4
labels = [[tk.Button() for j in range(columns)] for i in range(rows)]
for i in range(0, rows):
    labels[0][0] = tk.Label(frame_buttons, text="CONCEPTE", font=font_titol, bg=cb, fg='white')
    labels[0][0].grid(row=0, column=0, sticky=tk.NSEW, ipadx=50, ipady=10)
    labels[0][1] = tk.Label(frame_buttons, text="DATA", font=font_titol, bg=cb, fg='white')
    labels[0][1].grid(row=0, column=1, sticky=tk.NSEW, ipadx=50, ipady=10)
    labels[0][2] = tk.Label(frame_buttons, text="IMPORT", font=font_titol, bg=cb, fg='white')
    labels[0][2].grid(row=0, column=2, sticky=tk.NSEW, ipadx=50, ipady=10)
    labels[0][3] = tk.Label(frame_buttons, text="CLASSIFICACIÓ", font=font_titol, bg=cb, fg='white')
    labels[0][3].grid(row=0, column=3, sticky=tk.NSEW, ipadx=50, ipady=10)

# Update buttons frames idle tasks to let tkinter calculate buttons sizes
frame_buttons.update_idletasks()

# Resize the canvas frame to show exactly 5-by-5 buttons and the scrollbar
first5columns_width = sum([labels[0][j].winfo_width() for j in range(0, 4)])
first5rows_height = sum([labels[i][0].winfo_height() for i in range(0, 4)])
frame_canvas.config(width=first5columns_width + vsb.winfo_width(),
                    height=first5rows_height)

# Set the canvas scrolling region
canvas.config(scrollregion=canvas.bbox("all"))

# Launch the GUI
root.mainloop()