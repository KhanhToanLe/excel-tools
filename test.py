import tkinter as tk
from tkinter import ttk

root = tk.Tk()

tabControl = ttk.Notebook(root) 
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tabControl.add(tab1, text ='Change Data') 
tabControl.add(tab2, text ='Ctrl Home') 
tabControl.pack(expand = 1, fill ="both")


container = ttk.Frame(tab1)
canvas = tk.Canvas(container)
scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

canvas.configure(yscrollcommand=scrollbar.set)

for i in range(50):
    test = ttk.Frame(scrollable_frame)
    test.pack()
    ttk.Entry(test).pack(side="left")
    ttk.Entry(test).pack(side='left')



container.pack()
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

root.mainloop()