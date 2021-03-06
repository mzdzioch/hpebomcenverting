import tkinter as tk
from tkinter import filedialog as fd

from exconv import run

root = tk.Tk()
root.title("ABB Power Grid HPE Converter")

canvas1 = tk.Canvas(root, width=300, height=300)
canvas1.pack()

def hello():
    filename = fd.askopenfilename()
    print(filename)
    if(run(filename)):
        label1 = tk.Label(root, text='Done', fg='green', font=('helvetica', 12, 'bold'))
    else:
        label1 = tk.Label(root, text='Please find correct file!', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 200, window=label1)


button1 = tk.Button(text='Find a file and convert', command=hello, bg='brown', fg='white')
canvas1.create_window(150, 150, window=button1)

root.mainloop()