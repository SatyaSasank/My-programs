import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import style
from openpyxl import load_workbook
import sys
import os


root= tk.Tk()
root.title("AYTAS PLOT PROGRAM")

canvas1 = tk.Canvas(root, width = 600, height = 600, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    global df
    global filename

    import_file_path = filedialog.askopenfilename()
    df = pd.ExcelFile(import_file_path)
    filename = (os.path.basename(import_file_path))

browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('Calibri', 16, 'bold'))
canvas1.create_window(300, 200, window=browseButton_Excel)

def show_file ():
    canvas1.create_text(400, 250, text= filename, font=('calibri', 12, 'bold'))

list_Button = tk.Button(text=f'Files(s) Selected', command=show_file, bg='blue', fg='red', font=('Calibri', 12, 'bold'))
canvas1.create_window(300, 250, window=list_Button)

def plotgraph ():
    workbook = load_workbook(df)
    sheet = workbook.active
    style.use('ggplot')
    y = [sheet["C5"].value, sheet["C6"].value ,sheet["C7"].value ]
    x1 = [sheet["D5"].value, sheet["D6"].value, sheet["D7"].value]
    x2 = [sheet["E5"].value, sheet["E6"].value, sheet["E7"].value]
    plt.plot(y,x1,'b',label = sheet["D4"].value, linewidth = 5)
    plt.plot(y,x2,'r',label = sheet["E4"].value, linewidth = 5)
    plt.title('LIFT AND DRAG FOR VELOCITY')
    plt.ylabel(sheet["C4"].value)
    plt.xlabel('Forces(N)')
    plt.legend()
    plt.grid(True, color='k')
    plt.show()


plotButton_graph = tk.Button(text='Plot graph', command=plotgraph, bg='black', fg='cyan', font=('Calibri', 16, 'bold'))
canvas1.create_window(300, 300, window=plotButton_graph)

root.mainloop()
