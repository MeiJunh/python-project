from tkintertable import TableCanvas, TableModel
from tkinter import *


def tableFunc():
    tk = Tk()
    tk.geometry('800x500+200+100')
    tk.title('Test')
    f = Frame(tk)
    f.pack(fill=BOTH, expand=1)
    data = {}
    for i in range(0, 5):
        data[i] = {'col1': 99.88, 'col2': 108.79, 'label': 'rec1asdfasfasdfasdfasdfasdf'}
    table = TableCanvas(f, data=data)
    table.show()
    tk.mainloop()
    c = table.model.getRowCount()
    for i in range(0, c):
        print(table.model.getRecordAtRow(i))
    return


tableFunc()
