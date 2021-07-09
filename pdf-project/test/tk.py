import tkinter, os
from tkinter.messagebox import askokcancel, showinfo
# import TestDire
import windnd


def drag_files(urls):
    print(b'\n'.join(urls).decode())
    showinfo('确认路径', b'\n'.join(urls).decode())


if __name__ == '__main__':
    top = tkinter.Tk()
    top.title('测试工具')
    # top.geometry("350x200")
    ent = tkinter.Entry(top, width=100).grid(row=0)
    label = tkinter.Label(top, text='请拖拽文件夹至软件内上传', font=15, width=80, height=40)
    label.grid(row=1)

    # windnd.hook_dropfiles(top, func=drag_files)
    # 进入消息循环
    top.mainloop()
