from tkinter import *
from tkinter import ttk, filedialog

from marksheet import Marksheet

if __name__ == '__main__':
    root = Tk()
    root.title('Marksheet')
    root.geometry('500x200+500+200')
    root.resizable(0, 0)
    root.configure(bg='#ffffff')
    root.option_add('*foreground', 'black')
    root.option_add('*activeForeground', 'red')

    style = ttk.Style(root)
    style.configure('TLabel', foreground='white')

    file_path = StringVar()
    Entry(
        root,
        textvariable=file_path,
        font=('hack', 16),
        bg='#ffffff',
        fg='#000000'
    ).pack(fill=X, padx=10, pady=15)
    file_path.set('Enter File Location')


    def find_file_path():
        file_path.set(filedialog.askdirectory(initialdir='/home', title='Select directory'))


    Button(
        root,
        text='Browse',
        command=find_file_path,
        font=('hack', 12),
        bg='#ffffff',
        fg='#000000'
    ).pack(fill=X, padx=200)

    class_name = StringVar()
    Entry(
        root,
        textvariable=class_name,
        font=('hack', 16),
        bg='#ffffff',
        fg='#000000'
    ).pack(fill=X, padx=10, pady=15)
    class_name.set('Enter Class Name')


    def pressed_ok():
        root.destroy()
        Marksheet(file_path.get(), class_name.get())


    Button(
        root,
        text='Ok',
        command=pressed_ok,
        font=('hack', 12),
        bg='#ffffff',
        fg='#000000'
    ).pack(fill=X, padx=200)

    root.mainloop()
