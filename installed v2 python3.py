import tkinter as tk
from tkinter import filedialog
import openpyxl
import windows_tools.installed_software



def create_apps_file(dir):
    wb = openpyxl.Workbook()
    ws = wb.active
    for software in windows_tools.installed_software.get_installed_software():
        ws.append([software['name'], software['version'], software['publisher']])
    wb.save(dir + '\\apps.xlsx')



def save_file():
    dir = filedialog.askdirectory()
    create_apps_file(dir)


def main():
    root = tk.Tk()
    root.geometry('300x300')

    btn_save = tk.Button(root, text='Save', command=save_file)
    btn_save.pack()

    root.mainloop()


if __name__ == '__main__':
    main()
