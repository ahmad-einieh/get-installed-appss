from tkinter import *
from tkinter import filedialog
import xlsxwriter
from windows_tools.installed_software import get_installed_software
from xlsxwriter.workbook import Workbook
import webbrowser


def CreateApps(dir):
    row = 0
    try:
        workbook = xlsxwriter.Workbook(dir+'\\apps.xlsx')
        worksheet = workbook.add_worksheet()
        for software in get_installed_software():
            worksheet.write(row, 0, software['name'])
            worksheet.write(row, 1, software['version'])
            worksheet.write(row, 2, software['publisher'])
            row = row + 1
        workbook.close()
        message.set("done !")
        print("done !")
    except:
        print("error")
        message.set("error")


def saveFile():
    file = filedialog.askdirectory()
    print(file)
    CreateApps(file)


def mypaypal():
    webbrowser.open_new(r"https://paypal.me/progahmad?locale.x=ar_EG")


def mygithub():
    webbrowser.open_new(r"https://github.com/ahmad-prog")


window = Tk(className="Your Device Apps")
window.geometry("300x300")
window.configure(bg='#6C6C6C')
message = StringVar()
window.resizable(0, 0)

Label(bg='#6C6C6C').pack()
Label(bg='#6C6C6C').pack()
Label(bg='#6C6C6C').pack()
Button(text="get my apps", command=saveFile,
       bg='#0770A8', fg='yellow', width=22, height=2).pack()
Label(bg='#6C6C6C').pack()
Button(text="support", command=mypaypal,
       bg='#0770A8', fg='yellow', width=22, height=2).pack()
Label(bg='#6C6C6C').pack()
Button(text="github", command=mygithub,
       bg='#0770A8', fg='yellow', width=22, height=2).pack()
Label(bg='#6C6C6C').pack()
Label(bg='#6C6C6C', textvariable=message, fg='yellow').pack()

window.mainloop()
