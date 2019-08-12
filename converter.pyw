#   v1.0.0  #
#   v1.1.0  #

from tkinter import *
from tkinter import colorchooser
from tkinter import filedialog
from tkinter.ttk import *
from tkinter import messagebox
import sys
import os
import comtypes.client
from PIL import Image
import time


window = Tk()
window.geometry("700x450")
window.title("Converter")
#for app favicon
window.iconbitmap("E:\coding practice\python practice\pyhon-test\converter\converter_logo.ico")

def message():
    messagebox.showinfo(title="Alert", message="Please select File.")

def fun():
    global file
    file = filedialog.askopenfile()
    string_var = ""

    if(file!=None):
        string_var = file.name
        f = file.name.split(".")[-1]
        print("f---",f)#############
        if(f=="doc" or f=="docx"):
            l2.config(text=string_var)
            convertBtn.config(command=docToPdf)
        elif(f=="jpg" or f=="jpeg" or f=="png" or f=="JPG" or f=="JPEG"):
            l2.config(text=string_var)
            convertBtn.config(command=imgToPdf)
        else:
            messagebox.showinfo(title="Alert", message="File should be in doc, docx, jpg, jpeg, JPG, JPEG, png format.")
    else:
        l2.config(text=string_var)
        convertBtn.config(command=message)

def bar():
    global l3
    l3 = Label(window, text="Status: Converting...", font="sans-serif")
    l3.grid(row=6, column=1, padx=15, pady=15)

    progress['value'] = 20
    window.update_idletasks()
    time.sleep(0.1)

    progress['value'] = 40
    window.update_idletasks()
    time.sleep(0.1)

    progress['value'] = 50
    window.update_idletasks()
    time.sleep(0.1)

    progress['value'] = 60
    window.update_idletasks()
    time.sleep(0.1)

    progress['value'] = 80
    window.update_idletasks()
    time.sleep(0.2)
    progress['value'] = 100


def docToPdf():
    global progress

    wdFormatPDF = 17
    in_file = os.path.abspath(file.name)
    out = filedialog.asksaveasfilename() + ".pdf"

    progress = Progressbar(window, orient=HORIZONTAL, length=100, mode='determinate')
    progress.grid(row=5, column=1, padx=20, pady=20, ipadx=20, ipady=2)

    bar()
    out_file = os.path.abspath(out)
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.open(in_file)
    doc.SaveAs(out_file,FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    progress.stop()
    progress.grid_remove()

    l3.configure(text="Status: Converted")


def imgToPdf():
    global progress

    out = filedialog.asksaveasfilename() + ".pdf"
    print("out----",out)###########
    if(out[0]!='.'):
        progress = Progressbar(window, orient=HORIZONTAL, length=100, mode='determinate')
        progress.grid(row=5, column=1, padx=20, pady=20, ipadx=20, ipady=2)

        bar()
        image = Image.open(file.name)
        print("mode----",image.mode)##############
        print(image.bits, image.size, image.format)###########
        if(image.mode=="RGBA"):
            image = image.convert("RGB")
        image.save(out,"PDF",resolution=100.0)

        '''
        pdf_bytes = img2pdf.convert(image.filename)
        pdf_file = open(out, "wb")
        pdf_file.write(pdf_bytes)
        image.close()
        pdf_file.close()'''

        progress.stop()
        progress.grid_remove()

        l3.configure(text="Status: Converted")
    else:
        pass


l1 = Label(window, text="Convert docx/ doc or JPG/ jpg/ jpeg/ png to Pdf", font=("sans-serif",12,"bold"))
l1.grid(row=0, column=1, padx=150, pady=20, ipadx=30, ipady=10)

l4 = Label(window, text="Developed by Amit Singh", font=("sans-serif", 7))
l4.grid(row=1, column=1, padx=35, pady=2, columnspan=3)

selectBtn = Button(window,text="Select File", command=fun)
selectBtn.grid(row=2, column=1, padx=10, pady=20, ipadx=30, ipady=10)

l2 = Label(window, text="", font="sans-serif")
l2.grid(row=3, column=1, padx=35, pady=10)

convertBtn = Button(window, text="Convert", command="")
convertBtn.grid(row=4, column=1, padx=10, pady=20, ipadx=30, ipady=10)

window.mainloop()
