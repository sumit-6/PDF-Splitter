import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import pandas as pd
from tkinter import filedialog
from tkinter import *
from pdf2image import convert_from_path

excel_path=""
pdf_path=""
output_path=""
pdf=""
exdata=""
name=""
number=""

PDFFILEOPENOPTIONS = dict(defaultextension=".pdf",filetypes=[('pdf file', '*.pdf')])

def columncheck(List):

    print("printing column names of your excel file")
    print()

    for item in List:
        print(item)
    print()
    a=input("Enter name of column on the basis of which you want to segregate:")
    if(a in List):
        return a
    else:
        print("One of the column name is incorrect")
        print("please enter again")
        var=columncheck(List)
        return var
    
def excel_path1():
    global excel_path
    excel_path=filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx"),("Excel files","*.csv"),("Excel files", "*.xls")])

def pdf_path1():
    global pdf_path
    pdf_path=filedialog.askopenfilename(**PDFFILEOPENOPTIONS)

def output_path1():
    global output_path
    output_path=filedialog.askdirectory()


def get_excel():

    while(excel_path == ""):
        root1=Tk()
        root1.geometry('900x200')
        
        excel_button=Button(root1,text="SELECT EXCEL FILE",bg="yellow", fg="blue",font=("Arial Bold", 15),command=excel_path1)
        excel_button.grid(row = 2, column = 0)
        exit_button=Button(root1,text="NEXT",bg="yellow", fg="blue",font=("Arial Bold", 15),command=root1.destroy)
        exit_button.grid(row = 2, column = 2)
        root1.mainloop()
        

def get_pdf():
    while(pdf_path==""):
        root2=Tk()
        root2.geometry('900x200')
        pdf_button=Button(root2,text="SELECT PDF FILE",bg="yellow", fg="blue",font=("Arial Bold", 15),command=pdf_path1)
        pdf_button.grid(row = 2, column = 0)
        exit_button=Button(root2,text="NEXT",bg="yellow", fg="blue",font=("Arial Bold", 15),command=root2.destroy)
        exit_button.grid(row = 2, column = 2)
        back_button=Button(root2,text="BACK",bg="yellow", fg="blue",font=("Arial Bold", 15),command=get_excel)
        back_button.grid(row = 2, column = 4)
        root2.mainloop()

def start_func():
    get_excel()
    get_pdf()
      
def pdf_convert():
    global pdf_path
    global exdata
    global pdf
    global exdata
    global name
    global number
    global output_path
    global excel_path
    pdf=PdfFileReader(pdf_path)
    for i in range(pdf.numPages):
        val=exdata[name][i]     
        pdfw=PdfFileWriter()
        pdfw.addPage(pdf.getPage(i))
        with open(os.path.join(output_path,'{0}.pdf'.format(val)), 'wb') as f:
            pdfw.write(f)
            f.close()

def jpg_convert():
    global exdata
    global pdf
    global name
    global number
    global output_path
    global pdf_path
    global excel_path
    images = convert_from_path(pdf_path)
    for i in range(len(images)):
        val=exdata[name][i]
        images[i].save(output_path+ '\\{0}.jpg'.format(val), 'JPEG')

def convertfile():
    global exdata
    global pdf
    global name
    global number
    global pdf_path
    global output_path
    exdata=pd.read_excel(excel_path)
    column_list=list(exdata.columns)
    column_list=columncheck(column_list)
    name=column_list
    while(output_path==""):
        root3=Tk()
        root3.geometry('900x200')
        output_button=Button(root3,text="SELECT OUTPUT FOLDER",bg="yellow", fg="blue",font=("Arial Bold", 15),command=output_path1)
        output_button.grid(row = 2, column = 0)
        exit_button=Button(root3,text="NEXT",bg="yellow", fg="blue",font=("Arial Bold", 15),command=root3.destroy)
        exit_button.grid(row = 2, column = 2)
        back_button=Button(root3,text="BACK",bg="yellow", fg="blue",font=("Arial Bold", 15),command=get_pdf)
        back_button.grid(row = 2, column = 4)
        root3.mainloop()
    
    root4=Tk()
    root4.geometry('900x200')
    PDF_button=Button(root4,text="CLICK TO CONVERT IN PDF",bg="yellow", fg="blue",font=("Arial Bold", 15),command=pdf_convert)
    PDF_button.grid(row = 2, column = 0)
    JPG_button=Button(root4,text="CLICK TO CONVERT IN JPG",bg="yellow", fg="blue",font=("Arial Bold", 15),command=jpg_convert)
    JPG_button.grid(row = 2, column = 2)
    exit_button=Button(root4,text="Exit",bg="yellow", fg="blue",font=("Arial Bold", 15),command=root4.destroy)
    exit_button.grid(row = 4, column = 0)
    root4.mainloop() 

if __name__=="__main__":
    start_func()
    convertfile()
