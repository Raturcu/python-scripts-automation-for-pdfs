import tkinter as tk
from tkinter import filedialog
import tkinter.font as font
import easygui
from docx2pdf import convert
from PyPDF2 import PdfFileReader,PdfFileWriter,PdfReader
import fitz
import os
from os import DirEntry, curdir, getcwd, chdir, rename
from glob import glob as glob


root = tk.Tk();
root.title("Docx To PDF")

def add_file():
    file_path=easygui.fileopenbox(filetypes=["*.docx"])
    convert(file_path)

def split_file():
    pdf_file_path=easygui.fileopenbox(filetypes=["*.pdf"])
    file_base_name=pdf_file_path.replace('.pdf','')
    output_folder_path=easygui.filesavebox(filetypes=["*.pdf"])

    pdf=PdfFileReader(pdf_file_path)

    for page_num in range(pdf.numPages):
        pdfWriter=PdfFileWriter()
        pdfWriter.add_page(pdf.getPage(page_num))

        with open(os.path.join(output_folder_path,'{0}_Page_{1}.pdf'.format(file_base_name,page_num+1)), 'wb') as f:
            pdfWriter.write(f)
            f.close()

def rename_file():
    directory =easygui.diropenbox()
    chdir(directory)

    pdf_list=glob('*.pdf')
    pdf_list_actualized = pdf_list
    for pdf in pdf_list:
        with fitz.open(pdf) as pdf_obj:
            text=pdf_obj[0].get_text()
        new_file_name=text.split("\n",1)[0].strip()
        nr_files_same = 1
        f_new_pdf = new_file_name+ '_RBC'+str(nr_files_same) + '.pdf'
        while f_new_pdf in pdf_list_actualized:
            nr_files_same+=1
            f_new_pdf = new_file_name+ '_'+str(nr_files_same) + '.pdf'
        else:
            rename(pdf,f_new_pdf)
        pdf_list_actualized = glob('*.pdf')        

canvas=tk.Canvas(root, height=500,width=700)
canvas.pack()

frame=tk.Frame(root,bg="#263D42")
frame.place(relwidth=.8,relheight=.8,relx=.1,rely=.1)

button=tk.Button(frame,text="Convert DOCX to PDF",padx=500,pady=50,bg="#ff2200",fg="#fff",command=add_file)
button_font=font.Font(size=20)
button["font"]=button_font
button.pack()

split_button=tk.Button(frame,text="Split Pages",padx=500,pady=50,bg="#ff2200",fg="#fff",command=split_file)
split_button_font=font.Font(size=20)
split_button["font"]=split_button_font
split_button.pack()

rename_button=tk.Button(frame,text="Rename PDF Files",padx=500,pady=50,bg="#ff2200",fg="#fff",command=rename_file)
rename_button_font=font.Font(size=20)
rename_button["font"]=rename_button_font
rename_button.pack();


root.mainloop()