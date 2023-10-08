from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

def main():

    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("PDF to DOCX", "1º - Select the PDF file to be converted;\n\n2º - Select the destination path to save the DOCX file;\n\n3º - Then it's done, enjoy your DOCX file!")

    filetypes = [
            ("PDF files", "*.pdf")
        ]

    pdf_file = filedialog.askopenfilename(filetypes=filetypes,title='Select the PDF file to convert')

    if pdf_file:

        docx_output = filedialog.askdirectory(title='Select the DOCX destionation file') + f'/{os.path.splitext(os.path.basename(pdf_file))[0]}_converted.docx'

        if docx_output != f'/{os.path.splitext(os.path.basename(pdf_file))[0]}_converted.docx':

            try:

                cv = Converter(pdf_file)

                cv.convert(docx_output)

            except Exception as e:
                messagebox.showerror("Error", "Error to convert file!")
                print("Error: ",e)
                exit()

            messagebox.showinfo("Success", "PDF file converted!")
        
        else:
            messagebox.showwarning('Warning', 'No destination path selected, try again!')
            exit()

    else:
        messagebox.showwarning('Warning', 'No file selected, try again!')
        exit()

if __name__ == "__main__":
    main()