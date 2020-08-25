import comtypes.client, os, sys
from tkinter.filedialog import askdirectory
import tkinter as tk

root = tk.Tk()
root.withdraw()

def convert_to_pdf(_in, _out):
    pdf_format_key = 17
    file_in = os.path.abspath(_in)
    file_out = os.path.abspath(_out)
    worddoc = comtypes.client.CreateObject('word.Application')
    doc = worddoc.Documents.Open(file_in)
    doc.SaveAs(file_out, FileFormat = pdf_format_key)
    doc.Close()
    worddoc.Quit()

#destination = sys.argv[1]
destination = askdirectory(title='Select Folder')
#destination = "C:/Users/KellyBennetts/Desktop/tr"
for file in os.listdir(destination):
    if ".docx" in file:
        convert_to_pdf(destination + "\\" + file, destination + "\\" + file.replace(".docx", "") + ".pdf")
