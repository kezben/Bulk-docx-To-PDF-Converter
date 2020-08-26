import comtypes.client, os
from tkinter.filedialog import askdirectory
import tkinter as tk

# hide root window to not show on loading
root = tk.Tk()
root.withdraw()

# function to convert
def convert_to_pdf(_fin, _fout):
    pdf_format_key = 17
    file_in = os.path.abspath(_fin)
    file_out = os.path.abspath(_fout)
    doc = comtypes.client.CreateObject('word.Application')
    open = doc.Documents.Open(file_in)
    open.SaveAs(file_out, FileFormat = pdf_format_key)
    open.Close()
    doc.Quit()

# choose folder to convert and call convert function
save = askdirectory(title='Select Folder')
for file in os.listdir(save):
    if ".docx" in file:
        convert_to_pdf(save + "\\" + file, save + "\\" + file.replace(".docx", "") + ".pdf")
