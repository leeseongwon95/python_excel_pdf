from tkinter import *
from tkinter import filedialog
import os
import win32com.client

window = Tk()
window.title("엑셀 PDF 자동변환")
window.geometry("440x300")
window.resizable(False, False)
   
def open_excel_to_save_pdf():
    filename =  filedialog.askopenfilename( title="select excel", filetypes=[("Excel files", ".xlsx .xls")])
    print(filename)
    excel = win32com.client.Dispatch("Excel.Application")
    try:
        wb = excel.Workbooks.Open(filename)
        ws_chart = wb.Worksheets("Sheet1")
        ws_chart.Select()    
        s = os.path.splitext(filename)
        pdf_path = s[0]
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
    except:
        FileNotFoundError

label = Label(window, text="PDF 변환 키를 누르고 원하는 파일을 고르세용~" , height = 10)
label.pack()

btnRead = Button(window, height =5 , width = 10, text = "PDF 변환", command=open_excel_to_save_pdf)
btnRead.pack()

window.mainloop()