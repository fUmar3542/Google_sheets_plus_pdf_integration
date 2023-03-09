import sys, subprocess
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'gspread'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'oauth2client'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'PyPDF2'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'fpdf'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'paddlepaddle'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'paddleocr'])
# subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'barcode'])
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter
from fpdf import FPDF
from paddleocr import PaddleOCR
import os
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import glob

os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
ocr = PaddleOCR(use_angle_cls=True, lang='en')


def get_Text(x):

    text =[]
    result = ocr.ocr(x, cls=True)  # Getting results from the paddleocr library functions
    try:
        for idx in range(len(result)):
            res = result[idx]
            for line in res:
                text.append(line[-1][0])
                # print(line[-1][0])  # Displaying the latest result on screen

        return text

    except Exception as o:
        print("***********************************")
        print(o)  # In case of exception occurrence


# ******************************************************
# Function creates a new pdf of attachments
# ******************************************************
def write_attachments(inputpdf, attchments_dir_path):

    paths = []
    for i in range(0, inputpdf.numPages):
        name = attchments_dir_path + "\\" + 'attachment_' + str(i+1) + ".pdf"
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        with open(name, "wb") as outputStream:
            output.write(outputStream)
        paths.append(name)

    return paths

# *************************************
def write_pdf(all_text, result):
    pdf = FPDF('P', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font('Arial', '', 12)

    for text in all_text:

        pdf.cell(45, 8, text[0], border=True, align='C')
        pdf.cell(82, 8, " " + text[1] + ":    " + text[2], border=True, align='L')
        # pdf.multi_cell(50, 5, text[1] + '\n' + text[2], border=True)
        pdf.cell(63, 8, text[3],  border=True, align='C', ln=1)

        pdf.cell(45, 8, text[4] + ":  " + text[5], border=True, align='C')
        pdf.cell(82, 8, " " + text[6] + ":  " + text[7], border=True, align='L')
        pdf.cell(63, 8, text[8] + ":  " + text[9], border=True, align='C', ln=1)
        pdf.ln(1)

        pdf.code39(text[10], x=pdf.get_x(), y=pdf.get_y(), w=4, h=20)
        pdf.ln(20)
        pdf.cell(190, 8, text[10], align='C', ln=1, border=True)

        pdf.cell(130, 8, "  " + text[11] + ":  " + text[12], border=True, align='L')
        pdf.cell(60, 8, "  " + text[15] + ":  " + text[16], border=True, align='L', ln=1)
        pdf.cell(190, 8, "  " + text[13] + ":  " + text[14], border=True, align='L', ln=1)
        pdf.cell(190, 8, "  " + text[17] + ":  " + text[18], border=True, align='L', ln=1)
        pdf.cell(95, 8, "  " + text[19] + ":  " + text[20], border=True, align='L')
        pdf.cell(95, 8, "  " + text[21], border=True, align='L')
        pdf.ln(25)

    pdf.output(result)

# ****************************************
def sheets_automation(text, sheet):

    orders = sheet.col_values(9)
    skus = sheet.col_values(7)
    products = sheet.col_values(6)
    # for i in range(len(orders)):
    #     print(orders[i])

    for x in text:
        sku = product = ""
        if orders.index(x[2]) >= 0:
            index = orders.index(x[2])
            sku = skus[index]
            product = products[index]
        x[16] = sku
        x[18] = product
        sheet.update_cell(index+1, 10, int(x[10]))

    return text

# *************************************************
def return_all_text(filename):

    attachments_dir_path = os.getcwd() + '\\Temporary'
    if not os.path.isdir(attachments_dir_path):
        os.makedirs(attachments_dir_path)  # Creating a temporary folder to save attachments

    inputpdf = PdfFileReader(open(filename, "rb"))

    paths = write_attachments(inputpdf, attachments_dir_path)

    d = 1
    data_to_write = []
    for x in paths:
        if d == 4:
            break
        print("---------------------------------------------")
        print(d)
        text = get_Text(x)
        sku = product = ""
        dt = [text[1], text[0], text[3], text[2], text[4], text[7], text[5], text[8], text[6], text[9], text[10],
              text[11], text[12], text[13], text[14], "SKU", sku, "Product", product, text[16], text[18], text[17]]
        data_to_write.append(dt)
        d = d + 1
    for t in paths:
        os.remove(t)
    os.rmdir(attachments_dir_path)  # Deleting temporary folder

    return data_to_write

# *********************************
def main():

    scopes = {
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    }
    secret_key_path = os.getcwd() + '/Data/secret_key.json'
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret_key_path, scopes=scopes)
    file = gspread.authorize(creds)
    workbook = file.open("Data entry book- SC v3")
    sheet = workbook.worksheet('Falabella GSC')

    def_paths = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.pdf')]

    for x in def_paths:
        if "merged" not in x:
            text = return_all_text(x)
            text = sheets_automation(text, sheet)
            resultant_file = "merged_" + x.replace(os.getcwd(),"")
            write_pdf(text, resultant_file)

# *********************************
main()






# import sys, subprocess
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'gspread'])
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'oauth2client'])
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'PyPDF2'])
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'fpdf'])
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'paddlepaddle'])
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'paddleocr'])
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'barcode'])
# # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'PyQt5'])
# from PyQt5.QtGui import QFont
# from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QMessageBox)
# from PyPDF2 import PdfFileReader, PdfFileWriter
# from fpdf import FPDF
# from paddleocr import PaddleOCR
# import os
# from oauth2client.service_account import ServiceAccountCredentials
# import gspread
# import glob
#
# os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
# ocr = PaddleOCR(use_angle_cls=True, lang='en')
#
# class Window(QMainWindow):
#     def __init__(self):
#         super().__init__()
#
#         self.title = "Google Sheets Automation"
#         self.top = 100
#         self.left = 100
#         self.width = 480
#         self.height = 300
#
#         self.pushButton = QPushButton("Click me to start the process", self)
#         self.pushButton.move(110, 70)
#         self.pushButton.resize(260,120)
#         self.pushButton.setFont(QFont('Arial', 13))
#
#         self.pushButton.clicked.connect(lambda: self.process())
#
#         self.main_window()
#
#     def main_window(self):
#         self.setWindowTitle(self.title)
#         self.setGeometry(self.top, self.left, self.width, self.height)
#         self.show()
#
#     def get_Text(self, x):
#
#         text = []
#         result = ocr.ocr(x, cls=True)  # Getting results from the paddleocr library functions
#         try:
#             for idx in range(len(result)):
#                 res = result[idx]
#                 for line in res:
#                     text.append(line[-1][0])
#                     # print(line[-1][0])  # Displaying the latest result on screen
#
#             return text
#
#         except Exception as o:
#             print("***********************************")
#             print(o)  # In case of exception occurrence
#
#     # ******************************************************
#     # Function creates a new pdf of attachments
#     # ******************************************************
#     def write_attachments(self, inputpdf, attchments_dir_path):
#         try:
#             paths = []
#             for i in range(0, inputpdf.numPages):
#                 name = attchments_dir_path + "\\" + 'attachment_' + str(i + 1) + ".pdf"
#                 output = PdfFileWriter()
#                 output.addPage(inputpdf.getPage(i))
#                 with open(name, "wb") as outputStream:
#                     output.write(outputStream)
#                 paths.append(name)
#
#             return paths
#         except Exception as ex:
#             print(ex)
#
#     # *************************************
#     def write_pdf(self, all_text, result):
#         try:
#             pdf = FPDF('P', 'mm', 'A4')
#             pdf.add_page()
#             pdf.set_font('Arial', '', 12)
#
#             for text in all_text:
#                 pdf.cell(45, 8, text[0], border=True, align='C')
#                 pdf.cell(82, 8, " " + text[1] + ":    " + text[2], border=True, align='L')
#                 # pdf.multi_cell(50, 5, text[1] + '\n' + text[2], border=True)
#                 pdf.cell(63, 8, text[3], border=True, align='C', ln=1)
#
#                 pdf.cell(45, 8, text[4] + ":  " + text[5], border=True, align='C')
#                 pdf.cell(82, 8, " " + text[6] + ":  " + text[7], border=True, align='L')
#                 pdf.cell(63, 8, text[8] + ":  " + text[9], border=True, align='C', ln=1)
#                 pdf.ln(1)
#
#                 pdf.code39(text[10], x=pdf.get_x(), y=pdf.get_y(), w=4, h=20)
#                 pdf.ln(20)
#                 pdf.cell(190, 8, text[10], align='C', ln=1, border=True)
#
#                 pdf.cell(130, 8, "  " + text[11] + ":  " + text[12], border=True, align='L')
#                 pdf.cell(60, 8, "  " + text[15] + ":  " + text[16], border=True, align='L', ln=1)
#                 pdf.cell(190, 8, "  " + text[13] + ":  " + text[14], border=True, align='L', ln=1)
#                 pdf.cell(190, 8, "  " + text[17] + ":  " + text[18], border=True, align='L', ln=1)
#                 pdf.cell(95, 8, "  " + text[19] + ":  " + text[20], border=True, align='L')
#                 pdf.cell(95, 8, "  " + text[21], border=True, align='L')
#                 pdf.ln(25)
#
#             pdf.output(result)
#         except Exception as x:
#             print(x)
#
#     # ****************************************
#     def sheets_automation(self, text, sheet):
#         try:
#             orders = sheet.col_values(9)
#             skus = sheet.col_values(7)
#             products = sheet.col_values(6)
#             # for i in range(len(orders)):
#             #     print(orders[i])
#
#             for x in text:
#                 sku = product = ""
#                 if orders.index(x[2]) >= 0:
#                     index = orders.index(x[2])
#                     sku = skus[index]
#                     product = products[index]
#                 x[16] = sku
#                 x[18] = product
#                 sheet.update_cell(index + 1, 10, int(x[10]))
#
#             return text
#         except Exception as ex:
#             print(ex)
#
#     # *************************************************
#     def return_all_text(self, filename):
#         try:
#             attachments_dir_path = os.getcwd() + '\\Temporary'
#             if not os.path.isdir(attachments_dir_path):
#                 os.makedirs(attachments_dir_path)  # Creating a temporary folder to save attachments
#
#             inputpdf = PdfFileReader(open(filename, "rb"))
#
#             paths = self.write_attachments(inputpdf, attachments_dir_path)
#
#             d = 1
#             data_to_write = []
#             for x in paths:
#                 # if d == 4:
#                 #     break
#                 print("---------------------------------------------")
#                 print(d)
#                 text = self.get_Text(x)
#                 sku = product = ""
#                 dt = [text[1], text[0], text[3], text[2], text[4], text[7], text[5], text[8], text[6], text[9], text[10],
#                       text[11], text[12], text[13], text[14], "SKU", sku, "Product", product, text[16], text[18], text[17]]
#                 data_to_write.append(dt)
#                 d = d + 1
#             for t in paths:
#                 os.remove(t)
#             os.rmdir(attachments_dir_path)  # Deleting temporary folder
#
#             return data_to_write
#         except Exception as ex:
#             print(ex)
#
#     def process(self):
#         dlg = QMessageBox()
#         dlg.setIcon(QMessageBox.Information)
#         dlg.setWindowTitle("Google Sheets Automation")
#         dlg.setFont(QFont('Calibri', 13))
#
#         dlg.setText("Your document is under processing. It might take few minutes. Do not exit...")
#         dlg.exec()
#
#         try:
#             filename = 'qAr_2XIM.pdf'
#             resultant_file = 'merged.pdf'
#             scopes = {
#                 'https://www.googleapis.com/auth/spreadsheets',
#                 'https://www.googleapis.com/auth/drive'
#             }
#             secret_key_path = os.getcwd() + '/Data/secret_key.json'
#             creds = ServiceAccountCredentials.from_json_keyfile_name(secret_key_path, scopes=scopes)
#             file = gspread.authorize(creds)
#             workbook = file.open("Data entry book- SC v3")
#             sheet = workbook.worksheet('Falabella GSC')
#
#             text = self.return_all_text(filename)
#             text = self.sheets_automation(text, sheet)
#             self.write_pdf(text, resultant_file)
#
#             dlg.setText("Success")
#         except Exception as ex:
#             print(ex)
#             dlg.setText("Error!!!\n" + str(ex))
#         finally:
#             dlg.exec()
#
# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#     window = Window()
#     sys.exit(app.exec())



