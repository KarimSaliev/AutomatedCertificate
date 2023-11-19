import openpyxl
from docxtpl import DocxTemplate
import datetime


# Loading data from excel

path = "data.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.active

list_values = list(sheet.values)

doc = DocxTemplate("certificate.docx")

for value in list_values[1:]:
    doc.render({"name":value[0], "course_name":value[1], "date": value[3], "professor_name":value[2]})

    doc_name = "certificate" + value[0] + value[1] + ".docx"
    doc.save(doc_name)



