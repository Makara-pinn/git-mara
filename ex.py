import openpyxl
from docxtpl import DocxTemplate
filename= 'list.xlsx'
workbook = openpyxl.load_workbook(filename)
sheet = workbook.active
studentlist = list(sheet.values)
for student in studentlist:
    print(student)

# get template
template = DocxTemplate('word.docx')
for student in studentlist[1:]:
    template.render({
        'name':student[0],
        'year': student[1],
        
    })
    doc_name= str(student[0]) + '.docx'
    template.save(doc_name)