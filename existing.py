from docx import Document
document = Document('template.docx')

table = document.add_table(rows=2, cols=2)


document.save('new_file.docx')
