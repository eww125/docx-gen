from docx import Document
from docx.enum.style import WD_STYLE_TYPE
document = Document('template.docx')


table = document.add_table(rows=1, cols=2) #, style="Colorful Grid")
row = table.rows[0]
row.cells[0].text = 'info1'
row.cells[1].text = 'info2'

table.style = 'LightShading-Accent1'
styles = document.styles
table_styles = [
        s for s in styles if s.type == WD_STYLE_TYPE.TABLE
]
for style in table_styles:
        print(style.name)

document.save('new_file.docx')
