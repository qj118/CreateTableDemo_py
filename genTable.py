from docx import Document
from docx.shared import Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
import datetime

year = 2018

s = datetime.datetime(year, 1, 1)
year_end = datetime.datetime(year, 12, 31)
weeks = year_end.strftime("%W")

delta = datetime.timedelta(days = 6 - s.weekday())

document = Document()
# styles = document.styles
# for s in styles:
# 	if s.type == WD_STYLE_TYPE.TABLE:
# 			document.add_paragraph("Table style is : " + s.name)
# 			document.add_table(3, 3, style = s)
# 			document.add_paragraph("\n")

# if (year % 4 == 0) & (year % 100 != 0):
# 	year_long = 366
# elif year % 400 == 0:
# 	year_long = 366
# else:
# 	year_long = 365

#year_end = s + datetime.timedelta(days = year_long - 1)



table = document.add_table(rows = int(weeks), cols = 5, style = "Light Grid Accent 1")
table.alignment = WD_TABLE_ALIGNMENT.CENTER

heading_cells = table.rows[0].cells
heading_cells[0].text = '2018'
heading_cells[1].text = '书'
heading_cells[2].text = '菜'
heading_cells[3].text = '电影'
heading_cells[4].text = '算法'

first_col = table.columns[0].cells

for i in range(1, len(table.rows)):
	e = "(" + s.strftime("%m/%d") + '--' + (s + delta).strftime("%m/%d") + ")"
	first_col[i].text = "第%d周\n" % i + e
	s = s + datetime.timedelta(days = 7)
	delta = datetime.timedelta(days = 6)

remain = (year_end - (s - datetime.timedelta(days = 1))).days
if (remain >= 5):
	row = table.add_row()
	e = "(" + s.strftime("%m/%d") + '--' + year_end.strftime("%m/%d") + ")"
	row.cells[0].text = "第53周\n" + e

document.add_page_break()
document.save('demo.docx')
