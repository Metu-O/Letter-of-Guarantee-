# Letter-of-Guarantee-
#this allows you to ignore all the strikethroughs in a column and add everything without a strikethrought to a list
import openpyxl as opx
from openpyxl.styles import Font
def ignore_strikethrough(cell):
    if cell.font.strike or cell.value == None or cell.value == 'Company':
        return False
    else:
        return True
path = 'C:/Users/metu.osele/PycharmProjects/Customers.xlsx'
wb = opx.load_workbook(path)
ws = wb.active
colB = ws['B']
fColB = filter(ignore_strikethrough, colB)
list = []
for i in fColB:
    list.append("{2}".format(i.column, i.row, i.value))
print(list)

#this creates a document recursively and saves it to the assigned location
for i in list:
    import docx
    doc = docx.Document()
    doc.add_heading('CONTINUING LETTER OF GUARANTEE', 0)
    doc.add_paragraph('Safe Foods Corporation hereby guarantees that the Cecure product hereafter packed, shipped or consigned by it to' + ' ' + i + ' ' + 'shall conform to the following: ')
    doc.add_paragraph('It is not, as of the time of shipment, adulterated or misbranded within the meaning of the Federal Food, Drug and Cosmetic Act, or any amendment thereto; and specifically the Food Additives Amendment thereto. Safe Foods Corporation further guarantees that the articles and materials are not those which may not, under the provisions of the Section 404 and 505 of said Act, be introduced into Interstate commerce.')
    doc.add_paragraph('This guarantee is continuing and shall be in full force and effect until revoked in writing. Furthermore, every shipment is specifically covered by similar warranties on all invoices. The rights, remedies, and limitations of liabilities of the parties are as set forth in and governed by the Contract previously entered into by the parties, and this continuing letter of guarantee does not alter, add, or take away any of those rights, remedies, or limitations of liabilities.')
    doc.add_paragraph('Dated: January 1, 2021\nSafe Foods Corporation')
    doc.add_paragraph('BY: ____________________________________________\nBeatrice Maingi,\
                  \nSenior Manager, Regulatory Affairs & QA/QC Laboratory ')
    doc.save(i+'.docx')

