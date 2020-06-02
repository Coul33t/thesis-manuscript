from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

filename = 'Full distance PCA All data.xlsx'
wb = load_workbook(filename=filename)

for i, sheet in enumerate(wb.worksheets):
    name = sheet.title
    sheet.title = str(i + 1)
    for j in range(5, 43):
        if sheet[f'B{j}'].value and name in sheet[f'B{j}'].value:
            sheet[f'B{j}'].value = sheet[f'B{j}'].value.replace(f'{name}!', '')

wb.save(filename)
