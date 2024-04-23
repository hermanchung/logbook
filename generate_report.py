from openpyxl import load_workbook

# Generate CAD Format report
wb = load_workbook(filename='./HKCAD_logbook_format.xlsx')
sheet = wb['Sheet1']

column_range = {
    1:"A",
    2:"B",
    3:"C",
    4:"D",
    5:"E",
    6:"F",
    7:"G",
    8:"H",
    9:"I",
    10:"J",
    11:"K",
    12:"L",
    13:"M",
    14:"N",
    15:"O",
    16:"P",
    17:"Q",
    18:"R",
    19:"S",
    20:"T",
}

total_pages = 22

for i in range(2, total_pages):
    ws = wb.copy_worksheet(wb['Sheet1'])

for sheet in wb:
    for r in sheet:
        sheet.title = "new_sheet_title"
        for cell in r:
            if column_range[cell.column] == "A" and (int(cell.row)-4) % 26 == 0:
                cell.value = 2017

wb.save("./test.xlsx")

print(sheet['A4'].value,sheet['A30'].value)