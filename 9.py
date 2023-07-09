import openpyxl
from openpyxl.reader.excel import load_workbook
from russian_names import RussianNames

def generateFIO(amount:int) -> list:
    fio_list = []
    fio = ''
    for i in range(amount):
        fio = str(i + 1) + ' ' + RussianNames().get_person()
        fio_list.append(tuple(fio.split()))
    print(fio)
    print(fio_list)
    return fio_list

def writeToXlsx(write_list: list):
    try:
        wb = load_workbook('people.xlsx')
        sheet = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "People"
        sheet['A1'] = "Номер"
        sheet['B1'] = "Фамилия"
        sheet['C1'] = "Имя"
        sheet['D1'] = "Отчество"
    last_row = sheet.max_row

    for i, fio in enumerate(write_list):
        sheet.cell(row=last_row + i + 1, column=1).value = last_row + i
        sheet.cell(row=last_row + i + 1, column=2).value = fio[3]
        sheet.cell(row=last_row + i + 1, column=3).value = fio[1]
        sheet.cell(row=last_row + i + 1, column=4).value = fio[2]
    wb.save('people.xlsx')

def main():
    input_number = int(input('Enter a number: '))
    fio_list = generateFIO(input_number)
    writeToXlsx(fio_list)

if __name__ == "__main__":
    main()