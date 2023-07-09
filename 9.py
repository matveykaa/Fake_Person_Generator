import openpyxl
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
    wb = openpyxl.load_workbook('table.xlsx')

def main():
    input_number = int(input('Enter a number - '))
    generateFIO(input_number)
    writeToXlsx(generateFIO(input_number))

if __name__ == "__main__":
    main()