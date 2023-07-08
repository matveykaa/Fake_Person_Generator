import openpyxl
from russian_names import RussianNames

def generateFIO():
    fio = RussianNames().get_person()
    tuple_fio = tuple(fio.split())
    print(fio)
    print(tuple_fio)
    print(type(fio))

def main():
    generateFIO()

if __name__ == "__main__":
    main()