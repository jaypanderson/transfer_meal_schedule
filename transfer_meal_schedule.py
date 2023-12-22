import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from tkinter import filedialog



def open_excel_file() -> str:
    return filedialog.askopenfilename(title='献立表を選択してください。')


def find_date_ranges(sheet: openpyxl.Worksheet) -> list[list[int]]:
    start = find_start_of_dates(sheet)

    date_ranges = []
    for row in sheet.iter_rows():

        if row[0].value is not None:





def extract_meal_data_big_kids(path: str) -> dict:
    book = openpyxl.load_workbook(path)
    sheet = book.active
    date_ranges = find_date_ranges(sheet)


def transfer_meal_schedule_big_kids():
    path = open_excel_file()
    meal_data = extract_meal_data_big_kids(path)
    pass


def main():
    pass


if __name__ == '__main__':
    main()