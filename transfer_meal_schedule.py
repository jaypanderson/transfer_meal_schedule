import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from tkinter import filedialog
from typing import Union


def open_excel_file() -> str:
    return filedialog.askopenfilename(title='献立表を選択してください。')


def find_start_of_dates(sheet: Worksheet) -> int:
    for i, row in enumerate(sheet.iter_rows(), start=1):
        if isinstance(row[0].value, int):
            # we add one because openpyxl uses 1 indexing.
            return i


def find_date_ranges(sheet: Worksheet) -> dict[int, tuple[str, int, int]]:
    start_row = find_start_of_dates(sheet)
    date_ranges = {}
    date = None
    day = None
    start = None
    for i, row in enumerate(sheet.iter_rows(min_row=start_row), start=start_row):
        if i == start_row:
            start = i
            date = row[0].value
            day = row[1].value
        elif row[0].value is not None:
            end = i - 1
            date_ranges[date] = (day, start, end)
            start = i
            date = row[0].value
            day = row[1].value

    return date_ranges


def extract_meal_data_big_kids(path: str) -> dict:
    book = openpyxl.load_workbook(path)
    sheet = book.active
    date_ranges = find_date_ranges(sheet)
    meal_data_big_kids = {}
    for key, val in date_ranges.items():
        day = val[0]
        start = val[1]
        end = val[2]
        date_ranges[key] = (day,) + gather_text(start, end)
    print(date_ranges)


def transfer_meal_schedule_big_kids():
    path = open_excel_file()
    meal_data = extract_meal_data_big_kids(path)
    pass


def main():
    transfer_meal_schedule_big_kids()


if __name__ == '__main__':
    main()
