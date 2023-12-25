import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from tkinter import filedialog
from typing import Union


def choose_file(file_type: int) -> str:
    if file_type == 1:
        title = '献立表を選択してください。'
    elif file_type == 2:
        title = '検食簿原本を選択してください。'
    else:
        title = ''
    return filedialog.askopenfilename(title=title)


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


def gather_text(sheet: Worksheet, start: int, end: int) -> tuple[str]:
    breakfast = []
    lunch = []
    snack = []
    for row in sheet.iter_rows(min_row=start, max_row=end):
        if row[2].value is not None:
            breakfast.append(row[2].value)
        if row[6].value is not None:
            lunch.append(row[6].value)
        if row[7].value is not None:
            snack.append(row[7].value)
    return '\n'.join(breakfast), '\n'.join(lunch), '\n'.join(snack)


def extract_meal_data_big_kids(path: str) -> dict:
    book = openpyxl.load_workbook(path)
    sheet = book.active
    date_ranges = find_date_ranges(sheet)
    meal_data_big_kids = {}
    for key, val in date_ranges.items():
        day = val[0]
        start = val[1]
        end = val[2]
        meal_data_big_kids[key] = (day,) + gather_text(sheet, start, end)
    return meal_data_big_kids


def paste_meal_data_big_kids(path: str, meal_data_big_kids: dict):
    pass


def transfer_meal_schedule_big_kids():
    excel_path = choose_file(1)
    word_path = choose_file(2)
    meal_data_big_kids = extract_meal_data_big_kids(excel_path)
    print(meal_data_big_kids)
    paste_meal_data_big_kids(word_path, meal_data_big_kids)




def main():
    transfer_meal_schedule_big_kids()


if __name__ == '__main__':
    main()
