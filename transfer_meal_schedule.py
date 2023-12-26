import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from tkinter import filedialog
from typing import Union
from copy import copy



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


# add result to the end of the file name
def new_file_path(path: str, added_text: str = 'result') -> str:
    """
    This function creates a new name for the path of a save file. This is to avoid saving over the original Excel file
    that was used to create the new one. It places a new text between the name and the extension name. If no added_text
    is provided the default 'result' will be used.
    :param path: The path of the original Excel file.
    :param added_text: The text that will be added inbetween the name and the extension name of the original path.
    :return: The newly formed name path where the new Excel file will be saved to.
    """
    idx = path.find('.')
    ans = path[:idx] + added_text + path[idx:]
    return ans


def copy_sheet(sheet: Worksheet, new_sheet: Worksheet) -> None:
    """
    copy the contents from one sheet to another sheet. This function only copies the contents of the cells and its
    style, other aspects of the sheet are copied with other functions.
    :param sheet:The base sheet from which the cells are copied.
    :param new_sheet: The new sheet where the cell contents and style are pasted.
    :return: None
    """
    for row in sheet:
        for cell in row:
            new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)

            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
                new_cell.comment = copy(cell.comment)


# replicate the cell merges from base template.
def merge_cells(sheet: Worksheet, new_sheet: Worksheet) -> None:
    """
    When a new sheet is created and everything is copied, the one thing that will not be copied is the location of where
    there are merged cells.  This function iterates over all the merged cells of the base worksheet and then merges
    those cells in the new sheet.
    :param sheet: The base work sheet this function will use to iterate over the location of the merged cells.
    :param new_sheet: The new work sheet where the function will merge the cells from the base work sheet.
    :return: None
    """
    for merged_cell_range in sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_cell_range))


def paste_meal_data_big_kids(path: str, meal_data_big_kids: dict):
    book = openpyxl.load_workbook(path)
    sheet = book.active
    for key, val in meal_data_big_kids.items():
        new_sheet = book.create_sheet(f'{key}({val[0]})')
        copy_sheet(sheet, new_sheet)
        merge_cells(sheet, new_sheet)

    book.save(new_file_path(path, added_text='_test_complete'))








def transfer_meal_schedule_big_kids():
    copy_path = choose_file(1)
    paste_path = choose_file(2)
    meal_data_big_kids = extract_meal_data_big_kids(copy_path)
    print(meal_data_big_kids)
    paste_meal_data_big_kids(paste_path, meal_data_big_kids)




def main():
    transfer_meal_schedule_big_kids()


if __name__ == '__main__':
    main()
