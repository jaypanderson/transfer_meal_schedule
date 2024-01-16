import openpyxl
import os
import sys
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.chart import BarChart, Reference
from tkinter import filedialog
from typing import Union
from copy import copy
from itertools import zip_longest


def choose_file(file_type: int) -> str:
    if file_type == 1:
        title = '献立表を選択してください。'
    elif file_type == 2:
        title = '離乳食献立を選択してください'
    elif file_type == 3:
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


def gather_text_big_kids(sheet: Worksheet, start: int, end: int) -> tuple[str]:
    breakfast = []
    lunch = []
    snack = []
    for row in sheet.iter_rows(min_row=start, max_row=end):
        if row[2].value is not None:
            lunch.append(row[2].value)
        if row[6].value is not None:
            breakfast.append(row[6].value)
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
        meal_data_big_kids[key] = (day,) + gather_text_big_kids(sheet, start, end)
    return meal_data_big_kids


# todo There has to be a way to refactor so that this code isnt so bloated. The issue is that depending on the
# todo max_column some lists dont need to be appended and the location for where each list takes its value from
# todo is different as well. perhaps if i iterated through lists??
def gather_text_small_kids(sheet: Worksheet, start: int, end: int, max_column: str) -> tuple[str]:
    breakfast = []
    early = []
    middle = []
    late = []
    snack = []
    if max_column == 'G':
        for row in sheet.iter_rows(min_row=start, max_row=end):
            if row[2].value is not None:
                early.append(row[2].value)
            if row[3].value is not None:
                middle.append(row[3].value)
            if row[4].value is not None:
                late.append(row[4].value)
            if row[5].value is not None:
                breakfast.append(row[5].value)
            if row[6].value is not None:
                snack.append(row[6].value)
    elif max_column == 'F':
        for row in sheet.iter_rows(min_row=start, max_row=end):
            if row[2].value is not None:
                middle.append(row[2].value)
            if row[3].value is not None:
                late.append(row[3].value)
            if row[4].value is not None:
                breakfast.append(row[4].value)
            if row[5].value is not None:
                snack.append(row[5].value)
    elif max_column == 'E':
        for row in sheet.iter_rows(min_row=start, max_row=end):
            if row[2].value is not None:
                late.append(row[2].value)
            if row[3].value is not None:
                breakfast.append(row[3].value)
            if row[4].value is not None:
                snack.append(row[4].value)

    return '\n'.join(breakfast), '\n'.join(early), '\n'.join(middle), '\n'.join(late), '\n'.join(snack)


def extract_meal_data_small_kids(path: str) -> Union[dict, None]:
    if path == '':
        return None
    book = openpyxl.load_workbook(path)
    sheet = book.active
    meal_data_small_kids = {}
    date_ranges = find_date_ranges(sheet)
    max_column = get_column_letter(sheet.max_column)
    for key, val in date_ranges.items():
        day = val[0]
        start = val[1]
        end = val[2]
        meal_data_small_kids[key] = (day,) + gather_text_small_kids(sheet, start, end, max_column)
    return meal_data_small_kids


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


# copy the width and height of cells
def copy_dimensions(sheet: Worksheet, new_sheet: Worksheet) -> None:
    """
    A function that changes the new sheet's size of the cells to match that of a another sheet. Currently, it is set up
    so that it only copies the width of the cells.  If the height also needs to be copied, then Uncomment that portion
    of the code and vice versa.
    :param sheet: The sheet from which the cell dimensions will be copied from
    :param new_sheet: The sheet where the dimensions will be pasted into.
    :return: None
    """
    for row, col in zip_longest(sheet.row_dimensions, sheet.column_dimensions):
        if row is not None:
            new_sheet.row_dimensions[row].height = sheet.row_dimensions[row].height
        if col is not None:
            new_sheet.column_dimensions[col].width = sheet.column_dimensions[col].width


# copy the part of the worksheet that will be printed onto the new sheet.
def copy_print_area(sheet: Worksheet, new_sheet: Worksheet) -> None:
    """
    A certain area of a page is selected as the default area to be printed when the print button is pressed. To avoid
    having to change the print area for every single new sheet this function copies the print area attribute from the
    base sheet to the new sheet.
    :param sheet: the base sheet from which we will copy the print area attribute.
    :param new_sheet: The new sheet where we will paste the print area.
    :return: None
    """
    if sheet.print_area:
        new_sheet.print_area = sheet.print_area


def copy_margins(sheet: Worksheet, new_sheet: Worksheet):
    new_sheet.page_margins = sheet.page_margins


def copy_page_size(sheet: Worksheet, new_sheet: Worksheet):
    new_sheet.page_setup.paperSize = sheet.page_setup.paperSize


def copy_all_elements(sheet: Worksheet, new_sheet: Worksheet):
    copy_sheet(sheet, new_sheet)
    merge_cells(sheet, new_sheet)
    copy_dimensions(sheet, new_sheet)
    copy_print_area(sheet, new_sheet)
    copy_margins(sheet, new_sheet)
    copy_page_size(sheet, new_sheet)


# this function is so that when using the script the image retrieval works as well as when the executabel
# is created as well so that I don't have to write different code from when im developing and when
# im deploying.
def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# insert a image into the Excel sheet of three squares where the stamps of the teachers will go.
# todo add try except clause in case this cannot find the image, so that the program doesnt crash.
def add_shapes(new_sheet: Worksheet):
    path = resource_path('image_boxes.jpg')
    image = Image(path)
    image.anchor = 'F1'
    new_sheet.add_image(image)


# insert the big kids meal data into a single Excel sheet. As well as the date and day of the week.
def insert_data_big_kids(date: int, data: tuple[str], new_sheet: Worksheet):
    day = data[0]
    breakfast = data[1]
    lunch = data[2]
    snack = data[3]
    new_sheet['B4'].value = new_sheet['B4'].value.replace('@', str(date))
    new_sheet['B4'].value = new_sheet['B4'].value.replace('$', day)
    new_sheet['C7'].value = breakfast
    new_sheet['C16'].value = lunch
    new_sheet['C25'].value = snack


# todo refactor code because some values are not need. perhaps change the structure of the dict as well.
# insert the small kids meal data into a single Excel sheet.
def insert_data_small_kids(date: int, data: tuple[str], new_sheet: Worksheet):
    day = data[0]
    breakfast = data[1]
    early = data[2]
    middle = data[3]
    late = data[4]
    snack = data[5]
    new_sheet['F7'].value = breakfast
    new_sheet['F16'].value = early
    new_sheet['F18'].value = middle
    new_sheet['F20'].value = late
    new_sheet['F25'].value = snack


# todo I fixed the issue when the user didn't pick a file because there is none for that month. I was able to handle the
# todo error but it raised another question. what if there is mismatching dates on two of the files. which would lead
# todo the key not being in the second dict which would mean it wont be pasted. This can happen if the person creating
# todo the files accidentally inserts the wrong days.  i might have to refactor code like this but it feels messy.
'''
if meal_data_small_kids is None:
    iterate with only meal_data_big_kids
else:
    iterate by zipping small and big kids meal data
'''


# insert the collected data for the big and small kids into each new Excel sheet.
def paste_meal_data(path: str, meal_data_big_kids: dict, meal_data_small_kids: dict):
    """
    
    :param path:
    :param meal_data_big_kids:
    :param meal_data_small_kids:
    :return:
    """
    book = openpyxl.load_workbook(path)
    sheet = book.active
    for key, val_big in meal_data_big_kids.items():
        new_sheet = book.create_sheet(f'{key}({val_big[0]})')
        copy_all_elements(sheet, new_sheet)
        add_shapes(new_sheet)
        insert_data_big_kids(key, val_big, new_sheet)
        if meal_data_small_kids and key in meal_data_small_kids:
            val_small = meal_data_small_kids[key]
            insert_data_small_kids(key, val_small, new_sheet)

    del book['base']  # remove base sheet because it will not be needed when printing.
    book.save(new_file_path(path, added_text='_test_complete'))


# the function that steps through the large steps of transferring the data.
def main():
    """
    The main function of the script. First prompts the user to choose the file path for the three Excel documents needed
    to create the output documents. Extracts data from the meal schedule for the big kids and then the small kids
    (the babies) and then pastes this information into the new document this script produces.
    """
    big_kids_path = choose_file(1)
    small_kids_path = choose_file(2)
    output_path = choose_file(3)
    meal_data_big_kids = extract_meal_data_big_kids(big_kids_path)
    meal_data_small_kids = extract_meal_data_small_kids(small_kids_path)
    print(meal_data_big_kids)
    paste_meal_data(output_path, meal_data_big_kids, meal_data_small_kids)


if __name__ == '__main__':
    main()
