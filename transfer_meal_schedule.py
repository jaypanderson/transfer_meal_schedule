import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from tkinter import filedialog
from typing import Union
import docx
from copy import deepcopy


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


def copy_paragraph(original_para, new_doc):
    new_para = new_doc.add_paragraph()
    for run in original_para.runs:
        new_run = new_para.add_run(run.text)
        # Copy basic formatting (extend this as needed)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline

def copy_table(original_table, new_doc):
    new_table = new_doc.add_table(rows=0, cols=len(original_table.columns))
    for row in original_table.rows:
        cells = new_table.add_row().cells
        for i, cell in enumerate(row.cells):
            cells[i].text = cell.text


def paste_meal_data_big_kids(path: str, meal_data_big_kids: dict):
    doc = docx.Document(path)
    new_path = new_file_path(path, added_text='complete_test')
    new_doc = docx.Document(new_path)
    for i, _ in meal_data_big_kids:
        body_elements = deepcopy(doc.element.body)
        new_doc.element.body.extend(body_elements)
        if i+1 != len(meal_data_big_kids):
            new_doc.add_page_break()

    # elements = [(p, 'p') for p in doc.paragraphs] + [(t, 't') for t in doc.tables]
    # for i, _ in enumerate(meal_data_big_kids):
    #     for element, el_type in elements:
    #         if el_type == 'p':
    #             copy_paragraph(element, doc)
    #         elif el_type == 't':
    #             copy_table(element, doc)
    #     if i+1 != len(meal_data_big_kids):
    #         doc.add_page_break()

    doc.save(new_file_path(path, added_text='complete_test'))






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
