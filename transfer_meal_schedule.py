import openpyxl
from tkinter import filedialog


def open_excel_file() -> str:
    return filedialog.askopenfilename(title='献立表を選択してください。')


def transfer_meal_schedule_big_kids():
    path = open_excel_file()
    meal_data = extract_meal_data_big_kids(path)
    pass


def main():
    pass


if __name__ == '__main__':
    main()