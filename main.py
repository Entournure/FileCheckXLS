import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
import excel_utils
import threading


def check_file_exists(row_index, resource_path, filename):
    for root, dirs, files in os.walk(resource_path):
        for file in files:
            full_path = os.path.join(root, file)
            if os.path.isfile(full_path):
                filename_no_ext, _ = os.path.splitext(file)
                if filename_no_ext == filename:
                    print(f"Row {row_index}: {filename} 이 {resource_path} 경로에 존재합니다.")
                    return True
    return False


def print_missing_file(row_index, filename, resource_path):
    with open('result.txt', 'a') as file:
        file.write(f"Row {row_index}: {filename} 이 {resource_path} 경로에 존재하지 않습니다.\n")
    print('\033[91m' + f"Row {row_index}: {filename} 이 {resource_path} 경로에 존재하지 않습니다." + '\033[0m')


class App:
    def __init__(self):
        self.window = None
        self.path_excel = ""       # 엑셀 파일 경로
        self.path_resource = ""     # 리소스 검색할 경로
        self.create_window()

    def create_window(self):
        self.window = tk.Tk()
        self.window.geometry('300x200')
        self.create_button_select_path_excel()
        self.create_button_select_path_resource()
        self.create_button_reset()
        self.create_button_table_skill()
        self.create_file_result()
        self.window.mainloop()

    def create_button_select_path_excel(self):
        button = tk.Button(self.window, text="엑셀 파일 경로 선택", command=self.click_btn_select_path_excel)
        button.pack(side="top", pady=10)

    def create_button_select_path_resource(self):
        button = tk.Button(self.window, text="리소스 파일 경로 선택", command=self.click_btn_select_path_resource)
        button.pack(side="top", pady=10)

    def create_button_reset(self):
        button = tk.Button(self.window, text="경로 리셋", command=self.click_btn_reset)
        button.pack(side="top", pady=10)

    def create_button_table_skill(self):
        button = tk.Button(self.window, text="Skill.xls", command=self.click_btn_skill)
        button.pack(side="top", pady=10)

    # 결과 파일 생성
    def create_file_result(self):
        with open('result.txt', 'w') as self.file:
            self.file.write("")

    def click_btn_select_path_excel(self):
        self.path_excel = filedialog.askdirectory()

    def click_btn_select_path_resource(self):
        self.path_resource = filedialog.askdirectory()

    def click_btn_reset(self):
        self.path_excel = ""  # 초기화할 엑셀 파일 경로
        self.path_resource = ""  # 초기화할 리소스 파일 경로

    def click_btn_skill(self):
        table = ""  # 엑셀 파일 경로
        resource = ""  # 리소스 파일 경로
        cols = ['Unused', 'ClassDependent', 'Resource']
        # Skill.xls 테이블의 cols열들의 값을 5행부터 확인
        threading.Thread(target=self.check_table_skill, args=(table, resource, cols, 5)).start()

    def check_table_skill(self, table, resource, cols, cell_row):
        file = self.path_excel + table
        resource = self.path_resource + resource
        book = xw.Book(file)
        sheet = book.sheets[0]  # 첫번째 시트 지정
        last_row = sheet.range('A1').current_region.last_cell.row  # 테이블의 가장 마지막 행 번호
        dict_cols = excel_utils.return_cols_alpha(file, 0, cols)

        # 각 열에 대한 데이터 범위 생성 및 값 가져오기
        values_1, values_2, values_3 = [
            sheet.range(f"{dict_cols[col]}{cell_row}:{dict_cols[col]}{last_row}").value
            for col in ['Unused', 'ClassDependent', 'Resource']
        ]
        # 파일이 있는지 확인
        for row_index, (cell_value_1, cell_value_2, cell_value_3) in enumerate(zip(values_1, values_2, values_3),
                                                                               start=cell_row):
            if cell_value_1 is not None and int(cell_value_1) == 1:     # 비사용 컬럼 값이 1일 경우 확인하지 않음
                continue
            else:
                cell_value_3 = str(cell_value_3)
                # {Class}{Gender} 문자열 존재시 각 성별 체크
                if "{Class}{Gender}" in cell_value_3:
                    new_string_male = cell_value_3.replace("{Class}{Gender}", str(cell_value_2) + "Male")
                    new_string_female = cell_value_3.replace("{Class}{Gender}", str(cell_value_2) + "Female")
                    found = check_file_exists(row_index, resource, new_string_male)
                    if not found:
                        print_missing_file(row_index, new_string_male, resource)
                    found = check_file_exists(row_index, resource, new_string_female)
                    if not found:
                        print_missing_file(row_index, new_string_female, resource)
                # {Class}{Gender} 문자열이 없을 경우 체크
                else:
                    found = check_file_exists(row_index, resource, cell_value_3)
                    if not found:
                        print_missing_file(row_index, cell_value_3, resource)

        print("Verification complete")


app = App()