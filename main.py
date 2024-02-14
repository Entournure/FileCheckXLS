import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog


class App:
    def __init__(self):
        self.window = None
        self.path = ""
        self.create_window()

    def create_window(self):
        self.window = tk.Tk()
        self.create_button_SelectBranch()
        self.create_button_CharacterTransform()
        self.create_button_ChatBalloon()
        self.create_button_test()
        self.window.mainloop()

    def create_button_SelectBranch(self):
        button = tk.Button(self.window, text="경로 선택", command=self.select_path)
        button.pack()

    def create_button_CharacterTransform(self):
        button = tk.Button(self.window, text="CharacterTransform", command=self.check_CharacterTransform)
        button.pack()

    def create_button_ChatBalloon(self):
        button = tk.Button(self.window, text="ChatBalloon", command=self.check_ChatBalloon)
        button.pack()

    def create_button_test(self):
        button = tk.Button(self.window, text="test", command=self.check_test)
        button.pack()

    def select_path(self):
        self.path = filedialog.askdirectory()
        print(self.path)

    # CharacterTransform 버튼 클릭시 호출 함수
    def check_CharacterTransform(self):
        path_table = ""  # 엑셀 데이터파일 경로
        path_target = ""  # 파일 존재를 확인할 경로
        self.check_table(path_table, path_target, "N", 5)   # N열 5행부터 N열 탐색

    # ChatBalloon 버튼 클릭시 호출 함수
    def check_ChatBalloon(self):
        path_table = ""  # 엑셀 데이터파일 경로
        path_target = ""  # 파일 존재를 확인할 경로
        self.check_table(path_table, path_target, "I", 5)   # I열 5행부터 I열 탐색

    # 파일 직접 선택해 확인
    def check_test(self):
        self.path = filedialog.askopenfilename()
        print(self.path)
        file = self.path  # 엑셀 데이터파일 경로
        resource = ""  # 파일 존재를 확인할 경로
        book = xw.Book(file)

        sheet = book.sheets[0]  # 첫 번째 시트를 선택하려면 [0]을 사용합니다.
        # 마지막 행 인덱스 찾기
        last_row = sheet.range('A1').current_region.last_cell.row
        # print(f"행 개수: {last_row}")
        for row in range(1, last_row + 1):
            cell_value = sheet.range(f"A{row}").value  # A{row} 셀값을 저장
            if cell_value is not None and cell_value != "":
                cell_value = str(cell_value)
                for root, dirs, files in os.walk(resource):
                    for file in files:
                        full_path = os.path.join(root, file)
                        if os.path.isfile(full_path):
                            filename, file_ext = os.path.splitext(file)
                            if filename == cell_value:
                                print(f"Row {row}: {file} 이 {resource} 경로에 존재합니다.")
                                break
                    else:
                        continue
                    break
                else:
                    print('\033[91m' + f"Row {row}: {cell_value} 이 {resource} 경로에 존재하지 않습니다." + '\033[0m')
        print("확인 완료")

    # 파일 존재 여부를 확인하는 함수
    def check_table(self, path_table, path_target, cell_col, cell_row):
        file = self.path + path_table  # 엑셀 데이터파일 경로
        resource = self.path + path_target
        book = xw.Book(file)

        sheet = book.sheets[0]  # 첫 번째 시트를 선택하려면 [0]을 사용합니다.
        # 마지막 행 인덱스 찾기
        last_row = sheet.range('A1').current_region.last_cell.row
        # print(f"행 개수: {last_row}")
        for row in range(cell_row, last_row + 1):
            cell_value = sheet.range(cell_col + str(cell_row)).value  # 특정 셀의 값을 가져옴
            if cell_value is not None and cell_value != "":
                cell_value = str(cell_value)
                for root, dirs, files in os.walk(resource):
                    for file in files:
                        full_path = os.path.join(root, file)
                        if os.path.isfile(full_path):
                            filename, file_ext = os.path.splitext(file)
                            if filename == cell_value:
                                print(f"Row {row}: {file} 이 {resource} 경로에 존재합니다.")
                                break
                    else:
                        continue
                    break
                else:
                    print('\033[91m' + f"Row {row}: {cell_value} 이 {resource} 경로에 존재하지 않습니다." + '\033[0m')
        print("확인 완료")


app = App()
