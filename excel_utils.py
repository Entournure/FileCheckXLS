import xlwings as xw


def number_to_column_name(num):
    """
    Convert a number to Excel column name.
    """
    result = ""
    while num > 0:
        remainder = (num - 1) % 26
        result = chr(65 + remainder) + result
        num = (num - 1) // 26
    return result


def return_cols_alpha(file, num_sheet, cols):
    book = xw.Book(file)
    sheet = book.sheets[num_sheet]  # 시트 지정
    dict_cols = {col: '' for col in cols}

    # dict_cols의 각 키와 값을 순회하면서 시트에서 해당 값을 찾음
    for key, value in dict_cols.items():
        column = sheet.range('1:1').value.index(key) + 1 if key in sheet.range('1:1').value else -1

        # 만약 해당 값을 찾았다면 결과 딕셔너리에 추가
        if column != -1:
            # dict_cols의 해당 키에 찾은 열의 위치를 저장
            dict_cols[key] = number_to_column_name(column)

    # dict_cols 출력
    # print(dict_cols)
    return dict_cols
