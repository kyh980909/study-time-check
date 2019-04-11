"""
descrip       : 스터디 데이터 조회 프로그램
developer     : 김용호
develop date  : 2019-04-10
last update   : 2019-04-11
e-mail        : kyh980909@gmail.com
"""

from openpyxl import load_workbook


def header_fix(header_list):  # header 전처리 함수
    header1 = []
    j = 0

    for lst in header_list[0]:
        if lst is not None:
            header1.append(lst)

    for index in range(len(header_list[1])):
        if header_list[1][index] is None:   # header_list[1][index]값이 None이면 header1의 있는 값을 차례로 넣음
            header_list[1][index] = header1[j]
            j += 1

    temp = header_list[1]
    header_list.clear()
    temp[3] = str(temp[3]).replace('\n', '')  # 한학기 목표시간 개행 문자 변환

    return temp  # 전처리된 리스트 반환


def study_time_search(filename, user_name):
    load_wb = load_workbook(filename, data_only=True)
    load_ws = load_wb.active

    all_values = []
    headers = []

    for index, row in enumerate(load_ws.rows):
        row_value = []
        for cell in row:
            if cell.value == '필수활동' or cell.value == '출석 시간(분)':
                pass
            else:
                row_value.append(cell.value)
        if index == 0:
            headers.append(row_value)
        elif index == 1:
            headers.append(row_value)
        else:
            all_values.append(row_value)

    headers = header_fix(headers)

    for header in headers:
        print('%-014s' % header, end='')
    print('\n')

    is_check = False  # 목록에 있는지 없는지 체크 하는 변수
    for values in all_values:
        if user_name in values:
            for value in values:
                print('%-014s' % value, end='')
                is_check = True
        else:
            continue
        print('\n')

    if not is_check:   # 결과가 없는 경우에만 출력
        print('결과가 없습니다.')