# git_daul
"""
Python 3.6 이상에서 실행되어야 합니다.
Constant에 있는 값들을 수정하면 그것에 맞추어 실행되도록 하였으니, 해당 문자열 값을 수정해주세요.
실행 전에 터미널에서 다음 명령어를 실행해주세요.
pip install openpyxl
"""
# pip install openpyxl

from openpyxl import load_workbook
from string import ascii_uppercase
from unicodedata import normalize
from datetime import datetime

# Constant 파일 환경에 따라 수정 필요
target_workbook = 'data.xlsx'
target_worksheet = 'export'
target_header = 'Mention Content'
normalized_header = 'Normalized Mention Content'
start_row_number = 1


# Function
def find_target_column(headers, target_header):
    for cell in headers:
        if cell.value == target_header: # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.cell.cell.html
            return ascii_uppercase[cell.column - 1] # https://stackoverflow.com/a/23199756


wb = load_workbook(target_workbook) # https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file
ws = wb[target_worksheet]
headers = ws[start_row_number]

target_column = find_target_column(headers, target_header)
normalize_column = ascii_uppercase[len(headers)]


ws[f'{normalize_column}{start_row_number}'] = normalized_header # https://www.python.org/dev/peps/pep-0498/

for cell in ws[target_column][start_row_number:]:
    if (cell.value):
        ws[f'{normalize_column}{cell.row}'] = normalize('NFC', cell.value)

created_at = datetime.utcnow().isoformat()
# https://docs.python.org/3/library/datetime.html#datetime.datetime.utcnow
# https://docs.python.org/3/library/datetime.html#datetime.date.isoformat

#필요에 따라 저장명 변경
wb.save(f'result.xlsx') # https://openpyxl.readthedocs.io/en/stable/tutorial.html#saving-as-a-stream

