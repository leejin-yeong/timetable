import openpyxl 

# 엑셀 파일있는 경로
path = "C:/Users/cooki/Desktop/timetable/example.xlsx"

# workbook 객체
workbook = openpyxl.load_workbook(path, data_only=True)
worksheet = workbook['Sheet1']
#workbook.active로 하면 현재 활성화된 시트를 뜻하는 듯.

cell_obj = worksheet.cell(3,2)
print(cell_obj.value)
