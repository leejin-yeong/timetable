import openpyxl 

# 각 셀의 정보를 저장할 클래스
class Cells:
    #생성자 - 초기값 설정
    def __init__(self, row, col, people):
        self.row = row # 행이름
        self.KV = dict(col, people) #열, 사람list
        return self

class member:
    name = ""
    current_time = 0
    total_num = 0

    def __init__(self, name, total_num):
        self.name = name
        self.total_num = total_num

#근무자 리스트 생성
MemberList = [ member("재현",21), member("병국",16), member("예윤",20),
            member("혜지",23), member("태훈",19), member("유진",20),
            member("서연",19), member("한솔",20), member("희지",18),
            member("현빈",13), member("준범",5)
            ]

# 엑셀 파일있는 경로
path = "C:/Users/cooki/Desktop/timetable/example.xlsx"

# workbook 객체
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb['Sheet1']
#workbook.active로 하면 현재 활성화된 시트를 뜻하는 듯.
write_wb = openpyxl.Workbook()
result = write_wb.active

cols = 1
for c in ws.columns:#같은 열부터 읽음
    rows = 1
    for r in c:
        people = r.value.split() #공백 기준 나누기
        
        if len(people) < 3:
            result.cell(rows,cols,r.value)
            print(r.value)
        rows += 1
        if rows > 7 : break
    cols += 1
    if cols > 5 : break
write_wb.save("reuslt.xlsx")