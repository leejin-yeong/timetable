import openpyxl 

# 각 셀의 정보를 저장할 클래스
class Cells:
    #생성자 - 초기값 설정
    def __init__(self, row, col, people):
        self.row = row # 행이름
        self.KV = dict(col, people) #열, 사람list
        return self

class member:
    current_time = 0
    total_num = 0

    def __init__(self, total_num):
        self.total_num = total_num

#근무자 리스트 생성
MemberList = {
    "재현": member(21), "병국": member(16), "예윤": member(20),
    "혜지": member(23), "태훈": member(19), "유진": member(20),
    "서연": member(19), "한솔": member(20), "희지": member(18),
    "현빈": member(13), "준범": member(5)
}

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
            #현재 근무시간 업데이트
            for person in people:
                if rows == 7 : add = 4 # 저녁
                else: add = 1.5 # 낮
                MemberList[person].current_time += add

        rows += 1
        if rows > 7 : break
    cols += 1
    if cols > 5 : break
write_wb.save("reuslt.xlsx")