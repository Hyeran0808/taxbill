from openpyxl import load_workbook

# data_only=True로 해줘야 수식이 아닌 값으로 받아온다. 
load_wb = load_workbook(r"C:\Users\gpfks\OneDrive\바탕 화면\taxbill\exp.xlsx", data_only=True)
# 시트 이름으로 불러오기 
load_ws = load_wb['Sheet']
# 셀 주소로 값 출력
# print(load_ws.cell(row=7, column=3).value)
# 시트의 최대 행과 열의 값
max_ro = load_ws.max_row
max_col = load_ws.max_column
print("maxro : " , max_ro)
print("max_col : " , max_col)

# 값을 저장할 딕셔너리 명 : dic
# 순서 : 코드, 상호명, 사업자번호, 금액, 배수
dic = {}
for num in range(2,max_ro-1):
    print(num)
    dic[num] = [load_ws.cell(row=num, column=1).value, 
    load_ws.cell(row=num, column=3).value, 
    load_ws.cell(row=num, column=5).value,
    load_ws.cell(row=num, column=10).value, 1]
    print(dic[num])