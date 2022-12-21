from openpyxl import load_workbook

# data_only=True로 해줘야 수식이 아닌 값으로 받아온다. 
load_wb = load_workbook(r"C:\Users\gpfks\OneDrive\바탕 화면\taxbill\exp.xlsx", data_only=True)
# 시트 이름으로 불러오기 
load_ws = load_wb['Sheet']
# 셀 주소로 값 출력
print(load_ws['C3'].value)