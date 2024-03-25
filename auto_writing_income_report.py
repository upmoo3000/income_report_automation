'''
[필요 파일 목록]
- transaction_history = 거래내역조회서
- students_list = 입사자 명단
- income_report = 수입보고서

[사용 방법]
1. 작성할 기간의 거래내역조회서 다운 후 다른 이름으로 저장하여 xlsx 확장자로 저장
2. 파일명 '[월]_transaction_history'로 변경
3. auto_writing_income_report.py 내 18행의 불러올 파일명을 방금 다운 받은 거래내역조회서의 파일명과 똑같이 변경
4. ctrl + alt + N 하면 파이썬 코드가 실행되어 수입결의서에 학생 이름 자동 입력
5. 수입결의서 열어 확인 필요라고 입력된 학생 및 기타 서식 정리
'''
import pandas as pd
import re
import openpyxl as op

# 엑셀 파일 불러오기
transactions = pd.read_excel('3_transaction_history.xlsx')
studentsList = pd.read_excel('students_list.xlsx')
wb = op.load_workbook(r"income_report.xlsx") 
ws = wb.active
# 스타일 설정
styles = op.styles      

# 입금내역 변수 저장 및 입사자명단에서 학생명 추출하여 배열화
depositors = transactions['내역']
studentNamesFromStudentsList = [i for i in studentsList['이름']]
depositorNamesFromStudentsList = [i for i in studentsList['입금자명']]
onlyDepositorNames = []

# 입금내역에서 숫자 제거하여 문자만 추출
def extractNonNumeric(val):
    nonNumeric = re.sub(r'\d+', '', val).strip()
    return nonNumeric

# 입금자명 배열 전체 탐색하며 숫자 제거하여 onlyDepositorNames append
for val in depositors:
    if type(val) is float:
        continue
    onlyDepositorNames.append(extractNonNumeric(val))

# 입금자명행에서 입금자명 유무 확인
def searchStudentIdx(name):
    for i, studentName in enumerate(studentNamesFromStudentsList, 0):
        if name not in studentName:
            continue
        if name in studentName:
            return i

# 입금자명행에서 입금자명 유무 확인
def searchDepositorIdx(name):
    for i, val in enumerate(depositorNamesFromStudentsList, 0):
        depositorName = str(val)
        if depositorName == 'nan' or name not in depositorName:
            continue
        if name in depositorName:
            return i

# 결괏값 삽입
for i, name in enumerate(onlyDepositorNames, 0):
    studentIdx = searchStudentIdx(name) # 입사자명단 이름행에서 입금자명 검색
    if studentIdx: 
        ws['B'+ str(i + 2)].value = studentNamesFromStudentsList[studentIdx]
        continue
    depositorIdx = searchDepositorIdx(name) # 이름행에서 입금자명 검색 안 될 시 입금자명행에서 입금자명 검색 후 삽입
    if depositorIdx:
        ws['B'+ str(i + 2)].value = studentNamesFromStudentsList[depositorIdx]
        ws['B'+ str(i + 2)].fill = styles.PatternFill(fill_type ='solid', fgColor = styles.Color('FFFF00'))
    else: # 이름행 및 입금자명행 둘 다 검색 안 될 시 확인 필요 문구 삽입&빨간색칠
        ws['B'+ str(i + 2)].value = name + ' 확인 필요'
        ws['B'+ str(i + 2)].fill = styles.PatternFill(fill_type ='solid', fgColor = styles.Color('FFFF0000'))
                
# 수입결의서 저장
wb.save('income_report.xlsx')
