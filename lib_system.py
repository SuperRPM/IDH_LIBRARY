'''
1. 도서목록 데이터가 저장된 엑셀파일을 불러온다
2. row, column으로 구분하여 도서 목록을 분류한다.
3. 분류된 도서에 대출가능상태를 설정, 변경 할 수 있게 한다 -> 도서 검색 엔진, 도서 대출, 반납 엔진
4. 전체 도서목록을 대출가능상태와 같이 GUI로 구현한다.
'''

import openpyxl

'''1. 도서목록 데이터가 저장된 엑셀파일을 불러온다'''
book_excel_file = openpyxl.load_workbook('book_list.xlsx')
file_name = 'book_list.xlsx'
list_sheet = book_excel_file.worksheets[0]

'''2. row, column으로 구분하여 도서 목록을 리스트에 저장한다.'''
book_list = []
#분류번호, 제목, 대출가능 순으로 dict type으로 생성후 list에 저장
for row in list_sheet.rows:
    data = {}
    data['number'] = row[0].value
    data['name'] = row[1].value
    data['loan'] = row[2].value
    book_list.append(data)

#엑셀에 row1은 불필요하므로 삭제
del book_list[0]

################################
#데이터가 정상적으로 만들어졌는지 확인
#for data in book_list:
#    print(data)
#    print(data['name'])
################################

'''도서 검색 엔진'''

#제목으로 검색
def search_engine(book_name):
    for data in book_list:
        if book_name in data['name']:
            print(data['number'])
            print(data['name'])
            print(data['loan'])
            print("")
            check = True #중복되는 제목의 책을 찾고 책을 아예 찾지 못했을 경우 안내문을 print해주기 위한 변수
    if check == True:
        return
    else:
        print("찾으려는 책이 없습니다. 제목을 확인해주세요")

#분류번호로 검색
def search_engine_number(number):
    for data in book_list:
        if number == data['number']:
            print(data['number'])
            print(data['name'])
            print(data['loan'])
            return
    print('찾으려는 책이 없습니다. 분류번호를 확인해주세요')

#실제 검색 구현
def search():
    user_type = input('책 제목 또는 분류번호를 입력하세요 :')
    if user_type[0] == '2':
        search_engine_number(user_type)
    else:
        search_engine(user_type)

'''도서 대출 엔진'''
def book_loan_engine(book_name):
    for data in book_list:
        if book_name in data['name']:
            if data['loan'] == 1:
#                print(f"{data['name']} 대출 완료")
                print(data['number'])
                print(data['name'])
                print('대출 완료')
                row_number = book_list.index(data)
                input_cell = row_number + 2
                list_sheet[f'C{input_cell}'] = 0
                book_excel_file.save(filename = file_name) #변경된 데이터 엑셀에 다시 저장! 중요!!
                return
            else:
                print('이미 대출되어있는 도서입니다.')
                return
    print('대출하려는 책이 없습니다. 제목을 확인해주세요')

def book_loan_engine_number(number):
    for data in book_list:
        if number == data['number']:
            if data['loan'] == 1:
                print(data['number'])
                print(data['name'])
                print('대출 완료')
                row_number = book_list.index(data)
                input_cell = row_number + 2
                list_sheet[f'C{input_cell}'] = 0
                book_excel_file.save(filename = file_name)
                return
            else:
                print('이미 대출되어있는 도서입니다.')
                return
    print('대출하려는 책이 없습니다. 제목을 확인해주세요')


#실제 대출 구현
def loan():
    user_type = input('책 제목을 입력하세요 :')
    if user_type[0] == 2:
        book_loan_engine_number(user_type)
    else:
        book_loan_engine(user_type)

'''도서 반납 엔진'''
def ban_nap_engine(number):
    for data in book_list:
        if number == data['number']:
            if data['loan'] == 0:
                print(number)
                print(data['name'])
                print('반납 완료')
                row_number = book_list.index(data)
                input_cell = row_number + 2
                list_sheet[f'C{input_cell}'] = 1
                book_excel_file.save(filename = file_name)
                return
            else:
                print('이미 반납되어있는 도서입니다.')
                return
    print('반납하려는 책은 도서관에 없는 책입니다. 목록을 확인해주세요.')

#실제 도서 반납 구현
def ban_nap():
    user_type = input('분류 번호를 입력해주세요 (예)20-도B-01 : ')
    ban_nap_engine(user_type)

ban_nap()