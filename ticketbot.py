import gspread

#한글이 먼저, 영어 이름이 나중에 오도록 하는 정렬 규칙
def sort_rule(item):
    return item[0] if ord(item[0][0]) < ord('ㄱ') else chr(0) + item[0]

#시트 관리할 봇의 정보가 담긴 json파일 입력
gc = gspread.service_account(filename = "credentials.json")

#티켓 관리를 위한 스프레드시트 이름 입력
sh = gc.open("(23-2) 가을연주회 티켓의 사본")

#무료 티켓 폼과 연결된 시트 이름 입력
free = sh.worksheet("무료 티켓")

#유료 티켓 폼과 연결된 시트 이름 입력
all = sh.worksheet("전체 티켓")

#전체 티켓 관리용 시트 이름 입력
pay = sh.worksheet("유료 티켓")

#전체 티켓 관리용 리스트
tickets = []

#무료티켓 수합 결과에서 이름, 연락처, 필요 매수 가져옴
for i in free.get_all_records():
    tickets.append(list(i.values())[1:4])

#유료티켓 수합 결과에서 이름, 연락처, 구매 매수 가져옴
for i in pay.get_all_records():
    tickets.append(list(i.values())[1:4])

#sorting을 위해 int와 str이 섞여있는 전화번호의 형식 맞추기
for i in tickets:
    i[1] = str(i[1])

#이름 순으로 정렬
tickets.sort(key = sort_rule)

#전체 티켓 관리용 시트에 무료 티켓 수합 결과와 유료 티켓 수합 결과 합쳐서 작성
all.batch_update([{"range":f"B2:D{len(tickets) + 1}", "values": tickets}])

'''
개발 중
1. 동명이인 처리
    1-1. 연속된 서로 다른 동명이인 처리
2. 동일 인물 티켓 매수 합쳐서 처리 (cell merge 하는 방법 찾아보기)
'''

#동명이인 관리용 리스트
samenames = []

#동명이인이 있으면 이름 앞 행 cell을 노란색 배경으로 채우기
for i in range(len(tickets)):
    if tickets[i][0] == tickets[i-1][0]:
        temp = {"range": f"A{i+1}:A{i+2}", "format":{"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}}}
        samenames.append(temp)

#배경색 적용
all.batch_format(samenames)

