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
tickets = list()

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

#동명이인 관리용 리스트
samenames = list()

#1행 배경색 하얀색으로 초기화
samenames.append({"range": "A", "format": {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}})

#하나의 동명이인 그룹 앞에 다른 이름의 동명이인 그룹이 있었는지 확인
toggle = False

#동명이인이면 맨 앞 행의 배경을 노란색으로 칠함; 서로 다른 이름의 동명이인 그룹들이 연속된 경우 노란색과 주황색을 번갈아가며 칠함
i = 0
while i < len(tickets) - 1:
    if tickets[i][0] == tickets[i+1][0]:
        toggle = not toggle
        j = i + 1
        while tickets[j][0] == tickets[j+1][0] and j < len(tickets):
            j += 1
        if toggle:
            temp = {"range": f"A{i+2}:A{j+2}", "format": {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}}}
        else:
            temp = {"range": f"A{i+2}:A{j+2}", "format": {"backgroundColor": {"red": 1.0, "green": 0.5, "blue": 0.0}}}
        samenames.append(temp)
        i = j + 1
    else:
        i = i + 1
        toggle = False

#배경색 적용
all.batch_format(samenames)

