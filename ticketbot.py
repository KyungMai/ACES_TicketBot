import gspread

if __name__ == "__main__":

    #시트 관리할 봇의 정보가 담긴 json파일 입력
    gc = gspread.service_account(filename = "credentials.json")

    #티켓 관리용 스프레드시트 링크 입력
    sh = gc.open_by_url("")

    #시트 목록 가져오기
    worksheet_list = sh.worksheets()

    #무료 티켓 관리 시트 가져옴
    free = worksheet_list[0]

    #유로 티켓 관리 시트 가져옴
    pay = worksheet_list[1]

    #전체 티켓 관리 시트 초기화
    try:
        sh.del_worksheet(worksheet_list[2])
    except IndexError:
        pass
    finally:
        all = sh.add_worksheet(title = "전체 티켓", rows = 0, cols = 0)

    #전체 티켓 관리용 리스트 [[이름, 연락처, 티켓 매수], ...] 형식
    tickets = list()

    #무료티켓 수합 결과에서 이름, 연락처, 필요 매수 가져옴
    for i in free.get_all_records():
        tickets.append(list(i.values())[1:4])

    #유료티켓 수합 결과에서 이름, 연락처, 구매 매수 가져옴
    for i in pay.get_all_records():
        tickets.append(list(i.values())[1:4])

    #정렬을 위해 int와 str이 섞여있는 전화번호의 형식 맞춤
    for i in tickets:
        i[1] = str(i[1])

    #이름의 사전순으로 tickets를 정렬하되 한글이 다른 언어보다 먼저 오게 함
    #이름이 같은 경우 전화번호 순으로 정렬
    tickets.sort(key = lambda x: (x[0] if ord('ㄱ') <= ord(x[0][0]) <= ord('힣') else '힣' + x[0], x[1]))

    #전체 티켓 관리용 시트에 무료 티켓 수합 결과와 유료 티켓 수합 결과 합쳐서 작성
    all.batch_update([{"range":f"B1:D{len(tickets)}", "values": tickets}])

    #이름이 같은 사람들의 ticket에서의 index를 저장하는 리스트
    samename = list()

    #배경색 일괄 적용을 위한 리스트
    sheet_format = list()

    #하나의 동명인 그룹 앞에 다른 이름의 동명인 그룹이 있었는지 확인하는 변수
    toggle = False

    #이름이 같은 사람들의 tickets에서의 index를 samename리스트에 저장
    #동명인이면 맨 앞 행의 배경을 노란색으로 칠함
    #서로 다른 이름의 동명인 그룹들이 연속된 경우 노란색과 주황색을 번갈아가며 칠함
    i = 0
    while i < len(tickets) - 1:
        if tickets[i][0] == tickets[i + 1][0]:
            toggle = not toggle
            j = i + 1
            while tickets[j][0] == tickets[j + 1][0] and j < len(tickets):
                j += 1
            if toggle:

                sheet_format.append({"range": f"A{i + 1}:A{j + 1}",
                                    "format": {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}}}) #노란색
            else:
                sheet_format.append({"range": f"A{i + 1}:A{j + 1}",
                                    "format": {"backgroundColor": {"red": 1.0, "green": 0.6, "blue": 0.0}}}) #주황색
            samename.append(range(i,j+1))
            i = j + 1
        else:
            i = i + 1
            toggle = False


    #동일인 관리용 리스트 {[동일인 시작 index, 동일인 끝 index]: 동일인의 티켓 총 매수, ...} 형식
    sameperson = dict()

    #동명인 중 전화번호가 같은 사람은 동일인으로 판단
    #동일인의 tickets에서의 index를 sameperson의 key로, 동일인이 필요로하는 티켓의 총 매수를 value로 저장
    #동일인이면 마지막 행들의 배경을 초록색으로 칠함
    for i in samename:
        j = 0
        while j < len(i) - 1:
            if tickets[i[j]][1] == tickets[i[j + 1]][1]:
                k = j + 1
                while tickets[i[k]][1] != tickets[i[k]][1] and k < len(i) - 1:
                    k += 1
                sameperson[(i[j], i[k])] = sum(int(sublist[2].split("매")[0]) for sublist in tickets[i[j]: i[k] + 1])
                sheet_format.append({"range": f"E{i[j] + 1}:E{i[k] + 1}",
                                    "format": {"backgroundColor": {"red": 0.0, "green": 1.0, "blue": 0.0}}}) #초록색
                j = k + 1
            else:
                j += 1

    #배경색 변경사항 일괄 적용
    all.batch_format(sheet_format)

    #초록색으로 칠한 동일인의 마지막 행들을 병합
    #병합한 셀에 동일인이 필요로 하는 총 티켓 매수를 작성
    for i in sameperson.keys():
        all.merge_cells(f"E{i[0] + 1}:E{i[1] + 1}")
        all.update([[f"총 {sameperson[i]}매"]], f"E{i[0] + 1}")
