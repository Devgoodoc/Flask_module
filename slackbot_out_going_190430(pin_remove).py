#-*- coding:utf-8 -*-
import json
from flask import Flask, request, make_response
from slacker import Slacker
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from threading import Thread
from slackclient import SlackClient
from pprint import pprint
#============================
app = Flask(__name__)

# 봇 유저 토큰 값
token = 'xoxb-3917885633-580835534055-mO595N6OkPfjgk2nXk6bPj9o'
slack = Slacker(token)

# 처리완료 봇 토큰
token_1 = 'xoxb-3917885633-600487694965-iyxlCIpTrhdspZec3PZ7NnwF'
slack_1 = Slacker(token_1)

# 보류 봇 토큰
token_2 = 'xoxb-3917885633-599974395204-ZyzbupBx6WHzXbN1rIkjXWu4'
slack_2 = Slacker(token_2)

#---
# 구글스프레드 시트 인증
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('cred.json', scope)

gs = gspread.authorize(credentials)     # Key 정보 인증

#---
# CS접수현황 문서 가져오기
doc = gs.open_by_url('https://docs.google.com/spreadsheets/d/15N7K31hkeaqb8Snq7D2U2N6jTjvtLQ5pip4iW1DK4TU/edit?pli=1#gid=0')
ws = doc.get_worksheet(0)

# 병원 접수 도입현황 문서 가져오기
doc_2 = gs.open_by_url('https://docs.google.com/spreadsheets/d/1iRpmebKnV31cfS9xStu8GedjOxKPmObSAnnZaX-M65A/edit?pli=1#gid=1759562169')
ws_2 = doc_2.get_worksheet(0)

#---
message_info = ""
channel = ""
parent_ts = ""
#============================
def get_answer_1():
	return "안녕하세요."

def get_answer_2():
	return "CS 처리결과가 '처리완료' 상태로 변경되었습니다. "

def get_answer_3():
	return "CS 처리결과가 '보류' 상태로 변경되었습니다."
#============================
def update_gs_cell_2():
    global message_info, channel, parent_ts
    #print("message_info: ", message_info)

    print("핀 해제: ", slack.pins.remove(channel, timestamp=parent_ts))

    print("start update gs cell")

    # 구글 시트 재로그인
    if credentials.access_token_expired:
        gs.login()
        print("토큰이 만료되어 재로그인을 수행했습니다.")

    # 슬랙 메시지 추출
    text_value = message_info['event']['text']

    # 슬랙 봇 코드 식별
    get_slack_bot_code = text_value[0:12]
    print(get_slack_bot_code)

    # 슬랙 메시지 값
    slack_message = text_value.split(',')

    # *****
    # cs id가 입력된 실제 행 주소 추출
    cell = ws.find(slack_message[1].strip())
    print(cell)
    cell_2 = str(cell).split(" ")
    print(cell_2)
    cell_3 = cell_2[1]
    print(cell_3)
    cell_4 = cell_3[1:-3]
    print("행 주소: ", cell_4)
    print('---')
    # *****

    if get_slack_bot_code == '<@UHNEBLEUD>':     # 처리완료 봇
        ws.update_acell('K' + cell_4, "처리완료")  # 처리결과 업데이트
        ws.update_acell('M' + cell_4, slack_message[2])  # 처리내용 업데이트
        print("처리완료로 변경")

    elif get_slack_bot_code == '<@UHMUNBM60>':     # 보류 봇
        ws.update_acell('K' + cell_4, "보류")  # 처리결과 업데이트
        ws.update_acell('M' + cell_4, slack_message[2])  # 처리내용 업데이트
        print("보류로 변경")

#============================
def event_handler(event_type, slack_event):
    global message_info, channel, parent_ts

    # 메세지 발송 채널
    channel = slack_event["event"]["channel"]

    # 슬랙 메시지 타임스탬프
    thread_ts = slack_event['event']['ts']

    parent_ts = slack_event['event']['thread_ts']

    # 슬랙 메시지 추출
    text_value = slack_event['event']['text']
    # 슬랙 봇 코드 식별
    get_slack_bot_code = text_value[0:12]

    #print("핀 해제: ", slack.pins.remove(channel, timestamp=parent_ts))

    # 스프레드 셀 업데이트 쓰레드 실행
    # 정석은 call_me() 함수를 실행해서 등록 완료를 리턴값으로 받아서 해당 값 기준으로 슬랙 메시지를 보내는게 맞지만, 로직을 다시 재구성해야 함.
    call_me()

    if event_type == "app_mention" and get_slack_bot_code == '<@UHNEBLEUD>':     # 처리완료 봇
        #channel = slack_event["event"]["channel"]
        text = get_answer_2()
        print("슬랙 메시지 전송: ", slack_1.chat.post_message(channel, text, thread_ts=thread_ts))
        return make_response("앱 멘션 메시지가 보내졌습니다.", 200, )

    if event_type == "app_mention" and get_slack_bot_code == '<@UHMUNBM60>':     # 보류 봇
        #channel = slack_event["event"]["channel"]
        text = get_answer_3()
        print("슬랙 메시지 전송: ", slack_2.chat.post_message(channel, text, thread_ts=thread_ts))
        return make_response("앱 멘션 메시지가 보내졌습니다.", 200, )

    message = "[%s] 이벤트 핸들러를 찾을 수 없습니다." % event_type
    return make_response(message, 200, {"X-Slack-No-Retry": 1})

#============================
@app.route("/slack", methods=["GET", "POST"])
def hears():
    global message_info

    slack_event = json.loads(request.data)
    print("이벤트 수신")
    pprint(slack_event)
    print('---')

    message_info = slack_event

    if slack_event['event']['type'] == 'pin_removed':
        print("핀이 제거된 메세지")
        return 'done'

    if "challenge" in slack_event:
        return make_response(slack_event["challenge"], 200, {"content_type": "application/json"})

    if "event" in slack_event:
        event_type = slack_event["event"]["type"]
        return event_handler(event_type, slack_event)

    return make_response("슬랙 요청에 이벤트가 없습니다.", 404, {"X-Slack-No-Retry": 1})

#============================
def call_me():                                              # 봇 호출 후, 시트 업데이트 시 담아야 할 이벤트
    create_thread()                                         # 쓰레드 메서드 호출
    print("call thread done")
    print('---')

def create_thread():
    run_thread = Thread(target=update_gs_cell_2)           # 메서드 대상 지정
    print("create thread")
    print('---')
    run_thread.setDaemon(True)
    run_thread.start()
    print('셀 업데이트 쓰레드 시작 : ', run_thread)
    print('---')

#============================
@app.route('/')
def hello_world():
    return 'Hello World!'

#============================
@app.route('/dave')
def hello_dave():
	return 'Very Nice!'

#============================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)
