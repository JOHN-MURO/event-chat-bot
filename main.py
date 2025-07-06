# -*- coding: utf-8 -*-

from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from linebot.exceptions import InvalidSignatureError
from google.oauth2 import service_account
from googleapiclient.discovery import build
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta
from pytz import timezone
import re
import os
import json

app = Flask(__name__)

# 環境変数
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')
GOOGLE_CALENDAR_ID = os.getenv("GOOGLE_CALENDAR_ID")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
LINE_GROUP_ID = os.getenv("LINE_GROUP_ID")

# LINE SDK
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Calendar 認証
SCOPES = ['https://www.googleapis.com/auth/calendar']
credentials_info = json.loads(GOOGLE_CREDENTIALS_JSON)
creds = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
calendar_service = build('calendar', 'v3', credentials=creds)

JST = timezone('Asia/Tokyo')

def extract_event_info(message):
    message = message.replace("【予定】", "").strip()
    match = re.search(r'(\d{1,2})[/-](\d{1,2})\s+(\d{1,2})(?::(\d{2}))?\s*(.+)', message)
    if not match:
        return None
    month, day, hour, minute, title = match.groups()
    minute = minute or '00'
    year = datetime.now(JST).year

    # 修正前: dt = datetime(year, int(month), int(day), int(hour), int(minute), tzinfo=JST)
    naive_dt = datetime(year, int(month), int(day), int(hour), int(minute))
    dt = JST.localize(naive_dt)

    return title, dt.isoformat(), (dt + timedelta(hours=1)).isoformat()

def add_event(summary, start_time_str, end_time_str):
    event = {
        'summary': summary,
        'start': {'dateTime': start_time_str, 'timeZone': 'Asia/Tokyo'},
        'end': {'dateTime': end_time_str, 'timeZone': 'Asia/Tokyo'}
    }
    calendar_service.events().insert(calendarId=GOOGLE_CALENDAR_ID, body=event).execute()

def get_events_between(start_dt, end_dt):
    result = calendar_service.events().list(
        calendarId=GOOGLE_CALENDAR_ID,
        timeMin=start_dt.isoformat(),  # JST時間をそのまま使う
        timeMax=end_dt.isoformat(),
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    return result.get('items', [])

def handle_incoming_message(message_text):
    # 「【予定】」というキーワードがない場合は無視
    if "【予定】" not in message_text:
        return None

    parsed = extract_event_info(message_text)
    if not parsed:
        return "予定の形式が正しくありません。例: 【予定】 7/10 14:00 会議"

    title, start_str, end_str = parsed
    add_event(title, start_str, end_str)
    return f"予定を登録しました：{title}（{start_str}）"

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    reply = handle_incoming_message(event.message.text)
    if reply:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

def format_events(events, header):
    if not events:
        return header + "\n予定はありません。"
    lines = [header]
    for e in events:
        start = e['start'].get('dateTime', '')[11:16]
        lines.append(f"{start} - {e['summary']}")
    return '\n'.join(lines)

def notify_week_events(bot):
    today = datetime.now(JST)
    start = JST.localize(datetime(today.year, today.month, today.day, 0, 0, 0))
    end_date = today + timedelta(days=6 - today.weekday())
    end = JST.localize(datetime(end_date.year, end_date.month, end_date.day, 23, 59, 59))
    events = get_events_between(start, end)
    msg = format_events(events, "【今週の予定】")
    bot.push_message(LINE_GROUP_ID, TextSendMessage(text=msg))

def notify_tomorrow_events(bot):
    tomorrow = datetime.now(JST) + timedelta(days=1)
    start = JST.localize(datetime(tomorrow.year, tomorrow.month, tomorrow.day))
    end = start + timedelta(days=1)
    events = get_events_between(start, end)
    msg = format_events(events, "【明日の予定】")
    bot.push_message(LINE_GROUP_ID, TextSendMessage(text=msg))

def start_scheduler(line_bot_api):
    scheduler = BackgroundScheduler(timezone=JST)
    scheduler.add_job(lambda: notify_week_events(line_bot_api), 'cron', day_of_week='mon', hour=8)
    scheduler.add_job(lambda: notify_tomorrow_events(line_bot_api), 'cron', hour=20)
    scheduler.start()

@app.route("/")
def index():
    return "LINE Google Calendar Bot is running!"

@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers["X-Line-Signature"]
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

#if __name__ == "__main__":
#   port = int(os.environ.get("PORT", 5000))  # HerokuのPORTを取得
#    app.run(host="0.0.0.0", port=port)