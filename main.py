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
    # 各項目を正規表現で抽出
    title_match = re.search(r'【タイトル】(.+)', message)
    date_match = re.search(r'【日付】(\d{1,2})[/-](\d{1,2})', message)
    start_time_match = re.search(r'【開始時間】(\d{1,2}):(\d{2})', message)
    content_match = re.search(r'【内容】(.+)', message)
    url_match = re.search(r'【URL】(.+)', message)

    if not (title_match and date_match and start_time_match):
        return None

    title = title_match.group(1).strip()
    month = int(date_match.group(1))
    day = int(date_match.group(2))
    hour = int(start_time_match.group(1))
    minute = int(start_time_match.group(2))

    content = content_match.group(1).strip() if content_match else ""
    url = url_match.group(1).strip() if url_match else ""

    # 日付時刻をTokyo timezoneで作成
    year = datetime.now(JST).year
    naive_dt = datetime(year, month, day, hour, minute)
    dt = JST.localize(naive_dt)
    start_str = dt.isoformat()
    end_str = (dt + timedelta(hours=1)).isoformat()

    # 説明は内容＋URLを結合して作成
    description = content
    if url:
        description += "\nURL: " + url

    return title, start_str, end_str, description


def add_event(summary, start_time_str, end_time_str, description=None):
    event = {
        'summary': summary,
        'start': {'dateTime': start_time_str, 'timeZone': 'Asia/Tokyo'},
        'end': {'dateTime': end_time_str, 'timeZone': 'Asia/Tokyo'}
    }
    if description:
        event['description'] = description
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
    if "【タイトル】" not in message_text or "【日付】" not in message_text or "【開始時間】" not in message_text:
        return None

    parsed = extract_event_info(message_text)
    if not parsed:
        return "予定の形式が正しくありません。例:\n【タイトル】会議\n【日付】7/10\n【開始時間】14:00\n【内容】説明\n【URL】https://..."

    title, start_str, end_str, description = parsed
    add_event(title, start_str, end_str, description=description)
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
    lines.append("\nご参加ご希望の方は予定表に調整さんがあればそちらから出欠のご連絡をお願いします。")  # ← 注意文追加
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