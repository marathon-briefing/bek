#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
雲端版學員每日通知（GitHub Actions 用）
從 Dropbox API 讀 xlsx，寄 email 給每位學員，CC 教練。
"""
import os
import re
import smtplib
from io import BytesIO
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.header import Header

import dropbox
from openpyxl import load_workbook

from dropbox_auth import get_dbx
SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASS = os.environ["SMTP_PASS"]
COACH_EMAIL = "runningcoach.kevin@gmail.com"
BASE = "/教練業務管理"

CIVIL_DIR = f"{BASE}/學員專用/國考體測學員"
MARATHON_DIR = f"{BASE}/學員專用/馬拉松學員"

STUDENTS = [
    {"編號": "M26001", "姓名": "王奕翔", "email": "mike141431@gmail.com",
     "日誌": f"{MARATHON_DIR}/M26001-王奕翔/M26001-王奕翔_馬拉松訓練日誌.xlsx",
     "課表": f"{MARATHON_DIR}/M26001-王奕翔/M26001-王奕翔_馬拉松課表.xlsx"},
    {"編號": "M26002", "姓名": "林洋樂", "email": "alerler0817@gmail.com",
     "日誌": f"{MARATHON_DIR}/M26002-林洋樂/M26002-林洋樂_馬拉松訓練日誌.xlsx",
     "課表": f"{MARATHON_DIR}/M26002-林洋樂/M26002-林洋樂_馬拉松課表.xlsx"},
    {"編號": "T26001", "姓名": "阮筱軒", "email": "syuan000906@gmail.com",
     "日誌": f"{CIVIL_DIR}/T26001-阮筱軒/T26001-阮筱軒_國考訓練日誌.xlsx",
     "課表": f"{CIVIL_DIR}/T26001-阮筱軒/T26001-阮筱軒_國考課表.xlsx"},
    {"編號": "T26002", "姓名": "盧冠婷", "email": "tina19981217@gmail.com",
     "日誌": f"{CIVIL_DIR}/T26002-盧冠婷/T26002-盧冠婷_國考訓練日誌.xlsx",
     "課表": f"{CIVIL_DIR}/T26002-盧冠婷/T26002-盧冠婷_國考課表.xlsx"},
]

WEEKDAY_MAP = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}

_db = None
def dbx():
    global _db
    if _db is None:
        _db = get_dbx()
    return _db

def read_xlsx(path):
    try:
        _, res = dbx().files_download(path)
        return load_workbook(BytesIO(res.content), data_only=True)
    except Exception as e:
        print(f"⚠ 讀取失敗 {path}: {e}")
        return None

def read_schedule(path, target_date):
    wb = read_xlsx(path)
    if not wb or "每日課表" not in wb.sheetnames:
        if wb: wb.close()
        return None
    ws = wb["每日課表"]
    patterns = [
        target_date.strftime("%-m/%-d"),
        target_date.strftime("%m/%d"),
        target_date.strftime("%Y-%m-%d"),
    ]
    content_col = 4
    for r in range(1, min(ws.max_row+1, 10)):
        vals = [ws.cell(r,c).value for c in range(1, ws.max_column+1)]
        if any("日期" in str(v) for v in vals if v):
            for c, v in enumerate(vals, 1):
                if str(v) in ("訓練內容", "內容"):
                    content_col = c
            break
    for r in range(1, ws.max_row+1):
        for c in range(1, min(ws.max_column+1, 5)):
            v = ws.cell(r,c).value
            if v and any(p in str(v) for p in patterns):
                content = ws.cell(r, content_col).value
                wb.close()
                return str(content).strip() if content else None
    wb.close()
    return None

def read_log(path, target_date):
    wb = read_xlsx(path)
    if not wb: return None
    ws = wb.active
    patterns = [target_date.strftime("%-m/%-d"), target_date.strftime("%Y-%m-%d")]
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, 1).value
        if v and any(p in str(v) for p in patterns):
            # C=課表內容 D=距離 E=時間 F=平均心率 G=最高心率 H=教練評註
            record = {
                "距離": ws.cell(r, 4).value,
                "時間": ws.cell(r, 5).value,
                "平均心率": ws.cell(r, 6).value,
                "最高心率": ws.cell(r, 7).value,
                "教練評註": ws.cell(r, 8).value,
            }
            wb.close()
            return record
    wb.close()
    return None

def get_first_name(full_name):
    if not full_name or len(full_name) <= 1:
        return full_name
    return full_name[1:]

def format_content(content):
    if not content: return content
    content = str(content).strip()
    pace_match = re.search(r'[（(]([\d:–\-]+/km)[)）]', content)
    pace_info = pace_match.group(1) if pace_match else None
    if pace_info:
        content = re.sub(r'[（(][\d:–\-]+/km[)）]', '', content).strip()
    return content

def build_email(student, today):
    name = student["姓名"]
    first = get_first_name(name)
    yesterday = today - timedelta(days=1)
    tomorrow = today + timedelta(days=1)

    today_s = read_schedule(student["課表"], today) or "休息日或輕鬆恢復"
    today_s = format_content(today_s)

    tomorrow_s = read_schedule(student["課表"], tomorrow) or "休息日或輕鬆恢復"
    tomorrow_s = format_content(tomorrow_s)

    log = read_log(student["日誌"], yesterday)
    if log and log.get("距離"):
        parts = [f"距離 {log['距離']} km"]
        if log.get("時間"): parts.append(f"時間 {log['時間']}")
        if log.get("平均心率"): parts.append(f"心率 {log['平均心率']}")
        yest = "｜".join(parts)
        if log.get("教練評註"): yest += f"\n教練評註：{log['教練評註']}"
    else:
        yest = "昨日無訓練記錄\n如尚未回傳，請記得儘速補上"

    weekday = WEEKDAY_MAP[today.weekday()]
    subject = f"Hi {first}，今日訓練安排"

    body = f"""早安，{first}：

【今天 {today.strftime('%-m/%-d')}（{WEEKDAY_MAP[today.weekday()]}）】
{today_s}

【昨日訓練】
{yest}

【明天 {tomorrow.strftime('%-m/%-d')}（{WEEKDAY_MAP[tomorrow.weekday()]}）】
{tomorrow_s}

提醒：
1. 課表內容請照表執行，有問題隨時跟我說
2. 訓練後請用手機截圖資料畫面回傳
3. 身體不適或有任何狀況，請主動回報
4. 課表內容請勿隨意更換或改課，如有需要改動，請回報教練調整

加油！

執行教練｜Kevin Chang
0917060888｜Line: kc1225888
"""
    return subject, body

def send(to, subject, body):
    msg = MIMEText(body, "plain", "utf-8")
    msg["From"] = SMTP_USER
    msg["To"] = to
    msg["Cc"] = COACH_EMAIL
    msg["Subject"] = Header(subject, "utf-8")
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(SMTP_USER, [to, COACH_EMAIL], msg.as_string())
    print(f"✓ → {to} (CC {COACH_EMAIL})")

if __name__ == "__main__":
    today = datetime.now().date()
    for st in STUDENTS:
        try:
            subject, body = build_email(st, today)
            send(st["email"], subject, body)
        except Exception as e:
            print(f"✗ {st['姓名']}: {e}")
