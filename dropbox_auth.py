"""Dropbox 認證 helper：用 refresh token 自動換新 access token。"""
import os
import requests

APP_KEY = "v6uthgfsykttyx8"
APP_SECRET = "khiimn2sroe5vub"

def get_dbx():
    import dropbox
    # 相容兩種環境變數名稱：DROPBOX_REFRESH_TOKEN（正式）與 DROPBOX_TOKEN（舊 yml）
    refresh_token = (
        os.environ.get("DROPBOX_REFRESH_TOKEN")
        or os.environ.get("DROPBOX_TOKEN")
    )
    if not refresh_token:
        raise RuntimeError(
            "缺少 Dropbox refresh token：請在環境變數設定 DROPBOX_REFRESH_TOKEN"
        )
    # 用 refresh token 換新 access token
    r = requests.post(
        "https://api.dropbox.com/oauth2/token",
        auth=(APP_KEY, APP_SECRET),
        data={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        },
        timeout=30,
    )
    r.raise_for_status()
    access_token = r.json()["access_token"]
    return dropbox.Dropbox(access_token)
