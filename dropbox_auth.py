"""Dropbox 認證 helper：用 refresh token 自動換新 access token。"""
import os
import requests

APP_KEY = "v6uthgfsykttyx8"
APP_SECRET = "khiimn2sroe5vub"

def get_dbx():
    import dropbox
    refresh_token = os.environ["DROPBOX_REFRESH_TOKEN"]
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
