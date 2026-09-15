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
    # 自動去除貼上時可能夾帶的前後空白／換行
    raw = refresh_token
    refresh_token = refresh_token.strip()
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
    if r.status_code != 200:
        raise RuntimeError(
            f"Dropbox 換發 access token 失敗 HTTP {r.status_code}：{r.text[:300]} "
            f"| 收到的 token 開頭={refresh_token[:4]!r}、長度={len(refresh_token)}、"
            f"前後含空白或換行={raw != refresh_token}（refresh token 應為 sCTL 開頭，"
            f"若為 sl.u 開頭代表貼成 4 小時過期的 access token）"
        )
    access_token = r.json()["access_token"]
    return dropbox.Dropbox(access_token)
