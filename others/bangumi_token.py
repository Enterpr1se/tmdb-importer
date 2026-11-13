# 【2026-11-13】Bangumi token 到期提醒。Bangumi 個人 token 最長有效 365 日，API 唔會話幾時到期，
# 所以喺 configs.json 記錄建立日期（bangumi_token_created，YYYY-MM-DD），到期前 30 日起每次啟動都提醒。
# 呢個模組只用標準庫，監控面板都可以 import。
import datetime as dt
import json
import os

TOKEN_VALID_DAYS = 365      # configs.json 可以用 bangumi_token_valid_days 改（例如建立時揀咗較短有效期）
REMIND_DAYS_BEFORE = 30
TOKEN_PAGE = 'https://next.bgm.tv/demo/access-token'
_CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'configs.json')


def token_expiry_message(config_file=None, today=None):
    """快到期 / 已過期就回傳提醒文字，否則回傳 None。冇 token 或者冇記錄日期都回傳 None。"""
    try:
        with open(config_file or _CONFIG_FILE, encoding='utf-8') as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        return None
    if not str(cfg.get('bangumi_token') or '').strip():
        return None
    try:
        created = dt.date.fromisoformat(str(cfg.get('bangumi_token_created') or '').strip())
    except ValueError:
        return None
    try:
        valid_days = int(cfg.get('bangumi_token_valid_days') or TOKEN_VALID_DAYS)
    except (TypeError, ValueError):
        valid_days = TOKEN_VALID_DAYS
    expiry = created + dt.timedelta(days=valid_days)
    remaining = (expiry - (today or dt.date.today())).days
    how = (f"請去 {TOKEN_PAGE} 建立新 token，然後更新 configs.json 嘅 bangumi_token，"
           f"同埋將 bangumi_token_created 改做建立日期（YYYY-MM-DD）。")
    if remaining < 0:
        return f"Bangumi token 已於 {expiry} 過期（建立日期 {created}）。{how}"
    if remaining <= REMIND_DAYS_BEFORE:
        return f"Bangumi token 將於 {expiry} 到期（仲有 {remaining} 日）。{how}"
    return None
