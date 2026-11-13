# 【2026-11】Bangumi（bgm.tv）人物中文名：日韓人 TMDB 冇任何中文 / 漢字名、維基又搵唔到時用。
# 實測（2026-11-13）：漢字名嘅結果同本機新字體轉換一樣；真正有用嘅係假名名（阪口あや → 阪口彩）。
# 只接受「Bangumi 主名 == 日文名」（唔用別名，別名會配錯人），有生日就要吻合，而且只有一個中文名。
# 結果係简体中文名，交返主流程當 zh-CN 候選（會經 s2t / s2hk）。快取喺 wiki_cache.db（media_type person_bangumi）。
import json
import logging
import os
import re
import threading
import time

import requests

from others import utils
from others.wiki_search import _init_db, _save_to_wiki_cache, _wikidata_cache_lookup

BANGUMI_CACHE_TYPE = 'person_bangumi'
_SEARCH_URL = 'https://api.bgm.tv/v0/search/persons'
_USER_AGENT = 'tmdb-importer (https://github.com/Enterpr1se/tmdb-importer)'
_MIN_INTERVAL = 1.0          # 每秒最多一個請求
_CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'configs.json')

_lock = threading.Lock()     # 全部請求逐個做
_last_request = 0.0
_token = None
_token_loaded = False


def _get_token():
    global _token, _token_loaded
    if not _token_loaded:
        _token_loaded = True
        try:
            with open(_CONFIG_FILE, encoding='utf-8') as f:
                _token = (json.load(f).get('bangumi_token') or '').strip() or None
        except (OSError, ValueError) as e:
            logging.warning(f"[Bangumi] 讀取 configs.json 失敗（{e}），唔用 token。")
        if not _token:
            logging.info("[Bangumi] configs.json 冇 bangumi_token，以匿名方式查詢。")
    return _token


def _norm(s):
    return re.sub(r'[\s　・·]', '', str(s or ''))


def _has_japanese_script(s):
    return bool(s and re.search(r'[぀-ヿ㐀-鿿]', s))


def _birthday_matches(person, birthday):
    """TMDB 生日（YYYY-MM-DD）同 Bangumi 有填嘅部分要一致；任何一邊冇就當吻合。"""
    m = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', str(birthday or ''))
    if not m:
        return True
    y, mo, d = map(int, m.groups())
    return (person.get('birth_year') in (None, y) and person.get('birth_mon') in (None, mo)
            and person.get('birth_day') in (None, d))


def _search(keyword):
    """回傳搜尋結果 list；429 / 出錯回傳 None（唔寫快取，下次再試）。"""
    global _last_request
    headers = {'User-Agent': _USER_AGENT}
    token = _get_token()
    if token:
        headers['Authorization'] = f'Bearer {token}'
    with _lock:
        wait = _MIN_INTERVAL - (time.monotonic() - _last_request)
        if wait > 0:
            time.sleep(wait)
        try:
            r = requests.post(_SEARCH_URL, params={'limit': 10}, json={'keyword': keyword},
                              headers=headers, timeout=20)
        except requests.RequestException as e:
            logging.warning(f"[Bangumi] 搜尋 '{keyword}' 失敗: {e}")
            return None
        finally:
            _last_request = time.monotonic()
    if r.status_code == 429:
        logging.warning(f"[Bangumi] 搜尋 '{keyword}' 回覆 429（限速）。")
        return None
    if r.status_code == 401:
        logging.warning("[Bangumi] token 無效（401），可能已過期，請檢查 configs.json 嘅 bangumi_token / bangumi_token_created。")
        return None
    if not r.ok:
        logging.warning(f"[Bangumi] 搜尋 '{keyword}' 回覆 {r.status_code}。")
        return None
    try:
        return r.json().get('data') or []
    except ValueError:
        return None


def lookup_zh_name(item_id, japanese_names, birthday=None, original_name=None):
    """用日文名喺 Bangumi 搵人物嘅简体中文名。搵唔到 / 唔肯定回傳 None。"""
    _init_db()
    cached, need_query = _wikidata_cache_lookup(item_id, BANGUMI_CACHE_TYPE) if item_id else (None, True)
    if cached or not need_query:
        logging.info(f"[Bangumi] 快取命中 '{original_name}': {cached or '冇'}")
        return cached

    keys = list(dict.fromkeys(_norm(n) for n in japanese_names if _has_japanese_script(n)))[:3]
    keys = [k for k in keys if k]
    if not keys:
        return None

    zh_name, matched_key = None, None
    for key in keys:
        if utils._PROGRAM_INTERRUPTED:
            return None
        results = _search(key)
        if results is None:
            return None   # 網絡 / 限速問題：唔寫快取
        exact = [p for p in results if _norm(p.get('name')) == key and _birthday_matches(p, birthday)]
        names = {next((x.get('value') for x in (p.get('infobox') or [])
                       if x.get('key') == '简体中文名' and isinstance(x.get('value'), str)), '').strip()
                 for p in exact} - {''}
        if len(names) == 1:
            zh_name, matched_key = names.pop(), key
            break
        if len(names) > 1:
            logging.info(f"[Bangumi] '{key}' 有多個同名人物（{'、'.join(sorted(names))}），唔使用。")
            break

    if zh_name and not utils.is_chinese(zh_name):
        zh_name = None
    if zh_name:
        logging.info(f"[Bangumi] '{original_name}'（日文名 '{matched_key}'）中文名: '{zh_name}'")
    else:
        logging.info(f"[Bangumi] '{original_name}' 搵唔到中文名（搜尋: {'、'.join(keys)}）。")
    if item_id:
        _save_to_wiki_cache(item_id, BANGUMI_CACHE_TYPE, zh_name, 'FOUND' if zh_name else 'NOT_FOUND',
                            original_name, None)
    return zh_name
