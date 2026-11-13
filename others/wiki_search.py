import logging
import re
import requests
from bs4 import BeautifulSoup
import wikipediaapi
from .utils import is_chinese, normalize_to_hk, is_all_hanzi
from collections import Counter
import time
from urllib.parse import quote
from SPARQLWrapper import SPARQLWrapper, JSON
import sqlite3
import os
from datetime import datetime
# 【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import traceback

# Wikimedia 要求 User-Agent 寫明工具名稱同聯絡方式（網址），否則限速更嚴
WIKI_USER_AGENT = 'TMDB-HK-Name-Extractor/2.3 (https://github.com/Enterpr1se/tmdb-importer)'

# --- Rate-limit signal for wiki_queue adaptive delay ---
_search_hit_rate_limit: bool = False

def _mark_rate_limited():
    global _search_hit_rate_limit
    _search_hit_rate_limit = True

def _reset_rate_limit_flag():
    global _search_hit_rate_limit
    _search_hit_rate_limit = False

# --- 【新增】Wiki 查詢快取資料庫處理 ---

DB_FILE = 'data/wiki_cache.db'
_in_memory_cache = {}

def _init_db():
    """
    【簡化版 2025-08-25】初始化資料庫。
    此版本假設用戶總是從一個新的、不存在的資料庫開始，因此移除了處理舊資料庫結構的遷移邏輯。
    """
    try:
        db_dir = os.path.dirname(DB_FILE)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir)
            logging.info(f"已創建資料庫資料夾: {db_dir}")

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        # 直接創建最新結構的資料表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS wiki_cache (
                query_key TEXT PRIMARY KEY,
                item_id TEXT,
                original_name TEXT NOT NULL,
                media_type TEXT NOT NULL,
                year TEXT,
                hk_name TEXT,
                status TEXT NOT NULL, -- 'FOUND', 'NOT_FOUND', 'FAILED'
                timestamp DATETIME NOT NULL
            )
        ''')

        # 建立索引
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_item_id_type ON wiki_cache (item_id, media_type)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_name_type_year ON wiki_cache (original_name, media_type, year)')
        
        conn.commit()
        conn.close()
    except Exception as e:
        logging.error(f"初始化 wiki 快取資料庫時發生錯誤: {e}")
        logging.error(traceback.format_exc())

def _get_query_key(item_id, media_type):
    """【修改】根據 item_id 和 media_type 創建唯一的查詢鍵。"""
    return f"{media_type}-{item_id}"

def _query_wiki_cache(item_id, media_type):
    """
    【修改】使用 item_id 和 media_type 查詢快取。
    """
    if not item_id or not media_type:
        return None
        
    query_key = _get_query_key(item_id, media_type)

    if query_key in _in_memory_cache:
        logging.info(f"Wiki 快取命中 (記憶體) for key: '{query_key}'")
        return _in_memory_cache[query_key]

    if not os.path.exists(DB_FILE):
        return None

    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        # 使用新的索引進行查詢
        cursor.execute(
            "SELECT hk_name, status FROM wiki_cache WHERE item_id = ? AND media_type = ?",
            (str(item_id), media_type)
        )
        result = cursor.fetchone()
        conn.close()

        if result:
            hk_name, status = result
            logging.info(f"Wiki 快取命中 (資料庫) for key: '{query_key}'. 狀態: {status}, 名稱: '{hk_name}'")
            cached_result = {'hk_name': hk_name, 'status': status}
            _in_memory_cache[query_key] = cached_result
            return cached_result
        else:
            return None
    except Exception as e:
        logging.error(f"查詢 wiki 快取時發生錯誤 for key '{query_key}': {e}")
        return None

def _save_to_wiki_cache(item_id, media_type, hk_name, status, original_name, year):
    """
    【修改】將包含 item_id 的搜索結果儲存到資料庫。
    """
    if not item_id or not media_type:
        logging.warning(f"嘗試儲存快取失敗，缺少 item_id 或 media_type。ID: {item_id}, Type: {media_type}")
        return

    query_key = _get_query_key(item_id, media_type)
    timestamp = datetime.now()

    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(
            '''
            INSERT OR REPLACE INTO wiki_cache (query_key, item_id, original_name, media_type, year, hk_name, status, timestamp)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''',
            (query_key, str(item_id), original_name, media_type, str(year or 'None'), hk_name, status, timestamp)
        )
        conn.commit()
        conn.close()
        logging.info(f"已儲存至 wiki 快取: Key='{query_key}', Status='{status}', Name='{hk_name}'")

        cached_result = {'hk_name': hk_name, 'status': status}
        _in_memory_cache[query_key] = cached_result
    except Exception as e:
        logging.error(f"儲存至 wiki 快取時發生錯誤 for key '{query_key}': {e}")

def _interruptible_sleep(seconds):
    """Sleep in 1-second chunks so Ctrl+C / _PROGRAM_INTERRUPTED can break out early.
    Returns True if interrupted before the full duration elapsed, False otherwise."""
    try:
        from others import utils as _u
        check = lambda: _u._PROGRAM_INTERRUPTED
    except Exception:
        check = lambda: False
    for _ in range(max(1, int(seconds))):
        if check():
            return True
        time.sleep(1)
    return False


def _create_retry_session():
    """【新增】為 wiki_search.py 創建一個配置了重試機制的 requests Session 物件"""
    session = requests.Session()
    # 429 不放入 status_forcelist：讓 while True 迴圈自行處理並輸出日誌，
    # 避免 urllib3 指數退避在 session.get() 內部靜默等待長達數十分鐘。
    retry_strategy = Retry(
        total=5,
        status_forcelist=[500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"],
        backoff_factor=0.5
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

# 為不同語言的維基百科建立連線器
WIKI_CONNECTORS = {}

def get_wiki_connector(lang='en'):
    """
    【修改】動態獲取或創建指定語言的維基百科連線器。
    此版本修正了中文變體域名的問題，統一將 zh-XX 指向 'zh'。
    """
    # 【使用者要求】新增檢查，如果語言代碼是無效的 'xx'，則直接返回 None，避免程式因域名解析錯誤而崩潰。
    if lang == 'xx':
        logging.warning(f"偵測到無效的語言代碼 'xx'，將跳過此語言的維基百科連線器創建。")
        return None

    api_lang = 'zh' if lang.startswith('zh-') else lang
    
    if api_lang not in WIKI_CONNECTORS:
        logging.info(f"Creating new Wikipedia connector for language: '{api_lang}' (requested: '{lang}')")
        WIKI_CONNECTORS[api_lang] = wikipediaapi.Wikipedia(
            language=api_lang,
            extract_format=wikipediaapi.ExtractFormat.WIKI,
            user_agent=WIKI_USER_AGENT
        )
    return WIKI_CONNECTORS[api_lang]

def _get_hk_name_via_langlink(lang, search_term, media_type, year):
    """
    【全新輔助函數 2025-08-27】
    使用者要求：針對 en, ja, ko 原始語言的影片，從對應語言的維基百科開始搜索，
    透過頁面間的語言連結 (langlink) 找到中文維基頁面，最後提取香港繁體 (zh-hk) 變體標題。
    【修改】加入無限重試機制。
    """
    log = logging.getLogger() # 使用根 logger
    session = _create_retry_session()
    headers = {'User-Agent': WIKI_USER_AGENT}
    
    # 【使用者要求】新增檢查，如果語言代碼是無效的 'xx'，則直接返回 None，避免程式因域名解析錯誤而崩潰。
    if lang == 'xx':
        log.warning(f"[_get_hk_name_via_langlink] 偵測到無效的語言代碼 'xx'，將跳過在 '{lang}' 維基百科上對 '{search_term}' 的 langlink 搜索。")
        return None

    api_lang_host = 'zh' if lang.startswith('zh-') else lang
    api_url = f"https://{api_lang_host}.wikipedia.org/w/api.php"
    log.info(f"[_get_hk_name_via_langlink] 正在 '{lang}' 維基上為 '{search_term}' 啟動 langlink 搜索...")

    try:
        # 步驟 1 & 2: 搜索並獲取最終頁面標題和 langlinks
        params = {
            'action': 'query', 'format': 'json', 'redirects': 1, 'titles': search_term,
            'prop': 'info|pageprops|langlinks', 'inprop': 'url', 'ppprop': 'wikibase_item', 'lllang': 'zh', 'lllimit': 1
        }
        
        # 【修改】加入無限重試循環
# +       response = None
# +       while True:
# +           try:
# +               response = session.get(api_url, params=params, headers=headers, timeout=15)
# +               if response.status_code == 200:
# +                   break
# +               elif response.status_code == 404:
# +                   log.info(f"[_get_hk_name_via_langlink] 資源不存在 (404) - {search_term}")
# +                   return None
# +               else:
# +                   log.warning(f"[_get_hk_name_via_langlink] 請求失敗 ({response.status_code})，3秒後重試...")
# +                   time.sleep(3)
# +           except Exception as e:
# +               log.warning(f"[_get_hk_name_via_langlink] 網絡錯誤 ({e})，3秒後重試...")
# +               time.sleep(3)

# -       response = session.get(api_url, params=params, headers=headers, timeout=15)
# -       response.raise_for_status()

        # 這裡為了方便複製，直接給出整合後的代碼
        response = None
        while True:
            try:
                response = session.get(api_url, params=params, headers=headers, timeout=15)
                if response.status_code == 200:
                    break
                elif response.status_code == 404:
                    log.info(f"[_get_hk_name_via_langlink] 資源不存在 (404) - {search_term}")
                    return None
                else:
                    if response.status_code == 429:
                        _mark_rate_limited()
                    log.warning(f"[_get_hk_name_via_langlink] 請求失敗 ({response.status_code})，3秒後重試...")
                    time.sleep(3)
            except Exception as e:
                log.warning(f"[_get_hk_name_via_langlink] 網絡錯誤 ({e})，3秒後重試...")
                time.sleep(3)

        query_data = response.json().get('query', {})
        pages_data = query_data.get('pages', {})

        if not pages_data or '-1' in pages_data:
            log.info(f"[_get_hk_name_via_langlink] 在 '{lang}' 維基上未直接找到 '{search_term}' 的頁面。")
            return None

        page_id = next(iter(pages_data))
        page_content = pages_data[page_id]
        final_page_title = page_content.get('title')

        # 步驟 3: 驗證頁面
        wiki_connector = get_wiki_connector(lang)
        page_to_validate = wiki_connector.page(final_page_title)
        
        # wikipediaapi 的 .exists() 也可能需要網絡，建議簡單 try-catch 或視為穩定
        if not page_to_validate.exists():
            log.warning(f"[_get_hk_name_via_langlink] API 找到的頁面 '{final_page_title}' 在 wikipediaapi 中無效。")
            return None
        
        if not _validate_page_type(page_to_validate, media_type, year):
            log.warning(f"[_get_hk_name_via_langlink] 頁面 '{final_page_title}' 未通過類型 '{media_type}' 的驗證。")
            return None
        log.info(f"[_get_hk_name_via_langlink] 頁面 '{final_page_title}' 已通過類型驗證。")

        # 步驟 4 & 5: 尋找中文 langlink
        langlinks = page_content.get('langlinks', [])
        zh_title = None
        for link in langlinks:
            if link.get('lang') == 'zh':
                zh_title = link.get('*')
                break
        
        if not zh_title:
            log.info(f"[_get_hk_name_via_langlink] 頁面 '{final_page_title}' 上未找到指向中文維基的 langlink。")
            return None
        log.info(f"[_get_hk_name_via_langlink] 找到中文維基頁面標題: '{zh_title}'。")

        # 步驟 6: 獲取香港繁體變體
        is_person = (media_type == 'person')
        hk_name = _get_zh_hk_variant_title(zh_title, is_person=is_person)

        if hk_name:
            log.info(f"[_get_hk_name_via_langlink] 成功透過 langlink 流程找到香港譯名: '{hk_name}'")
            return hk_name
        else:
            log.warning(f"[_get_hk_name_via_langlink] 已找到中文頁面 '{zh_title}'，但無法從中提取 zh-hk 變體標題。")
            return None

    except Exception as e:
        log.error(f"[_get_hk_name_via_langlink] 處理 '{search_term}' ({lang}) 時發生未知錯誤: {e}\n{traceback.format_exc()}")
        return None

def _clean_wiki_name(name):
    """Cleans common Wikipedia artifacts from a potential name string."""
    if not name:
        return ""
    # This regex is an attempt to remove simple templates, but might not cover all cases.
    name = re.sub(r'\{\{[^}]+\}\}', '', name)
    # Remove file links
    name = re.sub(r'\[\[File:[^\]]+\]\]', '', name, re.IGNORECASE)
    # Handle piped links, keeping only the display text
    name = re.sub(r'\[\[(?:[^|]+\|)?([^\]]+)\]\]', r'\1', name)
    # Remove references
    name = re.sub(r'<ref[^>]*>.*?</ref>|<ref[^>]*/>', '', name, re.DOTALL | re.IGNORECASE)
    # Remove any other HTML tags
    name = re.sub(r'<[^>]+>', '', name)
    # Remove content in parentheses
    name = re.sub(r'\s*\(.*\)\s*|\s*（.*）\s*', '', name).strip()
    # Strip leading/trailing punctuation
    name = name.strip().strip(';:-()[]').strip()
    return name

def _query_wikidata_for_label(qid):
    """
    【新增】使用 SPARQL 直接從 Wikidata 查詢給定 QID 的地區語言標籤。
    優先級為 zh-hk -> zh-hant -> zh-tw -> zh -> en。
    【修改】移除導致報錯的 setRetryCount，改為使用 while True 實現無限重試。
    """
    if not SPARQLWrapper:
        logging.warning("SPARQLWrapper 未加載，跳過 Wikidata 查詢。")
        return None
    if not qid:
        return None

    endpoint_url = "https://query.wikidata.org/sparql"
    # 使用 SERVICE wikibase:label 來自動後備語言
    query = f"""SELECT ?itemLabel WHERE {{
        BIND(wd:{qid} AS ?item)
        SERVICE wikibase:label {{
            bd:serviceParam wikibase:language "zh-hk, zh-hant, zh-tw, zh, en".
        }}
    }}"""

    sparql = SPARQLWrapper(endpoint_url)
    sparql.setQuery(query)
    sparql.setReturnFormat(JSON)
    sparql.addCustomHttpHeader("User-Agent", WIKI_USER_AGENT)
    
    # 【使用者要求】無限重試機制 (修復 AttributeError: 'SPARQLWrapper' object has no attribute 'setRetryCount')
# -     sparql.setRetryCount(50)
# -     sparql.setRetrySleep(5) # 設置重試間隔
# +     # 手動實現無限重試循環
    
    results = None
    retry_429 = 0
    retry_total = 0
    MAX_RETRIES = 10
    _ssl_bypass_done = False  # SSL 憑證錯誤只嘗試一次 bypass，失敗立即放棄

    while True:
        try:
            results = sparql.query().convert()
            break
        except Exception as e:
            err_str = str(e)
            is_ssl = 'SSL' in err_str or 'CERTIFICATE' in err_str.upper() or 'certificate' in err_str.lower()

            # SSL 憑證錯誤：環境問題，重試無效，改為立即嘗試停用 SSL 驗證
            if is_ssl:
                if not _ssl_bypass_done:
                    _ssl_bypass_done = True
                    logging.warning(f"查詢 Wikidata ({qid}) SSL 憑證錯誤，嘗試停用 SSL 驗證後重試...")
                    try:
                        import ssl as _ssl_mod
                        import urllib.request as _urllib_mod
                        _ctx = _ssl_mod.create_default_context()
                        _ctx.check_hostname = False
                        _ctx.verify_mode = _ssl_mod.CERT_NONE
                        _urllib_mod.install_opener(
                            _urllib_mod.build_opener(_urllib_mod.HTTPSHandler(context=_ctx))
                        )
                        continue  # 以停用 SSL 驗證重試
                    except Exception as ssl_e:
                        logging.warning(f"查詢 Wikidata ({qid}) 停用 SSL 驗證失敗: {ssl_e}，放棄。")
                        return None
                else:
                    # SSL bypass 後仍失敗，立即放棄（不等待）
                    logging.warning(f"查詢 Wikidata ({qid}) 停用 SSL 驗證後仍失敗，放棄。")
                    return None

            # 服務中斷（outage）時直接放棄，重試無意義
            if 'outage' in err_str.lower():
                logging.warning(f"查詢 Wikidata ({qid}) 偵測到服務中斷，立即跳過。錯誤: {e}")
                return None

            retry_total += 1
            if retry_total >= MAX_RETRIES:
                logging.warning(f"查詢 Wikidata ({qid}) 已重試 {MAX_RETRIES} 次仍失敗，放棄。最後錯誤: {e}")
                return None
            is_429 = '429' in err_str
            wait = 15 if is_429 else 5
            if is_429:
                _mark_rate_limited()
                retry_429 += 1
                if retry_429 >= 2:
                    logging.warning(f"查詢 Wikidata ({qid}) 連續 429 達 2 次，跳過。")
                    return None
                logging.warning(f"查詢 Wikidata ({qid}) 時發生錯誤: {e}，{wait}秒後重試 ({retry_429}/2，總計 {retry_total}/{MAX_RETRIES})...")
            else:
                retry_429 = 0
                logging.warning(f"查詢 Wikidata ({qid}) 時發生錯誤: {e}，{wait}秒後重試 ({retry_total}/{MAX_RETRIES})...")
            if _interruptible_sleep(wait):
                return None

    try:
        bindings = results.get("results", {}).get("bindings", [])

        if bindings:
            label = bindings[0].get("itemLabel", {}).get("value")
            lang = bindings[0].get("itemLabel", {}).get("xml:lang")
            if label:
                logging.info(f"成功從 Wikidata ({qid}) 透過 SPARQL 查詢找到標籤: '{label}' (語言: {lang})")
                return label
        return None
    except Exception as e:
        logging.error(f"解析 Wikidata ({qid}) 結果時發生錯誤: {e}")
        return None

def _get_zhwiki_title_from_wikidata(qid):
    # 【新增的輔助函式】根據你的測試程式碼改寫，用於獲取中文維基頁面標題
    """
    從 Wikidata API 獲取給定 QID 對應的中文維基百科頁面標題。
    【修改】加入無限重試機制。
    """
    if not qid:
        return None
    try:
        url = f"https://www.wikidata.org/wiki/Special:EntityData/{qid}.json"
        headers = {'User-Agent': WIKI_USER_AGENT}
        # 【修改】使用帶有重試機制的 session
        session = _create_retry_session()
        
        # 【修改】加入無限重試循環
# +       response = None
# +       while True:
# +           try:
# +               response = session.get(url, headers=headers, timeout=10)
# +               if response.status_code == 200:
# +                   break
# +               elif response.status_code == 404:
# +                   logging.info(f"Wikidata ({qid}) 不存在 (404)。")
# +                   return None
# +               else:
# +                   logging.warning(f"獲取 Wikidata ({qid}) 失敗 ({response.status_code})，3秒後重試...")
# +                   time.sleep(3)
# +           except requests.exceptions.RequestException as e:
# +               logging.warning(f"獲取 Wikidata ({qid}) 網絡錯誤 ({e})，3秒後重試...")
# +               time.sleep(3)

# -       response = session.get(url, headers=headers, timeout=10)
# -       response.raise_for_status()

        response = None
        while True:
            try:
                response = session.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    break
                elif response.status_code == 404:
                    logging.info(f"Wikidata ({qid}) 不存在 (404)。")
                    return None
                else:
                    if response.status_code == 429:
                        _mark_rate_limited()
                    logging.warning(f"獲取 Wikidata ({qid}) 失敗 ({response.status_code})，3秒後重試...")
                    time.sleep(3)
            except requests.exceptions.RequestException as e:
                logging.warning(f"獲取 Wikidata ({qid}) 網絡錯誤 ({e})，3秒後重試...")
                time.sleep(3)

        data = response.json()
        entity = data.get('entities', {}).get(qid, {})
        sitelinks = entity.get('sitelinks', {})
        zh_link_title = sitelinks.get('zhwiki', {}).get('title')
        if zh_link_title:
            logging.info(f"從 Wikidata ({qid}) 找到對應的中文維基頁面標題: '{zh_link_title}'")
            return zh_link_title
        else:
            logging.info(f"Wikidata ({qid}) 中沒有 'zhwiki' 的 sitelink。")
            return None
    except Exception as e:
        logging.error(f"解析 Wikidata ({qid}) 的 'zhwiki' 標題時出錯: {e}")
        return None

def _get_zh_hk_variant_title(wiki_title, is_person=False):
    """
    【修改】訪問中文維基頁面的 zh-hk 變體，優先從 NoteTA 模板中提取香港譯名，並以 H1 標題作為後備和驗證。
    新增 is_person 參數以決定是否保護「里」字。
    【修改】加入無限重試機制。
    """
    if not wiki_title:
        return None
    try:
        # 使用下劃線替換空格，這更符合維基百科的URL格式
        encoded_title = quote(wiki_title.replace(" ", "_"))
        url = f"https://zh.wikipedia.org/wiki/{encoded_title}?variant=zh-hk"
        headers = {'User-Agent': WIKI_USER_AGENT}
        # 【修改】使用帶有重試機制的 session
        session = _create_retry_session()
        
        # 【修改】加入無限重試循環
# +       response = None
# +       while True:
# +           try:
# +               response = session.get(url, headers=headers, timeout=10)
# +               if response.status_code == 200:
# +                   break
# +               elif response.status_code == 404:
# +                   logging.warning(f"中文維基頁面 '{wiki_title}' 不存在 (404)。")
# +                   return None
# +               else:
# +                   logging.warning(f"獲取中文維基變體失敗 ({response.status_code})，3秒後重試...")
# +                   time.sleep(3)
# +           except requests.exceptions.RequestException as e:
# +               logging.warning(f"獲取中文維基變體網絡錯誤 ({e})，3秒後重試...")
# +               time.sleep(3)

# -       response = session.get(url, headers=headers, timeout=10)
# -       response.raise_for_status()

        response = None
        while True:
            try:
                response = session.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    break
                elif response.status_code == 404:
                    logging.warning(f"中文維基頁面 '{wiki_title}' 不存在 (404)。")
                    return None
                else:
                    if response.status_code == 429:
                        _mark_rate_limited()
                    logging.warning(f"獲取中文維基變體失敗 ({response.status_code})，3秒後重試...")
                    time.sleep(3)
            except requests.exceptions.RequestException as e:
                logging.warning(f"獲取中文維基變體網絡錯誤 ({e})，3秒後重試...")
                time.sleep(3)

        soup = BeautifulSoup(response.text, 'html.parser')

        # 步驟 1: 獲取 H1 作為後備和驗證基準
        h1_tag = soup.find("h1", {"id": "firstHeading"})
        h1_title = _clean_wiki_name(h1_tag.text) if h1_tag and h1_tag.text else None
        if not h1_title:
            logging.warning(f"無法從頁面 '{wiki_title}' 獲取 H1 標題，中止譯名提取。")
            return None
        
        normalized_h1 = normalize_to_hk(h1_title, field_type=('name' if is_person else 'title'))

        # 步驟 2: 尋找所有 NoteTA 區塊
        note_ta_blocks = soup.find_all("div", class_="noteTA")
        if not note_ta_blocks:
            logging.info(f"頁面 '{wiki_title}' 無 NoteTA 模板，將使用 H1 標題 '{h1_title}' 作為結果。")
            return h1_title

        logging.info(f"在頁面 '{wiki_title}' 找到 {len(note_ta_blocks)} 個 NoteTA 模板。開始尋找主要標題的譯名...")

        # 步驟 3: 遍歷 NoteTA 區塊，尋找與主標題相關的譯名
        for block in note_ta_blocks:
            # 在一個區塊內尋找所有 data-noteta-code 的 div
            code_divs = block.find_all("div", attrs={"data-noteta-code": True})
            for code_div in code_divs:
                code_str = code_div['data-noteta-code']
                # 解析 data-noteta-code 字串，例如 "zh-tw:膽大黨;zh-hk:DAN DA DAN;zh-cn:胆大党;"
                variants = {}
                parts = code_str.split(';')
                for part in parts:
                    if ':' in part:
                        key, val = part.split(':', 1)
                        if key.strip() and val.strip():
                            variants[key.strip().lower()] = val.strip()
                
                # 透過與 H1 標題比較，檢查此 NoteTA 是否為文章主標題的譯名
                is_main_title_note = False
                # 檢查的鍵列表，包含了常見的中文變體
                keys_to_check = ['zh-tw', 'zh-cn', 'zh-hant', 'zh-hans', 'zh-sg', 'zh-my', 'zh-mo']
                for key in keys_to_check:
                    if key in variants and normalize_to_hk(variants[key], field_type=('name' if is_person else 'title')) == normalized_h1:
                        is_main_title_note = True
                        logging.info(f"成功將 NoteTA 內容 ('{key}':'{variants[key]}') 與 H1 標題 ('{h1_title}') 關聯。")
                        break
                
                if is_main_title_note:
                    # 如果找到了正確的區塊，提取 zh-hk 譯名
                    if 'zh-hk' in variants:
                        hk_name = _clean_wiki_name(variants['zh-hk'])
                        if hk_name:
                            logging.info(f"成功從 NoteTA 中提取到香港譯名: '{hk_name}'。此為優先結果。")
                            return hk_name  # 返回找到的高優先級名稱
                    else:
                        logging.warning(f"已關聯到正確的 NoteTA，但其中未定義 'zh-hk' 譯名。內容: {code_str}")
                        # 即使這個 noteTA 沒有 zh-hk，我們也應該認為它就是主標題的 noteTA，
                        # 不應再檢查其他 noteTA，所以直接跳出內外層循環，使用 H1 作為後備。
                        logging.info(f"因無 zh-hk 譯名，將使用 H1 標題 '{h1_title}' 作為後備。")
                        return h1_title

        # 步驟 4: 如果遍歷完所有 NoteTA 都沒有找到關聯的，則回退到 H1
        logging.info(f"未能從 NoteTA 中找到與 H1 標題 '{h1_title}' 關聯的譯名，將使用 H1 標題作為後備。")
        return h1_title

    except Exception as e:
        logging.error(f"解析維基頁面 '{wiki_title}' 的 'zh-hk' 變體時出錯: {e}\n{traceback.format_exc()}")
        return None

def _find_hk_name_in_text(text, media_type):
    """
    【修改】透過解析 wikitext 來尋找香港地區的特定名稱。
    此版本增加了對 {{NoteTA}} 模板的解析，這是中文維基上最常見的譯名處理方式。
    """
    if not text:
        return None
    
    is_person = (media_type == 'person')
        
    # 優先策略：尋找 {{NoteTA}} 模板，這是最精確的來源
    # 範例: {{NoteTA|1=zh-cn:玛格特·罗比; zh-tw:瑪格·羅比; zh-hk:瑪歌羅比;}}
    noteta_match = re.search(r'\{\{NoteTA\|([^}]+)\}\}', text, re.IGNORECASE)
    if noteta_match:
        params_str = noteta_match.group(1)
        # 尋找 zh-hk, hk, 或 yue 的譯名定義
        hk_match = re.search(r'(?:zh-hk|hk|yue)\s*:\s*([^;|}]+)', params_str, re.IGNORECASE)
        if hk_match:
            hk_name = _clean_wiki_name(hk_match.group(1))
            if hk_name and is_chinese(hk_name):
                logging.info(f"從 {{NoteTA}} 模板中找到香港譯名: '{hk_name}'")
                return normalize_to_hk(hk_name, field_type=('name' if is_person else 'title'))

    # 後備策略一：尋找 "各地譯名" 章節
    section_match = re.search(r'==\s*(?:各地譯名|譯名|名稱|Chinese names|Regional titles|Other names)\s*==\n(.+?)(?=\n==|\Z)', text, re.DOTALL | re.IGNORECASE)
    if section_match:
        section_text = section_match.group(1)
        hk_pattern_match = re.search(r'(?:香港|港)\s*[:：\s]*\'*\{\{lang\|zh-hk\|([^}]+)\}\}\'*|(?<!\w)(?:香港|港)\s*[:：\s]*\'*([^\n|}<,]+)\'*', section_text, re.IGNORECASE)
        if hk_pattern_match:
            hk_name = _clean_wiki_name(hk_pattern_match.group(1) or hk_pattern_match.group(2))
            if hk_name and is_chinese(hk_name):
                logging.info(f"在 '各地譯名' 章節中找到香港譯名: '{hk_name}'")
                return normalize_to_hk(hk_name, field_type=('name' if is_person else 'title'))

    # 後備策略二：尋找 {{zh}} 或 {{Chinese}} 模板
    zh_template_match = re.search(r'\{\{(?:zh|Chinese)\s*\|([^}]+)\}\}', text, re.IGNORECASE)
    if zh_template_match:
        params_str = zh_template_match.group(1)
        param_dict = dict(re.findall(r'\|\s*([^=]+?)\s*=\s*([^|]+?)(?=\s*\||\s*\}\})', "|" + params_str))
        param_dict = {k.lower().strip(): v.strip() for k, v in param_dict.items()}
        for hk_key in ['hk', 'h', 'yue']:
            if hk_key in param_dict:
                hk_name = _clean_wiki_name(param_dict[hk_key])
                if hk_name and is_chinese(hk_name):
                    logging.info(f"在 {{zh/Chinese}} 模板中透過鍵 '{hk_key}' 找到香港譯名: '{hk_name}'")
                    return normalize_to_hk(hk_name, field_type=('name' if is_person else 'title'))
        if 't' in param_dict:
            t_name = _clean_wiki_name(param_dict['t'])
            if t_name and is_chinese(t_name):
                logging.info(f"在 {{zh/Chinese}} 模板中透過鍵 't' 找到繁體名稱，作為香港譯名的後備: '{t_name}'")
                return normalize_to_hk(t_name, field_type=('name' if is_person else 'title'))

    # 後備策略三：尋找資訊框（infobox）中的參數
    infobox_matches = re.finditer(r'\|\s*(?:traditional_chinese|trad_chi|hk_name)\s*=\s*(.+?)(?=\n\s*\|)', text, re.DOTALL | re.IGNORECASE)
    for match in infobox_matches:
        content = match.group(1)
        hk_name = _clean_wiki_name(content)
        if hk_name and is_chinese(hk_name):
             logging.info(f"在資訊框參數中找到香港譯名: '{hk_name}'")
             return normalize_to_hk(hk_name, field_type=('name' if is_person else 'title'))
             
    return None

def _validate_page_type(page, media_type, year):
    """
    【強化版】通過檢查標題、資訊框和分類來驗證維基頁面是否符合預期的媒體類型。
    """
    if not page or not page.exists():
        return False
        
    title_lower = page.title.lower()
    # Lazy load text and categories only when needed
    text_lower = page.text.lower() if page.text else ""
    summary_lower = page.summary.lower() if page.summary else ""
    
    try:
        categories_lower = [cat.lower() for cat in page.categories.keys()]
    except Exception:
        categories_lower = []


    if media_type == 'person':
        # --- 負面檢查 (NEGATIVE CHECKS for PERSON) ---
        if any(kw in title_lower for kw in ['list of', 'filmography', 'discography', '一覧', '목록', '列表', '電影作品', '唱片']):
             logging.warning(f"Validation failed for 'person': Page title '{page.title}' suggests it is a list or collection.")
             return False
        
        wrong_infoboxes = ['infobox film', 'infobox television', 'infobox album', 'infobox book', 'infobox video game']
        if any(wrong_box in text_lower for wrong_box in wrong_infoboxes):
            found_wrong_boxes = [b for b in wrong_infoboxes if b in text_lower]
            logging.warning(f"Validation failed for 'person': Page '{page.title}' contains a non-person infobox ({found_wrong_boxes}).")
            return False
            
        wrong_cat_keywords = ['films', 'television series', 'albums', 'songs', 'video games', 'lists of', '映画', 'テレビドラマ', 'アルバム', '목록', '列表']
        if any(any(kw in cat for kw in wrong_cat_keywords) for cat in categories_lower):
            logging.warning(f"Validation failed for 'person': Page '{page.title}' has disqualifying categories.")
            return False

        if re.search(r'\b(is|was)\s+a\s+([\w\s-]+\s+)?(film|movie|television\s+series|tv\s+show|anime|drama|book|novel)\b', summary_lower):
            logging.warning(f"Validation failed for 'person': Page '{page.title}' summary describes a film/show/book.")
            return False

        # --- 正面檢查 (POSITIVE CHECKS for PERSON) ---
        if 'infobox person' in text_lower or 'infobox officeholder' in text_lower or 'infobox sportsperson' in text_lower:
            return True
        if year and any(f'{year} births' in cat for cat in categories_lower):
            return True
        if any('births' in cat or 'deaths' in cat for cat in categories_lower):
            return True
        
        person_category_keywords = [
            'actors', 'actresses', 'directors', 'writers', 'singers', 'filmmakers', 'people from', 'alumni of', '人物', 
            '俳優', '声優', '監督', '배우', '가수', 'schauspieler', 'regisseur', 'autor', 'sänger', 'filmemacher', 'geboren',
            'acteur', 'actrice', 'réalisateur', 'écrivain', 'chanteur', 'cinéaste', 'naissance en',
            'attore', 'attrice', 'regista', 'scrittore', 'cantante', 'cineasta', 'nati nel'
        ]
        if any(any(person_keyword in cat for person_keyword in person_category_keywords) for cat in categories_lower):
            return True
        if re.search(r'\b(is|was)\s+an?\s+[\w\s-]+\s+(actor|actress|director|singer|politician|writer|artist|filmmaker)\b', summary_lower):
            return True

    elif media_type in ['movie', 'film', 'tv', 'television']:
        # --- 負面檢查 (NEGATIVE CHECKS for FILM/TV) ---
        if 'infobox person' in text_lower:
             logging.warning(f"Validation failed for '{media_type}': Page '{page.title}' contains a person infobox.")
             return False
        if any('births' in cat or 'deaths' in cat for cat in categories_lower):
            logging.warning(f"Validation failed for '{media_type}': Page '{page.title}' has person categories (e.g., 'births', 'deaths').")
            return False
        if any(kw in title_lower for kw in ['list of', 'filmography', 'discography', '一覧', '목록', '列表']):
             logging.warning(f"Validation failed for '{media_type}': Page title '{page.title}' suggests it is a list or collection.")
             return False

        # --- 正面檢查 (POSITIVE CHECKS for FILM/TV) ---
        media_type_std = 'film' if media_type in ['movie', 'film'] else 'television'
        
        # 1. Infobox 檢查 (多語言)
        film_infoboxes = ['infobox film', '電影資訊框']
        tv_infoboxes = ['infobox television', '電視資訊框', 'infobox tv series', '電視劇資訊框', 'infobox animanga', '애니메이션 정보']
        target_infoboxes = film_infoboxes if media_type_std == 'film' else tv_infoboxes
        if any(box in text_lower for box in target_infoboxes):
            return True

        # 2. 分類檢查 (多語言)
        film_cat_keywords = ['films', 'movies', '映画', '영화', '電影', '影片']
        tv_cat_keywords = ['television series', 'tv series', 'テレビドラマ', '텔레비전 드라마', '電視劇', '電視影集', '電視節目', '電視動畫', 'anime', '電視動畫', '애니메이션']
        target_cat_keywords = film_cat_keywords if media_type_std == 'film' else tv_cat_keywords
        if any(any(kw in cat for kw in target_cat_keywords) for cat in categories_lower):
            return True

        # 3. 摘要檢查 (多語言)
        film_summary_patterns = [
            r'\b(is|was)\s+a\s+([\w\s-]+\s+)?(film|movie)\b', 
            r'是一部\d{4}年.*?的(電影|影片)'
        ]
        tv_summary_patterns = [
            r'\b(is|was)\s+a\s+([\w\s-]+\s+)?(television\s+series|tv\s+show|anime|drama)\b', 
            r'是一部.*?的(電視劇|電視影集|電視節目|電視動畫|日本動畫)'
        ]
        target_summary_patterns = film_summary_patterns if media_type_std == 'film' else tv_summary_patterns
        if any(re.search(pattern, summary_lower) for pattern in target_summary_patterns):
            return True


    logging.warning(f"Page '{page.title}' could not be positively validated as media_type '{media_type}'.")
    return False

def _search_one_wiki(lang, search_term, media_type, year):
    """
    【修改】在單一語言維基上搜索，並返回(香港譯名, 維基數據QID)。
    此版本整合了 Wikidata SPARQL 查詢作為主要譯名來源，並使用 requests API 呼叫來處理重定向，
    使其更加可靠。
    【修改】加入無限重試機制。
    """
    # 【使用者要求】新增檢查，如果語言代碼是無效的 'xx'，則直接返回 None，避免程式因域名解析錯誤而崩潰。
    if lang == 'xx':
        logging.warning(f"偵測到無效的語言代碼 'xx'，將跳過在 '{lang}' 維基百科上對 '{search_term}' 的搜索。")
        return None, None

    api_lang_host = 'zh' if lang.startswith('zh-') else lang
    wiki_wiki = get_wiki_connector(lang)
    # 【修改】使用帶有重試機制的 session
    session = _create_retry_session()
    headers = {'User-Agent': WIKI_USER_AGENT}
    is_person = (media_type == 'person')

    logging.info(f"正在從 '{lang}' 維基百科（主機: {api_lang_host}.wikipedia.org）搜索 '{search_term}'...")

    # --- 步驟 1: API 呼叫，獲取頁面資料和 QID ---
    page_content = None
    api_url = f"https://{api_lang_host}.wikipedia.org/w/api.php"
    params = {
        'action': 'query', 'format': 'json', 'redirects': 1, 'titles': search_term,
        'prop': 'info|pageprops|langlinks', 'inprop': 'url', 'ppprop': 'wikibase_item', 'lllimit': 'max'
    }
    
    try:
        # 【修改】加入無限重試循環 (API 查詢 1)
        response = None
        while True:
            try:
                response = session.get(api_url, params=params, headers=headers, timeout=10)
                if response.status_code == 200:
                    break
                elif response.status_code == 404:
                     logging.info(f"API 查詢返回 404 for '{search_term}'")
                     return None, None
                else:
                    if response.status_code == 429:
                        _mark_rate_limited()
                    logging.warning(f"API 查詢失敗 ({response.status_code})，3秒後重試...")
                    time.sleep(3)
            except Exception as e:
                logging.warning(f"API 查詢網絡錯誤 ({e})，3秒後重試...")
                time.sleep(3)

        query_data = response.json().get('query', {})
        pages_data = query_data.get('pages', {})
        
        if not pages_data or '-1' in pages_data:
            logging.info(f"直接標題查找失敗，後備至 API 搜索功能 for '{search_term}'...")
            time.sleep(0.5)
            search_params = {'action': 'query', 'format': 'json', 'list': 'search', 'srsearch': search_term, 'srlimit': 1}
            
            # 【修改】加入無限重試循環 (API 查詢 2 - 搜索)
            search_response = None
            while True:
                try:
                    search_response = session.get(api_url, params=search_params, headers=headers, timeout=10)
                    if search_response.status_code == 200:
                        break
                    else:
                        if search_response.status_code == 429:
                            _mark_rate_limited()
                        logging.warning(f"API 搜索失敗 ({search_response.status_code})，3秒後重試...")
                        time.sleep(3)
                except Exception as e:
                    logging.warning(f"API 搜索網絡錯誤 ({e})，3秒後重試...")
                    time.sleep(3)

            search_results = search_response.json().get('query', {}).get('search', [])
            if search_results:
                found_title = search_results[0]['title']
                logging.info(f"API 搜索找到候選標題 '{found_title}', 重新查詢其詳細資料...")
                params['titles'] = found_title
                time.sleep(0.5)
                # 【修改】加入無限重試循環 (API 查詢 3 - 詳細資料)
                response = None
                while True:
                    try:
                        response = session.get(api_url, params=params, headers=headers, timeout=10)
                        if response.status_code == 200:
                            break
                        else:
                            if response.status_code == 429:
                                _mark_rate_limited()
                            logging.warning(f"API 詳細資料查詢失敗 ({response.status_code})，3秒後重試...")
                            time.sleep(3)
                    except Exception as e:
                        logging.warning(f"API 詳細資料查詢網絡錯誤 ({e})，3秒後重試...")
                        time.sleep(3)

                query_data = response.json().get('query', {})
                pages_data = query_data.get('pages', {})

        if pages_data and '-1' not in pages_data:
            page_id = next(iter(pages_data))
            page_content = pages_data[page_id]
        
    except Exception as e:
        logging.error(f"透過 API 查找頁面 '{search_term}' ({lang}) 時出錯: {e}")
        return None, None

    if not page_content:
        logging.warning(f"在 '{lang}' 維基百科上未能為 '{search_term}' 找到一個有效的頁面。")
        return None, None
        
    final_page_title = page_content.get('title')
    identity = page_content.get('pageprops', {}).get('wikibase_item')
    hk_name = None

    # --- 步驟 2: 驗證頁面並提取香港譯名 ---
    # 在 session.get() 之後加入短暫延遲，避免連續快速請求觸發 Wikipedia burst throttle。
    time.sleep(1)
    page_to_validate = wiki_wiki.page(final_page_title)
    if not page_to_validate.exists():
        logging.warning(f"API 返回的頁面標題 '{final_page_title}' 在 wikipediaapi 庫中似乎無效或不存在。")
        return None, identity

    logging.info(f"正在驗證候選頁面: '{page_to_validate.title}' (媒體類型: '{media_type}')")
    if not _validate_page_type(page_to_validate, media_type, year):
        logging.warning(f"頁面 '{page_to_validate.title}' 驗證失敗。")
        return None, None # 【關鍵修改】驗證失敗時，返回 None, None，避免傳遞錯誤的 identity

    logging.info(f"頁面 '{page_to_validate.title}' 驗證成功，接受此頁面。")
    
    # --- 步驟 3: 譯名提取 (優先使用 Wikidata) ---
    if identity:
        hk_name = _query_wikidata_for_label(identity)
        if hk_name:
            return normalize_to_hk(_clean_wiki_name(hk_name), field_type=('name' if is_person else 'title')), identity

    logging.info("Wikidata 未返回香港譯名，後備至解析頁面內文。")
    hk_name = _find_hk_name_in_text(page_to_validate.text, media_type)

    return hk_name, identity

def search_wiki_for_hk_name(item_id, original_name, media_type, year=None, search_targets_by_lang=None, original_language=None, tmdb_chinese_names=None):
    """
    【修改 2025-08-27】新增優先搜索路徑，以解決特定原始語言影片的譯名問題。
    使用者要求：如果影片的原始語言是 en, ja (日文), 或 ko (韓文)，則應優先從該語言的維基百科開始搜索。
    新邏輯:
    1.  檢查 `original_language` 參數。
    2.  如果為 'en', 'ja', 或 'ko'，則調用新的輔助函數 `_get_hk_name_via_langlink`，
        該函數會嘗試從外語維基頁面透過語言連結 (langlink) 跳轉到中文維基頁面，並提取香港繁體標題。
    3.  如果這個優先路徑成功找到譯名，則直接使用該結果並返回。
    4.  如果優先路徑失敗，或 `original_language` 不是指定的語言，則回退到原有的、更廣泛的多語言搜索和投票機制。
    """
    _reset_rate_limit_flag()
    _init_db()

    if not original_name and not item_id:
        return None

    # --- 步驟 1: 檢查快取 ---
    if item_id:
        cached_result = _query_wiki_cache(item_id, media_type)
        if cached_result:
            if cached_result['status'] == 'FOUND':
                logging.info(f"從快取中找到 ID '{item_id}'。結果: '{cached_result['hk_name']}'")
                return cached_result['hk_name']
            elif cached_result['status'] in ['NOT_FOUND', 'FAILED']:
                logging.info(f"從快取中得知 ID '{item_id}' 先前未找到或搜索失敗。跳過網路搜索。")
                return None

    # --- 步驟 2: 【全新】執行基於原始語言的優先搜索 ---
    final_hk_name = None
    media_type_std = 'film' if media_type == 'movie' else ('television' if media_type == 'tv' else media_type)
    
    # 此優先路徑僅對影片類型生效
    if original_language in ['en', 'ja', 'ko'] and media_type_std in ['film', 'television']:
        logging.info(f"檢測到優先語言 '{original_language}'，為 '{original_name}' 啟動 langlink 搜索...")
        # 使用原始標題進行搜索，因為這最可能對應到原始語言的維基頁面
        hk_name_from_langlink = _get_hk_name_via_langlink(original_language, original_name, media_type_std, year)
        if hk_name_from_langlink:
            final_hk_name = hk_name_from_langlink
            logging.info(f"優先搜索成功，找到譯名: '{final_hk_name}'。")

    # --- 步驟 3: 如果優先搜索失敗，則執行後備的廣泛搜索 ---
    if not final_hk_name:
        if original_language in ['en', 'ja', 'ko'] and media_type_std in ['film', 'television']:
            logging.info("優先 langlink 搜索未找到結果，將回退至標準的多語言搜索流程。")
        
        try:
            # ... (此處的搜索邏輯保持不變) ...
            is_person = (media_type_std == 'person')
            
            if media_type_std in ['film', 'television'] and search_targets_by_lang:
                unique_zh_titles = list(set(v for k, v in search_targets_by_lang.items() if k.startswith('zh') and v))
                
                if unique_zh_titles:
                    logging.info(f"第一階段搜索: 嘗試使用中文標題在中文維基上搜索: {unique_zh_titles}")
                    zh_results = []
                    for title in unique_zh_titles:
                        hk_name_res, identity_res = _search_one_wiki('zh', title, media_type_std, year)
                        if identity_res:
                            zh_results.append({'hk_name': hk_name_res, 'identity': identity_res})
                    
                    if zh_results:
                        zh_identities = {res['identity'] for res in zh_results}
                        if len(zh_identities) == 1:
                            winner_identity = list(zh_identities)[0]
                            logging.info(f"第一階段成功: 所有中文標題指向同一維基實體 {winner_identity}。")
                            
                            validation_passed = True
                            if tmdb_chinese_names and any(tmdb_chinese_names.values()):
                                tmdb_tw_raw, tmdb_cn_raw = tmdb_chinese_names.get('zh-TW'), tmdb_chinese_names.get('zh-CN')
                                wiki_zh_variants_raw = {}
                                try:
                                    url = f"https://www.wikidata.org/w/api.php?action=wbgetentities&ids={winner_identity}&props=labels&languages=zh|zh-hans|zh-hant|zh-hk|zh-mo|zh-cn|zh-sg|zh-tw|zh-my&format=json"
                                    response = requests.get(url, headers={'User-Agent': WIKI_USER_AGENT}, timeout=10)
                                    response.raise_for_status()
                                    labels_data = response.json().get('entities', {}).get(winner_identity, {}).get('labels', {})
                                    wiki_zh_variants_raw = {code: lbl['value'] for code, lbl in labels_data.items()}
                                except Exception as e:
                                    logging.error(f"Wiki 驗證: 無法為 QID {winner_identity} 獲取標籤。錯誤: {e}")
                                
                                if wiki_zh_variants_raw:
                                    tmdb_tw_norm = normalize_to_hk(tmdb_tw_raw, field_type=('name' if is_person else 'title')).replace('·', '') if tmdb_tw_raw else None
                                    tmdb_cn_norm = normalize_to_hk(tmdb_cn_raw, field_type=('name' if is_person else 'title')).replace('·', '') if tmdb_cn_raw else None
                                    wiki_names_norm_set = {normalize_to_hk(name, field_type=('name' if is_person else 'title')).replace('·', '') for name in wiki_zh_variants_raw.values()}
                                    
                                    if not ((tmdb_tw_norm and tmdb_tw_norm in wiki_names_norm_set) or \
                                            (tmdb_cn_norm and tmdb_cn_norm in wiki_names_norm_set)):
                                        logging.warning(f"Wiki 驗證失敗 for '{original_name}' ({winner_identity}): TMDB 中文名未在 Wikidata 變體中找到匹配。")
                                        validation_passed = False
                                    else:
                                        logging.info(f"Wiki 驗證成功 for '{original_name}' ({winner_identity})。")
                            
                            if validation_passed:
                                winner_names = [res['hk_name'] for res in zh_results if res['hk_name'] and is_chinese(res['hk_name'])]
                                if winner_names:
                                    final_hk_name = Counter(winner_names).most_common(1)[0][0]
                                    logging.info(f"從中文維基搜索中確定最終結果: '{final_hk_name}'")
                                else:
                                    logging.info("第一階段搜索找到頁面但無有效中文名，將繼續進行第二階段。")
                        else:
                            logging.warning(f"第一階段失敗: 中文標題指向多個不同實體 {zh_identities}，將繼續進行第二階段。")

            if not final_hk_name:
                logging.info(f"為 '{original_name}' ({media_type}) 啟動全語言目標搜索。")
                search_plan = search_targets_by_lang if search_targets_by_lang else {'en': original_name}
                results = []
                for lang, name in search_plan.items():
                    if not name: continue
                    # 各語言之間加 3 秒延遲，避免連續觸發 Wikipedia 429 導致程式停頓
                    if _interruptible_sleep(3):
                        break
                    try:
                        hk_name, identity = _search_one_wiki(lang, name, media_type_std, year)
                        if identity:
                            results.append({'hk_name': hk_name, 'identity': identity, 'source_lang': lang})
                    except Exception as e:
                        logging.error(f"在搜索 '{lang}' 維基時為 '{name}' 發生意外錯誤: {e}")
                        continue

                if results:
                    identity_groups = {}
                    for res in results:
                        if res['identity'] not in identity_groups: identity_groups[res['identity']] = []
                        identity_groups[res['identity']].append(res)
                    
                    grouped_langs = {qid: [r['source_lang'] for r in grp] for qid, grp in identity_groups.items()}
                    logging.info(f"搜索 '{original_name}' 找到的身份標識分組: {grouped_langs}")

                    if identity_groups:
                        sorted_groups = sorted(identity_groups.items(), key=lambda item: len(item[1]), reverse=True)
                        best_identity, best_group_size = sorted_groups[0][0], len(sorted_groups[0][1])
                        
                        if len(sorted_groups) > 1 and best_group_size <= len(sorted_groups[1][1]):
                            logging.warning(f"對 '{original_name}' 的搜索結果指向了多個不同的維基實體，且支持度最高者並未明確勝出。因存在衝突，放棄採納。")
                        else:
                            has_tmdb_chinese_name = tmdb_chinese_names and any(tmdb_chinese_names.values())
                            if not has_tmdb_chinese_name and best_group_size < 2:
                                logging.warning(f"對 '{original_name}' 的搜索確定了最可能的實體 '{best_identity}', 但它僅由 {best_group_size} 個語言支持。由於 TMDB 中沒有中文名稱，根據規則要求至少 2 個語言支持，故放棄採納此結果。")
                            else:
                                validation_passed = True
                                if tmdb_chinese_names and any(tmdb_chinese_names.values()):
                                    tmdb_tw_raw, tmdb_cn_raw = tmdb_chinese_names.get('zh-TW'), tmdb_chinese_names.get('zh-CN')
                                    wiki_zh_variants_raw = {}
                                    try:
                                        url = f"https://www.wikidata.org/w/api.php?action=wbgetentities&ids={best_identity}&props=labels&languages=zh|zh-hans|zh-hant|zh-hk|zh-mo|zh-cn|zh-sg|zh-tw|zh-my&format=json"
                                        response = requests.get(url, headers={'User-Agent': WIKI_USER_AGENT}, timeout=10)
                                        response.raise_for_status()
                                        labels_data = response.json().get('entities', {}).get(best_identity, {}).get('labels', {})
                                        wiki_zh_variants_raw = {code: lbl['value'] for code, lbl in labels_data.items()}
                                    except Exception as e:
                                        logging.error(f"Wiki 驗證: 無法為 QID {best_identity} 獲取標籤。錯誤: {e}")
                                    
                                    if not wiki_zh_variants_raw:
                                        logging.warning(f"Wiki 驗證: 未能從 Wikidata 獲取任何中文標籤 for QID {best_identity}，跳過驗證步驟。")
                                    else:
                                        tmdb_tw_norm = normalize_to_hk(tmdb_tw_raw, field_type=('name' if is_person else 'title')).replace('·', '') if tmdb_tw_raw else None
                                        tmdb_cn_norm = normalize_to_hk(tmdb_cn_raw, field_type=('name' if is_person else 'title')).replace('·', '') if tmdb_cn_raw else None
                                        wiki_names_norm_set = {normalize_to_hk(name, field_type=('name' if is_person else 'title')).replace('·', '') for name in wiki_zh_variants_raw.values()}
                                        
                                        if not ((tmdb_tw_norm and tmdb_tw_norm in wiki_names_norm_set) or \
                                                (tmdb_cn_norm and tmdb_cn_norm in wiki_names_norm_set)):
                                            logging.warning(f"Wiki 驗證失敗 for '{original_name}' ({best_identity}): TMDB 中文名未在 Wikidata 變體中找到匹配。拒絕此維基結果。")
                                            validation_passed = False
                                        else:
                                            logging.info(f"Wiki 驗證成功 for '{original_name}' ({best_identity})。")

                                if validation_passed:
                                    logging.info(f"已確定最可能的維基數據實體為 '{best_identity}'。現在嘗試使用新方法直接獲取其香港繁體名稱。")
                                    zhwiki_page_title = _get_zhwiki_title_from_wikidata(best_identity)
                                    if zhwiki_page_title:
                                        hk_name_from_variant = _get_zh_hk_variant_title(zhwiki_page_title, is_person=is_person)
                                        if hk_name_from_variant:
                                            final_hk_name = normalize_to_hk(hk_name_from_variant, field_type=('name' if is_person else 'title'))
                                            logging.info(f"新方法成功找到並確定最終結果: '{final_hk_name}'")

                                    if not final_hk_name:
                                        logging.info(f"新方法未能找到有效的香港繁體名稱，將回退到原有的投票機制。")
                                        winner_group = identity_groups[best_identity]
                                        found_names = [res['hk_name'] for res in winner_group if res['hk_name'] and is_chinese(res['hk_name'])]
                                        if found_names:
                                            name_counts = Counter(found_names)
                                            logging.info(f"對 '{original_name}' ({best_identity}) 的維基搜索投票結果: {dict(name_counts)}")
                                            most_common = name_counts.most_common(1)
                                            if most_common:
                                                final_hk_name = most_common[0][0]
                                                logging.info(f"對 '{original_name}' 的維基搜索確定最終結果: '{final_hk_name}' (票數: {most_common[0][1]})。")
                                        else:
                                            logging.info(f"對 '{original_name}' 的維基搜索找到了對應頁面 ({best_identity})，但未能從中提取任何有效的中文譯名。")
                else:
                    logging.info(f"對 '{original_name}' 的維基百科搜索在所有目標語言中均無有效結果。")

        except Exception as e:
            logging.error(f"執行 wiki 搜索時發生未預期的頂層錯誤: {e}", exc_info=True)
            if item_id:
                _save_to_wiki_cache(item_id, media_type, None, 'FAILED', original_name, year)
            return None

    # --- 步驟 4: 將結果存入快取 ---
    if item_id:
        if final_hk_name:
            _save_to_wiki_cache(item_id, media_type, final_hk_name, 'FOUND', original_name, year)
        else:
            _save_to_wiki_cache(item_id, media_type, None, 'NOT_FOUND', original_name, year)

    return final_hk_name


# ---------------------------------------------------------------------------
# 【2026-11】人物香港譯名：用 TMDB 提供嘅 Wikidata ID 直接搵（唔使用名搜尋，唔會撞名）
#   Wikidata ID → 中文維基頁面名稱 → 中文維基 zh-hk 顯示標題（會套用頁面嘅 NoteTA 香港譯名，
#   例如 斯蒂芬·斯皮尔伯格 → 史提芬·史匹堡）。粵語維基頁名只記入 log 作參考。
#   每人 2 次請求；結果存 wiki_cache（media_type = WIKIDATA_CACHE_TYPE），搵唔到嘅 90 日後再查。
# ---------------------------------------------------------------------------
WIKIDATA_CACHE_TYPE = 'person_qid'
WIKIDATA_NOT_FOUND_RECHECK_DAYS = 90


def _wikidata_cache_lookup(item_id, cache_type=None):
    """回傳 (hk_name 或 None, 是否要重新查詢)。cache_type 預設 WIKIDATA_CACHE_TYPE（中文名）。"""
    cache_type = cache_type or WIKIDATA_CACHE_TYPE
    try:
        if not os.path.exists(DB_FILE):
            return None, True
        conn = sqlite3.connect(DB_FILE)
        try:
            row = conn.execute("SELECT hk_name, status, timestamp FROM wiki_cache WHERE item_id = ? AND media_type = ?",
                               (str(item_id), cache_type)).fetchone()
        finally:
            conn.close()
    except Exception as e:
        logging.warning(f"[Wikidata] 讀取快取失敗: {e}")
        return None, True
    if not row:
        return None, True
    hk_name, status, ts = row
    if status == 'FOUND' and hk_name:
        return hk_name, False
    try:
        age_days = (datetime.now() - datetime.fromisoformat(str(ts))).days
    except ValueError:
        age_days = WIKIDATA_NOT_FOUND_RECHECK_DAYS
    return None, age_days >= WIKIDATA_NOT_FOUND_RECHECK_DAYS


def _wiki_api_get(url, params):
    """回傳 JSON；429 時標記限速並回傳 None（唔寫快取，下次再試）。"""
    r = requests.get(url, params=params, headers={'User-Agent': WIKI_USER_AGENT}, timeout=20)
    if r.status_code == 429:
        _mark_rate_limited()
        logging.warning(f"[Wikidata] {url} 回覆 429（限速）。")
        return None
    r.raise_for_status()
    return r.json()


def lookup_hk_name_by_wikidata(item_id, wikidata_id, original_name=None):
    """用 Wikidata ID 搵人物嘅中文維基 zh-hk 譯名。搵唔到或出錯回傳 None。"""
    _reset_rate_limit_flag()
    _init_db()
    if not wikidata_id or not re.fullmatch(r'Q\d+', str(wikidata_id)):
        return None
    cached, need_query = _wikidata_cache_lookup(item_id) if item_id else (None, True)
    if cached or not need_query:
        logging.info(f"[Wikidata] 快取命中 '{original_name}' ({wikidata_id}): {cached or '搵唔到'}")
        return cached
    try:
        j = _wiki_api_get('https://www.wikidata.org/w/api.php', {
            'action': 'wbgetentities', 'ids': wikidata_id, 'props': 'sitelinks',
            'sitefilter': 'zhwiki|zh_yuewiki', 'format': 'json'})
        if j is None:
            return None
        links = j.get('entities', {}).get(wikidata_id, {}).get('sitelinks', {})
        zh_title = links.get('zhwiki', {}).get('title')
        yue_title = links.get('zh_yuewiki', {}).get('title')
        hk_name = None
        if zh_title:
            j = _wiki_api_get('https://zh.wikipedia.org/w/api.php', {
                'action': 'parse', 'page': zh_title, 'prop': 'displaytitle', 'variant': 'zh-hk',
                'redirects': 1, 'format': 'json'})
            if j is None:
                return None
            hk_name = _clean_wiki_name(j.get('parse', {}).get('displaytitle', '')) or None
            if hk_name and not is_chinese(hk_name):
                hk_name = None
    except Exception as e:
        logging.warning(f"[Wikidata] 查詢 '{original_name}' ({wikidata_id}) 失敗: {e}")
        return None

    if hk_name:
        yue_note = ''
        if yue_title:
            yue_clean = re.sub(r'[\s·・]', '', _clean_wiki_name(yue_title))
            yue_note = '，粵語維基一致' if yue_clean and yue_clean in re.sub(r'[\s·・]', '', hk_name) \
                else f"，粵語維基頁名為「{yue_title}」"
        logging.info(f"[Wikidata] '{original_name}' ({wikidata_id}) 中文維基香港譯名: '{hk_name}'{yue_note}")
    else:
        logging.info(f"[Wikidata] '{original_name}' ({wikidata_id}) 冇中文維基香港譯名。")
    if item_id:
        _save_to_wiki_cache(item_id, WIKIDATA_CACHE_TYPE, hk_name, 'FOUND' if hk_name else 'NOT_FOUND',
                            original_name, None)
    return hk_name


# 【2026-11】B：日文漢字名。日本人（同用漢字名嘅韓國人）搵唔到中文名、TMDB 又冇漢字名時，
# 用 Wikidata 日文標籤 / 日文維基頁名，只接受全漢字（「々」換成前一個字），新字體轉繁體。
WIKIDATA_JA_CACHE_TYPE = 'person_qid_ja_kanji'


# 人名專用：日文新字體 → 繁體（廣 / 關 / 邊 / 澤…）。用 OpenCC 官方 JPVariants.txt 反向（只限一對一），加少量補充。
# 一般文字唔用呢個表：主流程嘅 jp2t 刻意只得幾個字，因為新字體好多同簡體重疊（見 AI_HANDOFF）。
_NAME_SHINJITAI_EXTRA = {'啓': '啟', '﨑': '崎', '髙': '高', '𠮷': '吉', '濵': '濱', '槇': '槙'}
_name_shinjitai_map = None


def _shinjitai_name_to_trad(name):
    global _name_shinjitai_map
    if _name_shinjitai_map is None:
        rev = {}
        path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            'opencc_data', 'dictionary', 'JPVariants.txt')
        try:
            with open(path, encoding='utf-8') as f:
                for line in f:
                    parts = line.rstrip('\n').split('\t')
                    if len(parts) == 2 and not line.startswith('#'):
                        for jp in parts[1].split():
                            rev.setdefault(jp, set()).add(parts[0])
        except OSError as e:
            logging.warning(f"[Wikidata] 讀取 JPVariants.txt 失敗（{e}），日文漢字名唔轉繁體。")
        _name_shinjitai_map = {jp: next(iter(t)) for jp, t in rev.items() if len(t) == 1 and jp != next(iter(t))}
        _name_shinjitai_map.update(_NAME_SHINJITAI_EXTRA)
    return ''.join(_name_shinjitai_map.get(c, c) for c in name)


def _to_kanji_name(name):
    name = re.sub(r'\s*[(（][^)）]*[)）]\s*$', '', str(name or '')).strip()   # 「渡辺謙 (俳優)」
    name = re.sub(r'[\s\u3000]', '', name)
    name = re.sub(r'(.)々', r'\1\1', name)                                   # 佐々木 → 佐佐木
    name = _shinjitai_name_to_trad(name)                                    # 平田広明 → 平田廣明
    return name if 2 <= len(name) <= 8 and is_all_hanzi(name) else None


def lookup_ja_kanji_name_by_wikidata(item_id, wikidata_id, original_name=None):
    """用 Wikidata ID 攞人物嘅日文漢字名（標籤優先，其次日文維基頁名）。唔係全漢字就回傳 None。"""
    _reset_rate_limit_flag()
    _init_db()
    if not wikidata_id or not re.fullmatch(r'Q\d+', str(wikidata_id)):
        return None
    cached, need_query = _wikidata_cache_lookup(item_id, WIKIDATA_JA_CACHE_TYPE) if item_id else (None, True)
    if cached or not need_query:
        logging.info(f"[Wikidata] 日文漢字名快取命中 '{original_name}' ({wikidata_id}): {cached or '冇'}")
        return cached
    try:
        j = _wiki_api_get('https://www.wikidata.org/w/api.php', {
            'action': 'wbgetentities', 'ids': wikidata_id, 'props': 'labels|sitelinks',
            'languages': 'ja', 'sitefilter': 'jawiki', 'format': 'json'})
        if j is None:
            return None
        ent = j.get('entities', {}).get(wikidata_id, {})
        label = ent.get('labels', {}).get('ja', {}).get('value')
        title = ent.get('sitelinks', {}).get('jawiki', {}).get('title')
    except Exception as e:
        logging.warning(f"[Wikidata] 查詢日文名 '{original_name}' ({wikidata_id}) 失敗: {e}")
        return None
    kanji = _to_kanji_name(label) or _to_kanji_name(title)
    if kanji:
        logging.info(f"[Wikidata] '{original_name}' ({wikidata_id}) 日文漢字名: '{kanji}'（標籤 '{label}'）")
    else:
        logging.info(f"[Wikidata] '{original_name}' ({wikidata_id}) 冇全漢字日文名（標籤 '{label}'，頁名 '{title}'）。")
    if item_id:
        _save_to_wiki_cache(item_id, WIKIDATA_JA_CACHE_TYPE, kanji, 'FOUND' if kanji else 'NOT_FOUND',
                            original_name, None)
    return kanji

