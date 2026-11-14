import httpx  # 【修改】替換 requests
import asyncio  # 【新增】非同步處理
import logging
import time
import random
import re
import traceback
import unicodedata
import os
import requests # 【新增】為同步函數導入 requests
from requests.adapters import HTTPAdapter # 【新增】為同步函數導入
from urllib3.util.retry import Retry # 【新增】為同步函數導入
from bs4 import BeautifulSoup
from urllib.parse import quote
from others.wiki_search import _clean_wiki_name, WIKI_USER_AGENT, _wikidata_cache_lookup
from importers.tmdb_uploader import extract_tv_show_id, extract_movie_id
from others.excel_handler import (
    save_series_info_to_excel, save_season_info_to_excel,
    save_episode_info_to_excel, save_movie_info_to_excel, check_create_excel,
    save_english_site_info_to_excel, save_unavailable_video_info_to_excel,
    replace_tw_hk_names as apply_tw_hk_name_replacements,
    save_skipped_identical_video_info_to_excel, save_skipped_identical_to_bulk_excel,
    save_processed_episode_info_to_excel, check_and_release_excel_control,
    save_actor_info_to_excel as save_person_info_to_excel
)
from others.analysis_excel import is_generic_title, clean_generic_title
# 【修改】導入整個 utils 模組，以便存取全局中斷標誌
from others import utils
from others.utils import (
    is_chinese, is_english, is_simplified_chinese,
    normalize_text, normalize_to_hk,
    add_spaces_between_en_zh_num,
    convert_runtime_to_minutes, process_translatable_item, NO_DESCRIPTION_TEXT,
    is_all_hanzi, is_valid_zh_hk_title, compare_with_zh_hk
)
from others.tw_hk_name import select_tw_hk_sheet
from others import bangumi
from others.wiki_queue import search_wiki_for_hk_name_queued as search_wiki_for_hk_name, lookup_hk_name_by_wikidata_queued, lookup_ja_kanji_name_by_wikidata_queued

# 【修改】移除會導致循環導入的 import
# from tmdb_importer import _PROGRAM_INTERRUPTED

"""
graph TD
    subgraph "tmdb_extractor_api.py 完整決策流程"
        direction TB
        B1[開始: _async_main(url)] --> B2{判斷 URL 類型<br>(/tv/, /movie/, /person/)};

        B2 -- "/tv/" --> B_TV[<b>TV Show 處理路徑</b>];
        B2 -- "/movie/" --> B_MOVIE[<b>Movie 處理路徑</b>];
        B2 -- "/person/" --> B_PERSON[<b>Person 處理路徑</b>];
        
        subgraph B_TV [TV Show 處理路徑]
            direction TB
            TV1["1. 提取 TV Show 主體資訊<br><b>extract_tv_show_details()</b>"] --> TV1_LOGIC["(執行 B.1 核心提取 & 決策邏輯)"];
            TV1_LOGIC --> TV2["2. 提取所有演職員資訊<br><b>extract_person_details_for_credits()</b>"];
            TV2 --> TV2_LOGIC["(對<b>每一位</b>演職員執行 B.1 核心提取 & 決策邏輯)"];
            TV2_LOGIC --> TV3["3. 提取所有季度資訊<br><b>extract_season_details()</b>"];
            TV3 --> TV3_LOGIC["(對<b>每一個</b>季度執行 B.1 核心提取 & 決策邏輯)"];
            TV3_LOGIC --> TV4["4. 提取所有集數資訊<br><b>extract_episode_details()</b>"];
            TV4 --> TV4_LOGIC["(對<b>每一集</b>執行 B.1 核心提取 & 決策邏輯)"];
            TV4_LOGIC --> TV_END[結束 TV 處理];
        end

        subgraph B_MOVIE [Movie 處理路徑]
            direction TB
            M1["1. 提取 Movie 主體資訊<br><b>extract_movie_details()</b>"] --> M1_LOGIC["(執行 B.1 核心提取 & 決策邏輯)"];
            M1_LOGIC --> M2["2. 提取所有演職員資訊<br><b>extract_person_details_for_credits()</b>"];
            M2 --> M2_LOGIC["(對<b>每一位</b>演職員執行 B.1 核心提取 & 決策邏輯)"];
            M2_LOGIC --> M_END[結束 Movie 處理];
        end

        subgraph B_PERSON [Person 處理路徑 (單獨)]
             direction TB
             P1["提取 Person 資訊<br><b>extract_person_details()</b>"] --> P1_LOGIC["(執行 B.1 核心提取 & 決策邏輯)"];
             P1_LOGIC --> P_END[結束 Person 處理];
        end
        
        TV_END & M_END & P_END --> B6{綜合所有部分的結果<br>有任何內容需要寫入 Excel 嗎?};
        B6 -- 是 --> B7[保存至 video_detail.xlsx];
        B6 -- 否 --> B8["日誌記錄: 內容無需更新<br>刪除空的快取 Excel 檔"];
        B7 & B8 --> B_FINAL_END[結束提取];
    end

    subgraph B1_LOGIC [B.1 核心提取 & 決策邏輯 (被重複調用)]
        style B1_LOGIC fill:#f2f2f2,stroke:#333
        direction TB
        B1_1["<b>get_api_data_for_item()</b><br>並行獲取多語言資料"];
        B1_1 --> B1_2{是人名且原始名稱為中文?};
        B1_2 -- 是 --> B1_SKIP_WIKI[跳過 Wiki 搜索];
        B1_2 -- 否 --> B1_3["<b>執行 Wiki 搜索</b><br>(參考 E. Wiki 搜索子圖)"];
        B1_3 & B1_SKIP_WIKI --> B1_4{"<b>處理 Title/Name</b>"};
        B1_1 --> B1_5{"<b>處理 Description/Bio</b>"};

        subgraph B4_detail [Title/Name 處理細節]
            B4_1["<b>1. 搜尋中文候選值 (優先)</b><br>- 維基譯名 (若為人名)<br>- zh-HK<br>- zh-TW / zh-CN<br>- original_name (若為中文)<br>- 日/韓漢字"]
            B4_1 --> B4_1_CHECK{找到中文候選值?}
            B4_1_CHECK -- 是 --> B4_2["<b>3. 執行文字正規化</b><br>(參考 D. 文字正規化子圖)"]
            B4_1_CHECK -- 否 --> B4_1_FALLBACK["<b>2. 回退搜尋英文候選值</b><br>- zh-TW (若為英文)<br>- zh-CN (若為英文)<br>- zh-HK (若為英文)<br>- en-US / original_name"]
            B4_1_FALLBACK --> B4_2
            
            subgraph B4_3 [4. 套用 tw_hk_name.xlsx 最終規則]
                direction TB
                B4_3_1["根據<b>上下文</b>選擇映射表"] --> B4_3_2{使用關鍵詞 Sheet?};
                B4_3_2 -- 否 --> B4_3_3["<b>僅載入 'Default' Sheet 規則</b>"];
                B4_3_2 -- 是 --> B4_3_4["載入對應的關鍵詞 Sheet 規則"];
                B4_3_3 & B4_3_4 --> B4_3_5[執行譯名替換];
            end

            B4_2 --> B4_3;
            B4_3 --> B4_4{"5. 判斷是否應寫入 Excel"};
            B4_4 -- 是 --> B4_5[準備<b>最終標題/名稱</b><br>和 <b>perfect='n'/'y'</b> 旗標];
            B4_4 -- 否 --> B4_6[不寫入此欄位];
        end
        
        B1_4 --> B4_detail

        subgraph B5_detail [Description/Bio 處理細節]
            B5_1[選擇候選值 (<b>高優先級</b>)<br>1. zh-HK<br>2. zh-TW...];
            B5_1 --> B5_2["<b>執行文字正規化</b><br>(參考 D. 文字正規化子圖)"];

            subgraph B5_3 [套用 tw_hk_name.xlsx 最終規則]
                direction TB
                B5_3_1["根據<b>上下文</b>選擇映射表"] --> B5_3_2{使用關鍵詞 Sheet?};
                B5_3_2 -- 否 --> B5_3_3["<b><font color=red>合併</font> 'Movie name' 和 'Default' Sheet 規則<br>('Movie name' 優先)</b>"];
                B5_3_2 -- 是 --> B5_3_4["載入對應的關鍵詞 Sheet 規則"];
                B5_3_3 & B5_3_4 --> B5_3_5[執行譯名替換];
            end

            B5_2 --> B5_3;
            B5_3 --> B5_4{"判斷是否應寫入 Excel"};
            B5_4 -- 是 --> B5_5[準備<b>最終描述/簡介</b><br>和 <b>perfect='n'/'y'</b> 旗標];
            B5_4 -- 否 --> B5_6[不寫入此欄位];
        end
        
        B1_5 --> B5_detail
    end
"""

# 【新增】並行請求數量限制
MAX_CONCURRENT_REQUESTS = 3

# 【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
# 此處為 httpx (非同步請求) 和 requests (同步請求) 配置了重試邏輯。
RETRY_COUNT = 50 # 為 httpx 設定一個較大的重試次數以模擬無限重試
RETRY_BACKOFF = 1

# 【移除】不再需要此處的獨立 converter
# converter = opencc.OpenCC('s2hk.json')

def is_japanese(text):
    """檢查文字是否包含日文字元（平假名、片假名或常用漢字）"""
    if not text: return False
    # CJK Unified Ideographs (Kanji) + Hiragana + Katakana
    return bool(re.search(r'[\u3040-\u30ff\u4e00-\u9faf]', text))

def _is_western_person_name(original_name, zh_tw_name, zh_cn_name):
    """【2026-11】判斷係咪西方人名（要用 Wikidata 搵香港譯名）。
    有中文名：含「·」「-」或者多過 4 個字（音譯名）→ 西方名；中日韓名（例如 梁朝偉、宋康昊）唔係。
    冇中文名：原名冇假名 / 韓文 → 西方名。"""
    zh = next((n for n in (zh_tw_name, zh_cn_name, original_name) if n and is_chinese(n)), '')
    if not zh:
        return not re.search(r'[぀-ヿ가-힯]', original_name or '')
    return bool(re.search(r'[·・•‧\-]', zh)) or len(zh) > 4


def _looks_like_jk_name(name):
    """【2026-11】日韓人名嘅寫法：有平假名，或者漢字夾假名；韓文要係 2–4 個字、冇空格（例如 최하나）。
    純片假名（トニー・リップ）同有空格嘅韓文（톰 크루즈）多數係西方人嘅音譯，唔算。"""
    name = (name or "").strip()
    if not name:
        return False
    if re.search(r'[぀-ゟ]', name):                                   # 平假名
        return True
    if re.search(r'[゠-ヿ]', name) and re.search(r'[㐀-鿿]', name):   # 漢字 + 片假名
        return True
    return bool(re.fullmatch(r'[가-힯]{2,4}', name))


def _is_jk_person(data_dict):
    """【2026-11】判斷係咪日本 / 韓國人（後備維基搜尋只保留畀呢類，歷史上命中幾乎全部係日韓名）。
    原名有假名 / 漢字 / 韓文；或者日韓譯名、別名似日韓人名（_looks_like_jk_name）；或者出生地喺日本 / 韓國。
    【2026-11-13】以前日韓譯名 / 別名有任何假名 / 韓文就算，會將有片假名譯名嘅西方人（Tony Lip 等）當日韓人。"""
    original = data_dict.get("original_title_or_name") or ""
    names = [v for k, v in (data_dict.get("all_name_translations") or {}).items()
             if v and k.split('-')[0] in ('ja', 'ko')]
    names += [a for a in (data_dict.get("also_known_as") or []) if isinstance(a, str)]
    if is_japanese(original) or re.search(r'[가-힯]', original):
        return True
    pob = (data_dict.get("place_of_birth") or "").lower()
    if any(k in pob for k in ('japan', 'korea', '日本', '韓國', '韩国', '한국', '대한민국', '朝鮮')):
        return True
    # 【2026-11-14】TMDB 已有普通中文名（例如 袁振洋）→ 日韓譯名 / 別名（원진양）只係外文寫法，唔當日韓人
    if _plain_chinese_names(data_dict):
        return False
    return any(_looks_like_jk_name(n) for n in names)


_PLAIN_CHINESE_NAME_RE = re.compile(r'[㐀-鿿]{2,4}')


def _plain_chinese_names(data_dict):
    """【2026-11-14】TMDB zh-HK / TW / CN / 原名入面嘅普通中文名（2–4 個漢字、冇「·」），例如 袁振洋、曾志偉。"""
    names = [data_dict.get(k) or "" for k in ("zh_hk_name", "zh_tw_name", "zh_cn_name", "original_title_or_name")]
    return [n.strip() for n in names if _PLAIN_CHINESE_NAME_RE.fullmatch(n.strip())]


def _is_other_name_than_tmdb(wiki_name, data_dict):
    """【2026-11-14】TMDB 已有普通中文名，而 wiki 名同佢哋全部都有一半以上嘅字唔同 → wiki 頁用咗另一個名
    （法號 / 藝名 / 筆名，例如 袁振洋 嘅頁叫「釋傳般」），唔係呢個名嘅香港寫法。只差一兩個字（宋康浩 → 宋康昊）唔算。"""
    plain = _plain_chinese_names(data_dict)
    if not plain or not wiki_name:
        return False
    s2t = utils.opencc_converters.get('s2t')
    to_t = (lambda s: s2t.convert(s)) if s2t else (lambda s: s)
    w = set(to_t(wiki_name))
    return not any(len(w & set(to_t(n))) * 2 >= len(n) for n in plain)


def _has_chinese_person_name(data_dict):
    """【2026-11】TMDB 已有可用嘅中文 / 漢字人名（zh-HK / TW / CN、中文原名、中文別名、全漢字 ja-JP）。"""
    names = [data_dict.get(k) or "" for k in ("zh_hk_name", "zh_tw_name", "zh_cn_name", "original_title_or_name")]
    names += [a for a in (data_dict.get("also_known_as") or []) if isinstance(a, str)]
    if any(is_chinese(n) for n in names):
        return True
    return is_all_hanzi((data_dict.get("all_name_translations") or {}).get("ja-JP") or "")


def _person_ja_value(data_dict):
    """【2026-11】人名嘅日文候選：TMDB ja-JP 係全漢字就用佢，否則用 Wikidata 日文漢字名（B）。"""
    ja = (data_dict.get("all_name_translations") or {}).get("ja-JP")
    if ja and is_all_hanzi(ja):
        return ja
    return data_dict.get("wiki_ja_kanji_name") or ja


def is_korean(text):
    """檢查文字是否包含韓文字元（Hangul）"""
    if not text: return False
    # Hangul Syllables
    return bool(re.search(r'[\uac00-\ud7af]', text))

_ORG_SUFFIXES = (
    '公司', '有限', '股份', '集團', '企業', '事業',
    '基金', '基金會', '協會', '學會', '公會',
    '委員會', '理事會', '機構', '組織',
    '出版社', '出版集團', '傳媒', '媒體', '廣播', '電視台', '電視臺',
    '大學', '學院', '學校', '研究所', '研究院', '實驗室',
    '政府', '部門', '局', '廳', '署', '院', '部',
    '銀行', '保險', '投資', '地產',
)

def _is_likely_org_name(name: str) -> bool:
    """判斷一個名稱是否可能是公司或組織，而非個人姓名。"""
    if not name:
        return False
    return any(name.endswith(suffix) or suffix in name for suffix in _ORG_SUFFIXES)

def extract_tmdb_id(url):
    try:
        if "/tv/" in url:
            match = re.search(r'/tv/([^/?]+)', url)
            if not match:
                logging.error(f"無效的 TMDB TV URL: {url}")
                return None, None, None
            tmdb_id = match.group(1).split('-')[0]
            season_match = re.search(r'/season/(\d+)', url)
            season_number = int(season_match.group(1)) if season_match else None
            episode_match = re.search(r'/episode/(\d+)', url)
            episode_number = int(episode_match.group(1)) if episode_match else None
            return tmdb_id, season_number, episode_number
        elif "/movie/" in url:
            match = re.search(r'/movie/([^/?]+)', url)
            if match:
                return match.group(1).split('-')[0], None, None
            logging.error(f"無效的 TMDB Movie URL: {url}")
            return None, None, None
        elif "/person/" in url:
            match = re.search(r'/person/([^/?]+)', url)
            if match:
                return match.group(1).split('-')[0], None, None
            logging.error(f"無效的 TMDB Person URL: {url}")
            return None, None, None
        logging.error(f"無效的 TMDB URL: {url}")
        return None, None, None
    except Exception as e:
        logging.error(f"提取 TMDB ID 失敗: {url}, 錯誤: {e}")
        return None, None, None

async def get_api_data_for_item(client, semaphore, api_key, base_url_tmdb, endpoint_suffix, item_is_person=False, preferred_language="en", fallback_language=None, context_for_logging="Item"):
    """
    【修改 2025-08-25】在呼叫維基百科搜尋時，傳入 item_id 以支援更可靠的快取機制。
    【修正 2025-08-25】修正了因缺少 traceback 導入而導致的 NameError。
    【使用者要求 2025-09-01】如果藝人的原始名稱 (original_name) 本身就是中文（例如 "于骞"），
    則沒有必要在維基百科上搜索譯名。此修改會跳過對這類姓名的維基搜索，
    直接交由後續流程處理簡繁轉換和正規化，以提高效率並避免不必要的網絡請求。
    【使用者要求 2025-09-18】在執行網絡請求前檢查中斷信號。
    【修改】將 _PROGRAM_INTERRUPTED 的檢查改為 utils._PROGRAM_INTERRUPTED。
    【修改】加入無限重試機制 (while True loop)，處理 429 限流與網絡中斷。
    """
    if utils._PROGRAM_INTERRUPTED:
        logging.warning(f"檢測到中斷信號，取消 API 請求: {endpoint_suffix}")
        return {} # 返回空字典以表示操作被取消

    headers = {"Authorization": f"Bearer {api_key}"}
    
    data_dict = {
        "zh_hk_title": "", "zh_hk_desc": "", "zh_hk_runtime": None, "zh_hk_name": "", "zh_hk_bio": "",
        "zh_tw_title": "", "zh_tw_desc": "", "zh_tw_runtime": None, "zh_tw_name": "", "zh_tw_bio": "",
        "zh_cn_title": "", "zh_cn_desc": "", "zh_cn_runtime": None, "zh_cn_name": "", "zh_cn_bio": "",
        "en_title": "", "en_desc": "", "en_name": "", "en_bio": "",
        "original_title_or_name": "", 
        "year": None,
        "original_language": "en",
        "place_of_birth": None,
        "also_known_as": [],
        "all_title_translations": {},
        "all_name_translations": {},
        "wiki_zh_hk_name": ""
    }
    api_raw_data_log = {}

    # 【新增】內部 helper function 用於無限重試請求
    async def _fetch_with_infinite_retry(url, context):
        while True:
            if utils._PROGRAM_INTERRUPTED: return None
            try:
                async with semaphore:
                    # 觸發條件：無限重試直到成功或遇到 404
                    response = await client.get(url, headers=headers, timeout=20)
                
                if response.status_code == 200:
                    return response
                elif response.status_code == 404:
                    logging.warning(f"{context}: 資源不存在 (404) - {url}")
                    return response # 回傳 response 讓外層處理
                elif response.status_code == 429:
                    retry_after = int(response.headers.get("Retry-After", 5))
                    logging.warning(f"{context}: API 限流 (429)，等待 {retry_after} 秒後重試...")
                    await asyncio.sleep(retry_after + 1)
                else:
                    logging.warning(f"{context}: 請求失敗 (Status {response.status_code})，3 秒後重試 - {url}")
                    await asyncio.sleep(3)
            except (httpx.RequestError, httpx.TimeoutException) as e:
                logging.warning(f"{context}: 網絡錯誤 ({e})，3 秒後重試 - {url}")
                await asyncio.sleep(3)
            except Exception as e:
                logging.error(f"{context}: 未知錯誤 ({e})，等待重試 - {url}")
                await asyncio.sleep(5)

    details_data_raw = None 
    try:
        # 【修改】使用無限重試邏輯獲取基本資料
        details_url = f"{base_url_tmdb}{endpoint_suffix}?language=en"
        if item_is_person:
            # 【2026-11】同一個請求攞埋 Wikidata ID（用嚟直接搵中文維基香港譯名）
            details_url += "&append_to_response=external_ids"
        details_resp = await _fetch_with_infinite_retry(details_url, f"{context_for_logging} [Base]")
        
        if utils._PROGRAM_INTERRUPTED: return {}
        if details_resp and details_resp.status_code == 200:
            details_data_raw = details_resp.json()
            logging.debug(f"{context_for_logging}: Successfully fetched base details (in English) for {endpoint_suffix}")
        else:
            logging.warning(f"{context_for_logging}: Failed to fetch base details for {endpoint_suffix}")

        if details_data_raw:
            if not item_is_person:
                data_dict["en_title"] = details_data_raw.get("title") or details_data_raw.get("name") or ""
                data_dict["en_desc"] = details_data_raw.get("overview", "")
                data_dict["original_title_or_name"] = details_data_raw.get("original_title") or details_data_raw.get("original_name") or data_dict["en_title"]
                data_dict["original_language"] = details_data_raw.get("original_language", "en")
                
                release_date = details_data_raw.get("release_date") or details_data_raw.get("first_air_date") or ""
                if release_date and isinstance(release_date, str) and len(release_date) >= 4:
                    data_dict["year"] = release_date[:4]
                
                api_raw_data_log["original_language"] = data_dict["original_language"]
                api_raw_data_log["en_title_details"] = data_dict["en_title"]
                api_raw_data_log["en_desc_details"] = data_dict["en_desc"]
                api_raw_data_log["original_title"] = data_dict["original_title_or_name"]
                api_raw_data_log["year"] = data_dict["year"]
                
                runtime_val = None
                if "runtime" in details_data_raw: 
                    runtime_val = details_data_raw.get("runtime")
                elif "episode_run_time" in details_data_raw and details_data_raw["episode_run_time"]: 
                    runtime_val = details_data_raw["episode_run_time"][0] if details_data_raw["episode_run_time"] else None
                
                converted_runtime = convert_runtime_to_minutes(runtime_val)
                data_dict["zh_hk_runtime"] = converted_runtime
                data_dict["zh_tw_runtime"] = converted_runtime
                data_dict["zh_cn_runtime"] = converted_runtime
                api_raw_data_log["runtime_details"] = str(runtime_val) if runtime_val is not None else ""
            else: # item_is_person
                data_dict["en_name"] = details_data_raw.get("name", "")
                data_dict["en_bio"] = details_data_raw.get("biography", "")
                data_dict["original_title_or_name"] = details_data_raw.get("name", "")
                data_dict["place_of_birth"] = details_data_raw.get("place_of_birth")
                data_dict["also_known_as"] = details_data_raw.get("also_known_as", [])
                data_dict["wikidata_id"] = (details_data_raw.get("external_ids") or {}).get("wikidata_id") or ""

                birthday = details_data_raw.get("birthday") or ""
                if birthday and isinstance(birthday, str) and len(birthday) >= 4:
                    data_dict["year"] = birthday[:4]
                    data_dict["birthday"] = birthday   # 【2026-11】Bangumi 用嚟核對同名人物

                api_raw_data_log["en_name_details"] = data_dict["en_name"]
                api_raw_data_log["en_bio_details"] = data_dict["en_bio"]
                api_raw_data_log["original_name"] = data_dict["original_title_or_name"]
                api_raw_data_log["year"] = data_dict["year"]
                api_raw_data_log["place_of_birth"] = data_dict["place_of_birth"]
                api_raw_data_log["also_known_as"] = str(data_dict["also_known_as"])
        else:
            logging.warning(f"{context_for_logging}: Failed to get any details_data_raw for {endpoint_suffix}")

        if utils._PROGRAM_INTERRUPTED: return {}
        
        # 【修改】使用無限重試邏輯獲取翻譯資料
        translations_url = f"{base_url_tmdb}{endpoint_suffix}/translations"
        translations_resp = await _fetch_with_infinite_retry(translations_url, f"{context_for_logging} [Trans]")

        if utils._PROGRAM_INTERRUPTED: return {}
        if translations_resp and translations_resp.status_code == 200:
            translations_data = translations_resp.json().get("translations", [])
            for t in translations_data:
                lang_code = t.get("iso_639_1", "")
                country_code = t.get("iso_3166_1", "")
                data = t.get("data", {})
                
                full_lang_code = f"{lang_code}-{country_code}" if country_code else lang_code

                if item_is_person:
                    name_val = (data.get("name") or "").strip() 
                    bio_val = (data.get("biography") or "").strip() 
                    api_raw_data_log[f"{full_lang_code}_name_trans"] = name_val
                    api_raw_data_log[f"{full_lang_code}_bio_trans"] = bio_val
                    
                    if name_val:
                        data_dict['all_name_translations'][full_lang_code] = name_val

                    if lang_code == "zh":
                        if country_code == "HK": data_dict["zh_hk_name"], data_dict["zh_hk_bio"] = name_val, bio_val
                        elif country_code == "TW": data_dict["zh_tw_name"], data_dict["zh_tw_bio"] = name_val, bio_val
                        elif country_code == "CN": data_dict["zh_cn_name"], data_dict["zh_cn_bio"] = name_val, bio_val
                else: # not item_is_person
                    title_val = (data.get("title") or data.get("name") or "").strip() 
                    desc_val = (data.get("overview") or "").strip() 
                    api_raw_data_log[f"{full_lang_code}_title_trans"] = title_val
                    api_raw_data_log[f"{full_lang_code}_desc_trans"] = desc_val
                    
                    if title_val:
                        data_dict['all_title_translations'][full_lang_code] = title_val

                    if lang_code == "zh":
                        if country_code == "HK": data_dict["zh_hk_title"], data_dict["zh_hk_desc"] = title_val, desc_val
                        elif country_code == "TW": data_dict["zh_tw_title"], data_dict["zh_tw_desc"] = title_val, desc_val
                        elif country_code == "CN": data_dict["zh_cn_title"], data_dict["zh_cn_desc"] = title_val, desc_val
        else:
            status_code = translations_resp.status_code if translations_resp else "Unknown"
            logging.warning(f"{context_for_logging}: 獲取翻譯數據失敗 for {endpoint_suffix}: 狀態碼 {status_code}")

        if item_is_person:
            if utils._PROGRAM_INTERRUPTED: return {}
            item_id_match = re.search(r'/(\d+)$', endpoint_suffix)
            item_id_for_cache = item_id_match.group(1) if item_id_match else None

            # 【使用者要求 2025-09-01】如果藝人的原始名稱 (original_name) 本身就是中文，
            # 則沒有必要在維基百科上搜索譯名，因為它很可能沒有一個不同的「香港譯名」。
            # 直接進行後續的簡繁轉換和正規化即可。
            # 這可以避免對像「于骞」這樣的中文姓名進行不必要的網絡請求和搜索。
            # 【使用者要求 2026-04-13】擴展此邏輯：即使 original_name 是拼音（如 "Kang Xuan"），
            # 只要 zh-TW 或 zh-CN 已有中文名，且原名並非日文，就同樣跳過 Wiki。
            # 原因：中文人名直接做 OpenCC 簡繁轉換即可，Wiki 反而容易產生錯誤重定向。
            # 日文名依然需要 Wiki（如 "渡辺謙" → 需要找香港譯名）。
            original_name = data_dict.get("original_title_or_name") or ""
            zh_tw_name = data_dict.get("zh_tw_name") or ""
            zh_cn_name = data_dict.get("zh_cn_name") or ""
            has_chinese_name = (
                is_chinese(original_name) or
                is_chinese(zh_tw_name) or
                is_chinese(zh_cn_name)
            )
            original_is_japanese = bool(original_name and is_japanese(original_name))
            # 【2026-11】西方名：就算已有 TW / CN 譯名，都先用 TMDB 嘅 Wikidata ID 搵中文維基 zh-hk 譯名
            # （例如 Tom Cruise：湯姆·克魯斯 → 湯·告魯斯）。搵唔到就照舊用以下流程。
            # 【2026-11】A：日韓人名都先用 Wikidata ID（以前只用名稱搜尋，成功率極低）。
            # 選項「維基譯名取代現有香港人名」= n 而 TMDB 已有中文香港人名 → 唔使查（結果唔會用）
            is_jk = _is_jk_person(data_dict)
            if (data_dict.get("wikidata_id") and (is_jk or _is_western_person_name(original_name, zh_tw_name, zh_cn_name))
                    and (utils.wiki_name_replaces_existing() or not is_chinese(data_dict.get("zh_hk_name") or ""))):
                hk_by_qid = await asyncio.to_thread(lookup_hk_name_by_wikidata_queued, item_id_for_cache,
                                                    data_dict["wikidata_id"], original_name,
                                                    [original_name, data_dict.get("en_name")],
                                                    data_dict.get("also_known_as"))
                if utils._PROGRAM_INTERRUPTED: return {}
                if hk_by_qid and _is_likely_org_name(hk_by_qid):
                    logging.warning(f"Wikidata 返回的名稱 '{hk_by_qid}' 疑似公司/組織名，已棄用。")
                elif hk_by_qid and _is_other_name_than_tmdb(hk_by_qid, data_dict):
                    logging.info(f"Wikidata 中文維基頁名 '{hk_by_qid}' 同 TMDB 中文名 {_plain_chinese_names(data_dict)} 唔係同一個名"
                                 f"（可能係法號 / 藝名），唔用。")
                elif hk_by_qid:
                    data_dict['wiki_zh_hk_name'] = hk_by_qid
            if data_dict.get("wiki_zh_hk_name"):
                logging.info(f"已由 Wikidata ID 取得香港譯名 '{data_dict['wiki_zh_hk_name']}'，跳過名稱搜尋。")
            elif has_chinese_name and not original_is_japanese:
                logging.info(f"人名已有中文版本（original='{original_name}', zh-TW='{zh_tw_name}', zh-CN='{zh_cn_name}'），且非日文，跳過維基百科搜索。")
            elif not utils.wiki_name_replaces_existing() and is_chinese(data_dict.get("zh_hk_name") or ""):
                # 選項 = n 而 TMDB 已有中文香港人名 → 搵到都唔會用，唔使搜
                logging.info(f"TMDB 已有香港人名 '{data_dict.get('zh_hk_name')}' 且選項為保留，跳過後備維基百科搜索。")
            elif (data_dict.get("wikidata_id") and item_id_for_cache and not is_jk
                    and _wikidata_cache_lookup(item_id_for_cache) == (None, False)):
                # 【2026-11】Wikidata 已確認冇中文維基頁（非日韓人）→ 名稱搜尋幾乎一定都搵唔到，跳過，避免阻塞維基隊列
                logging.info(f"Wikidata ({data_dict['wikidata_id']}) 已確認冇中文維基頁，且非日韓人名，跳過後備維基百科搜索。")
            else:
                # 只有當尚無中文名，或原名為日文時，才執行維基百科搜索
                # 先檢查 DB 快取 — 如果之前 wiki 查詢已為 NOT_FOUND/FAILED，直接跳過所有網路請求
                _wiki_skip = False
                if item_id_for_cache:
                    from others.wiki_search import _init_db as _wiki_init_db, _query_wiki_cache as _wiki_query_cache
                    _wiki_init_db()
                    _wiki_c = _wiki_query_cache(item_id_for_cache, 'person')
                    if _wiki_c and _wiki_c['status'] in ['NOT_FOUND', 'FAILED']:
                        logging.info(f"DB 快取顯示人員 ID '{item_id_for_cache}' wiki 查詢先前無結果，跳過維基百科搜索。")
                        _wiki_skip = True
                if not _wiki_skip:
                    if not data_dict.get("wiki_zh_hk_name"):
                        chinese_title_to_check = data_dict.get("zh_tw_name") or data_dict.get("zh_cn_name") or data_dict.get("zh_hk_name")
                        if chinese_title_to_check:
                            logging.info(f"檢測到中文名稱 '{chinese_title_to_check}'，嘗試從中文維基 HTML (zh-hk variant) 獲取香港譯名...")
                            hk_name_from_html = None
                            try:
                                encoded_title = quote(chinese_title_to_check.replace(" ", "_"))
                                url = f"https://zh.wikipedia.org/wiki/{encoded_title}?variant=zh-hk"

                                # 【修改】對維基百科的請求也加入重試 (雖然是次要請求)
                                async with semaphore:
                                    if utils._PROGRAM_INTERRUPTED: return {}
                                    # 簡單重試循環
                                    for _ in range(3):
                                        try:
                                            response = await client.get(url, headers={"User-Agent": WIKI_USER_AGENT}, timeout=15)
                                            if response.status_code == 200:
                                                soup = BeautifulSoup(response.text, 'html.parser')
                                                break
                                        except Exception:
                                            await asyncio.sleep(1)
                                    else:
                                        response = None

                                if response and response.status_code == 200:
                                    h1_tag = soup.find("h1", {"id": "firstHeading"})
                                    h1_title = _clean_wiki_name(h1_tag.text) if h1_tag and h1_tag.text else None

                                    if h1_title:
                                        hk_name_from_html = h1_title
                                    else:
                                        logging.warning(f"無法從頁面 '{chinese_title_to_check}' 的 HTML 中獲取 H1 標題。")
                                else:
                                    logging.warning(f"Wiki 請求失敗")

                            except httpx.RequestError as e:
                                logging.warning(f"抓取中文維基 HTML 變體時出錯: {e}")
                            except Exception as e:
                                logging.error(f"解析中文維基 HTML 變體時出錯: {e}")

                            if hk_name_from_html:
                                if _is_likely_org_name(hk_name_from_html):
                                    logging.warning(f"Wiki HTML 返回的名稱 '{hk_name_from_html}' 疑似公司/組織名，已棄用。")
                                else:
                                    data_dict['wiki_zh_hk_name'] = hk_name_from_html
                                    logging.info(f"成功從 HTML (zh-hk variant) 獲取香港譯名: '{hk_name_from_html}'")
                            else:
                                logging.info("從 HTML 變體頁面未能找到香港譯名。")

                    if not data_dict.get("wiki_zh_hk_name"):
                        if utils._PROGRAM_INTERRUPTED: return {}
                        search_targets = data_dict.get('all_name_translations', {}).copy()
                        en_name_to_use = data_dict.get('en_name') or data_dict.get('original_title_or_name')
                        if en_name_to_use:
                            search_targets['en'] = en_name_to_use

                        # 只搜索對查找中文名有意義的語言，避免搜索法文/德文等數十個無用語言
                        # 導致大量 Wikipedia 429 請求，造成程式停頓數十分鐘。
                        _USEFUL_PERSON_WIKI_LANGS = {'en', 'ja', 'ko'}
                        search_targets_by_lang = {
                            lang.split('-')[0]: name for lang, name in search_targets.items()
                            if name and not lang.startswith('zh') and lang.split('-')[0] in _USEFUL_PERSON_WIKI_LANGS
                        }
                        for lang, name in search_targets.items():
                            if name and lang.startswith('zh'):
                                search_targets_by_lang[lang] = name

                        if not search_targets_by_lang:
                            logging.info(f"{context_for_logging}: 沒有可用於維基百科後備搜尋的目標名稱。")
                        else:
                            logging.info(f"{context_for_logging}: 執行通用的後備維基百科搜尋...")

                            wiki_hk_name = search_wiki_for_hk_name(
                                item_id=item_id_for_cache,
                                original_name=data_dict.get("original_title_or_name"),
                                media_type='person',
                                year=data_dict.get("year"),
                                search_targets_by_lang=search_targets_by_lang
                            )

                            if wiki_hk_name:
                                if _is_likely_org_name(wiki_hk_name):
                                    logging.warning(f"後備 Wiki 搜尋返回的名稱 '{wiki_hk_name}' 疑似公司/組織名，已棄用。")
                                else:
                                    logging.info(f"後備維基百科搜尋成功找到香港譯名: '{wiki_hk_name}'")
                                    data_dict['wiki_zh_hk_name'] = wiki_hk_name
                            else:
                                logging.info(f"後備維基百科搜尋失敗，未能找到香港譯名。")

            # 【2026-11-13】Bangumi：日韓人名搵唔到任何中文名、TMDB 又冇漢字名 → 用日文名喺 Bangumi 搵中文名
            # （當 zh-CN 候選，排喺 zh-CN / 中文別名後面；搵到就唔使做 B）
            if is_jk and not data_dict.get("wiki_zh_hk_name") and not _has_chinese_person_name(data_dict):
                ja_names = [(data_dict.get("all_name_translations") or {}).get("ja-JP"), original_name]
                ja_names += [a for a in (data_dict.get("also_known_as") or []) if isinstance(a, str)]
                bgm_name = await asyncio.to_thread(bangumi.lookup_zh_name, item_id_for_cache, [n for n in ja_names if n],
                                                   data_dict.get("birthday"), original_name)
                if utils._PROGRAM_INTERRUPTED: return {}
                if bgm_name:
                    data_dict['bangumi_zh_name'] = bgm_name

            # 【2026-11】B：日韓人名搵唔到任何中文名、TMDB 又冇漢字名 → 用 Wikidata 日文漢字名（低優先，見 _person_ja_value）
            if (is_jk and data_dict.get("wikidata_id") and not data_dict.get("wiki_zh_hk_name")
                    and not data_dict.get("bangumi_zh_name") and not _has_chinese_person_name(data_dict)):
                kanji = await asyncio.to_thread(lookup_ja_kanji_name_by_wikidata_queued, item_id_for_cache,
                                                data_dict["wikidata_id"], original_name)
                if utils._PROGRAM_INTERRUPTED: return {}
                if kanji:
                    data_dict['wiki_ja_kanji_name'] = kanji
    
    except (KeyboardInterrupt, SystemExit):
        logging.warning(f"get_api_data_for_item 檢測到中斷信號。")
        return {}
    except httpx.RequestError as e:
        logging.error(f"{context_for_logging}: 處理 {endpoint_suffix} 時發生網絡錯誤: {e}")
    except Exception as e:
        logging.error(f"{context_for_logging}: 處理 {endpoint_suffix} 時發生未知錯誤: {e}")
        logging.error(traceback.format_exc())

    log_output = "\n".join([f"  {k}: '{v}'" for k, v in api_raw_data_log.items()])
    logging.info(f"{context_for_logging} API 返回原始數據:\n{log_output}")

    return data_dict

def create_retry_session():
    """【新增】創建一個配置了無限重試機制的 requests Session 物件，以處理網路中斷。"""
    session = requests.Session()
    retry_strategy = Retry(
        total=RETRY_COUNT, # 【使用者要求】設定一個較大的重試次數以模擬無限重試
        status_forcelist=[429, 500, 502, 503, 504], # 遇到這些伺服器錯誤時重試
        allowed_methods=["HEAD", "GET", "OPTIONS"],
        backoff_factor=RETRY_BACKOFF # 重試間的等待時間因子
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

def translate_to_hk(text):
    """
    【移除】此函數已被移除。
    所有轉換邏輯已統一由 utils.normalize_to_hk() 處理，
    該函數會使用由 start.py 全局修補過的、正確的 OpenCC 實例。
    """
    # 函數內容已移除
    pass

def convert_runtime_to_minutes(runtime):
    """將運行時間字符串轉換為分鐘並返回整數"""
    if not runtime:
        return None
    if isinstance(runtime, str):
        if 'h' in runtime and 'm' in runtime:
            hours, minutes = map(int, re.findall(r'\d+', runtime))
            return hours * 60 + minutes
        elif 'h' in runtime:
            hours = int(re.findall(r'\d+', runtime)[0])
            return hours * 60
        elif 'm' in runtime:
            minutes = int(re.findall(r'\d+', runtime)[0])
            return minutes
        return int(float(runtime)) if runtime.replace('.', '', 1).isdigit() else None
    return int(runtime) if isinstance(runtime, (int, float)) else None

async def get_available_translations(client, semaphore, api_key, tmdb_id, media_type):
    """【修改】查詢 TMDB API 獲取可用翻譯語言，包含 zh-HK (非同步版本)"""
    headers = {"Authorization": f"Bearer {api_key}"}
    base_url = "https://api.themoviedb.org/3"
    endpoint = f"{base_url}/{media_type}/{tmdb_id}/translations"

    try:
        async with semaphore:
            response = await client.get(endpoint, headers=headers, timeout=20)
        if response.status_code != 200:
            return []
        data = response.json()
        translations = data.get("translations", [])
        available_languages = []
        for t in translations:
            lang = t.get("iso_639_1", "").lower().strip()
            country = t.get("iso_3166_1", "").upper().strip()
            if lang and country:
                available_languages.append((lang, country))
        return available_languages
    except httpx.RequestError as e:
        logging.error(f"獲取可用翻譯時出錯 for {media_type}/{tmdb_id}: {e}")
        return []

async def fetch_with_fallback_api(client, semaphore, api_key, endpoint, tmdb_id, media_type, preferred_language="zh-TW", season_number=None, episode_number=None, fallback_language="zh-CN", overwrite_titles="ask", overwrite_descriptions="ask", tv_show_title=None):
    """【修改】獲取資料並提供後備選項 (非同步版本)"""
    headers = {"Authorization": f"Bearer {api_key}"}
    base_url = "https://api.themoviedb.org/3"
    
    # 確定上下文名稱
    context = "Unknown"
    if media_type == "tv":
        if season_number is not None and episode_number is not None:
            context = f"Episode {season_number}-{episode_number}"
        elif season_number is not None:
            context = f"Season {season_number}"
        else:
            context = "TV Show"
    elif media_type == "movie":
        context = "Movie"
    url = f"{base_url}/{endpoint}"

    # 請求 zh-HK, zh-TW, zh-CN, en 數據
    # 使用 asyncio.gather 來並行發送請求
    urls_to_fetch = {
        'zh-HK': f"{url}?language=zh-HK",
        'zh-TW': f"{url}?language=zh-TW",
        'zh-CN': f"{url}?language=zh-CN",
        'en': f"{url}?language=en"
    }

    async def fetch_url(lang, fetch_url):
        try:
            async with semaphore:
                response = await client.get(fetch_url, headers=headers, timeout=20)
            if response.status_code == 200:
                return lang, response.json()
            else:
                logging.error(f"{context} 獲取 {lang} 資料錯誤: status {response.status_code}")
                return lang, None
        except httpx.RequestError as e:
            logging.error(f"{context} 獲取 {lang} 資料錯誤: {e}")
            return lang, None

    tasks = [fetch_url(lang, u) for lang, u in urls_to_fetch.items()]
    results = await asyncio.gather(*tasks)
    
    api_data = {lang: data for lang, data in results if data}

    zh_hk_data = api_data.get('zh-HK', {})
    zh_tw_data = api_data.get('zh-TW', {})
    zh_cn_data = api_data.get('zh-CN', {})
    en_data = api_data.get('en', {})

    # 解析 zh-HK 數據
    zh_hk_title = zh_hk_data.get("name") or zh_hk_data.get("title") or ""
    zh_hk_description = zh_hk_data.get("overview") or ""
    zh_hk_runtime = None
    if not (season_number is not None and episode_number is None):
        zh_hk_runtime = convert_runtime_to_minutes(zh_hk_data.get("runtime"))
    zh_hk_has_data = bool(zh_hk_title.strip() or zh_hk_description.strip())
    logging.debug(f"{context} zh-HK 數據: 標題={zh_hk_title}, 描述={zh_hk_description[:50]}..., 時長={zh_hk_runtime}")

    # 解析 zh-TW 數據
    zh_tw_title = zh_tw_data.get("name") or zh_tw_data.get("title") or ""
    zh_tw_description = zh_tw_data.get("overview") or ""
    zh_tw_runtime = None
    if not (season_number is not None and episode_number is None):
        zh_tw_runtime = convert_runtime_to_minutes(zh_tw_data.get("runtime"))
    logging.debug(f"{context} zh-TW 數據: 標題={zh_tw_title}, 描述={zh_tw_description[:50]}...")

    # 解析 zh-CN 數據
    zh_cn_title = zh_cn_data.get("name") or zh_cn_data.get("title") or ""
    zh_cn_description = zh_cn_data.get("overview") or ""
    zh_cn_runtime = None
    if not (season_number is not None and episode_number is None):
        zh_cn_runtime = convert_runtime_to_minutes(zh_cn_data.get("runtime"))
    logging.debug(f"{context} zh-CN 數據: 標題={zh_cn_title}, 描述={zh_cn_description[:50]}...")

    # 解析 en 數據
    en_title = en_data.get("name") or en_data.get("title") or ""
    en_description = en_data.get("overview") or ""
    logging.debug(f"{context} en 數據: 標題={en_title}, 描述={en_description[:50]}...")

    # --- 後續處理邏輯保持不變 ---
    # 標題選擇邏輯
    final_title = ""
    if is_generic_title(zh_hk_title):
        zh_hk_title = ""
    if is_chinese(zh_hk_title):
        final_title = normalize_to_hk(zh_hk_title) if is_simplified_chinese(zh_hk_title) else zh_hk_title
    elif is_chinese(zh_tw_title) and not is_generic_title(zh_tw_title):
        final_title = normalize_to_hk(zh_tw_title)
    elif is_chinese(zh_cn_title) and not is_generic_title(zh_cn_title):
        final_title = normalize_to_hk(zh_cn_title)
    elif is_english(zh_hk_title):
        final_title = zh_hk_title
    elif is_english(zh_tw_title) and not is_generic_title(zh_tw_title):
        final_title = zh_tw_title
    elif is_english(zh_cn_title) and not is_generic_title(zh_cn_title):
        final_title = zh_cn_title
    else:
        final_title = zh_hk_title or zh_tw_title or zh_cn_title

    # 描述選擇邏輯
    final_description = ""
    if zh_hk_description in NO_DESCRIPTION_TEXT:
        zh_hk_description = ""
    if is_chinese(zh_hk_description) and not is_simplified_chinese(zh_hk_description):
        pass
    elif is_chinese(zh_tw_description) and zh_tw_description not in NO_DESCRIPTION_TEXT:
        final_description = normalize_to_hk(zh_tw_description)
    elif is_chinese(zh_cn_description) and zh_cn_description not in NO_DESCRIPTION_TEXT:
        final_description = normalize_to_hk(zh_cn_description)
    elif is_english(zh_hk_description):
        final_description = zh_hk_description
    elif is_english(zh_tw_description) and zh_tw_description not in NO_DESCRIPTION_TEXT:
        final_description = zh_tw_description
    elif is_english(zh_cn_description) and zh_cn_description not in NO_DESCRIPTION_TEXT:
        final_description = zh_cn_description
    else:
        final_description = zh_tw_description or zh_cn_description or zh_hk_description

    # 時長選擇邏輯
    final_runtime = zh_hk_runtime or zh_tw_runtime or zh_cn_runtime

    # 應用 tw_hk_name.xlsx 替換規則
    if final_title or final_description:
        try:
            replaced_title, replaced_description = apply_tw_hk_name_replacements(
                title=final_title,
                description=final_description,
                return_strings=True
            )
            final_title = replaced_title or final_title
            final_description = replaced_description or final_description
            logging.info(f"{context} 應用 tw_hk_name 替換: 標題={final_title}, 描述={final_description[:50] if final_description else ''}...")
        except Exception as e:
            logging.error(f"{context} 應用 tw_hk_name 替換失敗: {e}")

    # 與 zh-HK 比較
    should_write_to_excel, is_same_as_zh_hk, final_title, final_description, final_runtime, _, _ = compare_with_zh_hk(
        new_title=final_title, new_description=final_description, new_runtime=final_runtime,
        new_name=None, new_bio=None,
        zh_hk_title=zh_hk_title, zh_hk_description=zh_hk_description, zh_hk_runtime=zh_hk_runtime,
        zh_hk_name=None, zh_hk_bio=None,
        media_type=media_type, overwrite_titles=overwrite_titles, overwrite_descriptions=overwrite_descriptions,
        tv_show_title=tv_show_title, context=context
    )

    has_valid_translation = bool(final_title.strip() or final_description.strip())
    logging.info(f"{context} 最終狀態: 標題={final_title}, 描述={final_description[:50]}..., 時長={final_runtime}, "
                 f"有效翻譯={has_valid_translation}, 與 zh-HK 相同={is_same_as_zh_hk}, 需要寫入={should_write_to_excel}")

    # 返回所有必要的值，包括原始標題
    return (
        final_title, final_description, zh_hk_data, has_valid_translation, is_same_as_zh_hk, zh_hk_has_data, should_write_to_excel,
        zh_hk_title, zh_tw_title, zh_cn_title, en_title
    )

async def extract_tv_show_details(client, semaphore, tmdb_id, api_key, url, overwrite_titles_str="ask", overwrite_descriptions_str="ask", bulk_mode=False, row_index=None, file_path='data/video_detail.xlsx'):
    """
    提取 TV Show 主體資訊。
    # 觸發條件：採用直觀比較邏輯，確保標題/簡介內容與 TMDB 原始狀態同步。
    """
    context_main = f"TV Show {tmdb_id}"
    logging.info(f"開始提取 {context_main} 的主體和人員資訊...")
    base_url_tmdb = "https://api.themoviedb.org/3"
    endpoint_suffix = f"/tv/{tmdb_id}"
    
    api_data = await get_api_data_for_item(client, semaphore, api_key, base_url_tmdb, endpoint_suffix, item_is_person=False, context_for_logging=context_main)

    wiki_hk_name = ""
    if api_data.get("original_language") in ['en', 'ja', 'ko']:
        search_targets_for_wiki = {api_data.get("original_language"): api_data.get("original_title_or_name")}
        if api_data.get("en_title"): search_targets_for_wiki['en'] = api_data.get("en_title")
        wiki_hk_name = search_wiki_for_hk_name(item_id=tmdb_id, original_name=api_data.get("original_title_or_name"), media_type='tv', year=api_data.get("year"), search_targets_by_lang=search_targets_for_wiki, original_language=api_data.get("original_language"), tmdb_chinese_names={'zh-TW': api_data.get('zh_tw_title'), 'zh-CN': api_data.get('zh_cn_title')})
    
    desc_candidate, desc_should_write, desc_is_identical, desc_was_imperfect = process_translatable_item(
        item_type="description", media_category="video",
        zh_hk_value=api_data.get("zh_hk_desc"), zh_tw_value=api_data.get("zh_tw_desc"), zh_cn_value=api_data.get("zh_cn_desc"),
        context_for_logging=context_main
    )
    
    # 【直觀比較 - Description】直接與 TMDB 原始內容比對
    current_tmdb_desc = api_data.get("zh_hk_desc") or ""
    _current_tmdb_desc_norm = re.sub(r'\n[ \t]+', '\n', current_tmdb_desc.replace('\r\n', '\n').replace('\r', '\n'))
    if desc_candidate and desc_candidate != _current_tmdb_desc_norm:
        desc_should_write = True
        logging.info(f"{context_main} (Description): 內容與 TMDB 原始狀態不符，強制寫入。")

    title_candidate, title_should_write, title_is_identical, title_was_imperfect = process_translatable_item(
        item_type="title", media_category="video",
        zh_hk_value=api_data.get("zh_hk_title"), zh_tw_value=api_data.get("zh_tw_title"), zh_cn_value=api_data.get("zh_cn_title"),
        en_value=wiki_hk_name or api_data.get("en_title"), ja_value=api_data.get('all_title_translations', {}).get('ja-JP'),
        ko_value=api_data.get('all_title_translations', {}).get('ko-KR'), original_value=api_data.get("original_title_or_name"),
        sibling_desc_is_empty_for_eng_title=(not desc_candidate), context_for_logging=context_main
    )

    # 【直觀比較 - Title】清理後與 TMDB 原始標題比對（只在 process_translatable_item 允許寫入時才生效）
    _cleaned_candidate = clean_generic_title(title_candidate) if title_candidate else ""
    current_tmdb_title = api_data.get("zh_hk_title") or ""
    if title_should_write and _cleaned_candidate and _cleaned_candidate != current_tmdb_title:
        logging.info(f"{context_main} (Title): 標題內容與 TMDB 原始不一致 (含前綴差異)，強制寫入。")
    excel_title_to_save = _cleaned_candidate if title_should_write else current_tmdb_title
    excel_desc_to_save = desc_candidate if desc_should_write else current_tmdb_desc

    if title_should_write or desc_should_write:
        title_differs = (excel_title_to_save != current_tmdb_title)
        desc_differs = (excel_desc_to_save != current_tmdb_desc) if excel_desc_to_save else False
        perfect_title_flag = 'n' if title_was_imperfect or title_differs or not is_valid_zh_hk_title(excel_title_to_save, field_type='title') else 'y'
        perfect_desc_flag = 'n' if desc_was_imperfect or desc_differs or not is_valid_zh_hk_title(excel_desc_to_save, field_type='description') else 'y'
        save_series_info_to_excel(excel_title_to_save, excel_desc_to_save, perfect_title_flag, perfect_desc_flag, file_name=file_path)

    work_title_context_for_people = excel_title_to_save if excel_title_to_save.strip() else (api_data.get("original_title_or_name") or context_main)
    work_desc_context_for_people = excel_desc_to_save if excel_desc_to_save.strip() else api_data.get("en_desc")
    logging.info(f"{context_main}: 開始提取演職員資料...")
    people_are_all_identical_and_no_write_needed = await extract_person_details_for_credits(client, semaphore, api_key, tmdb_id, "tv", overwrite_titles=overwrite_titles_str, overwrite_descriptions=overwrite_descriptions_str, url=url, work_title_for_mapping_context=work_title_context_for_people, work_description_for_mapping_context=work_desc_context_for_people, file_path=file_path)
    logging.info(f"{context_main}: 演職員資料提取完成，是否全部相同: {people_are_all_identical_and_no_write_needed}")
    show_title_for_nested_context = excel_title_to_save if excel_title_to_save.strip() else (api_data.get("original_title_or_name") or context_main)
    show_desc_for_nested_context = excel_desc_to_save if excel_desc_to_save.strip() else (api_data.get("en_desc") or "")
    logging.info(f"{context_main}: 開始提取季度資料...")
    seasons_to_write, processed_s_numbers, seasons_are_identical = await extract_season_details(client, semaphore, tmdb_id, api_key, "zh-TW", "zh-CN", overwrite_titles_str, overwrite_descriptions_str, tv_show_title_context=show_title_for_nested_context, tv_show_description_context=show_desc_for_nested_context, url=url, file_path=file_path)
    logging.info(f"{context_main}: 季度資料提取完成，需寫入季數: {len(seasons_to_write)}，已處理季號: {processed_s_numbers}")
    if seasons_to_write: save_season_info_to_excel(seasons_to_write, file_name=file_path)
    logging.info(f"{context_main}: 開始提取 {len(processed_s_numbers)} 季的劇集資料...")
    episode_tasks = [extract_episode_details(client, semaphore, tmdb_id, s_num, api_key, "zh-TW", "zh-CN", overwrite_titles_str, overwrite_descriptions_str, tv_show_title_context=show_title_for_nested_context, tv_show_description_context=show_desc_for_nested_context, url=url, file_path=file_path) for s_num in processed_s_numbers]
    episode_results = await asyncio.gather(*episode_tasks)
    all_episodes_to_write = [ep for episodes_for_season, _ in episode_results for ep in episodes_for_season]
    logging.info(f"{context_main}: 劇集資料提取完成，需寫入集數: {len(all_episodes_to_write)}")
    if all_episodes_to_write: save_episode_info_to_excel(all_episodes_to_write, file_name=file_path)
    overall_item_is_identical_and_no_write_needed = not (title_should_write or desc_should_write or not people_are_all_identical_and_no_write_needed or seasons_to_write or all_episodes_to_write)
    no_chinese_content_body_skipped = not (title_should_write or desc_should_write or seasons_to_write or all_episodes_to_write)
    contextual_title_for_return = excel_title_to_save if excel_title_to_save.strip() else (api_data.get("zh_hk_title") or api_data.get("original_title_or_name") or context_main)
    return (contextual_title_for_return, excel_desc_to_save, "zh-TW", "zh-CN", overall_item_is_identical_and_no_write_needed, excel_title_to_save, no_chinese_content_body_skipped)

async def extract_season_details(client, semaphore, tmdb_id, api_key, preferred_language, fallback_language,
                                 overwrite_titles_str="ask", overwrite_descriptions_str="ask",
                                 tv_show_title_context=None, tv_show_description_context=None, url=None, file_path='data/video_detail.xlsx'):
    """
    提取季詳細資訊。
    # 觸發條件：採用直觀比較邏輯，確保季度標題不含「第 X 季」等冗餘前綴。
    """
    base_url_tmdb = "https://api.themoviedb.org/3"
    seasons_to_write_to_excel = []
    
    show_data = None
    try:
        if utils._PROGRAM_INTERRUPTED: return [], [], True
        while True:
            if utils._PROGRAM_INTERRUPTED: return [], [], True
            try:
                async with semaphore:
                    tv_show_details_url_en = f"{base_url_tmdb}/tv/{tmdb_id}?language=en"
                    response_en = await client.get(tv_show_details_url_en, headers={"Authorization": f"Bearer {api_key}"}, timeout=20)
                if response_en.status_code == 200:
                    show_data = response_en.json()
                    break
                elif response_en.status_code == 404: return [], [], True
                elif response_en.status_code == 429: await asyncio.sleep(int(response_en.headers.get("Retry-After", 5)) + 1)
                else: await asyncio.sleep(2)
            except httpx.RequestError: await asyncio.sleep(2)
    except Exception as e:
        logging.error(f"TV Show {tmdb_id} 獲取季列表錯誤: {e}")
        return [], [], True

    if not show_data or "seasons" not in show_data: return [], [], True
    season_list_from_api = show_data.get("seasons", [])
    tasks_with_context = []
    for season_api_data_item in season_list_from_api:
        if utils._PROGRAM_INTERRUPTED: break
        season_number = season_api_data_item.get("season_number")
        if season_number is None: continue
        tasks_with_context.append((season_number, f"TV Show {tmdb_id} Season {season_number}", get_api_data_for_item(client, semaphore, api_key, base_url_tmdb, f"/tv/{tmdb_id}/season/{season_number}", item_is_person=False, context_for_logging=f"TV Show {tmdb_id} Season {season_number}")))

    if not tasks_with_context: return [], [], True
    task_results = await asyncio.gather(*(t for _, _, t in tasks_with_context), return_exceptions=True)
    season_numbers_processed = [sn for sn, _, _ in tasks_with_context]

    for (season_number, context_prefix, _), api_season_data in zip(tasks_with_context, task_results):
        if utils._PROGRAM_INTERRUPTED or isinstance(api_season_data, Exception) or not api_season_data: continue
        
        desc_candidate, desc_should_write, _, desc_was_imperfect = process_translatable_item(item_type="description", media_category="video", zh_hk_value=api_season_data.get("zh_hk_desc"), zh_tw_value=api_season_data.get("zh_tw_desc"), zh_cn_value=api_season_data.get("zh_cn_desc"), context_for_logging=context_prefix)
        # 【直觀比較 - Description】
        current_tmdb_desc = api_season_data.get("zh_hk_desc") or ""
        _current_tmdb_desc_norm = re.sub(r'\n[ \t]+', '\n', current_tmdb_desc.replace('\r\n', '\n').replace('\r', '\n'))
        if desc_candidate and desc_candidate != _current_tmdb_desc_norm:
            desc_should_write = True
            logging.info(f"{context_prefix} (Description): 簡介內容與 TMDB 原始狀態不一致，觸發強制更新。")

        title_candidate, title_should_write, _, title_was_imperfect = process_translatable_item(item_type="title", media_category="video", zh_hk_value=api_season_data.get("zh_hk_title"), zh_tw_value=api_season_data.get("zh_tw_title"), zh_cn_value=api_season_data.get("zh_cn_title"), en_value=api_season_data.get("en_title"), sibling_desc_is_empty_for_eng_title=(not desc_candidate), context_for_logging=context_prefix)
        # 【直觀比較 - Title】（只在 process_translatable_item 允許寫入時才生效）
        _cleaned_candidate = clean_generic_title(title_candidate) if title_candidate else ""
        current_tmdb_title = api_season_data.get("zh_hk_title") or ""
        if title_should_write and _cleaned_candidate and _cleaned_candidate != current_tmdb_title:
            logging.info(f"{context_prefix} (Title): 標題內容與 TMDB 原始狀態不一致，觸發強制更新。")
        excel_title_to_save = _cleaned_candidate if title_should_write else current_tmdb_title

        if title_should_write or desc_should_write:
            excel_desc_to_save = desc_candidate if desc_should_write else current_tmdb_desc
            title_differs = (excel_title_to_save != current_tmdb_title)
            desc_differs = (excel_desc_to_save != current_tmdb_desc) if excel_desc_to_save else False
            perfect_name_flag = 'n' if title_was_imperfect or (excel_title_to_save != title_candidate) or title_differs or not is_valid_zh_hk_title(excel_title_to_save, field_type='title') else 'y'
            perfect_desc_flag = 'n' if desc_was_imperfect or desc_differs or not is_valid_zh_hk_title(excel_desc_to_save, field_type='description') else 'y'
            seasons_to_write_to_excel.append((excel_title_to_save, season_number, excel_desc_to_save, perfect_name_flag, perfect_desc_flag))

    return seasons_to_write_to_excel, season_numbers_processed, not bool(seasons_to_write_to_excel)

async def extract_episode_details(client, semaphore, tmdb_id, season_number, api_key, preferred_language, fallback_language,
                                  overwrite_titles_str="ask", overwrite_descriptions_str="ask",
                                  tv_show_title_context=None, tv_show_description_context=None, url=None, file_path='data/video_detail.xlsx'):
    """
    提取劇集詳細資訊。
    # 觸發：Episodes sheet有資料即上載；採用直觀比較邏輯。
    # 原要求：處理多來源extractor（netflix, disneyplus, appletv, primevideo等）；觸發：tv show/seasons/episodes/movie/person sheet任一有資料即上載。
    # 新增要求：修復 typo，確保 title 和 description 獨立處理，避免將 tw/hk 的 description 誤抄寫至 title。
    # 邏輯確認：如果 hk 無資料, 就找 tw, tw 無資料, 就找 cn, 如果 3 個都無資料, 就唔更新 keep 原樣。
    """
    base_url_tmdb = "https://api.themoviedb.org/3"
    episodes_to_write_to_excel = []
    
    season_data_from_api_list = None
    try:
        if utils._PROGRAM_INTERRUPTED: return [], True
        while True:
            if utils._PROGRAM_INTERRUPTED: return [], True
            try:
                async with semaphore:
                    season_details_url_en = f"{base_url_tmdb}/tv/{tmdb_id}/season/{season_number}?language=en"
                    response_en = await client.get(season_details_url_en, headers={"Authorization": f"Bearer {api_key}"}, timeout=20)
                if response_en.status_code == 200:
                    season_data_from_api_list = response_en.json()
                    break
                elif response_en.status_code == 404: return [], True
                elif response_en.status_code == 429: await asyncio.sleep(int(response_en.headers.get("Retry-After", 5)) + 1)
                else: await asyncio.sleep(2)
            except httpx.RequestError: await asyncio.sleep(2)
    except Exception as e:
        logging.error(f"TV Show {tmdb_id} S{season_number} 獲取劇集列表錯誤: {e}")
        return [], True

    if not season_data_from_api_list or "episodes" not in season_data_from_api_list: return [], True
    tasks_with_context = []
    for episode_api_data_item in season_data_from_api_list.get("episodes", []):
        if utils._PROGRAM_INTERRUPTED: break
        episode_number = episode_api_data_item.get("episode_number")
        if episode_number is None: continue
        tasks_with_context.append((episode_number, f"TV Show {tmdb_id} S{season_number}E{episode_number}", get_api_data_for_item(client, semaphore, api_key, base_url_tmdb, f"/tv/{tmdb_id}/season/{season_number}/episode/{episode_number}", item_is_person=False, context_for_logging=f"TV Show {tmdb_id} S{season_number}E{episode_number}")))

    if not tasks_with_context: return [], True
    task_results = await asyncio.gather(*(t for _, _, t in tasks_with_context), return_exceptions=True)

    for (episode_number, context_prefix, _), api_ep_data in zip(tasks_with_context, task_results):
        if utils._PROGRAM_INTERRUPTED or isinstance(api_ep_data, Exception) or not api_ep_data: continue

        desc_candidate, desc_should_write, _, desc_was_imperfect = process_translatable_item(item_type="description", media_category="video", zh_hk_value=api_ep_data.get("zh_hk_desc"), zh_tw_value=api_ep_data.get("zh_tw_desc"), zh_cn_value=api_ep_data.get("zh_cn_desc"), context_for_logging=context_prefix)
        # 【直觀比較 - Description】
        current_tmdb_desc = api_ep_data.get("zh_hk_desc") or ""
        _current_tmdb_desc_norm = re.sub(r'\n[ \t]+', '\n', current_tmdb_desc.replace('\r\n', '\n').replace('\r', '\n'))
        if desc_candidate and desc_candidate != _current_tmdb_desc_norm:
            desc_should_write = True
            logging.info(f"{context_prefix} (Description): 簡介內容與 TMDB 原始狀態不一致，觸發強制更新。")

        # 【強化 zh-TW fallback 至 zh-HK 目標】如果正常流程沒有產生有效的 description 候選值，
        # 但 API 有 zh-TW 內容，則強制使用轉換後的 zh-TW 作為寫入 zh-HK slot 的候選值。
        # 這樣可以確保當 zh-HK 缺失時，zh-TW 內容能被轉換並寫入 Excel 供後續上傳使用。
        zh_tw_desc_raw = api_ep_data.get("zh_tw_desc") or ""
        if not desc_candidate and zh_tw_desc_raw and is_chinese(zh_tw_desc_raw):
            converted = normalize_to_hk(zh_tw_desc_raw, field_type="description")
            converted = add_spaces_between_en_zh_num(converted)
            if converted and not is_generic_title(converted):  # 避免寫入無意義內容
                desc_candidate = converted
                desc_should_write = True
                desc_was_imperfect = False
                logging.info(f"{context_prefix} (Description): 無直接 zh-HK，強制使用轉換後的 zh-TW 內容作為 zh-HK 候選值。")

        # 【修改】已修復此處將 zh_tw_desc 誤傳給 zh_tw_value 的 typo，確保正確獲取 title
        title_candidate, title_should_write, _, title_was_imperfect = process_translatable_item(item_type="title", media_category="video", zh_hk_value=api_ep_data.get("zh_hk_title"), zh_tw_value=api_ep_data.get("zh_tw_title"), zh_cn_value=api_ep_data.get("zh_cn_title"), en_value=api_ep_data.get("en_title"), sibling_desc_is_empty_for_eng_title=(not desc_candidate), context_for_logging=context_prefix)
        # 【直觀比較 - Title】（只在 process_translatable_item 允許寫入時才生效）
        _cleaned_candidate = clean_generic_title(title_candidate) if title_candidate else ""
        current_tmdb_title = api_ep_data.get("zh_hk_title") or ""
        if title_should_write and _cleaned_candidate and _cleaned_candidate != current_tmdb_title:
            logging.info(f"{context_prefix} (Title): 標題內容與 TMDB 原始狀態不一致，觸發強制更新。")
        # 不需更新時保留現有 TMDB 值，避免 row 因其他欄位被加入 Excel 時 uploader 誤以為空值需清除
        excel_title_to_save = _cleaned_candidate if title_should_write else current_tmdb_title

        # 【強化 zh-TW fallback 至 zh-HK 目標 - Title】
        zh_tw_title_raw = api_ep_data.get("zh_tw_title") or ""
        if not title_candidate and zh_tw_title_raw and is_chinese(zh_tw_title_raw):
            converted_title = normalize_to_hk(zh_tw_title_raw, field_type="title")
            converted_title = add_spaces_between_en_zh_num(converted_title)
            if converted_title and not is_generic_title(converted_title):
                title_candidate = converted_title
                title_should_write = True
                title_was_imperfect = False
                _cleaned_candidate = clean_generic_title(title_candidate)
                excel_title_to_save = _cleaned_candidate
                logging.info(f"{context_prefix} (Title): 無直接 zh-HK，強制使用轉換後的 zh-TW 內容作為 zh-HK 候選值。")

        runtime_candidate, runtime_should_write, _, _ = process_translatable_item(item_type="runtime", media_category="video", zh_hk_value=None, zh_tw_value=None, zh_cn_value=None, current_runtime_for_video=api_ep_data.get("zh_hk_runtime"), zh_hk_runtime_for_video=api_ep_data.get("zh_hk_runtime"), context_for_logging=context_prefix)

        if title_should_write or desc_should_write or runtime_should_write:
            excel_desc_to_save = desc_candidate if desc_should_write else current_tmdb_desc
            actual_runtime_for_excel = int(runtime_candidate) if runtime_should_write and runtime_candidate else None
            title_differs = (excel_title_to_save != current_tmdb_title)
            desc_differs = (excel_desc_to_save != current_tmdb_desc) if excel_desc_to_save else False
            perfect_title_flag = 'n' if title_was_imperfect or (excel_title_to_save != title_candidate) or title_differs or not is_valid_zh_hk_title(excel_title_to_save, field_type='title') else 'y'
            perfect_desc_flag = 'n' if desc_was_imperfect or desc_differs or not is_valid_zh_hk_title(excel_desc_to_save, field_type='description') else 'y'
            episodes_to_write_to_excel.append((season_number, episode_number, actual_runtime_for_excel, excel_title_to_save, excel_desc_to_save, perfect_title_flag, perfect_desc_flag, None))

    return episodes_to_write_to_excel, not bool(episodes_to_write_to_excel)

async def extract_movie_details(client, semaphore, tmdb_id, api_key, url, overwrite_titles_str="ask", overwrite_descriptions_str="ask", bulk_mode=False, row_index=None, file_path='data/video_detail.xlsx'):
    """
    提取 Movie 詳細資訊。
    # 觸發條件：採用直觀比較邏輯，確保標題與簡介內容正確。
    """
    context_main = f"Movie {tmdb_id}"
    logging.info(f"開始提取 {context_main} 的詳細資訊...")
    base_url_tmdb = "https://api.themoviedb.org/3"
    endpoint_suffix = f"/movie/{tmdb_id}"
    api_data = await get_api_data_for_item(client, semaphore, api_key, base_url_tmdb, endpoint_suffix, item_is_person=False, context_for_logging=context_main)
    if not (api_data.get("zh_hk_title") or api_data.get("en_title")): return None, None, None, True, True 

    wiki_hk_name = ""
    if api_data.get("original_language") in ['en', 'ja', 'ko']:
        search_targets_for_wiki = {api_data.get("original_language"): api_data.get("original_title_or_name")}
        if api_data.get("en_title"): search_targets_for_wiki['en'] = api_data.get("en_title")
        wiki_hk_name = search_wiki_for_hk_name(item_id=tmdb_id, original_name=api_data.get("original_title_or_name"), media_type='movie', year=api_data.get("year"), search_targets_by_lang=search_targets_for_wiki, original_language=api_data.get("original_language"), tmdb_chinese_names={'zh-TW': api_data.get('zh_tw_title'), 'zh-CN': api_data.get('zh_cn_title')})

    desc_candidate, desc_should_write, _, desc_was_imperfect = process_translatable_item(item_type="description", media_category="video", zh_hk_value=api_data.get("zh_hk_desc"), zh_tw_value=api_data.get("zh_tw_desc"), zh_cn_value=api_data.get("zh_cn_desc"), context_for_logging=context_main)
    # 【直觀比較 - Description】
    current_tmdb_desc = api_data.get("zh_hk_desc") or ""
    _current_tmdb_desc_norm = re.sub(r'\n[ \t]+', '\n', current_tmdb_desc.replace('\r\n', '\n').replace('\r', '\n'))
    if desc_candidate and desc_candidate != _current_tmdb_desc_norm:
        desc_should_write = True
        logging.info(f"{context_main} (Description): 簡介不一致，強制寫入。")

    title_candidate, title_should_write, _, title_was_imperfect = process_translatable_item(item_type="title", media_category="video", zh_hk_value=api_data.get("zh_hk_title"), zh_tw_value=api_data.get("zh_tw_title"), zh_cn_value=api_data.get("zh_cn_title"), en_value=wiki_hk_name or api_data.get("en_title"), original_value=api_data.get("original_title_or_name"), sibling_desc_is_empty_for_eng_title=(not desc_candidate), context_for_logging=context_main)
    # 【直觀比較 - Title】（只在 process_translatable_item 允許寫入時才生效）
    current_tmdb_title = api_data.get("zh_hk_title") or ""
    if title_should_write and title_candidate and title_candidate != current_tmdb_title:
        logging.info(f"{context_main} (Title): 標題不一致，強制寫入。")

    runtime_candidate, runtime_should_write, _, _ = process_translatable_item(item_type="runtime", media_category="video", zh_hk_value=None, zh_tw_value=None, zh_cn_value=None, current_runtime_for_video=api_data.get("zh_hk_runtime"), zh_hk_runtime_for_video=api_data.get("zh_hk_runtime"), context_for_logging=context_main)

    excel_title_to_save = title_candidate if title_should_write else current_tmdb_title
    excel_desc_to_save = desc_candidate if desc_should_write else current_tmdb_desc
    final_runtime_to_save = int(runtime_candidate) if runtime_should_write and runtime_candidate else None

    if title_should_write or desc_should_write or runtime_should_write:
        title_differs = (excel_title_to_save != current_tmdb_title)
        desc_differs = (excel_desc_to_save != current_tmdb_desc) if excel_desc_to_save else False
        perfect_title_flag = 'n' if title_was_imperfect or title_differs or not is_valid_zh_hk_title(excel_title_to_save, field_type='title') else 'y'
        perfect_desc_flag = 'n' if desc_was_imperfect or desc_differs or not is_valid_zh_hk_title(excel_desc_to_save, field_type='description') else 'y'
        save_movie_info_to_excel(excel_title_to_save, excel_desc_to_save, final_runtime_to_save, perfect_title_flag, perfect_desc_flag, file_name=file_path)

    work_title_context_for_people = excel_title_to_save if excel_title_to_save.strip() else (api_data.get("original_title_or_name") or context_main)
    work_desc_context_for_people = excel_desc_to_save if excel_desc_to_save.strip() else api_data.get("en_desc")
    people_are_all_identical_and_no_write_needed = await extract_person_details_for_credits(client, semaphore, api_key, tmdb_id, "movie", overwrite_titles=overwrite_titles_str, overwrite_descriptions=overwrite_descriptions_str, url=url, work_title_for_mapping_context=work_title_context_for_people, work_description_for_mapping_context=work_desc_context_for_people, file_path=file_path)

    return (excel_title_to_save, excel_desc_to_save, final_runtime_to_save, not (title_should_write or desc_should_write or runtime_should_write or not people_are_all_identical_and_no_write_needed), not (title_should_write or desc_should_write or runtime_should_write))

async def extract_person_details(client, semaphore, person_id, api_key, url, driver=None, overwrite_titles_str="ask", overwrite_descriptions_str="ask", bulk_mode=False, row_index=None, file_path='data/video_detail.xlsx'):
    """
    提取單獨人員詳細資訊。
    # 觸發條件：執行人名替換保護後，執行直觀比較以應對「範/范」等轉換錯誤。
    """
    context_main = f"Person {person_id}"
    logging.info(f"開始提取單獨人員 ({context_main}) 的詳細資訊...")
    person_tmdb_url = f"https://www.themoviedb.org/person/{person_id}"
    base_url_tmdb = "https://api.themoviedb.org/3"
    endpoint_suffix = f"/person/{person_id}"

    api_data = await get_api_data_for_item(client, semaphore, api_key, base_url_tmdb, endpoint_suffix, item_is_person=True, context_for_logging=context_main)

    found_chinese_alias = None
    original_name = api_data.get("original_title_or_name")
    if original_name and not is_chinese(original_name):
        for alias in api_data.get("also_known_as", []):
            if is_chinese(alias):
                found_chinese_alias = alias
                break

    name_candidate, name_should_write, _, name_was_imperfect = process_translatable_item(item_type="name", media_category="person", zh_hk_value=api_data.get("zh_hk_name"), zh_tw_value=api_data.get("zh_tw_name"), zh_cn_value=api_data.get("zh_cn_name") or found_chinese_alias or api_data.get("bangumi_zh_name"), en_value=api_data.get('wiki_zh_hk_name'), ja_value=_person_ja_value(api_data), original_value=api_data.get("original_title_or_name"), context_for_logging=context_main)
    bio_candidate, bio_should_write, _, bio_was_imperfect = process_translatable_item(item_type="bio", media_category="person", zh_hk_value=api_data.get("zh_hk_bio"), zh_tw_value=api_data.get("zh_tw_bio"), zh_cn_value=api_data.get("zh_cn_bio"), context_for_logging=context_main,
                                                                                      protect_names=_bio_protect_names(api_data, name_candidate, found_chinese_alias))

    # 【關鍵修正】執行人名替換保護，還原簡介中被錯誤轉換的人名 (例如還原「範」為「范」)
    final_bio_to_compare = bio_candidate
    if name_candidate:
        # 【2026-11】加入 TMDB 現有香港人名：維基譯名取代佢時，簡介入面嘅舊香港名都要換
        old_names_for_bio_replace = [api_data.get("zh_cn_name"), api_data.get("zh_tw_name"), api_data.get("zh_hk_name"), api_data.get("original_title_or_name"), api_data.get("en_name")]
        if found_chinese_alias: old_names_for_bio_replace.append(found_chinese_alias)
        # 產生一個「被錯誤轉換」的版本用於搜尋並還原
        incorrectly_normalized_name = normalize_to_hk(name_candidate, field_type='bio')
        if incorrectly_normalized_name != name_candidate: old_names_for_bio_replace.append(incorrectly_normalized_name)
        # 執行保護替換
        final_bio_to_compare = replace_name_in_bio(bio_candidate, old_names_for_bio_replace, name_candidate)

    # 執行直觀比較
    current_tmdb_name = api_data.get("zh_hk_name") or ""
    current_tmdb_bio = api_data.get("zh_hk_bio") or ""
    
    if name_candidate and name_candidate != current_tmdb_name: name_should_write = True
    if final_bio_to_compare and final_bio_to_compare != current_tmdb_bio: bio_should_write = True

    person_actually_written_to_excel = False
    if name_should_write or bio_should_write:
        name_to_save_final = name_candidate if name_should_write else ""
        bio_to_save_final = final_bio_to_compare if bio_should_write else ""
        name_differs = (name_to_save_final != current_tmdb_name) if name_to_save_final else False
        bio_differs = (bio_to_save_final != current_tmdb_bio) if bio_to_save_final else False
        perfect_name_flag = 'n' if name_was_imperfect or name_differs or not is_valid_zh_hk_title(name_to_save_final, field_type='name') else 'y'
        perfect_bio_flag = 'n' if bio_was_imperfect or bio_differs or not is_valid_zh_hk_title(bio_to_save_final, field_type='bio') else 'y'
        save_person_info_to_excel([(person_tmdb_url, name_to_save_final, bio_to_save_final, perfect_name_flag, perfect_bio_flag)], file_name=file_path)
        person_actually_written_to_excel = True

    return (name_to_save_final if name_should_write else current_tmdb_name, not person_actually_written_to_excel)

async def extract_person_details_for_credits(client, semaphore, api_key, tmdb_id, media_type, preferred_language="zh-TW", fallback_language="zh-CN",
                                             overwrite_titles="ask", overwrite_descriptions="ask", url=None, 
                                             work_title_for_mapping_context=None, work_description_for_mapping_context=None, file_path='data/video_detail.xlsx'):
    """
    提取作品的演職員詳細資訊。
    # 觸發條件：執行直觀比較與人名保護。
    """
    headers = {"Authorization": f"Bearer {api_key}"}
    base_url = "https://api.themoviedb.org/3"
    credits_url = f"{base_url}/{media_type}/{tmdb_id}/{'aggregate_credits' if media_type == 'tv' else 'credits'}"
    people_to_write_to_excel = []
    
    all_person_items = []
    try:
        if utils._PROGRAM_INTERRUPTED: return True
        while True:
            if utils._PROGRAM_INTERRUPTED: return True
            try:
                async with semaphore:
                    credits_resp = await client.get(credits_url, headers=headers, timeout=20)
                if credits_resp.status_code == 200:
                    credits_data = credits_resp.json()
                    unique_people = {p['id']: p for p in (credits_data.get("cast", []) + credits_data.get("crew", [])) if p.get('id')}
                    all_person_items = list(unique_people.values())
                    break
                elif credits_resp.status_code == 404:
                    return True
                elif credits_resp.status_code == 429:
                    retry_after = int(credits_resp.headers.get("Retry-After", 5))
                    logging.warning(f"Credits {tmdb_id}: API 限流 (429)，等待 {retry_after} 秒後重試...")
                    await asyncio.sleep(retry_after + 1)
                else:
                    await asyncio.sleep(3)
            except httpx.RequestError as e:
                logging.warning(f"Credits {tmdb_id}: 網絡錯誤 ({e})，3 秒後重試...")
                await asyncio.sleep(3)
    except Exception as e:
        logging.error(f"Credits {tmdb_id}: 未知錯誤 ({e})")
        return True

    tasks_with_context = []
    for person_entry in all_person_items:
        if utils._PROGRAM_INTERRUPTED: break
        person_id = person_entry.get("id")
        if not person_id: continue
        tasks_with_context.append((person_id, f"Person {person_id}", get_api_data_for_item(client, semaphore, api_key, base_url, f"/person/{person_id}", item_is_person=True, context_for_logging=f"Person {person_id}")))

    if not tasks_with_context: return True
    task_results = await asyncio.gather(*(t for _, _, t in tasks_with_context), return_exceptions=True)

    for (person_id, context_prefix, _), api_person_data in zip(tasks_with_context, task_results):
        if utils._PROGRAM_INTERRUPTED or isinstance(api_person_data, Exception) or not api_person_data: continue

        person_tmdb_url = f"https://www.themoviedb.org/person/{person_id}"
        found_chinese_alias = None
        original_name = api_person_data.get("original_title_or_name")
        if original_name and not is_chinese(original_name):
            for alias in api_person_data.get("also_known_as", []):
                if is_chinese(alias):
                    found_chinese_alias = alias
                    break
        
        name_candidate, name_should_write, _, name_was_imperfect = process_translatable_item(item_type="name", media_category="person", zh_hk_value=api_person_data.get("zh_hk_name"), zh_tw_value=api_person_data.get("zh_tw_name"), zh_cn_value=api_person_data.get("zh_cn_name") or found_chinese_alias or api_person_data.get("bangumi_zh_name"), en_value=api_person_data.get('wiki_zh_hk_name'), ja_value=_person_ja_value(api_person_data), original_value=api_person_data.get("original_title_or_name"), context_for_logging=context_prefix)
        bio_candidate, bio_should_write, _, bio_was_imperfect = process_translatable_item(item_type="bio", media_category="person", zh_hk_value=api_person_data.get("zh_hk_bio"), zh_tw_value=api_person_data.get("zh_tw_bio"), zh_cn_value=api_person_data.get("zh_cn_bio"), context_for_logging=context_prefix,
                                                                                          protect_names=_bio_protect_names(api_person_data, name_candidate, found_chinese_alias))

        current_tmdb_name = api_person_data.get("zh_hk_name") or ""
        current_tmdb_bio = api_person_data.get("zh_hk_bio") or ""
        
        final_bio_to_compare = bio_candidate
        if name_candidate:
            old_names_for_bio_replace = [api_person_data.get("zh_cn_name"), api_person_data.get("zh_tw_name"), api_person_data.get("zh_hk_name"), api_person_data.get("original_title_or_name"), api_person_data.get("en_name")]
            if found_chinese_alias: old_names_for_bio_replace.append(found_chinese_alias)
            incorrectly_normalized_name = normalize_to_hk(name_candidate, field_type='bio')
            if incorrectly_normalized_name != name_candidate: old_names_for_bio_replace.append(incorrectly_normalized_name)
            final_bio_to_compare = replace_name_in_bio(bio_candidate, old_names_for_bio_replace, name_candidate)

        if name_candidate and name_candidate != current_tmdb_name: name_should_write = True
        if final_bio_to_compare:
            bio_should_write = (final_bio_to_compare != current_tmdb_bio)

        if name_should_write or bio_should_write:
            name_to_save = name_candidate if name_should_write else ""
            bio_to_save = final_bio_to_compare if bio_should_write else ""
            name_differs = (name_to_save != current_tmdb_name) if name_to_save else False
            bio_differs = (bio_to_save != current_tmdb_bio) if bio_to_save else False
            perfect_name_flag = 'n' if name_was_imperfect or name_differs or not is_valid_zh_hk_title(name_to_save, field_type='name') else 'y'
            perfect_bio_flag = 'n' if bio_was_imperfect or bio_differs or not is_valid_zh_hk_title(bio_to_save, field_type='bio') else 'y'
            people_to_write_to_excel.append((person_tmdb_url, name_to_save, bio_to_save, perfect_name_flag, perfect_bio_flag))

    if people_to_write_to_excel:
        save_person_info_to_excel(people_to_write_to_excel, file_name=file_path)
        return False
    return True

def _bio_protect_names(api_data, name_candidate, alias=None):
    """【2026-11】簡介轉換時保護呢個人自己嘅名：各種中文寫法 → 最終人名（name_candidate）。
    例如 zh-CN「于成龙」→「于成龍」，簡介入面嘅名唔會被轉成「於成龍」，名以外嘅「于」照常轉換。"""
    if not name_candidate or not is_chinese(name_candidate):
        return None
    sources = [api_data.get("zh_cn_name"), api_data.get("zh_tw_name"), api_data.get("zh_hk_name"),
               api_data.get("original_title_or_name"), alias, name_candidate]
    names = {s: name_candidate for s in sources if isinstance(s, str) and len(s) >= 2 and is_chinese(s)}
    # 【2026-11-13】連名 / 姓（≥2 字）都保護，例如「克鲁斯」→「告魯斯」：唔使行 OpenCC / AI，
    # 慳 token，亦唔會因為佢自己個名出假被拒用詞。結果同之後 _replace_name_parts 換名一樣。
    # 舊部份係新部份其中一截（例如 李維 → 李維斯）就唔保護，交返 _replace_name_parts 處理。
    s2t = utils.opencc_converters.get('s2t')
    for o, n in _name_part_pairs(sources, name_candidate).items():
        if o not in n and (s2t is None or s2t.convert(o) not in n):
            names.setdefault(o, n)
    return names or None


_NAME_SEP_RE = re.compile(r'[·・•‧]')
_BIO_TITLE_RE = re.compile(r'《[^》]*》|〈[^〉]*〉')


def _name_part_pairs(old_names, new_name):
    """舊名「部份」→ 新名對應部份 {舊: 新}（用「·」分開；數目一樣逐部份對應，唔一樣只對應姓；只取 ≥2 字）。"""
    new_parts = [p for p in _NAME_SEP_RE.split(new_name or '') if p]
    if len(new_parts) < 2:
        return {}
    pairs = {}
    for old in old_names or ():
        old_parts = [p for p in _NAME_SEP_RE.split(old or '') if p]
        if len(old_parts) < 2:
            continue
        cand = zip(old_parts, new_parts) if len(old_parts) == len(new_parts) else [(old_parts[-1], new_parts[-1])]
        for o, n in cand:
            if len(o) >= 2 and o != n and is_chinese(o):
                pairs.setdefault(o, n)
    return pairs


def _replace_name_parts(bio, old_names, new_name):
    """【2026-11】將簡介入面舊名嘅「部份」（名或姓）換成新名對應部份。回傳 (簡介, 有冇換)。
    - 新舊名用「·」分開；部份數目一樣就逐部份對應，唔一樣就只對應最後一部份（姓）。
    - 只換 ≥2 字嘅部份（單字太易誤中）；《》〈〉作品名同已經係新全名嘅位置唔會改。
    例如 舊「讓·雷諾」→ 新「尚·連奴」：「雷諾憑《這個殺手不太冷》成名」→「連奴憑《這個殺手不太冷》成名」。"""
    if not bio:
        return bio, False
    pairs = _name_part_pairs(old_names, new_name)
    if not pairs:
        return bio, False

    # 保護：新全名、作品名 → 佔位符（私用區字元），換完再還原
    keep = []
    def _hold(m):
        keep.append(m.group(0))
        return f'{len(keep) - 1}'
    work = re.sub('|'.join([re.escape(new_name), _BIO_TITLE_RE.pattern]), _hold, bio)

    replaced = False
    for o in sorted(pairs, key=len, reverse=True):
        n = pairs[o]
        pattern = re.escape(o)
        if o in n:   # 例如 舊「李維」→ 新「李維斯」：已經係「李維斯」嘅位置唔再加字
            i = n.index(o)
            if n[:i]:
                pattern = r'(?<!' + re.escape(n[:i]) + r')' + pattern
            if n[i + len(o):]:
                pattern += r'(?!' + re.escape(n[i + len(o):]) + r')'
        new_work = re.sub(pattern, n, work)
        if new_work != work:
            logging.info(f"Bio 姓名替換 (部分): 將 '{o}' 替換為 '{n}'。")
            work, replaced = new_work, True
    return re.sub('(\\d+)', lambda m: keep[int(m.group(1))], work), replaced


def replace_name_in_bio(final_bio, old_names_list, new_name):
    """
    【全新修改】在個人簡介中智能地替換人名。
    1. 優先替換完整的舊中文名（來自 zh-CN, zh-TW 等）。
    2. 若無法替換完整名稱，且新舊名稱均包含分隔符「·」，則嘗試逐部分替換。
    3. 保護已有 `中文名（英文名）` 的結構，避免錯誤替換。
    【修改 2025-09-19】更新對 normalize_to_hk 的呼叫，傳入 field_type 參數。
    """
    if not final_bio or not new_name or not old_names_list:
        return final_bio

    modified_bio = str(final_bio)
    
    # 【2026-11-14】去前後空格；中文舊名再加香港寫法（同簡介一樣用 normalize_to_hk(field_type='bio')），
    # 例如 zh-CN「大卫·爱登堡」→「大衛·愛登堡」，先配對到轉換後簡介入面嘅「大衛·愛登堡」。
    stripped = [n.strip() for n in old_names_list if isinstance(n, str) and n.strip()]
    stripped += [normalize_to_hk(n, field_type='bio') for n in stripped if is_chinese(n)]
    # 過濾空值、去重，並按長度倒序排序以優先匹配長名稱
    all_old_names_to_try = sorted(list(set(filter(None, stripped))), key=len, reverse=True)
    
    logging.info(f"Bio 姓名替換：嘗試使用 '{new_name}' 替換簡介中的舊姓名: {all_old_names_to_try}")
    
    # 分離出中文舊名和英文舊名，並過濾掉與新名稱相同的舊名
    chinese_old_names = [n for n in all_old_names_to_try if is_chinese(n) and n != new_name]
    english_old_names = [n for n in all_old_names_to_try if not is_chinese(n) and n != new_name]

    name_was_replaced = False

    # --- 策略一：替換完整的中文舊名 ---
    for old_name in chinese_old_names:
        if old_name in modified_bio:
            # 【修正】當 old_name 是 new_name 的子字串時（例如：
            #   - Prefix 情況：old='康軒', new='康軒文教事業' → 防止把已有的 '康軒文教事業' 變成 '康軒文教事業文教事業'
            #   - Suffix 情況：old='普京', new='弗拉基米爾·弗拉基米羅維奇·普京' → 防止把名字前段也替換，每次 run 多加一層
            # 解決方案：用 lookbehind (prefix) 和/或 lookahead (suffix) 避免在已完整嵌入的位置再次替換。
            if old_name in new_name and len(new_name) > len(old_name):
                idx = new_name.index(old_name)
                prefix_in_new = new_name[:idx]
                suffix_in_new = new_name[idx + len(old_name):]
                pattern = ''
                if prefix_in_new:
                    pattern += r'(?<!' + re.escape(prefix_in_new) + r')'
                pattern += re.escape(old_name)
                if suffix_in_new:
                    pattern += r'(?!' + re.escape(suffix_in_new) + r')'
                new_modified = re.sub(pattern, new_name, modified_bio)
            else:
                new_modified = modified_bio.replace(old_name, new_name)

            if new_modified != modified_bio:
                modified_bio = new_modified
                logging.info(f"Bio 姓名替換 (完整): 成功將 '{old_name}' 替換為 '{new_name}'。")
                name_was_replaced = True
                # 替換成功一個完整的就跳出，避免後續部分替換造成問題
                break
            else:
                logging.info(f"Bio 姓名替換：'{old_name}' 存在於 bio 中但已是 '{new_name}' 的一部分，跳過以防重複附加。")

    # --- 策略二【2026-11 改寫】：部份名（例如簡介後段只寫姓「雷諾」或名「基努」）---
    # 以前只喺完整名搵唔到時先做，而且要新舊名部份數目一樣；而家完整名換咗之後都會做。
    modified_bio, parts_replaced = _replace_name_parts(modified_bio, chinese_old_names, new_name)
    name_was_replaced = name_was_replaced or parts_replaced

    # --- 策略三：保護邏輯與英文名替換 ---
    # 如果已經替換了中文名，或者新的中文名已經存在於簡介中
    # 就不要再嘗試用新的中文名去替換舊的英文名，以防破壞 "新中文名 (舊英文名)" 這樣的結構
    if name_was_replaced or new_name in modified_bio:
        if name_was_replaced:
            logging.info(f"Bio 姓名替換成功，簡介內容已更新。")
        else:
            logging.info(f"Bio 姓名替換：新姓名 '{new_name}' 已存在於簡介中。為避免重複，將跳過替換英文姓名。")
        return modified_bio

    # 最後，如果以上都沒發生，才嘗試替換英文名
    for old_name in english_old_names:
        # 使用正則表達式確保只替換獨立的單詞，避免替換 "Weaver" 中的 "weaver"
        # \b 是單詞邊界
        pattern = r'\b' + re.escape(old_name) + r'\b'
        if re.search(pattern, modified_bio):
            modified_bio = re.sub(pattern, new_name, modified_bio)
            logging.info(f"Bio 姓名替換 (英文): 成功將 '{old_name}' 替換為 '{new_name}'。")
            logging.info(f"Bio 姓名替換成功，簡介內容已更新。")
            return modified_bio

    if modified_bio == final_bio:
        logging.info(f"Bio 姓名替換：未作出任何修改 (未找到可替換的姓名)。")
    
    return modified_bio

def fetch_actor_translation(api_key, person_id, preferred_language="zh-TW", fallback_language="zh-CN"):
    """
    【修改】使用 TMDB API 獲取人員的翻譯，並加入無限重試機制。
    【修改 2025-09-19】更新對 normalize_to_hk 的呼叫，傳入 field_type 參數。
    """
    headers = {"Authorization": f"Bearer {api_key}"}
    base_url = "https://api.themoviedb.org/3"
    endpoint = f"{base_url}/person/{person_id}/translations"
    # 【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
    # 【修改】使用帶有重試機制的 session
    session = create_retry_session()

    try:
        response = session.get(endpoint, headers=headers, timeout=10)
        if response.status_code != 200:
            logging.error(f"獲取人員翻譯失敗: {endpoint}, 狀態碼: {response.status_code}")
            return None, None, None, None, None, None
        data = response.json()
        translations = data.get("translations", [])
        zh_hk_name, zh_tw_name, zh_cn_name = None, None, None
        zh_hk_bio, zh_tw_bio, zh_cn_bio = None, None, None
        for t in translations:
            lang = t.get("iso_639_1", "").lower().strip()
            country = t.get("iso_3166_1", "").upper().strip()
            name = t.get("data", {}).get("name", "").strip()
            bio = t.get("data", {}).get("biography", "").strip()
            if lang == "zh" and country == "HK":
                zh_hk_name = name
                # 【修改 2025-09-19】更新呼叫，傳入 field_type
                zh_hk_bio = add_spaces_between_en_zh_num(normalize_to_hk(bio, field_type='bio'), is_long_text=True) if bio else ""
            elif lang == "zh" and country == "TW":
                zh_tw_name = name
                # 【修改 2025-09-19】更新呼叫，傳入 field_type
                zh_tw_bio = add_spaces_between_en_zh_num(normalize_to_hk(bio, field_type='bio'), is_long_text=True) if bio else ""
            elif lang == "zh" and country == "CN":
                # 【修改 2025-09-19】更新呼叫，傳入 field_type
                zh_cn_name = normalize_to_hk(name, field_type='name')  # 確保繁體化
                zh_cn_bio = add_spaces_between_en_zh_num(normalize_to_hk(bio, field_type='bio'), is_long_text=True) if bio else ""
        logging.info(f"人員 {person_id} 翻譯: zh-HK={zh_hk_name}, zh-TW={zh_tw_name}, zh-CN={zh_cn_name}, "
                       f"zh-HK Bio={zh_hk_bio[:50] if zh_hk_bio else ''}, zh-TW Bio={zh_tw_bio[:50] if zh_tw_bio else ''}, "
                       f"zh-CN Bio={zh_cn_bio[:50] if zh_cn_bio else ''}")
        return zh_hk_name, zh_tw_name, zh_cn_name, zh_hk_bio, zh_tw_bio, zh_cn_bio
    except requests.RequestException as e:
        # 只有在所有重試都失敗後，才會觸發此異常
        logging.error(f"獲取人員翻譯錯誤 (已達最大重試次數): {endpoint}, 錯誤: {e}")
        return None, None, None, None, None, None

async def _async_main(url, api_key, driver=None, overwrite_titles_str="ask", overwrite_descriptions_str="ask", bulk_mode=False, row_index=None, file_path='data/video_detail.xlsx'):
    """
    【修改】內部非同步主邏輯函數，新增 httpx 的無限重試機制。
    【使用者要求 2025-08-16】
    - 如果最終判斷無需寫入任何資料到 Excel，則應在此處刪除由 check_create_excel 在函式開頭創建的空檔案。
    【使用者要求 2025-08-17】
    - 在每次提取開始時，明確地重新載入所有字典檔 (如 tw_hk_name.xlsx) 和轉換器，確保使用最新規則。
    """
    # 【使用者要求 2025-08-17】在每次提取新項目之前，重新載入所有字典檔。
    # 這確保了在程序運行期間對 customize.txt 等檔案的任何修改都能立即生效。
    # 注意：tw_hk_name.xlsx 是在 excel_handler.py 中被獨立讀取的。
    from others import utils
    utils.reload_dictionaries_and_converters()

    tmdb_id, season_number_from_url, episode_number_from_url = extract_tmdb_id(url)
    if not tmdb_id:
        logging.error(f"無效的 TMDB URL 或無法提取 ID: {url}")
        return False, False

    check_create_excel(file_path)
    media_type = "unknown"
    if "/tv/" in url: media_type = "tv"
    elif "/movie/" in url: media_type = "movie"
    elif "/person/" in url: media_type = "person"

    overall_all_parts_are_identical_and_no_excel_write_needed = True
    no_chinese_content_body_skipped = False

    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)
    # 【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
    # 為 httpx 客戶端配置重試傳輸層，使其在遇到網路錯誤或特定伺服器錯誤時自動重試。
    transport = httpx.AsyncHTTPTransport(retries=RETRY_COUNT)
    async with httpx.AsyncClient(transport=transport, timeout=30.0) as client:
        if media_type == "tv":
            logging.info(f"處理電視節目: {url}")
            tv_show_data_tuple = await extract_tv_show_details(client, semaphore, tmdb_id, api_key, url, overwrite_titles_str, overwrite_descriptions_str, bulk_mode, row_index, file_path=file_path)
            if tv_show_data_tuple is None:
                logging.error(f"extract_tv_show_details 返回 None for TV ID {tmdb_id}")
                return False, False
            
            _, _, _, _, tv_show_parts_identical_status, _, no_chinese_content_body_skipped = tv_show_data_tuple
            
            overall_all_parts_are_identical_and_no_excel_write_needed = tv_show_parts_identical_status

        elif media_type == "movie":
            logging.info(f"處理電影: {url}")
            movie_data_tuple = await extract_movie_details(client, semaphore, tmdb_id, api_key, url, overwrite_titles_str, overwrite_descriptions_str, bulk_mode, row_index, file_path=file_path)
            if movie_data_tuple is None or \
               (movie_data_tuple[0] is None and movie_data_tuple[1] is None and movie_data_tuple[2] is None and \
                movie_data_tuple[3] and movie_data_tuple[4]):
                logging.error(f"extract_movie_details 返回無效數據 Movie ID {tmdb_id}")
                return False, False

            _, _, _, movie_and_people_identical_status, no_chinese_content_body_skipped = movie_data_tuple
            overall_all_parts_are_identical_and_no_excel_write_needed = movie_and_people_identical_status

        elif media_type == "person":
            logging.info(f"處理演職員: {url}")
            person_data_tuple = await extract_person_details(client, semaphore, person_id=tmdb_id, api_key=api_key, url=url, driver=driver,
                                                             overwrite_titles_str=overwrite_titles_str,
                                                             overwrite_descriptions_str=overwrite_descriptions_str,
                                                             bulk_mode=bulk_mode, row_index=row_index, file_path=file_path)
            if person_data_tuple is None or (person_data_tuple[0] is None and not person_data_tuple[1]):
                logging.error(f"extract_person_details 返回無效數據 Person ID {tmdb_id}")
                return False, False
            _, person_identical_status = person_data_tuple
            overall_all_parts_are_identical_and_no_excel_write_needed = person_identical_status
            no_chinese_content_body_skipped = False

        else:
            logging.error(f"無法識別的 TMDB URL 類型: {url}")
            return False, False

    logging.info(f"TMDB Extractor API 最終判斷 for {url}: 所有部分是否都無需寫入 Excel 且與 zh-HK 一致 = {overall_all_parts_are_identical_and_no_excel_write_needed}, 無中文內容跳過 = {no_chinese_content_body_skipped}")
    
    # 【使用者要求 2025-09-05】如果最終判斷影片無需任何更新 (即準備生成的 excel 檔案會是空的)，
    # 則不生成該空檔案，並直接在 bulk_videos.xlsx 中將該項目標記為完成 (Done=1)。
    # 這樣可以跳過 uploader 階段對空檔案的處理，加快整體流程。
    if overall_all_parts_are_identical_and_no_excel_write_needed:
        # 僅在 bulk 模式下，且提供了有效的行索引時，才更新 bulk_videos.xlsx
        if bulk_mode and row_index is not None and row_index >= 0:
            # 直接調用 excel_handler 中的函數，在 bulk_videos.xlsx 中將此 URL 對應的行標記為 Done=1
            save_skipped_identical_to_bulk_excel(
                row_index=row_index,
                video_url=url,
                tmdb_url=url, # 在此腳本的上下文中，video_url 和 tmdb_url 是相同的
                file_name='data/bulk_videos.xlsx'
            )

        # 【使用者要求】如果最終判斷無需寫入任何資料到 Excel，則刪除此函式開頭創建的空檔案。
        # 觸發條件：使用 PermissionError 循環嘗試刪除
        max_delete_retries = 50
        for attempt in range(max_delete_retries):
            try:
                if os.path.exists(file_path):
                    check_and_release_excel_control(file_path)
                    os.remove(file_path)
                    logging.info(f"TMDB Extractor API: 因無任何內容更新，已刪除空的快取檔案: {file_path}")
                break # Success or not exists
            except OSError as e:
                if attempt < max_delete_retries - 1:
                    time.sleep(1)
                else:
                    logging.error(f"TMDB Extractor API: 無法刪除空的快取檔案 {file_path}: {e}")

    # 確保檔案控制權被釋放 (即使檔案已被刪除，此操作也應能安全處理)
    check_and_release_excel_control(file_path)
    return overall_all_parts_are_identical_and_no_excel_write_needed, no_chinese_content_body_skipped

def main(url, api_key, driver=None, overwrite_titles_str="ask", overwrite_descriptions_str="ask", bulk_mode=False, row_index=None, file_path='data/video_detail.xlsx'):
    """
    【修改】同步主執行函數，作為非同步邏輯的入口點。
    此函數會啟動 asyncio 事件循環來執行真正的非同步工作。
    """
    try:
        # 檢查是否已有正在運行的事件循環 (例如在 Jupyter Notebook 或某些框架中)
        loop = asyncio.get_running_loop()
        # 如果有，則在此循環上運行任務
        return loop.run_until_complete(_async_main(url, api_key, driver, overwrite_titles_str, overwrite_descriptions_str, bulk_mode, row_index, file_path))
    except RuntimeError:
        # 如果沒有正在運行的事件循環，則使用 asyncio.run() 創建一個新的
        return asyncio.run(_async_main(url, api_key, driver, overwrite_titles_str, overwrite_descriptions_str, bulk_mode, row_index, file_path))