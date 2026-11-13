import time
import json
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
import re
import urllib.parse
import logging
import os
import pandas as pd
from others.split_title import split_combined_episodes
from others.excel_handler import (
    save_series_info_to_excel, save_season_info_to_excel,
    save_episode_info_to_excel, save_movie_info_to_excel,
    save_unavailable_video_info_to_excel, save_english_site_info_to_excel
)
from others.utils import add_spaces_between_en_zh_num
from others.credentials import ensure_netflix_logged_in, _select_first_netflix_profile, load_configs
import requests as _req_nf
from others.analysis_excel import is_generic_title

def extract_video_id(url):
    parsed_url = urllib.parse.urlparse(url)
    query_params = urllib.parse.parse_qs(parsed_url.query)
    path_segments = parsed_url.path.split('/')
    if 'jbv' in query_params:
        return query_params['jbv'][0]
    for segment in path_segments:
        if segment.isdigit():
            return segment
    return None

def is_chinese(text):
    return any('\u4e00' <= char <= '\u9fff' for char in text)

def handle_english_description(driver, url, description, max_retries=5):
    for _ in range(max_retries):
        if not description or is_chinese(description):
            break
        logging.warning("描述全為英文，重新載入頁面...")
        driver.get(url)
        time.sleep(2)
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')
        description_element = soup.find('div', {'class': 'title-info-synopsis'})
        description = add_spaces_between_en_zh_num(description_element.get_text(strip=True)) if description_element else ""
        time.sleep(1)
    return description if description and is_chinese(description) else None

_GENERIC_PAGE_TITLES = frozenset({
    'netflix', '首頁', '首页', '熱播影片', '热播影片', '最新上映', 'home',
    '我的最愛', '我的口袋名單', '繼續觀看', '繼續收看',
})

# 影片頁（含 /browse?jbv=<id> 彈出視窗形式）載入完成時，至少會有其中一個元素（前 3 個 = 新版 UI 核心元素，用於等待載入）
_DETAIL_SELECTORS = (
    'h3.previewModal--section-header',
    'p.preview-modal-synopsis',
    'div.episodeSelector-container',
    'img[class*="playerModel--player__storyArt"]',
    'div[data-uia="episodes"]',
    'div.videoMetadata--line',
)

def get_title_and_description(soup):
    title = ""
    description = ""

    # 優先嘗試 og 標籤（未登入版）
    title_meta = soup.find('meta', property='og:title')
    description_meta = soup.find('meta', property='og:description')
    if title_meta and title_meta.get('content'):
        candidate = add_spaces_between_en_zh_num(title_meta['content'].split('| Netflix')[0].strip())
        # 過濾泛用性頁面標題（如首頁等，見 jbv= 重定向情境）
        if candidate and candidate.lower() not in _GENERIC_PAGE_TITLES:
            title = candidate
    if description_meta and description_meta.get('content'):
        description = add_spaces_between_en_zh_num(description_meta['content'].strip())

    # 新版登入 UI：標題從「關於《...》」標題取得（處理 <strong> 包覆的情況）
    if not title:
        about_header = soup.find('h3', class_='previewModal--section-header')
        if about_header:
            header_text = about_header.get_text(strip=True)
            m = re.search(r'《(.+?)》', header_text)
            if m:
                candidate = add_spaces_between_en_zh_num(m.group(1).strip())
                if candidate and candidate.lower() not in _GENERIC_PAGE_TITLES:
                    title = candidate

    # 新版登入 UI (jbv/modal 首頁重定向)：標題從 storyArt img alt 取得
    # 當 /title/<id> 被重定向至 /latest?jbv=<id> 時，節目名稱在 storyArt 圖片的 alt 屬性
    # 注意：class 可能帶有額外樣式類別 (has-smaller-buttons detail-modal 等)，需使用更寬鬆的匹配
    if not title:
        # 使用 CSS 選擇器或 attrs 寬鬆比對，避免因額外 class 而抓不到
        story_art = soup.select_one('img.playerModel--player__storyArt, img[class*="playerModel--player__storyArt"]')
        if not story_art:
            story_art = soup.find('img', attrs={'class': re.compile(r'playerModel--player__storyArt')})
        if story_art and story_art.get('alt'):
            alt_text = story_art['alt'].strip()
            if alt_text and alt_text.lower() not in _GENERIC_PAGE_TITLES:
                title = add_spaces_between_en_zh_num(alt_text)
                logging.info(f"從 storyArt img alt 取得標題: {title}")

    # 新版登入 UI：描述從 preview-modal-synopsis 取得
    if not description:
        synopsis = soup.find('p', class_='preview-modal-synopsis')
        if synopsis:
            ptrack = synopsis.find('div', class_='ptrack-content')
            description = add_spaces_between_en_zh_num(
                ptrack.get_text(strip=True) if ptrack else synopsis.get_text(strip=True)
            )

    # 從 HTML <title> 標籤取得（登入版 Netflix 標題頁，格式為「節目名稱 | Netflix」）
    if not title and soup.title:
        raw = soup.title.get_text(strip=True)
        for suffix in [' | Netflix', '| Netflix', ' - Netflix', '- Netflix']:
            if suffix in raw:
                raw = raw[:raw.index(suffix)].strip()
                break
        # 去除書名號
        raw = re.sub(r'^《|》$', '', raw).strip()
        if raw and raw.lower() not in _GENERIC_PAGE_TITLES:
            title = add_spaces_between_en_zh_num(raw)

    # 舊版後備
    if not title:
        h1 = soup.find('h1')
        if h1:
            title = add_spaces_between_en_zh_num(h1.get_text(strip=True))
    if not description:
        desc_el = soup.find('p', class_='description')
        if desc_el:
            description = add_spaces_between_en_zh_num(desc_el.get_text(strip=True))

    return title, description

def extract_runtime(text):
    hours = 0
    minutes = 0
    match_hours = re.search(r'(\d+)\s*(小時|小时|hr|h)', text)
    match_minutes = re.search(r'(\d+)\s*(分鐘|分钟|分|min|m)', text)
    if match_hours:
        hours = int(match_hours.group(1))
    if match_minutes:
        minutes = int(match_minutes.group(1))
    return hours * 60 + minutes

def get_movie_runtime(soup):
    # 新版 UI：videoMetadata--line 內的 span.duration（只取含時間單位的文字）
    metadata_line = soup.find('div', class_='videoMetadata--line')
    if metadata_line:
        dur = metadata_line.find('span', class_='duration')
        if dur:
            text = dur.get_text(strip=True)
            if any(unit in text for unit in ['小時', '分鐘', 'hr', 'min']):
                return extract_runtime(text)
    # 舊版後備
    runtime_element = soup.find('span', {'class': 'title-info-metadata-item item-runtime'})
    runtime_text = runtime_element.find('span', {'class': 'duration'}).get_text(strip=True) if runtime_element else ""
    return extract_runtime(runtime_text)

def get_season_info(soup, series_description, tmdb_url: str = None):
    """
    解析舊版 <select name="seasonSelect"> 的季數。
    當標籤不是「第 x 季」時，會嘗試使用 TMDB 資料或數字推斷。
    """
    season_data = []
    tmdb_seasons = _fetch_tmdb_seasons(tmdb_url) if tmdb_url else []

    try:
        season_select = soup.find('select', {'name': 'seasonSelect'})
        season_options = season_select.find_all('option') if season_select else []
        for option in season_options:
            season_name = add_spaces_between_en_zh_num(option.get_text(strip=True))
            match = re.search(r'第\s*(\d+)\s*季', season_name)
            if match:
                season_number = int(match.group(1))
            else:
                # 非標準格式處理
                season_number = None
                # 嘗試年份對應 TMDB
                yr = re.search(r'(20\d{2})', season_name)
                if yr and tmdb_seasons:
                    for s in tmdb_seasons:
                        if s.get("air_year") == yr.group(1):
                            season_number = s["season_number"]
                            break
                # 再嘗試直接取數字
                if season_number is None:
                    num_m = re.search(r'(\d+)', season_name)
                    if num_m:
                        candidate = int(num_m.group(1))
                        if tmdb_seasons and any(s["season_number"] == candidate for s in tmdb_seasons):
                            season_number = candidate
                if season_number is None:
                    season_number = 1  # 最後後備

            season_data.append((season_name, season_number, '', season_name, ''))
            logging.info(f"季度名稱：{season_name}, 編號：{season_number}")
    except Exception as e:
        logging.error(f"提取季數信息時出錯: {e}")
        if not season_data:
            season_data.append(('第 1 季', 1, '', '第 1 季', ''))
    return season_data

def _click_expand_buttons(driver):
    """捲動到集數容器觸發 lazy-load，再點擊所有展開按鈕。"""
    # 先把集數容器捲入視野，讓 Netflix 渲染展開按鈕
    try:
        container = driver.find_element(By.CSS_SELECTOR, 'div.episodeSelector-container')
        driver.execute_script("arguments[0].scrollIntoView(true);", container)
        time.sleep(1.5)
    except Exception:
        driver.execute_script("window.scrollBy(0, 600);")
        time.sleep(1.5)

    try:
        buttons = driver.find_elements(By.XPATH,
            '//button[@data-uia="section-expand"]'
            '[not(ancestor::div[contains(@class,"moreLikeThis")])]')
        for btn in buttons:
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1.5)
            except Exception:
                pass
        if buttons:
            logging.info(f"已點擊 {len(buttons)} 個展開按鈕。")
        else:
            logging.info("未找到展開按鈕，集數可能已全部顯示。")
    except Exception as e:
        logging.warning(f"點擊展開按鈕時發生錯誤: {e}")

def get_episode_info(soup, standardized_url, video_url, tmdb_url, season_number, split_value=None):
    all_episode_data = []
    final_split_value = split_value.lower() if split_value else 'n' if split_value is not None else ''
    needs_split_prompt = False

    try:
        # 新版 Netflix UI：episodeSelector-container
        container = soup.find('div', class_='episodeSelector-container')
        if container:
            episode_elements = container.find_all('div', {'data-uia': 'titleCard--container'})
            logging.info(f"新版 UI：找到 {len(episode_elements)} 個集數")
        else:
            # 舊版 Netflix UI
            episodes_section = soup.find('div', {'data-uia': 'episodes'})
            if not episodes_section:
                logging.warning("未找到包含集數的 section")
                return all_episode_data, final_split_value
            episodes_list = episodes_section.find('ul')
            if not episodes_list:
                logging.warning("未找到集數列表")
                return all_episode_data, final_split_value
            episode_elements = episodes_list.find_all('li')
            logging.info(f"舊版 UI：找到 {len(episode_elements)} 個集數")

        for ep_idx, episode_element in enumerate(episode_elements, start=1):
            try:
                # 集數編號：新版用 div.titleCard-title_index，舊版從標題文字提取
                num_el = episode_element.find('div', class_='titleCard-title_index')
                if num_el and num_el.get_text(strip=True).isdigit():
                    episode_number = int(num_el.get_text(strip=True))
                else:
                    old_title_el = episode_element.find('p', class_=re.compile(r'11kf3ei'))
                    old_title_txt = old_title_el.get_text(strip=True) if old_title_el else ""
                    ep_num_match = re.search(r'第\s*(\d+)\s*集', old_title_txt)
                    episode_number = int(ep_num_match.group(1)) if ep_num_match else None

                # 後備：從 ptrack-content 的 data-ui-tracking-context 取 rank
                if episode_number is None:
                    ptrack_el = episode_element.select_one('.titleCard-imageWrapper .ptrack-content[data-ui-tracking-context]')
                    if ptrack_el:
                        try:
                            ctx = json.loads(ptrack_el.get('data-ui-tracking-context', '{}'))
                            rank = ctx.get('rank')
                            if isinstance(rank, int) and rank > 0:
                                episode_number = rank
                        except Exception:
                            pass

                # 最後後備：使用位置索引
                if episode_number is None:
                    episode_number = ep_idx
                    logging.info(f"無法取得集數編號，使用位置索引 {ep_idx}。")

                # 集數標題：新版用 span.titleCard-title_text
                title_el = episode_element.find('span', class_='titleCard-title_text')
                if title_el:
                    episode_title = add_spaces_between_en_zh_num(title_el.get_text(strip=True))
                else:
                    old_title_el = episode_element.find('p', class_=re.compile(r'11kf3ei'))
                    episode_title = add_spaces_between_en_zh_num(old_title_el.get_text(strip=True)) if old_title_el else ""

                # 集數時長：新版用 span.duration
                dur_el = episode_element.find('span', class_='duration')
                if dur_el:
                    runtime_text = dur_el.get_text(strip=True)
                else:
                    old_dur = episode_element.find('span', text=re.compile(r'\d+\s*(小時\s*)?\d+\s*分'))
                    runtime_text = old_dur.get_text(strip=True) if old_dur else ""
                episode_runtime = extract_runtime(runtime_text)

                # 集數描述：新版用 p.titleCard-synopsis > div.ptrack-content
                synopsis_el = episode_element.find('p', class_='titleCard-synopsis')
                if synopsis_el:
                    ptrack = synopsis_el.find('div', class_='ptrack-content')
                    desc_raw = ptrack.get_text(strip=True) if ptrack else synopsis_el.get_text(strip=True)
                    episode_description = add_spaces_between_en_zh_num(desc_raw)
                else:
                    old_desc_el = episode_element.find('p', class_=re.compile(r'7aasd4'))
                    episode_description = add_spaces_between_en_zh_num(old_desc_el.get_text(strip=True)) if old_desc_el else ""

                # 集數劇照 URL：取 .titleCard-imageWrapper img[src]，去除 ?.. 參數
                backdrop_url = ""
                img_el = episode_element.select_one('.titleCard-imageWrapper img[src]')
                if img_el:
                    backdrop_url = img_el.get('src', '').split('?')[0]

                if not episode_title and not episode_description and not episode_runtime:
                    logging.info("集數標題、描述和時長都為空，跳過此集數。")
                    continue
                if is_generic_title(episode_title) and not episode_description and not episode_runtime:
                    logging.info(f"第 {episode_number} 集標題為泛用型且無描述/時長，可能為未上映集數，跳過。")
                    continue
                if not is_chinese(episode_title) and not is_chinese(episode_description):
                    if not backdrop_url and not episode_runtime:
                        logging.info("集數標題和描述不包含中文且無劇照/時長，跳過此集數。")
                        continue
                    logging.info(f"第 {episode_number} 集：標題/描述非中文，保留劇照/時長但清空文字內容。")
                    episode_title = ''
                    episode_description = ''

                if ('/' in episode_title or '／' in episode_title) and not final_split_value:
                    needs_split_prompt = True
                    logging.info(f"檢測到集數標題包含 '/' 或 '／': {episode_title}")

                all_episode_data.append((season_number, episode_number, episode_runtime, episode_title, episode_description, episode_title, episode_description, backdrop_url))
                logging.info(f"第 {episode_number} 集 - 標題：{episode_title}, 時長：{episode_runtime} 分鐘")
            except Exception as e:
                logging.error(f"抓取集數資訊時發生錯誤: {e}")
                continue

        if needs_split_prompt and split_value is None:
            user_input = input("檢測到集數標題包含 '/' 或 '／'，是否需要拆分成多集？(y/n): ").strip().lower()
            final_split_value = user_input if user_input in ['y', 'n'] else 'n'
            logging.info(f"用戶選擇拆分標題: {final_split_value}")

    except Exception as e:
        logging.error(f"抓取集數資訊時發生錯誤: {e}")

    return all_episode_data, final_split_value

def _fetch_tmdb_seasons(tmdb_url: str) -> list:
    """
    從 TMDB 取得完整的季數資訊。
    返回格式: [
        {"season_number": 1, "air_year": "2023", "episode_count": 12, "name": "Season 1"},
        ...
    ]
    排序為 season_number 升冪。
    若失敗或無 tmdb_url 則返回空 list。
    """
    if not tmdb_url:
        return []
    m = re.search(r'/tv/(\d+)', tmdb_url)
    if not m:
        return []
    tmdb_id = m.group(1)
    try:
        api_key = load_configs().get('tmdb_api', '')
        if not api_key:
            return []
        r = _req_nf.get(
            f"https://api.themoviedb.org/3/tv/{tmdb_id}",
            headers={"Authorization": f"Bearer {api_key}", "accept": "application/json"},
            timeout=10,
        )
        if r.status_code != 200:
            return []
        seasons = []
        for s in r.json().get('seasons', []):
            sn = s.get('season_number', 0)
            if sn <= 0:
                continue
            air_date = s.get('air_date') or ''
            air_year = air_date[:4] if len(air_date) >= 4 else ''
            seasons.append({
                "season_number": sn,
                "air_year": air_year,
                "episode_count": s.get('episode_count', 0),
                "name": s.get('name', '')
            })
        seasons.sort(key=lambda x: x["season_number"])
        return seasons
    except Exception as e:
        logging.warning(f"TMDB seasons 查詢失敗: {e}")
        return []


def _fetch_tmdb_season_map(tmdb_url: str) -> dict:
    """
    Return {air_year_str: season_number} for a TMDB TV show, e.g. {'2019': 5, '2021': 7}.
    Used to fix year-label dropdown items whose data-index+1 gives a wrong season number.
    Returns {} if the lookup fails or no tmdb_url is given.
    (Deprecated in favor of _fetch_tmdb_seasons for new logic)
    """
    seasons = _fetch_tmdb_seasons(tmdb_url)
    year_map = {}
    for s in seasons:
        if s["air_year"] and s["air_year"] not in year_map:
            year_map[s["air_year"]] = s["season_number"]
    return year_map


def _extract_clean_season_label(item):
    """
    從季數下拉選單的 <li> 元素中，提取乾淨的季名稱。
    會自動忽略右側的「(x 集)」數量提示。
    """
    try:
        option = item.find_element(By.CSS_SELECTOR, "div.episodeSelector--option")
        # 嘗試移除 numEpisodes span 的文字
        try:
            num_span = option.find_element(By.CSS_SELECTOR, "span.episodeSelector--option-label-numEpisodes")
            full_text = option.text
            clean = full_text.replace(num_span.text, "").strip()
            # 清理可能的 &nbsp; 和多餘空白
            clean = re.sub(r'[\s\u00A0]+$', '', clean).strip()
            return clean
        except Exception:
            # 沒有 span，就直接用文字
            return option.text.strip()
    except Exception:
        # 後備
        return item.text.strip()


def _get_all_season_numbers(driver, wait, tmdb_url: str = None):
    """
    開啟季數下拉選單，讀取所有季數標籤。
    盡量正確解析 season_number，特別是當標籤不是「第 x 季」形式時（例如年份、Season X、Part X 等）。

    策略優先順序：
    1. 明確的「第 x 季」
    2. 透過 TMDB 比對年份（最可靠）
    3. 從標籤文字中提取數字（較弱）
    4. data-index + 1（最不可靠，已知 Netflix 常倒序排列）

    返回 season_data = [(season_label, season_number, '', season_label, ''), ...]，依 season_number 升序。
    """
    season_data = []
    toggle = None
    tmdb_seasons = _fetch_tmdb_seasons(tmdb_url) if tmdb_url else []

    try:
        toggle = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'button[data-uia="dropdown-toggle"]')
        ))
        driver.execute_script("arguments[0].click();", toggle)
        time.sleep(1)
        menu = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'ul[data-uia="dropdown-menu"]')
        ))
        items = menu.find_elements(By.CSS_SELECTOR, 'li[data-uia="dropdown-menu-item"]')

        raw_items = []
        for idx, item in enumerate(items):
            # 使用結構化方式提取，避免帶入 (x 集) 數量提示
            text = _extract_clean_season_label(item)
            if '查看所有集數' in text:
                continue
            data_index = item.get_attribute('data-index')
            raw_items.append({
                'text': text,
                'data_index': int(data_index) if data_index and data_index.isdigit() else None,
                'dom_position': idx
            })

        # 記錄原始標籤，方便 debug 非標準格式
        if raw_items:
            labels = [r['text'] for r in raw_items]
            if not any(re.search(r'第\s*\d+\s*季', l) for l in labels):
                logging.warning(f"偵測到非標準季數標籤（非「第 x 季」形式）: {labels}")

                # 【加強 Debug】輸出每個 menu item 的關鍵 HTML 屬性，方便提供「相關 HTML」時分析
                for i, r in enumerate(raw_items):
                    try:
                        # 重新取得對應的 WebElement 以取得更多屬性
                        item = items[i] if i < len(items) else None
                        if item:
                            outer = item.get_attribute('outerHTML')[:800]  # 限制長度避免 log 過大
                            data_index = item.get_attribute('data-index')
                            classes = item.get_attribute('class')
                            logging.warning(
                                f"[SeasonItem {i}] text='{r['text']}' | data-index='{data_index}' | class='{classes}'\n"
                                f"  outerHTML: {outer}"
                            )
                    except Exception as debug_e:
                        logging.debug(f"無法取得 item {i} 的詳細 HTML: {debug_e}")

        for r in raw_items:
            text = r['text']
            data_index = r['data_index']

            # 1. 最優先：明確「第 x 季」
            match = re.search(r'第\s*(\d+)\s*季', text)
            if match:
                sn = int(match.group(1))
                season_data.append((text, sn, '', text, ''))
                logging.info(f"從下拉選單讀取：第 {sn} 季（標籤：{text}）")
                continue

            sn = None

            # 2. 嘗試從 TMDB 年份對應（最可靠的非標準標籤解法）
            year_match = re.search(r'(20\d{2})', text)
            if year_match and tmdb_seasons:
                yr = year_match.group(1)
                for s in tmdb_seasons:
                    if s["air_year"] == yr:
                        sn = s["season_number"]
                        break

            # 3. 嘗試用 TMDB season name 做模糊比對（針對「早上來一杯」「半夜喝濃縮」這類創意名稱）
            if sn is None and tmdb_seasons:
                text_lower = text.lower()
                for s in tmdb_seasons:
                    tmdb_name = (s.get("name") or "").strip()
                    tmdb_name_lower = tmdb_name.lower()
                    if tmdb_name and (tmdb_name in text or tmdb_name_lower in text_lower or text_lower in tmdb_name_lower):
                        sn = s["season_number"]
                        logging.info(f"透過 TMDB season name 匹配成功: '{text}' → 第 {sn} 季 (TMDB name: {tmdb_name})")
                        break

            # 4. 從文字中提取任何數字作為候選（例如 "Season 2", "Part 3", "Vol.1"）
            if sn is None:
                num_match = re.search(r'(\d+)', text)
                if num_match:
                    candidate = int(num_match.group(1))
                    if tmdb_seasons and any(s["season_number"] == candidate for s in tmdb_seasons):
                        sn = candidate

            # 5. 最後手段：data-index + 1（已知常因倒序而錯誤，僅作為後備）
            if sn is None and data_index is not None:
                sn = data_index + 1
                logging.warning(f"使用 data-index 後備推斷季數（可能不準確）：{text} → 第 {sn} 季")

            if sn is None:
                # 實在無法判斷，給一個很大的數字避免衝突，後續會被 TMDB 修正或跳過
                sn = 999
                logging.error(f"無法解析季數標籤，已跳過或給予臨時編號: {text}")

            # 再次保險清理，防止極少數情況仍有殘留
            clean_label = re.sub(r'[\s\u00A0]*\(?\d+\s*集\)?\s*$', '', text).strip() or f'第 {sn} 季'
            season_data.append((clean_label, sn, '', clean_label, ''))
            logging.info(f"從下拉選單讀取：{clean_label}（推斷為第 {sn} 季）")

        # 排序 + 去重（以 season_number 為主）
        season_data.sort(key=lambda x: x[1])
        seen = set()
        deduped = []
        for item in season_data:
            if item[1] not in seen:
                seen.add(item[1])
                deduped.append(item)
        season_data = deduped

    except Exception as e:
        logging.error(f"季數下拉選單讀取失敗: {e}")
    finally:
        if toggle:
            try:
                driver.execute_script("arguments[0].click();", toggle)
                time.sleep(0.5)
            except Exception:
                pass

    return season_data

def _select_season_from_dropdown(driver, wait, season_number, season_name=None, target_label=None):
    """
    從下拉選單選擇指定季數。成功返回 True，失敗返回 False。

    改進後的選擇策略（針對非「第 x 季」標籤）：
    1. 精確匹配「第 x 季」
    2. 標籤文字包含匹配（season_name 或 target_label）
    3. 透過 data-index 反推位置（當我們知道正確的 season_number 時）
    4. 最後手段：點擊位置對應的項目（依據 TMDB 排序後的索引）
    """
    toggle = None
    try:
        toggle = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'button[data-uia="dropdown-toggle"]')
        ))
        toggle_text_before = toggle.text.strip()

        driver.execute_script("arguments[0].click();", toggle)
        time.sleep(1)
        menu = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'ul[data-uia="dropdown-menu"]')
        ))
        items = menu.find_elements(By.CSS_SELECTOR, 'li[data-uia="dropdown-menu-item"]')

        # 過濾掉「查看所有集數」
        valid_items = [it for it in items if '查看所有集數' not in _extract_clean_season_label(it)]

        clicked_label = None
        clicked = False

        for item in valid_items:
            # 使用乾淨的季名稱提取，避免 (x 集) 影響精確比對
            text = _extract_clean_season_label(item)
            first_line = text.split('\n')[0].strip()
            if '查看所有集數' in text:
                continue

            # 1. 精確「第 x 季」
            match = re.search(r'第\s*(\d+)\s*季', text)
            if match and int(match.group(1)) == season_number:
                driver.execute_script("arguments[0].click();", item)
                clicked_label = first_line[:30]
                logging.info(f"已從下拉選單選擇第 {season_number} 季")
                clicked = True
                break

            # 2. 【最重要】精確標籤文字比對（針對「早上來一杯」「半夜喝濃縮」這類創意名稱最可靠）
            if target_label and target_label.strip() == first_line:
                driver.execute_script("arguments[0].click();", item)
                clicked_label = first_line[:30]
                logging.info(f"已從下拉選單選擇第 {season_number} 季（精確標籤匹配：{clicked_label}）")
                clicked = True
                break

            # 3. 寬鬆標籤文字匹配（處理 "2025", "Season 2", "Part 1" 等）
            candidates = [str(season_number)]
            if season_name:
                candidates.append(season_name)
            if target_label:
                candidates.append(target_label)

            for cand in candidates:
                if cand and cand in first_line:
                    driver.execute_script("arguments[0].click();", item)
                    clicked_label = first_line[:30]
                    logging.info(f"已從下拉選單選擇第 {season_number} 季（文字包含匹配：{cand} → {clicked_label}）")
                    clicked = True
                    break
            if clicked:
                break

        # 3. 如果還是沒選到，嘗試用 data-index 反推點擊正確位置
        if not clicked:
            for item in valid_items:
                data_index = item.get_attribute('data-index')
                if data_index and data_index.isdigit():
                    # 假設 data-index 是從 0 開始的順序
                    # 這裡我們無法 100% 確定對應關係，但可以作為最後嘗試
                    pass  # 目前先保留，未來可加強 position mapping

        if not clicked:
            logging.warning(f"下拉選單中找不到第 {season_number} 季（所有匹配策略均失敗）")
            # 關閉選單
            driver.execute_script("arguments[0].click();", toggle)
            return False

        # 等待季數真正切換完成
        expected_fragment = f'第 {season_number} 季'
        try:
            WebDriverWait(driver, 8).until(
                lambda d: (
                    expected_fragment in
                    d.find_element(By.CSS_SELECTOR, 'button[data-uia="dropdown-toggle"]').text
                ) or (
                    d.find_element(By.CSS_SELECTOR, 'button[data-uia="dropdown-toggle"]').text.strip()
                    != toggle_text_before
                )
            )
            new_label = driver.find_element(
                By.CSS_SELECTOR, 'button[data-uia="dropdown-toggle"]'
            ).text.strip()
            logging.info(f"確認季數已切換，當前顯示：{new_label}")
        except Exception:
            logging.warning(f"等待第 {season_number} 季載入超時，嘗試繼續...")
            time.sleep(3)

        return True

    except Exception as e:
        logging.error(f"選擇第 {season_number} 季失敗: {e}")
        return False

def get_all_seasons_episode_info(soup, split_value=None):
    """
    解析「查看所有集數」視圖，提取所有季數的集數資料。
    季數邊界由 span.allEpisodeSelector-season-label 識別，集數編號依季內順序計算。
    返回 (all_episode_data, final_split_value)。
    """
    all_episode_data = []
    final_split_value = split_value.lower() if split_value else 'n' if split_value is not None else ''
    needs_split_prompt = False

    container = soup.find('div', class_='episodeSelector-container')
    if not container:
        logging.warning("查看所有集數：未找到 episodeSelector-container")
        return all_episode_data, final_split_value

    current_season_number = None
    episode_counter = 0

    for child in container.find_all(recursive=False):
        # 季數標題
        season_label = child.find('span', class_='allEpisodeSelector-season-label')
        if season_label:
            label_text = season_label.get_text(strip=True)
            match = re.search(r'第\s*(\d+)\s*季', label_text)
            if match:
                current_season_number = int(match.group(1))
            else:
                # 非標準標籤時，嘗試從 TMDB 或數字中推斷（簡易版）
                num_match = re.search(r'(\d+)', label_text)
                if num_match:
                    current_season_number = int(num_match.group(1))
                    logging.warning(f"「查看所有集數」模式偵測到非標準季數標籤，使用數字推斷：{label_text} → 第 {current_season_number} 季")
                else:
                    current_season_number = None
                    logging.warning(f"「查看所有集數」模式無法解析季數標籤，已跳過該區塊: {label_text}")

            if current_season_number is not None:
                episode_counter = 0
                logging.info(f"切換至第 {current_season_number} 季（標籤：{label_text}）")
            continue

        # 集數卡片
        if 'episode-item' not in (child.get('class') or []) or current_season_number is None:
            continue

        episode_counter += 1

        title_el = child.find('span', class_='titleCard-title_text')
        episode_title = add_spaces_between_en_zh_num(title_el.get_text(strip=True)) if title_el else ""

        dur_el = child.find('span', class_='duration')
        runtime_text = dur_el.get_text(strip=True) if dur_el else ""
        episode_runtime = extract_runtime(runtime_text)

        synopsis_el = child.find('p', class_='titleCard-synopsis')
        if synopsis_el:
            ptrack = synopsis_el.find('div', class_='ptrack-content')
            desc_raw = ptrack.get_text(strip=True) if ptrack else synopsis_el.get_text(strip=True)
            episode_description = add_spaces_between_en_zh_num(desc_raw)
        else:
            episode_description = ""

        if not episode_title and not episode_description and not episode_runtime:
            continue
        if is_generic_title(episode_title) and not episode_description and not episode_runtime:
            logging.info(f"第 {current_season_number} 季第 {episode_counter} 集標題為泛用型且無描述/時長，可能為未上映集數，跳過。")
            continue
        if not is_chinese(episode_title) and not is_chinese(episode_description):
            if not episode_runtime:
                continue
            episode_title = ''
            episode_description = ''

        if ('/' in episode_title or '／' in episode_title) and not final_split_value:
            needs_split_prompt = True

        all_episode_data.append((current_season_number, episode_counter, episode_runtime, episode_title, episode_description, episode_title, episode_description))
        logging.info(f"第 {current_season_number} 季 第 {episode_counter} 集 - {episode_title}，時長：{episode_runtime} 分鐘")

    if needs_split_prompt and split_value is None:
        user_input = input("檢測到集數標題包含 '/' 或 '／'，是否需要拆分成多集？(y/n): ").strip().lower()
        final_split_value = user_input if user_input in ['y', 'n'] else 'n'
        logging.info(f"用戶選擇拆分標題: {final_split_value}")

    return all_episode_data, final_split_value

def extract_netflix_episodes(driver, url, output_path, tmdb_url=None, split_value=None):
    video_id = extract_video_id(url)
    if not video_id:
        logging.error(f"無法提取影片 ID: {url}")
        return False, None

    ensure_netflix_logged_in(driver)

    standardized_url = f"https://www.netflix.com/title/{video_id}"
    logging.info(f"標準化URL: {standardized_url}")
    wait = WebDriverWait(driver, 20)

    # 【2026-11 K1】剛登入（尤其輸入驗證碼後）Netflix 會自行跳去 /browse，
    # 蓋過剛開啟的影片頁 → 抓到「首頁」並誤判為電影。
    # 影片頁的 URL 必定含影片 ID（/title/<id> 或 /latest?jbv=<id> 等 modal 形式），不含就重開。
    max_load_attempts = 3
    for load_attempt in range(1, max_load_attempts + 1):
        driver.get(standardized_url)
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        except Exception as e:
            logging.error(f"頁面加載超時或出錯: {e}")
            return False, None

        # 若帳戶選擇畫面阻擋內容頁，選擇後重新載入
        if _select_first_netflix_profile(driver):
            logging.info("偵測到帳戶選擇畫面，已選擇後重新載入頁面。")
            driver.get(standardized_url)
            try:
                wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
            except Exception:
                pass

        time.sleep(2)  # 讓登入後的自動跳轉有時間發生，再檢查 URL
        current_url = driver.current_url or ''
        if video_id not in current_url:
            problem = f"被跳轉到 {current_url}"
        else:
            # 等待新版 UI 的核心元素載入（任一出現即可繼續）
            try:
                wait.until(lambda d: any(d.find_elements(By.CSS_SELECTOR, sel) for sel in _DETAIL_SELECTORS[:3]))
            except Exception:
                time.sleep(3)
            # 【2026-11 K1b】URL 含影片 ID（例如 /browse?jbv=<id>）但彈出視窗未出現時，頁面其實仍是首頁：
            # 網頁標題是「首頁 - Netflix」之類、且找不到任何影片資料元素 → 視為失敗重開
            has_detail = any(driver.find_elements(By.CSS_SELECTOR, sel) for sel in _DETAIL_SELECTORS)
            _soup_chk = BeautifulSoup(driver.page_source, 'html.parser')
            page_title = _soup_chk.title.get_text(strip=True) if _soup_chk.title else ''
            page_title_core = re.split(r'\s*[-|]\s*Netflix\s*$', page_title)[0].strip().lower()
            if has_detail or (page_title_core and page_title_core not in _GENERIC_PAGE_TITLES):
                break
            problem = f"仍是「{page_title or '(無標題)'}」且未見影片資料"
        logging.warning(
            f"Netflix 影片頁{problem}（第 {load_attempt}/{max_load_attempts} 次），"
            f"可能是剛登入後的自動跳轉，{'5 秒後重新開啟' if load_attempt < max_load_attempts else '放棄'}。"
        )
        if load_attempt < max_load_attempts:
            time.sleep(5)
    else:
        logging.error(f"多次開啟後仍無法停留在影片頁 {standardized_url}，提取失敗（不會寫入資料）。")
        return False, None

    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')
    title, description = get_title_and_description(soup)
    logging.info(f"節目標題: {title}")

    # 額外診斷：如果仍然抓到泛用標題，記錄目前頁面關鍵元素以利 debug
    if title and title.lower() in _GENERIC_PAGE_TITLES:
        logging.warning("警告：抓取到泛用標題，可能是 Netflix modal / 首頁重定向狀態導致選擇器失效")
        try:
            about_h3 = soup.find('h3', class_='previewModal--section-header')
            story_img = soup.select_one('img[class*="playerModel--player__storyArt"]')
            logging.warning(f"  about_h3 text: {about_h3.get_text(strip=True) if about_h3 else 'None'}")
            logging.warning(f"  storyArt alt: {story_img.get('alt') if story_img else 'None'}")
            logging.warning(f"  page title: {soup.title.get_text(strip=True) if soup.title else 'None'}")
        except Exception:
            pass

    if not is_chinese(title) and not is_chinese(description):
        logging.info("影片標題和描述不包含中文，仍繼續提取集數資料（如劇照、時長等），但不寫入標題/描述。")
        title = ''
        description = ''

    # 偵測 UI 版本
    has_new_ui = bool(soup.find('div', class_='episodeSelector-container'))
    has_old_ui_episodes_early = bool(soup.find('div', {'data-uia': 'episodes'}))

    # ── 快速電影偵測：有片長但無集數容器 ──────────────────────────────────────
    # 電影頁的 videoMetadata--line 含「2 小時 4 分鐘」之類的片長；
    # 電視劇頁不含此元素，或內容為「共 X 集」。
    if not has_new_ui and not has_old_ui_episodes_early:
        runtime_candidate = get_movie_runtime(soup)
        if runtime_candidate == 0:
            # videoMetadata 可能尚未渲染，等待後重試
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.videoMetadata--line span.duration'))
                )
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                runtime_candidate = get_movie_runtime(soup)
            except Exception:
                pass
        if runtime_candidate > 0:
            logging.info(f"偵測為電影，片長：{runtime_candidate} 分鐘")
            save_movie_info_to_excel(title, description, runtime_candidate, title, description, output_path)
            logging.info("電影處理完成")
            return True, False

    has_season_dropdown = bool(soup.find('button', {'data-uia': 'dropdown-toggle'}))

    # ── 新版多季：下拉選單 ─────────────────────────────────────────────────────
    if has_season_dropdown:
        logging.info("偵測到新版多季下拉選單")
        season_data = _get_all_season_numbers(driver, wait, tmdb_url=tmdb_url)
        if not season_data:
            season_data = [('第 1 季', 1, description, '第 1 季', description)]

        # Fix year-format season numbers (e.g. "2019", "2021") using TMDB air dates.
        # data-index+1 gives sequential 1,2,3... which is wrong when only seasons 5,7,8
        # are available.  _fetch_tmdb_season_map maps year→real_season_number.
        _has_year_labels = any(
            not re.search(r'第\s*\d+\s*季', s[0]) for s in season_data
        )
        if _has_year_labels and tmdb_url:
            _yr_map = _fetch_tmdb_season_map(tmdb_url)
            if _yr_map:
                fixed = []
                for entry in season_data:
                    lbl, sn = entry[0], entry[1]
                    yr_m = re.search(r'(\d{4})', lbl)
                    if yr_m:
                        real_sn = _yr_map.get(yr_m.group(1))
                        if real_sn and real_sn != sn:
                            logging.info(f"年份標籤 '{lbl}' 季數修正：{sn} → {real_sn}（TMDB）")
                            entry = (lbl, real_sn) + tuple(entry[2:])
                    fixed.append(entry)
                season_data = sorted(fixed, key=lambda x: x[1])

        all_episode_data = []
        final_split_value = split_value
        for season_name, season_number, *_ in season_data:
            logging.info(f"準備提取第 {season_number} 季")
            if len(season_data) > 1:
                if not _select_season_from_dropdown(driver, wait, season_number, season_name=season_name, target_label=season_name):
                    logging.warning(f"無法選擇第 {season_number} 季，跳過")
                    continue
                # Wait for the episode card count to stabilise after the season switch.
                # Netflix may show the previous season's cards briefly while loading the new
                # season — we poll until the count stops changing for 2 consecutive checks.
                _CARD_CSS = (
                    'div.episodeSelector-container '
                    'div[data-uia="titleCard--container"]'
                )
                prev_count = -1
                stable_ticks = 0
                for _ in range(20):          # up to ~10 s
                    time.sleep(0.5)
                    cur_count = len(driver.find_elements(By.CSS_SELECTOR, _CARD_CSS))
                    if cur_count > 0 and cur_count == prev_count:
                        stable_ticks += 1
                        if stable_ticks >= 2:
                            break
                    else:
                        stable_ticks = 0
                    prev_count = cur_count
                logging.info(f"第 {season_number} 季 episode card 穩定：{prev_count} 張")
            _click_expand_buttons(driver)
            time.sleep(1)
            season_soup = BeautifulSoup(driver.page_source, 'html.parser')
            ep_data, final_split_value = get_episode_info(
                season_soup, standardized_url, url, tmdb_url, season_number, final_split_value
            )
            all_episode_data.extend(ep_data)
            logging.info(f"第 {season_number} 季：取得 {len(ep_data)} 集")

        save_series_info_to_excel(title, description, title, description, output_path)
        save_season_info_to_excel(season_data, output_path)
        if all_episode_data:
            save_episode_info_to_excel(all_episode_data, output_path)
            if final_split_value == 'y':
                split_combined_episodes(output_path)
                logging.info("已執行集數標題拆分")
        else:
            save_episode_info_to_excel([], output_path)
            logging.warning("未找到任何集數資料可寫入。")
        return True, True

    # ── 新版單季 / 舊版 ───────────────────────────────────────────────────────
    season_select_element = None
    season_names = []
    season_values = []

    try:
        season_select_element = driver.find_element(By.NAME, 'seasonSelect')
        _sel = Select(season_select_element)
        season_values = [opt.get_attribute('value') for opt in _sel.options]
        season_names = [opt.text.strip() for opt in _sel.options]
        logging.info(f"找到舊版季數選擇器，共 {len(season_values)} 季")
        is_tv_show = True
    except NoSuchElementException:
        is_tv_show = has_new_ui or has_old_ui_episodes_early
        if has_new_ui:
            logging.info("未找到季數選擇器，偵測到新版單季 UI")
        elif has_old_ui_episodes_early:
            logging.info("未找到季數選擇器，偵測到舊版單季 UI")
        else:
            logging.info("未找到季數選擇器，偵測為電影")
    except Exception as e:
        logging.error(f"檢查季數選擇器時出錯: {e}")
        is_tv_show = has_new_ui

    if is_tv_show:
        save_series_info_to_excel(title, description, title, description, output_path)
        season_data = get_season_info(soup, description, tmdb_url=tmdb_url) if season_select_element else [('第 1 季', 1, description, '第 1 季', description)]
        save_season_info_to_excel(season_data, output_path)

        all_episode_data = []
        final_split_value = split_value
        for season_name, season_number, *_ in season_data:
            logging.info(f"處理季數 {season_number}: {season_name}")
            try:
                if season_select_element and len(season_data) > 1:
                    season_value = next(
                        (val for name, val in zip(season_names, season_values) if name == season_name), None
                    )
                    if season_value is None:
                        logging.warning(f"找不到季數值對應於季名: {season_name}")
                        continue
                    Select(season_select_element).select_by_value(season_value)
                    logging.info(f"已選擇季數: {season_name}")
                    time.sleep(3)

                if has_new_ui:
                    _click_expand_buttons(driver)
                    time.sleep(2)
                else:
                    try:
                        eps_el = driver.find_element(By.CSS_SELECTOR, 'div[data-uia="episodes"]')
                        driver.execute_script("arguments[0].scrollIntoView();", eps_el)
                        time.sleep(3)
                    except Exception as e:
                        logging.error(f"滾動到集數區域時發生錯誤: {e}")
                        continue
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-uia="episodes"] ul')))
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-uia="episodes"] ul li')))
                    wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')

                page_source = driver.page_source
                season_soup = BeautifulSoup(page_source, 'html.parser')
                episode_data, updated_split_value = get_episode_info(
                    season_soup, standardized_url, url, tmdb_url, season_number, final_split_value
                )
                final_split_value = updated_split_value
                all_episode_data.extend(episode_data)
            except Exception as e:
                logging.error(f"處理季數 {season_number} 時發生錯誤: {e}")
                continue

        if all_episode_data:
            save_episode_info_to_excel(all_episode_data, output_path)
            if final_split_value == 'y':
                split_combined_episodes(output_path)
                logging.info("已執行集數標題拆分")
        else:
            save_episode_info_to_excel([], output_path)
            logging.warning("未找到任何集數資料可寫入。")
        return True, True
    else:
        runtime = get_movie_runtime(soup)
        if runtime == 0:
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.videoMetadata--line span.duration'))
                )
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                runtime = get_movie_runtime(soup)
            except Exception:
                pass
        save_movie_info_to_excel(title, description, runtime if runtime > 0 else None, title, description, output_path)
        logging.info(f"電影處理完成，片長：{runtime} 分鐘" if runtime and runtime > 0 else "電影處理完成（未取得片長）")
        return True, False