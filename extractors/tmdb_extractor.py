import requests
from bs4 import BeautifulSoup
import re
import logging
import time
from importers.tmdb_uploader import extract_tv_show_id, extract_movie_id
from others.excel_handler import (
    save_series_info_to_excel, save_season_info_to_excel, 
    save_episode_info_to_excel, save_movie_info_to_excel, check_create_excel, 
    save_english_site_info_to_excel, save_unavailable_video_info_to_excel,
    save_skipped_identical_video_info_to_excel, save_processed_episode_info_to_excel
)
from others.analysis_excel import is_generic_title, clean_generic_title
# 【修改】導入中央正規化函數，取代本地的 OpenCC 實例
from others.utils import normalize_to_hk

# 【修改】移除了此處的 opencc import 和 converter = opencc.OpenCC('s2hk.json')
# 這是造成 .json.json 錯誤的直接原因。所有轉換操作現在都將透過 utils.py 中的 normalize_to_hk 函數處理。

# 不需要的描述文本
NO_DESCRIPTION_TEXT = [
    "目前沒有相關簡介，您可協助新增簡介資料。",
    "There are no episodes added to this season.",
    "暂无英文的剧情简介，請幫助我们添加到数据库。",
    "沒有翻譯為正體中文的概要，請協助我們新增資料。",
    "剧情简介缺失，請幫助我们更新。"
]

def extract_tmdb_id(url):
    if "/tv/" in url:
        return extract_tv_show_id(url)
    elif "/movie/" in url:
        return extract_movie_id(url)
    return None

def extract_details(tag, default=""):
    if tag:
        text = tag.text.strip()
        if text in NO_DESCRIPTION_TEXT:
            return ""
        return text
    return default

def contains_chinese(text):
    """檢查文本是否包含中文字符"""
    return any('\u4e00' <= char <= '\u9fff' for char in text) if text else False

# 【修改】移除了本地的 translate_to_hk 函數，因為現在將使用 utils.normalize_to_hk

def add_spaces_between_en_zh_num(text):
    """在英文、漢字和數字之間添加空格，與 tmdb_uploader.py 一致"""
    if not text or not isinstance(text, str):
        return text
    text = re.sub(r'([a-zA-Z0-9])([\u4e00-\u9fff])', r'\1 \2', text)
    text = re.sub(r'([\u4e00-\u9fff])([a-zA-Z0-9])', r'\1 \2', text)
    text = re.sub(r' +', ' ', text)
    return text.strip()

def fetch_with_fallback(url, tmdb_id, media_type, season_number=None, episode_number=None, fallback_language="zh-CN", overwrite_descriptions="ask"):
    """使用爬蟲獲取指定語言的內容，若 zh-TW 無效則 fallback 到 zh-CN，並與 zh-HK 比較"""
    time.sleep(0.5)  # 添加延遲
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    description_tag = soup.find('div', class_='overview')
    description = extract_details(description_tag.find('p') if description_tag else None)
    title_tag = soup.find('h2') or soup.find('h3')
    title = extract_details(title_tag.find('a') if title_tag else None)
    
    # 正規化 zh-TW 數據
    valid_title = contains_chinese(title) and not is_generic_title(title) and title.strip() != ""
    valid_description = description.strip() != "" and description not in NO_DESCRIPTION_TEXT
    
    logging.info(f"zh-TW 頁面數據: 標題={title}, 描述={description[:50]}...")
    
    # 如果 zh-TW 標題或描述無效，嘗試 zh-CN
    if not (valid_title and valid_description):
        logging.info(f"zh-TW 數據不完整，嘗試獲取 {fallback_language} 內容: {url}")
        fallback_url = url.replace("language=zh-TW", f"language={fallback_language}")
        time.sleep(0.5)  # 添加延遲
        try:
            fallback_response = requests.get(fallback_url)
            fallback_soup = BeautifulSoup(fallback_response.content, 'html.parser')
        except requests.RequestException as e:
            logging.error(f"zh-CN 頁面請求失敗: {fallback_url}, 錯誤: {e}")
            return None, "", soup, False, False, False, False
        
        description_tag = fallback_soup.find('div', class_='overview')
        fallback_description = extract_details(description_tag.find('p') if description_tag else None)
        fallback_title_tag = fallback_soup.find('h2') or fallback_soup.find('h3')
        fallback_title = extract_details(fallback_title_tag.find('a') if fallback_title_tag else None)
        
        # 【修改】使用來自 utils.py 的中央轉換函數
        # 對 zh-CN 數據進行繁體化
        if contains_chinese(fallback_title):
            fallback_title = normalize_to_hk(fallback_title)
        if contains_chinese(fallback_description) and fallback_description not in NO_DESCRIPTION_TEXT:
            fallback_description = normalize_to_hk(fallback_description)
        
        # 更新最終數據
        if not valid_title and contains_chinese(fallback_title) and not is_generic_title(fallback_title):
            title = fallback_title
            valid_title = True
            logging.info(f"使用 zh-CN 繁體化標題: {title}")
        if not valid_description and fallback_description.strip() != "" and fallback_description not in NO_DESCRIPTION_TEXT:
            description = fallback_description
            valid_description = True
            logging.info(f"使用 zh-CN 繁體化描述: {description[:50]}...")
    
    # 請求 zh-HK 數據進行比較
    zh_hk_url = url.replace("language=zh-TW", "language=zh-HK")
    time.sleep(0.5)  # 添加延遲
    try:
        zh_hk_response = requests.get(zh_hk_url)
        zh_hk_soup = BeautifulSoup(zh_hk_response.content, 'html.parser')
    except requests.RequestException as e:
        logging.error(f"zh-HK 頁面請求失敗: {zh_hk_url}, 錯誤: {e}")
        zh_hk_title, zh_hk_description = "", ""
        zh_hk_has_data = False
        zh_hk_has_non_generic_title = False
        zh_hk_has_valid_description = False
    else:
        description_tag = zh_hk_soup.find('div', class_='overview')
        zh_hk_description = extract_details(description_tag.find('p') if description_tag else None)
        title_tag = zh_hk_soup.find('h2') or zh_hk_soup.find('h3')
        zh_hk_title = extract_details(title_tag.find('a') if title_tag else None)
        
        zh_hk_has_data = bool(zh_hk_title.strip() or zh_hk_description.strip())
        zh_hk_has_non_generic_title = zh_hk_has_data and not is_generic_title(zh_hk_title)
        zh_hk_has_valid_description = zh_hk_description.strip() != "" and zh_hk_description not in NO_DESCRIPTION_TEXT
        
        # 如果 zh-HK 有非泛型標題，保留 zh-HK 標題
        if zh_hk_has_non_generic_title:
            title = zh_hk_title
            valid_title = True
            logging.info(f"zh-HK 有非泛型標題，保留 zh-HK 標題: {title}")
        
        # 格式化標題和描述以比較
        formatted_title = add_spaces_between_en_zh_num(title) if title else ""
        formatted_zh_hk_title = add_spaces_between_en_zh_num(zh_hk_title) if zh_hk_title else ""
        formatted_description = add_spaces_between_en_zh_num(description) if description else ""
        formatted_zh_hk_description = add_spaces_between_en_zh_num(zh_hk_description) if zh_hk_description else ""
        
        # 標題比較
        if zh_hk_has_non_generic_title or (is_generic_title(zh_hk_title) and not formatted_title):
            title_match = True
        else:
            title_match = formatted_title == formatted_zh_hk_title
        
        # 描述比較
        if overwrite_descriptions == "n":
            if zh_hk_has_valid_description:
                description_match = True
                if valid_description:
                    description = ""
                logging.info(f"zh-HK 描述非空，保留 zh-HK 描述")
            elif valid_description:
                description_match = False
                logging.info(f"zh-HK 描述為空，使用 zh-TW/zh-CN 描述")
            else:
                description_match = True
                description = ""
                logging.info(f"zh-HK 描述為空，無有效 zh-TW/zh-CN 描述，保持空描述")
        else:
            description_match = formatted_description == formatted_zh_hk_description
        
        is_same_as_zh_hk = title_match and description_match
    
    has_valid_translation = valid_title or valid_description
    
    logging.info(f"zh-HK 數據: 標題={zh_hk_title}, 描述={zh_hk_description[:50]}...")
    logging.info(f"比較結果: 標題匹配={title_match}, 描述匹配={description_match}, 整體相同={is_same_as_zh_hk}")
    logging.info(f"最終狀態: 有效標題={valid_title}, 有效描述={valid_description}, 整體有效={has_valid_translation}, 與 zh-HK 相同={is_same_as_zh_hk}, zh-HK 有資料={zh_hk_has_data}, zh-HK 非泛型標題={zh_hk_has_non_generic_title}")
    
    return title, description, soup, has_valid_translation, is_same_as_zh_hk, zh_hk_has_data, zh_hk_has_non_generic_title

def is_broadcast_info(text):
    return all(keyword in text for keyword in ["首播", "年", "月", "日"])

def is_default_description(text):
    """檢查描述是否為 TMDB 的預設格式：{title} 第 X 季已于 YYYY 年 MM 月 DD 日播出。"""
    pattern = r'《.+》第 \d+ 季已于 \d{4} 年 \d{2} 月 \d{2} 日播出。'
    return bool(re.match(pattern, text.strip()))

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

def extract_tv_show_details(tmdb_id, url, overwrite_descriptions="ask"):
    url_tw = f"https://www.themoviedb.org/tv/{tmdb_id}?language=zh-TW"
    title, description, soup, has_valid_translation, is_same_as_zh_hk, zh_hk_has_data, zh_hk_has_non_generic_title = fetch_with_fallback(
        url_tw, tmdb_id, media_type="tv", overwrite_descriptions=overwrite_descriptions
    )
    
    if title is None:
        logging.warning(f"TV Show {tmdb_id} 在 zh-TW 和 zh-CN 均無有效中文標題，跳過處理，URL: {url}")
        save_unavailable_video_info_to_excel(url, url)
        return None, None, "zh-TW", None, zh_hk_has_data, None
    
    if has_valid_translation and not is_same_as_zh_hk:
        logging.info(f"TV Show {tmdb_id} 需要寫入: 標題={title}, 描述={description[:50]}...")
        save_series_info_to_excel(title, description)
    elif zh_hk_has_data and is_same_as_zh_hk:
        logging.info(f"TV Show {tmdb_id} 與 zh-HK 相同，跳過處理")
        save_skipped_identical_video_info_to_excel(url, url)
        return None, None, "zh-TW", None, True, None
    else:
        logging.info(f"TV Show {tmdb_id} 無有效中文標題或 zh-HK 有非泛型標題，跳過處理，URL: {url}")
        save_skipped_identical_video_info_to_excel(url, url)
        return None, None, "zh-TW", None, True, None
    
    return title, description, "zh-TW", "zh-CN", is_same_as_zh_hk, title

def extract_season_details(tmdb_id, url, overwrite_descriptions="ask"):
    url_tw = f"https://www.themoviedb.org/tv/{tmdb_id}/seasons?language=zh-TW"
    time.sleep(0.5)  # 添加延遲
    response = requests.get(url_tw)
    soup = BeautifulSoup(response.content, 'html.parser')

    seasons = []
    season_numbers = []
    season_wrappers = soup.find_all('div', class_='season_wrapper')
    all_same_as_zh_hk = True
    
    for wrapper in season_wrappers:
        season_tag = wrapper.find('h2')
        season_number = re.search(r'/season/(\d+)', season_tag.find('a')['href']).group(1) if season_tag else ""
        season_numbers.append(int(season_number) if season_number.isdigit() else season_number)
        
        season_url_tw = f"https://www.themoviedb.org/tv/{tmdb_id}/season/{season_number}?language=zh-TW"
        season_name, season_description, _, has_valid_translation, is_same_as_zh_hk, zh_hk_has_data, zh_hk_has_non_generic_title = fetch_with_fallback(
            season_url_tw, tmdb_id, media_type="tv", season_number=season_number, overwrite_descriptions=overwrite_descriptions
        )
        
        if season_name is None:
            logging.warning(f"Season {season_number} 無有效中文標題，跳過該季")
            continue
        
        if season_description in NO_DESCRIPTION_TEXT or not contains_chinese(season_description):
            season_description = ""
        
        if is_same_as_zh_hk:
            logging.info(f"Season {season_number} 與 zh-HK 相同，跳過")
            continue
        
        if has_valid_translation and not is_same_as_zh_hk:
            seasons.append((season_name, season_number, season_description))
            all_same_as_zh_hk = False
            logging.info(f"Season {season_number} 需要寫入: 標題={season_name}, 描述={season_description[:50]}...")
        else:
            logging.info(f"跳過 Season {season_number}: 無有效翻譯或 zh-HK 有非泛型標題")
    
    return seasons, season_numbers, all_same_as_zh_hk

def extract_episode_details(driver, tmdb_id, season_number, url, overwrite_descriptions="ask"):
    url_tw = f"https://www.themoviedb.org/tv/{tmdb_id}/season/{season_number}?language=zh-TW"
    driver.get(url_tw)
    time.sleep(0.5)  # 添加延遲
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    episodes = []
    episode_list = soup.find_all('div', class_='info')
    all_same_as_zh_hk = True
    
    for episode in episode_list:
        episode_number_tag = episode.find('span', class_='episode_number')
        episode_number = extract_details(episode_number_tag)
        
        episode_url_tw = f"https://www.themoviedb.org/tv/{tmdb_id}/season/{season_number}/episode/{episode_number}?language=zh-TW"
        episode_title, episode_description, _, has_valid_translation, is_same_as_zh_hk, zh_hk_has_data, zh_hk_has_non_generic_title = fetch_with_fallback(
            episode_url_tw, tmdb_id, media_type="tv", season_number=season_number, episode_number=episode_number, overwrite_descriptions=overwrite_descriptions
        )
        
        if episode_title is None:
            logging.warning(f"Episode {season_number}-{episode_number} 無有效中文標題，跳過該集")
            save_processed_episode_info_to_excel(url, url, status="skipped")
            continue
        
        if episode_description in NO_DESCRIPTION_TEXT or not contains_chinese(episode_description):
            episode_description = ""
        
        runtime_tag = episode.find('span', class_='runtime')
        episode_runtime = extract_details(runtime_tag)
        episode_runtime = convert_runtime_to_minutes(episode_runtime) if episode_runtime else None
        
        if is_same_as_zh_hk:
            logging.info(f"Episode {season_number}-{episode_number} 與 zh-HK 相同，跳過")
            save_processed_episode_info_to_excel(url, url, status="skipped")
            continue
        
        if has_valid_translation and not is_same_as_zh_hk:
            episodes.append((season_number, int(episode_number) if episode_number.isdigit() else None, episode_runtime, episode_title, episode_description))
            all_same_as_zh_hk = False
            logging.info(f"Episode {season_number}-{episode_number} 需要寫入: 標題={episode_title}, 描述={episode_description[:50]}...")
        else:
            logging.info(f"跳過 Episode {season_number}-{episode_number}: 無有效翻譯或 zh-HK 有非泛型標題")
            save_processed_episode_info_to_excel(url, url, status="skipped")
    
    return episodes, all_same_as_zh_hk

def extract_movie_details(tmdb_id, url, overwrite_descriptions="ask"):
    url_tw = f"https://www.themoviedb.org/movie/{tmdb_id}?language=zh-TW"
    title, description, soup, has_valid_translation, is_same_as_zh_hk, zh_hk_has_data, zh_hk_has_non_generic_title = fetch_with_fallback(
        url_tw, tmdb_id, media_type="movie", overwrite_descriptions=overwrite_descriptions
    )
    
    if title is None:
        logging.warning(f"Movie {tmdb_id} 在 zh-TW 和 zh-CN 均無有效中文標題，跳過處理，URL: {url}")
        save_unavailable_video_info_to_excel(url, url)
        return None, None, None, zh_hk_has_data
    
    runtime_tag = soup.find('span', class_='runtime')
    runtime = extract_details(runtime_tag)
    runtime = convert_runtime_to_minutes(runtime) if runtime else None
    
    if has_valid_translation and not is_same_as_zh_hk:
        logging.info(f"Movie {tmdb_id} 需要寫入: 標題={title}, 描述={description[:50]}..., 運行時長={runtime}")
        return title, description, runtime, is_same_as_zh_hk
    else:
        logging.info(f"Movie {tmdb_id} 與 zh-HK 相同或無需更新，跳過處理")
        save_skipped_identical_video_info_to_excel(url, url)
        return None, None, None, True

def main(url, driver=None, overwrite_titles="ask", overwrite_descriptions="ask"):
    from selenium.webdriver.remote.webdriver import WebDriver

    if not isinstance(driver, WebDriver):
        logging.error("無效的 driver：預期為 WebDriver 實例，收到 %s", type(driver))
        print("錯誤：無效的瀏覽器驅動程式")
        return False
    
    tmdb_id = extract_tmdb_id(url)
    if not tmdb_id:
        logging.error(f"無效的 TMDB URL: {url}")
        print("無效的 TMDB URL。")
        return False

    check_create_excel()

    all_same_as_zh_hk = True
    if "/tv/" in url:
        logging.info(f"處理電視節目: {url}")
        result = extract_tv_show_details(tmdb_id, url, overwrite_descriptions)
        if result is None or result[0] is None:
            logging.warning(f"TV Show {tmdb_id} 無 zh-TW 或 zh-CN 翻譯，跳過處理")
            print(f"TV Show {tmdb_id} 無 zh-TW 或 zh-CN 翻譯，跳過處理。")
            return False
        title, description, preferred_language, fallback_language, is_same_as_zh_hk, tv_show_title = result
        all_same_as_zh_hk = is_same_as_zh_hk

        if not is_same_as_zh_hk:
            if not contains_chinese(title) and not contains_chinese(description):
                logging.info(f"僅有英文標題: {title}，記錄到 unavailable_video.xlsx 並繼續處理")
                save_english_site_info_to_excel(url, url)

        seasons, season_numbers, seasons_same_as_zh_hk = extract_season_details(tmdb_id, url, overwrite_descriptions)
        seasons_to_save = [(clean_generic_title(name) if is_generic_title(name) else name, num, desc) 
                           for name, num, desc in seasons]
        if seasons_to_save:
            save_season_info_to_excel(seasons_to_save)
            logging.info(f"保存季度到 video_detail.xlsx: {len(seasons_to_save)} 個季度")
        all_same_as_zh_hk = all_same_as_zh_hk and seasons_same_as_zh_hk

        all_episodes = []
        for season_number in season_numbers:
            episodes, episodes_same_as_zh_hk = extract_episode_details(driver, tmdb_id, season_number, url, overwrite_descriptions)
            all_episodes.extend(episodes)
            all_same_as_zh_hk = all_same_as_zh_hk and episodes_same_as_zh_hk
        if all_episodes:
            save_episode_info_to_excel(all_episodes)
            logging.info(f"保存集數到 video_detail.xlsx: {len(all_episodes)} 個集數")

    elif "/movie/" in url:
        logging.info(f"處理電影: {url}")
        title, description, runtime, is_same_as_zh_hk = extract_movie_details(tmdb_id, url, overwrite_descriptions)
        if title is None and description is None:
            logging.warning(f"Movie {tmdb_id} 無 zh-TW 或 zh-CN 翻譯，跳過處理")
            print(f"Movie {tmdb_id} 無 zh-TW 或 zh-CN 翻譯，跳過處理。")
            return False
        all_same_as_zh_hk = is_same_as_zh_hk
        if not is_same_as_zh_hk:
            save_movie_info_to_excel(title, description, runtime)
            logging.info(f"保存電影到 video_detail.xlsx: 標題={title}, 描述={description[:50]}..., 運行時長={runtime}")

    logging.info(f"數據處理完成，全部與 zh-HK 相同={all_same_as_zh_hk}")
    if all_same_as_zh_hk:
        print("所有資料與 zh-HK 相同，無需寫入 video_detail.xlsx。")
        save_skipped_identical_video_info_to_excel(url, url)
    else:
        print("數據已成功保存到 video_detail.xlsx。")
    return all_same_as_zh_hk
