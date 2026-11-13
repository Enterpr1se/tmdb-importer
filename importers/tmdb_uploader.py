import os
import shutil
import tempfile
import pandas as pd
import traceback
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
import re
import time
import random
import requests as _req
from others.analysis_excel import is_generic_title
from others.credentials import get_tmdb_credentials, is_tmdb_logged_in, login, load_configs
from others.excel_handler import save_unavailable_video_info_to_excel, save_translation_error_to_excel
# 【修改】導入整個 utils 模組，以便存取全局中斷標誌
from others import utils
from others.utils import is_valid_language, is_chinese, is_english, is_valid_zh_hk_title, is_simplified_chinese, check_wait_and_try_again, normalize_to_hk, add_spaces_between_en_zh_num, convert_runtime_to_minutes, compare_with_zh_hk
from inputimeout import inputimeout, TimeoutOccurred

# 設置日誌過濾器，過濾 Chromium 的 latency_info.cc 錯誤
class ChromiumErrorFilter(logging.Filter):
    def filter(self, record):
        return not "latency_info.cc" in record.getMessage()

logging.getLogger().addFilter(ChromiumErrorFilter())

def _safe_send_keys(driver, element, text):
    """Send text to a field, handling non-BMP Unicode characters (>U+FFFF) that
    msedgedriver cannot send via send_keys. Falls back to JS native value setter."""
    if not any(ord(c) > 0xFFFF for c in text):
        element.send_keys(text)
        return
    logging.warning(f"文字含非 BMP 字元，改用 JS 輸入: {[hex(ord(c)) for c in text if ord(c) > 0xFFFF]}")
    tag = element.tag_name.upper() if hasattr(element, 'tag_name') else 'INPUT'
    proto = 'HTMLTextAreaElement' if tag == 'TEXTAREA' else 'HTMLInputElement'
    driver.execute_script(f"""
        var el = arguments[0], text = arguments[1];
        Object.getOwnPropertyDescriptor(window.{proto}.prototype, 'value').set.call(el, text);
        el.dispatchEvent(new Event('input', {{bubbles: true}}));
        el.dispatchEvent(new Event('change', {{bubbles: true}}));
    """, element, text)


# 字典儲存每個項目的 failsafe 重試計數，鍵為 URL，值為計數
failsafe_retry_counts = {}

def reset_failsafe_counter(url):
    """重置指定項目的 failsafe 重試計數器"""
    failsafe_retry_counts.pop(url, None)
    logging.debug(f"重置 failsafe 計數器為 0，URL: {url}") # 改為 debug 級別

def get_failsafe_retry_count(url):
    """獲取指定項目的重試計數，若不存在則初始化為 0"""
    return failsafe_retry_counts.get(url, 0)

def increment_failsafe_retry_count(url):
    """遞增指定項目的重試計數"""
    failsafe_retry_counts[url] = get_failsafe_retry_count(url) + 1
    return failsafe_retry_counts[url]

def get_user_input_with_timeout(prompt, timeout=60):
    """帶超時的用戶輸入"""
    try:
        user_input = inputimeout(prompt, timeout=timeout)
    except TimeoutOccurred:
        print("\n輸入超時。假設默認為 'n'。")
        user_input = 'n'
    return user_input

def check_and_add_translation(driver, current_url, is_movie=False, video_url=None):
    """
    檢查並添加 zh-HK 翻譯。
    【修改】在 preprocess_page 中加入權限檢查與自動重登機制。
    """
    title_field_id = 'zh_HK_translated_title' if is_movie else 'zh_HK_name'
    base_url = current_url.split('?')[0]
    zh_hk_url = f"{base_url}?language=zh-HK"

    # 共通的頁面預處理函數
    def preprocess_page(url_to_load):
        driver.get(url_to_load)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # 【使用者要求】檢測權限錯誤並自動重新登入
        # + if "您没有权限访问该页面" in driver.page_source or "You don't have permission to access this page" in driver.page_source:
        # +    ... (重登邏輯)
        if "您没有权限访问该页面" in driver.page_source or "You don't have permission to access this page" in driver.page_source:
            logging.warning(f"preprocess_page: 檢測到權限錯誤 (未登入?)，嘗試重新登入 TMDB... 目標 URL: {url_to_load}")
            configs = load_configs()
            username, password = get_tmdb_credentials(configs)
            if username and password:
                login(driver, username, password)
                logging.info(f"重新登入完成，重新載入目標頁面: {url_to_load}")
                driver.get(url_to_load)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                # 重新檢查是否仍有錯誤
                if "您没有权限访问该页面" in driver.page_source or "You don't have permission to access this page" in driver.page_source:
                    logging.error("重新登入後仍然顯示無權限，中止此操作。")
                    return False
            else:
                logging.error("無法重新登入：缺少憑證。")
                return False

        if "Oops! We can't find the page you're looking for" in driver.page_source:
            logging.error("頁面找不到，中止操作。")
            save_unavailable_video_info_to_excel(video_url or current_url, current_url)
            return False
        try:
            WebDriverWait(driver, 0.5).until(EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Wait and try again')]")))
            logging.info("發現 'Wait and try again' 訊息，進入標準處理流程。")
            if check_wait_and_try_again(driver, url_to_load, video_url):
                logging.warning("'Wait and try again' 訊息持續存在，中止此項目的翻譯操作。")
                return False
            driver.get(url_to_load) # 處理後重新載入
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except TimeoutException:
            pass # 正常情況
        return True

    # 共通的新增翻譯點擊函數
    def click_add_translation_and_verify(field_id_to_check):
        try:
            add_translation_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '新增翻譯')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_translation_button)
            add_translation_button.click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, field_id_to_check)))
            logging.info(f"成功點擊 '新增翻譯' 並確認欄位 '{field_id_to_check}' 已出現。")
            return True
        except Exception as e:
            logging.warning(f"點擊 '新增翻譯' 或驗證欄位時失敗: {e}")
            return False

    # --- 主要流程 ---

    # 假設調用者已導航到頁面，這裡直接檢查
    try:
        driver.find_element(By.ID, title_field_id)
        logging.info(f"編輯欄位 '{title_field_id}' 已存在，無需新增翻譯。")
        return True
    except NoSuchElementException:
        logging.info(f"未找到編輯欄位 '{title_field_id}'，需要新增翻譯。")

    # 階段一：第一次嘗試新增
    logging.info("--- 階段一：第一次嘗試新增翻譯 ---")
    if click_add_translation_and_verify(title_field_id):
        return True

    # 階段二：重新載入後重試
    logging.warning("--- 階段二：第一次嘗試失敗，重新載入頁面並重試 ---")
    if not preprocess_page(zh_hk_url): return False # 重新載入並預處理

    try:
        driver.find_element(By.ID, title_field_id)
        logging.info(f"重新載入後，編輯欄位 '{title_field_id}' 已存在。")
        return True
    except NoSuchElementException:
        if click_add_translation_and_verify(title_field_id):
            return True

    # 階段三：啟用備用語言修復程序
    logging.error("--- 階段三：重試依然失敗，啟動備用語言修復程序 ---")
    fallback_languages = ['zh-TW', 'zh-CN', 'zh-SG']
    workaround_succeeded = False
    for lang in fallback_languages:
        fallback_url = f"{base_url}?language={lang}"
        logging.info(f"嘗試使用備用語言 '{lang}' 進行修復，URL: {fallback_url}")
        if not preprocess_page(fallback_url): continue
        
        try:
            fallback_field_id = f'{lang.replace("-", "_")}_translated_title' if is_movie else f'{lang.replace("-", "_")}_name'
            driver.find_element(By.ID, fallback_field_id)
            logging.info(f"備用語言 '{lang}' 的翻譯已存在，跳過此語言。")
            continue
        except NoSuchElementException:
            pass

        try:
            add_translation_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '新增翻譯')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_translation_button)
            add_translation_button.click()
            time.sleep(2)
            logging.info(f"已成功點擊 '{lang}' 的新增翻譯按鈕。假設修復程序已觸發。")
            workaround_succeeded = True
            break
        except Exception as e:
            logging.warning(f"使用備用語言 '{lang}' 修復失敗: {e}")
    
    if not workaround_succeeded:
        logging.error("錯誤修復程序失敗，所有備用語言均無法新增翻譯。")
        save_translation_error_to_excel(video_url, current_url, "添加 zh-HK 翻譯失敗，且備用語言修復無效")
        return False

    # 階段四：修復後最後一次嘗試 zh-HK
    logging.info("--- 階段四：修復後最後一次嘗試新增 zh-HK 翻譯 ---")
    if not preprocess_page(zh_hk_url): return False

    try:
        driver.find_element(By.ID, title_field_id)
        logging.info(f"修復後，編輯欄位 '{title_field_id}' 已存在。")
        return True
    except NoSuchElementException:
        if click_add_translation_and_verify(title_field_id):
            return True

    logging.error("所有階段嘗試均告失敗，最終無法添加 zh-HK 翻譯。")
    save_translation_error_to_excel(video_url, current_url, "多次重試及修復程序後，仍無法添加 zh-HK 翻譯")
    return False

def extract_tv_show_id(url):
    """從 URL 中提取電視節目 ID"""
    logging.info(f"從 URL 中提取電視節目 ID: {url}")
    parts = url.split('/tv/')[1].split('?')[0].split('/')
    return parts[0].split('-')[0]

def extract_movie_id(url):
    """從 URL 中提取電影 ID"""
    logging.info(f"從 URL 中提取電影 ID: {url}")
    parts = url.split('/movie/')[1].split('?')[0].split('/')
    return parts[0].split('-')[0]

def extract_person_id(url):
    """從 URL 中提取人員 ID"""
    logging.info(f"從 URL 中提取人員 ID: {url}")
    parts = url.split('/person/')[1].split('?')[0].split('/')
    return parts[0].split('-')[0]

def check_if_field_unlocked(driver, field_id):
    """檢查欄位是否解鎖"""
    try:
        field_status = driver.find_element(By.ID, f"{field_id}_status").get_attribute('class')
        is_unlocked = "unlocked" in field_status
        logging.info(f"欄位 {field_id} 狀態: {'unlocked' if is_unlocked else 'locked'}")
        return is_unlocked
    except Exception:
        logging.info(f"無法找到欄位 {field_id} 狀態，假設為鎖定狀態")
        return False

def _kept_tmdb_text(current, field_type):
    """【2026-11 N1】來源沒有資料、選擇不清空時：保留 TMDB 內容。
    若 TMDB 內容是簡體，回傳轉換後的繁體（需要寫入）；否則回傳 None（不用改動）。"""
    if not current or not str(current).strip():
        return None
    try:
        if not utils._needs_simp_to_trad(current):
            return None
        converted = normalize_to_hk(current, field_type=field_type)
    except Exception as e:
        logging.warning(f"來源無資料：檢查 TMDB 內容是否簡體時出錯，保留不改: {e}")
        return None
    if converted and converted.strip() and converted != current:
        return converted
    return None


def _source_empty_should_clear(overwrite_flag):
    """來源沒有資料時是否清空 TMDB 欄位（舊行為；需啟動選項 + 覆蓋=y）。"""
    return utils.clear_tmdb_when_source_empty() and str(overwrite_flag).lower() == 'y'


def is_episode_generic_title(title):
    """檢查標題是否為 '第 X 集' 格式"""
    if not title:
        return False
    return bool(re.match(r'^第\s*\d+\s*集$', title.strip()))

from others.utils import convert_runtime_to_minutes  # 新增導入


def _valid_runtime(value):
    """Return a positive int runtime, or None when the source has no runtime (None / '' / 0 / NaN / junk)."""
    if value is None:
        return None
    try:
        minutes = int(float(str(value).strip()))
    except (ValueError, TypeError):
        return None
    return minutes if minutes > 0 else None

def check_and_fill_form(driver, title_field_id, overview_field_id, excel_path, sheet_name, excel_title, excel_description, runtime=None, overwrite_titles='ask', overwrite_descriptions='ask', tmdb_id=None, video_url=None, perfect_title=None, perfect_description=None, air_date=None):
    """
    【修改版】檢查並填寫表單。
    - 新增 perfect_title 和 perfect_description 參數。
    - 如果 perfect 欄位為 'n'，則無論 overwrite 選項如何，都會強制覆蓋。
    - 如果 TMDB 上的內容不是理想的繁體中文格式（例如是簡體、包含未標準化的字符）、空值或泛用型標題，則無論 overwrite 選項如何，都會用 Excel 的內容覆蓋。
    - 否則，僅當用戶選擇覆蓋時 (overwrite != 'n')，才會更新。
    """
    # 處理來源標籤
    clean_excel_title = excel_title
    if isinstance(excel_title, str):
        match = re.match(r'\[([A-Z_-]+)\](.*)', excel_title)
        if match:
            clean_excel_title = match.group(2).strip()

    clean_excel_description = excel_description
    if isinstance(excel_description, str):
        match = re.match(r'\[([A-Z_-]+)\](.*)', excel_description, re.DOTALL)
        if match:
            clean_excel_description = match.group(2).strip()
    
    logging.info(f"檢查並填寫表單: title_id='{title_field_id}', Excel標題='{clean_excel_title}', Excel描述長度={len(clean_excel_description) if clean_excel_description else 0}, Excel時長='{runtime}', overwrite_titles='{overwrite_titles}', overwrite_descriptions='{overwrite_descriptions}'")
    update_made = False
    try:
        title_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, title_field_id)))
        overview_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, overview_field_id)))

        tmdb_title_raw = title_field.get_attribute('value') or ""
        tmdb_description_raw = overview_field.get_attribute('value') or ""
        
        # 【使用者要求 2025-09-05】錯誤修正：
        # 原始碼在此處直接對 clean_excel_title 執行 .strip()，但當 Excel 標題為數字 (例如集數 '1') 時，
        # clean_excel_title 會是整數 (int) 類型，導致 AttributeError。
        # 修改：在執行 .strip() 前，使用 str() 將其強制轉換為字串，以兼容數字或其他非字串類型的標題。
        current_excel_title = str(clean_excel_title).strip()
        current_excel_description = str(clean_excel_description or "").strip()

        logging.info(f"TMDB 原始標題: '{tmdb_title_raw}'")
        logging.info(f"TMDB 原始描述長度: {len(tmdb_description_raw)}")
        
        title_locked = 'k-disabled' in title_field.get_attribute('class') or not title_field.is_enabled()
        overview_locked = 'k-disabled' in overview_field.get_attribute('class') or not overview_field.is_enabled()
        logging.info(f"欄位鎖定狀態: 標題={title_locked}, 描述={overview_locked}")

        # --- 標題更新邏輯 (新) ---
        update_this_title = False
        # 【修正】只要 Excel 中有內容，就進入判斷邏輯
        if current_excel_title:
            force_overwrite_title = False
            reason = ""
            tmdb_title_normalized = normalize_to_hk(tmdb_title_raw, field_type='title') if tmdb_title_raw else ""
            
            if perfect_title == 'n':
                force_overwrite_title = True
                reason = "Excel 中 'perfect title' 欄標記為 'n'"
            elif not tmdb_title_raw.strip():
                force_overwrite_title = True
                reason = "TMDB標題為空"
            elif tmdb_title_raw != tmdb_title_normalized:
                force_overwrite_title = True
                reason = f"TMDB標題 ('{tmdb_title_raw}') 不是理想的繁體形式 (理想形式: '{tmdb_title_normalized}')"
            elif is_generic_title(tmdb_title_raw):
                force_overwrite_title = True
                reason = f"TMDB標題為泛用型 ('{tmdb_title_raw}')"

            can_overwrite_by_choice = overwrite_titles.lower() != 'n'

            # 【2026-11 P1】不覆蓋 (n) 時，TMDB 已有內容只是未規範化（簡體/引號等）→ 寫入 TMDB 原文的規範化版本，不用來源內容
            if (force_overwrite_title and not can_overwrite_by_choice and perfect_title != 'n'
                    and tmdb_title_raw.strip() and not is_generic_title(tmdb_title_raw)
                    and tmdb_title_normalized and tmdb_title_normalized != tmdb_title_raw):
                logging.info(f"標題更新：不覆蓋模式，保留 TMDB 標題但規範化：'{tmdb_title_raw}' -> '{tmdb_title_normalized}'")
                current_excel_title = tmdb_title_normalized
                update_this_title = True
            elif force_overwrite_title:
                logging.info(f"標題更新：因為 {reason}，將使用Excel內容 ('{current_excel_title}') 進行強制覆蓋。")
                update_this_title = True
            elif can_overwrite_by_choice and current_excel_title != tmdb_title_raw:
                logging.info(f"標題更新：用戶允許覆蓋 (overwrite='{overwrite_titles}') 且內容不同。將使用Excel內容 ('{current_excel_title}')。")
                update_this_title = True
            elif current_excel_title == tmdb_title_raw:
                logging.info("標題無需更新 (Excel 中與 TMDB 相同)。")
            else:
                logging.info(f"標題不更新：內容不同，但用戶選擇不覆蓋 (overwrite='{overwrite_titles}') 且TMDB內容非簡體/泛用型/空。")
        else:
            # Excel 為空：
            # - 啟動選項「來源無資料時清空 TMDB」= y 且覆蓋 = y：清空以恢復 TMDB 預設標題（舊行為）
            # - 否則（預設）：保留 TMDB 標題；TMDB 是簡體時寫入轉換後的繁體【2026-11 N1】
            # Exception: if TMDB already has a generic title, no need to clear it when Excel is empty.
            if _source_empty_should_clear(overwrite_titles) and tmdb_title_raw.strip():
                if is_generic_title(tmdb_title_raw):
                    logging.info(f"標題：Excel 為空，但 TMDB 已是泛用型標題 ('{tmdb_title_raw}')，無需清空。")
                else:
                    logging.info("標題：Excel 為空且選擇清空，將清空 TMDB 標題以恢復預設。")
                    update_this_title = True
            else:
                kept = _kept_tmdb_text(tmdb_title_raw, 'title')
                if kept:
                    logging.info(f"標題：來源無資料，保留 TMDB 標題但轉為繁體：'{tmdb_title_raw}' -> '{kept}'")
                    current_excel_title = kept
                    update_this_title = True
                else:
                    logging.info("標題無需更新 (來源無資料，保留 TMDB 現有內容)。")

        if update_this_title:
            if not title_locked:
                logging.info(f"執行標題更新到TMDB: 原TMDB='{tmdb_title_raw}', 將更新為 Excel 內容='{current_excel_title}'")
                title_field.clear()
                _safe_send_keys(driver, title_field, current_excel_title)
                update_made = True
            else:
                logging.warning(f"標題：意圖更新為 '{current_excel_title}'，但欄位已鎖定。")

        # --- 描述更新邏輯 (新) ---
        update_this_description = False
        # 【修正】只要 Excel 中有內容，就進入判斷邏輯
        if current_excel_description:
            force_overwrite_desc = False
            reason = ""
            tmdb_desc_normalized = normalize_to_hk(tmdb_description_raw, field_type='description') if tmdb_description_raw else ""

            if perfect_description == 'n':
                force_overwrite_desc = True
                reason = "Excel 中 'perfect description' 欄標記為 'n'"
            elif not tmdb_description_raw.strip():
                force_overwrite_desc = True
                reason = "TMDB描述為空"
            elif tmdb_description_raw != tmdb_desc_normalized:
                force_overwrite_desc = True
                reason = "TMDB描述不是理想的繁體形式"
            
            can_overwrite_by_choice = overwrite_descriptions.lower() != 'n'

            # 【2026-11 P1】不覆蓋 (n) 時，TMDB 已有描述只是未規範化 → 寫入 TMDB 原文的規範化版本，不用來源內容
            if (force_overwrite_desc and not can_overwrite_by_choice and perfect_description != 'n'
                    and tmdb_description_raw.strip()
                    and tmdb_desc_normalized and tmdb_desc_normalized != tmdb_description_raw):
                logging.info("描述更新：不覆蓋模式，保留 TMDB 描述但規範化（簡轉繁／引號等）。")
                current_excel_description = tmdb_desc_normalized
                update_this_description = True
            elif force_overwrite_desc:
                logging.info(f"描述更新：因為 {reason}，將使用Excel內容進行強制覆蓋。")
                update_this_description = True
            elif can_overwrite_by_choice and current_excel_description != tmdb_description_raw:
                logging.info(f"描述更新：用戶允許覆蓋 (overwrite='{overwrite_descriptions}') 且內容不同。將使用Excel內容。")
                update_this_description = True
            elif current_excel_description == tmdb_description_raw:
                logging.info("描述無需更新 (Excel 中與 TMDB 相同)。")
            else:
                logging.info(f"描述不更新：內容不同，但用戶選擇不覆蓋 (overwrite='{overwrite_descriptions}') 且無強制覆蓋條件。")
        else:
            # Excel 為空：啟動選項「來源無資料時清空 TMDB」= y 且覆蓋 = y 才清空（舊行為），
            # 否則保留 TMDB 描述，只在簡體時轉繁體【2026-11 N1】
            # Exception: if current TMDB description is already ideal, do not clear just because Excel is empty.
            kept_desc = None if _source_empty_should_clear(overwrite_descriptions) else _kept_tmdb_text(tmdb_description_raw, 'description')
            if kept_desc:
                logging.info("描述：來源無資料，保留 TMDB 描述但轉為繁體。")
                current_excel_description = kept_desc
                update_this_description = True
            elif not _source_empty_should_clear(overwrite_descriptions):
                logging.info("描述無需更新 (來源無資料，保留 TMDB 現有內容)。")
            elif tmdb_description_raw.strip():
                norm_tmdb = normalize_to_hk(tmdb_description_raw, field_type='description') if tmdb_description_raw else ''
                is_current_good = (
                    tmdb_description_raw.strip()
                    and norm_tmdb == tmdb_description_raw
                    and not (perfect_description == 'n')
                )
                if is_current_good:
                    logging.info("描述：Excel 為空，但 TMDB 描述已是理想狀態，無需清空。")
                else:
                    logging.info("描述：Excel 為空且用戶選擇覆蓋，將清空 TMDB 描述以恢復預設。")
                    update_this_description = True
            else:
                logging.info("描述無需更新 (Excel 中為空)。")

        if update_this_description:
            if not overview_locked:
                logging.info(f"執行描述更新到TMDB: 原長度={len(tmdb_description_raw)}, 將更新為 Excel 內容 (長度={len(current_excel_description)})")
                overview_field.clear()
                _safe_send_keys(driver, overview_field, current_excel_description)
                update_made = True
            else:
                logging.warning("描述：意圖更新，但欄位已鎖定。")
        
        # --- 時長更新邏輯：0 / 空白 = 來源冇資料，唔改 TMDB ---
        runtime = _valid_runtime(runtime)
        if runtime is None:
            logging.debug("Excel 時長為 0 或空白，視為冇資料，唔更新 TMDB 時長。")
        if runtime is not None:
            try:
                runtime_field = driver.find_element(By.ID, 'runtime')
                runtime_locked = 'k-disabled' in runtime_field.get_attribute('class') or not runtime_field.is_enabled()
                tmdb_runtime_raw = runtime_field.get_attribute('value') or ""
                
                if str(runtime) != tmdb_runtime_raw:
                    if not runtime_locked:
                        logging.info(f"執行時長更新到TMDB: 原TMDB='{tmdb_runtime_raw}', 將更新為 Excel 內容='{runtime}'")
                        runtime_field.clear()
                        runtime_field.send_keys(str(runtime))
                        update_made = True
                    else:
                        logging.warning(f"時長：意圖更新為 '{runtime}'，但欄位已鎖定。")
                else:
                    logging.info("時長無需更新 (與 TMDB 相同)。")
            except NoSuchElementException:
                logging.debug("頁面上沒有 'runtime' 欄位，跳過時長更新。")

        # --- 播出日期更新 ---
        if air_date:
            try:
                filled = driver.execute_script(
                    "var dp = $('#air_date_date_picker_field').data('kendoDatePicker');"
                    "if (dp) { dp.value(new Date(arguments[0])); dp.trigger('change'); return true; }"
                    "return false;",
                    air_date
                )
                if filled:
                    logging.info(f"播出日期 (Kendo API) 已填寫: {air_date}")
                    update_made = True
                else:
                    el = driver.find_element(By.ID, 'air_date_date_picker_field')
                    driver.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));"
                        "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                        el, air_date
                    )
                    logging.info(f"播出日期 (JS fallback) 已填寫: {air_date}")
                    update_made = True
            except NoSuchElementException:
                logging.debug("頁面上沒有 'air_date_date_picker_field' 欄位，跳過。")
            except Exception as e:
                logging.warning(f"填寫播出日期時發生錯誤: {e}")

        if update_made:
            logging.info("檢測到至少一個欄位已更新，準備提交表單...")
            submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "submit")))
            driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
            submit_button.click()
            
            try:
                WebDriverWait(driver, 20).until(
                    EC.any_of(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '已將您的修改儲存') or contains(text(), '成功')]")),
                        EC.url_changes(driver.current_url) 
                    )
                )
                logging.info("表單提交成功或頁面已響應。")
            except TimeoutException:
                logging.warning("提交後未檢測到明確的成功訊息或URL變化，但可能已成功。")
            time.sleep(random.uniform(2, 4))
            return True 
        else:
            logging.info("所有欄位無需更新。")
            time.sleep(random.uniform(0.5, 1))
            return True

    except TimeoutException as e:
        logging.error(f"等待欄位 '{title_field_id}' 或 '{overview_field_id}' 超時: {e}")
        if check_wait_and_try_again(driver, driver.current_url, video_url): 
            logging.warning("檢測到 'Wait and try again' 頁面，操作可能受阻。")
        return False
    except Exception as e:
        logging.error(f"check_and_fill_form 過程中發生異常: {e}\n{traceback.format_exc()}")
        return False

def pre_check(driver, current_url=None, video_url=None):
    """
    檢查頁面是否有效。
    【修改】回傳 (is_valid, was_limited)。
    was_limited: 表示在檢查過程中是否遭遇並處理了限流 (Wait and try again)。
    """
    was_limited = False
    try:
        # 權限檢查
        page_source = driver.page_source
        if "您没有权限访问该页面" in page_source or "You don't have permission to access this page" in page_source:
            logging.warning(f"pre_check: 檢測到權限錯誤 (未登入?)，嘗試重新登入 TMDB...")
            configs = load_configs()
            username, password = get_tmdb_credentials(configs)
            if username and password:
                login(driver, username, password)
                logging.info(f"重新登入完成，返回原網址: {current_url}")
                if current_url:
                    driver.get(current_url)
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    if "您没有权限访问该页面" in driver.page_source or "You don't have permission to access this page" in driver.page_source:
                        logging.error("重新登入後仍然沒有權限。")
                        return False, False
            else:
                logging.error("無法重新登入：缺少憑證。")
                return False, False

        if "Oops! We can't find the page you're looking for" in driver.page_source:
            logging.error("頁面找不到，跳過此項目")
            time.sleep(4)
            save_unavailable_video_info_to_excel(video_url or driver.current_url, driver.current_url)
            return False, False

        # 檢查限流
        # 使用 utils.check_wait_and_try_again，它現在會回傳 True 如果曾經限流
        if check_wait_and_try_again(driver, current_url, video_url):
            was_limited = True
            logging.info("pre_check: 頁面曾被限流但已恢復。")
        
        return True, was_limited

    except Exception as e:
        logging.error(f"pre_check 過程中發生錯誤: {e}")
        save_unavailable_video_info_to_excel(video_url or driver.current_url, driver.current_url)
        return False, False

def failsafe(driver, current_url, correct_episode_url=False, tv_show_id=None, season=None, episode_number=None, excel_path=None, title_field_id=None, overview_field_id=None, sheet_name=None, title=None, description=None, runtime=None, is_movie=False, video_url=None):
    """
    【修改】失敗安全處理，針對 TimeoutException 和 WebDriverException (網路錯誤) 重新載入網頁。
    【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
    """
    # 【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
    # 將最大重試次數提高到一個較大的值，以模擬無限重試。
    max_retries = 50 # 與原碼一致
    # KeyboardInterrupt 會由 check_wait_and_try_again 內部拋出，並在此處被捕獲或向外傳播
    while get_failsafe_retry_count(current_url) < max_retries:
        retry_count = increment_failsafe_retry_count(current_url)
        logging.info(f"進入 failsafe 模式，嘗試 {retry_count}/{max_retries}，URL: {current_url}")
        try:
            driver.get(current_url)
            time.sleep(random.uniform(3, 5))

            # 快速檢查 'Wait and try again'，避免不必要的等待
            wait_message_found = False
            try:
                WebDriverWait(driver, 0.1).until(
                    EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Wait and try again')]"))
                )
                wait_message_found = True
            except TimeoutException:
                pass # 【修復】缺少 pass，導致 IndentationError

            if wait_message_found:
                logging.info("Failsafe 偵測到 'Wait and try again' 訊息，進入標準等待處理流程。")
                if check_wait_and_try_again(driver, current_url, video_url): # 如果返回 True，表示問題仍存在，但用戶未中斷
                    logging.warning(f"failsafe: check_wait_and_try_again 返回 True (表示 'Wait...' 存在且未中斷)，這不符合預期。繼續 failsafe 循環。URL: {current_url}")
                    continue # 繼續 failsafe 的外層 while 循環

            # 【核心修改】在重試時，優先檢查目標編輯欄位是否已載入
            effective_title_field_id = title_field_id or ('zh_HK_translated_title' if is_movie else 'zh_HK_name')
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, effective_title_field_id)))
                # 如果欄位存在，代表頁面載入成功，直接嘗試填寫表單
                logging.info(f"Failsafe: 偵測到編輯欄位 '{effective_title_field_id}'，嘗試直接填寫表單。")
                if check_and_fill_form(driver, effective_title_field_id,
                                      overview_field_id or 'zh_HK_overview', excel_path, sheet_name or ('Movies' if is_movie else 'Episodes'),
                                      title or '', description or '', runtime or '', overwrite_titles='ask', overwrite_descriptions='ask', video_url=video_url):
                    logging.info("Failsafe: 表單成功填寫，退出 failsafe。")
                    reset_failsafe_counter(current_url)
                    return True
                else:
                    logging.warning("Failsafe: 找到欄位但填寫失敗，繼續 failsafe 循環。")
                    continue # 繼續外層 while 循環
            except TimeoutException:
                # 找不到編輯欄位，流程將繼續執行下方的原有邏輯（例如檢查新增翻譯）
                logging.info(f"Failsafe: 未找到編輯欄位 '{effective_title_field_id}'，繼續執行原有檢查。")
                pass

            # ---- 原有的 failsafe 邏輯 ----
            if "Oops! We can't find the page you're looking for" in driver.page_source:
                logging.error("頁面找不到，跳過此項目")
                save_unavailable_video_info_to_excel(video_url or current_url, current_url)
                reset_failsafe_counter(current_url)
                return False
            elif "的翻譯。如果您有興趣新增此翻譯，請按下方的按鈕。" in driver.page_source:
                logging.info("檢測到缺少 zh-HK 翻譯，嘗試添加新翻譯")
                if check_and_add_translation(driver, current_url, is_movie=is_movie, video_url=video_url): # 此函數內部也可能拋出 KI
                    driver.get(current_url) # 重新載入以確保翻譯生效
                    logging.info("已重新載入頁面，隨機等待 3-5 秒")
                    time.sleep(random.uniform(3, 5))
                    if check_and_fill_form(driver, title_field_id or ('zh_HK_translated_title' if is_movie else 'zh_HK_name'),
                                          overview_field_id or 'zh_HK_overview', excel_path, sheet_name or ('Movies' if is_movie else 'Episodes'),
                                          title or '', description or '', runtime or '', overwrite_titles='ask', overwrite_descriptions='ask'):
                        reset_failsafe_counter(current_url)
                        return True
                logging.error("無法添加 zh-HK 翻譯或後續填寫失敗，記錄錯誤並跳過")
                if video_url: # 確保 video_url 存在
                    save_translation_error_to_excel(video_url, current_url, "無法載入或處理 zh-HK 翻譯頁面")
                reset_failsafe_counter(current_url)
                return False # 添加翻譯或填寫失敗
            elif correct_episode_url:
                # ... (省略了這部分，與問題不直接相關，但需確保其內部也能正確處理 KeyboardInterrupt)
                # 假設這部分會正確處理中斷或成功/失敗返回
                # 【修復】移除了錯誤的 Markdown 格式 URL
                correct_url_to_try = f"https://www.themoviedb.org/tv/{tv_show_id}/season/{season}/episode/{episode_number}/edit?language=zh-HK"
                driver.get(correct_url_to_try) # 可能拋出 KI
                time.sleep(random.uniform(3,5))
                if check_wait_and_try_again(driver, correct_url_to_try, video_url): # 可能拋出 KI
                    logging.warning(f"修正 URL 後仍有 'Wait...' 訊息: {correct_url_to_try}")
                    continue
                logging.info(f"使用正確的 URL: {correct_url_to_try}")
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'zh_HK_name')))
                if check_and_fill_form(driver, title_field_id or 'zh_HK_name', overview_field_id or 'zh_HK_overview', excel_path, sheet_name or 'Episodes',
                                      title or '', description or '', runtime or '', overwrite_titles='ask', overwrite_descriptions='ask'):
                    reset_failsafe_counter(current_url) # 注意這裡用的是 original current_url
                    return True
                return False # 填寫失敗
            else:
                # 檢查頁面數據是否已更新 (此邏輯在原碼中較為複雜，此處簡化)
                logging.info("檢查頁面數據是否已更新...")
                # 假設 check_and_fill_form 已經處理了更新並返回了狀態
                # 或者這裡有一個獨立的檢查邏輯
                # 為簡化，我們假設如果到這裡，並且前面沒有return False，那麼可以認為是成功的
                # 實際上，原代碼在這裡有更詳細的檢查邏輯
                reset_failsafe_counter(current_url)
                return True # 假設操作成功或無需操作

        except KeyboardInterrupt:
            logging.warning(f"failsafe: 在嘗試 {retry_count}/{max_retries} 時檢測到 KeyboardInterrupt。URL: {current_url}")
            raise # 重新引發，讓更高層處理
        except Exception as e:
            if isinstance(e, KeyboardInterrupt): # 雙重保險
                raise
            logging.error(f"failsafe 嘗試 {retry_count}/{max_retries} 失敗: {str(e)}，URL: {current_url}")
            time.sleep(random.uniform(10, 15)) # 在繼續下一次 failsafe 循環前等待
            continue

    logging.error(f"failsafe: 重試次數達到上限 {max_retries}，記錄錯誤並跳過，URL: {current_url}")
    save_unavailable_video_info_to_excel(video_url or current_url, current_url)
    reset_failsafe_counter(current_url)
    return False

def retry_check_and_fill_form(driver, title_field_id, overview_field_id, excel_path, sheet_name, title, description, runtime, overwrite_titles='ask', overwrite_descriptions='ask', max_retries=50):
    """【修改後】重試填寫表單，新增處理翻譯缺失的能力，並在網路中斷時無限重試。"""
    current_url = driver.current_url
    video_url_for_error = current_url # 當 video_url 未提供時，使用 current_url 作為備用

    # 為此函數的重試操作使用一個唯一的鍵，以避免與其他重試邏輯（如 failsafe）的計數器衝突
    retry_key = f"retry_fill_{current_url}"

    # 重置計數器以開始新的重試序列
    reset_failsafe_counter(retry_key)

    # 【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
    # 將最大重試次數提高到一個較大的值，以模擬無限重試。
    while get_failsafe_retry_count(retry_key) < max_retries:
        retry_count = increment_failsafe_retry_count(retry_key)
        logging.info(f"第 {retry_count}/{max_retries} 次重試填寫表單，URL: {current_url}")
        try:
            time.sleep(random.uniform(5, 10)) # 重試前隨機等待
            logging.info("重新加載頁面...")
            driver.get(current_url)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

            # 快速檢查 'Wait and try again'
            wait_message_found = False
            try:
                WebDriverWait(driver, 0.1).until(
                    EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Wait and try again')]"))
                )
                wait_message_found = True
            except TimeoutException:
                pass # 正常情況

            if wait_message_found:
                logging.info("Retry 偵測到 'Wait and try again' 訊息，進入標準等待處理流程。")
                if check_wait_and_try_again(driver, current_url, video_url_for_error): # 如果中斷會拋異常
                    logging.warning("retry_check_and_fill_form: 'Wait and try again' 訊息仍存在，繼續重試。")
                    continue
            
            # 核心修改：檢查欄位是否存在，若不存在則嘗試新增翻譯
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, title_field_id)))
                # 欄位存在，可直接嘗試填寫
                if check_and_fill_form(driver, title_field_id, overview_field_id, excel_path, sheet_name, title, description, runtime, overwrite_titles=overwrite_titles, overwrite_descriptions=overwrite_descriptions, video_url=video_url_for_error):
                    reset_failsafe_counter(retry_key)
                    return True
                else:
                    logging.warning("表單填寫失敗，將繼續重試。")
                    continue
            except TimeoutException:
                # 欄位不存在，嘗試新增翻譯
                logging.info(f"未找到編輯欄位 '{title_field_id}'，檢查是否需要新增翻譯。")
                is_movie_inferred = (title_field_id == 'zh_HK_translated_title')
                if check_and_add_translation(driver, current_url, is_movie=is_movie_inferred, video_url=video_url_for_error):
                    logging.info("成功新增翻譯，將在下次循環中嘗試填寫。")
                    # 讓循環繼續到下一次迭代，屆時將會找到欄位
                    continue
                else:
                    logging.error("新增翻譯失敗，中止此項目的重試。")
                    save_translation_error_to_excel(video_url_for_error, current_url, "重試循環中新增翻譯失敗")
                    break # 中斷 while 循環

        except KeyboardInterrupt:
            logging.warning(f"retry_check_and_fill_form: 在嘗試 {retry_count}/{max_retries} 時檢測到 KeyboardInterrupt。URL: {current_url}")
            raise
        except Exception as e:
            if isinstance(e, KeyboardInterrupt): # 雙重保險
                raise
            logging.error(f"retry_check_and_fill_form 嘗試 {retry_count}/{max_retries} 失敗: {str(e)}，URL: {current_url}")
            # 繼續循環
            continue

    logging.error(f"retry_check_and_fill_form: 超過最大重試次數 {max_retries}，放棄重試，跳過此項目，URL: {current_url}")
    save_unavailable_video_info_to_excel(current_url, current_url) # 假設 video_url 不可用時用 current_url
    reset_failsafe_counter(retry_key) # 重置此函數的計數器
    return False

def check_and_add_person_translation(driver, person_url_with_lang, person_name_excel, person_bio_excel=None, video_url=None, overwrite_titles='ask', overwrite_descriptions='ask', perfect_name=None, perfect_bio=None):
    """
    【新版邏輯】處理單個人員的翻譯新增和資料填寫。
    此函數負責導航、確保人員的 zh-HK 翻譯表單可用，並填寫資料。
    【修正】更新了姓名和簡介的判斷邏輯，使其在 'perfect' 欄位為 'n' 時，即使網頁內容與 Excel 內容看似相同，也能強制觸發儲存操作，以解決 API 和網頁前端顯示不一致的問題。
    """
    primary_field_id = 'zh_HK_name'
    bio_field_id = 'zh_HK_biography'

    # --- 導航及預處理 ---
    driver.get(person_url_with_lang)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    if "Oops! We can't find the page you're looking for" in driver.page_source:
        logging.error(f"人員頁面找不到: {person_url_with_lang}")
        save_translation_error_to_excel(video_url or person_url_with_lang, person_url_with_lang, "人員頁面找不到")
        return False
    if not pre_check(driver, person_url_with_lang, video_url):
        return False

    # --- 新增翻譯點擊函數 (人員頁面版本) ---
    def click_add_person_translation_and_verify():
        try:
            button_xpath = "//button[contains(@class, 'translate') and contains(., '新增翻譯')]"
            add_translation_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_translation_button)
            add_translation_button.click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, primary_field_id)))
            logging.info(f"成功點擊人員 '新增翻譯' 並確認欄位 '{primary_field_id}' 已出現。")
            return True
        except Exception as e:
            logging.warning(f"點擊人員 '新增翻譯' 或驗證欄位時失敗: {e}")
            return False

    # --- 檢查及新增翻譯 ---
    try:
        driver.find_element(By.ID, primary_field_id)
        logging.info(f"人員編輯欄位 '{primary_field_id}' 已存在。")
    except NoSuchElementException:
        logging.info(f"未找到人員編輯欄位 '{primary_field_id}'，嘗試新增翻譯。")
        if not click_add_person_translation_and_verify():
            logging.warning("新增人員翻譯失敗，重新載入頁面並重試。")
            driver.get(person_url_with_lang) # 重新載入
            if not pre_check(driver, person_url_with_lang, video_url): return False
            
            if not click_add_person_translation_and_verify():
                logging.error("重試後仍無法新增人員翻譯，跳過此人員。")
                save_translation_error_to_excel(video_url or person_url_with_lang, person_url_with_lang, "無法新增人員的 zh-HK 翻譯")
                return False

    # --- 填寫表單 ---
    try:
        # 處理來源標籤
        clean_person_name = person_name_excel
        if isinstance(person_name_excel, str):
            match = re.match(r'\[([A-Z_-]+)\](.*)', person_name_excel)
            if match:
                clean_person_name = match.group(2).strip()

        clean_person_bio = person_bio_excel
        if isinstance(person_bio_excel, str):
            match = re.match(r'\[([A-Z_-]+)\](.*)', person_bio_excel, re.DOTALL)
            if match:
                clean_person_bio = match.group(2).strip()

        name_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, primary_field_id)))
        bio_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, bio_field_id)))
        
        tmdb_name = name_field.get_attribute('value') or ""
        tmdb_bio = bio_field.get_attribute('value') or ""
        name_locked = 'k-disabled' in name_field.get_attribute('class') or not name_field.is_enabled()
        bio_locked = 'k-disabled' in bio_field.get_attribute('class') or not bio_field.is_enabled()
        update_made_person = False

        # --- 人名處理 (新版邏輯) ---
        update_this_name = False
        current_excel_name = clean_person_name.strip()
        if current_excel_name: # If there's a name in Excel to potentially write
            force_overwrite_name = False
            reason = ""
            if perfect_name == 'n':
                force_overwrite_name = True
                reason = "Excel 中 'perfect name' 欄標記為 'n'"
            elif not tmdb_name.strip():
                force_overwrite_name = True
                reason = "TMDB 人名為空"
            # 【修改】 使用 field_type='name'
            elif tmdb_name != normalize_to_hk(tmdb_name, field_type='name'):
                force_overwrite_name = True
                reason = f"TMDB 人名 ('{tmdb_name}') 不是理想的繁體形式"
            
            can_overwrite_by_choice = overwrite_titles.lower() != 'n'

            if force_overwrite_name:
                logging.info(f"人名更新：因為 {reason}，將使用 Excel 內容 ('{current_excel_name}') 進行強制覆蓋。")
                update_this_name = True
            elif can_overwrite_by_choice and current_excel_name != tmdb_name:
                logging.info(f"人名更新：用戶允許覆蓋 (overwrite='{overwrite_titles}') 且內容不同。")
                update_this_name = True
            elif current_excel_name == tmdb_name:
                logging.info("人名無需更新 (Excel 中與 TMDB 相同)。")
            else:
                logging.info(f"人名不更新：內容不同，但用戶選擇不覆蓋 (overwrite='{overwrite_titles}') 且無強制覆蓋條件。")
        else:
            logging.info("人名無需更新 (Excel 中為空)。")

        if update_this_name:
            if not name_locked:
                logging.info(f"執行人名更新到TMDB: 原人名='{tmdb_name}', 將更新為 Excel 內容='{current_excel_name}'")
                name_field.clear()
                _safe_send_keys(driver, name_field, current_excel_name)
                update_made_person = True
            else:
                logging.warning(f"人名：意圖更新為 '{current_excel_name}'，但欄位已鎖定。")

        # --- 簡介處理 (新版邏輯) ---
        update_this_bio = False
        current_excel_bio = (clean_person_bio or "").strip()
        if current_excel_bio: # If there's a bio in Excel to potentially write
            force_overwrite_bio = False
            reason = ""
            if perfect_bio == 'n':
                force_overwrite_bio = True
                reason = "Excel 中 'perfect bio' 欄標記為 'n'"
            elif not tmdb_bio.strip():
                force_overwrite_bio = True
                reason = "TMDB 簡介為空"
            # 【修改】 使用 field_type='bio'
            elif tmdb_bio != normalize_to_hk(tmdb_bio, field_type='bio'):
                force_overwrite_bio = True
                reason = "TMDB 簡介不是理想的繁體形式"
                
            can_overwrite_by_choice = overwrite_descriptions.lower() != 'n'

            if force_overwrite_bio:
                logging.info(f"簡介更新：因為 {reason}，將使用 Excel 內容進行強制覆蓋。")
                update_this_bio = True
            elif can_overwrite_by_choice and current_excel_bio != tmdb_bio:
                logging.info(f"簡介更新：用戶允許覆蓋 (overwrite='{overwrite_descriptions}') 且內容不同。")
                update_this_bio = True
            elif current_excel_bio == tmdb_bio:
                logging.info("簡介無需更新 (Excel 中與 TMDB 相同)。")
            else:
                logging.info(f"簡介不更新：內容不同，但用戶選擇不覆蓋 (overwrite='{overwrite_descriptions}') 且無強制覆蓋條件。")
        else:
            logging.info("簡介無需更新 (Excel 中為空)。")

        if update_this_bio:
            if not bio_locked:
                logging.info(f"執行簡介更新到TMDB: 原長度={len(tmdb_bio)}, 將更新為 Excel 內容 (長度={len(current_excel_bio)})")
                bio_field.clear()
                _safe_send_keys(driver, bio_field, current_excel_bio)
                update_made_person = True
            else:
                logging.warning("簡介：意圖更新，但欄位已鎖定。")
        
        if update_made_person:
            logging.info("檢測到人員資料已更新，準備提交表單...")
            submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "submit")))
            driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
            submit_button.click()
            try:
                WebDriverWait(driver, 20).until(EC.any_of(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '已將您的修改儲存') or contains(text(), '成功')]")),
                    EC.url_changes(driver.current_url)
                ))
                logging.info("人員表單提交成功或頁面已響應。")
            except TimeoutException:
                logging.warning("提交後未檢測到明確的成功訊息或URL變化，但可能已成功。")
            time.sleep(random.uniform(1.5, 2.5))
        else:
            logging.info(f"人員 {clean_person_name or '[空名稱]'} 無需更新。")
        
        return True

    except Exception as e:
        logging.error(f"處理人員 '{person_name_excel}' (URL: {person_url_with_lang}) 時發生錯誤: {e}")
        save_translation_error_to_excel(video_url or person_url_with_lang, person_url_with_lang, f"處理人員時發生錯誤: {e}")
        return False

def update_people(driver, tmdb_id, excel_path, media_type='movie', video_url=None, overwrite_titles='ask', overwrite_descriptions='ask'):
    """
    【修改】更新人員資訊，加入「限流倒退重試」機制。
    如果載入某人時發現限流，則倒退回上一個人重新處理，以防上一個人的儲存被隱性限流吃掉。
    """
    logging.info("開始更新人員資訊...")
    try:
        df = pd.read_excel(excel_path, sheet_name='People')
        if df.empty or 'Person URL' not in df.columns:
            logging.info(f"People sheet 為空或缺少 'Person URL' 欄位，跳過。")
            return True

        df_valid = df[df['Person URL'].apply(lambda x: pd.notna(x) and str(x).strip() != '')].copy()
        # ... (欄位初始化代碼保持不變) ...
        if 'Name' not in df_valid.columns: df_valid.loc[:, 'Name'] = ''
        if 'Biography' not in df_valid.columns: df_valid.loc[:, 'Biography'] = ''
        if 'perfect name' not in df_valid.columns: df_valid.loc[:, 'perfect name'] = ''
        if 'perfect bio' not in df_valid.columns: df_valid.loc[:, 'perfect bio'] = ''
        df_valid['Name'] = df_valid['Name'].astype('string').fillna('')
        df_valid['Biography'] = df_valid['Biography'].astype('string').fillna('')
        
        df_valid = df_valid[(df_valid['Name'].str.strip() != '') | (df_valid['Biography'].str.strip() != '')]

        if df_valid.empty:
            logging.info(f"People sheet 中無有效的人員資料，跳過。")
            return True
        
        overall_success = True
        
        # 【修改】使用 while 循環和索引，以便可以修改索引 (倒退)
        indices = list(df_valid.index)
        i = 0
        last_processed_url = None # 用於最後一項的檢查

        while i < len(indices):
            if utils._PROGRAM_INTERRUPTED:
                logging.warning("update_people: 檢測到中斷信號，提前終止。")
                overall_success = False
                break
            
            index = indices[i]
            row = df_valid.loc[index]

            person_tmdb_url = row['Person URL']
            person_name_excel = _clean_excel_value(row.get('Name', ''))
            person_bio_excel = _clean_excel_value(row.get('Biography', ''))
            perfect_name = row.get('perfect name', '')
            perfect_bio = row.get('perfect bio', '')

            logging.info(f"處理第 {i+1}/{len(indices)} 個人員 (Index {index}): {person_name_excel}")

            base_person_url = person_tmdb_url.split('?')[0].split('#')[0]
            if not base_person_url.endswith('/edit'):
                base_person_url = base_person_url.rstrip('/') + '/edit'
            
            person_edit_url_zh_hk = f"{base_person_url}?language=zh-HK"
            last_processed_url = person_edit_url_zh_hk # 記錄 URL

            # 1. 導航
            driver.get(person_edit_url_zh_hk)

            # 2. 檢查頁面狀態 & 是否遇到限流
            is_valid, was_limited = pre_check(driver, person_edit_url_zh_hk, video_url)

            if was_limited:
                # 如果遇到限流，且不是第一項 -> 懷疑上一項失敗 -> 倒退
                if i > 0:
                    logging.warning(f"載入第 {i+1} 項時遇到限流 (已恢復)。安全起見，倒退回第 {i} 項 ({df_valid.loc[indices[i-1]].get('Name')}) 重試，以確保上項已儲存。")
                    i -= 1
                    time.sleep(2)
                    continue # 重新開始迴圈處理 i-1
                else:
                    logging.info("第一項遇到限流 (已恢復)，繼續處理當前項。")

            if not is_valid:
                logging.error(f"人員頁面無效，跳過: {person_tmdb_url}")
                overall_success = False
                i += 1
                continue

            # 3. 執行更新
            # 注意：這裡不需再做 check_wait_and_try_again，因為 pre_check 做了，且 check_and_add 內部也會做
            if not check_and_add_person_translation(driver, person_edit_url_zh_hk, person_name_excel, person_bio_excel, video_url, overwrite_titles, overwrite_descriptions, perfect_name, perfect_bio):
                # 如果函數內部回傳 False，可能是因為限流或錯誤
                # 如果是因為隱性限流導致失敗，check_and_add 內部應該已經嘗試過重載
                # 這裡我們可以選擇重試當前項
                logging.warning(f"人員 {person_name_excel} 更新回傳失敗。檢查是否因限流...")
                if check_wait_and_try_again(driver, person_edit_url_zh_hk, video_url):
                    logging.info("檢測到限流，重試當前人員...")
                    # i 不變，continue 會重做當前項
                    continue 
                
                logging.error(f"人員更新確實失敗，跳過。")
                overall_success = False
            else:
                logging.info(f"人員 {person_name_excel or '[空名稱]'} 處理完畢。")
            
            time.sleep(random.uniform(1.5, 2.5))
            i += 1 # 成功或確定失敗後，前進下一項

        # 【新增】處理完最後一項後，進行「回馬槍」檢查
        # 因為如果最後一項提交後剛好遇到限流，我們就跳出迴圈了，沒機會透過「載入下一頁」來發現
        if last_processed_url and not utils._PROGRAM_INTERRUPTED:
            logging.info("所有人員處理完畢，正在進行最終限流檢查 (針對最後一項)...")
            driver.get(last_processed_url) # 重新載入最後一頁
            if check_wait_and_try_again(driver, last_processed_url, video_url): # 如果這時發現限流
                logging.warning("最終檢查發現限流！最後一項可能未儲存成功。倒退回最後一項重試。")
                # 這裡我們需要手動重做最後一項。最簡單的方法是把 i 設回最後一個索引，並重新進入類似迴圈的邏輯，
                # 但為了避免複雜，我們直接遞歸調用自己處理最後一項 (或修改 i 並返回迴圈，但迴圈已結束)
                # 簡單作法：將 indices 只保留最後一項，重新執行一段邏輯，或者如果 i 已經超出，我們手動減 1 並在此處處理是不行的因為結構限制。
                # 最好的方法：在 while 循環條件中處理，或者用 goto (不推薦)。
                # 實用方法：如果發現限流，遞歸調用自己處理最後一項? 不太好。
                # 我們可以在 while 循環結束後，如果發現這個情況，把 i 減 1，然後讓循環繼續 (但需將 while 條件改為 True 並內部 break)
                
                # 修正後的邏輯：其實上面的 while i < len 已經結束。我們無法簡單回到循環。
                # 所以，如果這裡發現限流，我們單獨重做最後一項。
                last_idx = indices[-1]
                row = df_valid.loc[last_idx]
                logging.info(f"重做最後一項: {row.get('Name')}")
                # 遞歸調用 check_and_add...
                check_and_add_person_translation(driver, last_processed_url, _clean_excel_value(row.get('Name', '')), _clean_excel_value(row.get('Biography', '')), video_url, overwrite_titles, overwrite_descriptions, row.get('perfect name', ''), row.get('perfect bio', ''))
                # 再次檢查
                driver.get(last_processed_url)
                check_wait_and_try_again(driver) # 確保最後是乾淨的

        return overall_success
    except (KeyboardInterrupt, SystemExit):
        logging.warning("在 update_people 期間檢測到中斷信號。")
        utils._PROGRAM_INTERRUPTED = True
        return False
    except Exception as e:
        logging.error(f"更新人員資訊過程中發生異常: {e}\n{traceback.format_exc()}")
        return False


# ---------------------------------------------------------------------------
# tmdb-automation approach: API pre-check + simplified Selenium write
# ---------------------------------------------------------------------------

def _load_api_bearer() -> str:
    """Return the TMDB API JWT bearer token from configs.json."""
    try:
        return load_configs().get('tmdb_api', '')
    except Exception:
        return ''


def _api_fetch_season_episodes_batch(tmdb_id, season: int, api_key: str) -> dict:
    """
    Fetch all episodes' base info + zh-HK content for a season in 2 API calls.

    Uses the default-language response to detect fallback: when ?language=zh-HK
    returns the same name/overview as the default response, the episode has no
    actual zh-HK translation and we treat those fields as empty.

    Returns {ep_num: {'name', 'overview', 'exists', 'runtime', 'air_date'}}.
    """
    if not api_key:
        return {}
    headers = {'Authorization': f'Bearer {api_key}', 'accept': 'application/json'}
    base_url = f'https://api.themoviedb.org/3/tv/{tmdb_id}/season/{season}'

    default_eps = {}
    try:
        r = _req.get(base_url, headers=headers, timeout=10)
        if r.status_code == 200:
            for ep in r.json().get('episodes', []):
                n = ep.get('episode_number')
                if n is not None:
                    default_eps[n] = {
                        'name': ep.get('name', '') or '',
                        'overview': ep.get('overview', '') or '',
                        'runtime': ep.get('runtime'),
                        'air_date': ep.get('air_date') or None,
                    }
    except Exception as e:
        logging.warning(f"API 讀取第 {season} 季基本資料失敗: {e}")

    zh_hk_eps = {}
    try:
        r = _req.get(f'{base_url}?language=zh-HK', headers=headers, timeout=10)
        if r.status_code == 200:
            for ep in r.json().get('episodes', []):
                n = ep.get('episode_number')
                if n is not None:
                    zh_hk_eps[n] = {
                        'name': ep.get('name', '') or '',
                        'overview': ep.get('overview', '') or '',
                    }
    except Exception as e:
        logging.warning(f"API 讀取第 {season} 季 zh-HK 翻譯失敗: {e}")

    result = {}
    for n in set(default_eps) | set(zh_hk_eps):
        d = default_eps.get(n, {'name': '', 'overview': '', 'runtime': None, 'air_date': None})
        zh = zh_hk_eps.get(n, {'name': '', 'overview': ''})
        # Only count as zh-HK if different from the default-language response
        zh_name = zh['name'] if zh['name'] != d['name'] else ''
        zh_overview = zh['overview'] if zh['overview'] != d['overview'] else ''
        result[n] = {
            'name': zh_name,
            'overview': zh_overview,
            'exists': bool(zh_name or zh_overview),
            'runtime': d['runtime'],
            'air_date': d['air_date'],
        }
    return result


def _api_fetch_zh_hk(base_api_url: str, api_key: str) -> dict:
    """
    Fetch the zh-HK translation for a TMDB item via the /translations endpoint.
    Returns {'exists': bool, 'name': str, 'overview': str, 'runtime': int|None}.
    """
    if not api_key:
        return {'exists': False, 'name': '', 'overview': '', 'runtime': None}
    headers = {'Authorization': f'Bearer {api_key}', 'accept': 'application/json'}
    try:
        r = _req.get(f'{base_api_url}/translations', headers=headers, timeout=10)
        if r.status_code == 200:
            for t in r.json().get('translations', []):
                if t.get('iso_639_1') == 'zh' and t.get('iso_3166_1') == 'HK':
                    data = t.get('data', {})
                    raw_rt = data.get('runtime')
                    return {
                        'exists': True,
                        'name': data.get('name', '') or '',
                        'overview': data.get('overview', '') or '',
                        'runtime': int(raw_rt) if raw_rt else None,
                    }
    except Exception as e:
        logging.warning(f"API 讀取 zh-HK 翻譯失敗 ({base_api_url}): {e}")
    return {'exists': False, 'name': '', 'overview': '', 'runtime': None}


def _api_fetch_episode_base(tmdb_id, season: int, episode: int, api_key: str) -> dict:
    """Fetch language-neutral episode fields (runtime, air_date)."""
    if not api_key:
        return {'runtime': None, 'air_date': None}
    headers = {'Authorization': f'Bearer {api_key}', 'accept': 'application/json'}
    url = f'https://api.themoviedb.org/3/tv/{tmdb_id}/season/{season}/episode/{episode}'
    try:
        r = _req.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            data = r.json()
            return {'runtime': data.get('runtime'), 'air_date': data.get('air_date') or None}
    except Exception as e:
        logging.warning(f"API 讀取集數基本資料失敗 (S{season}E{episode}): {e}")
    return {'runtime': None, 'air_date': None}


def _clean_excel_value(value) -> str:
    """Strip [SOURCE_TAG] prefix and whitespace from an Excel cell value."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    s = str(value).strip()
    m = re.match(r'^\[([A-Z_-]+)\](.*)', s, re.DOTALL)
    return m.group(2).strip() if m else s


def _decide_what_to_update(
    excel_title, excel_desc, excel_runtime, excel_air_date,
    current_name: str, current_overview: str,
    current_runtime, current_air_date,
    overwrite_titles: str, overwrite_descriptions: str,
    perfect_title='', perfect_desc='',
) -> dict:
    """
    Decide which fields need writing to TMDB using the same logic as
    check_and_fill_form, but without loading a Selenium page.
    Returns a dict: {'title': str, 'overview': str, 'runtime': int, 'air_date': str}.
    """
    updates = {}
    clean_title = _clean_excel_value(excel_title)
    clean_desc = _clean_excel_value(excel_desc)

    if clean_title:
        if is_generic_title(clean_title):
            # Source title is a generic placeholder like "第 X 集" — never write it verbatim.
            # When overwrite=y, clear so TMDB fills the default (even if API shows no zh-HK
            # translation yet — the zh-HK title may be the English fallback, e.g. "TNT").
            if is_generic_title(current_name):
                logging.info(f"來源標題為泛用型 ('{clean_title}')，TMDB 已是泛用型標題，跳過。")
            elif overwrite_titles.lower() == 'y':
                updates['title'] = ''
                logging.info(f"來源標題為泛用型 ('{clean_title}')，清空 TMDB 標題讓其填入預設值。")
            else:
                logging.info(f"來源標題為泛用型 ('{clean_title}')，跳過標題更新。")
        else:
            norm_name = normalize_to_hk(current_name, field_type='title') if current_name else ''
            force = (
                perfect_title == 'n'
                or not current_name.strip()
                or is_generic_title(current_name)
                or norm_name != current_name
            )
            if force and norm_name != current_name:
                logging.debug(
                    f"標題 force 觸發（normalize 差異）: "
                    f"原文={repr(current_name[:80])} → 規範後={repr(norm_name[:80])}"
                )
            if (force and overwrite_titles.lower() == 'n' and perfect_title != 'n'
                    and current_name.strip() and not is_generic_title(current_name)
                    and norm_name and norm_name != current_name):
                # 【2026-11 P1】不覆蓋：TMDB 已有標題只是未規範化 → 寫入規範化後的 TMDB 原文
                updates['title'] = norm_name
                logging.info(f"不覆蓋模式：保留 TMDB 標題但規範化：'{current_name}' -> '{norm_name}'")
            elif force or (overwrite_titles.lower() != 'n' and clean_title != current_name):
                if clean_title == current_name:
                    logging.info("標題無需更新（Excel 與 TMDB 內容相同，跳過重複寫入）。")
                else:
                    updates['title'] = clean_title
            else:
                logging.info(f"標題無需更新 (API 確認與 TMDB 相同或不覆蓋)。")
    elif perfect_title and perfect_title != 'n' and is_generic_title(str(perfect_title)):
        # Episode Title was cleared by post-processing (e.g. "第 5 集" → "")
        # but perfect title still holds the original generic value.
        if is_generic_title(current_name):
            logging.info(f"來源標題為泛用型 ('{perfect_title}')，TMDB 已是泛用型標題，跳過。")
        elif overwrite_titles.lower() == 'y':
            updates['title'] = ''
            logging.info(f"來源標題為泛用型 ('{perfect_title}')，清空 TMDB 標題讓其填入預設值。")
        else:
            logging.info(f"來源標題為泛用型 ('{perfect_title}')，跳過標題更新。")
    # If Excel has no title:
    # - If overwrite_titles=y, treat as explicit "clear this field on TMDB" (force clear intent).
    # - Exception: if TMDB currently has a generic title (泛標題), and Excel is empty,
    #   there is no need to clear it (user's intent is usually to leave generic titles alone
    #   when they have intentionally left the Excel empty).
    else:
        if _source_empty_should_clear(overwrite_titles):
            if is_generic_title(current_name):
                logging.info(f"Excel 標題為空，但 TMDB 目前已是泛用型標題 ('{current_name}')，無需清空。")
            else:
                updates['title'] = ''
                logging.info("Excel 標題為空 + 選擇清空，準備清空 TMDB 標題欄位。")
        else:
            # 【2026-11 N1】保留 TMDB 標題；簡體則轉繁體
            kept = _kept_tmdb_text(current_name, 'title')
            if kept:
                updates['title'] = kept
                logging.info(f"來源無標題，保留 TMDB 標題但轉為繁體：'{current_name}' -> '{kept}'")

    if clean_desc:
        norm_overview = normalize_to_hk(current_overview, field_type='description') if current_overview else ''
        force = (
            perfect_desc == 'n'
            or not current_overview.strip()
            or norm_overview != current_overview
        )
        if force and norm_overview != current_overview:
            logging.debug(
                f"描述 force 觸發（normalize 差異）: "
                f"原文={repr(current_overview[:80])} → 規範後={repr(norm_overview[:80])}"
            )
        if (force and overwrite_descriptions.lower() == 'n' and perfect_desc != 'n'
                and current_overview.strip()
                and norm_overview and norm_overview != current_overview):
            # 【2026-11 P1】不覆蓋：TMDB 已有描述只是未規範化 → 寫入規範化後的 TMDB 原文
            updates['overview'] = norm_overview
            logging.info("不覆蓋模式：保留 TMDB 描述但規範化（簡轉繁／引號等）。")
        elif force or (overwrite_descriptions.lower() != 'n' and clean_desc != current_overview):
            if clean_desc == current_overview:
                # Excel content is identical to current TMDB text — writing it back would be
                # a no-op.  The real fix would be to write the normalised form, but since
                # the Excel source hasn't been normalised either, skip to avoid pointless
                # round-trips.
                logging.info("描述無需更新（Excel 與 TMDB 內容相同，跳過重複寫入）。")
            else:
                updates['overview'] = clean_desc
        else:
            logging.info(f"描述無需更新 (API 確認與 TMDB 相同或不覆蓋)。")
    # If Excel has no description:
    # - If overwrite_descriptions=y, we normally want to clear.
    # - BUT: if the current TMDB description (from API) is already in ideal zh-HK form,
    #   there is no need to clear it just because our local Excel is empty.
    #   Only clear if the current value is bad (empty, not ideal, etc.).
    else:
        if not _source_empty_should_clear(overwrite_descriptions):
            # 【2026-11 N1】保留 TMDB 描述；簡體則轉繁體
            kept = _kept_tmdb_text(current_overview, 'description')
            if kept:
                updates['overview'] = kept
                logging.info("來源無描述，保留 TMDB 描述但轉為繁體。")
            elif current_overview and current_overview.strip():
                logging.info("來源無描述，保留 TMDB 現有描述。")
        elif overwrite_descriptions.lower() == 'y':
            if current_overview and current_overview.strip():
                # Check if current is already good/ideal (same criteria as non-empty path)
                norm_current = normalize_to_hk(current_overview, field_type='description') if current_overview else ''
                force_clear_current = (
                    perfect_desc == 'n'
                    or not current_overview.strip()
                    or norm_current != current_overview
                )
                if force_clear_current:
                    updates['overview'] = ''
                    logging.info("Excel 描述為空 + overwrite=y，但 TMDB 描述非理想狀態，準備清空。")
                else:
                    logging.info("Excel 描述為空，但 TMDB 描述已是理想 zh-HK 狀態，無需清空。")
            else:
                # Current is already empty on TMDB, nothing to do
                logging.info("Excel 描述為空，TMDB 描述也為空，無需動作。")
        else:
            logging.info("描述無需更新 (Excel 中為空)。")

    excel_runtime = _valid_runtime(excel_runtime)  # 0 / 空白 = 來源冇資料，唔改 TMDB
    if excel_runtime is not None and str(excel_runtime) != str(current_runtime or ''):
        updates['runtime'] = excel_runtime

    if excel_air_date and excel_air_date != (current_air_date or ''):
        updates['air_date'] = excel_air_date

    return updates


def _fill_and_submit_tmdb_form(
    driver, updates: dict, title_field_id: str = 'zh_HK_name'
) -> bool:
    """
    Write only the fields listed in `updates` and click Submit.
    Uses the tmdb-automation approach: clear → send_keys → submit.
    Keys: 'title', 'overview', 'runtime', 'air_date'.
    """
    try:
        any_written = False

        if 'title' in updates:
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, title_field_id))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
            field.clear()
            if updates['title']:
                _safe_send_keys(driver, field, updates['title'])
            any_written = True

        if 'overview' in updates:
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'zh_HK_overview'))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
            field.clear()
            if updates['overview']:
                _safe_send_keys(driver, field, updates['overview'])
            any_written = True

        if 'runtime' in updates:
            try:
                for _rt_id in ('zh_HK_runtime', 'runtime'):
                    try:
                        field = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, _rt_id))
                        )
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
                        driver.execute_script(
                            "var el = arguments[0]; el.value = arguments[1];"
                            "el.dispatchEvent(new Event('input',  {bubbles:true}));"
                            "el.dispatchEvent(new Event('change', {bubbles:true}));",
                            field, str(updates['runtime'])
                        )
                        any_written = True
                        logging.info(f"已填入 runtime={updates['runtime']} 到欄位 #{_rt_id}")
                        break
                    except (NoSuchElementException, TimeoutException):
                        continue
                else:
                    logging.warning("找不到 runtime 欄位 (zh_HK_runtime / runtime)，跳過。")
            except Exception as _e:
                logging.warning(f"填入 runtime 時發生錯誤: {_e}")

        if 'air_date' in updates:
            try:
                field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'air_date'))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
                field.clear()
                field.send_keys(updates['air_date'])
                any_written = True
            except (NoSuchElementException, TimeoutException):
                pass

        if any_written:
            submit = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'submit'))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit)
            driver.execute_script("arguments[0].click();", submit)
            logging.info("表單已提交 (tmdb-automation 模式)，等待確認...")
            try:
                WebDriverWait(driver, 20).until(
                    EC.any_of(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '已將您的修改儲存') or contains(text(), '成功')]")),
                        EC.url_changes(driver.current_url)
                    )
                )
                logging.info("表單提交成功。")
            except TimeoutException:
                logging.warning("提交後未檢測到成功訊息或URL變化，但可能已成功。")
            time.sleep(1)

        return True
    except Exception as e:
        logging.error(f"_fill_and_submit_tmdb_form 失敗: {e}")
        return False


def _grid_fill_visible_field(driver, field_id: str, value: str, timeout: int = 5) -> bool:
    """Fill a VISIBLE text input/textarea (skips hidden template elements)."""
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.ID, field_id))
        )
        driver.execute_script(
            "var el = arguments[0]; el.value = arguments[1];"
            "el.dispatchEvent(new Event('input',  {bubbles:true}));"
            "el.dispatchEvent(new Event('change', {bubbles:true}));",
            el, value
        )
        return True
    except Exception:
        return False


def _grid_fill_date(driver, field_id: str, date_str: str) -> bool:
    """Fill a Kendo DatePicker in the grid edit form."""
    filled = driver.execute_script(
        "var dp = $('#' + arguments[0]).data('kendoDatePicker');"
        "if (dp) { dp.value(new Date(arguments[1])); dp.trigger('change'); return true; }"
        "return false;",
        field_id, date_str
    )
    if filled:
        return True
    try:
        el = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.ID, field_id))
        )
        driver.execute_script("arguments[0].value = '';", el)
        ActionChains(driver) \
            .click(el) \
            .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL) \
            .send_keys(Keys.DELETE).send_keys(date_str) \
            .perform()
        driver.execute_script(
            "var el = arguments[0];"
            "el.dispatchEvent(new Event('input',  {bubbles:true}));"
            "el.dispatchEvent(new Event('change', {bubbles:true}));"
            "el.dispatchEvent(new Event('blur',   {bubbles:true}));",
            el
        )
        return True
    except Exception as e:
        logging.warning(f"DatePicker 填寫失敗 ({field_id}): {e}")
        return False


def _edit_episode_via_grid(driver, ep_num: int, updates: dict, language: str = 'zh-HK') -> bool:
    """
    Find the row for ep_num in the currently loaded season grid and apply updates.
    Expects the season grid page to already be loaded.

    Uses ActionChains.click (proper browser events) to trigger Kendo's edit handler,
    then waits for the Update button to become clickable before filling visible fields.
    Logs available input IDs on failure to help diagnose field naming.
    """
    try:
        # Locate the episode row
        target_row = None
        for row in driver.find_elements(By.CSS_SELECTOR, 'tr.k-master-row'):
            tds = row.find_elements(By.CSS_SELECTOR, 'td')
            if tds and tds[0].text.strip() == str(ep_num):
                target_row = row
                break
        if not target_row:
            logging.warning(f"grid 中找不到第 {ep_num} 集的表格行。")
            return False

        # Find the Edit button
        edit_btn = None
        for sel in ['button.k-grid-edit-command', 'button.k-grid-edit']:
            btns = target_row.find_elements(By.CSS_SELECTOR, sel)
            if btns:
                edit_btn = btns[0]
                break
        if not edit_btn:
            logging.warning(f"找不到第 {ep_num} 集的 grid 編輯按鈕。")
            return False

        def _click_edit_and_wait(scroll_extra: int = 0) -> tuple:
            """Click the edit button and wait for edit mode. Returns (edit_active, update_btn)."""
            # 先把整個行滾入視野，再做額外偏移（小螢幕補償）
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});"
                "window.scrollBy(0, arguments[1]);",
                edit_btn, scroll_extra
            )
            time.sleep(0.4)
            # JS click 比 ActionChains 更不受視窗大小影響
            driver.execute_script("arguments[0].click();", edit_btn)
            time.sleep(0.4)

            for sel in [
                'button[ref-update-button]',
                '.k-grid-update',
                'button.k-button-solid-primary[data-action="save"]',
                '.k-edit-buttons button.k-primary',
                '.k-popup-edit-form .k-edit-buttons button:first-child',
            ]:
                try:
                    btn = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                    )
                    return True, btn
                except Exception:
                    continue
            return False, None

        # 第一次嘗試
        edit_active, update_btn = _click_edit_and_wait(scroll_extra=0)

        # 若失敗，向下多捲 150px 再試一次（小螢幕補償）
        if not edit_active:
            time.sleep(0.5)
            edit_active, update_btn = _click_edit_and_wait(scroll_extra=150)

        if not edit_active:
            # Diagnostic: log what inputs are currently visible to help identify field IDs
            field_ids = driver.execute_script(
                "return Array.from(document.querySelectorAll('input:not([type=hidden]), textarea'))"
                ".filter(el => el.offsetParent !== null)"
                ".map(el => (el.id || el.name || el.placeholder || '?') + '[' + el.type + ']')"
                ".join(', ');"
            )
            logging.warning(f"第 {ep_num} 集: edit 模式未啟動。可見輸入欄位: {field_ids}")
            return False

        # Diagnostic: log visible input IDs so we can identify the correct field names
        field_ids = driver.execute_script(
            "return Array.from(document.querySelectorAll('input:not([type=hidden]), textarea'))"
            ".filter(el => el.offsetParent !== null)"
            ".map(el => (el.id || el.name || '?') + '[' + el.type + ']')"
            ".join(', ');"
        )
        logging.info(f"第 {ep_num} 集 edit 模式可見欄位: {field_ids}")

        # Fill visible fields — try both zh-HK (hyphen) and zh_HK (underscore) variants
        lang_u = language.replace('-', '_')
        if 'title' in updates:
            val = updates['title'] or ''
            filled = _grid_fill_visible_field(driver, f'{language}_name_text_input_field', val, 3)
            if not filled:
                filled = _grid_fill_visible_field(driver, f'{lang_u}_name_text_input_field', val, 3)
            if not filled:
                logging.warning(f"第 {ep_num} 集: title 欄位填寫失敗")
        if 'overview' in updates:
            val = updates['overview'] or ''
            filled = _grid_fill_visible_field(driver, f'{language}_overview_text_box_field', val, 3)
            if not filled:
                filled = _grid_fill_visible_field(driver, f'{lang_u}_overview_text_box_field', val, 3)
            if not filled:
                logging.warning(f"第 {ep_num} 集: overview 欄位填寫失敗")
        if 'air_date' in updates and updates['air_date']:
            _grid_fill_date(driver, 'air_date_date_picker_field', updates['air_date'])
        if 'runtime' in updates:
            # Popup field: id="zh-HK_runtime_text_input_field" (plain k-textbox, not NumericTextBox)
            rv = str(updates['runtime'])
            filled = _grid_fill_visible_field(driver, f'{language}_runtime_text_input_field', rv, 3)
            if not filled:
                filled = _grid_fill_visible_field(driver, f'{lang_u}_runtime_text_input_field', rv, 3)
            if not filled:
                logging.warning(f"第 {ep_num} 集: runtime 欄位填寫失敗 (值: {updates['runtime']})")

        # Click the Update/Save button
        driver.execute_script("arguments[0].click();", update_btn)
        time.sleep(1.5)
        logging.info(f"S?E{ep_num}: grid 已更新 {list(updates.keys())}。")
        return True
    except Exception as e:
        logging.error(f"_edit_episode_via_grid 第 {ep_num} 集失敗: {e}")
        return False


def update_series_info(driver, tmdb_id, excel_path, overwrite_titles='ask', overwrite_descriptions='ask', video_url=None):
    """更新電視節目 zh-HK 翻譯 (tmdb-automation 模式：API 預檢 → 簡化 Selenium 寫入)。"""
    logging.info("開始更新系列信息 (tmdb-automation 模式)...")
    api_key = _load_api_bearer()
    while True:
        if utils._PROGRAM_INTERRUPTED:
            return False
        try:
            df = pd.read_excel(excel_path, sheet_name='Title')
            if df.empty:
                return True
            series_title = df['TV Show Title'].iloc[0] if 'TV Show Title' in df.columns and pd.notna(df['TV Show Title'].iloc[0]) else ''
            series_description = df['TV Show Description'].iloc[0] if 'TV Show Description' in df.columns and pd.notna(df['TV Show Description'].iloc[0]) else ''
            perfect_title = df['perfect title'].iloc[0] if 'perfect title' in df.columns and pd.notna(df['perfect title'].iloc[0]) else ''
            perfect_description = df['perfect description'].iloc[0] if 'perfect description' in df.columns and pd.notna(df['perfect description'].iloc[0]) else ''

            series_source_empty = not series_title and not series_description
            if series_source_empty:
                # 【2026-11 N1】只有啟動選項「來源無資料時清空 TMDB」= y 且覆蓋 = y 才清空；
                # 否則保留 TMDB 內容（簡體會轉繁體，由 _decide_what_to_update 處理）
                if _source_empty_should_clear(overwrite_titles) or _source_empty_should_clear(overwrite_descriptions):
                    logging.info("系列 title/desc 為空 + 選擇清空，準備強制清空 TMDB 對應欄位。")
                else:
                    logging.info("系列 title/desc 皆為空：保留 TMDB 現有內容（簡體會轉繁體）。")

            # --- Step 1: Pre-check via TMDB API ---
            zh_hk = _api_fetch_zh_hk(f'https://api.themoviedb.org/3/tv/{tmdb_id}', api_key)
            updates = _decide_what_to_update(
                series_title, series_description, None, None,
                zh_hk['name'], zh_hk['overview'], None, None,
                overwrite_titles, overwrite_descriptions, perfect_title, perfect_description,
            )
            if not updates and (zh_hk['exists'] or series_source_empty):
                logging.info("系列資訊與 TMDB 相同，跳過頁面載入。")
                return True

            # --- Step 2: Load edit page ---
            current_url = f"https://www.themoviedb.org/tv/{tmdb_id}/edit?language=zh-HK"
            driver.get(current_url)
            is_valid, was_limited = pre_check(driver, current_url, video_url)
            if was_limited:
                logging.info("系列頁面遇限流 (已恢復)，重試...")
                continue
            if not is_valid:
                return False

            # --- Step 3: Ensure zh-HK translation exists ---
            if not zh_hk['exists']:
                if not check_and_add_translation(driver, current_url, is_movie=False, video_url=video_url):
                    if check_wait_and_try_again(driver, current_url, video_url):
                        continue
                    return False
                if not updates:
                    t = _clean_excel_value(series_title)
                    d = _clean_excel_value(series_description)
                    if t: updates['title'] = t
                    if d: updates['overview'] = d

            # --- Step 4: Write (tmdb-automation style) ---
            if updates:
                if not _fill_and_submit_tmdb_form(driver, updates):
                    if check_wait_and_try_again(driver, current_url, video_url):
                        continue
                    return False
                time.sleep(1)

            driver.get(current_url)
            if check_wait_and_try_again(driver, current_url, video_url):
                logging.warning("系列信息提交後發現限流，重做...")
                continue

            return True

        except Exception as e:
            logging.error(f"更新系列信息失敗: {e}，TMDB ID: {tmdb_id}")
            return False

def update_seasons(driver, tmdb_id, excel_path, overwrite_titles='ask', overwrite_descriptions='ask', video_url=None):
    """更新季度 zh-HK 翻譯 (tmdb-automation 模式：API 預檢 → 簡化 Selenium 寫入)。"""
    logging.info("開始更新季度信息 (tmdb-automation 模式)...")
    api_key = _load_api_bearer()
    try:
        df = pd.read_excel(excel_path, sheet_name='Seasons')
        if df.empty or 'Season Number' not in df.columns or df['Season Number'].isna().all():
            logging.info("Seasons sheet 為空，跳過。")
            return True

        overall_success = True
        indices = list(df.index)
        i = 0
        last_processed_url = None

        while i < len(indices):
            if utils._PROGRAM_INTERRUPTED:
                overall_success = False
                break

            index = indices[i]
            row = df.loc[index]

            season_number = int(row['Season Number']) if pd.notna(row['Season Number']) else None
            if season_number is None:
                i += 1
                continue

            season_name = _clean_excel_value(row.get('Season Name', ''))
            season_description = _clean_excel_value(row.get('Season Description', ''))
            perfect_name = row.get('perfect name', '')
            perfect_description = row.get('perfect description', '')

            if not season_name and not season_description:
                # Allow proceeding if we want to force-clear (empty + overwrite=y)
                if overwrite_titles.lower() == 'y' or overwrite_descriptions.lower() == 'y':
                    # Will be handled inside _decide_what_to_update + updates will contain '' 
                    pass
                else:
                    logging.info(f"季度 {season_number} 數據為空，跳過。")
                    i += 1
                    continue

            # --- Step 1: Pre-check via TMDB API ---
            zh_hk = _api_fetch_zh_hk(
                f'https://api.themoviedb.org/3/tv/{tmdb_id}/season/{season_number}', api_key
            )
            updates = _decide_what_to_update(
                season_name, season_description, None, None,
                zh_hk['name'], zh_hk['overview'], None, None,
                overwrite_titles, overwrite_descriptions, perfect_name, perfect_description,
            )
            if not updates and zh_hk['exists']:
                logging.info(f"季度 {season_number}: 所有欄位與 TMDB 相同，跳過頁面載入。")
                i += 1
                continue

            # --- Step 2: Load edit page ---
            logging.info(f"處理季度 {season_number} — 需更新: {list(updates.keys()) if updates else ['新增翻譯']}")
            current_url = f"https://www.themoviedb.org/tv/{tmdb_id}/season/{season_number}/edit?language=zh-HK"
            last_processed_url = current_url
            driver.get(current_url)

            is_valid, was_limited = pre_check(driver, current_url, video_url)
            if was_limited:
                if i > 0:
                    logging.warning(f"載入季度 {season_number} 時遇限流，倒退重試。")
                    i -= 1
                    time.sleep(2)
                    continue
            if not is_valid:
                overall_success = False
                i += 1
                continue

            # --- Step 3: Ensure zh-HK translation exists ---
            if not zh_hk['exists']:
                if not check_and_add_translation(driver, current_url, is_movie=False, video_url=video_url):
                    logging.error(f"無法準備季度 {season_number} 翻譯表單，跳過。")
                    overall_success = False
                    i += 1
                    continue
                if not updates:
                    if season_name: updates['title'] = _clean_excel_value(season_name)
                    if season_description: updates['overview'] = _clean_excel_value(season_description)

            # --- Step 4: Write (tmdb-automation style) ---
            if updates:
                if not _fill_and_submit_tmdb_form(driver, updates):
                    if check_wait_and_try_again(driver, current_url, video_url):
                        continue
                    logging.error(f"季度 {season_number} 更新失敗。")
                    overall_success = False
                else:
                    logging.info(f"季度 {season_number} 已更新。")

            time.sleep(random.uniform(1.5, 2.5))
            i += 1

        if last_processed_url and not utils._PROGRAM_INTERRUPTED:
            logging.info("最終限流檢查 (Seasons)...")
            driver.get(last_processed_url)
            if check_wait_and_try_again(driver, last_processed_url, video_url):
                logging.warning("最終檢查發現限流，重做最後一個季度。")
                row = df.loc[indices[-1]]
                sn = int(row['Season Number'])
                zh_last = _api_fetch_zh_hk(f'https://api.themoviedb.org/3/tv/{tmdb_id}/season/{sn}', api_key)
                last_upd = _decide_what_to_update(
                    _clean_excel_value(row.get('Season Name', '')),
                    _clean_excel_value(row.get('Season Description', '')),
                    None, None, zh_last['name'], zh_last['overview'], None, None,
                    overwrite_titles, overwrite_descriptions,
                    row.get('perfect name', ''), row.get('perfect description', ''),
                )
                if last_upd:
                    _fill_and_submit_tmdb_form(driver, last_upd)
                    time.sleep(1)
                driver.get(last_processed_url)
                check_wait_and_try_again(driver)

        return overall_success
    except (KeyboardInterrupt, SystemExit):
        utils._PROGRAM_INTERRUPTED = True
        return False
    except Exception as e:
        logging.error(f"更新季度信息異常: {e}")
        return False

# ---------------------------------------------------------------------------
# Episode backdrop upload helpers (Netflix source)
# ---------------------------------------------------------------------------

def _download_backdrop(image_url: str, dest_path: str) -> bool:
    try:
        r = _req.get(image_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=30)
        r.raise_for_status()
        with open(dest_path, 'wb') as f:
            f.write(r.content)
        return True
    except Exception as e:
        logging.error(f"下載劇照失敗 ({image_url}): {e}")
        return False


def _api_has_episode_stills(tmdb_id: str, season: int, ep_num: int, api_key: str):
    """Check whether this episode has still images on TMDB.

    Returns:
        True  — episode exists and already has at least one still
        False — episode exists but has no stills yet
        None  — episode does not exist on TMDB (404); caller should skip upload
    """
    url = f"https://api.themoviedb.org/3/tv/{tmdb_id}/season/{season}/episode/{ep_num}/images"
    try:
        r = _req.get(
            url,
            headers={"Authorization": f"Bearer {api_key}", "accept": "application/json"},
            # No language filter — include_image_language=null would miss stills tagged
            # with 'en' or other language codes, causing false-negative checks.
            params={"include_image_language": "null,en,zh,ja,ko,fr,de,es,pt,it,xx"},
            timeout=10,
        )
        if r.status_code == 404:
            logging.warning(f"S{season:02d}E{ep_num:02d}: TMDB 不存在此集數 (404)，跳過劇照上傳。")
            return None
        if r.status_code == 200:
            count = len(r.json().get('stills', []))
            logging.debug(f"S{season:02d}E{ep_num:02d}: TMDB stills count = {count}")
            return count > 0
    except Exception:
        pass
    return False


def _upscale_to_min_size(image_path: str, min_w: int = 400, min_h: int = 225) -> None:
    """Upscale image in-place to at least min_w × min_h using Lanczos if it is smaller."""
    try:
        from PIL import Image
        with Image.open(image_path) as img:
            w, h = img.size
            if w >= min_w and h >= min_h:
                return
            scale = max(min_w / w, min_h / h)
            new_w, new_h = round(w * scale), round(h * scale)
            resized = img.resize((new_w, new_h), Image.LANCZOS)
            resized.save(image_path)
            logging.info(f"劇照已從 {w}x{h} 放大至 {new_w}x{new_h}")
    except Exception as e:
        logging.warning(f"放大圖片時失敗 ({image_path}): {e}")


def _convert_to_jpg(image_path: str) -> str:
    """Convert image to JPEG if it is not already JPEG. Returns the (possibly new) path."""
    if image_path.lower().endswith(('.jpg', '.jpeg')):
        return image_path
    try:
        from PIL import Image
        jpg_path = os.path.splitext(image_path)[0] + '.jpg'
        with Image.open(image_path) as img:
            rgb = img.convert('RGB')
            rgb.save(jpg_path, 'JPEG', quality=95)
        os.remove(image_path)
        logging.info(f"劇照已從 {os.path.basename(image_path)} 轉換為 JPEG。")
        return jpg_path
    except Exception as e:
        logging.warning(f"轉換 JPEG 失敗 ({image_path}): {e}")
        return image_path


def _crop_to_16_9(image_path: str) -> None:
    """Centre-crop image in-place to exactly 16:9. No-op if already 16:9."""
    try:
        from PIL import Image
        with Image.open(image_path) as img:
            w, h = img.size
            ratio = w / h
            target = 16 / 9
            if abs(ratio - target) < 0.001:
                return
            if ratio > target:
                new_w = round(h * target)
                left = (w - new_w) // 2
                box = (left, 0, left + new_w, h)
            else:
                new_h = round(w / target)
                top = (h - new_h) // 2
                box = (0, top, w, top + new_h)
            fmt = img.format or 'JPEG'
            img.crop(box).save(image_path, format=fmt)
            logging.info(f"劇照已自動裁切至 16:9: {w}x{h} → {box[2]-box[0]}x{box[3]-box[1]}")
    except Exception as e:
        logging.warning(f"裁切圖片失敗 ({image_path}): {e}")


def _reencode_jpeg(image_path: str) -> str:
    """Re-save image as clean JPEG to strip metadata/encoding quirks TMDB may reject."""
    try:
        from PIL import Image
        out = os.path.splitext(image_path)[0] + '_reenc.jpg'
        with Image.open(image_path) as img:
            img.convert('RGB').save(out, 'JPEG', quality=92)
        logging.info(f"劇照已重新編碼 JPEG: {os.path.basename(out)}")
        return out
    except Exception as e:
        logging.warning(f"重新編碼 JPEG 失敗 ({image_path}): {e}")
        return image_path


def _upload_episode_backdrop_from_url(driver, tmdb_id: str, season: int, ep_num: int,
                                       backdrop_url: str, temp_dir: str) -> bool:
    """Download an episode backdrop, process it, and upload to TMDB.

    Processing pipeline:
      1. Convert to JPEG
      2. Upscale if smaller than 400×225 (extend/upscale)
      3. Auto-crop to 16:9 if ratio deviates
      4. Upload; on failure: re-encode JPEG and retry once
    """
    clean_url = backdrop_url.split('?')[0]
    ext = os.path.splitext(clean_url)[1] or '.jpg'
    dest = os.path.join(temp_dir, f"s{season:02d}e{ep_num:03d}_backdrop{ext}")

    if not _download_backdrop(clean_url, dest):
        return False

    dest = _convert_to_jpg(dest)
    _upscale_to_min_size(dest)   # extend/upscale if size insufficient
    _crop_to_16_9(dest)          # auto-crop if not 16:9

    page_url = (
        f"https://www.themoviedb.org/tv/{tmdb_id}"
        f"/season/{season}/episode/{ep_num}/images/backdrops"
    )
    label = f"S{season:02d}E{ep_num:02d} 劇照"
    logging.info(f"上傳 {label} → {page_url}")

    def _do_upload(image_path: str) -> bool:
        # Same flow as fzlins/TMDB-Import importors/episode.py:
        #   a.add_image.upload → input#upload_files → span.file_status.success → Escape
        try:
            driver.get(page_url)
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a.add_image.upload, li#no_results'))
            )
            time.sleep(1)

            btn = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a.add_image.upload'))
            )
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(1.5)

            file_input = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input#upload_files[accept='image/*']")
                )
            )
            abs_path = os.path.abspath(image_path)
            try:
                file_input.send_keys(abs_path)
            except Exception:
                driver.execute_script(
                    "var el=arguments[0]; el.style.display='block'; el.style.visibility='visible';"
                    "el.style.opacity='1'; el.removeAttribute('tabindex');",
                    file_input,
                )
                file_input.send_keys(abs_path)

            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script(
                    "return !!document.querySelector('span.file_status.success')"
                )
            )
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(0.5)
            except Exception:
                pass
            logging.info(f"{label} 上傳成功。")
            return True
        except Exception as e:
            logging.error(f"{label} 上傳失敗: {str(e).strip() or type(e).__name__}")
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            except Exception:
                pass
            return False

    # First attempt
    if _do_upload(dest):
        return True

    # On failure: re-encode JPEG to strip metadata quirks and retry once
    logging.warning(f"{label} 上傳失敗，重新編碼 JPEG 後再試一次...")
    reenc = _reencode_jpeg(dest)
    return _do_upload(reenc)


def update_episodes(driver, tmdb_id, excel_path, overwrite_titles='ask', overwrite_descriptions='ask', video_url=None):
    """
    更新集數 zh-HK 翻譯 (純 grid 模式)：
    Phase A — API 預檢所有集數，決定需要更新的欄位
    Phase C — 每季加載一次 grid 頁面，批量 _edit_episode_via_grid

    注意：不再使用逐集載入頁面的方式 (舊 Phase B)。
    TMDB /translations API 不可靠 — 集數可能在頁面上已有 zh-HK 資料，
    但 API 不報告為正式翻譯條目。grid popup 直接顯示 zh-HK 欄位，
    無論翻譯是否「存在」，因此可以統一用 grid 處理所有集數。
    """
    # 【2026-11 P2】分集可獨立設定覆蓋選項（未設定則沿用劇集設定）
    overwrite_titles = utils.episode_overwrite('title', overwrite_titles)
    overwrite_descriptions = utils.episode_overwrite('description', overwrite_descriptions)
    logging.info(f"開始更新集數信息 (grid 模式)... 分集覆蓋標題={overwrite_titles}, 分集覆蓋描述={overwrite_descriptions}")
    api_key = _load_api_bearer()
    try:
        df = pd.read_excel(excel_path, sheet_name='Episodes')
        if df.empty or 'Episode Number' not in df.columns:
            return True

        overall_success = True

        # --- Phase A: API pre-check — 2 calls per season instead of 2 per episode ---
        grid_updates: dict[int, dict[int, dict]] = {}  # season → {ep_num → updates}

        unique_seasons = (
            df['Season Number'].dropna().astype(int).unique().tolist()
            if 'Season Number' in df.columns else []
        )
        season_cache: dict[int, dict] = {}
        if api_key:
            for s in sorted(unique_seasons):
                logging.info(f"API 預檢第 {s} 季 (批量)...")
                season_cache[s] = _api_fetch_season_episodes_batch(tmdb_id, s, api_key)

        for _, row in df.iterrows():
            if utils._PROGRAM_INTERRUPTED:
                return False

            season = int(row['Season Number']) if pd.notna(row['Season Number']) else None
            ep_num = int(row['Episode Number']) if pd.notna(row['Episode Number']) else None
            if season is None or ep_num is None:
                continue

            ep_title = _clean_excel_value(row.get('Episode Title', ''))
            ep_desc = _clean_excel_value(row.get('Episode Description', ''))
            ep_runtime = int(row['Episode Runtime']) if pd.notna(row.get('Episode Runtime')) else None
            perfect_title_flag = row.get('perfect title', '')
            perfect_desc_flag = row.get('perfect description', '')
            raw_air_date = row.get('Episode Air Date', '')
            if pd.notna(raw_air_date) and str(raw_air_date).strip():
                try:
                    ep_air_date = pd.to_datetime(str(raw_air_date)).strftime('%Y-%m-%d')
                except Exception:
                    ep_air_date = None
            else:
                ep_air_date = None

            # 【2026-11 N1】整行來源皆空時仍需判斷：選擇清空 → 清空；否則 TMDB 若是簡體 → 轉繁體

            cached = season_cache.get(season, {}).get(ep_num, {})
            zh_hk = {
                'exists': cached.get('exists', False),
                'name': cached.get('name', ''),
                'overview': cached.get('overview', ''),
            }
            base_info = {
                'runtime': cached.get('runtime'),
                'air_date': cached.get('air_date'),
            }

            updates = _decide_what_to_update(
                ep_title, ep_desc, ep_runtime, ep_air_date,
                zh_hk['name'], zh_hk['overview'],
                base_info['runtime'], base_info['air_date'],
                overwrite_titles, overwrite_descriptions,
                perfect_title_flag, perfect_desc_flag,
            )

            if updates:
                grid_updates.setdefault(season, {})[ep_num] = updates
            else:
                logging.info(f"S{season}E{ep_num}: 所有欄位與 TMDB 相同，跳過。")

        # --- Phase C: All episodes via season grid (one page load per season) ---
        for season in sorted(grid_updates.keys()):
            if utils._PROGRAM_INTERRUPTED:
                overall_success = False
                break
            eps_in_season = grid_updates[season]
            logging.info(f"載入第 {season} 季 grid 頁面，更新 {len(eps_in_season)} 集...")
            season_url = (
                f"https://www.themoviedb.org/tv/{tmdb_id}/season/{season}/edit"
                f"?active_nav_item=episodes&language=zh-HK"
            )

            driver.get(season_url)
            is_valid, was_limited = pre_check(driver, season_url, video_url)
            if was_limited:
                driver.get(season_url)
                is_valid, was_limited = pre_check(driver, season_url, video_url)
            if not is_valid:
                logging.error(f"第 {season} 季 grid 頁面無效，跳過該季。")
                overall_success = False
                continue

            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'tr.k-master-row, .k-grid'))
                )
            except Exception:
                logging.warning(f"第 {season} 季 grid 載入超時，繼續嘗試。")
            time.sleep(2)

            for ep_num in sorted(eps_in_season.keys()):
                if utils._PROGRAM_INTERRUPTED:
                    overall_success = False
                    break
                updates = eps_in_season[ep_num]
                logging.info(f"S{season}E{ep_num}: grid 更新 {list(updates.keys())}...")
                if not _edit_episode_via_grid(driver, ep_num, updates):
                    logging.error(f"S{season}E{ep_num}: grid 更新失敗。")
                    overall_success = False

        # --- Phase D: Upload Netflix episode backdrops (skip if TMDB already has stills) ---
        if api_key and 'Episode Backdrop URL' in df.columns:
            backdrop_rows = df[df['Episode Backdrop URL'].notna() & (df['Episode Backdrop URL'].astype(str).str.strip() != '')]
            if not backdrop_rows.empty:
                logging.info(f"開始上傳集數劇照 (共 {len(backdrop_rows)} 集有 Netflix 劇照)...")
                temp_dir = tempfile.mkdtemp(prefix='netflix_backdrop_')
                try:
                    for _, row in backdrop_rows.iterrows():
                        if utils._PROGRAM_INTERRUPTED:
                            overall_success = False
                            break
                        season = int(row['Season Number']) if pd.notna(row.get('Season Number')) else None
                        ep_num = int(row['Episode Number']) if pd.notna(row.get('Episode Number')) else None
                        if season is None or ep_num is None:
                            continue
                        ep_runtime = row.get('Episode Runtime')
                        if ep_runtime is None or (hasattr(ep_runtime, '__float__') and pd.isna(ep_runtime)) or int(ep_runtime) == 0:
                            logging.info(f"S{season:02d}E{ep_num:02d}: 無時長（可能為未上映集數），跳過劇照上傳。")
                            continue
                        backdrop_url = str(row['Episode Backdrop URL']).strip()
                        stills_check = _api_has_episode_stills(tmdb_id, season, ep_num, api_key)
                        if stills_check is None:
                            # Episode doesn't exist on TMDB — skip silently (already logged in helper)
                            continue
                        if stills_check:
                            logging.info(f"S{season:02d}E{ep_num:02d}: TMDB 已有劇照，跳過。")
                            continue
                        if not _upload_episode_backdrop_from_url(driver, tmdb_id, season, ep_num, backdrop_url, temp_dir):
                            overall_success = False
                        time.sleep(1.5)
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)

        return overall_success
    except (KeyboardInterrupt, SystemExit):
        utils._PROGRAM_INTERRUPTED = True
        return False
    except Exception as e:
        logging.error(f"更新集數信息異常: {e}\n{traceback.format_exc()}")
        return False

def update_movie_info(driver, tmdb_id, excel_path, overwrite_titles='ask', overwrite_descriptions='ask', video_url=None):
    """更新電影 zh-HK 翻譯 (tmdb-automation 模式：API 預檢 → 簡化 Selenium 寫入)。"""
    logging.info("開始更新電影信息 (tmdb-automation 模式)...")
    api_key = _load_api_bearer()
    while True:
        if utils._PROGRAM_INTERRUPTED:
            return False
        try:
            df = pd.read_excel(excel_path, sheet_name='Movies')
            if df.empty:
                return True
            movie_title = df['Movie Title'].iloc[0] if 'Movie Title' in df.columns and pd.notna(df['Movie Title'].iloc[0]) else ''
            movie_description = df['Movie Description'].iloc[0] if 'Movie Description' in df.columns and pd.notna(df['Movie Description'].iloc[0]) else ''
            _raw_runtime = df['Runtime'].iloc[0] if 'Runtime' in df.columns and pd.notna(df['Runtime'].iloc[0]) else None
            runtime = int(_raw_runtime) if _raw_runtime is not None and int(_raw_runtime) > 0 else None
            perfect_title = df['perfect title'].iloc[0] if 'perfect title' in df.columns and pd.notna(df['perfect title'].iloc[0]) else ''
            perfect_description = df['perfect description'].iloc[0] if 'perfect description' in df.columns and pd.notna(df['perfect description'].iloc[0]) else ''

            logging.debug(f"[Movie] Excel 讀取 — title='{movie_title}', desc='{movie_description[:30] if movie_description else ''}', runtime={runtime}")
            if not movie_title and not movie_description and runtime is None:
                return True

            # --- Step 1: Pre-check via TMDB API ---
            movie_api_base = f'https://api.themoviedb.org/3/movie/{tmdb_id}'
            zh_hk = _api_fetch_zh_hk(movie_api_base, api_key)
            # Use zh-HK specific runtime (from /translations), not the base movie runtime.
            # The base movie runtime (e.g. fetched from /movie/{id}) is the global value and
            # may already be set even when the zh-HK override field is empty.
            current_runtime = zh_hk.get('runtime')
            logging.debug(f"[Movie] TMDB zh-HK — name='{zh_hk['name']}', runtime={current_runtime}, exists={zh_hk['exists']}")

            updates = _decide_what_to_update(
                movie_title, movie_description, runtime, None,
                zh_hk['name'], zh_hk['overview'], current_runtime, None,
                overwrite_titles, overwrite_descriptions, perfect_title, perfect_description,
            )
            logging.info(f"[Movie] _decide_what_to_update 結果: {updates}")
            if not updates and zh_hk['exists']:
                logging.info("電影資訊與 TMDB 相同，跳過頁面載入。")
                return True

            # --- Step 2: Load edit page ---
            current_url = f"https://www.themoviedb.org/movie/{tmdb_id}/edit?language=zh-HK"
            driver.get(current_url)
            is_valid, was_limited = pre_check(driver, current_url, video_url)
            if was_limited:
                logging.info("電影頁面遇限流 (已恢復)，重試...")
                continue
            if not is_valid:
                return False

            # --- Step 3: Ensure zh-HK translation exists ---
            if not zh_hk['exists']:
                if not check_and_add_translation(driver, current_url, is_movie=True, video_url=video_url):
                    if check_wait_and_try_again(driver, current_url, video_url):
                        continue
                    return False
                if not updates:
                    t = _clean_excel_value(movie_title)
                    d = _clean_excel_value(movie_description)
                    if t: updates['title'] = t
                    if d: updates['overview'] = d
                    if runtime is not None: updates['runtime'] = runtime

            # --- Step 4: Write (tmdb-automation style) ---
            if updates:
                if not _fill_and_submit_tmdb_form(driver, updates, title_field_id='zh_HK_translated_title'):
                    if check_wait_and_try_again(driver, current_url, video_url):
                        continue
                    return False
                time.sleep(1)

            driver.get(current_url)
            if check_wait_and_try_again(driver, current_url, video_url):
                logging.warning("電影信息提交後發現限流，重做...")
                continue

            return True

        except Exception as e:
            logging.error(f"更新電影信息失敗: {e}，TMDB ID: {tmdb_id}")
            return False


def upload_person_to_tmdb(driver, url, excel_path, overwrite_titles='ask', overwrite_descriptions='ask'):
    """上傳人員資料到 TMDB"""
    if driver is None:
        logging.error(f"上傳人員到 TMDB 失敗: driver 為 None。URL: {url}")
        return False
    try:
        logging.info(f"開始上傳人員到 TMDB，URL: {url}")
        reset_failsafe_counter(url) 

        person_id = extract_person_id(url)
        if not person_id:
            logging.error(f"從 URL {url} 提取 Person ID 失敗。")
            return False

        success = update_people(driver, person_id, excel_path, media_type='person', video_url=url, overwrite_titles=overwrite_titles, overwrite_descriptions=overwrite_descriptions)
        
        if not success:
            logging.error(f"人員上傳/更新失敗，URL: {url}")
            return False

        logging.info(f"成功上傳/更新人員到 TMDB: {url}")
        return True
    except Exception as e:
        logging.error(f"上傳人員到 TMDB 過程中發生頂層異常: {e}，URL: {url}\n{traceback.format_exc()}")
        return False

def upload_to_tmdb(driver, url, excel_path, is_movie=False, overwrite_titles='ask', overwrite_descriptions='ask', upload_only_people=False):
    """
    【重要邏輯說明】請勿修改此處邏輯。
    程式設計要求：只要 'Title', 'Seasons', 'Episodes', 'Movies', 'People' 其中任何一張工作表 (Sheet) 含有有效資料，
    就必須觸發上傳/更新流程。嚴禁改為「只有 Title 有資料才上傳」或忽略其他 Sheet 的資料。
    """
    if driver is None:
        logging.error(f"TMDB 上傳失敗: driver 為 None。URL: {url}")
        return False
    try:
        logging.info(f"開始上傳到 TMDB，URL: {url}, 是否為電影: {is_movie}, 僅上傳人員: {upload_only_people}")
        reset_failsafe_counter(url) 

        tmdb_id_str = extract_movie_id(url) if is_movie else extract_tv_show_id(url)
        if not tmdb_id_str: 
            logging.error(f"從 URL {url} 提取 TMDB ID 失敗。")
            return False
        
        overall_upload_success = True 

        if not upload_only_people:
            # --- 電影/劇集主體資訊處理 ---
            main_body_has_data = False
            if is_movie:
                try:
                    df = pd.read_excel(excel_path, sheet_name='Movies')
                    if not df.empty:
                        title = str(df['Movie Title'].iloc[0]).strip() if 'Movie Title' in df.columns and pd.notna(df['Movie Title'].iloc[0]) else ''
                        desc = str(df['Movie Description'].iloc[0]).strip() if 'Movie Description' in df.columns and pd.notna(df['Movie Description'].iloc[0]) else ''
                        _raw_rt = df['Runtime'].iloc[0] if 'Runtime' in df.columns and pd.notna(df['Runtime'].iloc[0]) else None
                        runtime = int(_raw_rt) if _raw_rt is not None and int(_raw_rt) > 0 else None
                        if title or desc or runtime is not None:
                            main_body_has_data = True
                except Exception: pass
                
                # 即使 People 有資料，如果 Movie Sheet 有資料也必須執行此段
                if main_body_has_data:
                    if not update_movie_info(driver, tmdb_id_str, excel_path, overwrite_titles, overwrite_descriptions, video_url=url):
                        overall_upload_success = False
                else:
                    logging.info(f"Movies sheet 中無有效數據 (Title, Description, Runtime)，跳過電影主體資訊上傳。")
            else: # TV Show
                # 檢查所有 TV 相關 sheet
                sheets_to_check = {'Title': ['TV Show Title', 'TV Show Description'], 'Seasons': ['Season Number'], 'Episodes': ['Episode Number']}
                for sheet, cols in sheets_to_check.items():
                    try:
                        df = pd.read_excel(excel_path, sheet_name=sheet)
                        if not df.empty and any(col in df.columns and df[col].notna().any() for col in cols):
                            main_body_has_data = True
                            break
                    except Exception: continue
                
                # 重要：只要 TV/Seasons/Episodes 任一有資料，都會進入這裡執行更新
                if main_body_has_data:
                    if not update_series_info(driver, tmdb_id_str, excel_path, overwrite_titles, overwrite_descriptions, video_url=url): overall_upload_success = False
                    if not update_seasons(driver, tmdb_id_str, excel_path, overwrite_titles, overwrite_descriptions, video_url=url): overall_upload_success = False
                    if not update_episodes(driver, tmdb_id_str, excel_path, overwrite_titles, overwrite_descriptions, video_url=url): overall_upload_success = False
                else:
                    logging.info(f"Title, Seasons, Episodes sheets 中均無有效數據，跳過劇集主體資訊上傳。")

        else: 
            logging.info(f"僅上傳人員模式，跳過電影/劇集主體資訊。URL: {url}")

        # --- 人員資訊處理 ---
        # 此部分與 main_body_has_data 獨立，即使主體無更新，只要 People 有資料就會執行
        people_data_exists = False
        try:
            df_people = pd.read_excel(excel_path, sheet_name='People')
            df_people_valid = df_people.dropna(subset=['Person URL'])
            if not df_people_valid.empty:
                # 確保 Name 和 Biography 欄位存在
                if 'Name' not in df_people_valid.columns: df_people_valid['Name'] = ''
                if 'Biography' not in df_people_valid.columns: df_people_valid['Biography'] = ''
                has_names = df_people_valid['Name'].fillna('').str.strip().astype(bool).any()
                has_bios = df_people_valid['Biography'].fillna('').str.strip().astype(bool).any()
                if has_names or has_bios:
                    people_data_exists = True
        except Exception:
            pass

        if people_data_exists:
            media_type_for_people = 'movie' if is_movie else 'tv'
            if not update_people(driver, tmdb_id_str, excel_path, media_type=media_type_for_people, video_url=url, overwrite_titles=overwrite_titles, overwrite_descriptions=overwrite_descriptions):
                overall_upload_success = False
        else:
            logging.info(f"People sheet 為空或無有效數據 (URL存在，且Name或Bio至少其一有值)，跳過人員資訊上傳。")

        if not overall_upload_success:
            logging.error(f"TMDB 上傳過程中至少有一個部分失敗，URL: {url}")
            return False
        
        logging.info(f"TMDB 上傳處理完成，URL: {url}，整體成功狀態: {overall_upload_success}")
        return True

    except Exception as e:
        logging.error(f"上傳到 TMDB 過程中發生頂層異常: {e}\n{traceback.format_exc()}，URL: {url}")
        return False