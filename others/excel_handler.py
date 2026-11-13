import os
import logging
import pandas as pd
import openpyxl
import time
import traceback
import threading
import importlib.util
import re
import shutil
import zipfile
from datetime import datetime
from others.tw_hk_name import select_tw_hk_sheet


# 【2026-11】bulk_videos.xlsx 欄位。頭 4 欄固定（Done 一定係第 4 欄）；之後係每行設定，
# 平時留空 = 用程式選項；有值（y / n，覆蓋類亦可 ask）就取代程式選項（只影響嗰一行）。
BULK_BASE_COLUMNS = ['Video URL', 'TMDB URL', 'Split', 'Done']
BULK_OPTION_COLUMNS = {
    '覆蓋劇集標題': 'overwrite_titles',
    '覆蓋劇集描述': 'overwrite_descriptions',
    '覆蓋分集標題': 'episode_overwrite_titles',
    '覆蓋分集描述': 'episode_overwrite_descriptions',
    '來源無資料時清空TMDB': 'clear_when_source_empty',
    '維基人名取代現有': 'wiki_name_replace',
}
BULK_COLUMNS = BULK_BASE_COLUMNS + list(BULK_OPTION_COLUMNS)


def _add_missing_option_columns(workbook, bulk_sheet):
    """喺 bulk 格式嘅工作表（表頭頭兩欄係 Video URL / TMDB URL，或者係 bulk 選用嘅工作表）表頭尾加缺少嘅設定欄。
    回傳有冇改動。"""
    changed = False
    for ws in workbook.worksheets:
        header = [c.value for c in ws[1]] if ws.max_row >= 1 else []
        while header and header[-1] in (None, ''):
            header.pop()
        is_bulk = header[:2] == ['Video URL', 'TMDB URL'] or ws.title == bulk_sheet
        if not is_bulk or not header:
            continue
        missing = [c for c in BULK_OPTION_COLUMNS if c not in header]
        for i, col in enumerate(missing):
            ws.cell(row=1, column=len(header) + 1 + i, value=col)
        if missing:
            logging.info(f"已為工作表 '{ws.title}' 加入設定欄位: {missing}")
            changed = True
    return changed


def _bulk_sheet():
    """bulk 模式處理緊嘅工作表（啟動選項 utils.BULK_SHEET，默認 'Bulk'）。utils 會 import 本模組，所以延遲 import。"""
    from others import utils as _u
    return getattr(_u, 'BULK_SHEET', None) or 'Bulk'

# ... (保留變數定義 normalize_tmdb_url, clean_url 等不變) ...
# 全局變數，用於緩存 tw_hk_name.xlsx 的內容，避免在處理同一個項目時重複讀取檔案
_tw_hk_name_cache = {}
# 【2026-11 F3】每個 sheet 上一次成功讀取的內容；讀取失敗（例如 Excel 正在儲存）時沿用，不會清空
_tw_hk_name_last_good = {}

# 全局變數，用於 Done 值的批次寫入緩存和執行緒控制
_pending_done_updates = {}
_pending_done_lock = threading.Lock()
_batch_update_thread = None
_batch_worker_running = False
_excel_lock = None  # 由 start_batch_update_worker 注入，與 extractor 共用同一把鎖
_batch_permission_warned = False  # PermissionError 日誌節流，成功寫入後重置
_batch_flush_event = threading.Event()
_STALE_LOCK_MAX_AGE_SEC = 15
_last_batch_write_error = None  # 供結束提示區分「檔案損毀」與「Excel 開啟/lock」

def normalize_tmdb_url(url):
    """正規化 TMDB URL 為格式 https://www.themoviedb.org/(movie|tv)/(Video ID)"""
    if not isinstance(url, str) or not url:
        return ''
    try:
        url = url.strip()
        url = re.sub(r'[\u200b-\u200f\u202a-\u202e]', '', url)
        url = re.sub(r'\s+', ' ', url).strip()
        if "themoviedb.org" not in url:
            return url
        if "/tv/" in url:
            match = re.search(r'/tv/([^/?]+)', url)
            if match:
                video_id = match.group(1).split('-')[0]
                return f"https://www.themoviedb.org/tv/{video_id}"
        elif "/movie/" in url:
            match = re.search(r'/movie/([^/?]+)', url)
            if match:
                video_id = match.group(1).split('-')[0]
                return f"https://www.themoviedb.org/movie/{video_id}"
        logging.warning(f"無效的 TMDB URL，無法正規化: {url}")
        return url
    except Exception as e:
        logging.error(f"正規化 TMDB URL 失敗: {url}, 錯誤: {e}")
        return url

def clean_url(url):
    """清理 URL，移除多餘的空格、換行符和隱藏字符，並正規化 TMDB URL"""
    if pd.isna(url) or not isinstance(url, str):
        return ''
    url = normalize_tmdb_url(url)
    url = url.strip()
    url = re.sub(r'[\u200b-\u200f\u202a-\u202e]', '', url)
    url = re.sub(r'\s+', ' ', url)
    return url.strip()

def clean_sheet_data(file_name, sheet_name):
    """清理指定 sheet 的數據，People sheet 不再包含 Role。"""
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')
            if df.empty:
                logging.info(f"{sheet_name} sheet 為空，無需清理")
                return df

            if sheet_name == 'People':
                required_columns = ['Person URL', 'Name', 'Biography']
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = ''
                df = df.dropna(subset=['Person URL', 'Name'], how='all')
                is_name_valid = df['Name'].apply(lambda x: pd.notna(x) and str(x).strip() != '')
                is_bio_valid = df['Biography'].apply(lambda x: pd.notna(x) and str(x).strip() != '')
                df = df[is_name_valid | is_bio_valid]
                logging.info(f"清理後 People sheet 包含 {len(df)} 行有效數據")
            elif sheet_name == 'Title':
                is_title_valid = df['TV Show Title'].apply(lambda x: pd.notna(x) and str(x).strip() != '') if 'TV Show Title' in df.columns else pd.Series([False] * len(df), index=df.index)
                is_desc_valid = df['TV Show Description'].apply(lambda x: pd.notna(x) and str(x).strip() != '') if 'TV Show Description' in df.columns else pd.Series([False] * len(df), index=df.index)
                df = df[is_title_valid | is_desc_valid]
                logging.info(f"清理後 Title sheet 包含 {len(df)} 行有效數據")
            elif sheet_name == 'Seasons':
                if 'Season Number' in df.columns:
                    df = df[df['Season Number'].notna()]
                    df['Season Number'] = pd.to_numeric(df['Season Number'], errors='coerce')
                    df = df[df['Season Number'].notna()]
                logging.info(f"清理後 Seasons sheet 包含 {len(df)} 行有效數據")
            elif sheet_name == 'Episodes':
                if 'Season Number' in df.columns and 'Episode Number' in df.columns:
                    df = df[df['Season Number'].notna() & df['Episode Number'].notna()]
                    df['Season Number'] = pd.to_numeric(df['Season Number'], errors='coerce')
                    df['Episode Number'] = pd.to_numeric(df['Episode Number'], errors='coerce')
                    df = df[df['Season Number'].notna() & df['Episode Number'].notna()]
                logging.info(f"清理後 Episodes sheet 包含 {len(df)} 行有效數據")
            elif sheet_name == 'Movies':
                is_title_valid = df['Movie Title'].apply(lambda x: pd.notna(x) and str(x).strip() != '') if 'Movie Title' in df.columns else pd.Series([False] * len(df), index=df.index)
                is_desc_valid = df['Movie Description'].apply(lambda x: pd.notna(x) and str(x).strip() != '') if 'Movie Description' in df.columns else pd.Series([False] * len(df), index=df.index)
                is_runtime_valid = df['Runtime'].notna() if 'Runtime' in df.columns else pd.Series([False] * len(df), index=df.index)
                df = df[is_title_valid | is_desc_valid | is_runtime_valid]
                logging.info(f"清理後 Movies sheet 包含 {len(df)} 行有效數據")

            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            logging.info(f"成功清理並保存 {file_name} 中的 {sheet_name} sheet")
            return df
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"清理 {sheet_name} 時發生錯誤: {e}")
            raise

def check_create_excel(file_name='data/video_detail.xlsx'):
    """檢查並創建 Excel 文件，新增 'perfect' 相關欄位，並安全地更新現有文件結構。"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    required_sheets = {
        'Title': ['TV Show Title', 'TV Show Description', 'perfect title', 'perfect description'],
        'Seasons': ['Season Name', 'Season Number', 'Season Description', 'perfect name', 'perfect description'],
        'Episodes': ['Season Number', 'Episode Number', 'Episode Runtime', 'Episode Title', 'Episode Description', 'perfect title', 'perfect description'],
        'Movies': ['Movie Title', 'Movie Description', 'Runtime', 'perfect title', 'perfect description'],
        'People': ['Person URL', 'Name', 'Biography', 'perfect name', 'perfect bio']
    }
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    for sheet, columns in required_sheets.items():
                        pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet, index=False)
                logging.info(f"已創建新的 {file_name}，包含工作表: {list(required_sheets.keys())}")
            else:
                all_sheets_data = pd.read_excel(file_name, sheet_name=None, engine='openpyxl')
                structure_was_updated = False

                for sheet_name, required_columns in required_sheets.items():
                    if sheet_name not in all_sheets_data:
                        all_sheets_data[sheet_name] = pd.DataFrame(columns=required_columns)
                        logging.info(f"工作表 {sheet_name} 不存在於 {file_name}，已在記憶體中創建。")
                        structure_was_updated = True
                    else:
                        df = all_sheets_data[sheet_name]
                        current_columns = df.columns.tolist()
                        missing_columns = [col for col in required_columns if col not in current_columns]
                        
                        if missing_columns:
                            for col in missing_columns:
                                df[col] = ''
                            all_sheets_data[sheet_name] = df[required_columns]
                            logging.info(f"已為 {sheet_name} sheet 添加/修正欄位: {missing_columns}")
                            structure_was_updated = True
                
                if structure_was_updated:
                    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                        for sheet_name_to_write, df_to_write in all_sheets_data.items():
                            df_to_write.to_excel(writer, sheet_name=sheet_name_to_write, index=False)
                    logging.info(f"已成功更新 {file_name} 的文件結構。")
                else:
                    logging.info(f"{file_name} 已存在且結構有效。")

            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"創建或檢查 {file_name} 時發生錯誤: {e}\n{traceback.format_exc()}")
            raise

def check_create_unavailable_excel(file_name='data/unavailable_video.xlsx'):
    """檢查並創建 unavailable_video.xlsx，包含必要的 sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    required_sheets = {
        'English': ['Video URL', 'TMDB URL'],
        'Unavailable': ['Video URL', 'TMDB URL'],
        'Unsuccessful': ['Video URL', 'TMDB URL'],
        'SkippedIdentical': ['Video URL', 'TMDB URL'],
        'ProcessedEpisodes': ['Video URL', 'TMDB URL', 'Status', 'Processed At'],
        'TranslationErrors': ['Video URL', 'TMDB URL', 'Error Reason', 'Timestamp'],
        'NoChineseContent': ['Video URL', 'TMDB URL']
    }
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    for sheet, columns in required_sheets.items():
                        pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet, index=False)
                logging.info(f"已創建新的 {file_name}")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"創建或檢查 {file_name} 時發生錯誤: {e}")
            raise

def check_create_bulk_excel(file_name='data/bulk_videos.xlsx'):
    """檢查並創建 bulk_videos.xlsx，確保有 bulk 模式選用嘅工作表（默認 Bulk）"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    required_sheets = {
        _bulk_sheet(): BULK_COLUMNS
    }
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                write_excel_atomic(file_name, {
                    sheet: pd.DataFrame(columns=columns)
                    for sheet, columns in required_sheets.items()
                })
                logging.info(f"已創建新的 {file_name}，包含工作表: {list(required_sheets.keys())}")
            else:
                workbook = openpyxl.load_workbook(file_name)
                existing_sheets = workbook.sheetnames
                wb_modified = False
                for sheet, columns in required_sheets.items():
                    if sheet not in existing_sheets:
                        workbook.create_sheet(sheet)
                        df = pd.DataFrame(columns=columns)
                        # 注意：這裡不能同時用 pandas 寫入和 openpyxl save，會衝突。
                        # 正確做法是只用 openpyxl 添加 sheet 頭
                        ws = workbook[sheet]
                        ws.append(columns)
                        logging.info(f"工作表 {sheet} 不存在於 {file_name}，已創建")
                        wb_modified = True
                    else:
                        # 檢查欄位是否存在
                        ws = workbook[sheet]
                        header = [cell.value for cell in ws[1]]
                        if sheet == _bulk_sheet() and 'Done' not in header:
                            workbook.close()
                            df = pd.read_excel(file_name, sheet_name=sheet, engine='openpyxl')
                            df['Done'] = ''
                            replace_excel_sheet_atomic(file_name, sheet, df)
                            logging.info(f"已為 {file_name} 的 {sheet} 添加缺少的 'Done' 欄位")
                            return

                if _add_missing_option_columns(workbook, _bulk_sheet()):
                    wb_modified = True
                if wb_modified:
                    _atomic_save_workbook(workbook, file_name)
                else:
                    workbook.close()
                logging.info(f"{file_name} 已存在且有效")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"創建或檢查 {file_name} 時發生錯誤: {e}")
            raise

def remove_duplicate_video_urls(file_name='data/bulk_videos.xlsx', sheet_name=None):
    """移除 bulk_videos.xlsx 的指定 sheet 中重複的 Video URL 和 TMDB URL 組合"""
    sheet_name = sheet_name or _bulk_sheet()
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                logging.warning(f"{file_name} 不存在，無法執行去重操作")
                return
            try:
                df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')
            except ValueError:
                logging.warning(f"{sheet_name} sheet 不存在於 {file_name}，無法執行去重操作")
                return
            if df.empty:
                logging.info(f"{sheet_name} sheet 為空，無需去重")
                return
            if 'Video URL' in df.columns:
                df['Video URL'] = df['Video URL'].apply(clean_url)
            if 'TMDB URL' in df.columns:
                df['TMDB URL'] = df['TMDB URL'].apply(clean_url)
            original_rows = len(df)
            df = df.drop_duplicates(subset=['Video URL', 'TMDB URL'], keep='first')
            replace_excel_sheet_atomic(file_name, sheet_name, df)
            logging.info(f"成功移除 {original_rows - len(df)} 個重複 Video URL 行，更新 {file_name} 的 {sheet_name} sheet，現有行數: {len(df)}")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"移除重複 Video URL 時發生錯誤: {e}")
            raise

def save_no_chinese_content_info_to_excel(video_url, tmdb_url, file_name='data/unavailable_video.xlsx'):
    """將因無中文內容而跳過主體更新的影片資訊保存到 NoChineseContent sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            check_create_unavailable_excel(file_name)
            try:
                df = pd.read_excel(file_name, sheet_name='NoChineseContent', engine='openpyxl')
            except ValueError:
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])

            new_row = pd.DataFrame({'Video URL': [clean_url(video_url)], 'TMDB URL': [clean_url(tmdb_url)]})
            df = pd.concat([df, new_row], ignore_index=True)
            df.drop_duplicates(subset=['Video URL', 'TMDB URL'], keep='first', inplace=True)

            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='NoChineseContent', index=False)
            logging.info(f"因無中文內容跳過的影片資訊已保存到 {file_name} 的 NoChineseContent sheet: Video URL={video_url}, TMDB URL={tmdb_url}")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存無中文內容影片資訊到 {file_name} 時發生錯誤: {e}")
            raise

def clear_sheet(file_name, sheet_name):
    """清空指定 sheet 的內容，但保留欄位標題"""
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')
            columns = df.columns
            empty_df = pd.DataFrame(columns=columns)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
            logging.info(f"成功清空 {file_name} 中的 {sheet_name} 工作表")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"清空 {sheet_name} 時發生錯誤: {e}")
            raise

def clear_bulk_excel(file_name='data/bulk_videos.xlsx'):
    """清空 bulk_videos.xlsx 的 bulk 工作表內容，但保留欄位標題"""
    sheet = _bulk_sheet()
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            df = pd.read_excel(file_name, sheet_name=sheet, engine='openpyxl')
            columns = df.columns
            empty_df = pd.DataFrame(columns=columns)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                empty_df.to_excel(writer, sheet_name=sheet, index=False)
            logging.info(f"成功清空 {file_name} 中的 {sheet} 工作表")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"清空 Bulk 時發生錯誤: {e}")
            raise

def save_series_info_to_excel(title, description, perfect_title, perfect_description, file_name='data/video_detail.xlsx'):
    """將劇集資訊保存到 Excel 的 Title sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            title_df = pd.DataFrame({
                'TV Show Title': [title], 
                'TV Show Description': [description],
                'perfect title': [perfect_title],
                'perfect description': [perfect_description]
            })
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                title_df.to_excel(writer, sheet_name='Title', index=False)
            clean_sheet_data(file_name, 'Title')
            logging.info(f"劇集資訊已保存並清理到 {file_name} 的 Title sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存劇集資訊到 Title sheet 時發生錯誤: {e}")
            raise

def save_season_info_to_excel(seasons, file_name='data/video_detail.xlsx'):
    """將季度資訊保存到 Excel 的 Seasons sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            season_df = pd.DataFrame(seasons, columns=['Season Name', 'Season Number', 'Season Description', 'perfect name', 'perfect description'])
            season_df['Season Number'] = pd.to_numeric(season_df['Season Number'], errors='coerce')
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                season_df.to_excel(writer, sheet_name='Seasons', index=False)
            clean_sheet_data(file_name, 'Seasons')
            logging.info(f"季度資訊已保存並清理到 {file_name} 的 Seasons sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存季度資訊到 Seasons sheet 時發生錯誤: {e}")
            raise

def save_episode_info_to_excel(episodes, file_name='data/video_detail.xlsx'):
    """將集數資訊保存到 Excel 的 Episodes sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            episode_df = pd.DataFrame(
                episodes,
                columns=['Season Number', 'Episode Number', 'Episode Runtime', 'Episode Title', 'Episode Description', 'perfect title', 'perfect description', 'Episode Backdrop URL']
            )
            episode_df['Season Number'] = pd.to_numeric(episode_df['Season Number'], errors='coerce')
            episode_df['Episode Number'] = pd.to_numeric(episode_df['Episode Number'], errors='coerce')
            episode_df['Episode Runtime'] = pd.to_numeric(episode_df['Episode Runtime'], errors='coerce')
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                episode_df.to_excel(writer, sheet_name='Episodes', index=False)
            clean_sheet_data(file_name, 'Episodes')
            logging.info(f"集數資訊已保存並清理到 {file_name} 的 Episodes sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存集數資訊到 Episodes sheet 時發生錯誤: {e}")
            raise

def save_movie_info_to_excel(title, description, runtime, perfect_title, perfect_description, file_name='data/video_detail.xlsx'):
    """將電影資訊保存到 Excel 的 Movies sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            movie_df = pd.DataFrame({
                'Movie Title': [title],
                'Movie Description': [description],
                'Runtime': [runtime],
                'perfect title': [perfect_title],
                'perfect description': [perfect_description]
            })
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                movie_df.to_excel(writer, sheet_name='Movies', index=False)
            clean_sheet_data(file_name, 'Movies')
            logging.info(f"電影資訊已保存並清理到 {file_name} 的 Movies sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存電影資訊到 Movies sheet 時發生錯誤: {e}")
            raise

def save_actor_info_to_excel(actors, file_name='data/video_detail.xlsx'):
    """將人員資訊保存到 Excel 的 People sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            check_create_excel(file_name)
            try:
                existing_df = pd.read_excel(file_name, sheet_name='People', engine='openpyxl')
            except (ValueError, FileNotFoundError):
                existing_df = pd.DataFrame(columns=['Person URL', 'Name', 'Biography', 'perfect name', 'perfect bio'])
            
            new_df = pd.DataFrame(actors, columns=['Person URL', 'Name', 'Biography', 'perfect name', 'perfect bio'])
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)

            if 'Person URL' in combined_df.columns and not combined_df.empty:
                combined_df.drop_duplicates(subset=['Person URL'], keep='last', inplace=True)

            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                combined_df.to_excel(writer, sheet_name='People', index=False)
            
            logging.info(f"人員資訊已保存到 {file_name} 的 People sheet，包含 {len(combined_df)} 個人員")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存人員資訊到 People sheet 時發生錯誤: {e}\n{traceback.format_exc()}")
            raise

def save_english_site_info_to_excel(video_url, tmdb_url, file_name='data/unavailable_video.xlsx'):
    """將僅有英文標題的網站資訊保存到 Excel 的 English sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
                df.to_excel(file_name, index=False, engine='openpyxl')
                logging.info(f"已創建新的 {file_name}")
            try:
                df = pd.read_excel(file_name, sheet_name='English', engine='openpyxl')
            except ValueError:
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
            new_row = pd.DataFrame({'Video URL': [clean_url(video_url)], 'TMDB URL': [clean_url(tmdb_url)]})
            df = pd.concat([df, new_row], ignore_index=True)
            df.drop_duplicates(subset=['Video URL', 'TMDB URL'], keep='first', inplace=True)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='English', index=False)
            logging.info(f"英文網站資訊已保存到 {file_name} 的 English sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存英文網站資訊到 {file_name} 時發生錯誤: {e}")
            raise

def save_unavailable_video_info_to_excel(video_url, tmdb_url, file_name='data/unavailable_video.xlsx'):
    """將無法處理的影片資訊保存到 Excel 的 Unavailable sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
                df.to_excel(file_name, index=False, engine='openpyxl')
                logging.info(f"已創建新的 {file_name}")
            try:
                df = pd.read_excel(file_name, sheet_name='Unavailable', engine='openpyxl')
            except ValueError:
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
            new_row = pd.DataFrame({'Video URL': [clean_url(video_url)], 'TMDB URL': [clean_url(tmdb_url)]})
            df = pd.concat([df, new_row], ignore_index=True)
            df.drop_duplicates(subset=['Video URL', 'TMDB URL'], keep='first', inplace=True)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='Unavailable', index=False)
            logging.info(f"不可用影片資訊已保存到 {file_name} 的 Unavailable sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存不可用影片資訊到 {file_name} 時發生錯誤: {e}")
            raise

def save_translation_error_to_excel(video_url, tmdb_url, error_reason, file_name='data/unavailable_video.xlsx'):
    """將新增翻譯錯誤保存到 Excel 的 TranslationErrors sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            check_create_unavailable_excel(file_name)
            try:
                df = pd.read_excel(file_name, sheet_name='TranslationErrors', engine='openpyxl')
            except ValueError:
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL', 'Error Reason', 'Timestamp'])
            new_row = pd.DataFrame({
                'Video URL': [clean_url(video_url)],
                'TMDB URL': [clean_url(tmdb_url)],
                'Error Reason': [error_reason],
                'Timestamp': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
            })
            df = pd.concat([df, new_row], ignore_index=True)
            df.drop_duplicates(subset=['Video URL', 'TMDB URL', 'Error Reason'], keep='last', inplace=True)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='TranslationErrors', index=False)
            logging.info(f"新增翻譯錯誤已保存到 {file_name} 的 TranslationErrors sheet: {video_url}, 錯誤原因={error_reason}")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存新增翻譯錯誤到 {file_name} 時發生錯誤: {e}")
            raise

def save_unsuccessful_video_info_to_excel(video_url, tmdb_url, file_name='data/unavailable_video.xlsx'):
    """將上傳失敗的影片資訊保存到 Excel 的 Unsuccessful sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
                df.to_excel(file_name, index=False, engine='openpyxl')
                logging.info(f"已創建新的 {file_name}")
            try:
                df = pd.read_excel(file_name, sheet_name='Unsuccessful', engine='openpyxl')
            except ValueError:
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
            new_row = pd.DataFrame({'Video URL': [clean_url(video_url)], 'TMDB URL': [clean_url(tmdb_url)]})
            df = pd.concat([df, new_row], ignore_index=True)
            df.drop_duplicates(subset=['Video URL', 'TMDB URL'], keep='first', inplace=True)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='Unsuccessful', index=False)
            logging.info(f"不成功的影片資訊已保存到 {file_name} 的 Unsuccessful sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存不成功影片資訊到 {file_name} 時發生錯誤: {e}")
            raise

def save_skipped_identical_video_info_to_excel(video_url, tmdb_url, file_name='data/unavailable_video.xlsx'):
    """將因所有數據與 zh-HK 相同而跳過的影片資訊保存到 Excel 的 SkippedIdentical sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
                df.to_excel(file_name, index=False, engine='openpyxl')
                logging.info(f"已創建新的 {file_name}")
            try:
                df = pd.read_excel(file_name, sheet_name='SkippedIdentical', engine='openpyxl')
            except ValueError:
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL'])
            new_row = pd.DataFrame({'Video URL': [clean_url(video_url)], 'TMDB URL': [clean_url(tmdb_url)]})
            df = pd.concat([df, new_row], ignore_index=True)
            df.drop_duplicates(subset=['Video URL', 'TMDB URL'], keep='first', inplace=True)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='SkippedIdentical', index=False)
            logging.info(f"因所有數據與 zh-HK 相同而跳過的影片資訊已保存到 {file_name} 的 SkippedIdentical sheet")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存跳過的影片資訊到 {file_name} 時發生錯誤: {e}")
            raise

def save_processed_episode_info_to_excel(video_url, tmdb_url, status="skipped", file_name='data/unavailable_video.xlsx'):
    """將已處理的集數資訊保存到 Excel 的 ProcessedEpisodes sheet"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            if not os.path.exists(file_name):
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL', 'Status', 'Processed At'])
                df.to_excel(file_name, index=False, engine='openpyxl')
                logging.info(f"已創建新的 {file_name}")
            try:
                df = pd.read_excel(file_name, sheet_name='ProcessedEpisodes', engine='openpyxl')
            except ValueError:
                df = pd.DataFrame(columns=['Video URL', 'TMDB URL', 'Status', 'Processed At'])
            new_row = pd.DataFrame({
                'Video URL': [clean_url(video_url)],
                'TMDB URL': [clean_url(tmdb_url)],
                'Status': [status],
                'Processed At': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
            })
            df = pd.concat([df, new_row], ignore_index=True)
            df.drop_duplicates(subset=['Video URL', 'TMDB URL'], keep='last', inplace=True)
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='ProcessedEpisodes', index=False)
            logging.info(f"已處理的集數資訊已保存到 {file_name} 的 ProcessedEpisodes sheet: {video_url}, 狀態={status}")
            return
        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 {file_name}，檔案可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"保存已處理的集數資訊到 {file_name} 時發生錯誤: {e}")
            raise

def save_skipped_identical_to_bulk_excel(row_index, video_url, tmdb_url, file_name='data/bulk_videos.xlsx', sheet_name=None):
    """將因與 zh-HK 相同而跳過的影片記錄到 bulk_videos.xlsx 的 Done 欄位，設為 1"""
    try:
        if row_index is None or row_index < 0:
            logging.debug("非 bulk 模式（無效 row_index），跳過 Done 值寫入。")
            return
        update_done_value_fast(row_index, 1, file_name=file_name)
    except Exception as e:
        logging.error(f"更新 SkippedIdentical 記錄到 {file_name} 時發生頂層錯誤: {e}\n{traceback.format_exc()}")
        raise

def _remove_stale_lock_file(lock_file, force=False):
    """移除過期或強制指定的 .lock 檔。回傳 True 表示檔案已不存在、可建立新鎖。"""
    if not os.path.exists(lock_file):
        return True
    if not force:
        try:
            age = time.time() - os.path.getmtime(lock_file)
        except OSError:
            age = _STALE_LOCK_MAX_AGE_SEC + 1
        if age <= _STALE_LOCK_MAX_AGE_SEC:
            return False
    try:
        os.remove(lock_file)
        logging.warning(f"[BatchUpdate] 已移除過期 lock 檔: {lock_file}")
        return True
    except OSError as e:
        logging.warning(f"[BatchUpdate] 無法移除 lock 檔 {lock_file}: {e}")
        return False


def _xlsx_backup_path(file_name):
    directory = os.path.dirname(file_name) or '.'
    return os.path.join(directory, 'backup', os.path.basename(file_name))


def _xlsx_is_readable(path):
    """xlsx 是 zip；缺少 [Content_Types].xml 代表檔案已截斷/損毀。"""
    if not path or not os.path.exists(path) or os.path.getsize(path) < 64:
        return False
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            if '[Content_Types].xml' not in zf.namelist():
                return False
        wb = openpyxl.load_workbook(path, read_only=True)
        wb.close()
        return True
    except Exception:
        return False


def _is_corrupt_xlsx_error(exc):
    msg = str(exc).lower()
    return any(token in msg for token in (
        '[content_types].xml',
        'file is not a zip file',
        'bad zip file',
        'truncated',
        'central directory',
    ))


def _try_restore_corrupt_xlsx(file_name):
    """把損毀的 xlsx 另存，並用 data/backup 下的備份覆蓋。成功回傳 True。"""
    backup = _xlsx_backup_path(file_name)
    if not os.path.exists(backup):
        logging.error(f"{file_name} 已損毀，且找不到備份 {backup}。")
        return False
    if not _xlsx_is_readable(backup):
        logging.error(f"{file_name} 已損毀，備份 {backup} 同樣無法開啟。")
        return False
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    corrupt_copy = f"{file_name}.corrupt_{ts}"
    try:
        shutil.copy2(file_name, corrupt_copy)
    except OSError:
        corrupt_copy = None
    shutil.copy2(backup, file_name)
    logging.warning(
        f"{file_name} 已損毀（xlsx zip 不完整）。"
        f"已從備份還原: {backup}"
        + (f"；損毀檔另存為 {corrupt_copy}" if corrupt_copy else "")
    )
    return True


def _atomic_save_workbook(wb, file_name):
    """先寫入暫存檔並驗證，再 os.replace，避免中途中斷把正式檔截斷。"""
    tmp_path = file_name + '.tmp.xlsx'
    wb.save(tmp_path)
    _replace_xlsx_if_valid(tmp_path, file_name)


def _replace_xlsx_if_valid(tmp_path, file_name):
    """驗證暫存 xlsx 完整後，原子取代正式檔，並更新 data/backup。"""
    if not _xlsx_is_readable(tmp_path):
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise RuntimeError(f"暫存檔 {tmp_path} 寫入後驗證失敗，未覆蓋 {file_name}")
    os.replace(tmp_path, file_name)
    backup = _xlsx_backup_path(file_name)
    try:
        os.makedirs(os.path.dirname(backup), exist_ok=True)
        shutil.copy2(file_name, backup)
    except OSError as e:
        logging.debug(f"更新 xlsx 備份失敗（可忽略）: {e}")


def write_excel_atomic(file_name, sheets):
    """
    將完整工作簿寫入暫存檔，驗證 zip 後再取代 file_name。
    sheets: {sheet_name: DataFrame}
    不要用 pandas ExcelWriter 直接寫正式檔：中斷時會截斷 zip。
    """
    tmp_path = file_name + '.tmp.xlsx'
    with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            if df is None:
                pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    _replace_xlsx_if_valid(tmp_path, file_name)


def replace_excel_sheet_atomic(file_name, sheet_name, df):
    """保留其他工作表，只替換一個 sheet，全程走暫存檔。"""
    sheets = {}
    if os.path.exists(file_name) and _xlsx_is_readable(file_name):
        with pd.ExcelFile(file_name, engine='openpyxl') as xl:
            for name in xl.sheet_names:
                sheets[name] = pd.read_excel(xl, sheet_name=name)
    elif os.path.exists(file_name) and not _xlsx_is_readable(file_name):
        if not _try_restore_corrupt_xlsx(file_name):
            raise RuntimeError(f"{file_name} 已損毀，無法寫入。")
        with pd.ExcelFile(file_name, engine='openpyxl') as xl:
            for name in xl.sheet_names:
                sheets[name] = pd.read_excel(xl, sheet_name=name)
    sheets[sheet_name] = df
    write_excel_atomic(file_name, sheets)


def flush_pending_done_updates(file_name='data/bulk_videos.xlsx', force=False):
    """立即嘗試把記憶體中的 Done 值寫入 Excel。force=True 用於中斷/結束，會清過期 lock 並等待 EXCEL_LOCK。"""
    with _pending_done_lock:
        remaining = len(_pending_done_updates)
    if remaining == 0:
        return True
    logging.info(f"立即寫入 {remaining} 筆 Done 值 (force={force})...")
    return _perform_batch_update(file_name, max_attempts=5 if force else 1, force=force)


def start_batch_update_worker(excel_lock=None):
    """
    啟動背景執行緒，定期將緩存的 Done 值寫入 Excel。
    應在 Bulk 模式開始時調用。
    excel_lock: 與 extractor threads 共用的 threading.Lock，確保讀寫互斥。
    """
    global _batch_update_thread, _batch_worker_running, _excel_lock
    if _batch_update_thread is not None and _batch_update_thread.is_alive():
        logging.info("Batch update worker 已經在運行中。")
        return

    _excel_lock = excel_lock
    logging.info("啟動 Batch update worker 背景執行緒...")
    _batch_worker_running = True
    _batch_flush_event.clear()
    # name 要與 tmdb_importer.setup_loggers 的 importer.log 過濾名單一致，否則日誌只出現在 console
    _batch_update_thread = threading.Thread(target=_batch_update_worker, name="BatchUpdateThread", daemon=False)
    _batch_update_thread.start()

def stop_batch_update_worker(file_name='data/bulk_videos.xlsx'):
    """
    停止背景執行緒，並將剩餘的緩存數據強制寫入 Excel。
    應在程序結束前調用。
    """
    global _batch_worker_running, _batch_update_thread
    logging.info("正在停止 Batch update worker...")
    _batch_worker_running = False
    _batch_flush_event.set()
    
    if _batch_update_thread:
        _batch_update_thread.join(timeout=10)
        _batch_update_thread = None
    
    with _pending_done_lock:
        remaining = len(_pending_done_updates)
    if remaining > 0:
        logging.info(f"執行最後一次強制寫入剩餘的 {remaining} 筆 Done 值...")
        flush_pending_done_updates(file_name, force=True)
        with _pending_done_lock:
            still_pending = len(_pending_done_updates)
        while still_pending > 0:
            reason = _last_batch_write_error or f"{file_name} 可能仍被 Excel 開啟、檔案損毀或 lock 檔殘留"
            logging.warning(f"仍有 {still_pending} 筆 Done 值未能寫入。原因: {reason}")
            try:
                print(f"\n⚠  {still_pending} 筆 Done 值未能寫入 {file_name}。")
                print(f"   原因: {reason}")
                print("   請關閉 Excel（若檔案損毀，程式會嘗試從 data/backup 還原）後按 Enter 重試，或按 Ctrl+C 跳過...")
                from others import utils as _u
                with _u.hold_watchdog():  # 【2026-11 L1】等使用者處理期間，不要自動結束
                    input()
            except (EOFError, KeyboardInterrupt):
                logging.warning("使用者略過 Done 值寫入，部分紀錄未更新。")
                break
            flush_pending_done_updates(file_name, force=True)
            with _pending_done_lock:
                still_pending = len(_pending_done_updates)
        if still_pending == 0:
            logging.info("Done 值已成功全部寫入。")

def pending_done_count():
    """記憶體中仍未寫入 Excel 嘅 Done 值數目。"""
    with _pending_done_lock:
        return len(_pending_done_updates)


def get_pending_done_value(row_index: int):
    """Return the pending Done value for row_index if it's in the batch cache, else None."""
    with _pending_done_lock:
        return _pending_done_updates.get(row_index, None)

def update_done_value_fast(row_index: int, done_value: int, file_name='data/bulk_videos.xlsx'):
    """
    將 Done 值更新加入緩存，並立刻喚醒背景執行緒寫入 Excel。
    此函數本身不阻塞；實際寫入由 worker 在約 1 秒內完成。
    Single 模式的 row_index=-1 不可寫入，否則會覆蓋表頭。
    """
    global _pending_done_updates

    if row_index is None or row_index < 0:
        logging.debug(f"忽略無效 row_index={row_index} 的 Done 更新（非 bulk 模式）。")
        return
    
    with _pending_done_lock:
        _pending_done_updates[row_index] = done_value
        cache_size = len(_pending_done_updates)
    
    logging.info(f"已將 Done 值更新 (row_index={row_index}, value={done_value}) 加入批次緩存，當前緩存大小: {cache_size}。")
    
    if _batch_update_thread is None or not _batch_update_thread.is_alive():
        start_batch_update_worker()
    _batch_flush_event.set()

def _batch_update_worker():
    """
    背景執行緒：有新的 Done 值時約 1 秒內寫入 Excel。
    寫入失敗（Excel 被佔用、過期 lock）時保留記憶體，稍後再試。
    """
    from others import utils
    
    file_name = 'data/bulk_videos.xlsx'
    debounce_sec = 1.0
    
    while _batch_worker_running and not utils._PROGRAM_INTERRUPTED:
        _batch_flush_event.wait(timeout=2.0)
        _batch_flush_event.clear()
        if not _batch_worker_running or utils._PROGRAM_INTERRUPTED:
            break
        with _pending_done_lock:
            count = len(_pending_done_updates)
        if count == 0:
            continue
        time.sleep(debounce_sec)
        _perform_batch_update(file_name)

    with _pending_done_lock:
        leftover = len(_pending_done_updates)
    if leftover:
        logging.info(f"Batch worker 結束前寫入剩餘 {leftover} 筆 Done 值...")
        _perform_batch_update(file_name, max_attempts=3, force=True)

def _perform_batch_update(file_name, max_attempts=1, force=False):
    """
    執行批次寫入 Done 值到 Excel。
    force=True：清過期 lock、blocking 等待 EXCEL_LOCK（中斷/結束時使用）。
    回傳 True 表示已成功寫入或無需寫入，False 表示暫未寫入。
    """
    global _pending_done_updates, _batch_permission_warned
    
    with _pending_done_lock:
        if not _pending_done_updates:
            return True
        updates_to_write = _pending_done_updates.copy()

    lock_file = file_name + '.lock'

    def _acquire_lock_file():
        """建立跨 process 的 .lock 檔。回傳 (可繼續寫入, 是否由本次建立)。"""
        if not _remove_stale_lock_file(lock_file, force=force):
            logging.warning(
                f"[BatchUpdate] 偵測到有效 lock 檔 {lock_file}（另一程式正在寫入），延後寫入 "
                f"（記憶體仍保留 {len(updates_to_write)} 筆）。"
            )
            return False, False
        try:
            with open(lock_file, 'x') as f:
                f.write(str(os.getpid()))
            return True, True
        except FileExistsError:
            if force and _remove_stale_lock_file(lock_file, force=True):
                try:
                    with open(lock_file, 'x') as f:
                        f.write(str(os.getpid()))
                    return True, True
                except FileExistsError:
                    logging.warning(f"[BatchUpdate] 強制模式下仍無法取得 {lock_file}。")
                    return False, False
            logging.warning(f"[BatchUpdate] 無法取得 {lock_file}，延後 Done 批次寫入。")
            return False, False
        except Exception as e:
            logging.warning(f"[BatchUpdate] 無法建立 .lock 檔: {e}")
            return True, False

    def _do_write_once():
        global _batch_permission_warned, _last_batch_write_error
        try:
            try:
                wb = openpyxl.load_workbook(file_name)
            except Exception as load_err:
                if _is_corrupt_xlsx_error(load_err):
                    _last_batch_write_error = f"{file_name} 已損毀（xlsx zip 缺少 [Content_Types].xml），不是 lock 問題"
                    logging.error(_last_batch_write_error)
                    if _try_restore_corrupt_xlsx(file_name):
                        wb = openpyxl.load_workbook(file_name)
                    else:
                        return False
                else:
                    raise

            _sheet = _bulk_sheet()
            if _sheet not in wb.sheetnames:
                logging.error(f"檔案 {file_name} 中找不到 '{_sheet}' 工作表，無法寫入。將丟棄本次待寫入項目以免無限重試。")
                wb.close()
                with _pending_done_lock:
                    for idx in updates_to_write:
                        _pending_done_updates.pop(idx, None)
                return True

            ws = wb[_sheet]

            done_col = None
            for col_idx, cell in enumerate(ws[1], 1):
                if str(cell.value).strip() == 'Done':
                    done_col = col_idx
                    break

            if done_col is None:
                header_vals = [cell.value for cell in ws[1]]
                if (
                    len(header_vals) >= 4
                    and header_vals[0] == 'Video URL'
                    and header_vals[1] == 'TMDB URL'
                    and header_vals[2] == 'Split'
                ):
                    ws.cell(row=1, column=4).value = 'Done'
                    done_col = 4
                    logging.warning("偵測到 Bulk 表頭第 4 欄不是 'Done'（可能曾被誤寫），已還原為 'Done'。")

            if done_col is None:
                logging.error("找不到 'Done' 欄位，無法批次更新。將丟棄本次待寫入項目以免無限重試。")
                wb.close()
                with _pending_done_lock:
                    for idx in updates_to_write:
                        _pending_done_updates.pop(idx, None)
                return True

            batch_count = 0
            written_indices = []

            for row_index, done_value in updates_to_write.items():
                if row_index is None or row_index < 0:
                    logging.warning(f"忽略無效 row_index={row_index}，避免覆蓋表頭。")
                    written_indices.append(row_index)
                    continue
                target_row = row_index + 2
                if target_row < 2:
                    logging.warning(f"row_index {row_index} 對應 Excel 行 {target_row}，會覆蓋表頭，已跳過。")
                    written_indices.append(row_index)
                    continue
                if target_row <= ws.max_row:
                    ws.cell(row=target_row, column=done_col).value = done_value
                    batch_count += 1
                    written_indices.append(row_index)
                else:
                    logging.warning(f"row_index {row_index} (Excel行 {target_row}) 超出目前最大行數 {ws.max_row}，跳過此更新")
                    written_indices.append(row_index)

            if batch_count > 0:
                _atomic_save_workbook(wb, file_name)
                logging.info(f"成功批次寫入 {batch_count} 個 Done 值到 {file_name}。")
                _batch_permission_warned = False
                _last_batch_write_error = None
            if written_indices:
                with _pending_done_lock:
                    for idx in written_indices:
                        _pending_done_updates.pop(idx, None)

            wb.close()
            return True

        except PermissionError:
            _last_batch_write_error = f"無法訪問 {file_name}（可能被 Excel 佔用）"
            if not _batch_permission_warned or force:
                pending_n = len(updates_to_write)
                logging.warning(
                    f"批次寫入暫停：{_last_batch_write_error}。"
                    f"{pending_n} 筆 Done 值保留在記憶體，關閉 Excel 後將自動寫入。"
                )
                _batch_permission_warned = True
            return False
        except Exception as e:
            if _is_corrupt_xlsx_error(e):
                _last_batch_write_error = f"{file_name} 已損毀（xlsx zip 不完整），不是 lock 問題"
            else:
                _last_batch_write_error = str(e)
            logging.error(f"批次寫入 Done 值失敗: {e}\n{traceback.format_exc()}")
            return False

    # 【2026-11 I2】先等程式內部的 EXCEL_LOCK（最多 lock_wait 秒），取得後才建立 .lock 檔，
    # 避免「建立 lock 檔 → 搶不到 EXCEL_LOCK → 刪除」每 3 秒循環（lock 檔閃現 + 大量警告）。
    lock_wait = 15 if force else 20
    for attempt in range(max_attempts):
        lock_acquired = False
        created_lock_file = False
        try:
            if _excel_lock is not None:
                t0 = time.time()
                lock_acquired = _excel_lock.acquire(blocking=True, timeout=lock_wait)
                if not lock_acquired:
                    logging.warning(
                        f"[BatchUpdate] 程式內部讀寫鎖 (EXCEL_LOCK) 忙碌超過 {lock_wait} 秒"
                        f"（與 Excel 無關），延後 Done 批次寫入，記憶體仍保留 {len(updates_to_write)} 筆。"
                    )
                    if not force:
                        return False
                    continue
                waited = time.time() - t0
                if waited >= 5:
                    logging.info(f"[BatchUpdate] 等候程式內部讀寫鎖 {waited:.1f} 秒後開始寫入 Done 值。")
                # 等候期間可能有新 Done 值，取得鎖後重新取快照
                with _pending_done_lock:
                    if not _pending_done_updates:
                        return True
                    updates_to_write = _pending_done_updates.copy()
            ok, created_lock_file = _acquire_lock_file()
            if not ok:
                return False
            if _do_write_once():
                return True
        finally:
            if created_lock_file:
                try:
                    os.remove(lock_file)
                except OSError:
                    pass
            if lock_acquired:
                _excel_lock.release()
        if attempt < max_attempts - 1:
            time.sleep(1)
    return False

def update_done_value_with_lock_check(row_index, video_url, tmdb_url, value, file_name='data/bulk_videos.xlsx'):
    """
    更新指定行的 Done 值。
    注意：現在建議使用 update_done_value_fast 進行批次更新以提高效能。
    此函數保留作為單一、立即更新的選項。
    """
    import threading
    def background_update():
        log = logging.getLogger(__name__)
        lock_file = file_name + ".lock"
        output_filename = os.path.basename(file_name)
        waiting_message_logged = False
        while True:
            try:
                lock_fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.close(lock_fd)
                if waiting_message_logged:
                    log.info(f"檢測到 '{output_filename}' 已可用，現在開始寫入資料")
                break
            except FileExistsError:
                if not waiting_message_logged:
                    log.warning(f"'{output_filename}' 正在使用中（可能被其他腳本鎖定），現在等待其完成...")
                    waiting_message_logged = True
                time.sleep(1)
            except Exception as e:
                log.error(f"獲取檔案鎖 '{lock_file}' 時發生未知錯誤: {e}")
                return

        try:
            check_create_bulk_excel(file_name)
            update_done_value(row_index, video_url, tmdb_url, value, file_name=file_name)
        except Exception as e:
            log.error(f"在持有鎖的情況下更新 '{file_name}' 時發生非預期的錯誤: {e}\n{traceback.format_exc()}")
        finally:
            try:
                os.remove(lock_file)
            except OSError as e:
                log.error(f"釋放檔案鎖 '{lock_file}' 時失敗: {e}")

    thread = threading.Thread(target=background_update, name="DoneUpdateThread", daemon=False)
    thread.start()
    logging.info(f"已啟動背景執行緒來更新 Done 值 (row_index={row_index}, value={value})，主程序繼續執行。")

def check_and_release_excel_control(file_name):
    """檢查並釋放 Excel 文件的進程鎖定"""
    logger = logging.getLogger(__name__)
    max_retries = 3
    for attempt in range(max_retries):
        try:
            with open(file_name, 'ab'):
                pass
            return
        except PermissionError:
            logger.warning(f"無法訪問 {file_name}，可能被 Excel 占用，重試 {attempt + 1}/{max_retries}")
            try:
                if importlib.util.find_spec("win32com") is None:
                    logger.error("未安裝 pywin32 模組，無法釋放 Excel 進程")
                    raise ImportError("需要 pywin32 模組，請執行 'pip install pywin32'")
                import win32com.client
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                for wb in excel.Workbooks:
                    if wb.FullName == os.path.abspath(file_name):
                        wb.Close(SaveChanges=False)
                        logger.info(f"已關閉 Excel 文件: {file_name}")
                excel.Quit()
                time.sleep(1)
            except Exception as e:
                logger.error(f"釋放 Excel 進程失敗: {e}")
                if attempt == max_retries - 1:
                    raise PermissionError(f"無法釋放 {file_name} 的進程鎖定")
            time.sleep(2)

def release_excel_control():
    """釋放所有 Excel 進程的控制權"""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = False  # 嘗試設置 Visible 屬性
        except Exception as e:
            logging.warning(f"無法設置 Excel Visible 屬性: {e}")
        for wb in excel.Workbooks:
            wb.Close(SaveChanges=False)
        excel.Quit()
        logging.info("已關閉所有 Excel 進程")
        time.sleep(1)
    except ImportError as e:
        logging.error(f"未安裝 pywin32 模組，無法釋放 Excel 進程: {e}")
        logging.info("請執行 'pip install pywin32' 以安裝必要模組")
    except Exception as e:
        logging.error(f"釋放 Excel 控制權時發生錯誤: {e}")
    finally:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass

def _load_and_cache_sheet(mapping_file, sheet_name):
    """
    內部輔助函數，用於讀取、處理並緩存 tw_hk_name.xlsx 中的單個 sheet。
    如果 sheet 內容已在緩存中，則直接返回緩存內容，否則從檔案讀取並存入緩存。
    """
    global _tw_hk_name_cache
    if sheet_name in _tw_hk_name_cache:
        return _tw_hk_name_cache[sheet_name]

    try:
        logging.info(f"正在讀取並緩存 '{mapping_file}' 的 '{sheet_name}' sheet...")
        df = pd.read_excel(mapping_file, sheet_name=sheet_name, engine='openpyxl')
        df.dropna(subset=['TW Name', 'HK Name'], inplace=True)
        df['TW Name'] = df['TW Name'].astype(str).str.strip()
        df['HK Name'] = df['HK Name'].astype(str).str.strip()
        
        sheet_dict = dict(zip(df['TW Name'], df['HK Name']))
        sorted_keys = sorted(sheet_dict.keys(), key=len, reverse=True)
        sorted_sheet_dict = {k: sheet_dict[k] for k in sorted_keys}
        
        _tw_hk_name_cache[sheet_name] = sorted_sheet_dict
        _tw_hk_name_last_good[sheet_name] = sorted_sheet_dict
        logging.info(f"成功載入並緩存 '{sheet_name}' sheet，包含 {len(sorted_sheet_dict)} 條規則。")
        return sorted_sheet_dict
    except ValueError as e:
        # sheet 真的不存在（例如用家刪除/改名）→ 維持原有行為：空規則
        logging.warning(f"'{sheet_name}' Sheet 在 {mapping_file} 中不存在或讀取失敗。")
        _tw_hk_name_cache[sheet_name] = {}
        return {}
    except Exception as e:
        # 檔案暫時無法讀取（Excel 儲存中 / 被取代中 / 暫時損毀）→ 沿用上一次成功內容
        previous = _tw_hk_name_last_good.get(sheet_name)
        if previous is not None:
            logging.warning(f"讀取 '{sheet_name}' sheet 失敗 ({e})，沿用上一次成功載入的 {len(previous)} 條規則。")
            # 不寫入 _tw_hk_name_cache，下次使用時會再嘗試讀取檔案
            return previous
        logging.error(f"讀取並緩存 '{sheet_name}' sheet 時發生錯誤: {e}")
        _tw_hk_name_cache[sheet_name] = {}
        return {}

def clear_tw_hk_name_cache():
    """
    清除 tw_hk_name.xlsx 的內容緩存。
    """
    global _tw_hk_name_cache
    if _tw_hk_name_cache:
        logging.info("已清除 'tw_hk_name.xlsx' 的緩存，將在下次使用時重新載入。")
        _tw_hk_name_cache.clear()

def replace_tw_hk_names(
    input_file='data/video_detail.xlsx',
    mapping_file='data/tw_hk_name.xlsx',
    title=None,
    description=None,
    return_strings=False,
    context_title_for_sheet_selection=None,
    context_description_for_sheet_selection=None
):
    """
    根據 tw_hk_name.xlsx 中的映射替換 video_detail.xlsx 中的 TW 名稱為 HK 名稱，
    或直接對字串進行替換。
    """
    title_to_replace_is_empty = not (title and isinstance(title, str) and title.strip())
    description_to_replace_is_empty = not (description and isinstance(description, str) and description.strip())

    if return_strings and title_to_replace_is_empty and description_to_replace_is_empty:
        logging.debug("replace_tw_hk_names (字串模式): 傳入的 title 和 description 均為空或無效，不執行譯名轉換。")
        return title if title is not None else "", description if description is not None else ""

    if not os.path.exists(mapping_file):
        logging.warning(f"譯名映射檔案 '{mapping_file}' 不存在。")
        try:
            mapping_dir = os.path.dirname(mapping_file)
            if mapping_dir and not os.path.exists(mapping_dir):
                os.makedirs(mapping_dir)

            wb = openpyxl.Workbook()
            ws_default = wb.active
            ws_default.title = "Default"
            ws_default.append(["TW Name", "HK Name"])
            ws_default.append(["洛基", "洛奇"])
            ws_default.append(["汽車人", "博派"])
            ws_default.append(["霸天虎", "狂派"])
            ws_default.append(["擎天柱", "柯博文"])
            ws_default.append(["# 這是一般情況下使用的詞語替換", "#"])
            ws_movie = wb.create_sheet("Movie name")
            ws_movie.append(["TW Name", "HK Name"])
            ws_movie.append(["侏羅紀公園", "侏羅紀公園"])
            ws_movie.append(["# 這是僅在處理描述/簡介時，優先於Default的電影名稱替換", "#"])
            ws_gundam = wb.create_sheet("Gundam")
            ws_gundam.append(["TW Name", "HK Name", "Keyword"])
            ws_gundam.append(["鋼彈", "高達", "Gundam"])
            ws_gundam.append(["# 這是當標題或描述包含關鍵詞(Keyword)時，會優先使用的規則表", "#"])
            ws_protected = wb.create_sheet("Protected Name")
            ws_protected.append(["Protected Name"])
            ws_protected.append(["岩"])
            ws_protected.append(["占"])
            ws_protected.append(["里"])
            ws_protected.append(["# 此 sheet 中的人名/字詞在處理 'People' sheet 時不會被任何字典轉換", "#"])
            try:
                from others.ai_s2t import TRIGGER_HEADERS, DEFAULT_TRIGGERS
                ws_ai = wb.create_sheet("AI Convert")
                ws_ai.append(TRIGGER_HEADERS)
                for trig, opts, ex in DEFAULT_TRIGGERS:
                    ws_ai.append([trig, opts, ex, "Y"])
            except Exception as ai_sheet_e:
                logging.warning(f"建立 'AI Convert' sheet 失敗: {ai_sheet_e}")

            wb.save(mapping_file)
            logging.info(f"已創建範例檔案 '{mapping_file}'。請填寫後重新運行。本次將不執行譯名轉換。")
        except Exception as create_e:
            logging.error(f"創建範例檔案 '{mapping_file}' 失敗: {create_e}")

        if return_strings:
            return title if title is not None else "", description if description is not None else ""
        return

    try:
        sheet_select_title_ctx = title
        sheet_select_desc_ctx = description

        if return_strings:
            if context_title_for_sheet_selection is not None:
                sheet_select_title_ctx = context_title_for_sheet_selection
            if context_description_for_sheet_selection is not None:
                sheet_select_desc_ctx = context_description_for_sheet_selection
        elif (title is None and description is None):
            source_title_from_file = None
            source_description_from_file = None
            max_retries = 1
            for attempt in range(max_retries):
                try:
                    title_df = pd.read_excel(input_file, sheet_name='Title', engine='openpyxl')
                    if not title_df.empty and 'TV Show Title' in title_df.columns and pd.notna(title_df['TV Show Title'].iloc[0]):
                        source_title_from_file = title_df['TV Show Title'].iloc[0]
                        if 'TV Show Description' in title_df.columns and pd.notna(title_df['TV Show Description'].iloc[0]):
                            source_description_from_file = title_df['TV Show Description'].iloc[0]
                        break
                    movie_df = pd.read_excel(input_file, sheet_name='Movies', engine='openpyxl')
                    if not movie_df.empty and 'Movie Title' in movie_df.columns and pd.notna(movie_df['Movie Title'].iloc[0]):
                        source_title_from_file = movie_df['Movie Title'].iloc[0]
                        if 'Movie Description' in movie_df.columns and pd.notna(movie_df['Movie Description'].iloc[0]):
                            source_description_from_file = movie_df['Movie Description'].iloc[0]
                        break
                except Exception as e_read_ctx:
                    logging.debug(f"從 {input_file} 讀取 sheet 選擇上下文失敗 (嘗試 {attempt+1}): {e_read_ctx}")
                    if attempt < max_retries - 1: time.sleep(0.5)
                    else: logging.warning(f"無法從 {input_file} 讀取有效的 title/description 來選擇映射表。")
            sheet_select_title_ctx = source_title_from_file
            sheet_select_desc_ctx = source_description_from_file

        if not (sheet_select_title_ctx and str(sheet_select_title_ctx).strip()) and \
           not (sheet_select_desc_ctx and str(sheet_select_desc_ctx).strip()):
            logging.info("用於選擇譯名表 Sheet 的上下文 title 和 description 均為空，將使用 Default 規則。")
            selected_mapping_sheet = "Default"
        else:
            selected_mapping_sheet = select_tw_hk_sheet(sheet_select_title_ctx, sheet_select_desc_ctx)

        title_map_dict = {}
        description_map_dict = {}

        if selected_mapping_sheet == 'Default':
            title_map_dict = _load_and_cache_sheet(mapping_file, 'Default')
            movie_name_map = _load_and_cache_sheet(mapping_file, 'Movie name')
            default_map = _load_and_cache_sheet(mapping_file, 'Default')
            description_map_dict = {**default_map, **movie_name_map}
        else:
            keyword_map_dict = _load_and_cache_sheet(mapping_file, selected_mapping_sheet)
            title_map_dict = keyword_map_dict
            description_map_dict = keyword_map_dict
        
        logging.info(f"譯名轉換 - 上下文: title='{str(sheet_select_title_ctx)[:50]}...', desc='{str(sheet_select_desc_ctx)[:50]}...'. 已選擇並準備使用 Sheet 規則: '{selected_mapping_sheet}' (及適用的 'Default'/'Movie name')。")

        sorted_title_items = title_map_dict.items()
        sorted_description_items = description_map_dict.items()

        if return_strings:
            replaced_title_val = title if isinstance(title, str) else ''
            title_changes = []
            if replaced_title_val and sorted_title_items:
                for tw_name, hk_name in sorted_title_items:
                    if tw_name and tw_name in replaced_title_val:
                        new_val = replaced_title_val.replace(tw_name, hk_name)
                        if new_val != replaced_title_val:
                            title_changes.append(f"'{tw_name}' -> '{hk_name}'")
                            replaced_title_val = new_val
            if title_changes:
                logging.info(f"譯名轉換 (標題/名稱): {', '.join(title_changes)}")

            replaced_description_val = description if isinstance(description, str) else ''
            desc_changes = []
            if replaced_description_val and sorted_description_items:
                for tw_name, hk_name in sorted_description_items:
                    if tw_name and tw_name in replaced_description_val:
                        new_val = replaced_description_val.replace(tw_name, hk_name)
                        if new_val != replaced_description_val:
                            desc_changes.append(f"'{tw_name}' -> '{hk_name}'")
                            replaced_description_val = new_val
            if desc_changes:
                logging.info(f"譯名轉換 (描述/簡介): {', '.join(desc_changes)}")
            
            return replaced_title_val, replaced_description_val

        changes_made_in_excel = False
        if not sorted_title_items and not sorted_description_items:
            logging.info(f"無有效的譯名轉換規則，不對 Excel 文件 {input_file} 進行操作。")
            return

        for sheet_to_process in ['Title', 'Seasons', 'Episodes', 'Movies', 'People']:
            # 觸發條件：遇 PermissionError 時提示用戶並無限重試
            permission_logged = False
            while True:
                try:
                    df = pd.read_excel(input_file, sheet_name=sheet_to_process, engine='openpyxl')
                    if df.empty:
                        break # break loop to next sheet

                    sheet_changed = False
                    
                    title_columns_map = {
                        'Title': ['TV Show Title'], 'Seasons': ['Season Name'], 'Episodes': ['Episode Title'],
                        'Movies': ['Movie Title'], 'People': ['Name']
                    }
                    desc_columns_map = {
                        'Title': ['TV Show Description'], 'Seasons': ['Season Description'], 'Episodes': ['Episode Description'],
                        'Movies': ['Movie Description'], 'People': ['Biography']
                    }

                    if sheet_to_process in title_columns_map and sorted_title_items:
                        for col_name in title_columns_map[sheet_to_process]:
                            if col_name in df.columns:
                                original_col = df[col_name].copy()
                                for tw_name, hk_name in sorted_title_items:
                                    df[col_name] = df[col_name].apply(lambda x: x.replace(tw_name, hk_name) if isinstance(x, str) and tw_name in x else x)
                                if not df[col_name].equals(original_col):
                                    sheet_changed = True
                                    logging.info(f"在 {sheet_to_process} 的 '{col_name}' 欄位中應用了標題/名稱的譯名轉換規則。")
                    
                    if sheet_to_process in desc_columns_map and sorted_description_items:
                        for col_name in desc_columns_map[sheet_to_process]:
                            if col_name in df.columns:
                                original_col = df[col_name].copy()
                                for tw_name, hk_name in sorted_description_items:
                                    df[col_name] = df[col_name].apply(lambda x: x.replace(tw_name, hk_name) if isinstance(x, str) and tw_name in x else x)
                                if not df[col_name].equals(original_col):
                                    sheet_changed = True
                                    logging.info(f"在 {sheet_to_process} 的 '{col_name}' 欄位中應用了描述/簡介的譯名轉換規則。")

                    if sheet_changed:
                        changes_made_in_excel = True
                        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df.to_excel(writer, sheet_name=sheet_to_process, index=False)
                        logging.info(f"已更新 {input_file} 中的 {sheet_to_process} sheet (TW->HK譯名轉換)。")
                    
                    break # Success, break loop
                except PermissionError:
                    if not permission_logged:
                        logging.warning(f"無法訪問 {input_file} (用於寫入 {sheet_to_process} sheet)，可能正被其他程序占用。請關閉該檔案以繼續... (程式將自動每秒重試)")
                        permission_logged = True
                    time.sleep(1)
                except Exception as e:
                    logging.warning(f"處理 {input_file} 的 {sheet_to_process} sheet 時發生錯誤: {e}。跳過此 sheet。")
                    break

        if not changes_made_in_excel:
            logging.info(f"在 {input_file} 的所有相關 sheets 中未檢測到需要進行譯名轉換的內容。")

    except Exception as e:
        logging.error(f"執行 TW 到 HK 名稱替換時發生頂層錯誤: {e}\n{traceback.format_exc()}")
        if return_strings:
            return title if title is not None else "", description if description is not None else ""

def get_done_value(row_index, video_url, tmdb_url, sheet_name=None, file_name='data/bulk_videos.xlsx'):
    """從指定 sheet 獲取 Done 值"""
    sheet_name = sheet_name or _bulk_sheet()
    try:
        df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')
        if row_index < len(df):
            row = df.iloc[row_index]
            if (str(row.get('Video URL', '')).strip() == str(clean_url(video_url)).strip() and
                str(row.get('TMDB URL', '')).strip() == str(clean_url(tmdb_url)).strip()):
                done_value = row.get('Done', '')
                return str(done_value).strip() if pd.notna(done_value) else ''
        return ''
    except Exception as e:
        logging.error(f"獲取 Done 值時發生錯誤: {e}")
        return ''

def update_done_value(row_index, video_url, tmdb_url, done_value, sheet_name=None, file_name='data/bulk_videos.xlsx'):
    """
    更新指定行的 Done 值。
    優先使用 row_index，回退使用 URL 搜尋。
    """
    sheet_name = sheet_name or _bulk_sheet()
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    # 觸發條件：遇 PermissionError 時提示用戶並無限重試
    permission_logged = False
    while True:
        try:
            df_for_match = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl', usecols=['Video URL', 'TMDB URL', 'Done'])

            normalized_video_url = clean_url(video_url) if pd.notna(video_url) else ''
            normalized_tmdb_url = clean_url(tmdb_url) if pd.notna(tmdb_url) else ''
            
            target_index = -1
            
            if 0 <= row_index < len(df_for_match):
                row_data = df_for_match.iloc[row_index]
                row_video_url_clean = clean_url(row_data.get('Video URL', '')) if pd.notna(row_data.get('Video URL', '')) else ''
                row_tmdb_url_clean = clean_url(row_data.get('TMDB URL', '')) if pd.notna(row_data.get('TMDB URL', '')) else ''
                
                url_match_at_index = False
                if normalized_tmdb_url and row_tmdb_url_clean == normalized_tmdb_url:
                    url_match_at_index = True
                elif normalized_video_url and row_video_url_clean == normalized_video_url:
                     url_match_at_index = True
                elif not normalized_tmdb_url and not normalized_video_url and not row_tmdb_url_clean and not row_video_url_clean:
                    url_match_at_index = True

                if url_match_at_index:
                    target_index = row_index
                    logging.info(f"成功使用提供的行索引 {row_index} 定位目標行。")
                else:
                    logging.warning(f"提供的行索引 {row_index} 的 URL 與傳入的 URL 不匹配。將嘗試透過 URL 搜尋。")
            else:
                 logging.warning(f"提供的行索引 {row_index} 無效。將嘗試透過 URL 搜尋。")

            if target_index == -1:
                df_for_match['Video URL Clean'] = df_for_match['Video URL'].astype(str).replace('nan', '').apply(clean_url)
                df_for_match['TMDB URL Clean'] = df_for_match['TMDB URL'].astype(str).replace('nan', '').apply(clean_url)

                match_condition = pd.Series([False] * len(df_for_match))
                if normalized_tmdb_url:
                    match_condition = (df_for_match['TMDB URL Clean'] == normalized_tmdb_url)
                
                if not match_condition.any() and normalized_video_url:
                    match_condition = (df_for_match['Video URL Clean'] == normalized_video_url)
                
                if not normalized_tmdb_url and not normalized_video_url:
                    match_condition = (df_for_match['TMDB URL Clean'] == '') & (df_for_match['Video URL Clean'] == '')

                matched_indices = df_for_match.index[match_condition].tolist()

                if matched_indices:
                    if len(matched_indices) > 1:
                        if row_index in matched_indices:
                            target_index = row_index
                            logging.warning(f"透過 URL 找到多個匹配項，但包含原始索引，使用原始索引。")
                        else:
                            target_index = matched_indices[0]
                            logging.warning(f"透過 URL 找到多個匹配項。將更新第一個匹配項: {target_index}。")
                    else:
                        target_index = matched_indices[0]
                    logging.info(f"透過 URL 搜尋成功定位到行索引: {target_index}")
                else:
                    logging.error(f"無法透過行索引或 URL 找到匹配行。無法更新 Done 值。")
                    return

            if target_index != -1:
                logging.info(f"準備在行索引 {target_index} 更新 Done={done_value}")
                
                wb = openpyxl.load_workbook(file_name)
                ws = wb[sheet_name]
                
                header = [cell.value for cell in ws[1]]
                try:
                    done_col_index = header.index('Done') + 1
                except ValueError:
                    logging.error(f"在 '{sheet_name}' sheet 中找不到 'Done' 欄位。無法更新。")
                    wb.close()
                    return

                worksheet_row = target_index + 2
                
                if worksheet_row > ws.max_row:
                    logging.error(f"目標工作表行 {worksheet_row} 超出最大行 {ws.max_row}。無法更新 Done 值。")
                    wb.close()
                    return

                ws.cell(row=worksheet_row, column=done_col_index).value = done_value
                
                wb.save(file_name)
                wb.close()
                
                logging.info(f"成功更新 {file_name} (行索引 {target_index}, Done={done_value})")
            
            return

        except PermissionError:
            if not permission_logged:
                logging.warning(f"無法寫入 '{file_name}'，檔案可能已被其他程式鎖定。請關閉該檔案以繼續... (程式將自動每秒重試)")
                permission_logged = True
            time.sleep(1)

        except Exception as e:
            logging.error(f"更新 Done 值時發生未知錯誤: {e}\n{traceback.format_exc()}")
            raise