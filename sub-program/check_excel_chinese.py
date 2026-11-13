import asyncio
import logging
import os
import re
import sys
import time
import httpx
import pandas as pd
# --- 路徑設定 ---
# 獲取當前腳本的目錄 (sub-program)
script_dir = os.path.dirname(os.path.abspath(__file__))
# 獲取上層目錄 (專案根目錄)
project_root = os.path.dirname(script_dir)
# 將專案根目錄添加到 Python 的模組搜索路徑中
sys.path.append(project_root)
# 現在可以從 'others' 套件中導入模組
from others import credentials
from others import excel_handler
# --- 設定日誌 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# --- API 和並行設定 ---
MAX_CONCURRENT_REQUESTS = 15
TMDB_API_BASE_URL = "https://api.themoviedb.org/3"
def get_tmdb_api_key(configs):
    """
    從設定檔或使用者輸入獲取 TMDB API 金鑰。
    """
    if 'tmdb_api_v3' not in configs or not configs['tmdb_api_v3']:
        api_key = input("請輸入您的 TMDB API Key (v3): ")
        save_config = input("是否希望儲存此 API Key 留待下次使用? (y/n): ")
        if save_config.lower() == 'y':
            configs['tmdb_api_v3'] = api_key
            credentials.save_configs(configs)
    else:
        api_key = configs['tmdb_api_v3']
    return api_key
def parse_tmdb_url(url):
    """從 TMDB URL 中解析出影視類型和 ID。"""
    match = re.search(r'themoviedb\.org/(movie|tv)/(\d+)', url)
    if match:
        video_type = match.group(1)
        video_id = match.group(2)
        return video_type, video_id
    return None, None
async def check_chinese_translation_api(session, semaphore, video_type, video_id, api_key):
    """使用 API 檢查影片是否有 zh-HK, zh-TW, zh-CN 翻譯。"""
    url = f"{TMDB_API_BASE_URL}/{video_type}/{video_id}/translations?api_key={api_key}"
    try:
        async with semaphore:
            response = await session.get(url, timeout=20)
            response.raise_for_status()
            data = response.json()
            translations = data.get('translations', [])
            for t in translations:
                iso_639_1 = t.get('iso_639_1')
                iso_3166_1 = t.get('iso_3166_1')
                if iso_639_1 == 'zh' and iso_3166_1 in ['HK', 'TW', 'CN']:
                    logging.info(f"✔️ 找到指定中文翻譯 (zh-{iso_3166_1}): {video_type}/{video_id}")
                    return True
        logging.warning(f"❌ 未找到 zh-HK, zh-TW, zh-CN 翻譯: {video_type}/{video_id}")
        return False
    except httpx.HTTPStatusError as e:
        logging.error(f"檢查翻譯 {video_type}/{video_id} 時 API 錯誤: {e.response.status_code}")
    except Exception as e:
        logging.error(f"檢查翻譯 {video_type}/{video_id} 時出錯: {e}")
    return False
async def main():
    """主執行函數"""
    configs = credentials.load_configs()
    api_key = get_tmdb_api_key(configs)
    if not api_key:
        logging.error("未提供 TMDB API Key，無法繼續。")
        return
    output_file = os.path.join(project_root, 'data', 'bulk_videos.xlsx')
    try:
        df = pd.read_excel(output_file, sheet_name='Bulk')
    except ValueError:
        logging.info("Bulk sheet 不存在，無需處理。")
        return
    except FileNotFoundError:
        logging.info("檔案不存在，無需處理。")
        return
    except Exception as e:
        logging.error(f"讀取 Excel 時發生錯誤: {e}")
        return
    if 'Video URL' not in df.columns:
        logging.error("Bulk sheet 缺少 'Video URL' 欄位，無法繼續。")
        return
    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)
    valid_indices = []
    video_items = []
    for index, row in df.iterrows():
        url = row.get('Video URL')
        if pd.isna(url) or not url:
            logging.warning(f"第 {index+1} 列 Video URL 為空，將刪除此列。")
            continue
        video_type, video_id = parse_tmdb_url(url)
        if not video_type or not video_id:
            logging.warning(f"無法解析 URL: {url} (第 {index+1} 列)，將刪除此列。")
            continue
        video_items.append((index, video_type, video_id))
    if not video_items:
        logging.info("沒有有效的 URL 需要檢查，將清空 Bulk sheet。")
        new_df = pd.DataFrame(columns=df.columns)
    else:
        async with httpx.AsyncClient() as session:
            check_tasks = [check_chinese_translation_api(session, semaphore, v_type, v_id, api_key) for _, v_type, v_id in video_items]
            results = await asyncio.gather(*check_tasks)
            for i, has_translation in enumerate(results):
                if has_translation:
                    valid_indices.append(video_items[i][0])
        if valid_indices:
            new_df = df.loc[valid_indices].reset_index(drop=True)
            logging.info(f"檢查完成。保留 {len(new_df)} 列具有指定中文翻譯的記錄。")
        else:
            new_df = pd.DataFrame(columns=df.columns)
            logging.info("沒有找到任何具有指定中文翻譯的記錄，將清空 Bulk sheet。")
    # 寫回 Excel，只替換 Bulk sheet
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            new_df.to_excel(writer, sheet_name='Bulk', index=False)
        logging.info(f"成功更新 {output_file} 的 Bulk sheet。")
    except PermissionError:
        logging.error(f"寫入 Excel 失敗：權限錯誤。請確保 '{output_file}' 未被其他程式開啟。")
    except Exception as e:
        logging.error(f"寫入 Excel 時發生錯誤: {e}")
    logging.info("程式執行完畢。")
if __name__ == '__main__':
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())