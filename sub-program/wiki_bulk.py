import json
import logging
import re
import httpx
import asyncio
import sys
import urllib.parse
import os
from bs4 import BeautifulSoup
from urllib.parse import urlparse
import pandas as pd
import time # (新增) 為了 time.sleep
import traceback # (新增) 為了記錄詳細錯誤

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 並行請求設定
MAX_CONCURRENT_REQUESTS = 15

# 全域目標標題關鍵字列表
TARGET_TITLE_KEYWORDS = ['中文劇名', '韓文劇名', '原名', '片名', '劇名', '剧名', '名稱', '作品名', '작품명', '中文片名', '韓文片名', '電影', '제목']

def load_config():
    """讀取 configs.json 並檢查 TMDB v3 API 鍵"""
    config_path = '../configs.json'
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        api_key = config.get('tmdb_api_v3')
        if not api_key:
            logging.error(f"{config_path} 中缺少有效的 TMDB v3 API 鍵")
            return None
        return api_key
    except FileNotFoundError:
        logging.error(f"找不到 {config_path} 文件")
        return None
    except json.JSONDecodeError:
        logging.error(f"{config_path} 格式錯誤")
        return None

def validate_wiki_url(url):
    """驗證 URL 是否為 Wikipedia 頁面"""
    parsed_url = urlparse(url)
    if parsed_url.netloc not in ('zh.wikipedia.org', 'ko.wikipedia.org'):
        logging.error(f"輸入的 URL 不是有效的 Wikipedia 頁面: {url}")
        return False
    return True

def check_page_format(soup, url):
    """
    檢查頁面是否包含任何看起來有標題欄位的表格。
    返回一個布林值表示頁面是否可能包含可抓取的資料。
    """
    tables = soup.find_all('table')
    for table_idx, table in enumerate(tables):
        current_header_texts = []
        
        # 優先使用 <thead> 中的表頭
        thead = table.find('thead')
        if thead:
            thead_rows = thead.find_all('tr') # 移除 recursive=False
            if thead_rows:
                header_tags_in_row = thead_rows[0].find_all(['th', 'td']) # 移除 recursive=False
                # 在 check_page_format 中，我們更關心是否存在任一類型的表頭包含關鍵字
                current_header_texts = [h.get_text(strip=True) for h in header_tags_in_row]
        
        if not current_header_texts: # 如果 thead 中未找到有效表頭，嘗試表格的第一行
            all_table_rows = table.find_all('tr') # 移除 recursive=False
            if all_table_rows:
                first_row_tags = all_table_rows[0].find_all(['th', 'td']) # 移除 recursive=False
                current_header_texts = [h.get_text(strip=True) for h in first_row_tags]
        
        if not current_header_texts:
            logging.debug(f"Table {table_idx+1} in check_page_format: No clear headers found.")
            continue

        found_keywords_in_table = [kw for kw in TARGET_TITLE_KEYWORDS if kw in current_header_texts]
        
        if found_keywords_in_table:
            logging.info(f"頁面 {url} - check_page_format: 找到可能的表格 {table_idx+1}，偵測到的表頭文字: {current_header_texts}，包含關鍵字: {found_keywords_in_table}")
            return True
            
    logging.warning(f"頁面 {url} - check_page_format: 未找到包含任何目標標題關鍵字的表格。")
    return False

def check_create_bulk_excel(file_name, sheet_name='Bulk'):
    """檢查並創建 bulk_videos.xlsx，僅包含 Bulk 表"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    required_sheets = {
        sheet_name: ['Video URL', 'TMDB URL', 'Split', 'Done', '覆蓋劇集標題', '覆蓋劇集描述', '覆蓋分集標題', '覆蓋分集描述', '來源無資料時清空TMDB', '維基人名取代現有']
    }
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if not os.path.exists(file_name):
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    for sheet, columns in required_sheets.items():
                        pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet, index=False)
                logging.info(f"已創建新的 {file_name}，包含工作表: {list(required_sheets.keys())}")
            else:
                # 檢查 sheet 是否存在
                try:
                    all_sheets = pd.read_excel(file_name, sheet_name=None, engine='openpyxl')
                except FileNotFoundError:
                    # 另一種情況下的文件不存在
                    all_sheets = {}

                if sheet_name not in all_sheets:
                     with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                        pd.DataFrame(columns=required_sheets[sheet_name]).to_excel(writer, sheet_name=sheet_name, index=False)
                     logging.info(f"工作表 {sheet_name} 不存在於 {file_name}，已創建")
                else:
                    # 檢查欄位是否存在
                    df = all_sheets[sheet_name]
                    needs_update = False
                    for col in required_sheets[sheet_name]:
                        if col not in df.columns:
                            df[col] = ''
                            needs_update = True
                    
                    if needs_update:
                        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                        logging.info(f"已為 {file_name} 的 {sheet_name} 添加缺失的欄位")

            return  # 成功，退出循環
        except PermissionError:
            if attempt < max_retries - 1:
                logging.warning(f"無法訪問 {file_name}，可能正被其他程序占用，重試 {attempt + 1}/{max_retries}")
                input(f"請關閉 {file_name}，然後按 Enter 鍵繼續...")
            else:
                logging.error(f"超過最大重試次數 {max_retries}，無法訪問 {file_name}")
                raise
        except Exception as e:
            logging.error(f"創建或檢查 {file_name} 時發生錯誤: {e}")
            raise

def generate_title_variants(title):
    """生成標題變體，逐步移除空格後的詞和末尾數字"""
    variants = [title]
    current = title
    
    while ' ' in current:
        current = ' '.join(current.split(' ')[:-1]).strip()
        if current:
            variants.append(current)
    
    if re.search(r'\d+$', current):
        current = re.sub(r'\d+$', '', current).strip()
        if current:
            variants.append(current)
    
    return list(set(variants))

def clean_and_split_title(title):
    """清理標題中的備注、書名號並分拆多標題"""
    cleaned_title = re.sub(r'\*?\s*\[(註|注)\s*\d+\]', '', title).strip()
    cleaned_title = re.sub(r'[《》]', '', cleaned_title).strip()
    
    if re.search(r'[/\\]', cleaned_title):
        split_titles = [t.strip() for t in re.split(r'[/\\]', cleaned_title) if t.strip()]
        logging.debug(f"分拆標題: {cleaned_title} -> {split_titles}")
        return split_titles
    logging.debug(f"未分拆標題: {cleaned_title}")
    return [cleaned_title]

def extract_titles(soup, is_korean_wiki=False):
    """從所有包含相關表頭欄位的表格中提取標題。"""
    extracted_titles_list = []
    tables = soup.find_all('table')
    
    # This loop iterates through all tables on the page.
    for table_idx, table in enumerate(tables):
        logging.debug(f"Extract_titles: --- 開始處理表格 {table_idx + 1} (class: {table.get('class')}) ---")
        
        # Part 1: Identify the header row and its cells.
        header_tags_for_colspan_calc = [] 
        header_row_candidate = None
        is_header_from_thead = False

        thead = table.find('thead')
        if thead:
            thead_tr_list = thead.find_all('tr') 
            if thead_tr_list:
                header_row_candidate = thead_tr_list[0]
                is_header_from_thead = True
                logging.debug(f"表格 {table_idx + 1}: 找到 thead 中的表頭行。")
        
        if not header_row_candidate: 
            table_tr_list = table.find_all('tr') 
            if table_tr_list:
                header_row_candidate = table_tr_list[0]
                logging.debug(f"表格 {table_idx + 1}: 無 thead 表頭，嘗試使用表格的第一行作為表頭。")
            else:
                logging.debug(f"表格 {table_idx + 1}: 表格內無 tr 標籤。跳過此表格。")
                continue # Skip to the next table

        if header_row_candidate:
            header_tags_for_colspan_calc = header_row_candidate.find_all(['th', 'td']) 
        
        if not header_tags_for_colspan_calc:
            logging.debug(f"表格 {table_idx + 1}: 在識別的表頭行中未找到 'th' 或 'td' 標籤。跳過此表格。")
            continue
            
        current_header_texts_for_extraction = [h.get_text(strip=True) for h in header_tags_for_colspan_calc]
        logging.debug(f"表格 {table_idx + 1}: 偵測到的表頭文字: {current_header_texts_for_extraction}")

        # Part 2: Find which columns are relevant, prioritizing '原名'.
        current_table_title_columns_config = [] 
        original_title_header_index = -1
        # First, specifically look for '原名' (original name).
        for idx, header_text in enumerate(current_header_texts_for_extraction):
             if '原名' in header_text:
                 original_title_header_index = idx
                 break # Found the primary target, stop searching

        if original_title_header_index != -1:
            # If "原名" is found, only use that column.
            header_text = current_header_texts_for_extraction[original_title_header_index]
            current_table_title_columns_config.append((header_text, original_title_header_index))
            logging.debug(f"表格 {table_idx + 1}: 找到優先欄位 '原名' (索引 {original_title_header_index})。將只從此欄位提取。")
        else:
            # Fallback to original behavior if "原名" is not found.
            for idx, header_text in enumerate(current_header_texts_for_extraction):
                if header_text in TARGET_TITLE_KEYWORDS:
                    current_table_title_columns_config.append((header_text, idx))
            logging.debug(f"表格 {table_idx + 1}: 未找到 '原名' 欄位，將從以下欄位提取: {current_table_title_columns_config}")

        if not current_table_title_columns_config:
            logging.debug(f"表格 {table_idx + 1}: 標題欄位配置為空 (未找到目標關鍵字)。跳過此表格的標題提取。")
            continue
        
        logging.info(f"表格 {table_idx + 1}: 將使用以下欄位提取標題: {current_table_title_columns_config}")

        # Part 3: Find the data rows.
        data_rows = []
        tbody = table.find('tbody')
        if tbody:
            data_rows = tbody.find_all('tr', recursive=False) 
            logging.debug(f"表格 {table_idx + 1}: 從 <tbody> 中提取 {len(data_rows)} 行資料。")
        elif is_header_from_thead:
            all_table_trs = table.find_all('tr')
            thead_trs = thead.find_all('tr')
            if len(all_table_trs) > len(thead_trs):
                data_rows = [tr for tr in all_table_trs if tr not in thead_trs]
                logging.debug(f"表格 {table_idx + 1}: 有 thead ({len(thead_trs)}行) 但無 tbody。從 table 其餘 {len(data_rows)} 行提取資料。")
            else:
                logging.debug(f"表格 {table_idx + 1}: 有 thead 但無 tbody，且 table 中無額外 tr。")
        elif header_row_candidate:
            all_table_trs = table.find_all('tr')
            if len(all_table_trs) > 1:
                data_rows = all_table_trs[1:]
                logging.debug(f"表格 {table_idx + 1}: 無 thead/tbody。使用表格索引 1 開始的 {len(data_rows)} 行資料。")
            else:
                logging.debug(f"表格 {table_idx + 1}: 無 thead/tbody，且只有一行。無資料列。")
        else:
            logging.debug(f"表格 {table_idx + 1}: 無法確定資料列來源。")

        if not data_rows:
            logging.debug(f"表格 {table_idx + 1}: 未找到資料列。結束此表格處理。")
            continue

        # Part 4: ROW PROCESSING LOGIC to handle rowspan/colspan correctly.
        # Calculate total columns from header once per table
        header_col_count = sum(int(h.get('colspan', 1)) for h in header_tags_for_colspan_calc)
        logging.debug(f"表格 {table_idx + 1}: 計算出的表頭總欄數: {header_col_count}")

        for row_idx, row_element in enumerate(data_rows):
            all_cells_in_row = row_element.find_all(['th', 'td'], recursive=False)

            # Check for and skip rows that are likely notes/footer rows within the tbody
            if len(all_cells_in_row) == 1:
                single_cell = all_cells_in_row[0]
                if single_cell.name == 'td' and single_cell.get('colspan'):
                    # It's a single cell spanning multiple columns, a common pattern for notes.
                    # Check for more specific clues like the references list or "註釋" text.
                    note_cell_text = single_cell.get_text(strip=True)
                    if (single_cell.find('ol', class_='references') or 
                        '註釋' in note_cell_text or
                        'Notes' in note_cell_text):
                        
                        logging.debug(f"表格 {table_idx + 1}, 資料列 {row_idx + 1}: 偵測到註釋行 (colspan/references)，跳過。")
                        continue
            
            # Skip rows that are likely subheaders (all <th>) or empty
            if not all_cells_in_row or (len(row_element.find_all('td', recursive=False)) == 0 and len(all_cells_in_row) > 0):
                logging.debug(f"表格 {table_idx + 1}, 資料列 {row_idx + 1}: 行無內容，或可能為子標頭/分隔行，跳過。")
                continue

            # Calculate the effective width and offset for the current row to handle rowspans
            row_width = sum(int(c.get('colspan', 1)) for c in all_cells_in_row)
            row_offset = header_col_count - row_width
            
            if row_offset < 0:
                logging.warning(f"表格 {table_idx + 1}, 資料列 {row_idx + 1}: 計算出的行寬度 ({row_width}) 大於表頭寬度 ({header_col_count})。將偏移設為 0。")
                row_offset = 0

            # For each configured title column we want to extract from this row...
            for col_name_keyword, title_header_tag_index in current_table_title_columns_config:
                # Calculate the target conceptual index for this column based on header colspans
                conceptual_col_idx_start = 0
                for h_idx, h_tag in enumerate(header_tags_for_colspan_calc):
                    if h_idx == title_header_tag_index:
                        break
                    conceptual_col_idx_start += int(h_tag.get('colspan', 1))
                
                # Iterate through cells in the current row to find the one at the target conceptual position
                current_conceptual_pos = row_offset
                found_cell = None
                for cell_in_row in all_cells_in_row:
                    cell_width = int(cell_in_row.get('colspan', 1))
                    # Check if our target conceptual index falls within the column range covered by this cell
                    if current_conceptual_pos <= conceptual_col_idx_start < (current_conceptual_pos + cell_width):
                        found_cell = cell_in_row
                        break
                    current_conceptual_pos += cell_width

                # If a cell was found at the target position, extract its text
                if found_cell:
                    raw_title_text = ""
                    # Prioritize text from an 'a' tag if it exists and has text
                    link_tag = found_cell.find('a')
                    if link_tag and link_tag.get_text(strip=True):
                        raw_title_text = link_tag.get_text(strip=True)
                    else: 
                        raw_title_text = found_cell.get_text(strip=True)

                    # Handle multiple titles separated by <br>
                    if found_cell.find('br'):
                        cell_html_content = found_cell.encode_contents().decode('utf-8', 'replace')
                        parts_from_br = re.split(r'<br\s*/?>|\n', cell_html_content)
                        parts_text_from_br = [BeautifulSoup(part, 'html.parser').get_text(strip=True) for part in parts_from_br if BeautifulSoup(part, 'html.parser').get_text(strip=True)]
                        
                        for part_text in parts_text_from_br:
                            if part_text: 
                                cleaned_titles_from_part = clean_and_split_title(part_text)
                                extracted_titles_list.extend(cleaned_titles_from_part)
                                for t_split in cleaned_titles_from_part:
                                    logging.debug(f"表格 {table_idx+1} 資料列 {row_idx+1} 欄位 '{col_name_keyword}' (概念索引 {conceptual_col_idx_start}) (來自 <br> 拆分): 已提取標題: {t_split}")
                    elif raw_title_text: 
                        cleaned_titles = clean_and_split_title(raw_title_text)
                        extracted_titles_list.extend(cleaned_titles)
                        for t_direct in cleaned_titles:
                            logging.debug(f"表格 {table_idx+1} 資料列 {row_idx+1} 欄位 '{col_name_keyword}' (概念索引 {conceptual_col_idx_start}): 已提取標題: {t_direct}")
                    else:
                        logging.debug(f"表格 {table_idx+1} 資料列 {row_idx+1} 欄位 '{col_name_keyword}' (概念索引 {conceptual_col_idx_start}): 找到的儲存格為空或未提取到文字。")
                else:
                    logging.debug(f"表格 {table_idx + 1}, 資料列 {row_idx + 1}: 未能在行中定位到概念索引 {conceptual_col_idx_start} 對應的儲存格 (欄位: '{col_name_keyword}')。")
        
        logging.debug(f"Extract_titles: --- 完成處理表格 {table_idx + 1} ---")

    final_unique_titles = list(set(extracted_titles_list)) 
    logging.info(f"extract_titles 完成，總共提取到 {len(final_unique_titles)} 個獨立標題。")
    if not final_unique_titles:
        logging.warning("extract_titles 未提取到任何標題。")
    return final_unique_titles

def contains_chinese(text):
    """
    輔助函數：檢查一個字串是否包含任何中文字元，並排除包含日文假名的字串。
    - (使用者要求 2025-08-19) 修正將日文漢字誤判為中文的問題。
    - 實現方式：
    -   1. 首先，使用正規表示式檢查字串中是否存在日文的平假名或片假名。
    -      日文假名的 Unicode 範圍為 U+3040 到 U+30FF。
    -   2. 如果檢測到任何假名字符，則立即將該字串判定為「非中文」並返回 False。
    -   3. 只有在確認字串「不含」任何日文假名的前提下，才繼續檢查其是否包含 CJK 統一表意文字（漢字）。
    -      這樣可以有效過濾掉日文標題，同時正確識別純漢字標題。
    """
    if not text or not isinstance(text, str):
        return False
    
    # 步驟 1: 檢查是否有日文假名（平假名: \u3040-\u309f, 片假名: \u30a0-\u30ff）
    # 如果找到任何假名，就立即返回 False，因為中文文本不應包含這些字符。
    if re.search(r'[\u3040-\u30ff]', text):
        return False
        
    # 步驟 2: 在確認沒有日文假名後，再檢查是否包含中文字符（漢字）
    # \u4e00-\u9fff 是 CJK 統一表意文字的基本範圍
    return bool(re.search(r'[\u4e00-\u9fff]', text))

async def has_chinese_translation(session, semaphore, api_key, tmdb_url, api_base_url):
    """
    異步檢查指定的 TMDB URL 是否有中文內容。
    (使用者需求 2025-08-19) 採用更有效率的方式檢查影片是否有中文內容。
    (使用者需求 2025-09-02) 新增錯誤重試機制，如果 API 請求失敗，最多重試 100 次。
    """
    match = re.search(r'/(movie|tv)/(\d+)', tmdb_url)
    if not match:
        logging.warning(f"無法從 URL 中解析 media type 或 ID: {tmdb_url}。將假定其沒有翻譯並跳過。")
        return False

    media_type = match.group(1)
    item_id = match.group(2)
    
    url = f"{api_base_url}/{media_type}/{item_id}?api_key={api_key}&language=en-US&append_to_response=translations"
    
    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101):
        try:
            async with semaphore:
                response = await session.get(url, timeout=20)
                
                if response.status_code == 404:
                    logging.warning(f"❌ 檢查 {tmdb_url} 時項目不存在 (404)。")
                    return False

                response.raise_for_status()
                data = response.json()

            # --- 第一階段：檢查官方翻譯 ---
            translations = data.get('translations', {}).get('translations', [])
            for t in translations:
                lang_code = t.get("iso_639_1")
                country_code = t.get("iso_3166_1")
                if lang_code == 'zh' and country_code in ['HK', 'TW', 'CN']:
                    logging.info(f"✔️  URL {tmdb_url} 找到官方中文翻譯 ({lang_code}-{country_code})。")
                    return True

            # --- 第二階段：檢查標題 ---
            title = data.get('title') if media_type == 'movie' else data.get('name')
            original_title = data.get('original_title') if media_type == 'movie' else data.get('original_name')
            
            title_to_log = ""
            is_chinese = False
            if title and contains_chinese(title):
                is_chinese = True
                title_to_log = title
            elif original_title and contains_chinese(original_title):
                is_chinese = True
                title_to_log = original_title

            if is_chinese:
                logging.info(f"✔️  URL {tmdb_url} 在 en-US 資料中找到中文標題/原標題: '{title_to_log}'。")
                return True

            logging.warning(f"❌ URL {tmdb_url} 未找到任何指定的中文內容 (官方翻譯和 en-US 標題/原標題均不符合)。")
            return False # 檢查成功但無中文，返回 False

        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"檢查 {tmdb_url} 時發生網路錯誤 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"檢查 {tmdb_url} 時發生網路錯誤，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            logging.error(f"檢查 {tmdb_url} 時發生非預期錯誤: {traceback.format_exc()}")
            break # 遇到非網路問題，直接中斷

    logging.warning(f"❌ 檢查 {tmdb_url} 在多次重試後最終失敗。")
    return False

async def search_tmdb(session, semaphore, title, api_key):
    """使用 TMDB v3 API 異步搜尋標題，並行過濾並返回具有中文翻譯的結果的 TMDB URL
    (使用者需求 2025-09-02) 新增錯誤重試機制，如果 API 請求失敗，最多重試 100 次。
    """
    base_url = "https://api.themoviedb.org/3/search/multi"
    api_base_url = "https://api.themoviedb.org/3"
    variants = generate_title_variants(title)
    
    for variant in variants:
        params = {
            "api_key": api_key,
            "query": variant,
            "include_adult": "false",
            "language": "zh-HK,zh-CN,zh-TW,ko,en"
        }
        
        # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
        response_data = None
        for attempt in range(101):
            try:
                async with semaphore:
                    response = await session.get(base_url, params=params, timeout=20)
                
                # 處理 API 速率限制 (429)
                if response.status_code == 429:
                    wait_time = int(response.headers.get("Retry-After", 2)) # 優先使用 API 提供的等待時間
                    logging.warning(f"觸發 TMDB API 速率限制 (搜尋 '{variant}')，將等待 {wait_time} 秒後重試 (第 {attempt + 1}/101 次)")
                    await asyncio.sleep(wait_time)
                    # 這裡不 break，而是 continue 到下一次重試
                    continue

                if response.status_code == 404:
                    logging.debug(f"搜尋 '{variant}' 時項目不存在 (404)。")
                    break # 404 不重試，跳出此變體的搜尋

                response.raise_for_status()
                response_data = response.json()
                break # 成功，跳出重試迴圈
            except (httpx.RequestError, httpx.HTTPStatusError) as e:
                if attempt < 100:
                    wait_time = min(2 ** attempt, 60)
                    logging.warning(f"搜尋 '{variant}' 時發生網路錯誤 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                    await asyncio.sleep(wait_time)
                else:
                    logging.error(f"搜尋 '{variant}' 時發生 HTTP 錯誤，已達最大重試次數 (101 次): {e}")
            except Exception as ex:
                logging.error(f"搜尋 '{variant}' 時發生非預期錯誤: {traceback.format_exc()}")
                break # 非網路錯誤，跳出重試

        if not response_data:
            continue # 如果所有重試都失敗，嘗試下一個變體

        initial_results = []
        for result in response_data.get('results', []):
            media_type = result.get('media_type')
            if media_type in ['movie', 'tv']:
                tmdb_id = result.get('id')
                tmdb_url = f"https://www.themoviedb.org/{media_type}/{tmdb_id}"
                initial_results.append((tmdb_url, media_type))
        
        if not initial_results:
            continue

        filtered_results = []
        total_count = len(initial_results)
        processed_count = 0
        
        logging.info(f"對 '{variant}' 的 {total_count} 個初步結果進行並行中文翻譯檢查...")

        future_to_data = {
            has_chinese_translation(session, semaphore, api_key, res_url, api_base_url): (res_url, media_type)
            for res_url, media_type in initial_results
        }

        for future in asyncio.as_completed(future_to_data):
            original_url, original_media_type = future_to_data[future]
            has_translation = await future
            processed_count += 1
            
            if has_translation:
                filtered_results.append((original_url, original_media_type))
            
            logging.info(f"翻譯檢查進度 (標題: '{variant}'): {processed_count}/{total_count} (已找到 {len(filtered_results)} 個)")

        if filtered_results:
            logging.info(f"標題變體 '{variant}' 找到 {len(filtered_results)} 個有效(有中文翻譯)的結果。")
            return filtered_results
        else:
            logging.warning(f"標題變體 '{variant}' 的初步結果均無指定的中文翻譯，將嘗試下一個變體。")
            continue

    logging.warning(f"標題 '{title}' (及其變體) 在 TMDB 未找到任何具有中文翻譯的結果。")
    return []

async def process_single_url(session, semaphore, url, api_key):
    """
    異步處理單一 Wikipedia URL：抓取、解析、搜尋，並返回結果。
    (使用者需求 2025-09-02) 新增錯誤重試機制，如果網頁請求失敗，最多重試 100 次。
    """
    logging.info(f"開始處理 URL: {url}")
    if not validate_wiki_url(url):
        logging.error(f"跳過無效 URL: {url}")
        return ([], None)

    is_korean_wiki = urlparse(url).netloc == 'ko.wikipedia.org'
    
    # (使用者需求 2025-09-02) 如果網頁請求失敗，重試最多 100 次
    response_content = None
    for attempt in range(101):
        try:
            async with semaphore:
                response = await session.get(url, timeout=20)
            response.raise_for_status()
            response_content = response.content
            break # 成功，跳出重試
        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"無法訪問頁面 {url} (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"無法訪問頁面 {url}，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            logging.error(f"訪問頁面 {url} 時發生非預期錯誤: {traceback.format_exc()}")
            break
            
    if not response_content:
        return ([], None) # 所有重試失敗

    soup = BeautifulSoup(response_content, 'html.parser')
    
    if not check_page_format(soup, url):
        logging.error(f"頁面 {url} 未通過格式檢查，可能無可抓取的表格，跳過")
        return ([], None)

    logging.info(f"從 {url} 提取標題...")
    extracted_page_titles = extract_titles(soup, is_korean_wiki)
    
    if not extracted_page_titles:
        logging.warning(f"未從 {url} 找到任何標題。")
        return ([], None)
        
    logging.info(f"從 {url} 提取到 {len(extracted_page_titles)} 個標題，開始並行搜尋 TMDB...")

    search_tasks = [search_tmdb(session, semaphore, title, api_key) for title in extracted_page_titles]
    search_results_list = await asyncio.gather(*search_tasks)

    url_found_results = []
    url_unfound_titles = []
    seen_tmdb_urls = set()

    for i, title in enumerate(extracted_page_titles):
        tmdb_results_for_title = search_results_list[i]
        if tmdb_results_for_title:
            for res_tmdb_url, res_media_type in tmdb_results_for_title:
                if res_tmdb_url not in seen_tmdb_urls:
                    url_found_results.append((res_tmdb_url, res_media_type))
                    seen_tmdb_urls.add(res_tmdb_url)
        else:
            url_unfound_titles.append(title)
    
    unfound_with_source = (url, url_unfound_titles) if url_unfound_titles else None
    
    logging.info(f"完成處理 URL: {url}。找到 {len(url_found_results)} 個結果，{len(url_unfound_titles)} 個標題未找到。")
    return (url_found_results, unfound_with_source)

def save_to_excel(results, filename='../data/bulk_videos.xlsx', sheet_name='Bulk'):
    """將搜尋結果保存到 Excel 的 Bulk 工作表並刪除重複 Video URL"""
    check_create_bulk_excel(filename, sheet_name) 
    
    data_to_save = []
    for tmdb_url, _ in results: 
        data_to_save.append([tmdb_url, '', '', '']) 

    if not data_to_save:
        logging.info("沒有新的 TMDB 結果需要保存到 Excel。")
        return

    df_new_results = pd.DataFrame(data_to_save, columns=['Video URL', 'TMDB URL', 'Split', 'Done'])
    
    save_to_bulk_excel(df_new_results, filename, sheet_name) 

def save_to_bulk_excel(df, filename, sheet_name='Bulk'):
    """
    將數據追加到 bulk_videos.xlsx 的 Bulk sheet 末尾，保留現有數據。
    (使用者需求 2025-08-19) 升級檔案鎖定機制，使其與 tmdb_importer.py/bulk_video_person.py 同步，
    增加一個內部重試迴圈來處理來自外部程式（如 Excel）的檔案鎖定（PermissionError）以及其他暫時性寫入錯誤。
    """
    # --- 檔案鎖定機制 ---
    lock_file = filename + ".lock"
    output_filename = os.path.basename(filename)
    
    logging.info(f"準備寫入 '{output_filename}'，開始檢查檔案鎖...")

    # --- 階段一：獲取腳本鎖 (.lock 檔案)，防止其他 Python 腳本同時寫入 ---
    lock_acquired = False
    script_wait_message_logged = False
    while True:
        try:
            # 嘗試以獨佔模式創建鎖定檔
            lock_fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(lock_fd)
            lock_acquired = True
            if script_wait_message_logged:
                 logging.info(f"檢測到 '{output_filename}' 已可用，現在開始寫入資料")
            logging.info(f"成功獲取 '{output_filename}' 的腳本寫入權限。")
            break # 成功獲取鎖，跳出迴圈
        except FileExistsError:
            if not script_wait_message_logged:
                logging.info(f"'{output_filename}' 正被另一個程式實例處理中，現在等待其完成...")
                script_wait_message_logged = True
            time.sleep(1) # 等待 1 秒後重試
        except Exception as e:
            logging.error(f"獲取檔案鎖時發生未知錯誤: {e}")
            return # 發生其他錯誤時，放棄

    # --- 階段二：在持有腳本鎖的情況下，進行寫入並處理外部程式鎖 ---
    if lock_acquired:
        # 建立一個外層的 try...finally 結構，
        # 目的是確保無論內部重試邏輯成功或失敗，最終都一定會執行 finally 區塊來釋放檔案鎖。
        try:
            write_successful = False
            permission_error_logged = False
            while not write_successful:
                try:
                    if script_wait_message_logged and not permission_error_logged:
                        logging.info(f"檢測到 '{output_filename}' 已可用，現在開始寫入資料")

                    # --- 原有的寫入邏輯 ---
                    # 步驟 1: 讀取現有數據
                    try:
                        existing_df = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
                    except (FileNotFoundError, ValueError): # 處理檔案或工作表不存在的情況
                        existing_df = pd.DataFrame(columns=['Video URL', 'TMDB URL', 'Split', 'Done', '覆蓋劇集標題', '覆蓋劇集描述', '覆蓋分集標題', '覆蓋分集描述', '來源無資料時清空TMDB', '維基人名取代現有'])

                    # 步驟 2: 準備並篩選出真正新的 URL
                    if not existing_df.empty and 'Video URL' in existing_df.columns:
                        existing_urls = set(existing_df['Video URL'].dropna())
                        new_unique_df = df[~df['Video URL'].isin(existing_urls)]
                    else:
                        new_unique_df = df

                    # 步驟 3: 如果有新數據，則追加並保存
                    if not new_unique_df.empty:
                        final_df = pd.concat([existing_df, new_unique_df], ignore_index=True)
                        logging.info(f"發現 {len(new_unique_df)} 個新的 URL，將其追加到 Excel 中。")

                        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        logging.info(f"成功將新數據寫入 Excel。")
                    else:
                        logging.info("本次運行未發現新的 URL，Excel 檔案內容未變動。")
                    
                    # 成功後退出重試迴圈
                    write_successful = True
                
                except PermissionError:
                    # 如果是權限錯誤，表示檔案被 Excel 等外部程式鎖定
                    if not permission_error_logged:
                        logging.warning(f"'{output_filename}' 正被外部程式(如 Excel)使用中，請關閉後程式將自動重試...")
                        permission_error_logged = True
                    time.sleep(1) # 等待 1 秒後重試
                except Exception as e:
                    # 如果是其他錯誤，記錄詳細訊息並終止，因為可能不是暫時性問題
                    logging.error(f"在持有鎖的情況下寫入 Excel 時發生非預期的錯誤:")
                    logging.error(traceback.format_exc())
                    break # 跳出重試迴圈，放棄寫入
        finally:
            # (極其重要) 無論 try 區塊是否成功，都必須確保鎖定檔案被刪除
            try:
                os.remove(lock_file)
                logging.info(f"已釋放 '{output_filename}' 的寫入權限。")
            except OSError as e:
                logging.error(f"釋放檔案鎖 '{lock_file}' 時失敗: {e}")

def save_unfound_titles(titles, source_url=None, filename='../data/unfound_titles.xlsx', sheet_name='Unfound'):
    """
    將未找到的標題保存到 unfound_titles.xlsx 的 Unfound sheet，包含來源 URL。
    (使用者需求 2025-08-19) 升級檔案鎖定機制，使其與 tmdb_importer.py/bulk_video_person.py 同步，
    增加一個內部重試迴圈來處理來自外部程式（如 Excel）的檔案鎖定（PermissionError）以及其他暫時性寫入錯誤。
    """
    output_dir = os.path.dirname(filename)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # --- 檔案鎖定機制 ---
    lock_file = filename + ".lock"
    output_filename = os.path.basename(filename)
    
    logging.info(f"準備寫入 '{output_filename}'，開始檢查檔案鎖...")

    # --- 階段一：獲取腳本鎖 (.lock 檔案) ---
    lock_acquired = False
    script_wait_message_logged = False
    while True:
        try:
            lock_fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(lock_fd)
            lock_acquired = True
            if script_wait_message_logged:
                 logging.info(f"檢測到 '{output_filename}' 已可用，現在開始寫入資料")
            logging.info(f"成功獲取 '{output_filename}' 的腳本寫入權限。")
            break
        except FileExistsError:
            if not script_wait_message_logged:
                logging.info(f"'{output_filename}' 正被另一個程式實例處理中，現在等待其完成...")
                script_wait_message_logged = True
            time.sleep(1)
        except Exception as e:
            logging.error(f"獲取檔案鎖時發生未知錯誤: {e}")
            return

    # --- 階段二：在持有腳本鎖的情況下，進行寫入並處理外部程式鎖 ---
    if lock_acquired:
        try:
            write_successful = False
            permission_error_logged = False
            while not write_successful:
                try:
                    if script_wait_message_logged and not permission_error_logged:
                        logging.info(f"檢測到 '{output_filename}' 已可用，現在開始寫入資料")
                    
                    # --- 原有的寫入邏輯 ---
                    data_to_add = {'Title': titles}
                    if source_url:
                        data_to_add['Source URL'] = [source_url] * len(titles)
                    df = pd.DataFrame(data_to_add)
                    required_columns = ['Title', 'Source URL']

                    if not os.path.exists(filename):
                        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                            df.to_excel(writer, sheet_name=sheet_name, index=False, columns=required_columns)
                        logging.info(f"已創建 {filename} 並保存未找到的標題到 {sheet_name} sheet")
                    else:
                        try:
                            existing_df = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
                        except ValueError:
                            existing_df = pd.DataFrame(columns=required_columns)

                        for col in required_columns:
                            if col not in existing_df.columns:
                                existing_df[col] = ''

                        combined_df = pd.concat([existing_df, df], ignore_index=True)
                        combined_df.drop_duplicates(subset=required_columns, keep='first', inplace=True)

                        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        logging.info(f"成功將 {len(titles)} 個未找到的標題保存到 {filename} 的 {sheet_name} sheet，總行數: {len(combined_df)}")
                    
                    write_successful = True
                
                except PermissionError:
                    if not permission_error_logged:
                        logging.warning(f"'{output_filename}' 正被外部程式(如 Excel)使用中，請關閉後程式將自動重試...")
                        permission_error_logged = True
                    time.sleep(1)
                except Exception as e:
                    logging.error(f"在持有鎖的情況下寫入 Excel 時發生非預期的錯誤:")
                    logging.error(traceback.format_exc())
                    break
        finally:
            try:
                os.remove(lock_file)
                logging.info(f"已釋放 '{output_filename}' 的寫入權限。")
            except OSError as e:
                logging.error(f"釋放檔案鎖 '{lock_file}' 時失敗: {e}")

async def main():
    """
    主執行函數
    (使用者需求 2025-09-02) 新增整體的進度條與預計完成時間顯示
    """
    api_key = load_config()
    if not api_key:
        logging.error("程式將退出")
        return
    
    url_input = input("請輸入 Wikipedia 頁面的 URL（多個 URL 用 , 或 ， 分隔）: ").strip()
    urls = [url.strip() for url in re.split('[,，]', url_input) if url.strip()]
    
    if not urls:
        logging.error("未提供有效的 URL，程式將退出")
        return
    
    all_tmdb_results_for_excel = []
    all_unfound_titles_with_source = []
    final_seen_tmdb_urls = set()
    
    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)
    async with httpx.AsyncClient(http2=True, timeout=30) as session: # 使用 HTTP/2 並增加超時
        processing_tasks = [process_single_url(session, semaphore, url, api_key) for url in urls]
        
        # (使用者需求 2025-09-02) 新增整體的進度條與預計完成時間顯示
        total_tasks = len(processing_tasks)
        processed_tasks = 0
        start_time = time.time()
        
        logging.info(f"--- 開始處理 {total_tasks} 個 Wikipedia URL ---")

        for future in asyncio.as_completed(processing_tasks):
            found_results, unfound_info = await future
            processed_tasks += 1

            # 彙總結果
            if found_results:
                for tmdb_url, media_type in found_results:
                    if tmdb_url not in final_seen_tmdb_urls:
                        all_tmdb_results_for_excel.append((tmdb_url, media_type))
                        final_seen_tmdb_urls.add(tmdb_url)
            
            if unfound_info:
                all_unfound_titles_with_source.append(unfound_info)

            # 更新進度條
            elapsed_time = time.time() - start_time
            items_per_second = processed_tasks / elapsed_time if elapsed_time > 0 else 0
            remaining_items = total_tasks - processed_tasks

            if items_per_second > 0:
                estimated_remaining_seconds = remaining_items / items_per_second
                remaining_minutes = int(estimated_remaining_seconds // 60)
                remaining_seconds_val = int(estimated_remaining_seconds % 60)
                etc_str = f"預計剩餘 {remaining_minutes} 分 {remaining_seconds_val} 秒"
            else:
                etc_str = "預計剩餘時間計算中..."
                
            progress_bar = f"整體進度: {processed_tasks}/{total_tasks} 個 URL"
            found_count = f"已找到 {len(all_tmdb_results_for_excel)} 個"
            
            sys.stdout.write(f"\r{progress_bar} ({found_count}) - {etc_str}".ljust(80))
            sys.stdout.flush()

    print() # 完成後換行
    logging.info("所有 URL 處理完畢。")

    # 保存結果到 Excel
    if all_tmdb_results_for_excel:
        save_to_excel(all_tmdb_results_for_excel)
        logging.info(f"所有 URL 的搜尋結果已處理並嘗試保存到 ../data/bulk_videos.xlsx 的 Bulk 工作表")
    else:
        logging.warning("所有 URL 均未找到任何有效的 TMDB 結果。")
    
    if all_unfound_titles_with_source:
        for source_url, unfound_list_for_url in all_unfound_titles_with_source:
            if unfound_list_for_url:
                save_unfound_titles(unfound_list_for_url, source_url=source_url)
    
    logging.info("資料已處理完成。")

if __name__ == "__main__":
    # 在 Windows 上，為 asyncio 設定不同的事件循環策略以避免錯誤
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())