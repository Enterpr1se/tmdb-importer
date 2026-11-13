import pandas as pd
import os
import logging
import re
from openpyxl import load_workbook

def is_generic_title(title):
    """
    檢查標題是否為泛用型（完全匹配，不含額外標題文字）。
    # 原要求：處理多來源extractor；觸發：識別如「第1集」、「第 10 集」、「Season 1」等純編號標題。
    # 修改：優化正則表達式，容許數字前後有不同數量的空格。
    """
    if title is None:
        return True
        
    title_str = str(title).strip()

    if not title_str:
        return True

    # 使用 .isdigit() 檢查處理後的字串是否完全由數字組成。
    if title_str.isdigit():
        logging.debug(f"標題 '{title}' 被識別為純數字，視為泛用型。")
        return True
        
    # 中文數字正則
    chinese_numerals = r'[一二三四五六七八九十百千萬零]+'
    # 阿拉伯數字或中文數字
    num_pattern = fr'(?:\d+|{chinese_numerals})'
    
    # 【優化】使用 \s* 代替固定空格，確保能匹配「第1集」、「第 1 集」等各種空格組合
    generic_patterns = [
        fr'^第\s*{num_pattern}\s*集$',  # 匹配 "第1集", "第 10 集"
        fr'^第\s*{num_pattern}\s*季$',
        fr'^第\s*{num_pattern}\s*話$',
        fr'^第\s*{num_pattern}\s*輯$',
        r'^Episode\s*\d+$',
        r'^Season\s*\d+$',
        r'^特別篇$',
    ]
    for pattern in generic_patterns:
        if re.match(pattern, title_str, re.IGNORECASE):
            logging.debug(f"標題 '{title}' 匹配泛用模式 (完全匹配): {pattern}")
            return True
            
    return False

def clean_generic_title(title):
    """
    移除標題中的泛用型前綴（例如移除「第 10 集 」後保留後方的標題）。
    # 原要求：自動清除標題前綴；觸發：當標題包含「第 X 集」、「第 X 話」等前綴及分隔符時自動清理。
    # 修改：強化正則表達式，支援冒號、連字號、句號及任意空格的分隔。
    """
    if not title or not isinstance(title, str):
        return str(title) if title is not None else ""
    
    cleaned_title = title.strip()
    num_pattern = r'(?:\d+|[一二三四五六七八九十百千萬零]+)'
    
    # 【優化】匹配前綴及後方可能跟隨的分隔符
    # [ \t　]* : 匹配任意空白
    # [ :：.\-]* : 匹配常見分隔符 (空格, 冒號, 全形冒號, 點, 連字號)
    patterns_to_clean = [
        fr'^第\s*{num_pattern}\s*集[ :：.\-]*', # 匹配 "第 10 集", "第 10 集：", "第1集 "
        fr'^第\s*{num_pattern}\s*季[ :：.\-]*',
        fr'^第\s*{num_pattern}\s*話[ :：.\-]*',
        fr'^第\s*{num_pattern}\s*輯[ :：.\-]*',
        r'^特別篇[ :：.\-]*',
        r'^Episode\s*\d+[ :：.\-]*',
        r'^Season\s*\d+[ :：.\-]*',
    ]
    
    for pattern in patterns_to_clean:
        original_len = len(cleaned_title)
        # 使用正則替換前綴為空字串 (Flags: 忽略大小寫)
        cleaned_title = re.sub(pattern, '', cleaned_title, flags=re.IGNORECASE).strip()
        
        # 如果長度變短，代表有清除東西，記錄日誌並跳出循環 (通常只會有一個前綴)
        if len(cleaned_title) < original_len:
            logging.info(f"標題清理：原文='{title}' -> 清理後='{cleaned_title}'")
            break 

    return cleaned_title

def update_season_numbers(seasons_df):
    """從 Season Name 中提取季數並更新 Season Number"""
    # 確保 'Season Name' 列中的所有值都是字符串類型
    seasons_df['Season Name'] = seasons_df['Season Name'].astype(str)
    
    # 提取 Season Number 從 Season Name 中的 "第 x 季" 或 "第 x 輯" 或 "Season X"
    season_numbers = seasons_df['Season Name'].str.extract(r'(?:第 (\d+) (季|輯)|Season (\d+))')
    season_numbers = season_numbers.iloc[:, 0].fillna(season_numbers.iloc[:, 2]).dropna().astype(int)
    
    # 僅更新泛型標題的 Season Number
    mask = seasons_df['Season Name'].apply(is_generic_title) & ~season_numbers.isna()
    seasons_df.loc[mask, 'Season Number'] = season_numbers[mask].astype(int)
    
    # 清理 Season Name，去除泛型部分
    seasons_df['Season Name'] = seasons_df['Season Name'].apply(clean_generic_title)
    
    return seasons_df

def update_episodes_numbers(episodes_df, seasons_df):
    """根據 Seasons 的 Season Number 更新 Episodes 的 Season Number"""
    # 創建舊的 Season Number 到新的 Season Number 的映射
    old_to_new_season_number = {index + 1: row['Season Number'] for index, row in seasons_df.iterrows()}

    # 僅更新 Episodes 中匹配到的 Season Number
    episodes_df['Season Number'] = episodes_df['Season Number'].map(old_to_new_season_number).fillna(episodes_df['Season Number']).astype(int)
    return episodes_df

def process_excel(input_file, output_file):
    """處理 Excel 文件，更新 Seasons 和 Episodes 的編號"""
    logging.info(f"開始處理 Excel 文件: {input_file}")

    try:
        if os.path.exists(input_file):
            with pd.ExcelFile(input_file, engine='openpyxl') as reader:
                existing_sheets = {sheet: reader.parse(sheet) for sheet in reader.sheet_names}
        else:
            logging.error(f"文件 {input_file} 不存在，無法處理。")
            return

        # 處理 Seasons sheet
        if 'Seasons' in existing_sheets:
            existing_sheets['Seasons'] = update_season_numbers(existing_sheets['Seasons'])
            logging.info("已更新 Seasons sheet 的 Season Number 和 Season Name")

        # 處理 Episodes sheet
        if 'Episodes' in existing_sheets and 'Seasons' in existing_sheets:
            existing_sheets['Episodes'] = update_episodes_numbers(existing_sheets['Episodes'], existing_sheets['Seasons'])
            logging.info("已更新 Episodes sheet 的 Season Number")

        # 將所有工作表寫回 Excel 檔案
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
            for sheet_name, data_df in existing_sheets.items():
                data_df.to_excel(writer, sheet_name=sheet_name, index=False)
                logging.info(f"已保存 {sheet_name} 到 {output_file}")

    except Exception as e:
        logging.error(f"處理 {input_file} 時發生錯誤: {e}")
        raise e

