import pandas as pd
import re
import logging
import os
from openpyxl import load_workbook

def clear_title(title, title_type):
    if not isinstance(title, str):
        return title

    if title_type == 'show_movie':
        # 移除《 和 》引號
        title = title.replace('《', '').replace('》', '')
        return title

    elif title_type == 'episode':
        # 清理集數標題
        # 定義多個正則表達式模式來匹配不同格式的前綴
        patterns = [
            r'^第\s*\d+\s*季\s*第\s*\d+\s*集[:：\-]?\s*',       # 第3季第9集：共喰（其一）
            r'^第\s*[一二三四五六七八九十]+\s*季\s*第\s*[一二三四五六七八九十]+\s*集[:：\-]?\s*',  # 第三季第九集：共喰（其一）
            r'^第\s*\d+\s*季\s*第\s*\d+\s*集\s*$',              # 第5季第1集
            r'^第\s*[一二三四五六七八九十]+\s*季\s*第\s*[一二三四五六七八九十]+\s*集\s*$',  # 第五季第一集
            r'^第\s*\d+\s*集[:：\-]?\s*',                       # 第1集：你是為了什麼而玩遊戲？
            r'^第\s*[一二三四五六七八九十]+\s*集[:：\-]?\s*',    # 第一集：你是為了什麼而玩遊戲？
        ]
        previous_title = None
        while title != previous_title:
            previous_title = title
            for pattern in patterns:
                title = re.sub(pattern, '', title).strip()
        
        # 若清理後的標題為空或僅包含季集格式，則返回空字串
        if not title or re.match(r'^\s*$', title):
            return ""
        
        return title

    elif title_type == 'season':
        # 清理季數標題
        if re.match(r'^第\s*\d+\s*季$', title) or re.match(r'^第\d+季$', title):
            return ""
        if re.match(r'^第[一二三四五六七八九十]+季$', title):
            return ""
        return title

    else:
        return title

def process_titles(input_file, output_file):
    logging.info(f"開始處理標題: {input_file}")

    try:
        if not os.path.exists(input_file):
            logging.error(f"文件 {input_file} 不存在，無法處理標題。")
            return

        # 使用 openpyxl 加載整個工作簿
        workbook = load_workbook(input_file)

        # 處理 Title 工作表
        if 'Title' in workbook.sheetnames:
            title_df = pd.read_excel(input_file, sheet_name='Title')
            if 'TV Show Title' in title_df.columns:
                title_df['TV Show Title'] = title_df['TV Show Title'].apply(lambda x: clear_title(x, 'show_movie'))
            if 'Movie Title' in title_df.columns:
                title_df['Movie Title'] = title_df['Movie Title'].apply(lambda x: clear_title(x, 'show_movie'))
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                title_df.to_excel(writer, sheet_name='Title', index=False)
            logging.info("成功清理並保存 Title 工作表")

        # 處理 Movies 工作表
        if 'Movies' in workbook.sheetnames:
            movies_df = pd.read_excel(input_file, sheet_name='Movies')
            if 'Movie Title' in movies_df.columns:
                movies_df['Movie Title'] = movies_df['Movie Title'].apply(lambda x: clear_title(x, 'show_movie'))
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                movies_df.to_excel(writer, sheet_name='Movies', index=False)
            logging.info("成功清理並保存 Movies 工作表")

        # 處理 Seasons 工作表
        if 'Seasons' in workbook.sheetnames:
            seasons_df = pd.read_excel(input_file, sheet_name='Seasons')
            if 'Season Name' in seasons_df.columns:
                seasons_df['Season Name'] = seasons_df['Season Name'].apply(lambda x: clear_title(x, 'season'))
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                seasons_df.to_excel(writer, sheet_name='Seasons', index=False)
            logging.info("成功清理並保存 Seasons 工作表")

        # 處理 Episodes 工作表
        if 'Episodes' in workbook.sheetnames:
            episodes_df = pd.read_excel(input_file, sheet_name='Episodes')
            if 'Episode Title' in episodes_df.columns:
                episodes_df['Episode Title'] = episodes_df['Episode Title'].apply(lambda x: clear_title(x, 'episode'))
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                episodes_df.to_excel(writer, sheet_name='Episodes', index=False)
            logging.info("成功清理並保存 Episodes 工作表")

    except PermissionError:
        input(f"請關閉 {output_file}，然後按 Enter 鍵繼續...")
    except Exception as e:
        logging.error(f"處理標題時發生錯誤: {e}")

if __name__ == "__main__":

    # 示例用法
    input_excel = 'input.xlsx'
    output_excel = 'output.xlsx'
    process_titles(input_excel, output_excel)