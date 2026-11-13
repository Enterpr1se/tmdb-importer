"""sub-program 工具共用嘅輸出：將搵到嘅 TMDB 網址寫入 txt，唔直接改 data/bulk_videos.xlsx。

每個工具一個檔案：sub-program/<工具名>_urls.txt，一行一個網址，只會追加未有嘅網址。
寫入前會略過：txt 入面已有嘅、同埋 bulk_videos.xlsx 任何一張工作表（Video URL / TMDB URL 欄）已有嘅網址。
用家自己揀要貼邊啲入 bulk_videos.xlsx（Video URL 欄）。
"""
import logging
import os
import re

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
BULK_XLSX = os.path.join(os.path.dirname(OUTPUT_DIR), 'data', 'bulk_videos.xlsx')
URL_COLUMNS = ('Video URL', 'TMDB URL')

_TMDB_RE = re.compile(r'themoviedb\.org/(?:[a-z]{2}(?:-[A-Z]{2})?/)?(movie|tv|person|collection)/(\d+)', re.I)


def normalize_url(url):
    """用嚟比較嘅 key：TMDB 網址變成 'tmdb:tv/123'（唔理 slug、語言、query）；其他網址去 query / fragment / 尾 '/'、細楷。"""
    u = str(url or '').strip()
    if not u or u.lower() in ('nan', 'none'):
        return ''
    m = _TMDB_RE.search(u)
    if m:
        return f'tmdb:{m.group(1).lower()}/{m.group(2)}'
    u = re.sub(r'[?#].*$', '', u).rstrip('/').lower()
    return re.sub(r'^https?://(www\.)?', '', u)


def output_path(tool_name):
    return os.path.join(OUTPUT_DIR, f'{tool_name}_urls.txt')


def load_bulk_excel_urls(path=BULK_XLSX):
    """讀 bulk_videos.xlsx 所有工作表嘅 Video URL / TMDB URL → set(normalize_url)。讀唔到就回傳 None。"""
    if not os.path.exists(path):
        return set()
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        logging.warning(f"讀取 {os.path.basename(path)} 失敗（{e}），今次唔會略過 Excel 已有嘅網址。")
        return None
    keys = set()
    try:
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            header = next(rows, None) or ()
            cols = [i for i, h in enumerate(header) if str(h or '').strip() in URL_COLUMNS]
            if not cols:
                continue
            for row in rows:
                for i in cols:
                    if i < len(row) and row[i]:
                        k = normalize_url(row[i])
                        if k:
                            keys.add(k)
    finally:
        wb.close()
    return keys


def write_urls_txt(urls, tool_name):
    """追加網址到 <工具名>_urls.txt。回傳 (檔案路徑, 新增數目)。"""
    path = output_path(tool_name)
    in_txt = set()
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            in_txt = {normalize_url(line) for line in f if line.strip()}
    in_excel = load_bulk_excel_urls() or set()

    new, skipped_txt, skipped_excel = [], 0, 0
    for u in urls:
        u = str(u or '').strip()
        k = normalize_url(u)
        if not k:
            continue
        if k in in_excel:
            skipped_excel += 1
        elif k in in_txt:
            skipped_txt += 1
        else:
            in_txt.add(k)
            new.append(u)
    if new:
        with open(path, 'a', encoding='utf-8') as f:
            f.write('\n'.join(new) + '\n')
    notes = []
    if skipped_excel:
        notes.append(f"{skipped_excel} 個已經喺 bulk_videos.xlsx")
    if skipped_txt:
        notes.append(f"{skipped_txt} 個已經喺 txt")
    logging.info(f"已寫入 {len(new)} 個新網址到 {path}"
                 + (f"（略過：{'、'.join(notes)}）" if notes else "")
                 + "。需要嘅話請自行貼入 bulk_videos.xlsx。")
    return path, len(new)
