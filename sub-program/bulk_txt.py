"""sub-program 工具共用嘅輸出：將搵到嘅 TMDB 網址寫入 txt，唔直接改 data/bulk_videos.xlsx。

每個工具一個檔案：sub-program/<工具名>_urls.txt，一行一個網址，只會追加未有嘅網址。
寫入前會略過：txt 入面已有嘅、同埋 bulk_videos.xlsx 任何一張工作表（Video URL / TMDB URL 欄）已有嘅網址。
用家自己揀要貼邊啲入 bulk_videos.xlsx（Video URL 欄），或者喺小工具面板撳「匯入 Excel」（import_urls_to_bulk）。
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


def canonical_url(url):
    """TMDB 網址變成 https://www.themoviedb.org/<類型>/<ID>（同主程式 normalize_tmdb_url 一樣）；其他網址唔改。"""
    u = str(url or '').strip()
    m = _TMDB_RE.search(u)
    return f'https://www.themoviedb.org/{m.group(1).lower()}/{m.group(2)}' if m else u


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


# ---------------------------------------------------------------------------
# 【2026-11】由小工具面板（tools_gui.py）手動匯入：用家撳掣先會寫 bulk_videos.xlsx。
# 工具本身照舊只寫 txt。
# ---------------------------------------------------------------------------
LOCK_HEARTBEAT_SEC = 3   # 主程式當超過 15 秒冇更新嘅 .lock 係過期，所以匯入期間要不停更新


class BulkFileBusy(Exception):
    """主程式或者其他程式正在寫 bulk_videos.xlsx。"""


def list_bulk_sheets(path=BULK_XLSX):
    import openpyxl
    if not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _bulk_columns():
    try:
        import sys
        root = os.path.dirname(OUTPUT_DIR)
        if root not in sys.path:
            sys.path.insert(0, root)
        from others.excel_handler import BULK_COLUMNS
        return list(BULK_COLUMNS)
    except Exception:
        return ['Video URL', 'TMDB URL', 'Split', 'Done']


def import_urls_to_bulk(urls, sheet_name, create_sheet=False, path=BULK_XLSX, progress=None, lock_wait=60):
    """將網址加到 bulk_videos.xlsx 某張工作表底部（Video URL 欄），或者新增工作表。
    已經喺任何工作表嘅網址、同一批入面重複嘅網址都會略過。
    用同主程式一樣嘅 bulk_videos.xlsx.lock 同暫存檔方式儲存（保留格式），改之前備份到 data/backup/。
    回傳 (新增數目, 略過數目)。"""
    import shutil
    import threading
    import time
    import openpyxl

    say = progress or (lambda msg: None)
    lock_file = path + '.lock'

    # 1. 攞鎖（主程式寫 Done 時會建立同一個 .lock）
    deadline = time.time() + lock_wait
    while True:
        try:
            with open(lock_file, 'x') as f:
                f.write(str(os.getpid()))
            break
        except FileExistsError:
            try:
                age = time.time() - os.path.getmtime(lock_file)
            except OSError:
                continue
            if age > 15:                       # 同主程式一樣：15 秒冇更新當過期
                try:
                    os.remove(lock_file)
                except OSError:
                    pass
                continue
            if time.time() > deadline:
                raise BulkFileBusy('bulk_videos.xlsx 正被主程式寫入，請稍後再試。')
            say('主程式正在寫入 bulk_videos.xlsx，等緊…')
            time.sleep(1)

    stop = threading.Event()

    def heartbeat():
        while not stop.wait(LOCK_HEARTBEAT_SEC):
            try:
                os.utime(lock_file, None)
            except OSError:
                pass

    hb = threading.Thread(target=heartbeat, daemon=True)
    hb.start()
    try:
        say('檢查 Excel 已有嘅網址…')
        in_excel = load_bulk_excel_urls(path) or set()
        new, seen, skipped = [], set(), 0
        for u in urls:
            u = str(u or '').strip()
            k = normalize_url(u)
            if not k:
                continue
            if k in in_excel or k in seen:
                skipped += 1
                continue
            seen.add(k)
            new.append(canonical_url(u))
        if not new:
            return 0, skipped

        say('讀取 bulk_videos.xlsx…')
        if os.path.exists(path):
            backup_dir = os.path.join(os.path.dirname(path), 'backup')
            os.makedirs(backup_dir, exist_ok=True)
            shutil.copy2(path, os.path.join(backup_dir, 'bulk_videos.before_import.xlsx'))
            wb = openpyxl.load_workbook(path)
        else:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        columns = _bulk_columns()
        if create_sheet:
            if sheet_name in wb.sheetnames:
                raise ValueError(f"工作表「{sheet_name}」已經存在。")
            ws = wb.create_sheet(sheet_name)
            ws.append(columns)
        else:
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"搵唔到工作表「{sheet_name}」。")
            ws = wb[sheet_name]
        header = [c.value for c in ws[1]] if ws.max_row >= 1 else []
        if not any(header):
            for i, col in enumerate(columns, 1):
                ws.cell(row=1, column=i, value=col)
            header = columns
        if 'Video URL' not in header:
            raise ValueError(f"工作表「{sheet_name}」第一行冇「Video URL」欄，唔係 bulk 格式。")
        vcol = header.index('Video URL') + 1
        ncols = max(4, len(header))
        last = ws.max_row
        while last > 1 and all(ws.cell(row=last, column=c).value in (None, '') for c in range(1, ncols + 1)):
            last -= 1
        for i, u in enumerate(new, 1):
            ws.cell(row=last + i, column=vcol, value=u)

        say(f'儲存 bulk_videos.xlsx（加 {len(new)} 行）…')
        tmp = path + '.tmp.xlsx'
        wb.save(tmp)
        wb.close()
        check = openpyxl.load_workbook(tmp, read_only=True)
        check.close()
        try:
            os.replace(tmp, path)
        except PermissionError:
            os.remove(tmp)
            raise BulkFileBusy('bulk_videos.xlsx 正喺 Excel 開住，請先關閉再匯入。')
        return len(new), skipped
    finally:
        stop.set()
        hb.join(timeout=5)
        try:
            os.remove(lock_file)
        except OSError:
            pass
