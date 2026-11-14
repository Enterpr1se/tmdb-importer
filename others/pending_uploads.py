# 【2026-11-14】已提取、未上傳項目的記錄（data/details_cache/pending_uploads.json）。
# extractor 提取成功、放入上傳佇列時記錄；上傳完 / 跳過 / 失敗就刪除。下次 bulk 啟動時，用家揀保留的話，
# 符合條件的項目直接放返入上傳佇列，唔使重新提取（條件見 select_keepable）。
# 只用標準庫，start.py 未 import 主程式前都可以用（count_pending）。
import json
import os
import threading
import time

DETAILS_CACHE_DIR = os.path.join('data', 'details_cache')
MANIFEST_FILE = os.path.join(DETAILS_CACHE_DIR, 'pending_uploads.json')
_lock = threading.Lock()


def _load(path=MANIFEST_FILE):
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save(data, path=MANIFEST_FILE):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def _clean(v):
    if v is None or (isinstance(v, float) and v != v):
        return ''
    s = str(v).strip()
    return '' if s.lower() in ('nan', 'none') else s


def row_fingerprint(row):
    """bulk 表一行嘅內容（Done 以外所有欄），用嚟確認下次啟動時同一行號仲係同一個項目、內容冇改。"""
    try:
        items = row.items()
    except AttributeError:
        return {}
    return {str(k): _clean(v) for k, v in items if str(k) != 'Done' and not str(k).startswith('Unnamed')}


def record(task, sheet, row_fp, run_options):
    """放入上傳佇列時記錄。task = 上傳佇列項目（dict，可 JSON 化）。"""
    entry = {'task': task, 'sheet': sheet, 'row': row_fp, 'options': run_options,
             'recorded_at': time.strftime('%Y-%m-%d %H:%M:%S')}
    with _lock:
        data = _load()
        data[str(task['row_index'])] = entry
        _save(data)


def remove(row_index):
    with _lock:
        data = _load()
        if data.pop(str(row_index), None) is not None:
            _save(data)


def count_pending(path=MANIFEST_FILE):
    """(數目, 最後記錄時間) — start.py 用嚟決定問唔問保留。只計 details 檔案仲喺度嘅項目。"""
    data = _load(path)
    entries = [e for e in data.values() if isinstance(e, dict)
               and os.path.exists(str((e.get('task') or {}).get('details_file') or ''))]
    last = max((e.get('recorded_at') or '' for e in entries), default='')
    return len(entries), last


def select_keepable(sheet, run_options, get_row):
    """揀出可以保留嘅項目。get_row(row_index) → 而家 bulk 表嗰行（或 None）。
    條件：同一張工作表、啟動選項一樣、details 檔案仲喺度、嗰行內容（Done 以外）冇改、Done 仍然係空。
    回傳 (保留嘅 [task], {放棄原因: 數目})。"""
    kept, dropped = [], {}

    def drop(reason):
        dropped[reason] = dropped.get(reason, 0) + 1

    for key, e in sorted(_load().items(), key=lambda kv: int(kv[0]) if kv[0].isdigit() else 0):
        task = (e or {}).get('task') or {}
        if not isinstance(task, dict) or 'row_index' not in task:
            drop('記錄損壞')
            continue
        if e.get('sheet') != sheet:
            drop('工作表唔同')
            continue
        if e.get('options') != run_options:
            drop('啟動選項同上次唔同')
            continue
        if not os.path.exists(str(task.get('details_file') or '')):
            drop('details 檔案唔見咗')
            continue
        row = get_row(int(task['row_index']))
        if row is None or row_fingerprint(row) != e.get('row'):
            drop('Excel 嗰行內容改咗或者行號移咗')
            continue
        if _clean(row.get('Done')) not in ('', '0', '0.0'):
            drop('Done 已經有值')
            continue
        kept.append(task)
    return kept, dropped


def retain(row_indices):
    """只保留指定行號嘅記錄（啟動時保留咗嘅項目），其他刪除。"""
    keep = {str(i) for i in row_indices}
    with _lock:
        data = _load()
        _save({k: v for k, v in data.items() if k in keep})
