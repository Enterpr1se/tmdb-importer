"""
TMDB-parallel 監控面板（唯讀）

用法:  python monitor_gui.py            （或 pythonw monitor_gui.py / 雙擊 start.bat，不顯示命令視窗）
       python monitor_gui.py --base DIR  （監控其他資料夾，測試用）

設計原則
- 只「讀取」log/*.log、data/bulk_videos.xlsx、data/ai_s2t_cache.db，亦不 import 主程式，
  所以開住或關閉本面板都不會影響 start.py 運行。
  唯一例外：在「AI 被拒用詞」分頁按刪除／選用時，會寫入 ai_s2t_cache.db 的 rejected / accepted 表。
- 「▶ 啟動」在隱藏主控台執行 start.py：畫面輸出寫入 data/monitor_console.log（顯示於「主控台」分頁），
  啟動提問按面板設定經 stdin 自動回答，其餘提問可在面板輸入；「■ 停止」向該主控台送出 Ctrl+C。
  面板自己只寫 monitor_settings.json、data/monitor_console.log 及暫停旗標 data/.pause、data/.pause_upload。
- 「暫停」建立 data/.pause：主程式做完手上項目後不開始新項目；「繼續」刪除旗標。
- 「暫停上傳」建立 data/.pause_upload：uploader 喺下一個頁面步驟即時停低（2026-11-14 起，包括限流等待中），extractor 繼續提取，
  提取結果留喺上傳佇列；「繼續上傳」刪除旗標，uploader 由佇列繼續上傳。
- Windows 上以 FILE_SHARE_DELETE 方式開檔：主程式啟動時清除 log、log 輪替、
  Excel 以暫存檔取代 (os.replace) 時都不會因本面板而失敗。
- 讀取 log 只追加新內容（每秒一次），首次只讀最後約 1MB，數十 MB 的 log 亦不會拖慢。
"""
import argparse
import codecs
import collections
import datetime as dt
import io
import os
import queue
import re
import sqlite3
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, font as tkfont

NUM_EXTRACTORS = 5
INITIAL_TAIL_BYTES = 1_000_000
MAX_READ_PER_POLL = 4_000_000
MAX_VIEW_LINES = 5000
MAX_EVENTS = 300


# ---------------------------------------------------------------------------
# 共享模式開檔（不阻止主程式刪除 / 改名 / 取代檔案）
# ---------------------------------------------------------------------------
def open_shared(path):
    """以二進位唯讀開檔。Windows 上允許其他程式同時刪除/改名該檔。"""
    if os.name != 'nt':
        return open(path, 'rb')
    import ctypes
    import msvcrt
    from ctypes import wintypes
    GENERIC_READ = 0x80000000
    FILE_SHARE_ALL = 0x1 | 0x2 | 0x4          # READ | WRITE | DELETE
    OPEN_EXISTING = 3
    INVALID = ctypes.c_void_p(-1).value
    k32 = ctypes.WinDLL('kernel32', use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                                wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    h = k32.CreateFileW(path, GENERIC_READ, FILE_SHARE_ALL, None, OPEN_EXISTING, 0x80, None)
    if h is None or h == INVALID:
        err = ctypes.get_last_error()
        if err in (2, 3):
            raise FileNotFoundError(path)
        raise OSError(err, f'CreateFileW failed ({err})', path)
    fd = msvcrt.open_osfhandle(h, os.O_RDONLY | getattr(os, 'O_BINARY', 0))
    return os.fdopen(fd, 'rb')


# ---------------------------------------------------------------------------
# Log 追蹤
# ---------------------------------------------------------------------------
LINE_RE = re.compile(r'^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d),\d+ - (\S+) - ([A-Z]+) - (.*)$')


class LogEntry:
    __slots__ = ('src', 'ts', 'level', 'text', 'raw')

    def __init__(self, src, ts, level, text, raw):
        self.src, self.ts, self.level, self.text, self.raw = src, ts, level, text, raw


class LogTailer:
    """追蹤單一 log 檔：只讀新增部分；檔案被清除/重建/輪替時自動由頭開始。"""

    def __init__(self, key, path):
        self.key, self.path = key, path
        self.offset = None          # None = 未開始
        self.ident = None
        self.remainder = b''
        self.last_level = 'INFO'
        self.last_ts = None

    def poll(self):
        """回傳 (reset: bool, entries: list[LogEntry])"""
        try:
            st = os.stat(self.path)
        except OSError:
            if self.offset:
                self.offset, self.remainder = 0, b''
                return True, []
            return False, []
        ident = getattr(st, 'st_birthtime', None) or st.st_ctime
        reset = False
        if self.offset is None:
            start = max(0, st.st_size - INITIAL_TAIL_BYTES)
            skip_partial = start > 0
            self.ident = ident
        elif st.st_size < self.offset or ident != self.ident:
            start, skip_partial, reset = 0, False, True
            self.ident, self.remainder = ident, b''
        else:
            start, skip_partial = self.offset, False
        if st.st_size == start and not reset:
            self.offset = start
            return False, []
        try:
            with open_shared(self.path) as f:
                f.seek(start)
                data = f.read(MAX_READ_PER_POLL)
        except OSError:
            return reset, []
        self.offset = start + len(data)
        data = self.remainder + data
        lines = data.split(b'\n')
        self.remainder = lines.pop()  # 最後一段可能未寫完
        if skip_partial and lines:
            lines.pop(0)
        entries = []
        for raw in lines:
            line = raw.decode('utf-8', errors='replace').rstrip('\r')
            if not line:
                continue
            m = LINE_RE.match(line)
            if m:
                self.last_ts, self.last_level = m.group(1), m.group(3)
                entries.append(LogEntry(self.key, m.group(1), m.group(3), m.group(4), line))
            else:  # traceback 等多行內容，沿用上一行等級
                entries.append(LogEntry(self.key, self.last_ts, self.last_level, line, line))
        return reset, entries


# ---------------------------------------------------------------------------
# 由 log 推斷目前狀態（不依賴 Tk，可單獨測試）
# ---------------------------------------------------------------------------
RE_EXT_START = re.compile(r'\[Extractor\] 開始處理任務 \(索引: (\d+)\), Video URL: (.*?), TMDB URL: (.*)$')
RE_UP_START = re.compile(r'\[Uploader\] 開始處理上傳任務 \(索引: (\d+)\), TMDB URL: (.*)$')
RE_INDEX = re.compile(r'索引: (\d+)')
RE_DONE = re.compile(r'已將 Done 值更新 \(row_index=(\d+), value=(-?\d+)\)')
RE_BATCH_OK = re.compile(r'成功批次寫入 (\d+) 個 Done 值')
RE_SHEET = re.compile(r"\[工作表\] 開始處理工作表 '(.+?)'（第 (\d+)/(\d+) 張）")


def split_sheets(text):
    """「Bulk, Eng」→ ['Bulk', 'Eng']（同主程式一樣用逗號 / 分號分隔）。"""
    return [p.strip() for p in re.split(r'[,，;；]', text or '') if p.strip()]


def norm_url(v):
    s = '' if v is None else str(v).strip()
    return '' if s.lower() in ('', 'nan', 'none') else s
PHASES = [
    ('新一輪任務開始', '掃描 Excel'),
    ('等待當前批次任務處理中', '處理中'),
    ('[定期掃描] 發現', '處理中（發現新增項目）'),
    ('所有當前任務已處理完畢', '等待重新掃描'),
    ('所有任務已完成，程式將自動結束', '全部完成，正在結束'),
    ('主執行緒捕獲到 Ctrl+C', '收到 Ctrl+C，正在停止'),
    ('[Ctrl+C]', '收到 Ctrl+C，正在停止'),
    ('執行緒逾時仍未結束', '執行緒卡住，強制結束'),
    ('程序已安全結束', '已結束'),
    ('進入 Single 模式', 'Single 模式'),
]


def parse_ts(ts):
    try:
        return dt.datetime.strptime(ts, '%Y-%m-%d %H:%M:%S')
    except Exception:
        return None


class RunState:
    def __init__(self):
        self.workers = {}
        self.reset_all()

    def reset_all(self):
        self.workers = {'importer': self._idle('未開始')}
        for i in range(NUM_EXTRACTORS):
            self.workers[f'e{i}'] = self._idle('未開始')
        self.row_live = {}          # row_index -> '提取中' / '等待上傳' / '上傳中'
        # row_index -> {'val', 'urls': (video, tmdb), 'written'}（log 已記錄 Done，可能未寫入 Excel）
        self.done_pending = {}
        self.row_urls = {}          # row_index -> (Video URL, TMDB URL)，取自 log 的「開始處理任務」
        self.current_sheet = None   # 主程式 log 顯示處理緊嘅工作表（None = 未知，用面板選項第一張）
        self.sheet_progress = ''    # 例如「第 1/2 張」
        self.counts = collections.Counter()
        self.phase = '—'
        self.last_ts = None
        self.events = collections.deque(maxlen=MAX_EVENTS)
        self.changed_rows = set()

    @staticmethod
    def _idle(state='閒置'):
        return {'state': state, 'row': None, 'url': '', 'since': None}

    def reset_source(self, src):
        if src == 'importer':
            self.reset_all()
        else:
            self.workers[src] = self._idle('未開始')

    def _event(self, e, text):
        self.events.append(((e.ts or '')[11:], e.src, text))

    def _set_live(self, row, status):
        if status is None:
            self.row_live.pop(row, None)
        else:
            self.row_live[row] = status
        self.changed_rows.add(row)

    def feed(self, e):
        t = e.text
        if e.ts:
            self.last_ts = e.ts
        w = self.workers.get(e.src)
        if w is not None and w['state'] == '未開始':
            w['state'] = '閒置'

        m = RE_DONE.search(t)
        if m:
            row, val = int(m.group(1)), int(m.group(2))
            prev = self.done_pending.get(row)
            self.done_pending[row] = {'val': val, 'urls': self.row_urls.get(row), 'written': False}
            self.changed_rows.add(row)
            if prev is None or prev['val'] != val:
                self.counts['done_ok' if val == 1 else 'done_fail'] += 1
                self._event(e, f'列 {row + 2} Done={val}')
            self._set_live(row, None)
            return

        m = RE_BATCH_OK.search(t)
        if m:
            # 批次寫入會寫走當時記憶體中全部待寫值 → 之前記錄的 Done 均已寫入
            for row, p in self.done_pending.items():
                if not p['written']:
                    p['written'] = True
                    self.changed_rows.add(row)
            self._event(e, f'已寫入 Excel：{m.group(1)} 個 Done 值')
            return

        if '[Extractor]' in t:
            m = RE_EXT_START.search(t)
            if m:
                row = int(m.group(1))
                self.row_urls[row] = (norm_url(m.group(2)), norm_url(m.group(3)))
                url = m.group(2) if m.group(2) not in ('', 'nan', 'None') else m.group(3)
                self.workers[e.src] = {'state': '提取中', 'row': row, 'url': url, 'since': e.ts}
                self._set_live(row, '提取中')
                self.counts['extract_started'] += 1
                self._event(e, f'開始提取 列 {row + 2}  {url}')
                return
            m = RE_INDEX.search(t)
            if m and any(k in t for k in ('提取成功', '提取失敗', '跳過', '終止', '讀取任務')):
                row = int(m.group(1))
                if w is not None and w.get('row') == row:
                    self.workers[e.src] = self._idle()
                if '放入上傳隊列' in t:
                    self._set_live(row, '等待上傳')
                    self._event(e, f'提取完成 列 {row + 2} → 等待上傳')
                elif '重新讀取 Excel 再試' not in t:
                    if self.row_live.get(row) == '提取中':
                        self._set_live(row, None)
                    if '提取失敗' in t:
                        self.counts['extract_failed'] += 1
                        self._event(e, f'提取失敗 列 {row + 2}')
                return
            if '提取執行緒已結束' in t:
                self.workers[e.src] = self._idle('已結束')
                return

        if '[Uploader]' in t:
            m = RE_UP_START.search(t)
            if m:
                row = int(m.group(1))
                self.workers['uploader'] = {'state': '上傳中', 'row': row, 'url': m.group(2), 'since': e.ts}
                self._set_live(row, '上傳中')
                self._event(e, f'開始上傳 列 {row + 2}  {m.group(2)}')
                return
            m = RE_INDEX.search(t)
            if m and any(k in t for k in ('上傳成功', '上傳失敗', '跳過', '無需上傳')):
                row = int(m.group(1))
                up = self.workers.get('uploader')
                if up and up.get('row') == row:
                    self.workers['uploader'] = self._idle()
                if self.row_live.get(row) in ('上傳中', '等待上傳'):
                    self._set_live(row, None)
                if '上傳成功' in t:
                    self.counts['uploaded'] += 1
                    self._event(e, f'上傳成功 列 {row + 2}')
                elif '上傳失敗' in t:
                    self.counts['upload_failed'] += 1
                    self._event(e, f'上傳失敗 列 {row + 2}')
                return
            if '上傳執行緒已結束' in t:
                self.workers['uploader'] = self._idle('已結束')
                return
            if 'uploader' not in self.workers or self.workers['uploader']['state'] == '未開始':
                self.workers['uploader'] = self._idle()

        if e.src == 'importer':
            m = RE_SHEET.search(t)
            if m:
                # 轉工作表：行號由 0 重新計，之前嘅行狀態唔再適用
                self.current_sheet = m.group(1)
                self.sheet_progress = f'第 {m.group(2)}/{m.group(3)} 張'
                for row in set(self.row_live) | set(self.done_pending) | set(self.row_urls):
                    self.changed_rows.add(row)
                self.row_live, self.done_pending, self.row_urls = {}, {}, {}
                self._event(e, f'開始處理工作表 {m.group(1)}（{self.sheet_progress}）')
                return
            for key, phase in PHASES:
                if key in t:
                    self.phase = phase
                    self._event(e, phase)
                    if phase == '已結束':  # 程式已結束，不會再有進行中項目
                        for row in list(self.row_live):
                            self._set_live(row, None)
                    break


# ---------------------------------------------------------------------------
# 背景資料收集（log / Excel / 程式狀態 / AI 快取）
# ---------------------------------------------------------------------------
def add_user_site():
    """新電腦首次運行：面板開啟時 user site 未存在，Python 不會加入 sys.path。
    start.py 用 pip --user 裝好套件後在此補上，面板不用重開就認到 psutil / openpyxl。"""
    import importlib
    import site
    try:
        user_site = site.getusersitepackages()
        if user_site not in sys.path and os.path.isdir(user_site):
            site.addsitedir(user_site)
            importlib.invalidate_caches()
    except Exception:
        pass


CONSOLE_LOG_REL = os.path.join('data', 'monitor_console.log')   # 由面板啟動時 start.py 的畫面輸出
MAX_CONSOLE_LINES = 20000
SAME_AS_SERIES = '同劇集'   # 分集選項：自動代答時送出空白（start.py 默認沿用劇集設定）
# (設定鍵, 標籤, 可選值, 預設) —— 對應 start.py 啟動時的提問
RUN_OPTIONS = [
    # 劇集 = 系列/季/電影；分集 = 每一集（「同劇集」= 沿用劇集設定）
    ('overwrite_titles', '覆蓋劇集標題', ['ask', 'y', 'n'], 'ask'),
    ('overwrite_descriptions', '覆蓋劇集描述', ['ask', 'y', 'n'], 'ask'),
    ('episode_overwrite_titles', '覆蓋分集標題', [SAME_AS_SERIES, 'ask', 'y', 'n'], SAME_AS_SERIES),
    ('episode_overwrite_descriptions', '覆蓋分集描述', [SAME_AS_SERIES, 'ask', 'y', 'n'], SAME_AS_SERIES),
    # 來源（影片網站）無資料而 TMDB 有內容：n = 保留 TMDB（簡體轉繁體）；y = 覆蓋=y 時清空 TMDB
    ('clear_when_source_empty', '來源無資料時清空TMDB', ['n', 'y'], 'n'),
    # 西方人名：n = TMDB 已有香港人名就保留；y = 用中文維基 zh-hk 譯名取代
    ('wiki_name_replace', '維基人名取代現有', ['n', 'y'], 'n'),
    # 【2026-11-14】bulk：上次已提取未上傳嘅項目保留直接上傳（只有上次有剩先會問；程式會檢查該行內容同啟動選項冇改）
    ('keep_pending_uploads', '保留上次未上傳', ['n', 'y'], 'n'),
    ('mode', '模式', ['bulk', 'single'], 'bulk'),
    # bulk 模式處理 bulk_videos.xlsx 邊幾張工作表（逗號分隔，按次序處理；用「選擇…」對話框揀）
    ('bulk_sheet', 'bulk 工作表', [], 'Bulk'),
]
# 提問文字 → 用哪個設定自動回答（start.py 的提問字眼改了要同步更新）
AUTO_ANSWERS = [('是否覆蓋劇集標題', 'overwrite_titles'), ('是否覆蓋劇集描述', 'overwrite_descriptions'),
                ('是否覆蓋分集標題', 'episode_overwrite_titles'), ('是否覆蓋分集描述', 'episode_overwrite_descriptions'),
                ('來源無資料時是否清空', 'clear_when_source_empty'), ('維基香港譯名取代', 'wiki_name_replace'),
                ('運行模式', 'mode'), ('處理 bulk_videos.xlsx 邊一張工作表', 'bulk_sheet'),
                ('保留上次已提取', 'keep_pending_uploads')]
SECRET_PROMPT_RE = re.compile(r'password|密碼|驗證碼', re.I)
# 沒有提問文字但正等待 Enter 的情況（例如「請關閉 Excel 後按 Enter」）
ENTER_HINT_RE = re.compile(r'按 ?Enter')


class ConsoleTailer:
    """追蹤 start.py 的畫面輸出（純文字，不一定以換行結尾，例如 input() 的提問）。"""

    def __init__(self, path):
        self.path = path
        self.pos = None
        self.dec = None

    def reset(self):
        self.pos = 0
        self.dec = codecs.getincrementaldecoder('utf-8')('replace')

    def poll(self):
        """回傳 (是否重新開始, 新文字)。"""
        try:
            size = os.stat(self.path).st_size
        except OSError:
            return False, ''
        restarted = False
        if self.pos is None:                       # 首次：只讀最後一段
            self.reset()
            self.pos = max(0, size - INITIAL_TAIL_BYTES)
        elif size < self.pos:                      # 檔案被截斷 → 新一次運行
            self.reset()
            restarted = True
        if size == self.pos:
            return restarted, ''
        with open_shared(self.path) as f:
            f.seek(self.pos)
            data = f.read(MAX_READ_PER_POLL)
        self.pos += len(data)
        text = self.dec.decode(data).replace('\r\n', '\n').replace('\r', '\n')
        return restarted, text


class Collector(threading.Thread):
    def __init__(self, base, out_q):
        super().__init__(daemon=True, name='Collector')
        self.base, self.q = base, out_q
        self.console = ConsoleTailer(os.path.join(base, CONSOLE_LOG_REL))
        self.console_reset = threading.Event()
        log_dir = os.path.join(base, 'log')
        self.tailers = [LogTailer('importer', os.path.join(log_dir, 'importer.log'))]
        self.tailers += [LogTailer(f'e{i}', os.path.join(log_dir, f'extractor_{i}.log'))
                         for i in range(NUM_EXTRACTORS)]
        self.bulk_path = os.path.join(base, 'data', 'bulk_videos.xlsx')
        self.cache_path = os.path.join(base, 'data', 'ai_s2t_cache.db')
        self.bulk_mtime = None
        self.bulk_sheet = 'Bulk'          # 進度顯示用嘅工作表（跟面板「bulk 工作表」選項）
        self.force_excel = threading.Event()
        self.stop = threading.Event()

    def run(self):
        last_slow = 0
        while not self.stop.is_set():
            batch = []
            for t in self.tailers:
                try:
                    reset, entries = t.poll()
                except Exception as e:  # 不讓單一檔案錯誤停止整個面板
                    reset, entries = False, []
                    self.q.put(('error', f'讀取 {t.path} 失敗: {e}'))
                if reset:
                    self.q.put(('reset', t.key))
                batch.extend(entries)
            if batch:
                # 多個 log 檔按時間合併（穩定排序：同一秒內保持各檔原有次序），狀態推斷才正確
                batch.sort(key=lambda e: e.ts or '')
                self.q.put(('lines', batch))
            try:
                if self.console_reset.is_set():
                    self.console_reset.clear()
                    self.console.reset()
                    self.q.put(('console', (True, '')))
                restarted, text = self.console.poll()
                if restarted or text:
                    self.q.put(('console', (restarted, text)))
            except Exception as e:
                self.q.put(('error', f'讀取主控台輸出失敗: {e}'))
            now = time.monotonic()
            if now - last_slow >= 3 or self.force_excel.is_set():
                last_slow = now
                self.q.put(('proc', self._find_process()))
                self.q.put(('cache', self._cache_count()))
                rej = self._read_rejected()
                if rej is not None:
                    self.q.put(('rejected', rej))
                self._read_excel(force=self.force_excel.is_set())
                self.force_excel.clear()
            self.stop.wait(1.0)

    def _find_process(self):
        try:
            import psutil
        except ImportError:
            add_user_site()
            try:
                import psutil
            except ImportError:
                # 仍要讓「啟動」可按：psutil 正是由 start.py 首次運行時安裝
                return {'procs': [], 'error': '未安裝 psutil，首次啟動程式時會自動安裝'}
        base = os.path.normcase(os.path.abspath(self.base))
        found = []
        for p in psutil.process_iter(['pid', 'name']):   # 只取名稱（便宜），python 行程才讀 cmdline
            try:
                if p.pid == os.getpid() or 'python' not in (p.info['name'] or '').lower():
                    continue
                cmd = ' '.join(p.cmdline() or [])
                if not re.search(r'(start|tmdb_importer)\.py\b', cmd):
                    continue
                try:
                    if os.path.normcase(p.cwd()) != base:
                        continue
                except (psutil.AccessDenied, psutil.NoSuchProcess):
                    pass
                found.append({'pid': p.pid, 'create_time': p.create_time()})
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return {'procs': found}

    def _cache_count(self):
        if not os.path.exists(self.cache_path):
            return None
        try:
            uri = 'file:' + self.cache_path.replace('\\', '/') + '?mode=ro'
            con = sqlite3.connect(uri, uri=True, timeout=2)
            try:
                return con.execute('SELECT COUNT(*) FROM cache').fetchone()[0]
            finally:
                con.close()
        except Exception:
            return None

    def _read_rejected(self):
        """讀取 AI 被拒字詞（ai_s2t_cache.db 的 rejected 表）。無變化時回傳 None。"""
        if not os.path.exists(self.cache_path):
            return None
        try:
            uri = 'file:' + self.cache_path.replace('\\', '/') + '?mode=ro'
            con = sqlite3.connect(uri, uri=True, timeout=2)
            try:
                sig = con.execute('SELECT COUNT(*), SUM(count), MAX(last_seen) FROM rejected').fetchone()
                if sig == getattr(self, '_rej_sig', None):
                    return None
                rows = con.execute(
                    'SELECT count, draft_word, ai_word, draft_ctx, ai_ctx, model, last_seen, first_seen, '
                    'src, final_text, ai_text FROM rejected ORDER BY count DESC, last_seen DESC LIMIT 10000'
                ).fetchall()
            finally:
                con.close()
        except Exception:          # 表未建立（舊版主程式）或資料庫忙碌：稍後再試
            return None
        self._rej_sig = sig
        return rows

    def _read_excel(self, force=False):
        try:
            mtime = os.stat(self.bulk_path).st_mtime
        except OSError:
            self.q.put(('excel_error', '找不到 data/bulk_videos.xlsx'))
            return
        if mtime == self.bulk_mtime and not force:
            return
        try:
            add_user_site()
            import openpyxl
            with open_shared(self.bulk_path) as f:
                data = f.read()
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            sheet = self.bulk_sheet
            try:
                if sheet not in wb.sheetnames:
                    self.q.put(('excel_error', f'bulk_videos.xlsx 冇工作表「{sheet}」'))
                    return
                ws = wb[sheet]
                rows = [tuple('' if v is None else v for v in r) for r in ws.iter_rows(values_only=True)]
            finally:
                wb.close()
        except Exception as e:  # Excel 儲存中等暫時錯誤：保留上次結果，稍後再試
            self.q.put(('excel_error', f'讀取 Excel 失敗（稍後重試）: {e}'))
            return
        self.bulk_mtime = mtime
        header = [str(h).strip() for h in rows[0]] if rows else []
        self.q.put(('excel', (header, rows[1:], time.strftime('%H:%M:%S'), sheet)))


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
LEVEL_ORDER = {'DEBUG': 10, 'INFO': 20, 'WARNING': 30, 'ERROR': 40, 'CRITICAL': 50}
LEVEL_FILTERS = [('全部', 0), ('INFO 以上', 20), ('WARNING 以上', 30), ('只限 ERROR', 40)]
SRC_LABEL = {'importer': 'IMP', **{f'e{i}': f'E{i}' for i in range(NUM_EXTRACTORS)}}
STATUS_COLORS = {
    '提取中': '#dbeafe', '等待上傳': '#fef9c3', '上傳中': '#e0e7ff',
    '完成': '#ffffff', '完成 (未寫入)': '#ecfdf5', '失敗': '#fee2e2', '待處理': '#fff7ed',
}


class LogView(ttk.Frame):
    """一個 log 分頁：搜尋、等級篩選、自動捲動、暫停。"""

    def __init__(self, master, app, show_src=False, file_path=None, accept=None):
        super().__init__(master)
        self.app, self.show_src, self.file_path = app, show_src, file_path
        self.accept = accept or (lambda e: True)
        self.store = collections.deque(maxlen=MAX_VIEW_LINES)
        self.paused_buffer = []
        self.unseen = 0

        bar = ttk.Frame(self)
        bar.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(bar, text='搜尋').pack(side='left')
        self.search_var = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.search_var, width=28)
        ent.pack(side='left', padx=(4, 10))
        self.search_var.trace_add('write', lambda *_: self._schedule_rebuild())
        ttk.Label(bar, text='等級').pack(side='left')
        self.level_var = tk.StringVar(value=LEVEL_FILTERS[0][0])
        cb = ttk.Combobox(bar, textvariable=self.level_var, state='readonly', width=13,
                          values=[n for n, _ in LEVEL_FILTERS])
        cb.pack(side='left', padx=(4, 10))
        cb.bind('<<ComboboxSelected>>', lambda _: self.rebuild())
        self.auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text='自動捲動', variable=self.auto_var).pack(side='left', padx=4)
        self.pause_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text='暫停', variable=self.pause_var, command=self._on_pause).pack(side='left', padx=4)
        ttk.Button(bar, text='清空畫面', command=self.clear).pack(side='left', padx=4)
        if file_path:
            ttk.Button(bar, text='開啟檔案', command=lambda: app.open_path(file_path)).pack(side='left', padx=4)
        self.info = ttk.Label(bar, text='', foreground='#6b7280')
        self.info.pack(side='right')

        body = ttk.Frame(self)
        body.pack(fill='both', expand=True, padx=6, pady=(0, 6))
        self.text = tk.Text(body, wrap='none', font=app.log_font, undo=False, state='disabled',
                            background='#fbfbfc', relief='flat', borderwidth=1, highlightthickness=1,
                            highlightbackground='#d1d5db')
        ys = ttk.Scrollbar(body, orient='vertical', command=self.text.yview)
        xs = ttk.Scrollbar(body, orient='horizontal', command=self.text.xview)
        self.text.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.text.grid(row=0, column=0, sticky='nsew')
        ys.grid(row=0, column=1, sticky='ns')
        xs.grid(row=1, column=0, sticky='ew')
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.text.tag_configure('ERROR', foreground='#b91c1c')
        self.text.tag_configure('CRITICAL', foreground='#ffffff', background='#b91c1c')
        self.text.tag_configure('WARNING', foreground='#b45309')
        self.text.tag_configure('DEBUG', foreground='#9ca3af')
        self.text.tag_configure('INFO', foreground='#111827')
        self.text.tag_configure('AI', foreground='#7c3aed')
        self.text.tag_configure('src', foreground='#2563eb')
        self.text.tag_configure('hit', background='#fde68a')
        self._rebuild_job = None

    # -- 過濾
    def _min_level(self):
        return dict(LEVEL_FILTERS).get(self.level_var.get(), 0)

    def _filter(self, entries):
        """一次過讀取篩選設定（避免逐行讀取 Tk 變數），回傳符合的行。"""
        min_lv = self._min_level()
        q = self.search_var.get().strip().lower()
        return [e for e in entries
                if LEVEL_ORDER.get(e.level, 20) >= min_lv and (not q or q in e.raw.lower())]

    def _schedule_rebuild(self):
        if self._rebuild_job:
            self.after_cancel(self._rebuild_job)
        self._rebuild_job = self.after(250, self.rebuild)

    # -- 寫入
    def add(self, entries):
        entries = [e for e in entries if self.accept(e)]
        if not entries:
            return
        self.store.extend(entries)
        if self.pause_var.get():
            self.paused_buffer.extend(entries)
            self.info.configure(text=f'已暫停：{len(self.paused_buffer)} 行未顯示')
            return
        if not self.winfo_ismapped():
            # 分頁未被顯示：只記入 store，切換到此分頁時才重繪（避免背景分頁拖慢介面）
            self._dirty = True
            return
        self._insert(self._filter(entries))

    def on_shown(self):
        if getattr(self, '_dirty', False):
            self._dirty = False
            self.rebuild()

    def _insert(self, entries):
        if not entries:
            return
        entries = entries[-MAX_VIEW_LINES:]   # 畫面最多保留 MAX_VIEW_LINES 行，舊的不必插入
        t = self.text
        t.configure(state='normal')
        q = self.search_var.get().strip()
        start = t.index('end-1c')
        args = []                              # 一次過 insert（多段文字 + tag），比逐行快很多
        for e in entries:
            if self.show_src:
                args += [f'[{SRC_LABEL.get(e.src, e.src)}] ', 'src']
            tag = 'AI' if ('[AI簡轉繁]' in e.text and e.level == 'INFO') else e.level
            args += [e.raw + '\n', tag]
        t.insert('end', *args)
        if q:
            self._highlight(start, q)
        lines = int(t.index('end-1c').split('.')[0])
        if lines > MAX_VIEW_LINES + 500:
            t.delete('1.0', f'{lines - MAX_VIEW_LINES}.0')
        t.configure(state='disabled')
        if self.auto_var.get():
            t.see('end')
        self.info.configure(text=f'{min(len(self.store), MAX_VIEW_LINES)} 行')

    def _highlight(self, start, q):
        idx = start
        while True:
            idx = self.text.search(q, idx, nocase=True, stopindex='end')
            if not idx:
                break
            end = f'{idx}+{len(q)}c'
            self.text.tag_add('hit', idx, end)
            idx = end

    def rebuild(self):
        self._rebuild_job = None
        self.text.configure(state='normal')
        self.text.delete('1.0', 'end')
        self.text.configure(state='disabled')
        self._insert(self._filter(self.store))

    def _on_pause(self):
        if not self.pause_var.get():
            self.paused_buffer = []
            self.rebuild()

    def clear(self):
        self.store.clear()
        self.paused_buffer = []
        self.rebuild()


class MonitorApp:
    def __init__(self, root, base):
        self.root, self.base = root, base
        self.state = RunState()
        self.q = queue.Queue()
        self.excel_header, self.excel_rows = [], []
        self.excel_info = '尚未讀取'
        self.proc_info = {}
        self.cache_count = None
        self.last_error = ''

        self.bulk_path = os.path.join(base, 'data', 'bulk_videos.xlsx')
        self._setup_style()
        root.title(f'TMDB 監控面板 — {base}')
        self.switch_to = None
        root.geometry('1320x820')
        root.minsize(900, 560)

        self._build_header()
        self.nb = ttk.Notebook(root)
        self.nb.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self._build_overview()
        self._build_console_tab()
        self._build_excel_tab()
        self.views = {}
        log_dir = os.path.join(base, 'log')
        self.views['all'] = LogView(self.nb, self, show_src=True)
        self.nb.add(self.views['all'], text='全部 Log')
        self.views['importer'] = LogView(self.nb, self, file_path=os.path.join(log_dir, 'importer.log'),
                                         accept=lambda e: e.src == 'importer')
        self.nb.add(self.views['importer'], text='Importer / 上傳')
        for i in range(NUM_EXTRACTORS):
            key = f'e{i}'
            v = LogView(self.nb, self, file_path=os.path.join(log_dir, f'extractor_{i}.log'),
                        accept=lambda e, k=key: e.src == k)
            self.views[key] = v
            self.nb.add(v, text=f'Extractor {i}')
        self.views['problems'] = LogView(self.nb, self, show_src=True,
                                         accept=lambda e: LEVEL_ORDER.get(e.level, 20) >= 30)
        self.nb.add(self.views['problems'], text='⚠ 警告 / 錯誤')
        self.views['ai'] = LogView(self.nb, self, show_src=True, accept=lambda e: '[AI簡轉繁]' in e.text and '個觸發字/詞' not in e.text)
        self.nb.add(self.views['ai'], text='AI 簡轉繁')
        self._build_rejected_tab()
        self.problem_count = 0
        self.nb.bind('<<NotebookTabChanged>>', self._on_tab_changed)

        self.collector = Collector(base, self.q)
        self.collector.bulk_sheet = self._wanted_excel_sheet()
        self.collector.start()
        root.protocol('WM_DELETE_WINDOW', self.close)
        self.root.after(200, self._drain)
        self.root.after(1000, self._tick)
        self.root.after(1500, self._check_bangumi_token)

    def _check_bangumi_token(self):
        """【2026-11-13】Bangumi token 到期前 30 日起，每次開監控面板都彈出提醒。"""
        try:
            from others.bangumi_token import token_expiry_message
            msg = token_expiry_message(os.path.join(self.base, 'configs.json'))
        except Exception:
            return
        if msg:
            from tkinter import messagebox
            messagebox.showwarning('TMDB 監控面板：Bangumi token', msg, parent=self.root)

    # ---------------- 外觀
    def _setup_style(self):
        families = set(tkfont.families())
        ui = next((f for f in ('Microsoft JhengHei UI', 'Microsoft JhengHei', 'PingFang TC',
                               'Noto Sans CJK TC') if f in families), None)
        mono = next((f for f in ('Cascadia Mono', 'Consolas', 'Menlo') if f in families), 'TkFixedFont')
        for name in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont', 'TkHeadingFont'):
            try:
                f = tkfont.nametofont(name)
                if ui:
                    f.configure(family=ui)
                f.configure(size=10)
            except tk.TclError:
                pass
        self.log_font = tkfont.Font(family=mono, size=10)
        self.big_font = tkfont.Font(family=ui or 'TkDefaultFont', size=15, weight='bold')
        self.small_font = tkfont.Font(family=ui or 'TkDefaultFont', size=9)
        style = ttk.Style(self.root)
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        rh = tkfont.nametofont('TkDefaultFont').metrics('linespace') + 8
        style.configure('Treeview', rowheight=rh)
        style.configure('Card.TFrame', background='#ffffff', relief='solid', borderwidth=1)
        style.configure('CardTitle.TLabel', background='#ffffff', foreground='#6b7280', font=self.small_font)
        style.configure('CardValue.TLabel', background='#ffffff', foreground='#111827', font=self.big_font)
        style.configure('CardSub.TLabel', background='#ffffff', foreground='#6b7280', font=self.small_font)

    def _card(self, parent, title):
        f = ttk.Frame(parent, style='Card.TFrame', padding=(12, 8))
        ttk.Label(f, text=title, style='CardTitle.TLabel').pack(anchor='w')
        val = ttk.Label(f, text='—', style='CardValue.TLabel')
        val.pack(anchor='w')
        sub = ttk.Label(f, text='', style='CardSub.TLabel')
        sub.pack(anchor='w')
        return f, val, sub

    def _on_option_changed(self):
        save_settings({k: var.get() for k, var in self.opt_vars.items()})
        self._sync_excel_sheet()

    def _wanted_excel_sheet(self):
        """Excel 進度要顯示嘅工作表：主程式 log 講緊處理邊張就用嗰張，否則用面板選項第一張。"""
        st = getattr(self, 'state', None)
        if st is not None and st.current_sheet:
            return st.current_sheet
        return (split_sheets(self.opt_vars['bulk_sheet'].get()) or ['Bulk'])[0]

    def _sync_excel_sheet(self):
        collector = getattr(self, 'collector', None)
        if collector is None:
            return
        sheet = self._wanted_excel_sheet()
        if sheet != collector.bulk_sheet:
            collector.bulk_sheet = sheet
            collector.force_excel.set()

    def _choose_bulk_sheets(self):
        """揀 bulk 模式處理邊幾張工作表同次序（左邊全部工作表，右邊 = 已選，由上至下處理）。"""
        names = list_bulk_sheets(self.bulk_path)
        if not names:
            self.flash('讀取唔到 bulk_videos.xlsx 嘅工作表清單', 6)
            return
        chosen = [s for s in split_sheets(self.opt_vars['bulk_sheet'].get()) if s in names]
        win = tk.Toplevel(self.root)
        win.title('揀 bulk 工作表（右邊由上至下按次序處理）')
        win.transient(self.root)
        win.grab_set()
        body = ttk.Frame(win, padding=10)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text='全部工作表').grid(row=0, column=0, sticky='w')
        ttk.Label(body, text='處理次序').grid(row=0, column=2, sticky='w')
        left = tk.Listbox(body, height=16, width=30, exportselection=False, selectmode='extended')
        right = tk.Listbox(body, height=16, width=30, exportselection=False)
        left.grid(row=1, column=0, sticky='nsew')
        right.grid(row=1, column=2, sticky='nsew')
        body.columnconfigure(0, weight=1)
        body.columnconfigure(2, weight=1)
        body.rowconfigure(1, weight=1)

        def refresh():
            left.delete(0, 'end')
            for n in names:
                left.insert('end', n)
            right.delete(0, 'end')
            for i, n in enumerate(chosen, 1):
                right.insert('end', f'{i}. {n}')

        def add(_e=None):
            for i in left.curselection():
                if names[i] not in chosen:
                    chosen.append(names[i])
            refresh()

        def remove(_e=None):
            sel = right.curselection()
            if sel:
                chosen.pop(sel[0])
                refresh()
                if chosen:
                    right.selection_set(min(sel[0], len(chosen) - 1))

        def move(delta):
            sel = right.curselection()
            if not sel:
                return
            i, j = sel[0], sel[0] + delta
            if 0 <= j < len(chosen):
                chosen[i], chosen[j] = chosen[j], chosen[i]
                refresh()
                right.selection_set(j)

        mid = ttk.Frame(body, padding=(8, 0))
        mid.grid(row=1, column=1)
        for text, cmd in (('加入 →', add), ('← 移除', remove), ('↑ 上移', lambda: move(-1)), ('↓ 下移', lambda: move(1))):
            ttk.Button(mid, text=text, width=8, command=cmd).pack(pady=3)
        left.bind('<Double-Button-1>', add)
        right.bind('<Double-Button-1>', remove)

        def ok():
            self.opt_vars['bulk_sheet'].set(', '.join(chosen) or 'Bulk')
            self._on_option_changed()
            win.destroy()

        bottom = ttk.Frame(win, padding=(10, 0, 10, 10))
        bottom.pack(fill='x')
        ttk.Label(bottom, text='冇揀 = Bulk。程式運行中改，要下次啟動先生效。').pack(side='left')
        ttk.Button(bottom, text='取消', command=win.destroy).pack(side='right')
        ttk.Button(bottom, text='確定', command=ok).pack(side='right', padx=6)
        refresh()

    # ---------------- 頂部摘要
    def _build_header(self):
        top = ttk.Frame(self.root, padding=(8, 8, 8, 6))
        top.pack(fill='x')
        cards = [('程式狀態', 'proc'), ('目前階段', 'phase'), ('Excel 進度', 'excel'),
                 ('近期活動', 'session'), ('AI 快取', 'cache')]
        self.cards = {}
        for i, (title, key) in enumerate(cards):
            f, val, sub = self._card(top, title)
            f.grid(row=0, column=i, sticky='nsew', padx=(0 if i == 0 else 6, 0))
            top.columnconfigure(i, weight=1, uniform='cards')
            self.cards[key] = (val, sub)
        self.progress = ttk.Progressbar(self.cards['excel'][0].master, mode='determinate', length=160)
        self.progress.pack(anchor='w', fill='x', pady=(4, 0))
        # 第一行：啟動選項（視窗窄時自動換行）；第二行：啟動/停止按鈕
        opts_frame = ttk.Frame(self.root, padding=(8, 0, 8, 4))
        opts_frame.pack(fill='x')
        saved = load_settings()
        self.opt_vars = {}
        opt_cells = []
        for key, label, values, default in RUN_OPTIONS:
            cell = ttk.Frame(opts_frame)
            ttk.Label(cell, text=label).pack(side='left', padx=(0, 3))
            if key == 'bulk_sheet':
                v = tk.StringVar(value=', '.join(split_sheets(saved.get(key))) or default)
                ttk.Entry(cell, textvariable=v, state='readonly', width=24).pack(side='left')
                ttk.Button(cell, text='選擇…', width=6,
                           command=self._choose_bulk_sheets).pack(side='left', padx=(2, 12))
                self.opt_vars[key] = v
                opt_cells.append(cell)
                continue
            v = tk.StringVar(value=saved.get(key) if saved.get(key) in values else default)
            cb = ttk.Combobox(cell, textvariable=v, values=values, state='readonly', width=7)
            cb.pack(side='left', padx=(0, 12))
            cb.bind('<<ComboboxSelected>>', lambda _e: self._on_option_changed())
            self.opt_vars[key] = v
            opt_cells.append(cell)

        def _reflow(_e=None, cells=opt_cells, frame=opts_frame):
            """按目前寬度重新排列選項，放不下就換行。"""
            width = max(frame.winfo_width() - 16, 1)
            x = y = row_h = 0
            for c in cells:
                w, h = c.winfo_reqwidth(), c.winfo_reqheight()
                if x and x + w > width:
                    x, y, row_h = 0, y + row_h + 2, 0
                c.place(x=x, y=y)
                x, row_h = x + w, max(row_h, h)
            total = y + row_h
            if total and abs(frame.winfo_reqheight() - total - 4) > 1:
                frame.configure(height=total + 4)
        opts_frame.configure(height=30)
        opts_frame.bind('<Configure>', _reflow)
        self.root.after(50, _reflow)

        run = ttk.Frame(self.root, padding=(8, 0, 8, 6))
        run.pack(fill='x')
        self.start_btn = ttk.Button(run, text='▶ 啟動', command=self.start_program, state='disabled')
        self.start_btn.pack(side='left', padx=(0, 6))
        self.stop_btn = ttk.Button(run, text='■ 停止 (Ctrl+C)', command=self.stop_program, state='disabled')
        self.stop_btn.pack(side='left', padx=(0, 6))
        self.pause_path = os.path.join(self.base, 'data', '.pause')   # 與 utils.PAUSE_FLAG 相同
        self.pause_btn = ttk.Button(run, text='⏸ 暫停', command=self.toggle_pause, state='disabled')
        self.pause_btn.pack(side='left', padx=(0, 6))
        self.upload_pause_path = os.path.join(self.base, 'data', '.pause_upload')   # 與 utils.UPLOAD_PAUSE_FLAG 相同
        self.upload_pause_btn = ttk.Button(run, text='⏸ 暫停上傳', command=self.toggle_upload_pause,
                                           state='disabled')
        self.upload_pause_btn.pack(side='left', padx=(0, 6))
        ttk.Button(run, text='在命令視窗啟動', command=lambda: self.start_program(in_cmd=True)).pack(side='left')
        ttk.Label(run, text='（選項會自動代答 start.py 啟動提問；single 模式每次輸入 TMDB URL 時套用最新覆蓋選項）',
                  foreground='#6b7280', font=self.small_font).pack(side='left', padx=12)
        self._start_cooldown_until = 0
        self._stop_sent = 0
        self._flash_text, self._flash_until = '', 0
        btns = ttk.Frame(self.root, padding=(8, 0, 8, 4))
        btns.pack(fill='x')
        ttk.Button(btns, text='切換資料夾…', command=self.choose_folder).pack(side='left', padx=(0, 6))
        ttk.Button(btns, text='開啟 log 資料夾', command=lambda: self.open_path(os.path.join(self.base, 'log'))).pack(side='left')
        ttk.Button(btns, text='開啟 bulk_videos.xlsx', command=lambda: self.open_path(
            os.path.join(self.base, 'data', 'bulk_videos.xlsx'))).pack(side='left', padx=6)
        ttk.Button(btns, text='開啟 tw_hk_name.xlsx', command=lambda: self.open_path(
            os.path.join(self.base, 'data', 'tw_hk_name.xlsx'))).pack(side='left')
        self.status_label = ttk.Label(btns, text='', foreground='#6b7280')
        self.status_label.pack(side='right')

    # ---------------- 總覽分頁
    def _build_overview(self):
        page = ttk.Frame(self.nb)
        self.nb.add(page, text='總覽')
        pw = ttk.PanedWindow(page, orient='vertical')
        pw.pack(fill='both', expand=True, padx=6, pady=6)

        wf = ttk.LabelFrame(pw, text='工作執行緒（目前處理中的 Excel 項目）', padding=6)
        cols = ('thread', 'state', 'row', 'elapsed', 'url')
        self.worker_tree = ttk.Treeview(wf, columns=cols, show='headings', height=7)
        for c, label, w, anchor in (('thread', '執行緒', 110, 'w'), ('state', '狀態', 90, 'w'),
                                    ('row', 'Excel 列', 80, 'center'), ('elapsed', '已用時間', 90, 'center'),
                                    ('url', 'URL', 700, 'w')):
            self.worker_tree.heading(c, text=label)
            self.worker_tree.column(c, width=w, anchor=anchor, stretch=(c == 'url'))
        self.worker_tree.tag_configure('busy', background='#dbeafe')
        self.worker_tree.tag_configure('upload', background='#e0e7ff')
        self.worker_tree.tag_configure('ended', foreground='#9ca3af')
        self.worker_tree.pack(fill='both', expand=True)
        self.worker_tree.bind('<Double-1>', self._on_worker_double)
        for key, label in [('uploader', 'Uploader')] + [(f'e{i}', f'Extractor {i}') for i in range(NUM_EXTRACTORS)]:
            self.worker_tree.insert('', 'end', iid=key, values=(label, '未開始', '', '', ''))
        ttk.Label(wf, text='雙擊一行可跳到該執行緒的 log', foreground='#6b7280').pack(anchor='w', pady=(4, 0))
        pw.add(wf, weight=2)

        ef = ttk.LabelFrame(pw, text='最近事件', padding=6)
        self.event_tree = ttk.Treeview(ef, columns=('time', 'src', 'text'), show='headings', height=10)
        for c, label, w in (('time', '時間', 80), ('src', '來源', 60), ('text', '事件', 900)):
            self.event_tree.heading(c, text=label)
            self.event_tree.column(c, width=w, stretch=(c == 'text'))
        self.event_tree.tag_configure('bad', foreground='#b91c1c')
        self.event_tree.tag_configure('good', foreground='#047857')
        es = ttk.Scrollbar(ef, orient='vertical', command=self.event_tree.yview)
        self.event_tree.configure(yscrollcommand=es.set)
        self.event_tree.pack(side='left', fill='both', expand=True)
        es.pack(side='right', fill='y')
        pw.add(ef, weight=3)
        self._events_shown = 0

    # ---------------- AI 被拒用詞分頁
    REJ_COLS = [('count', '句數', 50), ('draft', '實際採用 (OpenCC)', 130), ('ai', 'AI 建議', 130),
                ('draft_ctx', '採用上下文', 220), ('ai_ctx', 'AI 上下文', 220), ('model', 'Model', 90),
                ('last', '最近', 140)]

    def _build_rejected_tab(self):
        page = ttk.Frame(self.nb)
        self.rej_page = page
        self.nb.add(page, text='AI 被拒用詞')
        bar = ttk.Frame(page, padding=(6, 6, 6, 0))
        bar.pack(fill='x')
        ttk.Label(bar, text='搜尋：').pack(side='left')
        self.rej_search = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.rej_search, width=24)
        ent.pack(side='left')
        self.rej_search.trace_add('write', lambda *_: self._rebuild_rejected())
        ttk.Button(bar, text='匯出 Excel 並開啟', command=self.export_rejected).pack(side='right')
        ttk.Button(bar, text='開啟 tw_hk_name.xlsx', command=lambda: self.open_path(
            os.path.join(self.base, 'data', 'tw_hk_name.xlsx'))).pack(side='right', padx=6)
        ttk.Button(bar, text='已選用 AI 清單', command=self.show_accepted).pack(side='right')
        ttk.Button(bar, text='以後用 AI 版本', command=self.accept_rejected).pack(side='right', padx=6)
        ttk.Button(bar, text='刪除', command=self.delete_rejected).pack(side='right')
        self.rej_info = ttk.Label(bar, text='', foreground='#6b7280')
        self.rej_info.pack(side='left', padx=12)
        ttk.Label(page, padding=(6, 4, 6, 0), foreground='#6b7280', font=self.small_font,
                  text='AI 建議了、但因規則「AI 只可在同一簡體字的繁體候選字之間選擇」而未被採用的字詞（同一句只計一次）。'
                       '檢查後可把合適的字詞加入 tw_hk_name.xlsx。雙擊一行看完整句子。'
                       '可按 Ctrl / Shift 多選，再按「刪除」或「以後用 AI 版本」。').pack(fill='x')
        pw = ttk.PanedWindow(page, orient='vertical')
        pw.pack(fill='both', expand=True, padx=6, pady=6)
        top = ttk.Frame(pw)
        cols = [c[0] for c in self.REJ_COLS]
        self.rej_tree = ttk.Treeview(top, columns=cols, show='headings')
        for key, title, width in self.REJ_COLS:
            self.rej_tree.heading(key, text=title)
            self.rej_tree.column(key, width=width, anchor='center' if key == 'count' else 'w',
                                 stretch=key in ('draft_ctx', 'ai_ctx'))
        sb = ttk.Scrollbar(top, command=self.rej_tree.yview)
        self.rej_tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.rej_tree.pack(fill='both', expand=True)
        self.rej_tree.bind('<<TreeviewSelect>>', self._on_rejected_select)
        pw.add(top, weight=4)
        self.rej_detail = tk.Text(pw, height=7, wrap='char', font=self.log_font)
        self.rej_detail.tag_configure('h', foreground='#6b7280')
        self.rej_detail.tag_configure('mark', background='#fde68a')
        pw.add(self.rej_detail, weight=1)
        self.rej_rows = []
        self.acc_win = None

    def _db_write(self, action):
        """在 ai_s2t_cache.db 執行 action(con) 並 commit（主程式運行中亦可）。回傳是否成功。"""
        try:
            con = sqlite3.connect(self.collector.cache_path, timeout=5)
            try:
                with con:
                    # 與 ai_s2t._rejected_init 相同；資料庫可能由未有此表的舊版主程式建立
                    con.execute('CREATE TABLE IF NOT EXISTS accepted (draft_word TEXT NOT NULL, ai_word TEXT NOT NULL, '
                                'draft_ctx TEXT, ai_ctx TEXT, created_at TEXT, PRIMARY KEY (draft_word, ai_word))')
                    action(con)
            finally:
                con.close()
            return True
        except Exception as e:
            self.flash(f'寫入 ai_s2t_cache.db 失敗: {e}', 20)
            return False

    def _selected_rejected(self, verb):
        rows = [self.rej_rows[int(i)] for i in self.rej_tree.selection()]
        if not rows:
            self.flash(f'請先揀選要{verb}的字詞（可按 Ctrl / Shift 多選）。')
        return rows

    @staticmethod
    def _pairs_preview(rows, limit=10):
        lines = [f'{r[1] or "（無）"} → {r[2] or "（無）"}' for r in rows[:limit]]
        if len(rows) > limit:
            lines.append(f'……另有 {len(rows) - limit} 組')
        return '\n'.join(lines)

    def _drop_rejected(self, rows):
        """由資料庫及畫面移除這些字詞（rejected_seen 保留，舊 log 重新匯入時不會復活）。"""
        def action(con):
            con.executemany('DELETE FROM rejected WHERE draft_word = ? AND ai_word = ?', [(r[1], r[2]) for r in rows])
        return action

    def _after_rejected_removed(self, rows):
        keys = {(r[1], r[2]) for r in rows}
        self.rej_rows = [r for r in self.rej_rows if (r[1], r[2]) not in keys]
        self._rebuild_rejected()
        self.rej_detail.delete('1.0', 'end')

    def delete_rejected(self):
        rows = self._selected_rejected('刪除')
        if not rows:
            return
        from tkinter import messagebox
        if not messagebox.askyesno('刪除被拒字詞',
                                   f'由清單刪除 {len(rows)} 組字詞？\n\n{self._pairs_preview(rows)}\n\n'
                                   '之後有新句子再出現同樣情況時，會再次列出。', parent=self.root):
            return
        if self._db_write(self._drop_rejected(rows)):
            self._after_rejected_removed(rows)
            self.flash(f'已刪除 {len(rows)} 組字詞。')

    def accept_rejected(self):
        rows = self._selected_rejected('選用')
        if not rows:
            return
        from tkinter import messagebox
        if not messagebox.askyesno('以後用 AI 版本',
                                   f'以後 AI 提出以下改法時一律採用？\n\n{self._pairs_preview(rows)}\n\n'
                                   'AI No Change 清單的字詞及受保護人名亦會照改；《》〈〉內的作品名不會改。\n'
                                   '已上傳到 TMDB 的內容不會自動更新。可在「已選用 AI 清單」取消。', parent=self.root):
            return
        drop = self._drop_rejected(rows)
        now = time.strftime('%Y-%m-%d %H:%M:%S')

        def action(con):
            con.executemany('INSERT OR REPLACE INTO accepted (draft_word, ai_word, draft_ctx, ai_ctx, created_at) '
                            'VALUES (?, ?, ?, ?, ?)', [(r[1], r[2], r[3], r[4], now) for r in rows])
            drop(con)
        if self._db_write(action):
            self._after_rejected_removed(rows)
            self.flash(f'已選用 {len(rows)} 組 AI 字詞，運行中的程式處理下一項時生效。', 20)

    def show_accepted(self):
        if self.acc_win is not None and self.acc_win.winfo_exists():
            self.acc_win.lift()
            return
        rows = []
        if os.path.exists(self.collector.cache_path):
            self._db_write(lambda con: rows.extend(con.execute(
                'SELECT draft_word, ai_word, ai_ctx, created_at FROM accepted ORDER BY created_at DESC').fetchall()))
        win = tk.Toplevel(self.root)
        win.title('已選用 AI 清單')
        win.geometry('640x380')
        self.acc_win, self.acc_rows = win, rows
        bar = ttk.Frame(win, padding=6)
        bar.pack(side='bottom', fill='x')
        ttk.Label(bar, text='取消後，下次處理到相關句子時按原有規則。', foreground='#6b7280').pack(side='left')
        ttk.Button(bar, text='關閉', command=win.destroy).pack(side='right')
        ttk.Button(bar, text='取消選用', command=self.cancel_accepted).pack(side='right', padx=6)
        cols = [('draft', 'OpenCC', 120), ('ai', 'AI', 120), ('ctx', 'AI 上下文', 240), ('at', '選用時間', 140)]
        tree = ttk.Treeview(win, columns=[c[0] for c in cols], show='headings')
        for key, title, width in cols:
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor='w', stretch=key == 'ctx')
        tree.pack(fill='both', expand=True, padx=6, pady=(6, 0))
        for i, r in enumerate(rows):
            tree.insert('', 'end', iid=str(i), values=r)
        self.acc_tree = tree

    def cancel_accepted(self):
        sel = self.acc_tree.selection()
        if not sel:
            self.flash('請先揀選要取消的字詞。')
            return
        pairs = [self.acc_rows[int(i)][:2] for i in sel]
        if self._db_write(lambda con: con.executemany(
                'DELETE FROM accepted WHERE draft_word = ? AND ai_word = ?', pairs)):
            self.acc_tree.delete(*sel)
            self.flash(f'已取消選用 {len(pairs)} 組字詞。')

    def _on_rejected(self, rows):
        self.rej_rows = rows
        self._rebuild_rejected()

    def _rebuild_rejected(self):
        q = self.rej_search.get().strip()
        tree = self.rej_tree
        tree.delete(*tree.get_children())
        shown = 0
        for i, r in enumerate(self.rej_rows):
            count, d, a, dc, ac, model, last = r[:7]
            if q and not any(q in (x or '') for x in (d, a, dc, ac, r[8], r[9], r[10])):
                continue
            tree.insert('', 'end', iid=str(i), values=(count, d, a, dc, ac, model, last))
            shown += 1
        total = len(self.rej_rows)
        self.rej_info.configure(text=f'共 {total} 組字詞' + (f'（符合 {shown} 組）' if q else ''))
        self.nb.tab(self.rej_page, text=f'AI 被拒用詞 ({total})' if total else 'AI 被拒用詞')

    def _on_rejected_select(self, _):
        sel = self.rej_tree.selection()
        if not sel:
            return
        count, d, a, dc, ac, model, last, first, src, final, ai = self.rej_rows[int(sel[0])]
        t = self.rej_detail
        t.delete('1.0', 'end')
        for label, text, word in (('實際採用：', final, d), ('AI 建議：', ai, a), ('簡體原文：', src, None)):
            if not text:
                continue
            t.insert('end', label, 'h')
            start = t.index('end-1c')
            t.insert('end', text + '\n')
            if word:
                pos = t.search(word, start, 'end')
                if pos:
                    t.tag_add('mark', pos, f'{pos}+{len(word)}c')
        t.insert('end', f'{model} · {count} 句 · 首次 {first} · 最近 {last}', 'h')

    def export_rejected(self):
        if not self.rej_rows:
            self.flash('暫時沒有被拒字詞。')
            return
        path = os.path.join(self.base, 'data', 'ai_s2t_rejected.xlsx')
        try:
            add_user_site()
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'AI Rejected'
            ws.append(['句數', '實際採用 (OpenCC)', 'AI 建議', '採用上下文', 'AI 上下文', 'Model',
                       '首次', '最近', '簡體原文', '實際採用句子', 'AI 建議句子'])
            for count, d, a, dc, ac, model, last, first, src, final, ai in self.rej_rows:
                ws.append([count, d, a, dc, ac, model, first, last, src, final, ai])
            for col, w in zip('ABCDEFGHIJK', (6, 16, 16, 28, 28, 12, 19, 19, 50, 50, 50)):
                ws.column_dimensions[col].width = w
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            wb.save(path)
        except PermissionError:
            self.flash(f'{path} 正被 Excel 開啟，請先關閉再匯出。', 20)
            return
        except Exception as e:
            self.flash(f'匯出失敗: {e}', 20)
            return
        self.flash(f'已匯出 {len(self.rej_rows)} 組字詞到 {path}')
        self.open_path(path)

    # ---------------- 主控台分頁（start.py 的畫面輸出 + 輸入）
    def _build_console_tab(self):
        page = ttk.Frame(self.nb)
        self.console_page = page
        self.nb.add(page, text='主控台')
        bar = ttk.Frame(page, padding=(6, 6, 6, 0))
        bar.pack(fill='x')
        self.hide_log_lines = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text='隱藏 log 行（已在其他分頁顯示）', variable=self.hide_log_lines).pack(side='left')
        ttk.Button(bar, text='清除畫面', command=lambda: self.console_text.delete('1.0', 'end')).pack(side='right')
        ttk.Button(bar, text='開啟輸出檔', command=lambda: self.open_path(
            os.path.join(self.base, CONSOLE_LOG_REL))).pack(side='right', padx=6)
        ttk.Label(bar, text=f'只顯示由本面板啟動的 start.py 輸出（{CONSOLE_LOG_REL}）',
                  foreground='#6b7280', font=self.small_font).pack(side='left', padx=12)

        body = ttk.Frame(page, padding=6)
        body.pack(fill='both', expand=True)
        self.console_text = tk.Text(body, wrap='char', font=self.log_font, undo=False,
                                    background='#0f172a', foreground='#e2e8f0', insertbackground='#e2e8f0')
        sb = ttk.Scrollbar(body, command=self.console_text.yview)
        self.console_text.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.console_text.pack(side='left', fill='both', expand=True)
        self.console_text.tag_configure('input', foreground='#fbbf24')
        self.console_text.tag_configure('note', foreground='#93c5fd')
        self.console_text.tag_configure('warn', foreground='#fca5a5')
        self.console_text.bind('<Key>', lambda e: None if (e.state & 0x4 and e.keysym.lower() in ('c', 'a')) else 'break')

        inp = ttk.Frame(page, padding=(6, 0, 6, 6))
        inp.pack(fill='x')
        self.prompt_label = ttk.Label(inp, text='', foreground='#b45309')
        self.prompt_label.pack(fill='x', anchor='w')
        row = ttk.Frame(inp)
        row.pack(fill='x', pady=(2, 0))
        ttk.Label(row, text='輸入：').pack(side='left')
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(row, textvariable=self.input_var, font=self.log_font)
        self.input_entry.pack(side='left', fill='x', expand=True)
        self.input_entry.bind('<Return>', lambda _e: self.send_input())
        self.send_btn = ttk.Button(row, text='送出 (Enter)', command=self.send_input)
        self.send_btn.pack(side='left', padx=(6, 0))
        ttk.Button(row, text='只按 Enter', command=lambda: self.send_input('')).pack(side='left', padx=(6, 0))

        self.child = None              # 由本面板啟動的 start.py (subprocess.Popen)
        self.child_reported = True
        self.console_tail = ''         # 最後一段未換行的文字（通常是 input() 的提問）
        self.console_tail_since = 0
        self.tail_shown = 0            # console_tail 已顯示在主控台畫面的字數
        self.auto_answered = set()
        self.waiting_prompt = None
        self._update_input_state()

    def _console_append(self, text, tag=None):
        t = self.console_text
        at_end = t.yview()[1] >= 0.999
        t.insert('end', text, (tag,) if tag else ())
        n = int(t.index('end-1c').split('.')[0])
        if n > MAX_CONSOLE_LINES:
            t.delete('1.0', f'{n - MAX_CONSOLE_LINES}.0')
        if at_end:
            t.see('end')

    def _on_console(self, restarted, text):
        if restarted:
            self.console_text.delete('1.0', 'end')
            self.console_tail = ''
            self.tail_shown = 0
        if not text:
            return
        data = self.console_tail + text
        lines = data.split('\n')
        self.console_tail = lines.pop()               # 未完成的一行留待下次（或作提問）
        already = self.tail_shown                     # 提問已顯示的字數（避免重複顯示）
        if lines:
            self.tail_shown = 0
        if self.console_tail:
            self.console_tail_since = time.monotonic()
        shown = []
        for i, ln in enumerate(lines):
            if i == 0 and already:
                rest = ln[already:]                   # 前半（提問）已顯示，只補上其餘部分
                if self.hide_log_lines.get() and LINE_RE.match(rest):
                    rest = ''
                shown.append(rest)
                continue
            if self.hide_log_lines.get() and LINE_RE.match(ln):
                continue
            shown.append(ln)
            if ENTER_HINT_RE.search(ln) and self._child_alive():
                self._set_waiting(ln.strip())
        if shown:
            self._console_append('\n'.join(shown) + '\n')
        if self.console_tail:
            self._handle_prompt(self.console_tail)

    def _handle_prompt(self, prompt):
        """有未換行文字 → 可能是 input() 提問：已知提問自動回答，其餘提示用家輸入。"""
        if not self._child_alive():
            return
        for key_text, opt in AUTO_ANSWERS:
            if key_text in prompt and opt not in self.auto_answered:
                self.auto_answered.add(opt)
                self.console_tail = ''
                self._console_append(prompt)
                answer = self.opt_vars[opt].get()
                self.send_input('' if answer == SAME_AS_SERIES else answer, auto=True)
                return

    def _set_waiting(self, prompt):
        self.waiting_prompt = prompt
        self.prompt_label.configure(text=f'⏳ 程式正在等待輸入：{prompt[-160:]}')
        if self.nb.select() != str(self.console_page):
            self.nb.select(self.console_page)
        self.input_entry.configure(show='*' if SECRET_PROMPT_RE.search(prompt) else '')
        self.input_entry.focus_set()
        try:
            self.root.bell()
        except tk.TclError:
            pass

    def _check_prompt_idle(self):
        """未換行文字停留超過 1.5 秒 → 視為等待輸入的提問。"""
        if (self.console_tail and self._child_alive() and self.waiting_prompt != self.console_tail
                and time.monotonic() - self.console_tail_since > 1.5):
            # 提問直接顯示在主控台畫面（之前只顯示在輸入框上方，容易看漏）
            if len(self.console_tail) > self.tail_shown:
                self._console_append(self.console_tail[self.tail_shown:])
                self.console_text.see('end')
                self.tail_shown = len(self.console_tail)
            self._set_waiting(self.console_tail)

    def _child_alive(self):
        return self.child is not None and self.child.poll() is None

    def send_input(self, text=None, auto=False):
        if text is None:
            text = self.input_var.get()
        if not self._child_alive():
            self.flash('只可以輸入到由本面板啟動、仍在運行的 start.py。')
            return
        try:
            self.child.stdin.write((text + '\n').encode('utf-8'))
            self.child.stdin.flush()
        except OSError as e:
            self.flash(f'送出失敗: {e}')
            return
        secret = SECRET_PROMPT_RE.search(self.waiting_prompt or self.console_tail or '')
        if self.console_tail:                        # 提問與答案同一行顯示
            self._console_append(self.console_tail[self.tail_shown:])
            self.console_tail = ''
        self.tail_shown = 0
        shown = ('*' * len(text) if secret and text else text) or '(Enter)'
        self._console_append(shown + ('   ← 面板按設定自動回答' if auto else '') + '\n', 'input')
        if not auto:
            self.input_var.set('')
        self.waiting_prompt = None
        self.prompt_label.configure(text='')
        self.input_entry.configure(show='')

    def _update_input_state(self):
        st = 'normal' if self._child_alive() else 'disabled'
        self.input_entry.configure(state=st)
        self.send_btn.configure(state=st)

    def _check_child(self):
        if self.child is None or self.child_reported or self.child.poll() is None:
            return
        self.child_reported = True
        rc = self.child.returncode
        try:
            self.child.stdin.close()
        except OSError:
            pass
        self._drain_console_now()
        msg = f'──── start.py 已結束（exit code {rc}）────'
        self._console_append(msg + '\n', 'note' if rc == 0 else 'warn')
        self.flash(msg, 30)
        self.waiting_prompt = None
        self.prompt_label.configure(text='')
        self._stop_sent = 0

    def _drain_console_now(self):
        if self.console_tail:
            self._console_append(self.console_tail[self.tail_shown:] + '\n')
            self.console_tail = ''
        self.tail_shown = 0

    def _on_tab_changed(self, _):
        try:
            w = self.root.nametowidget(self.nb.select())
        except (KeyError, tk.TclError):
            return
        if isinstance(w, LogView):
            w.after_idle(w.on_shown)

    def _on_worker_double(self, _):
        sel = self.worker_tree.selection()
        if not sel:
            return
        key = 'importer' if sel[0] == 'uploader' else sel[0]
        self.nb.select(self.views[key])

    # ---------------- Excel 分頁
    def _build_excel_tab(self):
        page = ttk.Frame(self.nb)
        self.nb.add(page, text='Excel 項目')
        bar = ttk.Frame(page)
        bar.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(bar, text='顯示').pack(side='left')
        self.excel_filter = tk.StringVar(value='全部')
        cb = ttk.Combobox(bar, textvariable=self.excel_filter, state='readonly', width=12,
                          values=['全部', '未完成', '進行中', '待處理', '失敗', '完成'])
        cb.pack(side='left', padx=(4, 10))
        cb.bind('<<ComboboxSelected>>', lambda _: self._rebuild_excel())
        ttk.Label(bar, text='搜尋 URL').pack(side='left')
        self.excel_search = tk.StringVar()
        self.excel_search.trace_add('write', lambda *_: self._schedule_excel_rebuild())
        ttk.Entry(bar, textvariable=self.excel_search, width=30).pack(side='left', padx=(4, 10))
        ttk.Button(bar, text='跳到進行中', command=self._jump_to_active).pack(side='left')
        ttk.Button(bar, text='立即重新讀取', command=self.collector_force_excel).pack(side='left', padx=6)
        self.excel_label = ttk.Label(bar, text='', foreground='#6b7280')
        self.excel_label.pack(side='right')

        body = ttk.Frame(page)
        body.pack(fill='both', expand=True, padx=6, pady=(0, 6))
        cols = ('row', 'status', 'done', 'video', 'tmdb', 'split')
        self.excel_tree = ttk.Treeview(body, columns=cols, show='headings')
        for c, label, w, st in (('row', '列', 60, False), ('status', '狀態', 110, False), ('done', 'Done', 70, False),
                                ('video', 'Video URL', 520, True), ('tmdb', 'TMDB URL', 380, True),
                                ('split', 'Split', 70, False)):
            self.excel_tree.heading(c, text=label)
            self.excel_tree.column(c, width=w, stretch=st, anchor='center' if not st else 'w')
        for status, color in STATUS_COLORS.items():
            self.excel_tree.tag_configure(status, background=color)
        ys = ttk.Scrollbar(body, orient='vertical', command=self.excel_tree.yview)
        self.excel_tree.configure(yscrollcommand=ys.set)
        self.excel_tree.pack(side='left', fill='both', expand=True)
        ys.pack(side='right', fill='y')
        legend = ttk.Frame(page)
        legend.pack(fill='x', padx=6, pady=(0, 6))
        for status in ('提取中', '等待上傳', '上傳中', '待處理', '失敗', '完成 (未寫入)'):
            tk.Label(legend, text=f' {status} ', background=STATUS_COLORS[status], relief='solid',
                     borderwidth=1).pack(side='left', padx=(0, 6))
        ttk.Label(legend, text='「完成 (未寫入)」= log 已記錄 Done，但仍未寫入 Excel（例如 Excel 正被開啟）。列號 = Excel 實際列號',
                  foreground='#6b7280').pack(side='left', padx=8)
        self._excel_job = None

    def collector_force_excel(self):
        self.collector.force_excel.set()

    def _schedule_excel_rebuild(self):
        if self._excel_job:
            self.root.after_cancel(self._excel_job)
        self._excel_job = self.root.after(300, self._rebuild_excel)

    def _col(self, name):
        try:
            return self.excel_header.index(name)
        except ValueError:
            return None

    def _row_status(self, idx, done_raw, video='', tmdb=''):
        live = self.state.row_live.get(idx)
        if live:
            return live
        pend = self.state.done_pending.get(idx)
        ds = str(done_raw).strip().lower()
        excel_done = not (ds in ('', 'nan', '0', '0.0', 'none'))
        if pend is not None and pend.get('urls'):
            # log 記錄的是該列「當時」的連結；Excel 已換成別的連結 → 舊結果不適用
            lv, lt = pend['urls']
            if (lv and lv != norm_url(video)) or (not lv and lt and lt != norm_url(tmdb)):
                pend = None
        if pend is not None and not pend['written']:
            if pend['val'] == 0:
                return '失敗'
            return '完成' if excel_done else '完成 (未寫入)'
        if ds in ('0', '0.0'):
            return '失敗'
        return '完成' if excel_done else '待處理'

    def _row_values(self, idx, r):
        def g(name):
            i = self._col(name)
            return r[i] if i is not None and i < len(r) else ''
        done = g('Done')
        status = self._row_status(idx, done, g('Video URL'), g('TMDB URL'))
        return status, (idx + 2, status, done, g('Video URL'), g('TMDB URL'), g('Split'))

    def _filter_ok(self, status, values):
        f = self.excel_filter.get()
        if f == '未完成' and status in ('完成', '完成 (未寫入)'):
            return False
        if f == '進行中' and status not in ('提取中', '等待上傳', '上傳中'):
            return False
        if f == '待處理' and status != '待處理':
            return False
        if f == '失敗' and status != '失敗':
            return False
        if f == '完成' and status not in ('完成', '完成 (未寫入)'):
            return False
        q = self.excel_search.get().strip().lower()
        return not q or q in f'{values[3]} {values[4]}'.lower()

    def _rebuild_excel(self):
        self._excel_job = None
        tree = self.excel_tree
        y = tree.yview()[0]
        tree.delete(*tree.get_children())
        for idx, r in enumerate(self.excel_rows):
            status, values = self._row_values(idx, r)
            if self._filter_ok(status, values):
                tree.insert('', 'end', iid=str(idx), values=values, tags=(status,))
        tree.yview_moveto(y)
        self.state.changed_rows.clear()

    def _refresh_changed_rows(self):
        changed, self.state.changed_rows = self.state.changed_rows, set()
        if not changed:
            return
        if self.excel_filter.get() != '全部':
            self._schedule_excel_rebuild()
            return
        for idx in changed:
            if 0 <= idx < len(self.excel_rows) and self.excel_tree.exists(str(idx)):
                status, values = self._row_values(idx, self.excel_rows[idx])
                self.excel_tree.item(str(idx), values=values, tags=(status,))

    def _jump_to_active(self):
        active = sorted(self.state.row_live)
        for idx in active:
            if self.excel_tree.exists(str(idx)):
                self.excel_tree.see(str(idx))
                self.excel_tree.selection_set(str(idx))
                return
        self.flash('目前沒有進行中的項目', 5)

    # ---------------- 資料更新
    def _drain(self):
        lines_by_batch = []
        excel_changed = False
        try:
            for _ in range(500):
                kind, payload = self.q.get_nowait()
                if kind == 'lines':
                    for e in payload:
                        self.state.feed(e)
                    lines_by_batch.extend(payload)
                elif kind == 'reset':
                    self.state.reset_source(payload)
                    self._sync_excel_sheet()
                    self.state.changed_rows.update(range(len(self.excel_rows)))
                    self._events_shown = 0
                    self.event_tree.delete(*self.event_tree.get_children())
                    marker = LogEntry(payload, time.strftime('%Y-%m-%d %H:%M:%S'), 'WARNING',
                                      '──── log 已被清除（程式重新啟動？）────',
                                      f'──── {payload} log 已被清除（程式重新啟動？）────')
                    lines_by_batch.append(marker)
                elif kind == 'excel':
                    self.excel_header, self.excel_rows, t, sheet = payload
                    self.excel_info = f'Excel（工作表 {sheet}）讀取於 {t}'
                    excel_changed = True
                elif kind == 'excel_error':
                    self.excel_info = payload
                elif kind == 'console':
                    self._on_console(*payload)
                elif kind == 'rejected':
                    self._on_rejected(payload)
                elif kind == 'proc':
                    self.proc_info = payload
                elif kind == 'cache':
                    self.cache_count = payload
                elif kind == 'error':
                    self.last_error = payload
        except queue.Empty:
            pass
        if lines_by_batch:
            for v in self.views.values():
                v.add(lines_by_batch)
            n_prob = sum(1 for e in lines_by_batch if LEVEL_ORDER.get(e.level, 20) >= 30 and e.ts)
            if n_prob:
                self.problem_count += n_prob
                self.nb.tab(self.views['problems'], text=f'⚠ 警告 / 錯誤 ({self.problem_count})')
        self._sync_excel_sheet()
        if excel_changed:
            self._rebuild_excel()
        else:
            self._refresh_changed_rows()
        self.root.after(300 if not self.q.empty() else 500, self._drain)

    def _tick(self):
        try:
            self._check_child()
            self._check_prompt_idle()
            self._update_input_state()
            self._update_summary()
            self._update_workers()
            self._update_events()
        finally:
            self.root.after(1000, self._tick)

    def _elapsed(self, ts):
        t0 = parse_ts(ts) if ts else None
        if not t0:
            return ''
        s = max(0, int((dt.datetime.now() - t0).total_seconds()))
        return f'{s // 3600}:{s // 60 % 60:02d}:{s % 60:02d}' if s >= 3600 else f'{s // 60}:{s % 60:02d}'

    def _update_summary(self):
        st = self.state
        procs = self.proc_info.get('procs')
        val, sub = self.cards['proc']
        if procs:
            p = procs[0]
            up = int(time.time() - p['create_time'])
            if os.path.exists(self.pause_path):
                val.configure(text='已暫停', foreground='#b45309')
            elif os.path.exists(self.upload_pause_path):
                val.configure(text='上傳已暫停', foreground='#b45309')
            else:
                val.configure(text='運行中', foreground='#047857')
            sub.configure(text=f"PID {p['pid']} · 已運行 {up // 3600}:{up // 60 % 60:02d}"
                               + (' · 由面板啟動' if self._child_alive() else ' · 外部啟動（不能在面板輸入）')
                               + (f' · 共 {len(procs)} 個!' if len(procs) > 1 else ''))
        elif procs is None:
            val.configure(text='檢查中…', foreground='#6b7280')
            sub.configure(text=self.proc_info.get('error', ''))
        else:
            val.configure(text='未運行', foreground='#6b7280')
            sub.configure(text=self.proc_info.get('error') or f'最後 log: {st.last_ts or "—"}')
        self._update_start_button(procs)

        val, sub = self.cards['phase']
        val.configure(text=st.phase)
        sub.configure(text=f'最後 log: {(st.last_ts or "—")[11:]}')

        total = len(self.excel_rows)
        counts = collections.Counter(self._row_values(i, r)[0] for i, r in enumerate(self.excel_rows))
        done = counts['完成'] + counts['完成 (未寫入)']
        val, sub = self.cards['excel']
        val.configure(text=f'{done} / {total}' if total else '—')
        _sheet = self.collector.bulk_sheet + (f'（{self.state.sheet_progress}）' if self.state.sheet_progress else '')
        # 【2026-11-14】進行中分開顯示，睇到有幾多已提取、排緊隊等上傳
        sub.configure(text=f'{_sheet} · 待處理 {counts["待處理"]} · 失敗 {counts["失敗"]}\n'
                           f'提取中 {counts["提取中"]} · 等待上傳 {counts["等待上傳"]} · 上傳中 {counts["上傳中"]}')
        self.progress.configure(maximum=max(total, 1), value=done)

        c = st.counts
        val, sub = self.cards['session']
        val.configure(text=f"提取 {c['extract_started']} · 上傳 {c['uploaded']}")
        sub.configure(text=f"失敗 {c['upload_failed'] + c['extract_failed']} · 只計面板開啟後及最近約 1MB log")

        val, sub = self.cards['cache']
        val.configure(text='—' if self.cache_count is None else f'{self.cache_count} 句')
        sub.configure(text='data/ai_s2t_cache.db')

        self.excel_label.configure(text=self.excel_info)
        if time.monotonic() < self._flash_until:
            self.status_label.configure(text=self._flash_text, foreground='#1d4ed8')
        else:
            self.status_label.configure(text=self.last_error or f'監控資料夾: {self.base}', foreground='#6b7280')

    def _update_workers(self):
        for key in ['uploader'] + [f'e{i}' for i in range(NUM_EXTRACTORS)]:
            w = self.state.workers.get(key) or {'state': '未開始', 'row': None, 'url': '', 'since': None}
            busy = w['state'] in ('提取中', '上傳中')
            row = '' if w['row'] is None else w['row'] + 2
            tag = ('upload' if w['state'] == '上傳中' else 'busy') if busy else (
                'ended' if w['state'] in ('已結束', '未開始') else '')
            label = 'Uploader' if key == 'uploader' else f'Extractor {key[1:]}'
            self.worker_tree.item(key, values=(label, w['state'], row,
                                               self._elapsed(w['since']) if busy else '', w['url'] if busy else ''),
                                  tags=(tag,))

    def _update_events(self):
        events = list(self.state.events)
        if len(events) == self._events_shown:
            return
        new = events[self._events_shown:] if self._events_shown < len(events) else events
        if self._events_shown > len(events) or len(events) == MAX_EVENTS:
            self.event_tree.delete(*self.event_tree.get_children())
            new = events
        for t, src, text in new:
            tag = 'bad' if '失敗' in text or 'Done=0' in text else ('good' if '成功' in text or 'Done=1' in text else '')
            src_label = 'UP' if src == 'importer' and ('上傳' in text) else SRC_LABEL.get(src, src)
            self.event_tree.insert('', 0, values=(t, src_label, text), tags=(tag,))
        self._events_shown = len(events)

    # ---------------- 其他
    def flash(self, text, seconds=10):
        """在右上狀態列顯示訊息數秒。"""
        self._flash_text, self._flash_until = text, time.monotonic() + seconds
        self.status_label.configure(text=text, foreground='#1d4ed8')

    def _running_pids(self):
        pids = [p['pid'] for p in (self.proc_info.get('procs') or [])]
        if self._child_alive() and self.child.pid not in pids:
            pids.insert(0, self.child.pid)
        return pids

    def _update_start_button(self, procs):
        running = bool(procs) or self._child_alive()
        self.pause_btn.configure(state='normal' if running else 'disabled',
                                 text='▶ 繼續' if running and os.path.exists(self.pause_path) else '⏸ 暫停')
        self.upload_pause_btn.configure(
            state='normal' if running else 'disabled',
            text='▶ 繼續上傳' if running and os.path.exists(self.upload_pause_path) else '⏸ 暫停上傳')
        if running:
            self.start_btn.configure(state='disabled', text='● 運行中')
            self.stop_btn.configure(state='normal',
                                    text='■ 強制停止' if self._stop_sent else '■ 停止 (Ctrl+C)')
        else:
            self.stop_btn.configure(state='disabled', text='■ 停止 (Ctrl+C)')
            self._stop_sent = 0
            if procs is None:                       # 仍在檢查程式狀態
                self.start_btn.configure(state='disabled', text='▶ 啟動')
            elif time.monotonic() < self._start_cooldown_until:
                self.start_btn.configure(state='disabled', text='啟動中…')
            else:
                self.start_btn.configure(state='normal', text='▶ 啟動')

    @staticmethod
    def _python_exe():
        import shutil
        return shutil.which('python') or shutil.which('py') or sys.executable.replace('pythonw', 'python')

    def start_program(self, in_cmd=False):
        """預設：在隱藏的主控台執行 start.py，畫面輸出寫入 data/monitor_console.log 並顯示於「主控台」分頁，
        三條啟動提問按上方設定自動回答，其餘提問可在面板輸入。
        in_cmd=True：舊方式，開新的命令視窗（在該視窗手動回答）。"""
        if self.proc_info.get('procs') or self._child_alive():
            self.flash('程式已在運行中，不會重複啟動。')
            return
        script = os.path.join(self.base, 'start.py')
        if not os.path.exists(script):
            self.flash(f'找不到 {script}')
            return
        import subprocess
        python = self._python_exe()
        opts = {k: v.get() for k, v in self.opt_vars.items()}
        save_settings(opts)
        try:
            if in_cmd:
                if os.name == 'nt':
                    comspec = os.environ.get('COMSPEC') or os.path.join(
                        os.environ.get('SystemRoot', r'C:\Windows'), 'System32', 'cmd.exe')
                    subprocess.Popen([comspec, '/k', python, 'start.py'], cwd=self.base,
                                     creationflags=subprocess.CREATE_NEW_CONSOLE)
                else:
                    subprocess.Popen([python, 'start.py'], cwd=self.base)
            else:
                out_path = os.path.join(self.base, CONSOLE_LOG_REL)
                os.makedirs(os.path.dirname(out_path), exist_ok=True)
                # TMDB_MONITOR_SETTINGS：single 模式每次輸入 TMDB URL 後讀取面板最新的覆蓋選項
                env = dict(os.environ, PYTHONUNBUFFERED='1', PYTHONIOENCODING='utf-8',
                           TMDB_MONITOR_SETTINGS=SETTINGS_PATH)
                kw = {}
                if os.name == 'nt':
                    # 有自己的（隱藏）主控台：Ctrl+C / 第二次 Ctrl+C 強制結束 與手動執行時完全相同
                    si = subprocess.STARTUPINFO()
                    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    si.wShowWindow = 0                      # SW_HIDE
                    kw = {'creationflags': subprocess.CREATE_NEW_CONSOLE, 'startupinfo': si}
                with open(out_path, 'wb') as out:
                    out.write((f'==== {time.strftime("%Y-%m-%d %H:%M:%S")} 由監控面板啟動 start.py '
                               f'(覆蓋劇集標題={opts["overwrite_titles"]}, 覆蓋劇集描述={opts["overwrite_descriptions"]}, '
                               f'覆蓋分集標題={opts["episode_overwrite_titles"]}, 覆蓋分集描述={opts["episode_overwrite_descriptions"]}, '
                               f'來源無資料時清空TMDB={opts["clear_when_source_empty"]}, '
                               f'維基人名取代現有={opts["wiki_name_replace"]}, '
                               f'模式={opts["mode"]}, bulk 工作表={opts["bulk_sheet"]}, '
                               f'保留上次未上傳={opts.get("keep_pending_uploads", "n")}) ====\n').encode('utf-8'))
                    out.flush()
                    self.child = subprocess.Popen([python, '-u', 'start.py'], cwd=self.base, env=env,
                                                  stdin=subprocess.PIPE, stdout=out, stderr=subprocess.STDOUT, **kw)
                self.child_reported = False
                self.auto_answered = set()
                self.console_tail = ''
                self.tail_shown = 0
                self.waiting_prompt = None
                self.collector.console_reset.set()
                self.nb.select(self.console_page)
        except Exception as e:
            self.flash(f'啟動失敗: {e}', 30)
            return
        self._start_cooldown_until = time.monotonic() + 15   # 等待偵測到新行程，期間防止重複按
        self.start_btn.configure(state='disabled', text='啟動中…')
        self.flash('已在新命令視窗啟動 start.py，請在該視窗回答啟動問題。' if in_cmd
                   else '已啟動 start.py（輸出見「主控台」分頁）。', 20)
        self.collector.force_excel.set()   # 盡快更新程式狀態

    def toggle_pause(self):
        try:
            if os.path.exists(self.pause_path):
                os.remove(self.pause_path)
                self.flash('已繼續。')
            else:
                os.makedirs(os.path.dirname(self.pause_path), exist_ok=True)
                open(self.pause_path, 'w').close()
                self.flash('已暫停：進行中的項目會完成，之後不會開始新項目。按「繼續」恢復。', 20)
        except OSError as e:
            self.flash(f'暫停／繼續失敗: {e}', 20)
        self._update_start_button(self.proc_info.get('procs'))

    def toggle_upload_pause(self):
        try:
            if os.path.exists(self.upload_pause_path):
                os.remove(self.upload_pause_path)
                self.flash('已繼續上傳：uploader 會由上傳佇列繼續。')
            else:
                os.makedirs(os.path.dirname(self.upload_pause_path), exist_ok=True)
                open(self.upload_pause_path, 'w').close()
                self.flash('已暫停上傳：uploader 會喺下一個頁面步驟停低（包括限流等待中），繼續後由停低位置接住做；'
                           'extractor 繼續提取（只限 bulk 模式）。'
                           '按「繼續上傳」恢復。', 20)
        except OSError as e:
            self.flash(f'暫停／繼續上傳失敗: {e}', 20)
        self._update_start_button(self.proc_info.get('procs'))

    def stop_program(self, confirm=True):
        """向 start.py 所在主控台送出 Ctrl+C（與在命令視窗按 Ctrl+C 相同）。
        第一次：程式自行安全結束（寫入 Done、關閉自己開的瀏覽器）；第二次：程式的強制結束機制。"""
        from tkinter import messagebox
        pids = self._running_pids()
        if not pids:
            self.flash('沒有偵測到運行中的 start.py。')
            return False
        if confirm:
            if self._stop_sent:
                q = '已送出過 Ctrl+C。再送一次會「強制結束」程式及其開啟的瀏覽器（未寫入的 Done 值可能遺失）。\n\n確定強制結束？'
            else:
                q = '送出 Ctrl+C 停止 start.py？\n\n程式會完成收尾（寫入 Done 值、關閉自己開啟的瀏覽器）後結束。'
            if not messagebox.askyesno('停止程式', q, parent=self.root):
                return False
        ok = send_ctrl_c(pids)
        if ok:
            self._stop_sent += 1
            self.flash('已送出 Ctrl+C，等待程式收尾結束…' if self._stop_sent == 1 else '已再次送出 Ctrl+C（強制結束）。', 30)
            if self._child_alive():
                self._console_append(f'^C  ← 面板送出 Ctrl+C（第 {self._stop_sent} 次）\n', 'input')
        else:
            self.flash('送出 Ctrl+C 失敗（可在工作管理員結束程式）。', 30)
        return ok

    def open_path(self, path):
        try:
            if os.name == 'nt':
                os.startfile(path)
            else:
                import subprocess
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            self.flash(f'無法開啟 {path}: {e}')

    def choose_folder(self):
        """選擇另一個專案資料夾（需含 start.py），記住選擇並以新資料夾重新開啟面板。"""
        new = ask_project_folder(self.root, self.base)
        new = os.path.abspath(new) if new else None
        if new and os.path.normcase(new) != os.path.normcase(self.base):
            save_settings({'base': new})
            self.switch_to = new
            self.close()

    def close(self):
        if self._child_alive():
            from tkinter import messagebox
            ans = messagebox.askyesnocancel(
                '關閉監控面板',
                'start.py 仍在運行（由本面板啟動）。\n\n'
                '「是」：先送出 Ctrl+C 停止程式，再關閉面板\n'
                '「否」：程式繼續在背景運行，只關閉面板\n'
                '　　（重開面板後仍可看進度及按「停止」，但不能再回答程式的提問）\n'
                '「取消」：不關閉',
                parent=self.root)
            if ans is None:
                self.switch_to = None
                return
            if ans:
                self.stop_program(confirm=False)
        self.collector.stop.set()
        self.root.destroy()


# ---------------------------------------------------------------------------
# 送出 Ctrl+C 到 start.py 所在的主控台
# ---------------------------------------------------------------------------
# 面板本身（pythonw）沒有主控台，不能直接 GenerateConsoleCtrlEvent；
# 改為開一個無視窗的小 helper：脫離自己的主控台 → 附加到目標主控台 → 忽略自己的 Ctrl+C → 送出。
# 效果等同用家在該命令視窗按 Ctrl+C（同一主控台內的行程都會收到，與手動操作相同）。
_CTRL_C_HELPER = r'''
import ctypes, sys
k = ctypes.windll.kernel32
k.FreeConsole()
for pid in map(int, sys.argv[1:]):
    if k.AttachConsole(pid):
        k.SetConsoleCtrlHandler(None, True)
        ok = k.GenerateConsoleCtrlEvent(0, 0)
        k.FreeConsole()
        sys.exit(0 if ok else 3)
sys.exit(2)
'''


def send_ctrl_c(pids):
    if os.name != 'nt':
        import signal
        try:
            os.kill(pids[0], signal.SIGINT)
            return True
        except OSError:
            return False
    import subprocess
    try:
        r = subprocess.run([sys.executable, '-c', _CTRL_C_HELPER, *map(str, pids)],
                           creationflags=subprocess.DETACHED_PROCESS, timeout=15,
                           stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return r.returncode == 0
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 專案資料夾設定（記住上次選擇）
# ---------------------------------------------------------------------------
SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_settings.json')


def is_project_folder(path):
    return bool(path) and os.path.isfile(os.path.join(path, 'start.py'))


def list_bulk_sheets(path):
    """讀取 bulk_videos.xlsx 嘅工作表名稱（唯讀、共享開啟；失敗回傳空清單）。"""
    try:
        add_user_site()
        import openpyxl
        with open_shared(path) as f:
            data = f.read()
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
        try:
            return list(wb.sheetnames)
        finally:
            wb.close()
    except Exception:
        return []


def load_settings():
    try:
        import json
        with open(SETTINGS_PATH, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        import json
        cur = load_settings()
        cur.update(data)
        with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
            json.dump(cur, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def ask_project_folder(parent, initial):
    """彈出資料夾選擇視窗，直至選到含 start.py 的資料夾或取消。"""
    from tkinter import filedialog, messagebox
    while True:
        path = filedialog.askdirectory(parent=parent, initialdir=initial or os.path.expanduser('~'),
                                       title='選擇 TMDB 程式資料夾（含 start.py）', mustexist=True)
        if not path:
            return None
        path = os.path.abspath(path)
        if is_project_folder(path):
            return path
        messagebox.showwarning('TMDB 監控面板', f'{path}\n\n這個資料夾沒有 start.py，請選擇程式所在的資料夾。',
                               parent=parent)
        initial = path


def main():
    ap = argparse.ArgumentParser(description='TMDB-parallel 監控面板（唯讀）')
    ap.add_argument('--base', default=None, help='專案資料夾（預設：上次選擇的資料夾，否則本檔所在資料夾）')
    args = ap.parse_args()
    if os.name == 'nt':
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [args.base, load_settings().get('base'), here]
    base = next((os.path.abspath(c) for c in candidates if c and is_project_folder(c)), None)
    if args.base and not is_project_folder(args.base):
        base = None
    while True:
        root = tk.Tk()
        if base is None:      # 找不到程式資料夾 → 請用家選擇
            root.withdraw()
            base = ask_project_folder(root, args.base or here)
            if base is None:
                root.destroy()
                return
            save_settings({'base': base})
            root.deiconify()
        app = MonitorApp(root, base)
        root.mainloop()
        if not getattr(app, 'switch_to', None):
            break
        base = app.switch_to


if __name__ == '__main__':
    main()
