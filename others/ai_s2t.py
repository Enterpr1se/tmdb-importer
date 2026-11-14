"""
AI-assisted Simplified -> Hong Kong Traditional conversion (簡轉繁 AI 校對)

只處理「包含多義簡體字」的句子（例如：后 → 後/后、干 → 乾/幹/干、发 → 發/髮）。
OpenCC 先轉換出草稿，再把這些句子連同草稿交給 AI 校對；其餘文字完全不經 AI。

安全機制（ai_config.json "validation"）：
  * relaxed（預設，2026-11 R1）：接受 AI 改字/換詞（荷里活、雪糕…），但《》〈〉內、@@PROTECTED@@、
    受保護人名、tw_hk_name.xlsx「AI No Change」頁及 Default / Movie name 譯名的字詞不准改；單段改動 > relaxed_max_chunk 字還原；
    整句改動 > relaxed_max_change_ratio → 整句保留 OpenCC。快取存 AI 原始回覆，每次重新驗證。
  * strict：AI 回覆逐字驗證：每個字只能是「原簡體字的合法繁體候選字」，否則該字還原為 OpenCC 結果。
  * 結果永久快取（data/ai_s2t_cache.db，SQLite），同一句子不會重複呼叫 AI。
  * AI 失敗 / 逾時 → 自動改用 fallback_providers，全部失敗則保留 OpenCC 結果。
  * 連續失敗達 circuit_breaker_failures 次 → 本次執行停用 AI，避免拖慢程式。

設定檔：專案根目錄 ai_config.json（不存在時自動建立預設檔）。
觸發字/詞：data/tw_hk_name.xlsx 的「AI Convert」頁（不存在時自動建立預設內容）。
"""

import json
import logging
import os
import re
import shutil
import subprocess
import tempfile
import threading

import requests

CONFIG_PATH = 'ai_config.json'
DICT_DIR = os.path.join('opencc_data', 'dictionary')

DEFAULT_CONFIG = {
    "_說明": "簡轉繁 AI 校對設定。enabled=false 則完全不使用 AI。provider 選擇下方 providers 其中一個名稱。",
    "enabled": True,
    "_provider_order_說明": "使用順序：第一個失敗/額度用完就用下一個，全部失敗則保留 OpenCC 結果。名稱對應下方 providers",
    "provider_order": ["claude_cli"],
    "_trigger_說明": "需要 AI 處理的字/詞記錄在 trigger_xlsx 的 trigger_sheet 頁（不存在時自動建立）",
    "trigger_xlsx": "data/tw_hk_name.xlsx",
    "trigger_sheet": "AI Convert",
    "_validation_說明": "relaxed（預設）= 接受 AI 的改字/換詞（如 荷里活、雪糕），但《》內、受保護人名、AI No Change 清單不准改，"
                        "改動超過 relaxed_max_change_ratio 或單段超過 relaxed_max_chunk 字則還原；strict = 舊的逐字候選字檢查",
    "validation": "relaxed",
    "relaxed_max_change_ratio": 0.2,
    "relaxed_max_chunk": 4,
    "no_change_sheet": "AI No Change",
    "max_concurrent_calls": 2,
    "timeout_seconds": 120,
    "max_items_per_call": 30,
    "circuit_breaker_failures": 5,
    "cache_file": "data/ai_s2t_cache.db",
    "providers": {
        "claude_cli": {
            "_說明": "Claude Code 訂閱 (Pro/Max)。需已安裝並登入 claude。最平: haiku",
            "type": "claude_cli", "command": "claude", "model": "haiku",
            "env": {"MAX_THINKING_TOKENS": "0"}
        },
        "codex_cli": {
            "_說明": "OpenAI Codex CLI (ChatGPT 訂閱)。model 留空 = Codex 預設；可填較平的 mini 型號",
            "type": "codex_cli", "command": "codex", "model": ""
        },
        "gemini_cli": {
            "_說明": "Google Gemini CLI。最平: gemini-2.5-flash-lite",
            "type": "gemini_cli", "command": "gemini", "model": "gemini-2.5-flash-lite"
        },
        "openai": {
            "_說明": "OpenAI API。api_key 直接填寫，或留空並設定環境變數 api_key_env",
            "type": "openai_compatible", "base_url": "https://api.openai.com/v1",
            "api_key": "", "api_key_env": "OPENAI_API_KEY", "model": "gpt-5-nano"
        },
        "deepseek": {
            "_說明": "DeepSeek API（中文佳、價錢低）",
            "type": "openai_compatible", "base_url": "https://api.deepseek.com/v1",
            "api_key": "", "api_key_env": "DEEPSEEK_API_KEY", "model": "deepseek-chat",
            "_model_說明": "deepseek-chat = 非推理模式（平、快）；deepseek-reasoner 為推理模式，唔建議"
        },
        "openrouter": {
            "_說明": "OpenRouter（一個 key 用多家 model）",
            "type": "openai_compatible", "base_url": "https://openrouter.ai/api/v1",
            "api_key": "", "api_key_env": "OPENROUTER_API_KEY", "model": "qwen/qwen-2.5-72b-instruct"
        },
        "openrouter_free": {
            "_說明": "自動掃描 OpenRouter 所有免費 model 並輪流使用；某 model 額度滿/限流就跳下一個，全部用完才放棄。需 OpenRouter API key（免費註冊）",
            "type": "openrouter_free", "base_url": "https://openrouter.ai/api/v1",
            "_extra_body_說明": "附加到每個請求；reasoning.enabled=false 關閉推理（較快、較慳額度）",
            "extra_body": {"reasoning": {"enabled": False}},
            "api_key": "", "api_key_env": "OPENROUTER_API_KEY",
            "_preferred_models_說明": "想優先使用的 model id（可留空），其後按 prefer_keywords 排序",
            "preferred_models": [],
            "prefer_keywords": ["qwen", "deepseek", "glm", "kimi", "moonshot", "gemma", "nemotron", "gemini", "llama", "mistral"],
            "_exclude_keywords_說明": "model id 含以下字眼則不使用（例如太細的模型中文差）",
            "exclude_keywords": ["1.7b", "2.6b", "code"],
            "cooldown_seconds": 600,
            "refresh_hours": 12,
            "_qualification_說明": "每個 model 首次使用前先做入門測試（qualification_tests 留空 = 用內建 6 句），合格才使用；結果存檔，retest_days 日後或測試句改變時重測",
            "qualification_enabled": True,
            "_qualification_mode_說明": "no_regression = 不可把正確字改錯且要比 OpenCC 好（預設）；strict = 測試句要全對（按 qualification_min_pass_ratio）",
            "qualification_mode": "no_regression",
            "qualification_min_pass_ratio": 1.0,
            "retest_days": 7,
            "qualification_file": "data/openrouter_model_tests.json",
            "qualification_tests": [],
            "models_cache_file": "data/openrouter_free_models.json"
        },
        "ollama": {
            "_說明": "本機 Ollama（免費），需先 ollama pull 模型",
            "type": "openai_compatible", "base_url": "http://localhost:11434/v1",
            "api_key": "ollama", "model": "qwen2.5:7b"
        },
        "anthropic": {
            "_說明": "Anthropic API（按量收費）。最平: claude-haiku-5-5",
            "type": "anthropic", "api_key": "", "api_key_env": "ANTHROPIC_API_KEY",
            "model": "claude-haiku-5-5"
        },
        "gemini": {
            "_說明": "Google Gemini API（有免費額度）",
            "type": "gemini", "api_key": "", "api_key_env": "GEMINI_API_KEY",
            "model": "gemini-2.5-flash-lite"
        }
    }
}

SYSTEM_PROMPT = (
    "你是「簡體中文 → 香港繁體中文」的校對員。\n"
    "每個項目有 src（簡體原文）和 draft（機器轉換的繁體草稿，字數與 src 相同）。\n"
    "機器轉換常錯在一簡對多繁的字，例如：后→後/后（天后、皇后、影后 用「后」；之後、後來 用「後」）、"
    "干→乾/幹/干、发→發/髮、里→裏/里、面→面/麪、表→表/錶、钟→鐘/鍾、只→只/隻、台→台/臺/颱、范→范/範、系→系/係/繫。\n"
    "請根據語境修正 draft 中選錯的字，並使用香港慣用字形（如 裏、着、麪）。\n"
    "規則：\n"
    "1. 只可替換字，不可增刪、改寫、重排；每個輸出字數必須與 draft 完全相同。\n"
    "2. 人名、專有名詞依常見香港寫法（如 范冰冰、天后宮）。\n"
    "3. 保留標點、空格、英文、數字及 @@PROTECTED_n@@ 之類標記原樣不變。\n"
    "4. 只輸出一個 JSON 字串陣列，順序與輸入相同，不要任何解釋或 markdown。"
)

# 【2026-11 R1】relaxed 模式：容許少量換詞（香港慣用譯名/用詞、明顯錯字），程式會再檢查
SYSTEM_PROMPT_RELAXED = (
    "你是「簡體中文 → 香港繁體中文」的校對員。\n"
    "每個項目有 src（簡體原文）和 draft（機器轉換的繁體草稿）。\n"
    "機器轉換常錯在一簡對多繁的字，例如：后→後/后（天后、皇后、影后 用「后」；之後、後來 用「後」）、"
    "干→乾/幹/干、发→發/髮、里→裏/里、面→面/麪、表→表/錶、钟→鐘/鍾、只→只/隻、台→台/臺/颱、范→范/範、系→系/係/繫。\n"
    "請根據語境修正 draft，並使用香港慣用字形（如 裏、着、麪）。\n"
    "規則：\n"
    "1. 以替換字為主；可把字詞改為香港慣用寫法或譯名（如 好萊塢→荷里活、冰淇淋→雪糕、戛納→康城），"
    "以及修正明顯錯別字（如 的/地/得 誤用），但不可改寫句子、不可增刪內容、不可重排。\n"
    "2. 《》、〈〉內的作品名稱保持 draft 原樣。\n"
    "3. 保留標點、空格、英文、數字及 @@PROTECTED_n@@ 之類標記原樣不變。\n"
    "4. 只輸出一個 JSON 字串陣列，順序與輸入相同，不要任何解釋或 markdown。"
)


def _validation_mode():
    mode = str(_load_config().get('validation') or 'relaxed').strip().lower()
    return 'strict' if mode == 'strict' else 'relaxed'


def _system_prompt():
    return SYSTEM_PROMPT if _validation_mode() == 'strict' else SYSTEM_PROMPT_RELAXED

_SENTENCE_RE = re.compile(r'[^。！？!?；;\n]*[。！？!?；;\n]?')


class _State:
    def __init__(self):
        self.lock = threading.Lock()
        self.config = None
        self.semaphore = None
        self.cache = None        # sqlite3 connection
        self.candidates = None   # simp char -> set of acceptable traditional chars
        self.trad_to_simp = None  # trad char -> set of simp chars
        self.triggers = None
        self.no_change = None    # AI 不准改的字詞（tw_hk_name.xlsx「AI No Change」頁）
        self.failures = 0
        self.disabled_reason = None


_state = _State()


# ---------------------------------------------------------------------------
# Config / cache / dictionaries
# ---------------------------------------------------------------------------
def _load_config():
    if _state.config is not None:
        return _state.config
    with _state.lock:
        if _state.config is not None:
            return _state.config
        cfg = json.loads(json.dumps(DEFAULT_CONFIG))
        if not os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                    json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
                logging.info(f"[AI簡轉繁] 已建立預設設定檔 {CONFIG_PATH}")
            except Exception as e:
                logging.warning(f"[AI簡轉繁] 無法建立 {CONFIG_PATH}: {e}")
        else:
            try:
                with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                    user_cfg = json.load(f)
                user_providers = user_cfg.pop('providers', {}) or {}
                cfg.update(user_cfg)
                for name, p in user_providers.items():
                    merged = dict(cfg['providers'].get(name, {}))
                    merged.update(p)
                    cfg['providers'][name] = merged
            except Exception as e:
                logging.error(f"[AI簡轉繁] 讀取 {CONFIG_PATH} 失敗，停用 AI: {e}")
                cfg['enabled'] = False
        _state.semaphore = threading.Semaphore(max(1, int(cfg.get('max_concurrent_calls', 2))))
        _state.config = cfg
        if cfg.get('enabled'):
            desc = []
            for n in _provider_order(cfg):
                pc = cfg['providers'].get(n, {})
                desc.append(f"{n}({'免費 model 輪換' if pc.get('type') == 'openrouter_free' else pc.get('model') or '預設'})")
            logging.info(f"[AI簡轉繁] 已啟用，使用順序: {' → '.join(desc)} → OpenCC")
        return cfg


# 【2026-11】快取改用 SQLite：每次只寫入新增的句子，不再整個 JSON 檔重寫。
# 舊的 data/ai_s2t_cache.json 會在首次開啟時自動匯入，並改名為 .json.migrated。
def _cache_db_path():
    path = _load_config().get('cache_file') or 'data/ai_s2t_cache.db'
    return os.path.splitext(path)[0] + '.db' if path.lower().endswith('.json') else path


def _load_cache():
    """開啟（必要時建立）快取資料庫。可重複呼叫。"""
    if _state.cache is not None:
        return
    import sqlite3
    with _state.lock:
        if _state.cache is not None:
            return
        path = _cache_db_path()
        os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
        conn = sqlite3.connect(path, check_same_thread=False, timeout=30)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("CREATE TABLE IF NOT EXISTS cache (k TEXT PRIMARY KEY, v TEXT NOT NULL)")
        _rejected_init(conn)
        conn.commit()
        legacy = os.path.splitext(path)[0] + '.json'
        if os.path.exists(legacy):
            try:
                with open(legacy, 'r', encoding='utf-8') as f:
                    old = json.load(f)
                conn.executemany("INSERT OR IGNORE INTO cache (k, v) VALUES (?, ?)", list(old.items()))
                conn.commit()
                os.replace(legacy, legacy + '.migrated')
                logging.info(f"[AI簡轉繁] 已把 {len(old)} 條舊快取由 {legacy} 匯入 {path}")
            except Exception as e:
                logging.warning(f"[AI簡轉繁] 匯入舊 JSON 快取失敗（不影響運作）: {e}")
        _state.cache = conn


def _cache_get(key):
    with _state.lock:
        row = _state.cache.execute("SELECT v FROM cache WHERE k = ?", (key,)).fetchone()
    return row[0] if row else None


def _cache_put(pairs):
    if not pairs:
        return
    try:
        with _state.lock:
            _state.cache.executemany("INSERT OR REPLACE INTO cache (k, v) VALUES (?, ?)", pairs)
            _state.cache.commit()
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 快取寫入失敗: {e}")


# ---------------------------------------------------------------------------
# AI 被拒字詞清單（同一個 data/ai_s2t_cache.db 的 rejected 表）
# AI 建議了、但因「只可在同一簡體字的繁體候選字之間選擇」而未被採用的字詞，
# 留給用家檢查是否應加入 tw_hk_name.xlsx。監控面板「AI 被拒用詞」分頁讀取此表。
# ---------------------------------------------------------------------------
def _rejected_init(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS rejected (
        draft_word TEXT NOT NULL,      -- 實際採用（OpenCC）的字詞
        ai_word TEXT NOT NULL,         -- AI 建議的字詞
        draft_ctx TEXT, ai_ctx TEXT,   -- 前後數字的上下文（最近一次）
        src TEXT, final_text TEXT, ai_text TEXT, model TEXT,
        count INTEGER NOT NULL DEFAULT 0,   -- 出現於多少句不同的句子
        first_seen TEXT, last_seen TEXT,
        PRIMARY KEY (draft_word, ai_word))""")
    conn.execute("CREATE TABLE IF NOT EXISTS rejected_seen (k TEXT PRIMARY KEY)")
    # 用家在監控面板選了「以後用 AI 版本」的字詞；monitor_gui 亦會建立同一個表（改欄位要同步）
    conn.execute("""CREATE TABLE IF NOT EXISTS accepted (
        draft_word TEXT NOT NULL, ai_word TEXT NOT NULL,
        draft_ctx TEXT, ai_ctx TEXT, created_at TEXT,
        PRIMARY KEY (draft_word, ai_word))""")


def load_accepted():
    """回傳用家選用的 {(OpenCC 字詞, AI 字詞)}。每次校對重讀，運行中在面板選用亦即時生效。需已 _load_cache()。"""
    try:
        with _state.lock:
            return set(_state.cache.execute("SELECT draft_word, ai_word FROM accepted").fetchall())
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 讀取已選用字詞失敗（本句不套用）: {e}")
        return set()


def _rejected_spans(final_s, ai_s, locked_words=(), ctx=3):
    """比較實際採用文字與 AI 回覆，回傳 [(採用字詞, AI 字詞, 採用上下文, AI 上下文)]。
    碰到 locked_words（tw_hk_name 譯名）的改動不列出：這些字詞其後會由譯名表修正。"""
    import difflib
    sm = difflib.SequenceMatcher(None, final_s, ai_s, autojunk=False)
    if sm.ratio() < 0.6:        # 差太遠（例如 AI 回了解釋文字）→ 不記錄
        return []
    locked = _word_spans(final_s, locked_words)
    spans = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal' or _touches(locked, i1, i2):
            continue
        d, a = final_s[i1:i2], ai_s[j1:j2]
        if not d.strip() and not a.strip():
            continue
        spans.append((d, a, final_s[max(0, i1 - ctx):i2 + ctx].strip(), ai_s[max(0, j1 - ctx):j2 + ctx].strip()))
    return spans


def record_rejected(src_s, final_s, ai_s, model, seen_at=None):
    """記錄 AI 被拒的字詞。同一句（採用文字 + AI 回覆）只計一次，可重複呼叫。回傳是否新記錄。"""
    import hashlib
    import time as _time
    ai_s = _answer_text(ai_s)
    if not isinstance(ai_s, str) or not final_s or final_s == ai_s:
        return False
    # 【2026-11-14】AI No Change 字詞同 tw_hk_name 譯名一樣唔記錄：用家已決定唔改（例如 台），列出只係噪音
    spans = _rejected_spans(final_s, ai_s, list(load_name_map_words() or ()) + list(load_no_change_words() or ()))
    if not spans:
        return False
    k = hashlib.sha1((final_s + '\x00' + ai_s).encode('utf-8')).hexdigest()
    now = seen_at or _time.strftime('%Y-%m-%d %H:%M:%S')
    try:
        _load_config()   # 必須先於 _load_cache：_load_cache 持鎖時會讀設定，設定未載入會自鎖
        _load_cache()
        accepted = load_accepted()
        spans = [s for s in spans if (s[0], s[1]) not in accepted]   # 用家已選用的字詞不再列出
        if not spans:
            return False
        with _state.lock:
            c = _state.cache
            if c.execute("SELECT 1 FROM rejected_seen WHERE k = ?", (k,)).fetchone():
                return False
            c.execute("INSERT INTO rejected_seen (k) VALUES (?)", (k,))
            for d, a, dc, ac in spans:
                c.execute(
                    """INSERT INTO rejected (draft_word, ai_word, draft_ctx, ai_ctx, src, final_text, ai_text,
                                             model, count, first_seen, last_seen)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                       ON CONFLICT(draft_word, ai_word) DO UPDATE SET
                           count = count + 1,
                           first_seen = MIN(first_seen, excluded.first_seen),
                           last_seen = MAX(last_seen, excluded.last_seen),
                           draft_ctx = excluded.draft_ctx, ai_ctx = excluded.ai_ctx,
                           src = COALESCE(NULLIF(excluded.src, ''), src),
                           final_text = excluded.final_text, ai_text = excluded.ai_text,
                           model = excluded.model""",
                    (d, a, dc, ac, src_s or '', final_s, ai_s, model or '', now, now))
            c.commit()
        return True
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 記錄被拒字詞失敗（不影響運作）: {e}")
        return False


_LOG_TS_RE = re.compile(r'^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d),\d+ - ', re.M)
_LOG_REJ_RE = re.compile(r"\[AI簡轉繁\] (.+?) 回覆中有 \d+ 個不合法字已還原：'(.*)' -> '(.*)'\s*$", re.S)
_LOG_LEN_RE = re.compile(r"\[AI簡轉繁\] (.+?) 回覆字數不符，保留 OpenCC：'(.*)' vs '(.*)'\s*$", re.S)


def import_rejected_from_logs(log_dir='log'):
    """由舊 log 補錄被拒字詞（主程式啟動清除 log 前呼叫；重複匯入不會重複計數）。回傳新記錄的句數。"""
    if not os.path.isdir(log_dir):
        return 0
    n = 0
    for name in sorted(os.listdir(log_dir)):
        if not name.endswith('.log'):
            continue
        records, cur = [], None
        try:
            with open(os.path.join(log_dir, name), 'r', encoding='utf-8', errors='replace') as f:
                for line in f:           # 逐行掃描（log 可達數百 MB），只收集 AI 被拒的紀錄（含續行）
                    if _LOG_TS_RE.match(line):
                        if cur:
                            records.append(''.join(cur))
                        cur = [line] if ('不合法字已還原' in line or '回覆字數不符' in line) else None
                    elif cur is not None:
                        cur.append(line)
                if cur:
                    records.append(''.join(cur))
        except OSError:
            continue
        for rec in records:
            ts = rec[:19]
            m = _LOG_REJ_RE.search(rec)
            if m:                      # 'AI 回覆' -> '實際採用'
                n += record_rejected('', m.group(3), m.group(2), m.group(1), seen_at=ts)
                continue
            m = _LOG_LEN_RE.search(rec)
            if m:                      # '實際採用 (OpenCC)' vs 'AI 回覆'
                n += record_rejected('', m.group(2), m.group(3), m.group(1), seen_at=ts)
    return n


def _read_pairs(filename):
    path = os.path.join(DICT_DIR, filename)
    if not os.path.exists(path):
        return
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            if not line.strip() or line.startswith('#'):
                continue
            parts = line.rstrip('\n').split('\t')
            if len(parts) >= 2:
                yield parts[0], parts[1].split()


def _load_candidates():
    """建立「簡體字 → 可接受繁體字集合」，用於驗證 AI 輸出。"""
    if _state.candidates is not None:
        return
    simp_to_trad, trad_to_simp = {}, {}
    for s, trads in _read_pairs('STCharacters.txt'):
        if len(s) == 1:
            for t in trads:
                if len(t) == 1:
                    simp_to_trad.setdefault(s, set()).add(t)
    # 【2026-11】唔再讀繁轉簡字典 TSCharacters.txt（用家要求停用，曾經令 AI「惡」→「噁」都被接受）
    # 香港異體字（裡→裏、麵→麪 等）
    hk, hk_norm = {}, {}
    for t, variants in _read_pairs('HKVariants.txt'):
        if len(t) == 1:
            singles = [v for v in variants if len(v) == 1]
            hk.setdefault(t, set()).update(singles)
            if singles:
                hk_norm.setdefault(t, singles[0])
    _state.hk_norm = hk_norm
    for s, trads in simp_to_trad.items():
        for t in list(trads):
            trads.update(hk.get(t, ()))
    for s, trads in simp_to_trad.items():
        for t in trads:
            trad_to_simp.setdefault(t, set()).add(s)
    _state.candidates = simp_to_trad
    _state.trad_to_simp = trad_to_simp


def _char_allowed(src_c, draft_c, ai_c):
    if ai_c == draft_c or ai_c == src_c:
        return True
    if ai_c in _state.candidates.get(src_c, ()):
        return True
    # src 字本身已是繁體（混雜文字）時：允許同一簡體來源的其他繁體字
    src_roots = _state.trad_to_simp.get(draft_c, set()) | {src_c}
    return bool(src_roots & _state.trad_to_simp.get(ai_c, set()))


def _answer_text(ai):
    """部分 model 回傳 {"draft": "..."} 之類物件而非字串，取出其中的文字。
    【2026-11-13】deepseek 試過將輸入物件變成 JSON 字串原樣回覆（'{"src":"…","draft":"…"}'），
    以前會當成 AI 建議而記錄成假嘅被拒用詞；而家先解開再取 draft 等欄位。"""
    if isinstance(ai, str):
        s = ai.strip()
        if s.startswith('{') and s.endswith('}'):
            try:
                obj = json.loads(s)
            except ValueError:
                obj = None
            if isinstance(obj, dict):
                return _answer_text(obj)
        return ai
    if isinstance(ai, dict):
        for k in ('result', 'output', 'corrected', 'answer', 'text', 'hk', 'draft'):
            if isinstance(ai.get(k), str):
                return ai[k]
        strs = [v for v in ai.values() if isinstance(v, str)]
        return strs[-1] if strs else ai
    if isinstance(ai, list) and len(ai) == 1 and isinstance(ai[0], str):
        return ai[0]
    return ai


def _merge_validated(src, draft, ai):
    """逐字驗證 AI 輸出；不合法的字還原為 draft。回傳 (結果, 被拒字數)。"""
    ai = _answer_text(ai)
    if not isinstance(ai, str) or len(ai) != len(draft) or len(src) != len(draft):
        return draft, -1
    out, rejected = [], 0
    for s, d, a in zip(src, draft, ai):
        if _char_allowed(s, d, a):
            out.append(a)
        else:
            out.append(d)
            rejected += 1
    return ''.join(out), rejected


# ---------------------------------------------------------------------------
# 【2026-11 R1】relaxed 驗證：接受 AI 換詞，但有防線
# ---------------------------------------------------------------------------
_TITLE_BRACKET_RE = re.compile(r'《[^》]*》|〈[^〉]*〉')
_PLACEHOLDER_RE = re.compile(r'@@PROTECTED_\d+@@')


def _protected_spans(draft, no_change_words):
    """回傳 draft 中不准 AI 改動的 [(start, end)]：《》〈〉作品名、@@PROTECTED_n@@、受保護人名、AI No Change 清單字詞。"""
    spans = [m.span() for m in _TITLE_BRACKET_RE.finditer(draft)]
    spans += [m.span() for m in _PLACEHOLDER_RE.finditer(draft)]
    words = list(no_change_words or ())
    try:
        from others import utils as _u
        words += [n for n in (getattr(_u, 'protected_person_names', None) or ()) if n and len(n) >= 2]
    except Exception:
        pass
    return spans + _word_spans(draft, words)


def _word_spans(text, words):
    """回傳 words 中每個字詞在 text 出現的所有 [(start, end)]。"""
    spans = []
    for w in set(words or ()):
        start = text.find(w)
        while start != -1:
            spans.append((start, start + len(w)))
            start = text.find(w, start + 1)
    return spans


def _touches(spans, i1, i2):
    """改動範圍 [i1, i2)（插入時 i1 == i2）是否碰到受保護範圍。"""
    for a, b in spans:
        if i1 == i2:
            if a < i1 < b:
                return True
        elif i1 < b and a < i2:
            return True
    return False


def _merge_relaxed(src, draft, ai, no_change_words=(), accepted=(), locked_words=()):
    """逐段比較 draft 與 AI 回覆：一般改動接受；碰到受保護範圍或單段太長的改動還原為 draft。
    改動總字數超過上限 → 整句保留 draft。回傳 (結果, 被拒段數)；-1 = 整句不採用。
    accepted = 用家選用的 {(draft 字詞, AI 字詞)}：優先於其他規則，但不可改《》〈〉作品名及 @@PROTECTED_n@@ 標記，
    整句不採用時亦照樣套用。
    locked_words = tw_hk_name 譯名：碰到的改動一律還原（包括 accepted 及合法候選字），不計被拒段數，
    否則 AI 改動後譯名表便對不上原詞。"""
    import difflib
    ai = _answer_text(ai)
    if not isinstance(ai, str) or not ai.strip():
        return draft, -1
    cfg = _load_config()
    max_chunk = max(1, int(cfg.get('relaxed_max_chunk', 4)))
    ratio = float(cfg.get('relaxed_max_change_ratio', 0.2))
    spans = _protected_spans(draft, no_change_words)
    always_kept = [m.span() for r in (_TITLE_BRACKET_RE, _PLACEHOLDER_RE) for m in r.finditer(draft)]
    # 【2026-11】AI No Change 字詞及受保護人名：連合法候選字（例如 台→臺）都唔准改；《》作品名內則照舊准改候選字
    word_kept = [s for s in spans if s not in always_kept]
    locked = _word_spans(draft, locked_words)
    sm = difflib.SequenceMatcher(None, draft, ai, autojunk=False)
    out, whole_rejected, rejected, changed = [], [], 0, 0   # whole_rejected = 整句不採用時的結果
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        d, a = draft[i1:i2], ai[j1:j2]
        if tag == 'equal' or _touches(locked, i1, i2):
            out.append(d)
            whole_rejected.append(d)
            continue
        if (d, a) in accepted and not _touches(always_kept, i1, i2):
            out.append(a)
            whole_rejected.append(a)
            continue
        whole_rejected.append(d)
        # 等長替換且每字都是合法候選字（舊 strict 規則）→ 一定接受（例如《》內 后→後）
        if (tag == 'replace' and len(d) == len(a) and len(src) == len(draft)
                and not _touches(word_kept, i1, i2)
                and all(_char_allowed(s, dc, ac) for s, dc, ac in zip(src[i1:i2], d, a))):
            out.append(a)
            continue
        if max(len(d), len(a)) > max_chunk or _touches(spans, i1, i2):
            out.append(d)
            rejected += 1
            continue
        out.append(a)
        changed += max(len(d), len(a))
    if changed > max(2, int(len(draft) * ratio)):
        return ''.join(whole_rejected), -1
    return ''.join(out), rejected


def load_no_change_words(force=False):
    """從 tw_hk_name.xlsx 的「AI No Change」頁載入 AI 不准改的字詞（OpenCC 結果中的寫法）。"""
    if getattr(_state, 'no_change', None) is not None and not force:
        return _state.no_change
    cfg = _load_config()
    path = cfg.get('trigger_xlsx') or 'data/tw_hk_name.xlsx'
    sheet = cfg.get('no_change_sheet') or 'AI No Change'
    words = None
    try:
        if _ensure_sheet(path, sheet, NO_CHANGE_HEADERS, DEFAULT_NO_CHANGE,
                         "# Word 欄填 OpenCC 轉換後的寫法；句子中出現時 AI 不可改動這些字詞；Enabled=N 停用該行"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            try:
                rows = list(wb[sheet].iter_rows(values_only=True))
            finally:
                wb.close()
            header = [str(h).strip().lower() if h is not None else '' for h in (rows[0] if rows else [])]
            wi = header.index('word') if 'word' in header else 0
            ei = header.index('enabled') if 'enabled' in header else None
            words = []
            for r in rows[1:]:
                if not r or wi >= len(r) or r[wi] is None:
                    continue
                w = str(r[wi]).strip()
                if not w or w.startswith('#'):
                    continue
                if ei is not None and ei < len(r) and r[ei] is not None and \
                        str(r[ei]).strip().upper() in ('N', 'NO', 'FALSE', '0'):
                    continue
                words.append(w)
            logging.info(f"[AI簡轉繁] 從 '{path}' [{sheet}] 載入 {len(words)} 個不准改字詞。")
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 讀取不准改清單失敗 ({e})。")
    if words is None:
        prev = getattr(_state, 'no_change', None)
        if prev is not None:
            return prev
        words = [w for w, _ in DEFAULT_NO_CHANGE]
    _state.no_change = list(dict.fromkeys(words))
    return _state.no_change


NAME_MAP_SHEETS = ('Default', 'Movie name')


def load_name_map_words():
    """回傳 tw_hk_name.xlsx 譯名表的 TW Name（≥2 字），校對時鎖定。內容來自 excel_handler 的快取。
    關鍵字專用頁（Gundam 等）只適用於個別作品，不鎖定，以免其他作品的同一字詞也不能校對。"""
    path = _load_config().get('trigger_xlsx') or 'data/tw_hk_name.xlsx'
    if not os.path.exists(path):
        return set()
    try:
        from others.excel_handler import _load_and_cache_sheet
        words = set()
        for sheet in NAME_MAP_SHEETS:
            words.update(k for k in _load_and_cache_sheet(path, sheet) if len(k) >= 2 and not k.startswith('#'))
        return words
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 讀取譯名表失敗（本句不鎖定譯名）: {e}")
        return set()


NO_CHANGE_HEADERS = ["Word", "Note", "Enabled"]
DEFAULT_NO_CHANGE = [
    ("著稱", "香港慣用「著稱」，不改為「着稱」"),
    ("憑藉", "保留「憑藉」，不改為「憑着」"),
]


# ---------------------------------------------------------------------------
# Trigger words sheet (data/tw_hk_name.xlsx → "AI Convert")
# ---------------------------------------------------------------------------
TRIGGER_HEADERS = ["Trigger", "Traditional Options", "Example", "Enabled"]
DEFAULT_TRIGGERS = [
    ("后", "後 / 后", "之後、後來 / 天后、皇后、影后"),
    ("干", "乾 / 幹 / 干", "乾杯、乾淨 / 幹練、能幹 / 干擾"),
    ("发", "發 / 髮", "發表 / 頭髮"),
    ("里", "裏 / 里", "這裏 / 公里、鄰里"),
    ("面", "面 / 麪", "面對 / 麪條"),
    ("台", "台 / 臺 / 颱 / 檯", "舞台 / 颱風 / 枱"),
    ("系", "系 / 係 / 繫", "系列 / 關係 / 聯繫"),
    ("只", "只 / 隻", "只有 / 一隻"),
    ("制", "制 / 製", "制度 / 製作"),
    ("复", "復 / 複 / 覆", "恢復 / 複雜 / 反覆"),
    ("历", "歷 / 曆", "歷史 / 日曆"),
    ("冲", "沖 / 衝", "沖涼 / 衝擊"),
    ("准", "准 / 準", "批准 / 準備"),
    ("斗", "斗 / 鬥", "北斗 / 戰鬥"),
    ("范", "范 / 範", "范冰冰 / 範圍"),
    ("松", "松 / 鬆", "松樹 / 放鬆"),
    ("征", "征 / 徵", "出征 / 特徵"),
    ("钟", "鐘 / 鍾", "鐘錶 / 鍾情"),
    ("余", "余 / 餘", "余文樂 / 多餘"),
    ("云", "云 / 雲", "人云亦云 / 白雲"),
    ("丑", "丑 / 醜", "小丑 / 醜陋"),
    ("表", "表 / 錶", "表演 / 手錶"),
    ("采", "采 / 採", "風采 / 採訪"),
    ("划", "划 / 劃", "划船 / 計劃"),
    ("尽", "盡 / 儘", "盡力 / 儘管"),
    ("谷", "谷 / 穀", "山谷 / 穀物"),
    ("卷", "卷 / 捲", "試卷 / 捲髮"),
    ("蒙", "蒙 / 濛 / 矇", "蒙古 / 迷濛 / 矇騙"),
    ("折", "折 / 摺", "折扣 / 摺疊"),
    ("借", "借 / 藉", "借錢 / 藉口"),
    ("涂", "塗 / 涂", "塗鴉 / 姓涂"),
    ("郁", "郁 / 鬱", "濃郁 / 憂鬱"),
    ("致", "致 / 緻", "一致 / 精緻"),
    ("胡", "胡 / 鬍", "胡說 / 鬍子"),
    ("须", "須 / 鬚", "必須 / 鬍鬚"),
    ("咸", "咸 / 鹹", "老少咸宜 / 鹹魚"),
    ("舍", "舍 / 捨", "宿舍 / 捨棄"),
    ("志", "志 / 誌", "志願 / 雜誌"),
    ("获", "獲 / 穫", "獲得 / 收穫"),
    ("签", "簽 / 籤", "簽名 / 抽籤"),
    ("汇", "匯 / 彙", "匯款 / 詞彙"),
    ("托", "托 / 託", "托盤 / 委託"),
    ("伙", "伙 / 夥", "伙食 / 夥伴"),
    ("朴", "朴 / 樸", "朴槿惠 / 樸素"),
    ("吁", "吁 / 籲", "長吁短嘆 / 呼籲"),
]


def _ensure_sheet(path, sheet, headers, default_rows, note, widths=(12, 22, 40, 10)):
    """若 xlsx 中沒有指定頁，建立並填入預設內容（修改前先備份）。default_rows 每行不含 Enabled 欄。"""
    import openpyxl
    if not os.path.exists(path):
        return False
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        if sheet in wb.sheetnames:
            return True
    finally:
        wb.close()
    wb = openpyxl.load_workbook(path)
    backup = path + '.before_' + re.sub(r'\W+', '_', sheet).strip('_') + '.bak'
    if not os.path.exists(backup):
        shutil.copy2(path, backup)
    ws = wb.create_sheet(sheet)
    ws.append(headers)
    for row in default_rows:
        ws.append(list(row) + ["Y"])
    ws.append([note] + [""] * (len(headers) - 1))
    for col, width in zip('ABCD', widths):
        ws.column_dimensions[col].width = width
    wb.save(path)
    logging.info(f"[AI簡轉繁] 已在 '{path}' 建立 '{sheet}' 頁 ({len(default_rows)} 個預設項)，備份: {backup}")
    return True


def _ensure_trigger_sheet(path, sheet):
    """若 xlsx 中沒有觸發字頁，建立並填入預設內容（修改前先備份）。"""
    return _ensure_sheet(path, sheet, TRIGGER_HEADERS, DEFAULT_TRIGGERS,
                         "# 簡體原文句子包含 Trigger 欄任何字/詞 (可填多字詞) 時交 AI 校對；Enabled=N 停用該行")


def load_triggers(force=False):
    """從 xlsx 載入觸發字/詞。失敗時使用內建預設清單。"""
    if _state.triggers is not None and not force:
        return _state.triggers
    cfg = _load_config()
    path = cfg.get('trigger_xlsx') or 'data/tw_hk_name.xlsx'
    sheet = cfg.get('trigger_sheet') or 'AI Convert'
    triggers = None
    try:
        if _ensure_trigger_sheet(path, sheet):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            try:
                rows = list(wb[sheet].iter_rows(values_only=True))
            finally:
                wb.close()
            header = [str(h).strip().lower() if h is not None else '' for h in (rows[0] if rows else [])]
            ti = header.index('trigger') if 'trigger' in header else 0
            ei = header.index('enabled') if 'enabled' in header else None
            triggers = []
            for r in rows[1:]:
                if not r or ti >= len(r) or r[ti] is None:
                    continue
                t = str(r[ti]).strip()
                if not t or t.startswith('#'):
                    continue
                if ei is not None and ei < len(r) and r[ei] is not None and                         str(r[ei]).strip().upper() in ('N', 'NO', 'FALSE', '0'):
                    continue
                triggers.append(t)
            logging.info(f"[AI簡轉繁] 從 '{path}' [{sheet}] 載入 {len(triggers)} 個觸發字/詞。")
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 讀取觸發字頁失敗 ({e})。")
    if triggers is None:
        # 【2026-11 F3】讀取失敗（例如 Excel 正在儲存）時沿用上一次成功讀取的清單；從未成功過才用內建預設
        if _state.triggers is not None:
            logging.warning(f"[AI簡轉繁] 沿用上一次載入的 {len(_state.triggers)} 個觸發字/詞。")
            return _state.triggers
        logging.warning("[AI簡轉繁] 改用內建預設觸發字清單。")
        triggers = [t for t, _, _ in DEFAULT_TRIGGERS]
    # 長詞優先無所謂（只做包含判斷），去重即可
    _state.triggers = list(dict.fromkeys(triggers))
    return _state.triggers


def _has_trigger(text, triggers):
    return any(t in text for t in triggers)


# ---------------------------------------------------------------------------
# Providers
# ---------------------------------------------------------------------------
def _api_key(p):
    key = p.get('api_key') or ''
    if not key and p.get('api_key_env'):
        key = os.environ.get(p['api_key_env'], '')
    return key


def _run_cli(args, prompt, timeout, extra_env=None):
    exe = shutil.which(args[0])
    if not exe:
        raise RuntimeError(f"找不到指令 '{args[0]}'，請確認已安裝並在 PATH 中")
    # 在暫存目錄執行，避免 CLI 讀取本專案檔案作為上下文
    with tempfile.TemporaryDirectory(prefix='ai_s2t_') as cwd:
        env = dict(os.environ, **{k: str(v) for k, v in (extra_env or {}).items()})
        proc = subprocess.run([exe] + args[1:], input=prompt.encode('utf-8'),
                              capture_output=True, timeout=timeout, cwd=cwd, env=env)
    out = proc.stdout.decode('utf-8', errors='replace')
    if proc.returncode != 0:
        err = proc.stderr.decode('utf-8', errors='replace').strip()[-500:]
        raise RuntimeError(f"{args[0]} 回傳碼 {proc.returncode}: {err or out[-500:]}")
    return out


def _call_provider(p, user_prompt, timeout):
    ptype = p.get('type')
    model = p.get('model') or ''
    full_prompt = _system_prompt() + "\n\n" + user_prompt

    if ptype == 'claude_cli':
        args = [p.get('command', 'claude'), '-p', '--output-format', 'text']
        if model:
            args += ['--model', model]
        return _run_cli(args, full_prompt, timeout, p.get('env'))

    if ptype == 'codex_cli':
        fd, out_file = tempfile.mkstemp(suffix='.txt', prefix='codex_out_')
        os.close(fd)
        try:
            args = [p.get('command', 'codex'), 'exec', '--skip-git-repo-check',
                    '--sandbox', 'read-only', '--output-last-message', out_file]
            if model:
                args += ['-m', model]
            args.append('-')  # 從 stdin 讀取 prompt
            stdout = _run_cli(args, full_prompt, timeout, p.get('env'))
            with open(out_file, 'r', encoding='utf-8') as f:
                return f.read() or stdout
        finally:
            try:
                os.remove(out_file)
            except OSError:
                pass

    if ptype == 'gemini_cli':
        args = [p.get('command', 'gemini')]
        if model:
            args += ['-m', model]
        return _run_cli(args, full_prompt, timeout, p.get('env'))

    if ptype == 'openai_compatible':
        headers = {'Content-Type': 'application/json'}
        key = _api_key(p)
        if key:
            headers['Authorization'] = f'Bearer {key}'
        body = {'model': model, 'messages': [
            {'role': 'system', 'content': _system_prompt()},
            {'role': 'user', 'content': user_prompt}]}
        if p.get('temperature') is not None:
            body['temperature'] = p['temperature']
        body.update(p.get('extra_body') or {})
        r = requests.post(p['base_url'].rstrip('/') + '/chat/completions',
                          headers=headers, json=body, timeout=timeout)
        r.raise_for_status()
        return r.json()['choices'][0]['message']['content']

    if ptype == 'anthropic':
        r = requests.post(p.get('base_url', 'https://api.anthropic.com').rstrip('/') + '/v1/messages',
                          headers={'x-api-key': _api_key(p), 'anthropic-version': '2023-06-01',
                                   'content-type': 'application/json'},
                          json={'model': model, 'max_tokens': 4096, 'system': _system_prompt(),
                                'messages': [{'role': 'user', 'content': user_prompt}]},
                          timeout=timeout)
        r.raise_for_status()
        return ''.join(b.get('text', '') for b in r.json().get('content', []))

    if ptype == 'gemini':
        url = (p.get('base_url', 'https://generativelanguage.googleapis.com/v1beta').rstrip('/')
               + f'/models/{model}:generateContent')
        r = requests.post(url, headers={'x-goog-api-key': _api_key(p), 'Content-Type': 'application/json'},
                          json={'systemInstruction': {'parts': [{'text': _system_prompt()}]},
                                'contents': [{'role': 'user', 'parts': [{'text': user_prompt}]}]},
                          timeout=timeout)
        r.raise_for_status()
        parts = r.json()['candidates'][0]['content']['parts']
        return ''.join(x.get('text', '') for x in parts)

    raise ValueError(f"未知的 provider type: {ptype}")


# ---------------------------------------------------------------------------
# OpenRouter free models: scan + rotate (一個 model 爆額度就跳下一個)
# ---------------------------------------------------------------------------
class _ORState:
    def __init__(self):
        self.lock = threading.Lock()
        self.models = None          # 排好序的 free model id list
        self.cooldown = {}          # model id -> 可再使用的 time.time()
        self.dead = set()           # 本次執行永久跳過 (404/400 等)
        self.daily_exhausted_until = 0.0  # 帳戶每日免費額度用完


_or = _ORState()


def _is_free(m):
    pr = m.get('pricing') or {}
    try:
        return float(pr.get('prompt', 1)) == 0 and float(pr.get('completion', 1)) == 0
    except (TypeError, ValueError):
        return False


def scan_openrouter_free_models(p=None, force=False, verbose=False):
    """
    從 OpenRouter /models 掃描所有免費 model（prompt 與 completion 價錢皆為 0、只輸出文字），
    依 prefer_keywords 排序，並快取到 models_cache_file（refresh_hours 內不重新掃描）。
    """
    import time
    if p is None:
        p = _load_config()['providers'].get('openrouter_free', {})
    cache_file = p.get('models_cache_file') or 'data/openrouter_free_models.json'
    refresh = float(p.get('refresh_hours', 12)) * 3600
    data = None
    if not force and os.path.exists(cache_file) and time.time() - os.path.getmtime(cache_file) < refresh:
        try:
            with open(cache_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            data = None
    if data is None:
        base = (p.get('base_url') or 'https://openrouter.ai/api/v1').rstrip('/')
        r = requests.get(base + '/models', timeout=30)
        r.raise_for_status()
        data = [
            {'id': m['id'], 'name': m.get('name', ''), 'context_length': m.get('context_length')}
            for m in r.json().get('data', [])
            if _is_free(m) and (m.get('architecture') or {}).get('output_modalities', ['text']) == ['text']
        ]
        try:
            os.makedirs(os.path.dirname(cache_file) or '.', exist_ok=True)
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=1)
        except Exception as e:
            logging.warning(f"[AI簡轉繁] 無法寫入 {cache_file}: {e}")

    ids = [m['id'] for m in data]
    excl = [k.lower() for k in p.get('exclude_keywords', [])]
    ids = [i for i in ids if not any(k in i.lower() for k in excl)]
    prefer = [k.lower() for k in p.get('prefer_keywords', [])]
    pinned = [m for m in p.get('preferred_models', []) if m]

    def rank(mid):
        if mid in pinned:
            return (0, pinned.index(mid))
        if mid == 'openrouter/free':      # 自動路由器放最後當後備
            return (3, 0)
        low = mid.lower()
        for i, k in enumerate(prefer):
            if k in low:
                return (1, i)
        return (2, 0)

    ordered = sorted(dict.fromkeys(pinned + ids), key=rank)
    if verbose:
        print(f"OpenRouter 免費 model 共 {len(ordered)} 個（使用順序）：")
        for i, mid in enumerate(ordered, 1):
            print(f"  {i:2d}. {mid}")
    with _or.lock:
        _or.models = ordered
    logging.info(f"[AI簡轉繁] OpenRouter 免費 model {len(ordered)} 個: {', '.join(ordered[:8])}{' ...' if len(ordered) > 8 else ''}")
    return ordered


def _strip_reasoning(text):
    return re.sub(r'<think>.*?</think>', '', text or '', flags=re.S)


class _ORRateLimited(Exception):
    pass


class _ORDailyExhausted(Exception):
    pass


class _ORDead(Exception):
    pass


class _ORBadKey(Exception):
    pass


def _or_request(p, model, user_prompt, timeout, expected_len):
    """單次呼叫一個 OpenRouter model。成功回傳 list；失敗拋出分類例外。"""
    import time
    key = _api_key(p)
    if not key:
        raise _ORBadKey("openrouter_free 需要 api_key（或環境變數 OPENROUTER_API_KEY）")
    base = (p.get('base_url') or 'https://openrouter.ai/api/v1').rstrip('/')
    headers = {'Authorization': f'Bearer {key}', 'Content-Type': 'application/json',
               'X-Title': 'TMDB-parallel s2hk'}
    try:
        r = requests.post(base + '/chat/completions', headers=headers, timeout=timeout, json={
            'model': model, 'temperature': 0, **(p.get('extra_body') or {}),
            'messages': [{'role': 'system', 'content': _system_prompt()},
                         {'role': 'user', 'content': user_prompt}]})
    except requests.RequestException as e:
        raise _ORRateLimited(f"連線/逾時: {e}")

    if r.status_code == 200:
        try:
            body = r.json()
            if body.get('error'):
                raise ValueError(body['error'])
            content = _strip_reasoning(body['choices'][0]['message'].get('content'))
            return _parse_json_array(content, expected_len)
        except Exception as e:
            raise _ORRateLimited(f"回覆無效: {str(e)[:150]}")

    msg = r.text[:300]
    if r.status_code == 429 and 'per-day' in msg:
        reset = r.headers.get('X-RateLimit-Reset')
        try:
            until = float(reset) / 1000.0 if reset else time.time() + 3600
        except ValueError:
            until = time.time() + 3600
        raise _ORDailyExhausted(until)
    if r.status_code in (429, 402, 408, 500, 502, 503, 504, 529):
        raise _ORRateLimited(f"額度滿/限流 ({r.status_code})")
    if r.status_code in (401, 403) and 'model' not in msg.lower():
        raise _ORBadKey(f"OpenRouter API key 無效 ({r.status_code}): {msg}")
    raise _ORDead(f"不可用 ({r.status_code}): {msg[:120]}")


# ---- 入門測試 (qualification) ------------------------------------------------
DEFAULT_QUALIFICATION_TESTS = [
    {"src": "她是乐坛天后，几天后她回来了。", "expected": "她是樂壇天后，幾天後她回來了。"},
    {"src": "他把头发吹干后去干活。", "expected": "他把頭髮吹乾後去幹活。"},
    {"src": "这只钟表很精致。", "expected": "這隻鐘錶很精緻。"},
    {"src": "三天后，影后回到了后台。", "expected": "三天後，影后回到了後台。"},
    {"src": "她的后妈很干练。", "expected": "她的後媽很幹練。"},
    {"src": "面条店老板范先生很高兴。", "expected": "麪條店老闆范先生很高興。"},
]

_qual_lock = threading.RLock()
_qual_results = None   # model -> {passed, score, tested_at, tests_hash, failures}
_s2hk_conv = None


def _qual_file(p):
    return p.get('qualification_file') or 'data/openrouter_model_tests.json'


def _qual_tests(p):
    return p.get('qualification_tests') or DEFAULT_QUALIFICATION_TESTS


def _tests_hash(tests, p=None):
    """測試句、草稿版本或請求設定 (extra_body) 改變 → hash 改變 → 自動重測。"""
    import hashlib
    if p is None:
        p = _load_config()['providers'].get('openrouter_free', {})
    sig = json.dumps({'v': 3, 't': tests, 'x': p.get('extra_body') or {}}, ensure_ascii=False, sort_keys=True)
    return hashlib.md5(sig.encode('utf-8')).hexdigest()[:12]


def _load_qual_results(p):
    global _qual_results
    if _qual_results is None:
        data = {}
        if os.path.exists(_qual_file(p)):
            try:
                with open(_qual_file(p), 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception:
                data = {}
        _qual_results = data
    return _qual_results


def _save_qual_results(p):
    try:
        path = _qual_file(p)
        os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(_qual_results, f, ensure_ascii=False, indent=1)
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 無法寫入測試結果檔: {e}")


def _opencc_draft(text):
    """與 utils.normalize_text 相同：先 s2t 再 s2hk（此 opencc 套件不支援設定檔互相 chain）。"""
    global _s2hk_conv
    if _s2hk_conv is None:
        import opencc
        cfg_dir = os.path.abspath(os.path.join('opencc_data', 'config'))
        _s2hk_conv = (opencc.OpenCC(os.path.join(cfg_dir, 'custom_s2t')),
                      opencc.OpenCC(os.path.join(cfg_dir, 'custom_s2hk')))
    s2t, s2hk = _s2hk_conv
    return s2hk.convert(s2t.convert(text))


# 同字異形（兩種寫法都正確），測試比對時視為相同
_EQUIV_VARIANTS = ["麵麪", "裡裏", "線綫", "衛衞", "啟啓", "群羣", "峰峯", "眾衆", "鉤鈎", "污汙", "唇脣", "床牀", "痴癡"]
_EQUIV_MAP = {c: grp[0] for grp in _EQUIV_VARIANTS for c in grp}


def _hk_normalize(text):
    hk = getattr(_state, 'hk_norm', None) or {}
    return ''.join(_EQUIV_MAP.get(hk.get(c, c), hk.get(c, c)) for c in text)


def _evaluate(p, answers):
    """
    依 qualification_mode 評分已收到的答案：
      no_regression (預設)：不可把 OpenCC 已正確的字改錯，且整體要比單用 OpenCC 好
      strict              ：全部測試句完全正確（或達 qualification_min_pass_ratio）
    """
    _load_candidates()
    tests = _qual_tests(p)
    correct = base_correct = regressions = 0
    failures = []
    for t, ans in zip(tests, answers):
        src, expected = t['src'], t['expected']
        draft = _opencc_draft(src)
        merged, _ = _merge_validated(src, draft, ans)
        e, d, m = _hk_normalize(expected), _hk_normalize(draft), _hk_normalize(merged)
        base_correct += (d == e)
        if m == e:
            correct += 1
            continue
        reg = sum(1 for dc, mc, ec in zip(d, m, e) if dc == ec and mc != ec) if len(d) == len(m) == len(e) else 1
        regressions += reg
        failures.append({'src': src, 'expected': expected,
                         'got': ans if isinstance(ans, str) else repr(ans),
                         'kind': '改錯' if reg else '未修正'})
    n = len(tests)
    if p.get('qualification_mode', 'no_regression') == 'strict':
        passed = correct / n >= float(p.get('qualification_min_pass_ratio', 1.0)) - 1e-9
    else:
        passed = regressions == 0 and correct > base_correct
    return {'passed': passed, 'score': f"{correct}/{n}", 'opencc_score': f"{base_correct}/{n}",
            'regressions': regressions, 'failures': failures}


def _qual_status(p, model):
    """回傳 'passed' / 'failed' / None(未測試或已過期)。每次以目前規則重新評分已存答案。"""
    import time
    res = _load_qual_results(p).get(model)
    if not res or res.get('tests_hash') != _tests_hash(_qual_tests(p), p):
        return None
    if time.time() - res.get('tested_at', 0) > float(p.get('retest_days', 7)) * 86400:
        return None
    if res.get('unavailable'):
        return 'failed'
    answers = res.get('answers')
    if answers is None and res.get('score', '').endswith(f"/{len(_qual_tests(p))}"):
        # 舊格式：由錯誤記錄還原全部答案（答對的即 expected）
        got = {f['src']: f['got'] for f in res.get('failures', [])}
        answers = [got.get(t['src'], t['expected']) for t in _qual_tests(p)]
        res['answers'] = answers
    if answers is None:
        return None
    ev = _evaluate(p, answers)
    if ev['passed'] != res.get('passed') or ev['score'] != res.get('score'):
        res.update(ev)
        with _qual_lock:
            _save_qual_results(p)
    res.update(ev)
    return 'passed' if ev['passed'] else 'failed'


def _record_unavailable(p, model, reason):
    """地區限制 / 已下架等長期不可用的 model 記錄為不合格（retest_days 後再試）。"""
    import time
    if not any(k in reason.lower() for k in ('region', 'not found', 'no endpoints', 'deprecated',
                                              'reasoning is mandatory')):
        return
    with _qual_lock:
        _load_qual_results(p)[model] = {
            'passed': False, 'score': '0/0', 'tested_at': time.time(),
            'tested_at_text': time.strftime('%Y-%m-%d %H:%M'), 'tests_hash': _tests_hash(_qual_tests(p), p),
            'failures': [], 'unavailable': reason[:200]}
        _save_qual_results(p)


def _run_qualification(p, model, timeout):
    """用已知答案的測試句考核 model。合格 → True；答錯 → False；限流等 → 拋出例外（不作判定）。"""
    import time
    _load_candidates()
    tests = _qual_tests(p)
    items = [(t['src'], _opencc_draft(t['src']), t['expected']) for t in tests]
    prompt = "輸入：\n" + json.dumps([{'src': s, 'draft': d} for s, d, _ in items], ensure_ascii=False, indent=1)
    answers = _or_request(p, model, prompt, timeout, len(items))
    ev = _evaluate(p, answers)
    with _qual_lock:
        _load_qual_results(p)[model] = dict(ev, answers=answers, tested_at=time.time(),
                                            tested_at_text=time.strftime('%Y-%m-%d %H:%M'),
                                            tests_hash=_tests_hash(tests, p))
        _save_qual_results(p)
    if ev['passed']:
        logging.info(f"[AI簡轉繁] OpenRouter {model} 入門測試合格 ({ev['score']}，OpenCC {ev['opencc_score']})，開始使用。")
    else:
        detail = '; '.join(f"[{f['kind']}] {f['got']} (應為 {f['expected']})" for f in ev['failures'][:3])
        logging.info(f"[AI簡轉繁] OpenRouter {model} 入門測試不合格 ({ev['score']}，OpenCC {ev['opencc_score']})，跳過: {detail}")
    return ev['passed']


def _call_openrouter_free(p, user_prompt, timeout, expected_len):
    """
    逐個免費 model 嘗試：
      * 未測試的 model 先做入門測試，答錯就跳過（結果記錄在 qualification_file，retest_days 後重測）
      * 額度滿 / 限流 / 逾時 / 回覆格式錯 → 冷卻 cooldown_seconds，跳下一個
      * 帳戶每日免費額度用完 → 全部停用至重設時間
    """
    import time
    if time.time() < _or.daily_exhausted_until:
        raise RuntimeError("OpenRouter 今日免費額度已用完")
    if _or.models is None:
        scan_openrouter_free_models(p)
    cooldown_s = float(p.get('cooldown_seconds', 600))
    use_qual = p.get('qualification_enabled', True)
    tried = skipped_failed = 0
    for model in list(_or.models):
        with _or.lock:
            if model in _or.dead or _or.cooldown.get(model, 0) > time.time():
                continue
        try:
            if use_qual:
                with _qual_lock:  # 避免兩條執行緒同時測試同一 model
                    status = _qual_status(p, model)
                    if status is None:
                        tried += 1
                        status = 'passed' if _run_qualification(p, model, timeout) else 'failed'
                if status == 'failed':
                    skipped_failed += 1
                    continue
            tried += 1
            return _or_request(p, model, user_prompt, timeout, expected_len), model
        except _ORDailyExhausted as e:
            until = e.args[0]
            with _or.lock:
                _or.daily_exhausted_until = until
            raise RuntimeError(f"OpenRouter 每日免費額度已用完 (重設時間 {time.strftime('%Y-%m-%d %H:%M', time.localtime(until))})")
        except _ORBadKey as e:
            raise RuntimeError(str(e))
        except _ORRateLimited as e:
            logging.info(f"[AI簡轉繁] OpenRouter {model} {e}，跳下一個。")
            with _or.lock:
                _or.cooldown[model] = time.time() + cooldown_s
        except _ORDead as e:
            logging.info(f"[AI簡轉繁] OpenRouter {model} {e}，本次執行跳過。")
            with _or.lock:
                _or.dead.add(model)
            _record_unavailable(p, model, str(e))
    raise RuntimeError(f"OpenRouter 所有免費 model 暫時都用不到 (嘗試 {tried} 次，{skipped_failed} 個測試不合格)")


def qualify_all_models(force=False):
    """逐個測試所有免費 model 並列印報告（python -m others.ai_s2t qualify）。"""
    p = _load_config()['providers'].get('openrouter_free', {})
    models = scan_openrouter_free_models(p)
    timeout = _load_config().get('timeout_seconds', 120)
    for m in models:
        status = None if force else _qual_status(p, m)
        if status is None:
            try:
                _run_qualification(p, m, timeout)
            except _ORDailyExhausted:
                print("!! 每日免費額度已用完，停止測試")
                break
            except _ORDead as e:
                _record_unavailable(p, m, str(e))
            except Exception as e:
                print(f"  ⚠ 未能測試  {m} ({e})", flush=True)
                continue
        r = _load_qual_results(p).get(m, {})
        _qual_status(p, m)  # 以目前規則重新評分
        mark = '✅ 合格  ' if r.get('passed') else '❌ 不合格'
        extra = f"（{r['unavailable'][:60]}）" if r.get('unavailable') else f"（OpenCC {r.get('opencc_score', '?')}，改錯 {r.get('regressions', '?')} 字）"
        print(f"  {mark} {r.get('score', '?'):>5}  {m} {extra}", flush=True)
        for f in r.get('failures', [])[:3]:
            print(f"        [{f.get('kind', '')}] 得 {f['got']}  應為 {f['expected']}", flush=True)


def _parse_json_array(text, expected_len):
    text = text.strip()
    m = re.search(r'\[.*\]', text, re.S)
    if not m:
        raise ValueError(f"AI 回覆中找不到 JSON 陣列: {text[:200]}")
    arr = json.loads(m.group(0))
    if not isinstance(arr, list) or len(arr) != expected_len:
        raise ValueError(f"AI 回覆陣列長度 {len(arr) if isinstance(arr, list) else '?'} != {expected_len}")
    return arr


def _provider_order(cfg):
    """provider_order (list) 優先；否則 provider + fallback_providers。"""
    order = cfg.get('provider_order')
    if isinstance(order, str):
        order = [x.strip() for x in order.split(',')]
    if not order:
        order = [cfg.get('provider')] + list(cfg.get('fallback_providers') or [])
    return [n for n in dict.fromkeys(order) if n]


def _ask_ai(items):
    """items: list of (src, draft). 回傳 AI 原始字串 list；全部 provider 失敗則 raise。"""
    cfg = _load_config()
    timeout = cfg.get('timeout_seconds', 120)
    payload = [{'src': s, 'draft': d} for s, d in items]
    user_prompt = "輸入：\n" + json.dumps(payload, ensure_ascii=False, indent=1)
    names = _provider_order(cfg)
    last_err = None
    for name in names:
        p = cfg['providers'].get(name)
        if not p:
            last_err = RuntimeError(f"ai_config.json 中沒有 provider '{name}'")
            continue
        try:
            with _state.semaphore:
                if p.get('type') == 'openrouter_free':
                    arr, model = _call_openrouter_free(p, user_prompt, timeout, len(items))
                    return arr, f"{name}:{model}"
                raw = _call_provider(p, user_prompt, timeout)
            return _parse_json_array(raw, len(items)), name
        except Exception as e:
            last_err = e
            logging.warning(f"[AI簡轉繁] provider '{name}' 失敗: {e}")
    raise last_err or RuntimeError("沒有可用 provider")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def is_enabled():
    cfg = _load_config()
    return bool(cfg.get('enabled')) and _state.disabled_reason is None


def refine(src_text, draft_text):
    """
    src_text: 簡體原文（OpenCC 轉換前）；draft_text: OpenCC 轉換結果（字數須相同）。
    回傳經 AI 校對後的文字；任何問題都安全地回傳 draft_text。
    """
    if not src_text or not draft_text or not is_enabled():
        return draft_text
    if len(src_text) != len(draft_text):
        logging.debug("[AI簡轉繁] 轉換前後長度不同，跳過 AI 校對。")
        return draft_text

    cfg = _load_config()
    triggers = load_triggers()
    if not _has_trigger(src_text, triggers):
        return draft_text

    # 切句並找出需要 AI 的句子 (以 index 對應 src/draft)
    spans = []
    for m in _SENTENCE_RE.finditer(src_text):
        if m.end() == m.start():
            continue
        spans.append((m.start(), m.end()))
    targets = [(a, b) for a, b in spans if _has_trigger(src_text[a:b], triggers)]
    if not targets:
        return draft_text

    try:
        _load_candidates()
        _load_cache()
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 初始化失敗，跳過: {e}")
        return draft_text

    relaxed = _validation_mode() == 'relaxed'
    no_change = load_no_change_words() if relaxed else ()
    accepted = load_accepted() if relaxed else ()
    locked = load_name_map_words() if relaxed else ()

    results = {}
    pending = []
    for a, b in targets:
        src_s, draft_s = src_text[a:b], draft_text[a:b]
        if relaxed:
            # 【2026-11 R1】快取 AI 原始回覆，每次重新驗證（不准改清單修改後即時生效）
            key = 'relaxed\x00' + src_s + '\x00' + draft_s
            cached = _cache_get(key)
            if cached is not None:
                merged, _rej = _merge_relaxed(src_s, draft_s, cached, no_change, accepted, locked)
                results[(a, b)] = merged
                continue
        else:
            key = src_s + '\x00' + draft_s
            cached = _cache_get(key)
            if cached is not None:
                results[(a, b)] = cached
                continue
        pending.append(((a, b), src_s, draft_s, key))

    batch_size = max(1, int(cfg.get('max_items_per_call', 30)))
    for i in range(0, len(pending), batch_size):
        if _state.disabled_reason:
            break
        chunk = pending[i:i + batch_size]
        try:
            answers, used = _ask_ai([(s, d) for _, s, d, _ in chunk])
            with _state.lock:
                _state.failures = 0
        except Exception as e:
            with _state.lock:
                _state.failures += 1
                if _state.failures >= int(cfg.get('circuit_breaker_failures', 5)):
                    _state.disabled_reason = str(e)
                    logging.error(f"[AI簡轉繁] 連續失敗 {_state.failures} 次，本次執行停用 AI，改用 OpenCC 結果。")
            continue
        new_entries = []
        for (span, src_s, draft_s, key), ans in zip(chunk, answers):
            if relaxed:
                ans_text = _answer_text(ans)
                merged, rejected = _merge_relaxed(src_s, draft_s, ans_text, no_change, accepted, locked)
                if rejected == -1:
                    logging.warning(f"[AI簡轉繁] {used} 改動過多或回覆無效，保留 OpenCC：'{draft_s}' vs '{ans_text}'")
                    record_rejected(src_s, draft_s, ans_text, used)
                elif rejected:
                    logging.warning(f"[AI簡轉繁] {used} 回覆中有 {rejected} 段受保護/過長改動已還原：'{ans_text}' -> '{merged}'")
                    record_rejected(src_s, merged, ans_text, used)
                if merged != draft_s:
                    logging.info(f"[AI簡轉繁] {used} 修正：'{draft_s.strip()}' -> '{merged.strip()}'")
                results[span] = merged
                if isinstance(ans_text, str):
                    new_entries.append((key, ans_text))
                continue
            merged, rejected = _merge_validated(src_s, draft_s, ans)
            if rejected == -1:
                logging.warning(f"[AI簡轉繁] {used} 回覆字數不符，保留 OpenCC：'{draft_s}' vs '{ans}'")
                record_rejected(src_s, draft_s, ans, used)
                continue  # 不快取，下次可重試
            if rejected:
                logging.warning(f"[AI簡轉繁] {used} 回覆中有 {rejected} 個不合法字已還原：'{ans}' -> '{merged}'")
                record_rejected(src_s, merged, ans, used)
            if merged != draft_s:
                logging.info(f"[AI簡轉繁] {used} 修正：'{draft_s.strip()}' -> '{merged.strip()}'")
            results[span] = merged
            new_entries.append((key, merged))
        _cache_put(new_entries)

    if not results:
        return draft_text
    # 【2026-11 R1】relaxed 模式句子長度可能改變 → 按位置由頭重組
    out, pos = [], 0
    for (a, b) in sorted(results):
        text = results[(a, b)]
        if not relaxed and len(text) != b - a:
            continue
        out.append(draft_text[pos:a])
        out.append(text)
        pos = b
    out.append(draft_text[pos:])
    return ''.join(out)


if __name__ == '__main__':
    # 用法（在專案根目錄）：
    #   python -m others.ai_s2t scan            掃描並列出 OpenRouter 免費 model
    #   python -m others.ai_s2t qualify [--force] 逐個測試所有免費 model（結果存檔）
    #   python -m others.ai_s2t test 几天后她回来了   用目前設定的 provider 測試一句
    import sys
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'scan'
    if cmd == 'scan':
        scan_openrouter_free_models(force=True, verbose=True)
    elif cmd == 'qualify':
        qualify_all_models(force='--force' in sys.argv)
    elif cmd == 'test':
        text = ' '.join(sys.argv[2:]) or '她是乐坛天后，几天后她回来了，头发干了。'
        draft = _opencc_draft(text)
        print('OpenCC :', draft)
        print('AI 校對:', refine(text, draft))
