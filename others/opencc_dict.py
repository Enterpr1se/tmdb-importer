"""
OpenCC 字典管理（start.py 啟動時呼叫 setup()）。

  * 字典來自 OpenCC 官方最新正式版（GitHub release），每日最多檢查一次；有新版先下載。
    強制檢查：python start.py --update-deps，或 python -m others.opencc_dict update
  * 下載後刪走註解行同空行（opencc-python-reimplemented 讀到會出錯）。多個寫法時套件只用第一個。
  * 換新版前，舊字典備份到 opencc_data/dictionary_backup/<版本>_<時間>/。
    下載或檢查失敗就繼續用現有字典；完全冇字典又下載唔到先會停止。
  * 修正字典：opencc_overrides/*.txt（會上傳 git），格式同 OpenCC 一樣「簡體<Tab>繁體」，可加 # 註解。
    同官方同名字典合併（opencc_data/dictionary/Merged_*.txt）：同一個 key 用修正版本，新 key 直接加入。
    所以長詞語照樣優先，例如修正「一出→一出」之後「一出戏」仍然係「一齣戲」；
    單字修正亦唔會蓋過官方詞語。只影響 OpenCC 階段，唔會改到本身係繁體嘅文字。
  * 日文新字體字典（JPShinjitaiCharacters.txt）唔用官方版本，整份放喺 opencc_overrides/（見 LOCAL_DICTS）。
  * 由比較表匯入選擇：python -m others.opencc_dict import-choices [data/opencc_choices.xlsx]

只用標準函式庫：start.py 會喺其他套件 import 之前呼叫。
"""
import glob
import json
import os
import shutil
import sys
import tempfile
import time
import urllib.request

REPO = 'BYVoid/OpenCC'
BASE_DIR = 'opencc_data'
DICT_DIR = os.path.join(BASE_DIR, 'dictionary')
CONFIG_DIR = os.path.join(BASE_DIR, 'config')
BACKUP_DIR = os.path.join(BASE_DIR, 'dictionary_backup')
STATE_FILE = os.path.join(BASE_DIR, 'dictionary_version.json')
OVERRIDE_DIR = 'opencc_overrides'
# 唔可以用 JP / TW / HK 開頭：utils.load_manual_dictionaries 會將嗰啲檔案當成手動替換字典
MERGED_PREFIX = 'Merged_'

# 轉換設定實際用到嘅官方字典（其他官方字典唔下載）
# 【2026-11】唔用繁轉簡字典（TSPhrases / TSCharacters）：用家要求停用；判斷繁體改用 utils._trad_only_chars()
REQUIRED = ('STPhrases.txt', 'STCharacters.txt', 'HKVariants.txt')
OVERRIDES = ('STPhrases.txt', 'STCharacters.txt', 'HKVariants.txt')
# 唔用官方版本、整個字典由 opencc_overrides/ 提供：
# 官方 JPShinjitaiCharacters.txt 有 400 個字，好多同簡體中文一樣（例如 来→來），
# 佢喺簡轉繁之前執行，會令程式誤以為簡體句子已經係繁體而跳過轉換。
LOCAL_DICTS = ('JPShinjitaiCharacters.txt',)
LOCAL_DICT_HEADER = ("# 完整字典（唔係修正字典），取代官方同名字典。格式「原文<Tab>結果」；# 開頭係註解。\n"
                     "# 只放日文新字體獨有、唔會喺簡體中文出現嘅字。\n")
OVERRIDE_HEADER = ("# OpenCC 修正字典：格式「原文<Tab>結果」，每行一條；# 開頭係註解。\n"
                   "# 同官方同名字典合併：同一個原文用呢度嘅結果，長詞語照樣優先。\n"
                   "# 想某個詞唔轉換，就寫「原文<Tab>原文」，例如「一出<Tab>一出」。\n")
USER_AGENT = 'TMDB-parallel-opencc-dictionary-check'
TIMEOUT = 15


def _get(url):
    req = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return r.read()


def latest_tag():
    """OpenCC 最新正式版嘅 tag（例如 ver.1.4.2）。"""
    data = json.loads(_get(f'https://api.github.com/repos/{REPO}/releases/latest').decode('utf-8'))
    tag = data.get('tag_name')
    if not tag:
        raise ValueError('GitHub 回覆冇 tag_name')
    return tag


def clean_lines(text):
    """回傳 (有效行, 格式錯誤行)。刪走空行同 # 註解；有效行 = 「key<Tab>value」。"""
    good, bad = [], []
    for line in text.splitlines():
        line = line.strip('\r\n')
        if not line.strip() or line.startswith('#'):
            continue
        parts = line.split('\t')
        if len(parts) == 2 and parts[0].strip() and parts[1].strip():
            good.append(f'{parts[0].strip()}\t{parts[1].strip()}')
        else:
            bad.append(line)
    return good, bad


def _read_state():
    try:
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _write_state(state):
    tmp = STATE_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STATE_FILE)


def _have_all():
    return all(os.path.exists(os.path.join(DICT_DIR, f)) for f in REQUIRED)


def ensure_dictionaries(force=False, log=print):
    """確保官方字典存在並係最新正式版。每日最多查一次；force=True 即時查。"""
    os.makedirs(DICT_DIR, exist_ok=True)
    state = _read_state()
    have = _have_all()
    today = time.strftime('%Y-%m-%d')
    current = state.get('tag') or '本機舊版'
    if have and not force and state.get('checked') == today:
        log(f"今日 ({today}) 已檢查過 OpenCC 字典（{current}），跳過。")
        return
    try:
        tag = latest_tag()
    except Exception as e:
        if have:
            log(f"警告：檢查 OpenCC 字典更新失敗（{e}），繼續使用現有字典（{current}）。")
            return
        raise RuntimeError(f"冇 OpenCC 字典，而且無法連線 GitHub 下載：{e}") from e
    if have and state.get('tag') == tag:
        state['checked'] = today
        _write_state(state)
        log(f"OpenCC 字典已係最新版 {tag}。")
        return

    log(f"正在下載 OpenCC 字典 {tag}（現有：{current if have else '冇'}）...")
    tmp = tempfile.mkdtemp(prefix='download_', dir=BASE_DIR)
    try:
        for f in REQUIRED:
            text = _get(f'https://raw.githubusercontent.com/{REPO}/{tag}/data/dictionary/{f}').decode('utf-8')
            good, bad = clean_lines(text)
            if bad or not good:
                raise ValueError(f"{f} 格式唔符合（{len(bad)} 行無法讀取，例如 {bad[:1]}）")
            with open(os.path.join(tmp, f), 'w', encoding='utf-8', newline='\n') as out:
                out.write('\n'.join(good) + '\n')
            log(f"  {f}：{len(good)} 條")
        backup = None
        existing = [f for f in REQUIRED if os.path.exists(os.path.join(DICT_DIR, f))]
        if existing:
            name = ((state.get('tag') or 'local') + '_' + time.strftime('%Y%m%d_%H%M%S')).replace('/', '_')
            backup = os.path.join(BACKUP_DIR, name)
            os.makedirs(backup, exist_ok=True)
            for f in existing:
                shutil.copy2(os.path.join(DICT_DIR, f), os.path.join(backup, f))
            log(f"  舊字典已備份到 {backup}")
        for f in REQUIRED:
            os.replace(os.path.join(tmp, f), os.path.join(DICT_DIR, f))
        _write_state({'tag': tag, 'checked': today, 'installed': time.strftime('%Y-%m-%d %H:%M:%S'),
                      'previous_backup': backup or state.get('previous_backup')})
        log(f"OpenCC 字典已更新到 {tag}。")
    except Exception as e:
        if have:
            log(f"警告：下載 OpenCC 字典 {tag} 失敗（{e}），繼續使用現有字典（{current}）。")
            return
        raise RuntimeError(f"下載 OpenCC 字典失敗：{e}") from e
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def prepare_overrides(log=print):
    """將 opencc_overrides/*.txt 同官方字典合併成 Merged_*.txt。冇修正檔就建立只有說明嘅檔案。"""
    os.makedirs(OVERRIDE_DIR, exist_ok=True)
    total = 0
    for f in OVERRIDES:
        src = os.path.join(OVERRIDE_DIR, f)
        if not os.path.exists(src):
            with open(src, 'w', encoding='utf-8', newline='\n') as out:
                out.write(OVERRIDE_HEADER)
        with open(src, 'r', encoding='utf-8') as fin:
            good, bad = clean_lines(fin.read())
        for b in bad:
            log(f"警告：{src} 有一行格式唔啱，已略過：{b!r}（要用 Tab 分隔原文同結果）")
        merged = {}
        official = os.path.join(DICT_DIR, f)
        if os.path.exists(official):
            with open(official, 'r', encoding='utf-8') as fin:
                merged = dict(l.split('\t') for l in clean_lines(fin.read())[0])
        merged.update(l.split('\t') for l in good)
        with open(os.path.join(DICT_DIR, MERGED_PREFIX + f), 'w', encoding='utf-8', newline='\n') as out:
            out.write(''.join(f'{k}\t{v}\n' for k, v in merged.items()))   # 唔可以有空行，套件會出錯
        total += len(good)
    log(f"OpenCC 修正字典：{total} 條。")
    for f in LOCAL_DICTS:
        src, dst = os.path.join(OVERRIDE_DIR, f), os.path.join(DICT_DIR, f)
        if not os.path.exists(src):
            existing = []
            if os.path.exists(dst):
                with open(dst, 'r', encoding='utf-8') as fin:
                    existing, _bad = clean_lines(fin.read())
            with open(src, 'w', encoding='utf-8', newline='\n') as out:
                out.write(LOCAL_DICT_HEADER + ''.join(l + '\n' for l in existing))
        with open(src, 'r', encoding='utf-8') as fin:
            good, bad = clean_lines(fin.read())
        for b in bad:
            log(f"警告：{src} 有一行格式唔啱，已略過：{b!r}")
        with open(dst, 'w', encoding='utf-8', newline='\n') as out:
            out.write('\n'.join(good) + ('\n' if good else ''))
        log(f"{f}：使用 {src}（{len(good)} 條）。")


def write_configs(dict_dir=DICT_DIR, config_dir=CONFIG_DIR, overrides=True):
    """產生 custom_*.json（絕對路徑）。此 opencc 套件唔支援 json chain，utils 會分開執行 s2t 同 s2hk。"""
    os.makedirs(config_dir, exist_ok=True)
    d = os.path.abspath(dict_dir).replace(os.sep, '/') + '/'

    def txt(name):
        return {"type": "txt", "file": d + name}

    def name(f):
        return MERGED_PREFIX + f if overrides else f

    def group(phrases, chars):
        return {"type": "group", "dicts": [txt(name(phrases)), txt(name(chars))]}

    def single(f):
        return txt(name(f))

    configs = {
        'custom_s2t.json': {
            "name": "Custom Simplified to Traditional",
            "segmentation": {"type": "mmseg", "dict": txt(name("STPhrases.txt"))},
            # 詞組字典必須排在單字字典之前，否則「皇后、頭髮、一乾二淨、一隻」等詞組規則不會生效
            "conversion_chain": [{"dict": group("STPhrases.txt", "STCharacters.txt")}],
        },
        'custom_s2hk.json': {
            "name": "Custom Simplified to Hong Kong Traditional",
            "conversion_chain": [
                {"dict": {"type": "chain", "dicts": [{"type": "json", "file": "custom_s2t.json"}]}},
                {"dict": single("HKVariants.txt")},
            ],
        },
        'custom_jp2t.json': {
            "name": "Custom Japanese Shinjitai to Traditional",
            "conversion_chain": [
                {"dict": txt("JPShinjitaiCharacters.txt")},
            ],
        },
    }
    for name, content in configs.items():
        with open(os.path.join(config_dir, name), 'w', encoding='utf-8') as f:
            json.dump(content, f, ensure_ascii=False, indent=2)
    old = os.path.join(config_dir, 'custom_t2s.json')   # 已停用嘅繁轉簡設定
    if os.path.exists(old):
        os.remove(old)


def setup(force=False, log=print):
    ensure_dictionaries(force=force, log=log)
    prepare_overrides(log=log)
    write_configs()
    log("已成功建立所有動態 OpenCC 設定檔。")


# ---------------------------------------------------------------------------
# 由比較表（data/opencc_choices.xlsx）匯入「本機 / 官方」選擇
# ---------------------------------------------------------------------------
def _converter(tmpdir, name, chain):
    import opencc
    path = os.path.join(tmpdir, name + '.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({"name": name, "conversion_chain": [{"dict": c} for c in chain]}, f)
    return opencc.OpenCC(path[:-5])   # 套件會自己加 .json


def _pair_chain(dict_dir, phrases, chars):
    d = os.path.abspath(dict_dir).replace(os.sep, '/') + '/'
    return [{"type": "group", "dicts": [{"type": "txt", "file": d + phrases}, {"type": "txt", "file": d + chars}]}]


def _simplifier(dict_dir):
    """用簡轉繁單字字典反向，將繁體字轉返簡體（只用嚟產生人名嘅簡體寫法，唔使用繁轉簡字典）。"""
    rev, keys = {}, set()
    with open(os.path.join(dict_dir, 'STCharacters.txt'), 'r', encoding='utf-8') as fin:
        for line in clean_lines(fin.read())[0]:
            k, v = line.split('	')
            keys.add(k)
            for t in v.split():
                if t != k:
                    rev.setdefault(t, k)
    return lambda text: ''.join(c if c in keys else rev.get(c, c) for c in text)


def _read_override(f):
    path = os.path.join(OVERRIDE_DIR, f)
    if not os.path.exists(path):
        return {}
    with open(path, 'r', encoding='utf-8') as fin:
        good, _bad = clean_lines(fin.read())
    return dict(line.split('\t') for line in good)


def import_choices(xlsx='data/opencc_choices.xlsx', local_dir=None, log=print):
    """「本機」→ 寫入修正字典（用本機舊字典嘅結果）；「官方」→ 由修正字典刪走；留空 → 唔改。"""
    import openpyxl
    if local_dir is None:
        backups = sorted(glob.glob(os.path.join(BACKUP_DIR, 'local_*')))
        if not backups:
            raise RuntimeError(f"搵唔到本機舊字典備份（{BACKUP_DIR}/local_*），請用 --from 指定資料夾")
        local_dir = backups[-1]
    log(f"本機字典：{local_dir}")
    with tempfile.TemporaryDirectory(prefix='opencc_import_') as tmp:
        for f in REQUIRED:   # 舊字典可能有註解行，先清理
            src = os.path.join(local_dir, f)
            if os.path.exists(src):
                with open(src, 'r', encoding='utf-8') as fin:
                    good, _bad = clean_lines(fin.read())
                with open(os.path.join(tmp, f), 'w', encoding='utf-8', newline='\n') as out:
                    out.write('\n'.join(good) + '\n')
        loc_s2t = _converter(tmp, 'loc_s2t', _pair_chain(tmp, 'STPhrases.txt', 'STCharacters.txt'))
        loc_hk = _converter(tmp, 'loc_hk', [{"type": "txt", "file": os.path.abspath(os.path.join(tmp, 'HKVariants.txt')).replace(os.sep, '/')}])
        up_s2t = _converter(tmp, 'up_s2t', _pair_chain(DICT_DIR, 'STPhrases.txt', 'STCharacters.txt'))
        # 以前手動加入嘅詞語有啲用繁體寫（例如「于成龍」），簡體原文「于成龙」配對唔到 → 一併加入簡體寫法
        to_simp = _simplifier(DICT_DIR)

        files = {f: _read_override(f) for f in OVERRIDES}
        added = removed = 0
        wb = openpyxl.load_workbook(xlsx, read_only=True)
        try:
            for sheet in ('簡轉繁',):   # 「繁轉簡」頁已停用
                if sheet not in wb.sheetnames:
                    continue
                for row in wb[sheet].iter_rows(min_row=2, values_only=True):
                    if not row or len(row) < 5 or not row[1]:
                        continue
                    key, choice = str(row[1]), (row[4] or '').strip()
                    if choice not in ('本機', '官方'):
                        continue
                    own = 'STCharacters.txt' if len(key) == 1 else 'STPhrases.txt'
                    ls, us = loc_s2t.convert(key), up_s2t.convert(key)
                    places = [(own, key), ('HKVariants.txt', ls)]
                    simp = to_simp(key)
                    if simp != key:
                        places.append(('STCharacters.txt' if len(simp) == 1 else 'STPhrases.txt', simp))
                    if ls != us:
                        target, tkey, value = own, key, ls
                    else:      # 分別喺香港字形階段
                        target, tkey, value = 'HKVariants.txt', ls, loc_hk.convert(ls)
                    if choice == '本機':
                        extra = [(f, k) for f, k in places[2:]] if target != 'HKVariants.txt' else []
                        for f, k in [(target, tkey)] + extra:
                            if files[f].get(k) != value:
                                files[f][k] = value
                                added += 1
                    else:          # 官方：由所有可能嘅修正字典刪走
                        for f, k in places:
                            if k in files[f]:
                                del files[f][k]
                                removed += 1
        finally:
            wb.close()
    for f, entries in files.items():
        with open(os.path.join(OVERRIDE_DIR, f), 'w', encoding='utf-8', newline='\n') as out:
            out.write(OVERRIDE_HEADER + ''.join(f'{k}\t{v}\n' for k, v in sorted(entries.items())))
    log(f"修正字典：新增/更新 {added} 條，刪除 {removed} 條。")
    prepare_overrides(log=log)


if __name__ == '__main__':
    # 用法（喺專案根目錄）：
    #   python -m others.opencc_dict update                     即時檢查並下載最新官方字典
    #   python -m others.opencc_dict import-choices [xlsx] [--from 本機字典資料夾]
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    args = sys.argv[1:]
    cmd = args[0] if args else 'update'
    if cmd == 'update':
        setup(force=True)
    elif cmd == 'import-choices':
        src = None
        if '--from' in args:
            i = args.index('--from')
            src = args[i + 1]
            args = args[:i] + args[i + 2:]
        import_choices(args[1] if len(args) > 1 else 'data/opencc_choices.xlsx', local_dir=src)
    else:
        print(__doc__)
