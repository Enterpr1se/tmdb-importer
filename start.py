import sys
import subprocess
import logging
import importlib.util
import traceback
import time
import os
import functools
import shutil
import contextlib # 【新增】導入 contextlib 模組，用於建立上下文管理器
import json # 【新增】導入 json 模組

"""
graph TD
    A[Start: start.py] --> B_SUB;

    subgraph B_SUB [1. 初始化 & 設定]
        direction LR
        B1[檢查並安裝<br>依賴模組] --> B2[下載 OpenCC<br>字典與設定檔] --> B3[運行時修補<br>OpenCC] --> B4[載入其他模組];
    end

    B_SUB --> C[2. 讀取使用者輸入<br>模式 (Single/Bulk), 是否覆蓋 (y/n/ask)];
    C --> D{3. 執行<br>tmdb_importer.main()};
    D --> E{4. 判斷運行模式};

    E -- Single Mode --> S_SUB[單一模式];
    E -- Bulk Mode --> B_BULK_SUB[並行模式];

    subgraph S_SUB [單一模式]
        direction TB
        S_START[開始] --> S_LOOP{循環處理};
        S_LOOP -- 輸入 URL --> S_EXTRACT_CALL["<b>執行資料提取 & 決策</b><br>(詳見 B. 資料提取子圖)"];
        S_EXTRACT_CALL -- 產生 video_detail.xlsx --> S_UPLOAD_CALL["<b>執行資料上傳 & 覆蓋</b><br>(詳見 C. 資料上傳子圖)"];
        S_UPLOAD_CALL --> S_ASK_CONTINUE{詢問是否繼續?};
        S_ASK_CONTINUE -- 是 --> S_LOOP;
        S_ASK_CONTINUE -- 否 --> S_CLEANUP[清理瀏覽器];
        S_LOOP -- 輸入 eof --> S_CLEANUP;
    end

    subgraph B_BULK_SUB [並行模式 (多執行緒)]
        direction TB
        B_START[開始] --> B_INIT[初始化: 清理快取, 正規化URL<br>建立佇列 & 信號量, 啟動執行緒];
        B_INIT --> M_THREAD["<b>主執行緒:</b><br>讀取 bulk_videos.xlsx<br>將任務放入<b>提取佇列</b>"];
        M_THREAD --> E_THREAD["<b>提取執行緒:</b><br>從<b>提取佇列</b>取任務<br>執行<b>資料提取 & 決策 (詳見 B)</b><br>產出快取檔<br>將任務放入<b>上傳佇列</b>"];
        E_THREAD --> U_THREAD["<b>上傳執行緒:</b><br>從<b>上傳佇列</b>取任務<br>執行<b>資料上傳 & 覆蓋 (詳見 C)</b><br>更新 bulk_videos.xlsx"];
        U_THREAD --> B_END[等待所有執行緒結束];
    end
    
    S_CLEANUP --> X[結束];
    B_END --> X;
"""

# 【修改】移除自定義日誌處理器，日誌設定完全由 tmdb_importer.py 處理

# 確保 log 目錄存在
os.makedirs('log', exist_ok=True)

# 【修改】移除 logging.basicConfig，避免干擾主程式的日誌設定

def pre_run_check():
    """
    【核心修正 2025-09-15】在程式啟動前執行關鍵的環境檢查。
    此函數專門用來偵測是否存在與 `opencc` 套件衝突的本地檔案，
    這是導致 `AttributeError` 的最常見原因。
    """
    conflicting_file = 'opencc.py'
    if os.path.exists(conflicting_file):
        print("="*80)
        print("【致命錯誤：偵測到檔案衝突】")
        print(f"在您的專案目錄中發現了一個名為 '{conflicting_file}' 的檔案。")
        print("這個檔案會與您安裝的 'opencc-python-reimplemented' 套件產生衝突，")
        print("導致 'AttributeError: module 'opencc' has no attribute 'OpenCC'' 錯誤。")
        print("\n【解決方案】")
        print(f"請手動刪除或重新命名專案目錄下的 '{conflicting_file}' 檔案，然後再重新運行此程式。")
        print("="*80)
        sys.exit(1)

def _ensure_user_site_on_path():
    """pip install --user 裝到 user site-packages，同一進程的 sys.path 需手動加入。
    首次運行時該目錄在 Python 啟動後才出現：要用 addsitedir 處理 .pth（pywin32 靠 .pth 才找到 pywintypes）。"""
    import site as _site
    user_site = _site.getusersitepackages()
    if user_site not in sys.path and os.path.isdir(user_site):
        sys.path.insert(0, user_site)
        _site.addsitedir(user_site)
    importlib.invalidate_caches()


def _run_pip(pip_args, description):
    """執行 pip，回傳是否成功。輸出直接印到終端，方便看到有沒有更新。"""
    print(description)
    result = subprocess.run(
        [sys.executable, "-m", "pip", *pip_args],
        check=False,
        encoding='utf-8',
        errors='replace',
    )
    return result.returncode == 0


def check_dependencies():
    """
    每日第一次啟動時檢查並升級必要的 Python 模組（含 pip 自身）；同日再次啟動只確認模組存在。
    已安裝的套件若 PyPI 有新版本會升級；已是最新則 pip 會顯示 already satisfied。
    """
    print("開始檢查並更新必要的依賴模組...")
    required_modules = {
        'setuptools': 'setuptools',
        'langid': 'langid',
        'pandas': 'pandas',
        'openpyxl': 'openpyxl',
        'opencc-python-reimplemented': 'opencc',
        'requests': 'requests',
        'httpx': 'httpx',
        'psutil': 'psutil',
        'selenium': 'selenium',
        'beautifulsoup4': 'bs4',
        'urllib3': 'urllib3',
        'inputimeout': 'inputimeout',
        'numpy': 'numpy',
        'webdriver_manager': 'webdriver_manager',
        'wikipedia-api': 'wikipediaapi',
        'SPARQLWrapper': 'SPARQLWrapper',
    }
    if os.name == 'nt':
        required_modules['pywin32'] = 'win32com'

    # 【2026-11】每日只執行一次 pip 升級；當日已檢查過且所有模組都找得到，就跳過 pip。
    # 強制更新：python start.py --update-deps
    stamp_file = os.path.join('data', '.dependency_check.json')
    today = time.strftime('%Y-%m-%d')
    force_update = '--update-deps' in sys.argv
    last_check = None
    try:
        with open(stamp_file, 'r', encoding='utf-8') as f:
            last_check = json.load(f).get('date')
    except Exception:
        pass
    if not force_update and last_check == today:
        _ensure_user_site_on_path()
        missing = [m for m in required_modules.values() if importlib.util.find_spec(m) is None]
        if not missing:
            print(f"今日 ({today}) 已檢查過依賴模組，跳過 pip 更新。（強制更新: python start.py --update-deps）")
            _verify_modules(required_modules)
            return
        print(f"發現缺少模組 {missing}，執行 pip 安裝/更新...")

    if not _run_pip(['install', '--user', '--upgrade', 'pip'], "正在檢查並更新 pip..."):
        print("警告：更新 pip 失敗，將繼續用現有 pip 更新套件。")

    packages = list(required_modules.keys())
    if not _run_pip(
        ['install', '--user', '--upgrade', *packages],
        "正在一次過檢查並更新所有依賴套件（已是最新則會略過）...",
    ):
        print("一次過更新失敗，改為逐個套件檢查更新...")
        for package in packages:
            if not _run_pip(
                ['install', '--user', '--upgrade', package],
                f"正在檢查並更新 {package}...",
            ):
                print(f"錯誤：無法安裝/更新 {package}，請手動執行：pip install --user --upgrade {package}")
                sys.exit(1)

    _ensure_user_site_on_path()
    _verify_modules(required_modules)
    try:
        os.makedirs('data', exist_ok=True)
        with open(stamp_file, 'w', encoding='utf-8') as f:
            json.dump({'date': today, 'time': time.strftime('%H:%M:%S')}, f)
    except Exception as e:
        print(f"警告：無法寫入依賴檢查記錄 {stamp_file}: {e}")


def _verify_modules(required_modules):
    """確認所有模組可導入（opencc 另檢查 OpenCC 類別，缺少時強制重裝）。"""
    for package, module_name in required_modules.items():
        check_module = 'SPARQLWrapper' if package == 'SPARQLWrapper' else module_name
        try:
            spec = importlib.util.find_spec(check_module)
            if spec is None:
                raise ImportError(f"更新後仍然找不到模組 {check_module}。請檢查 pip 安裝路徑。")
            if package == 'opencc-python-reimplemented':
                module = importlib.import_module(check_module)
                if not hasattr(module, 'OpenCC'):
                    print(f"警告：模組 '{check_module}' 缺少 OpenCC 類，將強制重新安裝...")
                    if not _run_pip(
                        ['install', '--user', '--upgrade', '--force-reinstall', package],
                        f"正在強制重新安裝 {package}...",
                    ):
                        raise AttributeError(f"重新安裝 {package} 後仍然缺少 OpenCC 類。")
                    _ensure_user_site_on_path()
                    module = importlib.import_module(check_module)
                    if not hasattr(module, 'OpenCC'):
                        raise AttributeError(f"重新安裝 {package} 後，其模組 '{check_module}' 仍然缺少 OpenCC 類。")
            print(f"模組 {package} 已就緒。")
        except Exception as e:
            print(f"檢查 {package} 時發生錯誤: {e}\n{traceback.format_exc()}")
            print(f"錯誤：無法使用 {package}，請手動執行：pip install --user --upgrade {package}")
            sys.exit(1)
    print("依賴模組檢查和更新完成。")

def setup_and_create_opencc_configs():
    """
    OpenCC 設定（詳見 others/opencc_dict.py）：
    1. 字典來自 OpenCC 官方最新正式版，每日最多檢查一次（--update-deps 強制檢查）；換版前備份舊字典。
    2. 加入修正字典 opencc_overrides/*.txt。
    3. 產生使用【絕對路徑】的 custom_*.json 設定檔。
    完全冇字典而又下載唔到時會拋出例外，由 main() 停止程式；其他失敗只會警告並沿用現有字典。
    """
    print("正在設定 OpenCC 環境...")
    from others import opencc_dict
    opencc_dict.setup(force='--update-deps' in sys.argv, log=print)

def _print_bulk_sheets(file_path=os.path.join('data', 'bulk_videos.xlsx')):
    """列出 bulk_videos.xlsx 嘅工作表，方便用編號揀。"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
    except Exception as e:
        print(f"（未能讀取 {file_path} 工作表清單: {e}）")
        return
    print("bulk_videos.xlsx 工作表：")
    for i, n in enumerate(names, 1):
        print(f"  {i}. {n}")


def get_user_input_with_timeout(prompt, timeout=120, default='ask'):
    """獲取用戶輸入，支援簡寫（已移除超時功能）"""
    user_input = input(prompt)
    
    user_input = user_input.lower().strip()
    if "是否覆蓋" in prompt:
        if user_input in ['y', 'n', 'ask']:
            return user_input
        if not user_input:
            return default
        # logging.warning 在此階段可能尚未設定
        print(f"警告：無效輸入: {user_input}，使用默認值 '{default}'")
        return default
    elif "運行模式" in prompt:
        if user_input in ['s', 'single']:
            return 'single'
        elif user_input in ['b', 'bulk']:
            return 'bulk'
        print(f"警告：無效輸入: {user_input}，使用默認值 '{default}'")
        return default
    elif "Video site" in prompt or "excel" in prompt:
        if user_input in ['f', 'eof']:
            return 'eof'
        elif user_input in ['b', 'bulk']:
            return 'excel'
        return user_input
    return user_input or default

@contextlib.contextmanager
def suppress_stderr():
    """
    【新增】一個上下文管理器，用於暫時抑制標準錯誤 (STDERR) 的輸出。
    """
    original_stderr = sys.stderr
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')
    try:
        yield
    finally:
        sys.stderr.close()
        sys.stderr = original_stderr

def main():
    """
    程式入口點
    【修改】採用了基於動態建立設定檔的 OpenCC 初始化流程。
    """
    pre_run_check()
    try:
        check_dependencies()
    except Exception as e:
        print(f"錯誤：依賴模組檢查或安裝失敗: {e}\n{traceback.format_exc()}")
        sys.exit(1)

    try:
        # 【核心修正】在導入任何其他模組之前，先執行 OpenCC 的完整設定。
        setup_and_create_opencc_configs()
    except Exception as e:
        print(f"錯誤：設定 OpenCC 環境時發生嚴重錯誤，程式將終止: {e}\n{traceback.format_exc()}")
        sys.exit(1)
        
    try:
        # 【2026-11 P2】劇集（系列/季/電影）與分集分開設定；監控面板以問題字眼自動代答（改字眼要同步 monitor_gui.AUTO_ANSWERS）
        overwrite_titles = get_user_input_with_timeout("是否覆蓋劇集標題 (系列/季/電影)？(y/n/ask, 默認 ask): ", default='ask')
        overwrite_descriptions = get_user_input_with_timeout("是否覆蓋劇集描述 (系列/季/電影)？(y/n/ask, 默認 ask): ", default='ask')
        episode_overwrite_titles = get_user_input_with_timeout("是否覆蓋分集標題？(y/n/ask, 默認 = 同劇集): ", default='') or None
        episode_overwrite_descriptions = get_user_input_with_timeout("是否覆蓋分集描述？(y/n/ask, 默認 = 同劇集): ", default='') or None
        # 【2026-11 N1】監控面板以「來源無資料時是否清空」字眼自動代答（改字眼要同步 monitor_gui.AUTO_ANSWERS）
        clear_when_source_empty = get_user_input_with_timeout("來源無資料時是否清空 TMDB 現有標題/描述？(y/n, 默認 n = 保留 TMDB，簡體會轉繁體): ", default='n')
        # 【2026-11】監控面板以「維基香港譯名取代」字眼自動代答（改字眼要同步 monitor_gui.AUTO_ANSWERS）
        wiki_name_replace = get_user_input_with_timeout("人名：是否用維基香港譯名取代 TMDB 現有香港人名？(y/n, 默認 n = 保留 TMDB): ", default='n')
        mode = get_user_input_with_timeout("請選擇運行模式 (single(s)/bulk(b), 默認 single): ", default='single')
        # 【2026-11】bulk 模式揀工作表；監控面板以「處理 bulk_videos.xlsx 邊一張工作表」字眼自動代答（改字眼要同步 monitor_gui.AUTO_ANSWERS）
        bulk_sheet = None
        if mode == 'bulk':
            _print_bulk_sheets()
            bulk_sheet = get_user_input_with_timeout("bulk 模式要處理 bulk_videos.xlsx 邊一張工作表？(名稱或編號；多張用逗號分隔，按次序處理，例如 1, 13；默認 Bulk): ", default='') or None
    except Exception as e:
        print(f"錯誤：無法收集用戶輸入，請檢查終端環境: {e}\n{traceback.format_exc()}")
        sys.exit(1)

    try:
        with suppress_stderr():
            # 在 OpenCC 設定完成後，才導入其他會使用到 OpenCC 的模組。
            # 【2026-11】移除 importlib.reload：此處是首次導入，OpenCC 設定檔已建立，
            # reload 只會令 OpenCC / 字典 / 日誌設定重複執行一次（且舊日誌 handler 未關閉）。
            import tmdb_importer
        
        print("成功導入 tmdb_importer，開始主程序...")
        tmdb_importer.main(mode=mode, overwrite_titles=overwrite_titles, overwrite_descriptions=overwrite_descriptions,
                           clear_when_source_empty=clear_when_source_empty,
                           wiki_name_replace=wiki_name_replace,
                           bulk_sheet=bulk_sheet,
                           episode_overwrite_titles=episode_overwrite_titles,
                           episode_overwrite_descriptions=episode_overwrite_descriptions)
    except Exception as e:
        print(f"錯誤：運行 tmdb_importer 失敗: {e}\n{traceback.format_exc()}")
        sys.exit(1)

def _disable_quickedit():
    """禁用 Windows 終端機 QuickEdit Mode，防止滑鼠點擊/選取時意外暫停程式。"""
    if sys.platform != 'win32':
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-10)  # STD_INPUT_HANDLE
        mode = ctypes.c_uint32()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            # 清除 ENABLE_QUICK_EDIT_MODE (0x0040)，保留 ENABLE_EXTENDED_FLAGS (0x0080)
            mode.value = (mode.value & ~0x0040) | 0x0080
            kernel32.SetConsoleMode(handle, mode.value)
    except Exception:
        pass


if __name__ == "__main__":
    _disable_quickedit()
    try:
        main()
    except KeyboardInterrupt:
        # logging 可能尚未設定，使用 print
        print("檢測到 Ctrl+C，已完成清理程序，程式現在退出。")
        sys.exit(0)
    except Exception as e:
        # logging 可能尚未設定，使用 print
        print(f"程式運行失敗: {e}\n{traceback.format_exc()}")
        print("錯誤：程式運行失敗，請檢查 tmdb_importer.log")
        sys.exit(1)
