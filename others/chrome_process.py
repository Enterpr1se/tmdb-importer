import os
import time
import psutil
import shutil
import tempfile
import logging
import platform
import subprocess
import requests
import importlib.util
import traceback
import zipfile
import stat  # 【新增】用於修改檔案權限
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.core.os_manager import OperationSystemManager
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import re
import getpass

def _create_retry_session(retries=50, backoff_factor=1, status_forcelist=None):
    """
    【新增】創建一個帶有強大重試機制的 requests Session。
    """
    if status_forcelist is None:
        status_forcelist = [429, 500, 502, 503, 504]
    
    session = requests.Session()
    retry_strategy = Retry(
        total=retries,
        backoff_factor=backoff_factor,
        status_forcelist=status_forcelist,
        allowed_methods=["HEAD", "GET", "OPTIONS"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

def check_and_update_dependencies():
    """檢查並更新所有依賴模組"""
    logger = logging.getLogger(__name__)
    dependencies = {
        'setuptools': 'setuptools', 'pandas': 'pandas', 'openpyxl': 'openpyxl',
        'requests': 'requests', 'psutil': 'psutil', 'selenium': 'selenium',
        'bs4': 'beautifulsoup4', 'opencc': 'opencc-python-reimplemented',
        'urllib3': 'urllib3', 'langid': 'langid', 'inputimeout': 'inputimeout',
        'win32com': 'pywin32', 'pythoncom': 'pywin32', 'numpy': 'numpy',
        'webdriver_manager': 'webdriver-manager', 'wikipediaapi': 'wikipedia-api',
    }
    session = _create_retry_session(retries=3)
    for module_name, package_name in dependencies.items():
        if platform.system() != "Windows" and module_name in ['win32com', 'pythoncom']: continue
        try:
            if importlib.util.find_spec(module_name) is None:
                logger.info(f"安裝 {package_name}...")
                subprocess.run([os.sys.executable, "-m", "pip", "install", "--user", package_name], check=True, capture_output=True)
        except Exception: pass

def get_system_architecture():
    return platform.machine()

def _reg_query_value(hive_name, subkey, value_name):
    """讀取 Windows registry 字串值，失敗回傳 None。"""
    if platform.system() != "Windows":
        return None
    try:
        import winreg
        hive = {'HKCU': winreg.HKEY_CURRENT_USER, 'HKLM': winreg.HKEY_LOCAL_MACHINE}[hive_name]
        with winreg.OpenKey(hive, subkey) as key:
            val, _ = winreg.QueryValueEx(key, value_name)
            return str(val) if val else None
    except OSError:
        return None


def _reg_app_path(exe_name):
    """從 App Paths 取得瀏覽器執行檔路徑。"""
    subkeys = [
        ('HKLM', rf'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}'),
        ('HKCU', rf'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}'),
        ('HKLM', rf'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}'),
    ]
    for hive, subkey in subkeys:
        val = _reg_query_value(hive, subkey, '')
        if val and os.path.exists(val):
            return val
    return None


def _read_windows_file_version(exe_path):
    """用 PowerShell 讀 exe 的 ProductVersion（不依賴已淘汰的 wmic）。"""
    if not exe_path or not os.path.exists(exe_path):
        return None
    try:
        cmd = (
            'powershell -NoProfile -Command '
            f'(Get-Item -LiteralPath "{exe_path}").VersionInfo.ProductVersion'
        )
        res = subprocess.run(cmd, capture_output=True, text=True, shell=True, timeout=8, errors='ignore')
        m = re.search(r'(\d+\.\d+\.\d+\.\d+)', (res.stdout or ''))
        if m:
            return m.group(1)
    except Exception:
        pass
    try:
        res = subprocess.run(
            [exe_path, '--product-version'],
            capture_output=True, text=True, timeout=8, errors='ignore'
        )
        m = re.search(r'(\d+\.\d+\.\d+\.\d+)', (res.stdout or '') + (res.stderr or ''))
        if m:
            return m.group(1)
    except Exception:
        pass
    try:
        res = subprocess.run(
            [exe_path, '--version'],
            capture_output=True, text=True, timeout=8, errors='ignore'
        )
        m = re.search(r'(\d+\.\d+\.\d+\.\d+)', (res.stdout or '') + (res.stderr or ''))
        if m:
            return m.group(1)
    except Exception:
        pass
    parent = os.path.basename(os.path.dirname(exe_path))
    m = re.match(r'(\d+\.\d+\.\d+\.\d+)$', parent)
    if m:
        return m.group(1)
    return None


def _first_existing(paths):
    for p in paths:
        if p and os.path.exists(p):
            return p
    return None


def find_edge_exe():
    """找出本機 msedge.exe。"""
    found = _reg_app_path('msedge.exe')
    if found:
        return found
    return _first_existing([
        os.path.join(os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)'), r'Microsoft\Edge\Application\msedge.exe'),
        os.path.join(os.environ.get('PROGRAMFILES', r'C:\Program Files'), r'Microsoft\Edge\Application\msedge.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), r'Microsoft\Edge\Application\msedge.exe'),
    ])


def find_chrome_exe():
    """找出本機 chrome.exe。"""
    found = _reg_app_path('chrome.exe')
    if found:
        return found
    return _first_existing([
        os.path.join(os.environ.get('PROGRAMFILES', r'C:\Program Files'), r'Google\Chrome\Application\chrome.exe'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)'), r'Google\Chrome\Application\chrome.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), r'Google\Chrome\Application\chrome.exe'),
    ])


def get_chrome_version():
    """獲取 Chrome 版本 (跨平台)"""
    logger = logging.getLogger(__name__)
    system = platform.system()
    if system == "Windows":
        ver = (
            _reg_query_value('HKCU', r'Software\Google\Chrome\BLBeacon', 'version')
            or _reg_query_value('HKLM', r'SOFTWARE\Google\Chrome\BLBeacon', 'version')
        )
        if ver:
            m = re.search(r'(\d+\.\d+\.\d+\.\d+)', ver)
            if m:
                return m.group(1)
        exe = find_chrome_exe()
        if exe:
            return _read_windows_file_version(exe)
        logger.warning("未找到本機 Chrome 執行檔，無法讀取版本。")
        return None
    elif system == "Darwin":
        paths = ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
        for path in paths:
            if os.path.exists(path):
                try:
                    result = subprocess.run([path, "--version"], capture_output=True, text=True)
                    match = re.search(r'Google Chrome (\d+\.\d+\.\d+\.\d+)', result.stdout)
                    if match: return match.group(1)
                except Exception:
                    pass
    elif system == "Linux":
        try:
            result = subprocess.run(["google-chrome", "--version"], capture_output=True, text=True)
            match = re.search(r'Google Chrome (\d+\.\d+\.\d+\.\d+)', result.stdout)
            if match: return match.group(1)
        except Exception:
            pass
    return None

def get_edge_version():
    """獲取 Edge 版本 (跨平台)"""
    logger = logging.getLogger(__name__)
    system = platform.system()
    if system == "Windows":
        ver = (
            _reg_query_value('HKCU', r'Software\Microsoft\Edge\BLBeacon', 'version')
            or _reg_query_value('HKLM', r'SOFTWARE\Microsoft\Edge\BLBeacon', 'version')
        )
        if ver:
            m = re.search(r'(\d+\.\d+\.\d+\.\d+)', ver)
            if m:
                return m.group(1)
        exe = find_edge_exe()
        if exe:
            return _read_windows_file_version(exe)
        logger.warning("未找到本機 Edge 執行檔，無法讀取版本。")
        return None
    elif system == "Darwin":
        path = "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
        if os.path.exists(path):
            try:
                result = subprocess.run([path, "--version"], capture_output=True, text=True)
                match = re.search(r'Microsoft Edge (\d+\.\d+\.\d+\.\d+)', result.stdout)
                if match: return match.group(1)
            except Exception:
                pass
    elif system == "Linux":
        try:
            result = subprocess.run(["microsoft-edge", "--version"], capture_output=True, text=True)
            match = re.search(r'Microsoft Edge (\d+\.\d+\.\d+\.\d+)', result.stdout)
            if match: return match.group(1)
        except Exception:
            pass
    return None

def get_latest_edge_driver_version_api(major_version):
    """
    1. 使用 Python 檢查 (API: msedgedriver.azurewebsites.net)
    Uses a short timeout with no retries — this endpoint frequently fails DNS resolution.
    """
    logger = logging.getLogger(__name__)
    url = f"https://msedgedriver.azurewebsites.net/LATEST_RELEASE_{major_version}"

    logger.info(f"步驟 1: 檢查 API 最新版本 (主版本 {major_version})...")
    try:
        resp = requests.get(url, timeout=2)
        if resp.status_code == 200:
            content = resp.content.decode('utf-16' if b'\xff\xfe' in resp.content else 'utf-8').strip()
            version = re.sub(r'[^\d\.]', '', content)
            if re.match(r'^\d+\.\d+\.\d+\.\d+$', version):
                logger.info(f"API 返回最新版本: {version}")
                return version
    except Exception as e:
        logger.warning(f"API 檢查失敗 ({e})，將嘗試其他方法。")
    return None

def _get_platform_string(driver_type):
    system = platform.system()
    machine = platform.machine().lower()
    if driver_type == 'msedgedriver':
        if system == "Windows":
            return "arm64" if 'arm64' in machine else ("win32" if '32' in machine else "win64")
        elif system == "Darwin": return "mac64_m1" if 'arm64' in machine else "mac64"
        elif system == "Linux": return "linux64"
    elif driver_type == 'chromedriver':
        if system == "Windows": return "win64" # Chrome for Testing mostly uses win64 for x64
        elif system == "Darwin": return "mac-arm64" if 'arm64' in machine else "mac-x64"
        elif system == "Linux": return "linux64"
    return None

def scrape_edge_driver_version_and_url(major_version_strict=None):
    """
    2. 使用官網檢查 msedgedriver (解析 HTML)
    返回: (version, download_url)
    """
    logger = logging.getLogger(__name__)
    url = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/?form=MA13LH#downloads"
    session = _create_retry_session(retries=5)

    logger.info(f"步驟 2: 嘗試從 Edge 官網解析版本: {url}")
    try:
        resp = session.get(url, timeout=30)
        if resp.status_code != 200: return None, None

        soup = BeautifulSoup(resp.content, 'html.parser')
        
        # 尋找 "Stable Channel" 卡片
        stable_card = None
        for card in soup.find_all('div', class_='common-card'):
            if "Stable Channel" in card.get_text():
                stable_card = card
                break
        
        if not stable_card: return None, None

        platform_key = _get_platform_string('msedgedriver')
        target_zip = f"edgedriver_{platform_key}.zip"
        
        best_ver = None
        best_url = None

        # 解析版本區塊
        for block in stable_card.find_all('div', class_='block-web-driver__versions'):
            text = block.get_text().strip()
            v_match = re.search(r'Version\s*([\d\.]+)', text)
            if not v_match: continue
            
            ver = v_match.group(1)
            
            if major_version_strict and ver.split('.')[0] != major_version_strict:
                continue

            for link in block.find_all('a', href=True):
                if target_zip in link['href']:
                    if best_ver is None or _compare_versions(ver, best_ver) > 0:
                        best_ver = ver
                        best_url = link['href']
                    break

        if best_ver:
            logger.info(f"官網解析到 Stable 版本: {best_ver}")
            return best_ver, best_url
            
    except Exception as e:
        logger.error(f"官網解析失敗: {e}")
    return None, None

def scrape_chrome_driver_version_and_url(major_version_strict=None):
    """
    5. 使用官網檢查 Chrome driver
    返回: (version, download_url)
    """
    logger = logging.getLogger(__name__)
    url = "https://googlechromelabs.github.io/chrome-for-testing/"
    session = _create_retry_session(retries=5)
    
    logger.info(f"步驟 5: 嘗試從 Chrome 官網解析版本: {url}")
    try:
        resp = session.get(url, timeout=30)
        if resp.status_code != 200: return None, None

        soup = BeautifulSoup(resp.content, 'html.parser')
        stable_section = soup.find('section', id='stable')
        if not stable_section: return None, None

        platform_key = _get_platform_string('chromedriver')
        
        detected_version = None
        code_tag = stable_section.find('code')
        if code_tag: detected_version = code_tag.get_text().strip()

        if major_version_strict and detected_version:
             if detected_version.split('.')[0] != major_version_strict:
                 logger.warning(f"Chrome 官網版本 {detected_version} 與本地 {major_version_strict} 不匹配")
                 return None, None

        download_url = None
        for row in stable_section.find_all('tr'):
            cols = [c.get_text().strip() for c in row.find_all(['th', 'td'])]
            if len(cols) >= 3 and cols[0] == 'chromedriver' and cols[1] == platform_key:
                download_url = cols[2]
                break
        
        if download_url:
            logger.info(f"官網解析到 Chrome 版本: {detected_version}")
            return detected_version, download_url

    except Exception: pass
    return None, None

def _cleanup_old_driver_folders(cache_root, driver_type, version_to_keep):
    """
    【強化版】清理舊版本驅動程式資料夾。
    1. 掃描驅動程式目錄下所有符合版本號格式的子目錄。
    2. 使用 onerror 處理唯讀檔案，並加入重試機制解決 WinError 5 (拒絕存取)。
    """
    logger = logging.getLogger(__name__)
    driver_base_path = os.path.join(cache_root, driver_type)
    
    if not os.path.exists(driver_base_path):
        return

    def on_rm_error(func, path, exc_info):
        """
        處理刪除時的錯誤：如果是權限問題，修改權限為可寫後重試。
        """
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception:
            pass

    # 觸發條件：清理除了 version_to_keep 以外的所有版本目錄
    try:
        cleaned_count = 0
        # 遍歷驅動程式根目錄
        for root, dirs, files in os.walk(driver_base_path, topdown=False):
            for d in dirs:
                # 檢查資料夾名稱是否完全匹配版本號格式 (例如 143.0.3650.75)
                if re.match(r'^\d+\.\d+\.\d+\.\d+$', d):
                    if d != version_to_keep:
                        full_path = os.path.join(root, d)
                        
                        # 加入重試機制
                        max_retries = 5
                        for attempt in range(max_retries):
                            try:
                                # 先嘗試移除唯讀屬性
                                for root_inner, dirs_inner, files_inner in os.walk(full_path):
                                    for fname in files_inner:
                                        os.chmod(os.path.join(root_inner, fname), stat.S_IWRITE)
                                    for dname in dirs_inner:
                                        os.chmod(os.path.join(root_inner, dname), stat.S_IWRITE)
                                
                                shutil.rmtree(full_path, onerror=on_rm_error)
                                logger.info(f"已清理舊版本驅動目錄: {full_path}")
                                cleaned_count += 1
                                break # 成功後跳出重試
                            except Exception as e:
                                if attempt < max_retries - 1:
                                    time.sleep(0.5) # 等待 OS 釋放控制權
                                else:
                                    logger.warning(f"無法刪除目錄 {full_path} (嘗試 {max_retries} 次均失敗): {e}")
        
        if cleaned_count > 0:
            logger.info(f"驅動程式清理完成，共移除 {cleaned_count} 個舊版本目錄。")
    except Exception as e:
        logger.error(f"執行驅動清理時發生錯誤: {e}")

def _download_and_extract(url, version, driver_type, cache_dir):
    """通用下載解壓函數"""
    logger = logging.getLogger(__name__)
    session = _create_retry_session()
    
    os_folder = "win64" if platform.system() == "Windows" else "linux64"
    if platform.system() == "Darwin": os_folder = "mac64"
    
    # 目標目錄: browserdrive/msedgedriver/win64/143.0.3650.75/
    base_ver_dir = os.path.join(cache_dir, driver_type, os_folder, version)
    os.makedirs(base_ver_dir, exist_ok=True)
    
    exe_name = f"{driver_type}.exe" if platform.system() == "Windows" else driver_type
    
    # 檢查是否已存在
    existing_path = _find_file_in_dir(base_ver_dir, exe_name)
    if existing_path:
        # 下載前/後都觸發清理，確保環境乾淨
        _cleanup_old_driver_folders(cache_dir, driver_type, version)
        return existing_path

    try:
        logger.info(f"下載中: {url}")
        zip_path = os.path.join(base_ver_dir, "driver.zip")
        with session.get(url, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(zip_path, 'wb') as f: shutil.copyfileobj(r.raw, f)
            
        with zipfile.ZipFile(zip_path, 'r') as z: z.extractall(base_ver_dir)
        try: os.remove(zip_path)
        except Exception: pass
        
        final_path = _find_file_in_dir(base_ver_dir, exe_name)
        if final_path:
            if platform.system() != "Windows":
                 try: os.chmod(final_path, 0o755)
                 except Exception: pass
            
            # 下載成功後清理舊版
            _cleanup_old_driver_folders(cache_dir, driver_type, version)
            logger.info(f"下載成功並已清理舊版本: {final_path}")
            return final_path
    except Exception as e:
        logger.error(f"下載失敗: {e}")
    return None

def _find_file_in_dir(directory, filename):
    for root, _, files in os.walk(directory):
        if filename in files: return os.path.join(root, filename)
    return None

def _compare_versions(v1, v2):
    def norm(v): return [int(x) for x in re.sub(r'(\.0+)*$','', v).split(".")]
    return (norm(v1) > norm(v2)) - (norm(v1) < norm(v2))

def get_cached_driver_path(driver_name, target_version, cache_root, strict_match=True):
    """
    3. & 6. 使用 local browser driver
    """
    logger = logging.getLogger(__name__)
    search_path = os.path.join(cache_root, driver_name)
    if not os.path.exists(search_path): return None
    
    exe_name = f"{driver_name}.exe" if platform.system() == "Windows" else driver_name
    major_ver = target_version.split('.')[0]
    
    matches = []
    for root, _, files in os.walk(search_path):
        if exe_name in files:
            ver_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', root)
            if ver_match:
                cached_ver = ver_match.group(1)
                if strict_match:
                    if cached_ver == target_version:
                        logger.info(f"快取精確命中: {cached_ver}")
                        return os.path.join(root, exe_name)
                else:
                    if cached_ver.split('.')[0] == major_ver:
                        matches.append((cached_ver, os.path.join(root, exe_name)))

    if not strict_match and matches:
        matches.sort(key=lambda x: [int(i) for i in x[0].split('.') if i.isdigit()], reverse=True)
        best = matches[0]
        logger.info(f"快取寬鬆命中 (同主版本 {major_ver}): {best[0]}")
        return best[1]
    return None


def get_any_cached_driver_path(driver_name, cache_root):
    """回傳快取中版本號最高的 driver exe，找不到則 None。"""
    logger = logging.getLogger(__name__)
    search_path = os.path.join(cache_root, driver_name)
    if not os.path.exists(search_path):
        return None
    exe_name = f"{driver_name}.exe" if platform.system() == "Windows" else driver_name
    matches = []
    for root, _, files in os.walk(search_path):
        if exe_name in files:
            ver_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', root)
            ver = ver_match.group(1) if ver_match else '0.0.0.0'
            matches.append((ver, os.path.join(root, exe_name)))
    if not matches:
        return None
    matches.sort(key=lambda x: [int(i) for i in x[0].split('.') if i.isdigit()], reverse=True)
    logger.info(f"使用快取中最新的 {driver_name}: {matches[0][0]}")
    return matches[0][1]

def download_via_wdm(driver_type, version, cache_dir):
    """嘗試使用 WebDriver Manager 下載"""
    logger = logging.getLogger(__name__)
    logger.info(f"嘗試使用 WDM 下載 {driver_type} {version}...")
    path = None
    try:
        if driver_type == 'msedgedriver':
            try:
                manager = EdgeChromiumDriverManager(driver_version=version)
                path = manager.install()
            except TypeError:
                manager = EdgeChromiumDriverManager(path=cache_dir, driver_version=version)
                path = manager.install()
        else:
            try:
                manager = ChromeDriverManager(driver_version=version)
                path = manager.install()
            except TypeError:
                manager = ChromeDriverManager(path=cache_dir, driver_version=version)
                path = manager.install()
        
        if path and os.path.exists(path):
            return path
    except Exception as e:
        logger.warning(f"WDM 下載失敗: {e}")
    return None

DEAD_BROWSER_MARKERS = (
    'no such window',
    'target window already closed',
    'web view not found',
    'chrome not reachable',
    'disconnected',
    'invalid session id',
    'unable to connect to renderer',
)


def is_dead_browser_error(exc):
    """True if the exception means the Selenium session/window is already gone."""
    msg = str(exc).lower()
    return any(marker in msg for marker in DEAD_BROWSER_MARKERS)


def _kill_stale_profile_processes(profile_path):
    """結束仍佔用指定 user-data-dir 的殘留 Edge/Chrome 行程，避免新視窗立刻被關掉。"""
    logger = logging.getLogger(__name__)
    profile_abs = os.path.abspath(profile_path)
    variants = {profile_abs, profile_abs.replace('\\', '/'), profile_path}
    browser_names = {
        'msedge.exe', 'chrome.exe', 'msedgedriver.exe', 'chromedriver.exe',
        'microsoftedge.exe',
    }
    killed = 0
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            name = (proc.info.get('name') or '').lower()
            if name not in browser_names:
                continue
            cmdline = proc.info.get('cmdline') or []
            joined = ' '.join(cmdline)
            if not any(v and v in joined for v in variants):
                continue
            logger.warning(f"偵測到殘留瀏覽器行程佔用 profile，結束 PID {proc.pid} ({name})")
            proc.kill()
            killed += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    if killed:
        time.sleep(1.5)
        for lock_name in ('SingletonLock', 'SingletonSocket', 'SingletonCookie'):
            lock_path = os.path.join(profile_path, lock_name)
            try:
                if os.path.exists(lock_path):
                    os.remove(lock_path)
            except Exception:
                pass
        logger.info(f"已結束 {killed} 個佔用 {os.path.basename(profile_path)} 的殘留行程。")
    return killed


def ensure_browser_window(driver, logger=None):
    """確保 driver 有一個可用視窗；若初始視窗已關，切換到剩餘視窗或開 about:blank。"""
    log = logger or logging.getLogger(__name__)
    try:
        handles = driver.window_handles
    except Exception as e:
        log.warning(f"無法取得 window_handles: {e}")
        return False

    def _try_handle(handle, allow_internal=True):
        try:
            driver.switch_to.window(handle)
            url = driver.current_url or ''
            # Edge 內部頁（edge://newtab 等）可能係隱藏 webview，導航時會出
            # "no such window: target window already closed / web view not found"
            return allow_internal or not url.startswith(('edge://', 'chrome://'))
        except Exception:
            return False

    # 【2026-11】先用目前視窗：Edge 157 由「新分頁」第一次導航後，會留低一個隱藏嘅 edge://newtab handle
    # （排喺 window_handles 最後）。以前一律揀最後一個，會切去嗰個隱藏 handle，下一次導航即刻斷線。
    try:
        current = driver.current_window_handle
        if current in handles and _try_handle(current):
            return True
    except Exception:
        pass

    if handles:
        # 目前視窗已關：先揀一般網頁，冇先至用內部頁（Edge 常把 automation 初始視窗關掉，留下還原的 session 視窗）
        for allow_internal in (False, True):
            for handle in reversed(handles):
                if _try_handle(handle, allow_internal):
                    return True

    try:
        driver.execute_script("window.open('about:blank','_blank');")
        time.sleep(0.4)
        handles = driver.window_handles
        if handles and _try_handle(handles[-1]):
            log.info("原視窗已關閉，已開啟新的 about:blank 視窗。")
            return True
    except Exception as e:
        log.warning(f"無法以 script 開啟新視窗: {e}")

    try:
        driver.get('about:blank')
        return True
    except Exception as e:
        log.warning(f"無法導向 about:blank: {e}")
        return False


def _browser_args(profile_path, remote_port):
    return [
        '--log-level=3', '--ignore-certificate-errors', '--disable-dev-shm-usage',
        '--no-sandbox', '--disable-gpu', '--disable-web-security',
        '--disable-blink-features=AutomationControlled',
        '--no-first-run', '--no-default-browser-check',
        '--disable-session-crashed-bubble', '--hide-crash-restore-bubble',
        f'--user-data-dir={profile_path}', f'--remote-debugging-port={remote_port}',
    ]


def _launch_chromium(kind, profile_path, remote_port, driver_path=None, binary_path=None, logger=None):
    """啟動 Edge 或 Chrome。driver_path 為 None 時交給 Selenium Manager 自動找 driver。"""
    log = logger or logging.getLogger(__name__)
    os.makedirs(profile_path, exist_ok=True)
    _kill_stale_profile_processes(profile_path)

    if kind == 'edge':
        ops = EdgeOptions()
        from selenium.webdriver.edge.service import Service
        webdriver_cls = webdriver.Edge
    else:
        ops = ChromeOptions()
        from selenium.webdriver.chrome.service import Service
        webdriver_cls = webdriver.Chrome

    for o in _browser_args(profile_path, remote_port):
        ops.add_argument(o)
    ops.add_experimental_option('excludeSwitches', ['enable-automation'])
    ops.add_experimental_option('useAutomationExtension', False)
    if binary_path and os.path.exists(binary_path):
        ops.binary_location = binary_path
        log.info(f"{kind} binary_location={binary_path}")

    if driver_path:
        log.info(f"使用 {kind} driver: {driver_path}")
        driver = webdriver_cls(service=Service(driver_path), options=ops)
    else:
        log.info(f"未指定 {kind} driver 路徑，改由 Selenium Manager 自動下載/選用。")
        driver = webdriver_cls(options=ops)

    time.sleep(0.8)
    if not ensure_browser_window(driver, log):
        try:
            driver.quit()
        except Exception:
            pass
        raise Exception(f"{kind} 啟動後沒有可用視窗")
    try:
        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
        })
    except Exception as cdp_err:
        log.warning(f"{kind} CDP 注入失敗（可忽略）: {cdp_err}")
    # 【2026-11 L2】記錄本程式開啟的 driver / 瀏覽器行程，結束時（包括 Ctrl+C）一定關閉
    try:
        from others import owned_procs
        owned_procs.register_driver(driver)
    except Exception as reg_err:
        log.debug(f"記錄瀏覽器行程失敗（不影響運作）: {reg_err}")
    return driver


def _resolve_edge_driver(edge_ver, drive_cache_dir, logger):
    """盡力取得 msedgedriver 路徑；全部失敗則回傳 None（交 Selenium Manager）。"""
    driver_path = None
    dl_url = None
    target_ver = edge_ver
    if edge_ver:
        major_ver = edge_ver.split('.')[0]
        api_ver = get_latest_edge_driver_version_api(major_ver)
        if api_ver:
            target_ver = api_ver
        else:
            online_ver, online_url = scrape_edge_driver_version_and_url(major_version_strict=major_ver)
            if online_ver:
                target_ver = online_ver
                dl_url = online_url
        driver_path = get_cached_driver_path('msedgedriver', target_ver, drive_cache_dir, strict_match=True)
        if not driver_path:
            logger.info(f"本地無 msedgedriver {target_ver}，準備下載...")
            if dl_url:
                driver_path = _download_and_extract(dl_url, target_ver, 'msedgedriver', drive_cache_dir)
            if not driver_path:
                driver_path = download_via_wdm('msedgedriver', target_ver, drive_cache_dir)
        if not driver_path:
            logger.warning("精確版本下載失敗，嘗試同主版本快取...")
            driver_path = get_cached_driver_path('msedgedriver', edge_ver, drive_cache_dir, strict_match=False)
    if not driver_path:
        driver_path = get_any_cached_driver_path('msedgedriver', drive_cache_dir)
    return driver_path


def _resolve_chrome_driver(chrome_ver, drive_cache_dir, logger):
    """盡力取得 chromedriver 路徑；全部失敗則回傳 None（交 Selenium Manager）。"""
    driver_path = None
    dl_url = None
    target_ver = chrome_ver
    if chrome_ver:
        major_ver = chrome_ver.split('.')[0]
        online_ver, online_url = scrape_chrome_driver_version_and_url(major_version_strict=major_ver)
        if online_ver:
            target_ver = online_ver
            dl_url = online_url
        driver_path = get_cached_driver_path('chromedriver', target_ver, drive_cache_dir, strict_match=True)
        if not driver_path:
            logger.info(f"本地無精確版本 {target_ver}，準備下載...")
            if dl_url:
                driver_path = _download_and_extract(dl_url, target_ver, 'chromedriver', drive_cache_dir)
            if not driver_path:
                driver_path = download_via_wdm('chromedriver', target_ver, drive_cache_dir)
        if not driver_path:
            logger.warning("下載失敗，嘗試使用本地快取的同主版本驅動...")
            driver_path = get_cached_driver_path('chromedriver', chrome_ver, drive_cache_dir, strict_match=False)
    if not driver_path:
        driver_path = get_any_cached_driver_path('chromedriver', drive_cache_dir)
    return driver_path


def get_browser_driver(profile_suffix="", remote_port=9222, base_dir=None):
    """
    主流程 orchestrator.
    優先 Edge，失敗時一定回退 Chrome（使用獨立 profile，避免沿用 Edge 的 user-data-dir）。
    即使讀不到瀏覽器版本，仍會嘗試快取 driver 或 Selenium Manager。
    """
    logger = logging.getLogger(__name__)
    if base_dir is None:
        base_dir = os.getcwd()
    drive_cache_dir = os.path.join(base_dir, 'browserdrive')
    os.makedirs(drive_cache_dir, exist_ok=True)

    edge_profile = os.path.join(base_dir, f"browser_profile{profile_suffix}")
    chrome_profile = os.path.join(base_dir, f"browser_profile{profile_suffix}_chrome")
    errors = []

    edge_exe = find_edge_exe() if platform.system() == "Windows" else None
    edge_ver = get_edge_version()
    logger.info(f"Edge 偵測: exe={edge_exe or '未找到'}, version={edge_ver or '未知'}")

    try:
        driver_path = _resolve_edge_driver(edge_ver, drive_cache_dir, logger)
        try:
            driver = _launch_chromium(
                'edge', edge_profile, remote_port,
                driver_path=driver_path, binary_path=edge_exe, logger=logger
            )
        except Exception as e1:
            if driver_path:
                logger.warning(f"指定 msedgedriver 啟動失敗，改試 Selenium Manager: {e1}")
                driver = _launch_chromium(
                    'edge', edge_profile, remote_port,
                    driver_path=None, binary_path=edge_exe, logger=logger
                )
            else:
                raise
        logger.info("Edge 啟動成功。")
        return driver, edge_profile, edge_profile
    except Exception as e:
        logger.error(f"Edge 啟動失敗，將回退至 Chrome: {e}")
        errors.append(f"Edge: {e}")

    chrome_exe = find_chrome_exe() if platform.system() == "Windows" else None
    chrome_ver = get_chrome_version()
    logger.info(f"Chrome 偵測: exe={chrome_exe or '未找到'}, version={chrome_ver or '未知'}；使用獨立 profile {chrome_profile}")

    try:
        driver_path = _resolve_chrome_driver(chrome_ver, drive_cache_dir, logger)
        try:
            driver = _launch_chromium(
                'chrome', chrome_profile, remote_port + 10,
                driver_path=driver_path, binary_path=chrome_exe, logger=logger
            )
        except Exception as e1:
            if driver_path:
                logger.warning(f"指定 chromedriver 啟動失敗，改試 Selenium Manager: {e1}")
                driver = _launch_chromium(
                    'chrome', chrome_profile, remote_port + 10,
                    driver_path=None, binary_path=chrome_exe, logger=logger
                )
            else:
                raise
        logger.info("Chrome 啟動成功（Edge 回退）。")
        return driver, chrome_profile, chrome_profile
    except Exception as e:
        logger.error(f"Chrome 啟動失敗: {e}")
        errors.append(f"Chrome: {e}")

    raise Exception("無法啟動任何瀏覽器。 " + " | ".join(errors))

def get_process_ids(driver):
    try:
        driver_process = psutil.Process(driver.service.process.pid)
        child_processes = driver_process.children(recursive=True)
        process_ids = [driver_process.pid] + [proc.pid for proc in child_processes]
        return process_ids
    except Exception:
        return []

def cleanup_browser(driver):
    if not driver:
        return
    process_ids = get_process_ids(driver)
    try:
        driver.quit()
        time.sleep(2)
    except Exception:
        pass
    for pid in process_ids:
        try:
            proc = psutil.Process(pid)
            proc.terminate()
            proc.wait(timeout=3)
        except Exception:
            pass
    # 【2026-11 L2】driver 行程已不在時 get_process_ids 會是空的，Edge 主行程仍可能殘留 → 以啟動時的記錄補殺
    try:
        from others import owned_procs
        owned_procs.kill_driver_procs(driver)
    except Exception:
        pass

if __name__ == "__main__":
    try:
        driver, _, _ = get_browser_driver()
        driver.quit()
        print("Test passed.")
    except Exception as e:
        print(f"Test failed: {e}")