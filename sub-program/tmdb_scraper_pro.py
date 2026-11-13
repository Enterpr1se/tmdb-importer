import json
import subprocess
import sys
import time
import re
import os
import random
from urllib.parse import urljoin

# === 新增：帳據與登入管理 function (參考 credentials.py 整合) ===

def load_configs():
    """
    # 觸發條件：讀取位於父目錄 (C:\\TMDB-parallel\\) 的 configs.json 設定檔
    """
    # 使用 .. 指向上一層目錄
    config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'configs.json'))
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as config_file:
            return json.load(config_file)
    return {}

def save_configs(configs):
    """
    # 觸發條件：儲存帳號資訊至父目錄 (C:\\TMDB-parallel\\) 的 configs.json
    """
    # 使用 .. 指向上一層目錄
    config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'configs.json'))
    with open(config_path, 'w', encoding='utf-8') as config_file:
        json.dump(configs, config_file, indent=4, ensure_ascii=False)

def get_tmdb_credentials(configs):
    """
    # 觸發條件：獲取 TMDB 帳號密碼。若無設定檔且用戶未輸入，則回傳 None 以便轉向無登入模式。
    """
    try:
        if 'tmdb_username' not in configs or 'tmdb_password' not in configs:
            print(" [i] 尚未設定 TMDB 帳號資訊。")
            tmdb_username = input("請輸入 TMDB 的 username (直接按 Enter 跳過登入): ").strip()
            if not tmdb_username:
                return None, None
                
            tmdb_password = input("請輸入 TMDB 的 password: ").strip()
            if not tmdb_password:
                return None, None

            save_config = input("是否希望儲存這些資料留待下次使用? (y/n): ").strip().lower()
            if save_config == 'y':
                configs['tmdb_username'] = tmdb_username
                configs['tmdb_password'] = tmdb_password
                save_configs(configs)
            return tmdb_username, tmdb_password
        else:
            return configs['tmdb_username'], configs['tmdb_password']
    except Exception as e:
        print(f" [!] 獲取憑據時發生非預期錯誤: {e}")
        return None, None

def is_tmdb_logged_in(page):
    """
    # 觸發條件：檢查 Playwright 頁面是否處於登入狀態
    """
    try:
        # 訪問登入頁面，如果已登入通常會跳轉到個人主頁 /u/username
        page.goto("https://www.themoviedb.org/login", wait_until="domcontentloaded", timeout=30000)
        time.sleep(2)
        current_url = page.url
        if '/u/' in current_url:
            print(" [i] 偵測到已登入狀態。")
            return True
        return False
    except Exception:
        return False

def tmdb_login(page, username, password):
    """
    # 觸發條件：執行 TMDB 登入。
    # 邏輯更新：加入無限重試機制，處理限流、Cloudflare 驗證與網路超時。
    """
    if not username or not password:
        return False

    print(f" >>> 正在啟動登入流程 (帳號: {username})...")
    
    while True:
        try:
            # 導航至登入頁面
            page.goto("https://www.themoviedb.org/login", wait_until="domcontentloaded", timeout=60000)
            handle_tmdb_popups(page)
            
            content = page.content()
            title = page.title()

            # 檢查是否被限流
            if is_throttled(content):
                print(f"\n [!] 登入頁面被限流。")
                wait_with_countdown(60, "登入冷卻中")
                continue

            # 檢查 Cloudflare 驗證
            if "Just a moment" in title or "Cloudflare" in title:
                print(f"\n [!] 登入時遇到 Cloudflare 驗證...")
                wait_with_countdown(20, "等待驗證")
                continue

            # 執行填寫與提交
            page.fill('input[name="username"]', username)
            page.fill('input[name="password"]', password)
            page.keyboard.press("Enter")
            
            # 等待跳轉並再次檢查狀態
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except:
                pass 

            # 檢查結果
            current_url = page.url
            current_content = page.content()

            if "/u/" in current_url:
                print(" 🎉 TMDB 登入成功！")
                return True
            
            # 如果還在登入頁面，檢查是否因為限流導致提交失敗
            if is_throttled(current_content):
                print(f"\n [!] 提交登入資訊時被限流。")
                wait_with_countdown(60, "登入提交冷卻中")
                continue

            # 檢查是否為帳密錯誤（若是人為錯誤則不應無限重試，避免帳號被鎖）
            if "Invalid username and/or password" in current_content or "無效的使用者名稱及/或密碼" in current_content:
                print(" [!] 登入失敗：帳號或密碼錯誤。將以【未登入狀態】繼續。")
                return False

            # 其他未知原因導致未登入成功，重試
            print(" [!] 登入未成功且未發現明確錯誤，準備重試...")
            wait_with_countdown(10, "重新嘗試登入")
            continue

        except PlaywrightTimeoutError:
            print(f"\n [!] 登入頁面載入超時。")
            wait_with_countdown(60, "網路超時重試登入")
            continue
        except Exception as e:
            if "target closed" in str(e).lower():
                raise e 
            print(f"\n [!] 登入過程發生錯誤: {e}")
            wait_with_countdown(60, "連線錯誤重試登入")
            continue

# === 1. 自動安裝套件與環境檢測 ===
def install_package(package):
    try:
        __import__(package)
    except ImportError:
        print(f"正在自動安裝 {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package])
        except Exception as e:
            print(f" [!] 安裝 {package} 失敗: {e}")

def check_environment():
    """檢查並安裝必要的套件與瀏覽器"""
    print("正在檢查執行環境...")
    requirements = ['playwright', 'bs4']
    for req in requirements:
        install_package(req)
    
    # 檢查 Playwright 瀏覽器是否安裝
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            try:
                # 嘗試啟動一次以確認瀏覽器存在
                p.chromium.launch(headless=True).close()
                print("Chromium 環境確認完成！")
            except Exception:
                print("偵測到 Playwright 瀏覽器未安裝，正在執行安裝指令...")
                # 解決 Windows 編碼問題的環境變數設定
                env = dict(os.environ, PYTHONIOENCODING="utf-8")
                subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"], env=env)
    except Exception as e:
        print(f" [!] 環境檢查發生錯誤: {e}")

# 執行環境檢查
check_environment()

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from bs4 import BeautifulSoup

# === 2. 輔助函式 ===

def parse_complex_pages(input_str):
    """
    解析複雜的頁數設定
    支援分隔符：空格、逗號(,)、全形逗號(，)
    支援格式：1 3 5 或 1-5 11-25 或 1 3-5 10
    """
    if not input_str or input_str.lower() == "all":
        return None, True # None 代表全部

    pages = set()
    parts = re.split(r'[,\s，]+', input_str)
    
    try:
        for part in parts:
            part = part.strip()
            if not part: continue
            
            if "-" in part:
                try:
                    start_s, end_s = part.split("-", 1)
                    start, end = int(start_s), int(end_s)
                    if start > end: start, end = end, start
                    for p in range(start, end + 1):
                        pages.add(p)
                except ValueError:
                    pass
            else:
                try:
                    pages.add(int(part))
                except ValueError:
                    pass
        
        if not pages:
            return None, True
            
        return sorted(list(pages)), False
    except Exception:
        return None, True

def normalize_tmdb_url(url):
    """
    將 TMDB 網址標準化，移除標題 slug，只保留 ID
    """
    match = re.search(r"/(movie|tv|person)/(\d+)", url)
    if match:
        media_type, media_id = match.groups()
        return f"https://www.themoviedb.org/{media_type}/{media_id}"
    return url

def is_throttled(html):
    """檢查是否被限流"""
    txt = html.lower()
    return "temporarily throttled" in txt or "wait and try again" in txt or "retry-after" in txt or "are being rate limited" in txt

def check_404_title(title):
    """嚴謹的 404 檢查"""
    if re.search(r"\b404\b", title):
        return True
    if "page not found" in title.lower():
        return True
    return False

def wait_with_countdown(seconds, message="等待重試"):
    """顯示倒數計時"""
    print()
    for i in range(seconds, 0, -1):
        sys.stdout.write(f"\r [!] {message}: {i} 秒...   ")
        sys.stdout.flush()
        time.sleep(1)
    print(f"\r [i] {message} 結束，繼續執行...               ")

# === 3. JAVHD 抓取類別 (移植自您的程式) ===

class JavHdScraper:
    def __init__(self, page):
        self.page = page
        self.base_url = "https://javhd.com/ja/model"
        self.categories = [
            "popular",    # 人氣
            "justadded",  # 最新
            "top"         # 高評價
        ]
        self.names = set()

    def scrape_all(self):
        """主抓取流程"""
        print(f"\n=== 開始 JAVHD 女優名單抓取 ===")
        
        try:
            print(">>> 正在前往 JAVHD 首頁進行初始化檢查...")
            self.page.goto(self.base_url, timeout=60000, wait_until="domcontentloaded")
            self._handle_popups()
        except Exception as e:
            print(f" [!] 首頁載入警告檢查失敗: {e}")

        for category in self.categories:
            url = f"{self.base_url}/{category}"
            print(f"\n>>> 正在進入分類: {category}")
            self._scrape_category(url)
        
        print(f"\n=== JAVHD 抓取完成，共收集到 {len(self.names)} 個不重複名字 ===")
        return list(self.names)

    def _handle_popups(self):
        """檢查並處理 JAVHD 各種彈出視窗"""
        try:
            # 1. 18+ 警告
            warning_btn = self.page.locator('.js-warningPopupAgreeBtn')
            if warning_btn.is_visible(timeout=1000):
                warning_btn.click()
                self.page.wait_for_load_state("domcontentloaded")
                time.sleep(0.5)

            # 2. 通知訂閱
            notify_cancel_btn = self.page.locator('#onesignal-slidedown-cancel-button')
            if notify_cancel_btn.is_visible(timeout=1000):
                notify_cancel_btn.click()
                time.sleep(0.5)

            # 3. Cookie 同意
            cookie_btn = self.page.locator('.js-jcookie-allow')
            if cookie_btn.is_visible(timeout=1000):
                cookie_btn.click()
                time.sleep(0.5)

        except Exception:
            pass 

    def _scrape_category(self, url):
        """處理單一分類的分頁抓取"""
        try:
            self.page.goto(url, timeout=60000, wait_until="domcontentloaded")
            self._handle_popups()
        except Exception as e:
            print(f" [!] 無法載入頁面 {url}: {e}")
            return

        page_num = 1
        empty_page_count = 0

        while True:
            self._handle_popups()

            try:
                # 等待列表元素出現，縮短 timeout 避免卡太久
                self.page.wait_for_selector('.section__thumbs-list', timeout=5000)
            except:
                print("    [Info] 此分類已無更多資料。")
                break

            elements = self.page.locator('.thumb--model .thumb__title').all_inner_texts()
            if not elements:
                print("    [Info] 讀取不到名字，可能已達末頁。")
                break

            new_count = 0
            for name in elements:
                clean_name = name.strip()
                if clean_name and clean_name not in self.names:
                    self.names.add(clean_name)
                    new_count += 1
            
            print(f"    [第 {page_num} 頁] 掃描到 {len(elements)} 個名字，新增 {new_count} 個 (目前總數: {len(self.names)})")

            # 如果連續幾頁都沒抓到新名字，可以考慮提前停止 (視需求開啟)
            # if new_count == 0: empty_page_count += 1
            # else: empty_page_count = 0
            # if empty_page_count >= 5: break

            if not self._go_to_next_page():
                break
            
            page_num += 1
            time.sleep(random.uniform(1.0, 2.0)) # 稍作休息

    def _go_to_next_page(self):
        """檢查並點擊下一頁"""
        next_btn_selector = '.list-pagination__button--next'
        try:
            btn = self.page.locator(next_btn_selector)
            if not btn.is_visible(timeout=2000): return False
            
            class_attr = btn.get_attribute("class")
            if class_attr and "list-pagination__button--disable" in class_attr:
                return False

            self._handle_popups()
            btn.click()
            return True
        except Exception:
            return False

# === 4. 核心操作邏輯 (TMDB) ===

def handle_tmdb_popups(page):
    """
    處理 TMDB 的彈出視窗
    """
    try:
        # --- Cookie 同意處理 ---
        # 觸發條件：頁面出現 Cookie 同意橫幅時，點擊「接受所有 Cookie」
        cookie_selectors = [
            '#onetrust-accept-btn-handler', # 用戶提供的標準 ID
            'button:has-text("Accept All Cookies")',
            'button:has-text("同意")',
            '.js-accept-cookies'
        ]
        for selector in cookie_selectors:
            if page.locator(selector).is_visible(timeout=500):
                page.locator(selector).click()
                time.sleep(0.5)
                break

        # --- 廣告/宣傳彈窗處理 ---
        # 觸發條件：頁面出現廣告彈窗（如 Apple TV 宣傳）時，點擊右上角關閉按鈕
        close_btns = [
            '#cf_exit_close',  # 用戶提供的廣告關閉 ID
            '.close-button',
            '#cf_exit_modal .close'
        ]
        for btn in close_btns:
            if page.locator(btn).is_visible(timeout=500):
                # 有些 div 類型的按鈕需要強制點擊或模擬點擊
                try:
                    page.locator(btn).click(timeout=1000)
                except:
                    page.evaluate(f'document.querySelector("{btn}").click()')
                time.sleep(0.5)
    except Exception:
        pass

def safe_navigate(page, url, description="頁面", on_throttle_callback=None):
    """
    通用導航函式：整合了彈窗處理、抗限流與重試
    # 觸發條件：新增 on_throttle_callback 參數，允許在發生限流時通知外部（如重置計數器）
    """
    while True:
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            handle_tmdb_popups(page)
            
            content = page.content()
            if is_throttled(content):
                print(f"\n [!] 載入 {description} 時被限流。")
                # 如果有提供 callback，則執行（例如重置 session_count）
                if on_throttle_callback:
                    on_throttle_callback()
                
                wait_with_countdown(60, "冷卻中")
                continue 

            title = page.title()
            if "Just a moment" in title or "Cloudflare" in title:
                print(f"\n [!] 遇到 Cloudflare 驗證，嘗試等待...")
                wait_with_countdown(20, "等待驗證")
                if "Just a moment" in page.title():
                    continue
                return True

            return True

        except PlaywrightTimeoutError:
            print(f"\n [!] 載入 {description} 超時。")
            wait_with_countdown(60, "網路超時重試")
            continue
        except Exception as e:
            if "target closed" in str(e).lower():
                raise e 
            print(f"\n [!] 載入 {description} 錯誤: {e}")
            wait_with_countdown(60, "連線錯誤重試")
            continue

def extract_valid_links(soup):
    """提取並標準化連結"""
    links = set()
    exclude = [
        "season", "episode", "company", "network", "collection", 
        "discuss", "edit?", "images", "videos", "remote", "export",
        "rating", "watchlist", "list", "changes"
    ]
    
    for a in soup.find_all("a", href=re.compile(r"^/(movie|tv|person)/\d+")):
        href = a["href"]
        if any(ex in href for ex in exclude): continue
        full_url = "https://www.themoviedb.org" + href
        links.add(normalize_tmdb_url(full_url))
        
    return links

def has_next_page(soup):
    """檢查下一頁"""
    nav = soup.find("nav", class_="pagy") or soup.find("div", class_="pagination_wrapper")
    if not nav: return False
    
    next_btn = nav.find("a", {"aria-label": re.compile(r"next|下一頁", re.I)}) or \
               nav.find("a", string=re.compile(r">|下一頁|next", re.I))
    
    if not next_btn: return False
    if next_btn.get("aria-disabled") == "true": return False
    return bool(next_btn.get("href"))

# === 5. 特殊任務函式 ===

def get_users_from_leaderboard(page):
    """抓取排行榜使用者"""
    url = "https://www.themoviedb.org/leaderboard"
    print(f"\n>>> 正在抓取排行榜 ({url})...")
    
    if not safe_navigate(page, url, "排行榜"):
        return set()

    try:
        time.sleep(2)
        soup = BeautifulSoup(page.content(), "html.parser")
        users = set()
        for a in soup.select('a[href^="/u/"]'):
            href = a.get("href", "")
            if any(x in href for x in ["/u/new", "/u/login", "/u/signup"]): continue
            if "/u/" in href:
                try:
                    u = href.split("/u/")[1].split("/")[0].split("?")[0]
                    if u: users.add(u)
                except: pass
        print(f">>> 排行榜找到 {len(users)} 位貢獻者")
        return users
    except Exception as e:
        print(f" [!] 排行榜解析失敗: {e}")
        return set()

def fetch_homepage_content(page):
    """抓取首頁內容 (含互動)"""
    url = "https://www.themoviedb.org/"
    print(f"\n>>> 正在前往首頁 ({url})...")
    
    if not safe_navigate(page, url, "首頁"):
        return set()
    
    try:
        print(">>> 頁面載入成功，開始執行自動操作 (捲動/點擊)...")
        
        for _ in range(3):
            page.mouse.wheel(0, 1000)
            time.sleep(0.3)
        page.mouse.wheel(0, -5000) 
        time.sleep(0.5)

        collected = set()
        
        soup = BeautifulSoup(page.content(), "html.parser")
        links = extract_valid_links(soup)
        collected.update(links)
        print(f" -> 初始掃描: {len(links)} 個連結")

        anchors = page.locator('.selector .anchor h3 a')
        count = anchors.count()
        print(f" -> 偵測到 {count} 個標籤，開始讀取...")

        for i in range(count):
            try:
                anchor = anchors.nth(i)
                label = anchor.inner_text().strip()
                handle_tmdb_popups(page)
                
                try:
                    anchor.click(force=True, timeout=2000)
                except:
                    anchor.evaluate("node => node.click()")
                
                time.sleep(random.uniform(1.2, 2.0))
                
                if is_throttled(page.content()):
                    wait_with_countdown(60, "操作被限流")
                    continue

                page.evaluate("""() => {
                    document.querySelectorAll('.scroller').forEach(div => {
                        div.scrollLeft = div.scrollWidth;
                    });
                }""")
                time.sleep(1.0) 
                
                soup = BeautifulSoup(page.content(), "html.parser")
                new_links = extract_valid_links(soup)
                before = len(collected)
                collected.update(new_links)
                print(f"    [{i+1}/{count}] '{label}': 掃描 {len(new_links)} (新增 {len(collected)-before})")
            except Exception:
                pass

        return collected
    except Exception as e:
        print(f" [!] 首頁抓取失敗: {e}")
        return set()

def search_people_on_tmdb(page, names):
    """
    在 TMDB 搜尋人名列表
    使用 safe_navigate 確保穩定性
    """
    print(f"\n=== 開始在 TMDB 搜尋 {len(names)} 位演員 ===")
    found_links = set()
    
    for idx, name in enumerate(names, 1):
        print(f"[{idx}/{len(names)}] 正在搜尋: {name}")
        
        # 搜尋 URL
        search_url = f"https://www.themoviedb.org/search/person?query={name}"
        
        # 重試迴圈 (針對每一位演員)
        while True:
            if safe_navigate(page, search_url, f"搜尋 {name}"):
                break
            # 如果 safe_navigate 內部失敗會重試，這裡 break 只是保險
        
        # 解析結果
        try:
            # 檢查是否有結果 (限定 .search_results.person)
            soup = BeautifulSoup(page.content(), "html.parser")
            
            # 抓取人物區塊的連結
            # TMDB 搜尋結果結構: .card .content .header a (href="/person/123")
            person_links = set()
            
            # 尋找結果區塊
            results_container = soup.select_one('.search_results.person')
            if results_container:
                for a in results_container.select('.card .header a, .item.profile .name a'):
                    href = a.get('href')
                    if href and '/person/' in href:
                        full_url = "https://www.themoviedb.org" + href
                        person_links.add(normalize_tmdb_url(full_url))
            
            if person_links:
                found_links.update(person_links)
                print(f"    -> 找到 {len(person_links)} 個結果")
            else:
                print("    -> 未找到結果")

        except Exception as e:
            print(f"    [!] 解析錯誤: {e}")

        # 搜尋間隔，避免過快
        time.sleep(random.uniform(1.5, 3.0))
        
    return found_links

# === 6. 主程式進入點 ===

def main():
    """
    # 觸發條件：修改休息邏輯為 10-12 頁隨機，並在限流發生時重置計數
    """
    print("\n=== TMDB 貢獻連結抓取器 (Pro Turbo + JAVHD) ===")
    print("說明：")
    print(" - 輸入 '1' 或 'leaderboard': 排行榜")
    print(" - 輸入 '2' 或 'homepage': 首頁")
    print(" - 輸入 '3' 或 'javhd': 抓取 JAVHD 女優並搜尋")
    print(" - 輸入使用者名稱 (如: Simnovo)")
    print(" - 輸入空白行並按 Enter 開始")

    input_lines = []
    while True:
        try:
            line = input("請輸入 (空行結束): ").strip()
            if not line: break
            input_lines.append(line)
        except KeyboardInterrupt: return

    if not input_lines: return

    # 1. 讀取設定檔
    configs = load_configs()
    
    # 2. 詢問登入意願
    do_login_input = input("是否需要執行 TMDB 登入以降低限流風險? (y/n): ").strip().lower()
    tmdb_user, tmdb_pw = (None, None)
    
    if do_login_input == 'y':
        tmdb_user, tmdb_pw = get_tmdb_credentials(configs)
        if not tmdb_user:
            print(" [i] 未提供有效帳號，取消登入嘗試。")

    usernames = set()
    tasks = set()
    
    for line in input_lines:
        parts = re.split(r'[,\s，]+', line)
        for part in parts:
            part = part.strip()
            if not part: continue
            
            if part == "2" or "homepage" in part.lower():
                tasks.add("homepage")
            elif part == "1" or "leaderboard" in part.lower():
                tasks.add("leaderboard")
            elif part == "3" or "javhd" in part.lower():
                tasks.add("javhd")
            elif "/u/" in part:
                try:
                    u = part.split("/u/")[1].split("/")[0].split("?")[0]
                    usernames.add(u)
                except: pass
            elif part.startswith("http"):
                print(f"忽略網址: {part}")
            else:
                usernames.add(part)

    target_pages_list = None
    fetch_all_pages = False
    if usernames or "leaderboard" in tasks:
        print("\n設定使用者貢獻抓取頁數 (Enter=全部, '5'=第5頁, '1-5'=範圍):")
        pg_input = input(">>> ").strip()
        target_pages_list, fetch_all_pages = parse_complex_pages(pg_input)
    else:
        fetch_all_pages = True

    all_links = set()

    print("\n正在啟動瀏覽器 (Headless)...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080},
            locale="zh-HK"
        )
        context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => false});")
        page = context.new_page()

        # 3. 嘗試登入流程
        if tmdb_user and tmdb_pw:
            tmdb_login(page, tmdb_user, tmdb_pw)
        else:
            print(" [i] 程式將以【未登入狀態】執行抓取任務。")

        # 4. 執行各項任務
        if "javhd" in tasks:
            jav_scraper = JavHdScraper(page)
            actor_names = jav_scraper.scrape_all()
            if actor_names:
                jav_links = search_people_on_tmdb(page, actor_names)
                all_links.update(jav_links)

        if "homepage" in tasks:
            home_links = fetch_homepage_content(page)
            all_links.update(home_links)

        if "leaderboard" in tasks:
            lb_users = get_users_from_leaderboard(page)
            usernames.update(lb_users)

        if usernames:
            print(f"\n=== 準備抓取 {len(usernames)} 位貢獻者 ===")
            for idx, user in enumerate(sorted(usernames), 1):
                print(f"[{idx}/{len(usernames)}] 使用者: {user}")
                
                consecutive_empty = 0
                session_count = 0
                
                # 初始化隨機休息閾值 (10-12)
                rest_threshold = random.randint(10, 12)

                # 定義重置 callback：當 safe_navigate 觸發限流等待時，重置計數
                def reset_session_count():
                    nonlocal session_count
                    session_count = 0
                    # print(" [Debug] 觸發限流，計數器已重置")

                if fetch_all_pages:
                    import itertools
                    iter_pages = itertools.count(1)
                else:
                    iter_pages = target_pages_list

                for cur in iter_pages:
                    url = f"https://www.themoviedb.org/u/{user}/activity?page={cur}"
                    print(f"   → 第 {cur} 頁", end="", flush=True)

                    # 檢查是否達到隨機休息閾值
                    if session_count > 0 and session_count >= rest_threshold:
                        wait_with_countdown(random.randint(5, 10), "休息中")
                        print(f"   → 繼續...", end="", flush=True)
                        session_count = 0 # 重置計數
                        rest_threshold = random.randint(10, 12) # 設定下一次的隨機閾值

                    page_html = None
                    while True:
                        # 傳入 callback，確保被限流時重置計數
                        if safe_navigate(page, url, f"第 {cur} 頁", on_throttle_callback=reset_session_count):
                            page_html = page.content()
                            break
                        else:
                            break
                    
                    if not page_html:
                        print(" [X] 跳過此頁。")
                        break

                    soup = BeautifulSoup(page_html, "html.parser")
                    new = extract_valid_links(soup)
                    
                    if new:
                        all_links.update(new)
                        print(f" ✓ 發現 {len(new)} 筆 (總計 {len(all_links)})")
                        consecutive_empty = 0
                    else:
                        title = page.title()
                        if check_404_title(title):
                            print(f" [END] 頁面 404")
                            break
                        else:
                            print(" - 無新連結")
                            consecutive_empty += 1

                    session_count += 1
                    time.sleep(random.uniform(1.5, 3.0))

                    if not has_next_page(soup) or consecutive_empty >= 20:
                        break
                print()

        browser.close()

    output_file = "tmdb_contributed_links.txt"
    if all_links:
        try:
            with open(output_file, "w", encoding="utf-8") as f:
                for link in sorted(all_links):
                    f.write(link + "\n")
            print(f"\n🎉 完成！共收集 {len(all_links)} 個連結。")
            print(f"檔案: {os.path.abspath(output_file)}")
        except Exception as e:
            print(f"存檔失敗: {e}")
    else:
        print("\n⚠️ 無資料。")

if __name__ == "__main__":
    main()