import re
import time
import random
import os
import sys
import subprocess

# === 1. 自動安裝套件與環境檢測 ===
def install_package(package):
    try:
        __import__(package)
    except ImportError:
        print(f"正在自動安裝 {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except Exception as e:
            print(f" [!] 安裝 {package} 失敗: {e}")

def check_environment():
    """檢查並安裝必要的套件與瀏覽器"""
    print("正在檢查執行環境...")
    requirements = ['playwright']
    for req in requirements:
        install_package(req)
    
    # 檢查 Playwright 瀏覽器是否安裝
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            try:
                p.chromium.launch(headless=True).close()
            except Exception:
                print("偵測到 Playwright 瀏覽器未安裝，正在執行安裝指令...")
                subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
    except Exception as e:
        print(f" [!] 環境檢查發生錯誤: {e}")

# 執行環境檢查
check_environment()

from playwright.sync_api import sync_playwright

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
        print(f"=== 開始 JAVHD 女優名單抓取 ===")
        
        try:
            print(">>> 正在前往 JAVHD 首頁進行初始化檢查...")
            self.page.goto(self.base_url, timeout=60000, wait_until="domcontentloaded")
            self._handle_popups()
        except Exception as e:
            print(f" [!] 首頁載入警告檢查失敗: {e}")

        for category in self.categories:
            url = f"{self.base_url}/{category}"
            print(f"\n>>> 正在進入分類: {category} ({url})")
            self._scrape_category(url)
        
        print(f"\n=== JAVHD 抓取完成，共收集到 {len(self.names)} 個不重複名字 ===")
        return list(self.names)

    def _handle_popups(self):
        """檢查並處理 JAVHD 各種彈出視窗"""
        try:
            # 1. 18+ 警告
            warning_btn = self.page.locator('.js-warningPopupAgreeBtn')
            if warning_btn.is_visible(timeout=1000):
                # print(" [Info] JAVHD: 偵測到 18+ 警告按鈕，正在點擊同意...")
                warning_btn.click()
                self.page.wait_for_load_state("domcontentloaded")
                time.sleep(0.5)

            # 2. 通知訂閱
            notify_cancel_btn = self.page.locator('#onesignal-slidedown-cancel-button')
            if notify_cancel_btn.is_visible(timeout=1000):
                # print(" [Info] JAVHD: 偵測到通知訂閱視窗，正在點擊 Cancel...")
                notify_cancel_btn.click()
                time.sleep(0.5)

            # 3. Cookie 同意
            cookie_btn = self.page.locator('.js-jcookie-allow')
            if cookie_btn.is_visible(timeout=1000):
                # print(" [Info] JAVHD: 偵測到 Cookie 同意視窗，正在點擊許可...")
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
        while True:
            self._handle_popups()

            try:
                self.page.wait_for_selector('.section__thumbs-list', timeout=10000)
            except:
                print(" [!] 找不到列表元素，可能已無資料或載入失敗。")
                break

            elements = self.page.locator('.thumb--model .thumb__title').all_inner_texts()
            new_count = 0
            for name in elements:
                clean_name = name.strip()
                if clean_name and clean_name not in self.names:
                    self.names.add(clean_name)
                    new_count += 1
            
            print(f"    [第 {page_num} 頁] 掃描到 {len(elements)} 個名字，新增 {new_count} 個 (目前總數: {len(self.names)})")

            if not self._go_to_next_page():
                print("    [Info] 已到達最後一頁或無下一頁，停止該分類抓取。")
                break
            
            page_num += 1
            time.sleep(random.uniform(1.5, 3.0))

    def _go_to_next_page(self):
        """檢查並點擊下一頁 (JAVHD 專用)"""
        next_btn_selector = '.list-pagination__button--next'
        try:
            btn = self.page.locator(next_btn_selector)
            if not btn.is_visible(): return False
            
            class_attr = btn.get_attribute("class")
            if class_attr and "list-pagination__button--disable" in class_attr:
                return False

            self._handle_popups()
            btn.click()
            return True
        except Exception as e:
            print(f" [!] 翻頁檢查發生錯誤: {e}")
            return False

class TmdbSearcher:
    def __init__(self, page):
        self.page = page
        self.base_search_url = "https://www.themoviedb.org/search/person"

    def _wait_with_countdown(self, seconds, message="等待重試"):
        """顯示倒數計時器"""
        print() # 換行
        for i in range(seconds, 0, -1):
            sys.stdout.write(f"\r [!] {message}: {i} 秒...   ")
            sys.stdout.flush()
            time.sleep(1)
        print(f"\r [i] {message} 結束，正在繼續執行...               ")

    def _is_throttled(self):
        """檢查頁面內容是否顯示限流訊息"""
        try:
            # 檢查頁面文字是否包含限流關鍵字
            content = self.page.content().lower()
            if "temporarily throttled" in content or "retry-after" in content:
                return True
            return False
        except:
            return False

    def _safe_goto(self, url):
        """
        [核心功能] 安全導航函式
        包含：無限重試機制、限流檢測與等待、隨機延遲
        """
        while True:
            try:
                # 嘗試前往網址
                self.page.goto(url, wait_until="domcontentloaded", timeout=60000)
                
                # 1. 檢查是否被限流
                if self._is_throttled():
                    self._wait_with_countdown(60, "偵測到 TMDB 限流 (Temporarily Throttled)")
                    continue # 倒數結束後，重新執行 while 迴圈，再次 goto
                
                # 2. 檢查是否有彈窗 (順便處理)
                self._handle_tmdb_popups()

                # 成功載入，跳出迴圈
                return True

            except Exception as e:
                # 捕捉所有錯誤 (Playwright Timeout, 斷線, DNS 錯誤等)
                print(f"\n [!] 連線發生錯誤: {e}")
                self._wait_with_countdown(10, "網路異常或斷線，準備重試")
                # 這裡不 break，會自動繼續 while 迴圈進行重試

    def search_actors(self, names):
        """搜尋演員並儲存結果"""
        total_names = len(names)
        print(f"\n=== 開始 TMDB 搜尋 (共 {total_names} 位) ===")
        
        output_file = "tmdb_jav_results.txt"
        
        for idx, name in enumerate(names, 1):
            # [修改] 移除這裡的總數顯示，只保留進度
            print(f"[{idx}/{total_names}] 正在搜尋: {name}")
            
            # 執行深度搜尋
            found_people = self._search_all_results(name)
            
            if found_people:
                print(f"    -> 找到 {len(found_people)} 筆結果，正在寫入 TXT...")
                self._batch_write_to_txt(output_file, found_people)
            else:
                print("    -> 未找到任何結果。")
            
            time.sleep(random.uniform(1.5, 3.0))

    def _handle_tmdb_popups(self):
        """處理 TMDB 的彈出視窗"""
        try:
            cookie_btn = self.page.locator('#onetrust-accept-btn-handler')
            if cookie_btn.is_visible(timeout=1000):
                cookie_btn.click()
                time.sleep(0.5)

            ad_close_btn = self.page.locator('#cf_exit_close')
            if ad_close_btn.is_visible(timeout=1000):
                ad_close_btn.click()
                time.sleep(0.5)
        except Exception:
            pass 

    def _search_all_results(self, search_query):
        """
        執行單一名字的搜尋，並遍歷所有分頁抓取所有結果
        """
        collected_results = []
        # 初始化搜尋總數計數器
        total_collected_count = 0
        
        # 初始搜尋頁面 URL
        current_url = f"{self.base_search_url}?query={search_query}"
        current_page = 1
        
        try:
            while True:
                # [修改] 使用 safe_goto 替代直接 goto，處理斷線和限流
                self._safe_goto(current_url)

                # 檢查是否有結果 (等待任一種視圖出現)
                try:
                    self.page.wait_for_selector('.search_results.person .item.profile, .search_results.person .card', timeout=3000)
                except:
                    # 如果第一頁就沒結果，直接返回；後續頁面沒結果則結束
                    if current_page == 1:
                        return []
                    else:
                        break

                # 抓取資料
                items = self.page.locator('.search_results.person .item.profile, .search_results.person .card').all()
                count_on_page = 0
                
                for item in items:
                    try:
                        link_el = item.locator('.name a, .content .header a').first
                        if link_el.count() > 0:
                            href = link_el.get_attribute('href')
                            found_name = link_el.inner_text().strip()
                            if href:
                                full_link = f"https://www.themoviedb.org{href}"
                                collected_results.append((found_name, full_link))
                                count_on_page += 1
                    except Exception:
                        continue 
                
                # 累加總數
                total_collected_count += count_on_page
                
                # [修改] 在此處顯示該頁抓取數量與目前總數
                print(f"       [Page {current_page}] 抓取到 {count_on_page} 筆資料... (目前總數: {total_collected_count})")

                # 取得下一頁的 URL
                next_url = self._get_next_page_url()
                
                if not next_url:
                    break
                
                # 更新 URL 準備進入下一次迴圈
                current_url = next_url
                current_page += 1
                
                # 翻頁間的隨機冷卻
                time.sleep(random.uniform(1.0, 2.0))
                
            return collected_results

        except Exception as e:
            print(f"    [!] 搜尋過程錯誤: {e}")
            return collected_results

    def _get_next_page_url(self):
        """
        獲取下一頁的 URL (如果有的話)
        不再使用點擊方式，改為解析 href 屬性，更穩定
        """
        try:
            # 限定在人物區塊
            nav = self.page.locator('div.search_results.person div.pagination_wrapper nav.pagy')
            if nav.count() == 0: return None
            if not nav.first.is_visible(): return None

            next_btn = nav.locator('a[aria-label="下一頁"]')
            if next_btn.count() == 0: return None

            is_disabled = next_btn.get_attribute("aria-disabled")
            if is_disabled == "true": return None
            
            href = next_btn.get_attribute("href")
            if not href: return None

            # 組合完整 URL
            if href.startswith("http"):
                return href
            else:
                return f"https://www.themoviedb.org{href}"

        except Exception:
            return None

    def _batch_write_to_txt(self, filename, found_people_list):
        """批次將資料寫入 TXT"""
        if not found_people_list:
            return
        try:
            with open(filename, "a", encoding="utf-8") as f:
                for _, link in found_people_list:
                    f.write(f"{link}\n")
            print(f"       (已成功寫入 {len(found_people_list)} 筆 URL)")
        except Exception as e:
            print(f" [!] 寫入 TXT 失敗: {e}")

def main():
    with sync_playwright() as p:
        print("啟動瀏覽器中...")
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
            locale="zh-TW"
        )
        page = context.new_page()

        jav_scraper = JavHdScraper(page)
        # 如需跳過 JAVHD 抓取，請手動註解下面這行並自定義 actor_names
        actor_names = jav_scraper.scrape_all()
        # actor_names = ["Test Name"]

        if not actor_names:
            print("未抓取到任何名字，程式結束。")
            browser.close()
            return

        print(f"\n準備搜尋 {len(actor_names)} 位演員...")
        time.sleep(2)

        tmdb_searcher = TmdbSearcher(page)
        tmdb_searcher.search_actors(actor_names)

        print(f"\n程式執行完畢！結果已儲存至 tmdb_jav_results.txt")
        browser.close()

if __name__ == "__main__":
    main()