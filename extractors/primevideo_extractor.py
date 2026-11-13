import time
import re
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from others.excel_handler import (
    save_series_info_to_excel,
    save_season_info_to_excel,
    save_episode_info_to_excel,
    save_movie_info_to_excel
)
from others.utils import add_spaces_between_en_zh_num


def convert_runtime_to_minutes(runtime_str):
    """將時間字串轉換為分鐘數"""
    if not runtime_str:
        return 0
    hours = int(m.group(1)) if (m := re.search(r'(\d+)\s*小時', runtime_str)) else 0
    minutes = int(m.group(1)) if (m := re.search(r'(\d+)\s*分鐘', runtime_str)) else 0
    return hours * 60 + minutes


def _set_language_zh_tw(driver):
    """確認介面語言為繁體中文；若否則透過 JS 提交語言切換表單。"""
    try:
        trigger = driver.find_element(
            By.CSS_SELECTOR,
            '[data-testid="pv-nav-locale-selector-dropdown-trigger"]'
        )
        label = (trigger.get_attribute('aria-label') or '').upper()
        if 'ZH' in label or '繁' in label:
            return
        switched = driver.execute_script("""
            const inp = document.querySelector('input[name="persistedLanguage"][value="zh_TW"]');
            if (inp) { inp.closest('form').submit(); return true; }
            return false;
        """)
        if switched:
            time.sleep(5)
            return
        # fallback: open dropdown and click
        logging.warning("未找到 zh_TW 隱藏表單，改為點擊下拉選單。")
        trigger.click()
        WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button[data-testid="pv-nav-locale-zh-tw"]')
            )
        ).click()
        time.sleep(5)
    except Exception as e:
        logging.error(f"設置語言為繁體中文時發生錯誤: {e}")


def _get_main_description(driver):
    """
    取得頁面主描述（劇集/季描述），排除集數列表裡的個別集描述。
    集數描述的外層 span 帶有 data-testid^="synopsis-"，主描述沒有。
    """
    # 優先：div.F3YXI2 裡的描述（多季節目主頁）
    for sel in [
        'div.F3YXI2 span.fbl-expandable-text:not([data-testid^="synopsis"]) span._1H6ABQ',
        'span.fbl-expandable-text:not([data-testid^="synopsis"]) span._1H6ABQ',
    ]:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                text = el.text.strip()
                if text:
                    return add_spaces_between_en_zh_num(text)
        except Exception:
            continue
    return ''


def _click_tab(driver, tab_testid, wait=2):
    """點擊指定 data-testid 的 tab 按鈕。"""
    try:
        btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, f'button[data-testid="{tab_testid}"]')
            )
        )
        btn.click()
        time.sleep(wait)
    except Exception as e:
        logging.warning(f"點擊 tab [{tab_testid}] 時發生錯誤: {e}")


def _wait_for_episodes(driver, timeout=15):
    """等待集數列表出現。"""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'li[data-testid="episode-list-item"]')
            )
        )
    except Exception:
        time.sleep(3)


def _get_season_links(driver):
    """取得所有季數的 URL 清單（按季序升序排列）。
    優先從「第 X 季」文字解析季數，其次從 URL ref 中的 _sN 參數。
    無論 Prime Video 回傳的是正序還是倒序（5→1），都會正確排序。
    返回格式： list of (season_num, url)
    """
    season_data = []  # list of (season_num, url)

    for sel in [
        '[data-testid="dp-season-selector"] ul li a',
        'div._3R4jka ul li a',
    ]:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                href = el.get_attribute('href')
                if not href:
                    continue
                # 優先從 span 取文字，再從整個元素取
                text = ''
                try:
                    span = el.find_element(By.CSS_SELECTOR, 'span._36qUej')
                    text = span.text.strip()
                except Exception:
                    text = el.text.strip()

                sn = None
                m = re.search(r'第\s*(\d+)\s*季', text)
                if m:
                    sn = int(m.group(1))
                else:
                    # 後備：從 ref 參數解析，例如 ..._s5
                    m2 = re.search(r'_s(\d+)', href)
                    if m2:
                        sn = int(m2.group(1))

                if sn is not None:
                    season_data.append((sn, href))
            if season_data:
                break
        except Exception:
            continue

    if not season_data:
        return []

    # 按季數升序排序
    season_data.sort(key=lambda x: x[0])
    logging.info(f"偵測到季數順序（已排序）: {[sn for sn, _ in season_data]}")
    return season_data  # list of (season_num, url)


def _detect_current_season_info(driver):
    """單季或無下拉清單時，嘗試從頁面抓取當前顯示的季數文字與數字。
    例如從 aria-label、span._36qUej 等處找「第 X 季」。
    找不到則回退 (1, '第 1 季')。
    """
    selectors = [
        '[data-testid="dp-season-selector"] input[aria-label]',
        'label[for*="season-selector"] span._36qUej',
        'label[for*="av-droplist"] span',
        'div._3R4jka input[aria-label]',
        'span._36qUej',
        '[data-testid="dp-season-selector"] label',
        'button[data-testid*="season"]',
    ]
    for sel in selectors:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                aria = el.get_attribute('aria-label') or ''
                text = (el.text or aria).strip()
                m = re.search(r'第\s*(\d+)\s*季', text)
                if m:
                    sn = int(m.group(1))
                    s_name = add_spaces_between_en_zh_num(f'第 {sn} 季')
                    return sn, s_name
        except Exception:
            continue
    # 後備
    return 1, add_spaces_between_en_zh_num('第 1 季')


def _parse_ep_num_and_title(raw_text):
    """
    解析集數標題文字。
    「第 1 集。遊戲的名字」→ (1, '遊戲的名字')
    「第 1 集」           → (1, '')
    """
    raw_text = raw_text.strip()
    m = re.match(r'^第\s*(\d+)\s*集[。：:]\s*(.*)$', raw_text, re.DOTALL)
    if m:
        return int(m.group(1)), m.group(2).strip()
    m = re.match(r'^第\s*(\d+)\s*集$', raw_text)
    if m:
        return int(m.group(1)), ''
    return None, raw_text


def _extract_episode_items(driver, season_num):
    """從當前頁面抽取所有集數資料，回傳 list of [season, ep_num, runtime, title, desc, perfect_title, perfect_desc, backdrop_url]。"""
    episodes = []
    ep_els = driver.find_elements(By.CSS_SELECTOR, 'li[data-testid="episode-list-item"]')
    logging.info(f"第 {season_num} 季找到 {len(ep_els)} 集")
    for ep_el in ep_els:
        try:
            # 標題：div.ncJWWt 是可見的標題區域；div.EymYSd 雖有 data-automation-id
            # 但為隱藏 accessibility 元素，Selenium .text 回傳空字串，故排在後備
            raw = ''
            for sel in [
                'div.ncJWWt span._36qUej',
                'div.ncJWWt h3',
                'h3 span._36qUej',
                'h3',
            ]:
                try:
                    raw = ep_el.find_element(By.CSS_SELECTOR, sel).text.strip()
                    if raw:
                        break
                except Exception:
                    continue
            logging.debug(f"  ep raw text: {repr(raw)}")

            ep_num, ep_title = _parse_ep_num_and_title(raw)
            # 後備集數：從 id 屬性取 0-indexed 值
            if ep_num is None:
                ep_id = ep_el.get_attribute('id') or ''
                m = re.search(r'-(\d+)$', ep_id)
                ep_num = (int(m.group(1)) + 1) if m else None

            # 描述
            ep_desc = ''
            for sel in [
                'span[data-testid^="synopsis-"] span._1H6ABQ',
                'div.kqn1jW',
                'span.fbl-expandable-text span._1H6ABQ',
            ]:
                try:
                    ep_desc = ep_el.find_element(By.CSS_SELECTOR, sel).text.strip()
                    if ep_desc:
                        break
                except Exception:
                    continue

            # 時長
            ep_runtime = 0
            try:
                ep_runtime = convert_runtime_to_minutes(
                    ep_el.find_element(By.CSS_SELECTOR, 'div[data-testid="episode-runtime"]').text
                )
            except Exception:
                pass

            # 劇照 URL：優先取 1080w 版本（JPEG → webp → png → img src 後備）
            backdrop_url = ''
            try:
                packshot = ep_el.find_element(By.CSS_SELECTOR, '[data-testid="episode-packshot"]')
                for src_type in ['image/jpeg', 'image/webp', 'image/png']:
                    try:
                        srcset = packshot.find_element(
                            By.CSS_SELECTOR, f'source[type="{src_type}"]'
                        ).get_attribute('srcset') or ''
                        bm = re.search(r'(https://\S+)\s+1080w', srcset)
                        if bm:
                            backdrop_url = bm.group(1)
                            break
                    except Exception:
                        continue
                if not backdrop_url:
                    try:
                        backdrop_url = packshot.find_element(
                            By.CSS_SELECTOR, 'img[data-testid="base-image"]'
                        ).get_attribute('src') or ''
                    except Exception:
                        pass
            except Exception:
                pass

            ep_title = add_spaces_between_en_zh_num(ep_title)
            ep_desc = add_spaces_between_en_zh_num(ep_desc)
            episodes.append([season_num, ep_num, ep_runtime, ep_title, ep_desc, ep_title, ep_desc, backdrop_url])
            logging.info(
                f"  S{season_num:02d}E{str(ep_num).zfill(2)}: {ep_title[:25] or '(無標題)'} | "
                f"{ep_desc[:35] or '(無描述)'} | {ep_runtime}min | {'有劇照' if backdrop_url else '無劇照'}"
            )
        except Exception as e:
            logging.error(f"捉取集數資料時發生錯誤: {e}")
    return episodes


def _extract_seasons_and_episodes(driver, output_path, fallback_description, preloaded_season_links=None):
    """提取所有季數和集數資料並儲存。"""
    seasons_all = []
    episodes_all = []

    season_data = preloaded_season_links if preloaded_season_links is not None else _get_season_links(driver)

    # 相容舊回傳格式（list of urls）與新格式（list of (num, url)）
    if season_data and isinstance(season_data[0], str):
        # 舊格式：嘗試從 URL ref 解析季數，否則退回位置（可能不準）
        parsed = []
        for url in season_data:
            m = re.search(r'_s(\d+)', url)
            if m:
                parsed.append((int(m.group(1)), url))
            else:
                parsed.append((len(parsed) + 1, url))  # 保守
        season_data = sorted(parsed, key=lambda x: x[0])  # 確保排序

    if not season_data:
        # 單季（或迷你影集）：也嘗試從頁面抓季數文字（例如 aria-label 或 span）
        _click_tab(driver, 'btf-episodes-tab')
        _wait_for_episodes(driver)
        season_desc = _get_main_description(driver) or fallback_description
        sn, s_name = _detect_current_season_info(driver)
        seasons_all.append([s_name, sn, season_desc, s_name, season_desc])
        episodes_all.extend(_extract_episode_items(driver, sn))
    else:
        for season_num, season_link in season_data:
            logging.info(f"處理第 {season_num} 季: {season_link}")
            driver.get(season_link)
            time.sleep(4)

            # 先點「詳細內容」取季描述（依需求：季 description = 詳細資料內容）
            season_desc = ''
            try:
                _click_tab(driver, 'btf-details-tab', wait=2)
                season_desc = _get_main_description(driver)
            except Exception:
                pass
            if not season_desc:
                season_desc = _get_main_description(driver) or fallback_description

            s_name = add_spaces_between_en_zh_num(f'第 {season_num} 季')
            seasons_all.append([s_name, season_num, season_desc, s_name, season_desc])

            # 再點「劇集」取集數
            _click_tab(driver, 'btf-episodes-tab', wait=2)
            _wait_for_episodes(driver)
            episodes_all.extend(_extract_episode_items(driver, season_num))

    if seasons_all:
        save_season_info_to_excel(seasons_all, output_path)
    if episodes_all:
        save_episode_info_to_excel(episodes_all, output_path)


def extract_primevideo_data(driver, url, output_path):
    """抓取 Prime Video 頁面資料並儲存至 Excel。回傳 True 代表電影，False 代表劇集。"""
    driver.get(url)
    time.sleep(5)

    _set_language_zh_tw(driver)

    try:
        # ── 節目名稱 ──────────────────────────────────────────────────
        title = ''
        for sel in ['span.bFCPl8', 'h1[data-automation-id="title"]', 'h1.title']:
            try:
                text = driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                if text:
                    title = add_spaces_between_en_zh_num(text)
                    break
            except Exception:
                continue
        if not title:
            try:
                raw = driver.find_element(
                    By.CSS_SELECTOR, 'meta[name="title"]'
                ).get_attribute('content') or ''
                raw = raw.replace('觀看 ', '').replace('Watch ', '').replace(' – Prime Video', '')
                raw = re.sub(r'\s*第.*季', '', raw).strip()
                title = add_spaces_between_en_zh_num(raw) if raw else ''
            except Exception:
                pass

        # ── 主描述 ────────────────────────────────────────────────────
        description = _get_main_description(driver)
        logging.info(f"名稱: {title}")
        logging.info(f"簡介: {description}")

        # ── 判斷劇集 / 電影 ───────────────────────────────────────────
        has_season_selector = bool(
            driver.find_elements(By.CSS_SELECTOR, '[data-testid="dp-season-selector"]')
        )
        has_episodes_tab = bool(
            driver.find_elements(By.CSS_SELECTOR, 'button[data-testid="btf-episodes-tab"]')
        )
        is_tv_show = has_season_selector or has_episodes_tab

        if is_tv_show:
            # 取得季數連結；若有多季，先跳至「最低季」（通常是第 1 季）取正確的 TV show 描述
            # （預設頁面顯示最新季，description 若不修正會是最新季的描述）
            # 支援不完整季數（如只有 2,5,6,8），會使用實際解析到的最低季。
            season_data = _get_season_links(driver)
            # 相容舊/新格式
            if season_data and isinstance(season_data[0], str):
                # 舊格式：嘗試從 URL ref 解析季數
                parsed = []
                for url in season_data:
                    m = re.search(r'_s(\d+)', url)
                    if m:
                        parsed.append((int(m.group(1)), url))
                    else:
                        parsed.append((len(parsed) + 1, url))
                season_data = sorted(parsed, key=lambda x: x[0])

            if len(season_data) > 1:
                # 現在 season_data 已按季數排序，第一個就是最低季
                first_sn, first_url = season_data[0]
                logging.info(f"多季節目，先跳至第 {first_sn} 季取節目描述: {first_url}")
                driver.get(first_url)
                time.sleep(4)
                s1_desc = _get_main_description(driver)
                if s1_desc:
                    description = s1_desc
                    logging.info(f"第 {first_sn} 季描述（作為節目描述）: {description[:60]}")
            save_series_info_to_excel(title, description, title, description, output_path)
            _extract_seasons_and_episodes(driver, output_path, description,
                                          preloaded_season_links=season_data)
            return False
        else:
            # 電影：取時長
            runtime = 0
            for sel in [
                'span[data-automation-id="runtime-badge"]',
                'div[data-testid="episode-runtime"]',
                'span.runtime',
            ]:
                try:
                    el = driver.find_element(By.CSS_SELECTOR, sel)
                    runtime = convert_runtime_to_minutes(
                        el.get_attribute('aria-label') or el.text
                    )
                    if runtime:
                        break
                except Exception:
                    continue
            if not runtime:
                # 從 metadata badges 的 aria-label 尋找時長
                try:
                    for b in driver.find_elements(By.CSS_SELECTOR, 'div.vrq5k3 span[role="img"]'):
                        lbl = b.get_attribute('aria-label') or ''
                        if '分鐘' in lbl or '小時' in lbl:
                            runtime = convert_runtime_to_minutes(lbl)
                            break
                except Exception:
                    pass
            save_movie_info_to_excel(title, description, runtime, title, description, output_path)
            return True

    except Exception as e:
        logging.error(f"捉取 Prime Video 資料時發生錯誤: {e}")
        return False
