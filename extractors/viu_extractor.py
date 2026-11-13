import math
import time
import re
import logging
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from others.excel_handler import (
    save_series_info_to_excel, save_season_info_to_excel,
    save_episode_info_to_excel
)
from others.utils import add_spaces_between_en_zh_num


def _parse_duration_to_minutes(duration_text):
    """Convert HH:MM:SS or MM:SS to minutes, ceiling-rounded."""
    parts = duration_text.strip().split(':')
    try:
        if len(parts) == 3:
            total_seconds = int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            total_seconds = int(parts[0]) * 60 + int(parts[1])
        else:
            return 0
        return math.ceil(total_seconds / 60)
    except (ValueError, IndexError):
        pass
    return 0


def _expand_all_episode_descriptions(driver, panel_id):
    """Click all collapsed 'show more' buttons in a panel to load episode descriptions."""
    buttons = driver.find_elements(
        By.CSS_SELECTOR,
        f'#{panel_id} button[aria-label="show more"][aria-expanded="false"]'
    )
    for btn in buttons:
        try:
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(0.2)
        except Exception:
            pass
    if buttons:
        time.sleep(0.5)
        logging.info(f"Viu: 已展開 {len(buttons)} 個集數描述")


def _extract_episodes_from_panel(panel, season_number):
    """Extract episode data from a BeautifulSoup tab panel element."""
    episodes = []
    ep_items = panel.find_all(class_='series_ep')
    for item in ep_items:
        ep_title_el = item.find(id='series_ep_title')
        if not ep_title_el:
            continue

        full_title = ep_title_el.get_text(strip=True)
        match = re.match(r'^(\d+)\.\s*(.*)', full_title)
        if not match:
            logging.warning(f"Viu: 無法解析集數標題格式: {full_title}，跳過。")
            continue

        episode_number = int(match.group(1))
        episode_title = add_spaces_between_en_zh_num(match.group(2).strip())

        dur_el = item.find(class_='Typography-bodySm')
        duration_text = dur_el.get_text(strip=True) if dur_el else ''
        episode_runtime = _parse_duration_to_minutes(duration_text)

        desc_el = item.find(class_='Typography-bodyMd')
        episode_description = add_spaces_between_en_zh_num(desc_el.get_text(strip=True)) if desc_el else ''

        episodes.append((
            season_number, episode_number, episode_runtime,
            episode_title, episode_description,
            episode_title, episode_description, ''
        ))
        logging.info(f"Viu 第 {season_number} 季 第 {episode_number} 集: {episode_title} ({episode_runtime} 分鐘)")
    return episodes


def extract_viu_data(driver, url, output_path):
    """
    Extract TV show data from viu.com.
    Returns True (TV show) or False on failure.
    """
    logging.info(f"開始抓取 Viu 資料，目標 URL: {url}")
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.ID, 'series_title')))
        time.sleep(2)

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        title_el = soup.find(id='series_title')
        title = add_spaces_between_en_zh_num(title_el.get_text(strip=True)) if title_el else ''

        desc_el = soup.find(id='synopsis')
        description = ''
        if desc_el:
            raw = desc_el.get_text(strip=True)
            if raw.startswith('節目簡介: '):
                raw = raw[len('節目簡介: '):]
            description = add_spaces_between_en_zh_num(raw.strip())

        logging.info(f"Viu 標題: {title}, 簡介: {description[:40]}...")

        # Only select episode list tabs (ep_list_tab), ignoring 推介 (recommend_tab)
        ep_tabs = driver.find_elements(By.CSS_SELECTOR, 'button.ep_list_tab[role="tab"]')

        all_episode_data = []
        num_seasons = max(len(ep_tabs), 1)
        season_data = [
            (f'第 {i+1} 季', i+1, description, f'第 {i+1} 季', description)
            for i in range(num_seasons)
        ]

        if len(ep_tabs) > 1:
            for i, tab in enumerate(ep_tabs):
                season_number = i + 1
                driver.execute_script("arguments[0].click();", tab)
                time.sleep(1.5)
                panel_id = tab.get_attribute('aria-controls')
                if panel_id:
                    _expand_all_episode_descriptions(driver, panel_id)
                panel_soup = BeautifulSoup(driver.page_source, 'html.parser')
                active_panel = panel_soup.find(id=panel_id) if panel_id else panel_soup.find(
                    'div', attrs={'role': 'tabpanel'}, hidden=False
                )
                if active_panel:
                    all_episode_data.extend(
                        _extract_episodes_from_panel(active_panel, season_number)
                    )
        else:
            panel_id = ep_tabs[0].get_attribute('aria-controls') if ep_tabs else None
            if panel_id:
                _expand_all_episode_descriptions(driver, panel_id)
                panel_soup = BeautifulSoup(driver.page_source, 'html.parser')
                panel = panel_soup.find(id=panel_id)
            else:
                panel = soup.find('div', attrs={'role': 'tabpanel'})
            if panel:
                all_episode_data.extend(_extract_episodes_from_panel(panel, 1))

        save_series_info_to_excel(title, description, title, description, output_path)
        save_season_info_to_excel(season_data, output_path)
        if all_episode_data:
            save_episode_info_to_excel(all_episode_data, output_path)
        else:
            logging.warning("Viu: 未找到任何集數資料。")

        return True

    except Exception as e:
        logging.error(f"Viu 資料抓取失敗: {e}")
        return False
