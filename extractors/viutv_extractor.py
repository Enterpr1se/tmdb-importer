import math
import time
import re
import logging
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from others.excel_handler import (
    save_series_info_to_excel, save_season_info_to_excel,
    save_episode_info_to_excel
)
from others.utils import add_spaces_between_en_zh_num
from others.credentials import load_configs


def _parse_duration_to_minutes(duration_text):
    """Convert MM:SS or HH:MM:SS to minutes, ceiling-rounded."""
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


def _is_time_badge(text):
    return bool(re.match(r'^\d+:\d{2}(:\d{2})?$', text.strip()))


def _extract_card_description(card) -> str:
    """Navigate the known DOM path to extract the episode description <p> text."""
    content_wrapper = card.find('div', class_='card__content-wrapper')
    if content_wrapper:
        ep_content = content_wrapper.find('div', class_='card__content--episode')
        if ep_content:
            p = ep_content.find('p')
            if p:
                return p.get_text(strip=True)
    for p in card.find_all('p'):
        cls = ' '.join(p.get('class') or [])
        if 'card__episode-content' in cls:
            return p.get_text(strip=True)
    return ''


def _nudge_episode_list_scroll(driver, direction: str = 'down') -> bool:
    cards = driver.find_elements(By.CSS_SELECTOR, 'a.card--episode')
    if not cards:
        return False
    result = driver.execute_script(
        """
        var direction = arguments[0];
        var cards = document.querySelectorAll('a.card--episode');
        if (!cards.length) return {moved: false};
        var sign = direction === 'up' ? -1 : 1;
        var moved = false;
        function tryScroll(el) {
            if (!el || el === window) return;
            var max = (el.scrollHeight || 0) - (el.clientHeight || 0);
            if (!(max > 10)) return;
            var before = el.scrollTop || 0;
            var step = Math.max(Math.floor((el.clientHeight || 400) * 0.75), 280);
            el.scrollTop = Math.max(0, Math.min(before + sign * step, max));
            if (Math.abs((el.scrollTop || 0) - before) > 2) moved = true;
        }
        var p = cards[0].parentElement;
        while (p) { tryScroll(p); p = p.parentElement; }
        tryScroll(document.scrollingElement || document.documentElement);
        tryScroll(document.body);
        var y0 = window.pageYOffset || 0;
        window.scrollBy(0, sign * 500);
        if (Math.abs((window.pageYOffset || 0) - y0) > 2) moved = true;
        try {
            cards[0].parentElement.dispatchEvent(new WheelEvent('wheel', {
                deltaY: sign * 800, bubbles: true, cancelable: true
            }));
        } catch (e) {}
        return {moved: moved};
        """,
        direction,
    ) or {}
    return bool(result.get('moved'))


def _scroll_and_collect_episode_soups(driver) -> list:
    """
    viu.tv virtualizes the episode list. Linear-TV pages have newest at the
    bottom (scroll up); VOD pages have newest on top (scroll down).
    """
    soups = []
    seen = set()
    direction = None
    end_hits = 0
    no_progress = 0
    try:
        for _ in range(120):
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            soups.append(soup)
            nums = set()
            order = []
            for card in soup.find_all('a', class_='card--episode'):
                el = card.find('div', attrs={'data-testid': 'content-card__episode-title'})
                if not el:
                    continue
                m = re.search(r'第(\d+)集', el.get_text(strip=True))
                if m:
                    n = int(m.group(1))
                    nums.add(n)
                    order.append(n)
            gained = nums - seen
            seen |= nums
            if direction is None and order:
                n = max(1, len(order) // 4)
                at_bottom = (sum(order[-n:]) / n) > (sum(order[:n]) / n)
                direction = 'up' if at_bottom else 'down'
                logging.info(
                    f"ViuTV: 集數列表方向: {'最新在底部 → 向上滾' if at_bottom else '最新在頂部 → 向下滾'}"
                )
            if gained:
                logging.info(f"ViuTV: 滾動累積 {sorted(gained)}（目前共 {len(seen)} 集）")
                no_progress = 0
                end_hits = 0
            else:
                no_progress += 1
            moved = _nudge_episode_list_scroll(driver, direction or 'up')
            if not moved:
                other = 'down' if direction == 'up' else 'up'
                logging.info(f"ViuTV: 已到盡頭，改為向{other}滾")
                direction = other
                end_hits += 1
                if end_hits >= 4:
                    logging.info(f"ViuTV: 滾動完成，共累積 {len(seen)} 集。")
                    break
            elif no_progress >= 6:
                logging.info(f"ViuTV: 滾動完成，共累積 {len(seen)} 集。")
                break
            time.sleep(1.5)
    except Exception as e:
        logging.warning(f"ViuTV: 滾動集數列表時發生問題: {e}")
    return soups


def _fetch_season_episodes(tmdb_id: str, season_num: int, headers: dict) -> list[int]:
    """Fetch sorted episode_numbers for one season from TMDB API."""
    try:
        r = requests.get(
            f'https://api.themoviedb.org/3/tv/{tmdb_id}/season/{season_num}',
            headers=headers, timeout=10
        )
        if r.status_code == 200:
            eps = r.json().get('episodes', [])
            return sorted(e['episode_number'] for e in eps if 'episode_number' in e)
    except Exception as e:
        logging.warning(f"ViuTV: 取得第 {season_num} 季集數資料失敗: {e}")
    return []


def _fetch_tmdb_season_map(tmdb_id: str, api_key: str) -> dict | None:
    """
    Query TMDB API to auto-build a season map for a multi-season show.

    Detects numbering mode automatically:
    - renumber  : each season's episodes start from E1 (standard)
    - continuous: episode numbers span across seasons (e.g. S05E53 follows S04E52)

    For continuous mode, each season's details are fetched to get the actual
    episode number ranges (viu_start / viu_end = TMDB episode numbers).
    For renumber mode, episode counts from the show summary are sufficient.

    Returns None if the show has only 1 regular season (no mapping needed).
    """
    if not tmdb_id or not api_key:
        return None
    headers = {'Authorization': f'Bearer {api_key}', 'accept': 'application/json'}
    try:
        r = requests.get(
            f'https://api.themoviedb.org/3/tv/{tmdb_id}',
            headers=headers, timeout=10
        )
        if r.status_code != 200:
            logging.warning(f"ViuTV: TMDB API 查詢節目失敗 (status {r.status_code})，無法自動建立 season map。")
            return None

        seasons = sorted(
            (s for s in r.json().get('seasons', [])
             if s.get('season_number', 0) > 0 and s.get('episode_count', 0) > 0),
            key=lambda x: x['season_number']
        )
        if len(seasons) <= 1:
            return None  # single season — no mapping needed

        # Detect numbering mode: fetch season 2's episode list and check if
        # the first episode_number is 1 (renumber) or > 1 (continuous).
        s2_eps = _fetch_season_episodes(tmdb_id, seasons[1]['season_number'], headers)
        is_continuous = bool(s2_eps) and s2_eps[0] > 1

        if is_continuous:
            # Fetch each season's actual episode numbers to build exact ranges.
            # Season 2 is already fetched; reuse it.
            season_eps_cache = {seasons[1]['season_number']: s2_eps}
            season_list = []
            for s in seasons:
                sn = s['season_number']
                eps = season_eps_cache.get(sn) or _fetch_season_episodes(tmdb_id, sn, headers)
                if not eps:
                    logging.warning(f"ViuTV: 第 {sn} 季無集數資料，跳過。")
                    continue
                season_list.append({
                    'season': sn,
                    'viu_start': eps[0],
                    'viu_end': eps[-1],
                })
            numbering = 'continuous'
        else:
            # Renumber: each season resets to E1; use cumulative counts.
            season_list = []
            viu_cursor = 1
            for s in seasons:
                ep_count = s['episode_count']
                season_list.append({
                    'season': s['season_number'],
                    'viu_start': viu_cursor,
                    'viu_end': viu_cursor + ep_count - 1,
                })
                viu_cursor += ep_count
            numbering = 'renumber'

        logging.info(
            f"ViuTV: 從 TMDB API 自動建立 season map — "
            f"{len(season_list)} 季，編號方式: {numbering}。"
        )
        return {'numbering': numbering, 'seasons': season_list}

    except Exception as e:
        logging.warning(f"ViuTV: 從 TMDB API 取得 season map 失敗: {e}")
        return None


def _load_season_map(url: str, tmdb_id: str | None = None) -> dict | None:
    """
    Look up a season map for this viu.tv URL from configs.json.

    configs.json format:
    {
      "viutv_season_maps": {
        "https://viu.tv/encore/talker-one-beside-you-3": {
          "numbering": "continuous",   // "continuous" = keep viu.tv ep numbers (S02E43)
                                       // "renumber"   = reset to E1 each season (S02E01)
          "seasons": [
            {"season": 1, "viu_start": 1,  "viu_end": 54},
            {"season": 2, "viu_start": 55, "viu_end": 108}
          ]
        }
      }
    }
    """
    try:
        configs = load_configs()
        maps = configs.get('viutv_season_maps', {})
        # Normalise URL for matching (strip trailing slash)
        url_clean = url.rstrip('/')
        for key, value in maps.items():
            if key.rstrip('/') == url_clean:
                return value
    except Exception as e:
        logging.warning(f"ViuTV: 讀取 season map 設定失敗: {e}")

    # No manual config found — try to auto-build from TMDB API
    if tmdb_id:
        try:
            configs = load_configs()
            api_key = configs.get('tmdb_api', '')
        except Exception:
            api_key = ''
        return _fetch_tmdb_season_map(tmdb_id, api_key)

    return None


def _apply_season_map(episodes_by_number: dict, season_map: dict) -> tuple[list, list]:
    """
    Map {viu_ep_num: (runtime, title, desc)} to Excel episode rows and season rows
    using the configured season map.

    Returns (episode_rows, season_nums) where:
      episode_rows = list of (season_num, tmdb_ep_num, runtime, title, desc, title, desc)
      season_nums  = sorted list of unique TMDB season numbers
    """
    numbering = season_map.get('numbering', 'renumber')
    seasons_config = season_map.get('seasons', [])

    # Build lookup: viu_ep_num → (tmdb_season, tmdb_ep_num)
    ep_map: dict[int, tuple[int, int]] = {}
    for sc in seasons_config:
        s = int(sc['season'])
        viu_start = int(sc['viu_start'])
        viu_end = int(sc['viu_end'])
        offset = (viu_start - 1) if numbering == 'renumber' else 0
        for viu_ep in range(viu_start, viu_end + 1):
            ep_map[viu_ep] = (s, viu_ep - offset)

    result = []
    for ep_num in sorted(episodes_by_number.keys()):
        runtime, title, desc, backdrop_url = episodes_by_number[ep_num]
        if ep_num in ep_map:
            season_num, tmdb_ep_num = ep_map[ep_num]
        else:
            # Not covered by any season config → fall back to S1, keep original number
            logging.warning(f"ViuTV: 第 {ep_num} 集不在 season map 範圍內，使用 S01E{ep_num} 作為後備。")
            season_num, tmdb_ep_num = 1, ep_num
        result.append((season_num, tmdb_ep_num, runtime, title, desc, title, desc, backdrop_url))
        desc_preview = desc[:40] if desc else ''
        logging.info(f"ViuTV S{season_num:02d}E{tmdb_ep_num:02d} (viu.tv 第 {ep_num} 集): {title} ({runtime} 分鐘) | {desc_preview}")

    season_nums = sorted({row[0] for row in result})
    return result, season_nums


def _read_episode_cards(soup) -> dict:
    """Parse cards into {ep_num: (is_prevent, runtime, title, desc, backdrop_url)}."""
    episodes_by_number = {}
    for card in soup.find_all('a', class_='card--episode'):
        ep_num_el = card.find('div', attrs={'data-testid': 'content-card__episode-title'})
        if not ep_num_el:
            continue
        match = re.search(r'第(\d+)集', ep_num_el.get_text(strip=True))
        if not match:
            continue
        episode_number = int(match.group(1))

        is_prevent = card.get('data-prevent-progress') == 'true'

        ep_title_el = card.find('h3', attrs={'data-testid': 'content-card__card-title'})
        episode_title = add_spaces_between_en_zh_num(ep_title_el.get_text(strip=True)) if ep_title_el else ''

        episode_description = add_spaces_between_en_zh_num(_extract_card_description(card))

        episode_runtime = 0
        for badge in card.find_all(class_='badge-indicator'):
            if 'badge--member' not in badge.get('class', []):
                text = badge.get_text(strip=True)
                if _is_time_badge(text):
                    episode_runtime = _parse_duration_to_minutes(text)
                    break

        backdrop_url = ''
        img_el = card.find('img', attrs={'data-testid': 'content-card__image'})
        if img_el:
            srcset = img_el.get('srcset', '')
            bm = re.search(r'(https://[^\s,]+)\s+1280w', srcset)
            if bm:
                backdrop_url = bm.group(1)
            elif srcset:
                urls = re.findall(r'(https://[^\s,]+)', srcset)
                if urls:
                    backdrop_url = urls[-1]

        existing = episodes_by_number.get(episode_number)
        if existing is None or (existing[0] and not is_prevent):
            episodes_by_number[episode_number] = (
                is_prevent, episode_runtime, episode_title, episode_description, backdrop_url
            )
    return episodes_by_number


def _extract_episodes(soup, season_map: dict | None = None, cards: dict | None = None):
    """
    Parse episode cards from BeautifulSoup (or a pre-merged cards dict).
    Deduplicates by episode number: prefers regular cards over data-prevent-progress ones
    (which are 足本版 / special extended cuts).

    If season_map is provided (from configs.json), episodes are mapped to the correct
    TMDB season and episode numbers. Otherwise everything goes to Season 1.
    """
    episodes_by_number = cards if cards is not None else _read_episode_cards(soup)

    # Strip is_prevent flag, keep only (runtime, title, desc, backdrop_url)
    flat = {ep_num: (data[1], data[2], data[3], data[4]) for ep_num, data in episodes_by_number.items()}

    if season_map:
        return _apply_season_map(flat, season_map)

    # Default: all Season 1, keep viu.tv episode numbers
    result = []
    for ep_num in sorted(flat.keys()):
        runtime, title, desc, backdrop_url = flat[ep_num]
        result.append((1, ep_num, runtime, title, desc, title, desc, backdrop_url))
        desc_preview = desc[:40] if desc else ''
        logging.info(f"ViuTV 第 1 季 第 {ep_num} 集: {title} ({runtime} 分鐘) | {desc_preview}")
    return result, [1]


def extract_viutv_data(driver, url, output_path, tmdb_id: str | None = None):
    """
    Extract TV show data from viu.tv.
    If tmdb_id is provided and no manual season map exists in configs.json,
    the TMDB API is queried to auto-build the season map.
    Returns True on success, False on failure.
    """
    logging.info(f"開始抓取 ViuTV 資料，目標 URL: {url}")
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'h1')))
        time.sleep(2)

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        title_el = soup.find('h1')
        title = add_spaces_between_en_zh_num(title_el.get_text(strip=True)) if title_el else ''

        desc_el = soup.find('p', class_='intro__p')
        description = add_spaces_between_en_zh_num(desc_el.get_text(strip=True)) if desc_el else ''

        logging.info(f"ViuTV 標題: {title}, 簡介: {description[:40]}...")

        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a.card--episode'))
            )
        except Exception:
            logging.warning("ViuTV: 等待集數卡片超時，繼續解析現有內容。")

        season_map = _load_season_map(url, tmdb_id)
        if season_map:
            logging.info(f"ViuTV: 使用 season map (numbering={season_map.get('numbering')}, "
                         f"{len(season_map.get('seasons', []))} 季)")
        soups = _scroll_and_collect_episode_soups(driver) or [
            BeautifulSoup(driver.page_source, 'html.parser')
        ]
        merged_cards = {}
        for soup in soups:
            for ep_num, data in _read_episode_cards(soup).items():
                existing = merged_cards.get(ep_num)
                if existing is None or (existing[0] and not data[0]):
                    merged_cards[ep_num] = data
        all_episode_data, season_nums = _extract_episodes(None, season_map, cards=merged_cards)

        season_data = [(f'第 {s} 季', s, description, f'第 {s} 季', description) for s in season_nums]

        save_series_info_to_excel(title, description, title, description, output_path)
        save_season_info_to_excel(season_data, output_path)
        if all_episode_data:
            save_episode_info_to_excel(all_episode_data, output_path)
        else:
            logging.warning("ViuTV: 未找到任何集數資料。")

        return True

    except Exception as e:
        logging.error(f"ViuTV 資料抓取失敗: {e}")
        return False
