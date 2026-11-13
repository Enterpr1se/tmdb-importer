"""
viutv_show_creator.py
從 viu.tv 節目頁面抓取資料，在 TMDB 建立新的電視劇條目，並上傳封面圖片。
用法: python sub-program/viutv_show_creator.py
"""

import math
import os
import re
import sys
import time
import shutil
import logging
import tempfile
import requests

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
sys.path.insert(0, project_root)

from bs4 import BeautifulSoup
from datetime import date as _date, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

from others.chrome_process import get_browser_driver, cleanup_browser
from others.credentials import load_configs, get_tmdb_credentials, is_tmdb_logged_in, login, ensure_netflix_logged_in

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
log = logging.getLogger(__name__)


_DATE_LIKE_RE = re.compile(r'^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$')


def add_spaces_between_en_zh_num(text: str) -> str:
    """Add spaces between CJK characters and Latin letters/digits, normalize whitespace.
    Date-like tokens (e.g. 2026-5-14, 14/5/2026) are left untouched."""
    if not text or not isinstance(text, str):
        return text or ''
    # Normalize line endings and full-width spaces
    text = text.replace('\r\n', '\n').replace('\r', '\n').replace('　', '\n')
    lines = text.split('\n')
    result = []
    for line in lines:
        line = re.sub(r'[ \t\f\v]+', ' ', line)
        # Skip space-insertion for date-like tokens
        if not _DATE_LIKE_RE.match(line.strip()):
            line = re.sub(r'([a-zA-Z0-9])([一-鿿])', r'\1 \2', line)
            line = re.sub(r'([一-鿿])([a-zA-Z0-9])', r'\1 \2', line)
        result.append(line.strip())
    return '\n'.join(result).strip()


# ---------------------------------------------------------------------------
# viu.tv page scraping
# ---------------------------------------------------------------------------

def _parse_episode_cards(soup) -> dict[int, dict]:
    """Parse viu.tv episode cards from a BeautifulSoup page. Returns {ep_num: ep_dict}."""
    episodes_by_num: dict[int, dict] = {}
    for card in soup.find_all('a', class_='card--episode'):
        ep_num_el = card.find('div', attrs={'data-testid': 'content-card__episode-title'})
        if not ep_num_el:
            continue
        m = re.search(r'第(\d+)集', ep_num_el.get_text(strip=True))
        if not m:
            continue
        ep_num = int(m.group(1))
        is_prevent = card.get('data-prevent-progress') == 'true'

        ep_title_el = card.find('h3', attrs={'data-testid': 'content-card__card-title'})
        ep_title = ep_title_el.get_text(strip=True) if ep_title_el else ''
        ep_desc = _extract_card_description(card)

        ep_runtime = 0
        for badge in card.find_all(class_='badge-indicator'):
            if 'badge--member' not in badge.get('class', []):
                text = badge.get_text(strip=True)
                if _is_time_badge(text):
                    ep_runtime = _parse_duration_to_minutes(text)
                    break

        ep_image_url = ''
        img_el = card.find('img', attrs={'data-testid': 'content-card__image'})
        if img_el:
            srcset = img_el.get('srcset', '')
            if srcset:
                best_w, best_url = 0, ''
                for part in srcset.split(','):
                    pieces = part.strip().rsplit(' ', 1)
                    if len(pieces) == 2:
                        try:
                            w = int(pieces[1].rstrip('w'))
                            if w > best_w:
                                best_w, best_url = w, pieces[0].strip()
                        except ValueError:
                            pass
                ep_image_url = best_url or img_el.get('src', '')
            else:
                ep_image_url = img_el.get('src', '')
            if ep_image_url and 'placeholder' not in ep_image_url:
                if ep_image_url.startswith('/'):
                    ep_image_url = 'https://www.viu.tv' + ep_image_url
                ep_image_url = _viu_full_size_url(ep_image_url)
            else:
                ep_image_url = ''

        existing = episodes_by_num.get(ep_num)
        if existing is None or (existing['is_prevent'] and not is_prevent):
            episodes_by_num[ep_num] = {
                'ep_num': ep_num,
                'is_prevent': is_prevent,
                'title': ep_title,
                'description': ep_desc,
                'runtime': ep_runtime,
                'image_url': ep_image_url,
            }
    return episodes_by_num


def _episode_dom_order(driver) -> list[int]:
    """Episode numbers in DOM order (not sorted)."""
    try:
        return driver.execute_script(
            """
            var out = [];
            document.querySelectorAll('a.card--episode').forEach(function(c) {
                var t = c.querySelector('[data-testid="content-card__episode-title"]');
                if (!t) return;
                var m = (t.textContent || '').match(/第(\\d+)集/);
                if (m) out.push(parseInt(m[1], 10));
            });
            return out;
            """
        ) or []
    except Exception:
        return []


def _newest_at_bottom(order: list[int]) -> bool:
    """True when later DOM cards have higher episode numbers (linear-TV / encore)."""
    if len(order) < 2:
        return True
    n = max(1, len(order) // 4)
    return (sum(order[-n:]) / n) > (sum(order[:n]) / n)


def _nudge_episode_list_scroll(driver, direction: str) -> bool:
    """
    Incrementally scroll the episode list by ~one viewport.
    direction='up'   → decrease scrollTop (needed when newest episodes sit at the bottom)
    direction='down' → increase scrollTop (needed when newest episodes sit at the top)
    Returns True if a scroll position actually changed.
    """
    cards = driver.find_elements(By.CSS_SELECTOR, 'a.card--episode')
    if not cards:
        return False
    result = driver.execute_script(
        """
        var direction = arguments[0];
        var cards = document.querySelectorAll('a.card--episode');
        if (!cards.length) return {moved: false, reason: 'no-cards'};
        var sign = direction === 'up' ? -1 : 1;
        var moved = false;
        var details = [];

        function tryScroll(el, label) {
            if (!el || el === window) return;
            var max = (el.scrollHeight || 0) - (el.clientHeight || 0);
            if (!(max > 10)) return;
            var before = el.scrollTop || 0;
            var step = Math.max(Math.floor((el.clientHeight || 400) * 0.75), 280);
            var next = Math.max(0, Math.min(before + sign * step, max));
            el.scrollTop = next;
            if (Math.abs((el.scrollTop || 0) - before) > 2) {
                moved = true;
                details.push(label + ':' + Math.round(before) + '→' + Math.round(el.scrollTop));
            }
        }

        var seen = [];
        var p = cards[0].parentElement;
        while (p) {
            seen.push(p);
            p = p.parentElement;
        }
        var se = document.scrollingElement || document.documentElement;
        seen.push(se, document.documentElement, document.body);

        for (var i = 0; i < seen.length; i++) {
            var node = seen[i];
            if (!node) continue;
            tryScroll(node, (node.tagName || '?') + (node.id ? '#' + node.id : ''));
        }
        var y0 = window.pageYOffset || 0;
        window.scrollBy(0, sign * 500);
        var y1 = window.pageYOffset || 0;
        if (Math.abs(y1 - y0) > 2) {
            moved = true;
            details.push('window:' + Math.round(y0) + '→' + Math.round(y1));
        }
        try {
            cards[0].parentElement.dispatchEvent(new WheelEvent('wheel', {
                deltaY: sign * 800, bubbles: true, cancelable: true
            }));
        } catch (e) {}
        return {moved: moved, details: details};
        """,
        direction,
    ) or {}
    moved = bool(result.get('moved'))
    details = result.get('details') or result.get('reason') or ''
    log.info(f"viu.tv 滾動 {direction}{' ✓' if moved else ' (無位移)'} {details}")
    return moved


def _collect_viutv_episode_cards(
    driver,
    needed_from: int | None = None,
    needed_to: int | None = None,
) -> dict[int, dict]:
    """
    Scroll the (virtualized) viu.tv episode list, accumulating cards as they
    mount/unmount. Linear-TV / encore pages put the newest episode at the
    bottom, so we scroll UP from the latest cluster; VOD pages (newest on
    top) scroll DOWN. If one direction hits the end, reverse to fill gaps.

    If needed_from/needed_to are set, stop once that inclusive range is complete.
    """
    collected: dict[int, dict] = {}
    direction: str | None = None
    no_progress = 0
    end_hits = 0
    for _ in range(120):
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        batch = _parse_episode_cards(soup)
        before = set(collected)
        for ep_num, ep in batch.items():
            existing = collected.get(ep_num)
            if existing is None or (existing.get('is_prevent') and not ep.get('is_prevent')):
                collected[ep_num] = ep
        gained = sorted(set(collected) - before)

        if direction is None:
            order = _episode_dom_order(driver)
            at_bottom = _newest_at_bottom(order)
            direction = 'up' if at_bottom else 'down'
            log.info(
                f"集數列表方向: {'最新在底部 → 向上滾' if at_bottom else '最新在頂部 → 向下滾'}"
                + (f"（DOM 首尾 {order[0]}…{order[-1]}）" if order else '')
            )

        missing: list[int] = []
        if needed_from is not None and needed_to is not None:
            expected = set(range(needed_from, needed_to + 1))
            missing = sorted(expected - set(collected))
            if gained:
                log.info(
                    f"滾動累積 {gained}；已有 {len(expected) - len(missing)}/{len(expected)} 目標集"
                    + (f"，仍缺 {missing[:12]}{'...' if len(missing) > 12 else ''}" if missing else '')
                )
            if not missing:
                log.info(f"已收集全部目標集數 {needed_from}–{needed_to}。")
                break
        elif gained:
            log.info(f"滾動累積集數 {gained}（目前共 {len(collected)} 集）")

        if gained:
            no_progress = 0
            end_hits = 0
        else:
            no_progress += 1

        moved = _nudge_episode_list_scroll(driver, direction)
        if not moved:
            other = 'down' if direction == 'up' else 'up'
            log.info(f"已到{'頂' if direction == 'up' else '底'}，改為向{other}滾")
            direction = other
            end_hits += 1
            if end_hits >= 4:
                if missing:
                    log.warning(f"滾動結束，仍缺集數: {missing}")
                else:
                    log.info(f"滾動完成，共累積 {len(collected)} 集。")
                break
        elif no_progress >= 6:
            if missing:
                log.warning(f"滾動結束，仍缺集數: {missing}")
            else:
                log.info(f"滾動完成，共累積 {len(collected)} 集。")
            break
        time.sleep(1.5)
    return collected


def _extract_card_description(card) -> str:
    """
    Extract the episode description from a card--episode <a> element.

    The HTML layout is:
      <a class="card--episode">
        <div class="card mb-3">          ← title / image
        <div class="card__content-wrapper">
          <div class="card__content--episode">
            <p class="card__episode-content line-clamp-1">...</p>   ← target

    We navigate the specific path rather than relying on class_= string matching
    (which can be unreliable when an element has multiple classes in some BS4 builds).
    """
    # Preferred path: card__content-wrapper → card__content--episode → first <p>
    content_wrapper = card.find('div', class_='card__content-wrapper')
    if content_wrapper:
        ep_content = content_wrapper.find('div', class_='card__content--episode')
        if ep_content:
            p = ep_content.find('p')
            if p:
                return p.get_text(strip=True)
    # Fallback: any <p> whose class string contains the expected substring
    for p in card.find_all('p'):
        cls = ' '.join(p.get('class') or [])
        if 'card__episode-content' in cls:
            return p.get_text(strip=True)
    return ''


def _parse_duration_to_minutes(duration_text: str) -> int:
    """Convert MM:SS or HH:MM:SS to minutes, rounded to nearest."""
    parts = duration_text.strip().split(':')
    try:
        if len(parts) == 3:
            total_seconds = int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            total_seconds = int(parts[0]) * 60 + int(parts[1])
        else:
            return 0
        return round(total_seconds / 60)
    except (ValueError, IndexError):
        return 0


def _is_time_badge(text: str) -> bool:
    return bool(re.match(r'^\d+:\d{2}(:\d{2})?$', text.strip()))


def _viu_full_size_url(url: str) -> str:
    """Strip the resized-N/ CDN prefix from a viu.tv image URL to get full resolution."""
    if not url:
        return ''
    return re.sub(r'^(https?://[^/]+/)resized-\d+/', r'\1', url)


def scrape_viutv_show(driver, viutv_url: str) -> dict:
    """
    Load a viu.tv show page and return a dict with:
      title, description, poster_url, episode_count, genres, first_air_date
    """
    log.info(f"載入 viu.tv 頁面: {viutv_url}")
    driver.get(viutv_url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.TAG_NAME, 'h1'))
    )
    time.sleep(3)

    # viu.tv og:image is always a 16:9 landscape backdrop, not a portrait poster.
    og_image = driver.execute_script(
        "var el = document.querySelector('meta[property=\"og:image\"]');"
        "return el ? el.getAttribute('content') : null;"
    ) or ''
    log.info(f"og:image: {og_image}")

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a.card--episode'))
        )
    except Exception:
        log.warning("等待集數卡片超時，繼續解析現有頁面內容。")

    episodes_by_num = _collect_viutv_episode_cards(driver)
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    title_el = soup.find('h1')
    title = title_el.get_text(strip=True) if title_el else ''

    desc_el = soup.find('p', class_='intro__p')
    description = desc_el.get_text(strip=True) if desc_el else ''

    episodes = [episodes_by_num[n] for n in sorted(episodes_by_num)]
    episode_count = len(episodes)
    for ep in episodes:
        desc_preview = ep['description'][:40] if ep['description'] else ''
        log.info(f"  第 {ep['ep_num']:>3} 集: {ep['title']} | {desc_preview}")

    # Genre tags — viu.tv lists genres as <a href="/genre/..."> links
    genres = []
    for el in soup.find_all('a', href=re.compile(r'/genre/')):
        text = el.get_text(strip=True)
        # Skip cast-member entries (contain parentheses like "王智德 (Alton)")
        # and short channel labels (Ch.99, 自家製, etc. are not genres)
        if text and '(' not in text and len(text) <= 10:
            genres.append(text)

    # First air date — try meta or structured data
    first_air_date = ''
    date_el = soup.find('meta', attrs={'property': 'video:release_date'})
    if date_el:
        first_air_date = date_el.get('content', '')
    if not first_air_date:
        date_el = soup.find('meta', attrs={'name': 'pubdate'})
        if date_el:
            first_air_date = date_el.get('content', '')

    # og:image (full-size by stripping resized-N/) is the 16:9 backdrop.
    # Fall back to in-page original_horizontal images (old viu.tv URL format).
    banner_url = _viu_full_size_url(og_image)
    if not banner_url:
        similar_section = soup.find('section', id='page__similar')
        if similar_section:
            similar_section.decompose()
        for img in soup.find_all('img'):
            src = img.get('src', '') or img.get('data-src', '')
            if 'original_horizontal' in src:
                banner_url = _viu_full_size_url(src)
                break
            srcset = img.get('srcset', '')
            if 'original_horizontal' in srcset:
                parts = [p.strip().split()[0] for p in srcset.split(',') if p.strip()]
                if parts:
                    banner_url = _viu_full_size_url(parts[-1])
                break

    # viu.tv has no portrait poster — poster_url stays empty
    poster_url = ''
    log.info(f"橫幅 (backdrop): {banner_url}")

    result = {
        'title': title,
        'description': description,
        'poster_url': poster_url,
        'banner_url': banner_url,
        'episode_count': episode_count,
        'episodes': episodes,
        'genres': genres,
        'first_air_date': first_air_date,
    }
    log.info(f"標題: {title}")
    log.info(f"集數: {episode_count}")
    log.info(f"類型: {genres}")
    log.info(f"首播日期: {first_air_date}")
    return result


# ---------------------------------------------------------------------------
# Netflix page scraping
# ---------------------------------------------------------------------------

def scrape_netflix_show(driver, netflix_url: str) -> dict:
    """
    Scrape a Netflix show page and return show_info dict in the same format as
    scrape_viutv_show().  Multi-season data is stored in '_seasons' and
    '_ep_by_season' for the caller to choose from.

    Returns keys:
      title, description, poster_url, banner_url, episode_count,
      episodes (empty list — caller fills from _ep_by_season),
      genres, first_air_date,
      _seasons: [(season_name, season_number), ...],
      _ep_by_season: {season_number: [episode_dicts...]}
    """
    from extractors.netflix_extractor import (
        extract_video_id, get_title_and_description, is_chinese,
        _select_first_netflix_profile, _get_all_season_numbers,
        _select_season_from_dropdown, _click_expand_buttons,
        get_episode_info, get_season_info,
    )
    from selenium.webdriver.support.ui import Select as _NfSelect
    from selenium.common.exceptions import NoSuchElementException as _NfNoElem
    from bs4 import BeautifulSoup as _NfBS

    video_id = extract_video_id(netflix_url)
    if not video_id:
        log.error(f"無法提取 Netflix 影片 ID: {netflix_url}")
        return {}

    ensure_netflix_logged_in(driver)

    standardized_url = f"https://www.netflix.com/title/{video_id}"
    log.info(f"載入 Netflix 頁面: {standardized_url}")
    driver.get(standardized_url)
    wait = WebDriverWait(driver, 20)

    try:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    except Exception as exc:
        log.error(f"Netflix 頁面加載超時: {exc}")
        return {}

    if _select_first_netflix_profile(driver):
        log.info("已選擇 Netflix 帳戶，重新載入頁面。")
        driver.get(standardized_url)
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        except Exception:
            pass

    try:
        wait.until(lambda d: any(
            d.find_elements(By.CSS_SELECTOR, sel) for sel in [
                'h3.previewModal--section-header',
                'p.preview-modal-synopsis',
                'div.episodeSelector-container',
            ]
        ))
    except Exception:
        time.sleep(3)

    soup = _NfBS(driver.page_source, 'html.parser')
    title, description = get_title_and_description(soup)
    log.info(f"Netflix 節目標題: {title}")

    has_season_dropdown = bool(soup.find('button', {'data-uia': 'dropdown-toggle'}))

    all_season_data: list[tuple[str, int]] = []
    all_ep_by_season: dict[int, list] = {}

    def _convert_ep_tuples(ep_tuples: list) -> list:
        """Convert netflix extractor tuples to standard episode dicts."""
        result = []
        for tup in ep_tuples:
            # tuple: (season_num, ep_num, runtime, title, desc, title_dup, desc_dup, backdrop_url)
            if len(tup) >= 8:
                _, ep_num, runtime, ep_title, ep_desc, _, _, image_url = tup[:8]
            elif len(tup) >= 5:
                _, ep_num, runtime, ep_title, ep_desc = tup[:5]
                image_url = tup[7] if len(tup) > 7 else ''
            else:
                continue
            result.append({
                'ep_num': ep_num,
                'title': ep_title,
                'description': ep_desc,
                'runtime': runtime,
                'image_url': image_url,
            })
        return result

    if has_season_dropdown:
        raw_seasons = _get_all_season_numbers(driver, wait)
        if not raw_seasons:
            raw_seasons = [('第 1 季', 1, '', '第 1 季', '')]
        for season_name, season_number, *_ in raw_seasons:
            all_season_data.append((season_name, season_number))
            if len(raw_seasons) > 1:
                if not _select_season_from_dropdown(driver, wait, season_number):
                    log.warning(f"無法選擇第 {season_number} 季，跳過")
                    continue
            _click_expand_buttons(driver)
            time.sleep(2)
            season_soup = _NfBS(driver.page_source, 'html.parser')
            ep_data, _ = get_episode_info(
                season_soup, standardized_url, netflix_url, None, season_number, 'n'
            )
            all_ep_by_season[season_number] = _convert_ep_tuples(ep_data)
            log.info(f"第 {season_number} 季：取得 {len(all_ep_by_season[season_number])} 集")
    else:
        try:
            season_select_el = driver.find_element(By.NAME, 'seasonSelect')
            _sel = _NfSelect(season_select_el)
            old_season_names = [opt.text.strip() for opt in _sel.options]
            old_season_values = [opt.get_attribute('value') for opt in _sel.options]
            old_season_data = get_season_info(soup, description)
            for season_name, season_number, *_ in old_season_data:
                all_season_data.append((season_name, season_number))
                if len(old_season_data) > 1:
                    sv = next(
                        (v for n, v in zip(old_season_names, old_season_values) if n == season_name),
                        None,
                    )
                    if sv:
                        _NfSelect(season_select_el).select_by_value(sv)
                        time.sleep(3)
                        soup_s = _NfBS(driver.page_source, 'html.parser')
                    else:
                        soup_s = soup
                else:
                    soup_s = soup
                ep_data, _ = get_episode_info(
                    soup_s, standardized_url, netflix_url, None, season_number, 'n'
                )
                all_ep_by_season[season_number] = _convert_ep_tuples(ep_data)
                log.info(f"第 {season_number} 季：取得 {len(all_ep_by_season[season_number])} 集")
        except _NfNoElem:
            has_new_ui = bool(soup.find('div', class_='episodeSelector-container'))
            if has_new_ui:
                _click_expand_buttons(driver)
                time.sleep(2)
                season_soup = _NfBS(driver.page_source, 'html.parser')
            else:
                season_soup = soup
            ep_data, _ = get_episode_info(
                season_soup, standardized_url, netflix_url, None, 1, 'n'
            )
            all_ep_by_season[1] = _convert_ep_tuples(ep_data)
            all_season_data.append(('第 1 季', 1))
            log.info(f"第 1 季：取得 {len(all_ep_by_season[1])} 集")

    total_eps = sum(len(v) for v in all_ep_by_season.values())
    log.info(f"Netflix 共抓取 {total_eps} 集，季數: {[s[1] for s in all_season_data]}")

    return {
        'title': title,
        'description': description,
        'poster_url': '',
        'banner_url': '',
        'episode_count': total_eps,
        'episodes': [],
        'genres': [],
        'first_air_date': '',
        '_seasons': all_season_data,
        '_ep_by_season': all_ep_by_season,
    }


# ---------------------------------------------------------------------------
# Image download
# ---------------------------------------------------------------------------

def download_image(image_url: str, dest_path: str) -> bool:
    """Download image_url to dest_path. Returns True on success."""
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        r = requests.get(image_url, headers=headers, timeout=30)
        r.raise_for_status()
        with open(dest_path, 'wb') as f:
            f.write(r.content)
        log.info(f"已下載圖片至: {dest_path}")
        return True
    except Exception as e:
        log.error(f"下載圖片失敗 ({image_url}): {e}")
        return False


# ---------------------------------------------------------------------------
# TMDB new show creation
# ---------------------------------------------------------------------------

def _wizard_click_next(driver):
    """Click the wizard Next button (li[aria-disabled=false] a[href='#next'])."""
    btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'li[aria-disabled="false"] a[href="#next"]')
        )
    )
    btn.click()
    time.sleep(1.5)


def _wizard_wait_for_step(driver, step_index: int, timeout: int = 20):
    """Block until wizard panel `step_index` (0-based) is active (aria-hidden=false)."""
    panel_id = f"new_tv_wizard-p-{step_index}"
    WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(By.ID, panel_id).get_attribute('aria-hidden') == 'false'
    )
    time.sleep(0.5)
    log.info(f"已進入 wizard 步驟 {step_index + 1}。")


def _wizard_wait_for_any_step(driver, step_indices: list, timeout: int = 30) -> int:
    """Wait until any of the given panels becomes active. Returns the active panel index."""
    def check(d):
        for idx in step_indices:
            els = d.find_elements(By.ID, f"new_tv_wizard-p-{idx}")
            if els and els[0].get_attribute('aria-hidden') == 'false':
                return idx
        return False
    result = WebDriverWait(driver, timeout).until(check)
    time.sleep(0.5)
    log.info(f"已進入 wizard 步驟 {result + 1}。")
    return result


def _fill_visible_field(driver, field_id: str, value: str):
    """JS-set value on a field that may be in a hidden panel, then fire change event."""
    driver.execute_script(
        "var el = document.getElementById(arguments[0]);"
        "if (el) { el.value = arguments[1];"
        "  el.dispatchEvent(new Event('input', {bubbles:true}));"
        "  el.dispatchEvent(new Event('change', {bubbles:true})); }",
        field_id, value
    )


def _fill_runtime_field(driver, runtime: int, language: str = 'zh-HK'):
    """
    Fill the runtime field in the Kendo grid popup, per-episode page, or wizard.

    TMDB-Import uses #{language}_runtime_text_input_field (e.g. zh-HK_runtime_...).
    Do not gate on offsetParent — Kendo Window is position:fixed, so offsetParent
    is null even when the field is on screen.
    """
    if not runtime:
        return

    ids = [
        f'{language}_runtime_text_input_field',
        f'{language.replace("-", "_")}_runtime_text_input_field',
        'runtime',
        'episode_runtime',
    ]
    done = driver.execute_script(
        """
        var ids = arguments[0], val = arguments[1], n = parseInt(val, 10);
        function setInput(el) {
            el.value = val;
            el.dispatchEvent(new Event('input',  {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
        }
        for (var i = 0; i < ids.length; i++) {
            var el = document.getElementById(ids[i]);
            if (!el) continue;
            var $el = (window.jQuery || window.$) ? (window.jQuery || window.$)(el) : null;
            var widget = $el && ($el.data('kendoNumericTextBox') || $el.data('kendoTextBox'));
            if (widget) {
                widget.value(n);
                widget.trigger('change');
                return ids[i] + ':kendo';
            }
            setInput(el);
            return ids[i] + ':input';
        }
        var el3 = document.querySelector('[data-bind*="value:runtime"]');
        if (el3) { setInput(el3); return 'bind'; }
        return false;
        """,
        ids, str(runtime),
    )

    if done:
        log.info(f"runtime 已填寫 ({done}): {runtime}")
    else:
        fields = driver.execute_script(
            "return Array.from(document.querySelectorAll('input,textarea'))"
            ".map(function(e){return (e.id||'-')+'/'+(e.name||'-')+'['+e.type+']';})"
            ".join(', ');"
        )
        log.warning(f"runtime 欄位填寫失敗 (值: {runtime}) — 頁面欄位: {fields}")


def _fill_form_field(driver, field_id: str, value: str):
    """Fill a visible text input/textarea by ID, with triple event dispatch."""
    try:
        el = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, field_id))
        )
        driver.execute_script(
            "var el = arguments[0]; el.value = arguments[1];"
            "el.dispatchEvent(new Event('input',  {bubbles:true}));"
            "el.dispatchEvent(new Event('change', {bubbles:true}));",
            el, value
        )
    except Exception as e:
        log.warning(f"填寫欄位 {field_id} 失敗: {e}")


def _fill_kendo_date_picker(driver, field_id: str, date_str: str):
    """
    Fill a Kendo DatePicker.
    Tries the jQuery Kendo API first; falls back to ActionChains + triple event dispatch
    (technique from tmdb-automation/auto-add-episodes.py).
    """
    filled = driver.execute_script(
        "var dp = $('#' + arguments[0]).data('kendoDatePicker');"
        "if (dp) { dp.value(new Date(arguments[1])); dp.trigger('change'); return true; }"
        "return false;",
        field_id, date_str
    )
    if filled:
        log.info(f"DatePicker {field_id} 已透過 Kendo API 填寫: {date_str}")
        return
    try:
        el = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, field_id))
        )
        driver.execute_script("arguments[0].value = '';", el)
        ActionChains(driver) \
            .click(el) \
            .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL) \
            .send_keys(Keys.DELETE).send_keys(date_str) \
            .perform()
        driver.execute_script(
            "var el = arguments[0];"
            "el.dispatchEvent(new Event('input',  {bubbles:true}));"
            "el.dispatchEvent(new Event('change', {bubbles:true}));"
            "el.dispatchEvent(new Event('blur',   {bubbles:true}));",
            el
        )
        log.info(f"DatePicker {field_id} 已透過 ActionChains 填寫: {date_str}")
    except Exception as e:
        log.warning(f"DatePicker 填寫失敗 ({field_id}): {e}")


def _set_kendo_episode_number(driver, target: int):
    """Set the episode-number NumericTextBox. Must target the popup field, not the pager."""
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'episode_number_numeric_text_box_field'))
    )
    done = driver.execute_script(
        """
        var el = document.getElementById('episode_number_numeric_text_box_field');
        if (!el) return false;
        var $ = window.jQuery || window.$;
        var sp = $ && $(el).data('kendoNumericTextBox');
        if (sp) { sp.value(arguments[0]); sp.trigger('change'); return 'kendo'; }
        var wrap = el.closest('.k-numerictextbox') || el.parentElement;
        var spin = wrap && wrap.querySelector('input[role="spinbutton"], input.k-input-inner');
        if (spin) {
            spin.value = arguments[0];
            spin.dispatchEvent(new Event('input',  {bubbles:true}));
            spin.dispatchEvent(new Event('change', {bubbles:true}));
            return 'spin';
        }
        el.value = arguments[0];
        el.dispatchEvent(new Event('change', {bubbles:true}));
        return 'input';
        """,
        target,
    )
    if done:
        log.debug(f"集數欄位已設為 {target} ({done})")
        return
    try:
        spin = driver.find_element(
            By.CSS_SELECTOR,
            '#episode_number_numeric_text_box_field ~ input[role="spinbutton"], '
            '.k-edit-form-container input[role="spinbutton"]',
        )
        spin.click()
        ActionChains(driver) \
            .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL) \
            .send_keys(str(target)) \
            .perform()
    except Exception as e:
        log.warning(f"設定集數 {target} 失敗: {str(e).strip() or type(e).__name__}")


def _wait_for_episode_popup(driver, timeout: int = 20) -> None:
    """Wait until the Kendo add/edit popup is ready (TMDB-Import waits on this field)."""
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, 'episode_number_numeric_text_box_field'))
    )
    time.sleep(0.4)


def _wait_for_episode_popup_closed(driver, timeout: int = 60) -> None:
    """Wait until the Update button from the popup is gone (grid finished saving)."""
    def _closed(d):
        for el in d.find_elements(By.CSS_SELECTOR, 'button[ref-update-button], button.k-grid-update'):
            cls = el.get_attribute('class') or ''
            if 'translate' in cls:
                continue
            try:
                if el.is_displayed():
                    return False
            except Exception:
                continue
        return True
    WebDriverWait(driver, timeout).until(_closed)


def _click_save_button(driver) -> bool:
    """Click the episode grid Update button (TMDB-Import: button[ref-update-button])."""
    selectors = [
        'button[ref-update-button]',
        'button.k-grid-update',
        '.k-edit-buttons button.k-primary',
        '.k-popup-edit-form .k-edit-buttons button:first-child',
    ]
    deadline = time.time() + 15
    while time.time() < deadline:
        for sel in selectors:
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                cls = el.get_attribute('class') or ''
                if 'translate' in cls:
                    continue
                try:
                    driver.execute_script("arguments[0].click();", el)
                    return True
                except Exception:
                    continue
        time.sleep(0.3)
    return False


# ---------------------------------------------------------------------------
# Broadcast schedule
# ---------------------------------------------------------------------------

def _extract_year_from_table_context(table) -> int | None:
    """
    Try to find the year associated with a wikitable from surrounding page structure.

    Handles several layouts:
    • zh-yue.wikipedia.org: <section aria-labelledby="2016年"> ancestor
    • zh.wikipedia.org collapsible format: <div class="mw-heading mw-heading3">
      <h3 id="2016年"> precedes the <div class="mw-collapsible"> that contains
      this table — found by walking up ancestors and checking their preceding siblings
    """
    # Check ancestor <section> aria-labelledby (e.g. aria-labelledby="2016年")
    for ancestor in table.parents:
        if ancestor.name == 'section':
            label = ancestor.get('aria-labelledby', '')
            m = re.search(r'(\d{4})', label)
            if m:
                return int(m.group(1))
        if ancestor.name in ('article', 'body', 'html'):
            break

    # Nearest preceding sibling of the table itself containing a 4-digit year
    for sibling in table.find_previous_siblings():
        if not hasattr(sibling, 'name'):
            continue
        text = sibling.get_text(strip=True)
        m = re.search(r'(\d{4})年', text)
        if m:
            return int(m.group(1))

    # Walk up through ancestors and check each ancestor's preceding siblings.
    # Handles the mw-collapsible format where <div class="mw-heading"><h3 id="2016年">
    # precedes the collapsible wrapper div that contains this table.
    for ancestor in table.parents:
        if ancestor.name in ('article', 'body', 'html'):
            break
        for sibling in ancestor.find_previous_siblings():
            if not hasattr(sibling, 'name'):
                continue
            text = sibling.get_text(strip=True)
            m = re.search(r'(\d{4})年', text)
            if m:
                return int(m.group(1))

    return None


def _scrape_wiki_episodes(wiki_url: str,
                          fallback_year: int | None = None,
                          section_name: str | None = None,
                          year_filter: int | None = None,
                          ) -> tuple[dict[int, str], dict[int, str], bool]:
    """
    Fetch a Wikipedia page and return (dates, titles, year_missing).

    dates:  {episode_number: 'YYYY-MM-DD'}
    titles: {episode_number: 'Episode title'} — populated when the table has a 主題/題目 column
    year_missing: True when date data was found but no year could be determined

    Supports two wikitable layouts:
    • Keyword-detected columns — header row contains 集數 / 播映 / 主題 cells.
      Handles rowspan carry and year-rollover.
    • Fallback (no column headers) — year header rows (colspan) followed by plain
      data rows; episode = col 0, date = col 1, title = col 2 (if present).

    section_name: restrict to <section aria-labelledby="..."> with that label (e.g. '第二季').
    year_filter:  restrict to tables whose context year equals this value; used for the
                  mw-collapsible year-section format where each year has its own table
                  and episode numbers restart from 1 each year.
    """
    try:
        from bs4 import BeautifulSoup as _BS
        resp = requests.get(wiki_url, timeout=15, headers={'User-Agent': 'Mozilla/5.0'})
        resp.raise_for_status()
        soup = _BS(resp.text, 'html.parser')
    except Exception as e:
        log.warning(f"無法獲取 Wiki 頁面: {e}")
        return {}, {}, False

    dates: dict[int, str] = {}
    titles: dict[int, str] = {}
    year_missing = False

    # When a specific season section is requested, restrict to its tables only.
    # This handles pages where each season has its own <section aria-labelledby="第X季">
    # and episodes restart from 1 each season — without restricting, later seasons
    # would silently overwrite earlier ones for the same episode numbers.
    if section_name:
        target_sec = soup.find('section', attrs={'aria-labelledby': section_name})
        if target_sec:
            _tables = target_sec.find_all('table', class_='wikitable')
        else:
            log.warning(f"找不到季節區塊 '{section_name}'，改用全頁資料。")
            _tables = soup.find_all('table', class_='wikitable')
    else:
        _tables = soup.find_all('table', class_='wikitable')

    for table in _tables:
        # Pre-initialise from surrounding DOM (handles per-section-per-year layout)
        current_year: int | None = _extract_year_from_table_context(table)
        if current_year is None and fallback_year is not None:
            current_year = fallback_year
        # For year-filtered mode (mw-collapsible per-year tables), skip other years
        if year_filter is not None and current_year is not None and current_year != year_filter:
            continue
        ep_col: int | None = None
        ep_col_alts: list[int] = []  # additional ep columns when multiple 集數 columns exist
        date_col: int | None = None
        title_col: int | None = None
        # rowspan carry: {logical_col_index: [text, remaining_rows]}
        rowspan_slots: dict[int, list] = {}
        prev_month: int | None = None  # for automatic year-rollover detection

        for row in table.find_all('tr'):
            raw_cells = row.find_all(['th', 'td'])
            if not raw_cells:
                continue

            # Section header — single cell spanning all columns with a year
            if len(raw_cells) == 1 and raw_cells[0].get('colspan'):
                text = raw_cells[0].get_text(strip=True)
                m = re.search(r'(\d{4})年', text)
                if m:
                    current_year = int(m.group(1))
                    rowspan_slots.clear()
                    prev_month = None  # reset rollover tracking on explicit year header
                continue

            # Column header detection — Wikipedia sometimes uses <td> with <b> instead of <th>
            # Detect by content keywords, regardless of tag name.
            # Guard: skip re-detection once both columns are established — prevents data rows
            # whose notes contain e.g. "在播出時標為" from resetting the column indices.
            cell_texts = [c.get_text(strip=True) for c in raw_cells]
            has_ep_kw = any('集數' in t or '集数' in t for t in cell_texts)
            has_date_kw = any(any(kw in t for kw in ('播映', '首播', '播出', '播放')) for t in cell_texts)
            if (has_ep_kw or has_date_kw) and (ep_col is None or date_col is None):
                ep_col = date_col = title_col = None
                ep_col_alts = []
                for i, t in enumerate(cell_texts):
                    if '集數' in t or '集数' in t:
                        if ep_col is None:
                            ep_col = i          # first 集數 column is primary
                        else:
                            ep_col_alts.append(i)  # extras tried when primary is ※
                    elif any(kw in t for kw in ('播映', '首播', '播出', '播放')):
                        date_col = i
                    elif '主題' in t or '題目' in t:
                        title_col = i
                continue
            elif all(c.name == 'th' for c in raw_cells) and len(raw_cells) > 1:
                # Sub-header row (e.g. individual host names in a multi-row header);
                # skip without resetting column indices.
                continue

            if ep_col is None or date_col is None:
                continue

            # Build logical cell map accounting for rowspan carry
            logical: dict[int, str] = {}
            for col_i, slot in list(rowspan_slots.items()):
                logical[col_i] = slot[0]
                slot[1] -= 1
                if slot[1] <= 0:
                    del rowspan_slots[col_i]

            raw_i = 0
            col_i = 0
            while raw_i < len(raw_cells):
                while col_i in logical:
                    col_i += 1
                cell = raw_cells[raw_i]
                span = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))
                # Strip <sup> footnote markers before reading text
                for sup in cell.find_all('sup'):
                    sup.decompose()
                text = cell.get_text(strip=True)
                for c in range(colspan):
                    logical[col_i + c] = text
                    if span > 1:
                        rowspan_slots[col_i + c] = [text, span - 1]
                col_i += colspan
                raw_i += 1

            ep_text = logical.get(ep_col, '')
            # Alternate ep columns: when primary is "※" or has no digits, try extras
            if ep_col_alts and (not ep_text or ep_text.strip() in ('※', '─') or not re.search(r'\d', ep_text)):
                for _ac in ep_col_alts:
                    _alt = logical.get(_ac, '')
                    if _alt and _alt.strip() not in ('※', '─') and re.search(r'\d', _alt):
                        ep_text = _alt
                        break
            date_text = logical.get(date_col, '')
            title_text = logical.get(title_col, '') if title_col is not None else ''

            # Extract year if explicitly in the date cell (e.g. "2008年12月13日").
            # This handles tables where the full date is in each cell rather than a
            # separate year context row — must be done before the current_year is None check.
            _date_year_m = re.search(r'(\d{4})年', date_text)
            if _date_year_m:
                current_year = int(_date_year_m.group(1))
                prev_month = None  # explicit year resets month-rollover tracking

            if current_year is None:
                year_missing = True
                continue

            # Parse episode number(s) — strip footnote brackets, handle "1", "1-2", "1,2"
            ep_clean = re.sub(r'\[.*?\]|（.*?）|【.*?】', '', ep_text)
            ep_nums: list[int] = []
            for p in re.split(r'[,，、\-–—至]', ep_clean):
                p = p.strip()
                if p.isdigit():
                    ep_nums.append(int(p))
            if not ep_nums:
                # Last resort: grab the first number in the cell
                m0 = re.match(r'\d+', ep_clean.strip())
                if m0:
                    ep_nums.append(int(m0.group()))
            if not ep_nums:
                continue

            # Parse date — "4月14號", "4月14日", "2008年12月13日"
            dm = re.search(r'(\d+)月(\d+)[號日]?', date_text)
            if not dm:
                continue

            month = int(dm.group(1))
            day = int(dm.group(2))
            # Year rollover only when year comes from context, not explicit in date cell
            if _date_year_m is None and prev_month is not None and month < prev_month:
                current_year += 1
            prev_month = month

            try:
                from datetime import date as _date
                d = _date(current_year, month, day)
                date_str = d.isoformat()
            except ValueError:
                continue

            for ep_num in ep_nums:
                dates[ep_num] = date_str
                if title_text and ep_num not in titles:
                    titles[ep_num] = title_text

        # ── Fallback: year-section header rows but no 集數/播映 column headers ──
        # Format: <tr><th colspan="N">YYYY年</th></tr>  followed by
        #         <tr><td>01</td><td>4月3日</td><td>Title</td><td>…</td></tr>
        if ep_col is None and date_col is None:
            _nf_year: int | None = None
            _nf_prev_month: int | None = None
            for _nf_row in table.find_all('tr'):
                _nf_cells = _nf_row.find_all(['th', 'td'])
                if not _nf_cells:
                    continue
                # Year header row (single cell with colspan containing YYYY年)
                if len(_nf_cells) == 1 and _nf_cells[0].get('colspan'):
                    _nf_text = _nf_cells[0].get_text(strip=True)
                    _nf_ym = re.search(r'(\d{4})年', _nf_text)
                    if _nf_ym:
                        _nf_year = int(_nf_ym.group(1))
                        _nf_prev_month = None
                    continue
                # Skip all-th header rows (column labels)
                if all(c.name == 'th' for c in _nf_cells):
                    continue
                if _nf_year is None or len(_nf_cells) < 2:
                    continue
                # Strip <sup> footnote markers before reading text
                for _nf_cell in _nf_cells:
                    for _sup in _nf_cell.find_all('sup'):
                        _sup.decompose()
                _nf_ep_text = _nf_cells[0].get_text(strip=True)
                _nf_date_text = _nf_cells[1].get_text(strip=True)
                _nf_title_text = _nf_cells[2].get_text(strip=True) if len(_nf_cells) > 2 else ''
                # Episode number — must be a plain integer (strip leading zeros)
                _nf_ep_m = re.match(r'^(\d+)$', _nf_ep_text)
                if not _nf_ep_m:
                    continue
                _nf_ep_num = int(_nf_ep_m.group(1))
                # Date — "4月3日" / "4月3號"
                _nf_dm = re.search(r'(\d+)月(\d+)[號日]?', _nf_date_text)
                if not _nf_dm:
                    continue
                _nf_month = int(_nf_dm.group(1))
                _nf_day = int(_nf_dm.group(2))
                # Auto-advance year on month rollover (Dec→Jan)
                if _nf_prev_month is not None and _nf_month < _nf_prev_month:
                    _nf_year += 1
                _nf_prev_month = _nf_month
                try:
                    from datetime import date as _date
                    _nf_d = _date(_nf_year, _nf_month, _nf_day)
                    dates[_nf_ep_num] = _nf_d.isoformat()
                    if _nf_title_text and _nf_ep_num not in titles:
                        titles[_nf_ep_num] = _nf_title_text
                except ValueError:
                    continue

    log.info(f"Wiki 播映日期: 共 {len(dates)} 集，標題: 共 {len(titles)} 集 ({wiki_url})")
    return dates, titles, year_missing


def _scrape_wiki_air_dates(wiki_url: str,
                           fallback_year: int | None = None,
                           section_name: str | None = None,
                           year_filter: int | None = None) -> tuple[dict[int, str], bool]:
    """Backward-compatible wrapper around _scrape_wiki_episodes; returns only dates."""
    dates, _, year_missing = _scrape_wiki_episodes(
        wiki_url, fallback_year, section_name=section_name, year_filter=year_filter
    )
    return dates, year_missing


def _find_wiki_year_sections(soup) -> list[int]:
    """
    Return sorted list of years from mw-heading year-section divs that have an
    associated wikitable. Detects the zh.wikipedia.org collapsible format:
      <div class="mw-heading mw-heading3"><h3 id="2016年">…</h3></div>
      <div class="mw-collapsible …"><div class="mw-collapsible-content">
        <table class="wikitable">…
    """
    years: list[int] = []
    for div in soup.find_all('div', class_='mw-heading'):
        h3 = div.find('h3')
        if h3 is None:
            continue
        m = re.match(r'^(\d{4})年', h3.get('id', ''))
        if not m:
            continue
        year = int(m.group(1))
        nxt = div.find_next_sibling()
        if nxt and hasattr(nxt, 'find') and nxt.find('table', class_='wikitable'):
            if year not in years:
                years.append(year)
    return sorted(years)


def _fetch_wiki_dates(wiki_url: str) -> dict[int, str]:
    """
    Scrape Wikipedia air dates.
    • If the page has multiple season sections (第X季), prompt the user to choose one.
    • If the page has year-based collapsible sections (2016年, 2017年 …), prompt to choose.
    • If the page has no year information, prompt for a fallback year and re-scrape.
    """
    from bs4 import BeautifulSoup as _BS

    # Pre-fetch once to detect multi-section layouts.
    section_name: str | None = None
    year_filter: int | None = None
    try:
        _resp = requests.get(wiki_url, timeout=15, headers={'User-Agent': 'Mozilla/5.0'})
        _resp.raise_for_status()
        _soup = _BS(_resp.text, 'html.parser')
        season_secs = [
            sec['aria-labelledby']
            for sec in _soup.find_all('section', attrs={'aria-labelledby': True})
            if '季' in sec.get('aria-labelledby', '')
            and sec.find('table', class_='wikitable')
        ]
        year_secs = _find_wiki_year_sections(_soup)
    except Exception:
        season_secs = []
        year_secs = []

    if len(season_secs) > 1:
        # Format A: <section aria-labelledby="第X季"> — choose a season
        print(f"\n此 Wikipedia 頁面包含 {len(season_secs)} 個季節的資料:")
        for i, name in enumerate(season_secs, 1):
            print(f"  {i}. {name}")
        raw = input(f"請選擇使用哪一季? (1–{len(season_secs)}，默認 1): ").strip()
        idx = int(raw) - 1 if raw.isdigit() and 1 <= int(raw) <= len(season_secs) else 0
        section_name = season_secs[idx]
        print(f"已選擇: {section_name}")
    elif len(year_secs) > 1:
        # Format B: mw-collapsible year sections (e.g. 2016年, 2017年) — choose a year
        print(f"\n此 Wikipedia 頁面按年份分組，共有 {len(year_secs)} 個年份的資料:")
        for i, yr in enumerate(year_secs, 1):
            print(f"  {i}. {yr}年")
        raw = input(f"請選擇使用哪一年的資料? (1–{len(year_secs)}，默認 1): ").strip()
        idx = int(raw) - 1 if raw.isdigit() and 1 <= int(raw) <= len(year_secs) else 0
        year_filter = year_secs[idx]
        print(f"已選擇: {year_filter}年")
    elif len(year_secs) == 1:
        # Only one year-section: use it automatically without prompting
        year_filter = year_secs[0]

    dates, year_missing = _scrape_wiki_air_dates(
        wiki_url, section_name=section_name, year_filter=year_filter
    )
    if year_missing:
        print("Wiki 頁面沒有年份資料。")
        while True:
            raw = input("請輸入節目播映年份 (例如 2025): ").strip()
            if raw.isdigit() and len(raw) == 4:
                dates, _ = _scrape_wiki_air_dates(
                    wiki_url, fallback_year=int(raw),
                    section_name=section_name, year_filter=year_filter
                )
                break
            print("請輸入有效的四位年份。")
    return dates


_DATE_TITLE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')


def _group_episodes_by_year(episodes: list) -> dict[int, list]:
    """
    Group episodes by the year in their YYYY-MM-DD title.
    Episodes whose title doesn't match the date pattern are dropped.
    Returns {year: [ep, ...]} sorted ascending by year.
    """
    groups: dict[int, list] = {}
    for ep in episodes:
        title = (ep.get('title') or '').strip()
        m = re.match(r'^(\d{4})-\d{2}-\d{2}$', title)
        if m:
            groups.setdefault(int(m.group(1)), []).append(ep)
    return dict(sorted(groups.items()))


def _dates_from_titles(episodes: list) -> dict[int, str] | None:
    """
    If ≥80% of episode titles look like 'YYYY-MM-DD', treat the title as
    the air date and return {ep_num: title}.  Returns None otherwise.
    """
    if not episodes:
        return None
    dated = [e for e in episodes if _DATE_TITLE_RE.match((e.get('title') or '').strip())]
    if len(dated) / len(episodes) < 0.8:
        return None
    return {e['ep_num']: e['title'].strip() for e in dated}


def _ask_air_dates(episodes: list, default_start: str = '',
                   cached_wiki_dates: dict | None = None) -> dict[int, str]:
    """
    Return {ep_num: 'YYYY-MM-DD'}.
    Priority: date-format titles → pre-fetched wiki_dates → new Wikipedia URL → manual schedule.
    """
    ep_nums = {e['ep_num'] for e in episodes}

    # Episode titles ARE the air dates (e.g. "2025-11-12")
    title_dates = _dates_from_titles(episodes)
    if title_dates is not None:
        print(f"偵測到集數標題為日期格式，自動使用標題作為播映日期 ({len(title_dates)}/{len(episodes)} 集)。")
        return title_dates

    # Use pre-fetched dates from an earlier call (mode 3 reuses across seasons)
    if cached_wiki_dates is not None:
        filtered = {k: v for k, v in cached_wiki_dates.items() if k in ep_nums}
        if filtered:
            print(f"使用 Wikipedia 播映日期 ({len(filtered)}/{len(ep_nums)} 集)。")
            return filtered

    print("\n=== 設定播出日期 ===")
    wiki_url = input("Wikipedia 頁面 URL (留空手動設定): ").strip()
    if wiki_url:
        dates = _fetch_wiki_dates(wiki_url)
        filtered = {k: v for k, v in dates.items() if k in ep_nums}
        if filtered:
            print(f"從 Wikipedia 成功抓取 {len(filtered)}/{len(ep_nums)} 集播映日期。")
            return filtered
        print("未能從 Wikipedia 抓取到有效播映日期，改為手動設定。")

    return _ask_broadcast_schedule(episodes, default_start)


def _ask_broadcast_schedule(episodes: list, default_start: str = '') -> dict:
    """
    Prompt the user for a start date and broadcast pattern.
    Returns {ep_num: 'YYYY-MM-DD'} for all episodes with a computed air date.
    """
    print("\n=== 設定播出日期 ===")
    prompt = f"第 1 集播出日期 (YYYY-MM-DD{f'，默認 {default_start}' if default_start else ''}，留空跳過): "
    start_raw = input(prompt).strip() or default_start
    if not start_raw:
        return {}
    try:
        start = _date.fromisoformat(start_raw[:10])
    except ValueError:
        print("日期格式錯誤，跳過播出日期設定。")
        return {}

    print("\n播放頻率:")
    print("  1. 逢星期一至五 (Mon–Fri，每集一天)")
    print("  2. 逢星期一至六 (Mon–Sat，每集一天)")
    print("  3. 每週一集 → 選擇星期幾")
    print("  4. 每週多集 → 輸入播出日 (例如 1,3,5 = 週一三五)")
    print("  5. 每日一集")
    print("  6. 不填 / 稍後手動設定")
    freq = input("選擇 (1-6，默認 6): ").strip() or '6'

    if freq == '1':
        weekdays = {0, 1, 2, 3, 4}
    elif freq == '2':
        weekdays = {0, 1, 2, 3, 4, 5}
    elif freq == '3':
        d = input("播出星期幾? (1=週一 … 7=週日): ").strip()
        weekdays = {int(d) - 1} if d.isdigit() and 1 <= int(d) <= 7 else {start.weekday()}
    elif freq == '4':
        raw = input("播出星期 (例如 1,3,5 表示週一三五): ").strip()
        weekdays = {int(p) - 1 for p in raw.split(',') if p.strip().isdigit() and 1 <= int(p.strip()) <= 7}
        if not weekdays:
            weekdays = {start.weekday()}
    elif freq == '5':
        weekdays = set(range(7))
    else:
        return {}

    ep_nums = sorted(e['ep_num'] for e in episodes)
    air_dates: dict = {}
    current = start
    while current.weekday() not in weekdays:
        current += timedelta(days=1)
    for ep_num in ep_nums:
        air_dates[ep_num] = current.isoformat()
        current += timedelta(days=1)
        while current.weekday() not in weekdays:
            current += timedelta(days=1)

    log.info(f"已計算 {len(air_dates)} 集播出日期，首集: {air_dates.get(1, '–')}")
    return air_dates


# ---------------------------------------------------------------------------
# Season episode batch-add
# ---------------------------------------------------------------------------

def _read_grid_episodes(driver) -> dict[int, tuple[str, str, str]]:
    """
    Read all existing episodes from the Kendo grid, navigating all pages.
    Returns {episode_number: (name_text, overview_text, runtime_text)}.

    Uses JavaScript to read rows so the result is independent of Kendo's
    row CSS class (k-master-row, k-table-row, etc. differ by version).
    Iterates the Kendo pager to collect all pages.
    """
    result: dict[int, tuple[str, str, str]] = {}

    # JS that reads all data rows from the grid content area.
    # Episode numbers may be formatted with thousands separators (e.g. "1,360"),
    # so commas are stripped before parseInt.
    _JS_READ = """
        var out = [];
        var c = document.querySelector('.k-grid-content') ||
                document.querySelector('.k-grid-table-wrap') ||
                document.querySelector('.k-grid');
        if (!c) return out;
        c.querySelectorAll('tr').forEach(function(row) {
            var cells = row.querySelectorAll('td');
            if (cells.length < 3) return;
            var t = cells[0].textContent.trim().replace(/,/g, '');
            var ep = parseInt(t, 10);
            if (isNaN(ep) || String(ep) !== t) return;
            out.push([ep,
                cells[1].textContent.trim(),
                cells[2].textContent.trim(),
                cells.length > 4 ? cells[4].textContent.trim() : '']);
        });
        return out;
    """

    def _read_current_page() -> int:
        try:
            rows_data = driver.execute_script(_JS_READ) or []
        except Exception as exc:
            log.debug(f"JS grid read error: {exc}")
            rows_data = []

        if not rows_data:
            # Selenium fallback for environments where JS return is restricted
            for sel in ['tr[data-uid]', 'tr.k-master-row', 'tr.k-table-row']:
                rows = driver.find_elements(By.CSS_SELECTOR, sel)
                if rows:
                    for row in rows:
                        tds = row.find_elements(By.CSS_SELECTOR, 'td')
                        if len(tds) < 3:
                            continue
                        try:
                            # Strip thousands-separator commas (e.g. "1,360" → 1360)
                            ep_n = int(tds[0].text.strip().replace(',', ''))
                        except ValueError:
                            continue
                        name_text = (tds[1].get_attribute('textContent') or '').strip()
                        overview_text = (tds[2].get_attribute('textContent') or '').strip()
                        runtime_text = (tds[4].get_attribute('textContent') or '').strip() if len(tds) > 4 else ''
                        result[ep_n] = (name_text, overview_text, runtime_text)
                    return len(rows)
            return 0

        for rd in rows_data:
            result[rd[0]] = (rd[1], rd[2], rd[3] if len(rd) > 3 else '')
        return len(rows_data)

    # Wait up to 12 s for grid data rows to appear before starting to read.
    # (The .k-grid container appears before the AJAX data load completes.)
    for _w in range(12):
        try:
            has_rows = driver.execute_script("""
                var c = document.querySelector('.k-grid-content, .k-grid-table-wrap, .k-grid');
                if (!c) return false;
                var rows = c.querySelectorAll('tr');
                for (var i = 0; i < rows.length; i++) {
                    if (rows[i].querySelectorAll('td').length >= 3) return true;
                }
                return false;
            """)
            if has_rows:
                break
        except Exception:
            pass
        # Also stop waiting if the grid shows an explicit "no records" message
        if driver.find_elements(By.CSS_SELECTOR, '.k-grid-norecords, .k-nodata, [class*="norecords"]'):
            break
        time.sleep(1)
    else:
        # Log what the DOM actually contains to help diagnose selector mismatches
        try:
            dbg = driver.execute_script("""
                var c = document.querySelector('.k-grid-content, .k-grid-table-wrap, .k-grid');
                if (!c) return 'grid container not found';
                var trs = c.querySelectorAll('tr');
                return 'grid found; tr count=' + trs.length + '; first tr classes=' +
                       (trs[0] ? trs[0].className : 'n/a');
            """)
            log.warning(f"Grid wait exhausted — DOM state: {dbg}")
        except Exception:
            pass

    _MAX_PAGES = 200  # safety cap
    for _page in range(_MAX_PAGES):
        _read_current_page()

        # Locate the Kendo "next page" button
        next_btn = None
        for sel in [
            'a.k-pager-nav.k-pager-next',
            '.k-grid-pager a.k-pager-next',
            'a[aria-label*="next page" i]',
            'button[aria-label*="next page" i]',
            '.k-pager-next',
        ]:
            btns = driver.find_elements(By.CSS_SELECTOR, sel)
            if btns:
                next_btn = btns[0]
                break

        if next_btn is None:
            break  # No pager → single-page grid

        classes = next_btn.get_attribute('class') or ''
        aria_disabled = next_btn.get_attribute('aria-disabled') or ''
        if 'k-disabled' in classes or aria_disabled == 'true':
            break  # Last page reached

        driver.execute_script("arguments[0].click();", next_btn)
        time.sleep(1.5)

    return result


def _is_placeholder_name(name: str, ep_num: int) -> bool:
    """Return True if name is the TMDB wizard auto-generated placeholder '第 N 集'."""
    return bool(re.match(r'^第\s*\d+\s*集$', name.strip()))


def _find_grid_row(driver, ep_num: int):
    """Find the grid row for ep_num, navigating pages if needed. Returns the row element or None."""
    ep_str = str(ep_num)
    _MAX_PAGES = 200
    for _ in range(_MAX_PAGES):
        # Try data-uid rows first (Kendo-specific), then class-based selectors
        rows = []
        for sel in ['tr[data-uid]', 'tr.k-master-row', 'tr.k-table-row']:
            rows = driver.find_elements(By.CSS_SELECTOR, sel)
            if rows:
                break
        if not rows:
            # Last resort: any tr inside the grid content area
            rows = driver.find_elements(By.CSS_SELECTOR, '.k-grid-content tr, .k-grid-table-wrap tr')
        for row in rows:
            tds = row.find_elements(By.CSS_SELECTOR, 'td')
            # Strip thousands-separator commas before comparing (e.g. "1,360" → "1360")
            if tds and tds[0].text.strip().replace(',', '') == ep_str:
                return row

        next_btn = None
        for sel in ['a.k-pager-nav.k-pager-next', '.k-grid-pager a.k-pager-next',
                    'a[aria-label*="next page" i]', 'button[aria-label*="next page" i]',
                    '.k-pager-next']:
            btns = driver.find_elements(By.CSS_SELECTOR, sel)
            if btns:
                next_btn = btns[0]
                break
        if next_btn is None:
            break
        classes = next_btn.get_attribute('class') or ''
        aria_disabled = next_btn.get_attribute('aria-disabled') or ''
        if 'k-disabled' in classes or aria_disabled == 'true':
            break
        driver.execute_script("arguments[0].click();", next_btn)
        time.sleep(1.5)
    return None


def _edit_episode_in_grid(driver, ep_num: int, ep: dict, air_dates: dict, language: str) -> bool:
    """Click Edit on an existing episode row and fill in its data."""
    try:
        target_row = _find_grid_row(driver, ep_num)
        if not target_row:
            log.warning(f"找不到第 {ep_num} 集的表格行。")
            return False

        edit_btn = None
        for sel in ['button.k-grid-edit-command', 'button.k-grid-edit', 'a.k-button']:
            btns = target_row.find_elements(By.CSS_SELECTOR, sel)
            if btns:
                edit_btn = btns[0]
                break
        if not edit_btn:
            log.warning(f"找不到第 {ep_num} 集的編輯按鈕。")
            return False

        driver.execute_script("arguments[0].scrollIntoView(true);", edit_btn)
        driver.execute_script("arguments[0].click();", edit_btn)
        time.sleep(1)

        _wait_for_episode_popup(driver)
        _fill_form_field(driver, f'{language}_name_text_input_field', add_spaces_between_en_zh_num(ep.get('title', '')))
        _fill_form_field(driver, f'{language}_overview_text_box_field', add_spaces_between_en_zh_num(ep.get('description', '')))
        if ep_num in air_dates:
            _fill_kendo_date_picker(driver, 'air_date_date_picker_field', air_dates[ep_num])
        _fill_runtime_field(driver, ep.get('runtime', 0), language)

        if _click_save_button(driver):
            try:
                _wait_for_episode_popup_closed(driver)
            except Exception:
                time.sleep(1.5)
            log.info(f"第 {ep_num} 集已補填資料。")
            return True
        log.warning(f"第 {ep_num} 集補填儲存失敗。")
        return False
    except Exception as e:
        log.error(f"編輯第 {ep_num} 集時發生錯誤: {str(e).strip() or type(e).__name__}")
        return False


def add_episodes_to_season(
    driver,
    show_id: str,
    season: int,
    episodes: list,
    air_dates: dict,
    language: str = 'zh-HK',
) -> tuple[set[int], set[int]]:
    """
    Add all episodes to the TMDB season edit page via the Kendo grid Add button.
    - Episodes that already exist with data → skipped (safe to resume after interruption).
    - Episodes that exist but are empty (e.g. a placeholder ep 1) → edited via Edit button.
    - Episodes that don't exist yet → added via Add button.
    Returns (newly_added, edited) where:
      newly_added = episode numbers that did not exist in the grid before
      edited      = episode numbers that existed as empty placeholders and were filled in
    """
    season_url = (
        f"https://www.themoviedb.org/tv/{show_id}/season/{season}/edit"
        f"?active_nav_item=episodes"
    )
    log.info(f"前往季數編輯頁面: {season_url}")
    driver.get(season_url)
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'button.k-grid-add, .k-grid'))
        )
    except Exception:
        log.warning("季數編輯頁面載入超時，繼續嘗試。")
    time.sleep(2)

    # Read all existing episodes from the grid (episode_number → (name, overview))
    grid_eps = _read_grid_episodes(driver)
    log.info(f"TMDB grid 已偵測到集數: {sorted(grid_eps.keys())}")

    # Build a lookup: ep_num → runtime we intend to write (0 if none)
    ep_runtime_map = {ep['ep_num']: ep.get('runtime', 0) or 0 for ep in episodes}

    # Split into: all fields present → skip; any field missing → edit
    # Whether viu.tv itself has a description for each episode
    ep_has_desc = {ep['ep_num']: bool((ep.get('description') or '').strip()) for ep in episodes}

    existing_with_name: set[int] = set()
    existing_empty: set[int] = set()
    for ep_n, (name, overview, runtime_grid) in grid_eps.items():
        name_ok = bool(name) and not _is_placeholder_name(name, ep_n)
        intended_runtime = ep_runtime_map.get(ep_n, 0)
        ep_has_runtime = intended_runtime > 0
        # Parse the grid runtime as an integer (avoids bool('0') == True trap)
        try:
            runtime_grid_int = int(runtime_grid.strip()) if runtime_grid and runtime_grid.strip() else 0
        except (ValueError, TypeError):
            runtime_grid_int = 0
        # Runtime is OK only if it already matches the intended value
        runtime_ok = (runtime_grid_int == intended_runtime) if ep_has_runtime else True
        # Overview is only required when viu.tv itself has a description to fill
        overview_ok = bool(overview) if ep_has_desc.get(ep_n, True) else True
        if name_ok and overview_ok and runtime_ok:
            existing_with_name.add(ep_n)
        else:
            existing_empty.add(ep_n)

    existing_eps = existing_with_name | existing_empty
    log.info(f"TMDB 已存在集數 (標題+描述+片長均有): {sorted(existing_with_name)}")
    if existing_empty:
        log.info(f"TMDB 已存在集數 (有欄位缺失，將補填): {sorted(existing_empty)}")

    to_edit = [ep for ep in sorted(episodes, key=lambda e: e['ep_num'])
               if ep['ep_num'] in existing_empty]
    to_add = [ep for ep in sorted(episodes, key=lambda e: e['ep_num'])
              if ep['ep_num'] not in existing_eps]

    if not to_add and not to_edit:
        log.info("所有集數已存在且有資料，無需新增或補填。")
        return set(), set()

    # Edit empty existing episodes first
    if to_edit:
        log.info(f"補填空集數: {[ep['ep_num'] for ep in to_edit]}")
        for ep in to_edit:
            _edit_episode_in_grid(driver, ep['ep_num'], ep, air_dates, language)

    if not to_add:
        return set(), {ep['ep_num'] for ep in to_edit}
    log.info(f"需要新增集數: {[ep['ep_num'] for ep in to_add]}")

    success, failed = 0, []
    for ep in to_add:
        ep_num = ep['ep_num']
        try:
            add_btn = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'button.k-grid-add'))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_btn)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", add_btn)
            _wait_for_episode_popup(driver)

            _set_kendo_episode_number(driver, ep_num)
            _fill_form_field(driver, f'{language}_name_text_input_field', add_spaces_between_en_zh_num(ep.get('title', '')))
            _fill_form_field(driver, f'{language}_overview_text_box_field', add_spaces_between_en_zh_num(ep.get('description', '')))
            if ep_num in air_dates:
                _fill_kendo_date_picker(driver, 'air_date_date_picker_field', air_dates[ep_num])
            _fill_runtime_field(driver, ep.get('runtime', 0), language)

            if _click_save_button(driver):
                try:
                    _wait_for_episode_popup_closed(driver)
                except Exception:
                    time.sleep(1.5)
                log.info(f"第 {ep_num} 集已新增。")
                success += 1
            else:
                log.warning(f"第 {ep_num} 集儲存失敗。")
                failed.append(ep_num)
                time.sleep(1)

        except Exception as e:
            log.error(f"新增第 {ep_num} 集時發生錯誤: {str(e).strip() or type(e).__name__}")
            failed.append(ep_num)
            try:
                cancel = driver.find_element(By.CSS_SELECTOR, '.k-grid-cancel, button[ref-cancel-button]')
                driver.execute_script("arguments[0].click();", cancel)
            except Exception:
                pass
            time.sleep(1)

    log.info(f"集數新增完成: {success} 成功, {len(failed)} 失敗。")
    if failed:
        log.warning(f"失敗集數: {failed}")
    added_nums = {ep['ep_num'] for ep in to_add if ep['ep_num'] not in failed}
    edited_nums = {ep['ep_num'] for ep in to_edit}
    return added_nums, edited_nums


def create_tmdb_show(driver, show_info: dict) -> str | None:
    """
    Fill the 7-step TMDB /tv/new wizard and return the new show ID string.

    Wizard steps (0-based panel index):
      0 — Getting Started         → click Next
      1 — TV Show Details         → original_name, series_overview, [imdb_id]
      2 — Translation Check       → auto; click Next when ready
      3 — Duplicate Check         → ensure duplicate_verified checked; click Next
      4 — TV Season Details       → season_name, season_number (keep defaults)
      5 — TV Episode Details      → episode_name, episode_overview, episode_air_date
      6 — Verify & Save           → click Finish → redirect → extract show ID
    """
    log.info("前往 TMDB 新增電視劇頁面: https://www.themoviedb.org/tv/new")
    driver.get('https://www.themoviedb.org/tv/new')

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, 'new_tv_wizard'))
        )
    except TimeoutException:
        log.error("TMDB 新增頁面未能載入 wizard。")
        return None

    try:
        # ── Step 0: Getting Started ──────────────────────────────────────────
        _wizard_wait_for_step(driver, 0)
        log.info("Step 1/7: Getting Started → 點擊 Next")
        _wizard_click_next(driver)

        # ── Step 1: TV Show Details ──────────────────────────────────────────
        _wizard_wait_for_step(driver, 1)
        log.info("Step 2/7: TV Show Details")

        show_title = add_spaces_between_en_zh_num(show_info['title'])
        _fill_visible_field(driver, 'original_name', show_title)
        log.info(f"已填寫 original_name: {show_title}")

        overview = add_spaces_between_en_zh_num(show_info.get('description', '').strip())
        if len(overview) < 10:
            overview = (overview + ' ' * 10)[:10]  # pad to minimum 10 chars
        _fill_visible_field(driver, 'series_overview', overview)
        log.info(f"已填寫 series_overview ({len(overview)} 字)。")

        _wizard_click_next(driver)

        # After clicking Next on TV Show Details, TMDB may auto-process
        # Translation Check (p-2) and Duplicate Check (p-3) in the background
        # and jump directly to TV Season Details (p-4). Handle all cases.
        active = _wizard_wait_for_any_step(driver, [2, 3, 4])

        if active == 2:
            # ── Step 2: Translation Check ────────────────────────────────────
            log.info("Step 3/7: Translation Check → 等待自動檢測...")
            try:
                WebDriverWait(driver, 30).until(
                    lambda d: (
                        not d.find_element(By.CSS_SELECTOR, '#new_tv_wizard-p-2 img.loading').is_displayed()
                        if d.find_elements(By.CSS_SELECTOR, '#new_tv_wizard-p-2 img.loading')
                        else True
                    )
                )
            except TimeoutException:
                log.warning("Translation Check 等待超時，繼續。")
            time.sleep(1)
            driver.execute_script(
                "var cb = document.getElementById('translation_verified');"
                "if (cb) cb.checked = true;"
            )
            _wizard_click_next(driver)
            active = _wizard_wait_for_any_step(driver, [3, 4])

        if active == 3:
            # ── Step 3: Duplicate Check ──────────────────────────────────────
            log.info("Step 4/7: Duplicate Check → 確認無重複並繼續")
            driver.execute_script(
                "var cb = document.getElementById('duplicate_verified');"
                "if (cb) cb.checked = true;"
            )
            time.sleep(1)
            _wizard_click_next(driver)
            _wizard_wait_for_step(driver, 4)
        else:
            log.info("Translation Check / Duplicate Check 已自動通過，直接進入 TV Season Details。")

        # ── Step 4: TV Season Details ────────────────────────────────────────
        log.info("Step 5/7: TV Season Details → 清空 Season name 令 TMDB 自動套用本地語言")
        # Clear the season name so TMDB falls back to its localised default
        driver.execute_script(
            "var el = document.getElementById('season_name');"
            "if (el) { el.value = ''; el.dispatchEvent(new Event('input', {bubbles:true}));"
            "          el.dispatchEvent(new Event('change', {bubbles:true})); }"
        )
        _wizard_click_next(driver)

        # ── Step 5: TV Episode Details ───────────────────────────────────────
        _wizard_wait_for_step(driver, 5)
        log.info("Step 6/7: TV Episode Details")

        # Use the first (lowest ep_num) episode scraped from viu.tv
        episodes_sorted = sorted(show_info.get('episodes', []), key=lambda e: e['ep_num'])
        ep1 = episodes_sorted[0] if episodes_sorted else None
        ep1_num  = ep1['ep_num'] if ep1 else 1
        ep1_name = add_spaces_between_en_zh_num(ep1['title'] if ep1 and ep1['title'] else f"{show_info['title']} 第 {ep1_num} 集")
        ep1_desc  = add_spaces_between_en_zh_num(ep1['description'] if ep1 and ep1['description'] else overview)
        if len(ep1_desc) < 10:
            ep1_desc = overview  # fallback to show overview if episode description too short

        _fill_visible_field(driver, 'episode_name', ep1_name)
        log.info(f"已填寫 episode_name: {ep1_name}")

        _fill_visible_field(driver, 'episode_overview', ep1_desc)
        log.info(f"已填寫 episode_overview ({len(ep1_desc)} 字)。")

        # Set the episode number (TMDB wizard defaults to 1)
        driver.execute_script(
            "var el = document.getElementById('episode_number');"
            "if (el) { el.value = arguments[0];"
            "  el.dispatchEvent(new Event('input',  {bubbles:true}));"
            "  el.dispatchEvent(new Event('change', {bubbles:true})); }",
            str(ep1_num)
        )
        log.info(f"已填寫 episode_number: {ep1_num}")

        ep1_runtime = ep1.get('runtime', 0) if ep1 else 0
        if ep1_runtime:
            _fill_runtime_field(driver, ep1_runtime)  # may silently fail if wizard has no runtime field

        air_date = show_info.get('first_air_date', '').strip()
        if not air_date:
            print("\n未能自動取得第 1 集播出日期。")
            air_date = input("請輸入第 1 集播出日期 (格式 YYYY-MM-DD，例如 2024-01-15): ").strip()

        if air_date:
            _fill_kendo_date_picker(driver, 'episode_air_date', air_date[:10])
        else:
            log.warning("未填寫 episode_air_date，wizard 可能拒絕進入下一步。")

        _wizard_click_next(driver)

        # ── Step 6: Verify & Save ────────────────────────────────────────────
        # If validation failed (e.g. air date rejected) the wizard stays on step 5
        # with class="current error". Detect this and let the user fix it manually.
        try:
            _wizard_wait_for_step(driver, 6, timeout=10)
        except TimeoutException:
            els = driver.find_elements(By.ID, 'new_tv_wizard-p-5')
            if els and els[0].get_attribute('aria-hidden') == 'false':
                print("\nWizard 仍停留在 TV Episode Details（可能有欄位錯誤）。")
                print("請在瀏覽器中補填或更正欄位，然後手動點擊 Next。")
                input("完成後按 Enter，程式將繼續等待 Verify & Save 步驟...")
                _wizard_wait_for_step(driver, 6, timeout=120)
            else:
                raise
        log.info("Step 7/7: Verify & Save → 點擊 Finish")

        finish_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="#finish"]'))
        )
        finish_btn.click()
        log.info("已點擊 Finish，等待跳轉...")
        time.sleep(5)

        # ── Extract show ID from redirect URL ───────────────────────────────
        current_url = driver.current_url
        log.info(f"提交後 URL: {current_url}")
        show_id = _extract_show_id_from_url(current_url)
        if show_id:
            log.info(f"新節目 TMDB ID: {show_id}")
            return show_id

        # Fallback: give user time to handle any prompt (e.g. CAPTCHA)
        print(f"\n當前 URL: {current_url}")
        _pause_for_manual("如需手動完成請操作後按 Enter，程式將從當前頁面提取節目 ID...")
        current_url = driver.current_url
        show_id = _extract_show_id_from_url(current_url)
        if show_id:
            log.info(f"新節目 TMDB ID: {show_id}")
            return show_id

        log.error(f"無法從 URL 提取節目 ID: {current_url}")
        return None

    except Exception as e:
        log.error(f"建立節目時發生錯誤: {e}", exc_info=True)
        return None


def _extract_show_id_from_url(url: str) -> str | None:
    """Extract TMDB show ID from a URL like /tv/12345 or /tv/12345/edit."""
    m = re.search(r'themoviedb\.org/tv/(\d+)', url)
    return m.group(1) if m else None


def _pause_for_manual(prompt: str):
    input(prompt)


# ---------------------------------------------------------------------------
# TMDB image upload
# ---------------------------------------------------------------------------

def _has_existing_images(driver) -> bool:
    """Return True if the current TMDB image page already contains images."""
    try:
        sys.path.insert(0, script_dir)
        import viutv_image_uploader as _img
        return _img._has_existing_images(driver)
    except Exception:
        try:
            WebDriverWait(driver, 8).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'div.image_content, li#no_results, a.add_image.upload')
                )
            )
        except Exception:
            pass
        return len(driver.find_elements(By.CSS_SELECTOR, 'div.image_content')) > 0


def _upload_to_tmdb_image_page(
    driver,
    page_url: str,
    image_path: str,
    label: str,
    skip_if_exists: bool = True,
    upload_all_state: list | None = None,
) -> bool:
    """Generic TMDB image-page uploader (TMDB-Import cropper flow). Skips if images already exist."""
    log.info(f"前往 TMDB 圖片頁面: {page_url}")
    driver.get(page_url)
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'h2, h3, div.image_content, a.add_image.upload, li#no_results')
            )
        )
    except Exception:
        pass
    time.sleep(1)

    if skip_if_exists and _has_existing_images(driver):
        if upload_all_state and upload_all_state[0]:
            log.info(f"{label} 已有圖片，自動上載（all 模式）。")
        else:
            log.info(f"{label} 已有圖片，跳過上傳。")
            return True

    sys.path.insert(0, script_dir)
    import viutv_image_uploader as _img
    return _img._upload_via_tmdb_cropper(driver, image_path, label)


def upload_show_poster(driver, show_id: str, image_path: str, language: str = 'zh-HK') -> bool:
    url = f"https://www.themoviedb.org/tv/{show_id}/images/posters?language={language}"
    return _upload_to_tmdb_image_page(driver, url, image_path, f"節目 {show_id} 封面")


def upload_show_backdrop(driver, show_id: str, image_path: str, language: str = 'zh-HK') -> bool:
    url = f"https://www.themoviedb.org/tv/{show_id}/images/backdrops?language={language}"
    return _upload_to_tmdb_image_page(driver, url, image_path, f"節目 {show_id} 橫幅")


def _reencode_jpeg(path: str) -> str:
    """
    Re-save an image as a clean JPEG to strip metadata or encoding quirks that
    some CDNs produce and that TMDB's upload endpoint may reject.
    Returns the path of the new file (same dir, .jpeg extension).
    """
    from PIL import Image as _PILImage
    out = os.path.splitext(path)[0] + '_reenc.jpeg'
    with _PILImage.open(path) as _im:
        _im.convert('RGB').save(out, 'JPEG', quality=92)
    log.info(f"已重新編碼 JPEG: {os.path.basename(out)}")
    return out


def upload_episode_backdrop(
    driver,
    show_id: str,
    season: int,
    episode: int,
    image_path: str,
    language: str = 'zh-HK',
    skip_if_exists: bool = True,
    upload_all_state: list | None = None,
) -> bool:
    url = (
        f"https://www.themoviedb.org/tv/{show_id}"
        f"/season/{season}/episode/{episode}"
        f"/images/backdrops?language={language}"
    )
    return _upload_to_tmdb_image_page(
        driver, url, image_path, f"第 {episode} 集劇照",
        skip_if_exists=skip_if_exists, upload_all_state=upload_all_state,
    )


def _upload_episode_backdrops(
    driver,
    show_id: str,
    viutv_url: str,
    is_viu_com: bool,
    ep_mapping: list,  # [(viu_ep_num, tmdb_season, tmdb_ep), ...]
    tmdb_lang: str,
    temp_dir: str,
    cached_images: dict | None = None,  # {ep_num: image_url} from first scrape
    auto_upload: bool = False,  # skip interactive prompt (for bulk mode)
):
    """
    Upload episode backdrops from viu.tv/viu.com to TMDB.
    If cached_images is provided (from scrape_viutv_show), the viu.tv page is NOT
    re-scraped.  Otherwise viutv_image_uploader.extract_episode_images is called.
    ep_mapping defines the viu episode number → (TMDB season, TMDB episode) translation.
    """
    if not auto_upload:
        ans = input("\n是否上傳集數劇照至 TMDB? (y/n，默認 n): ").strip().lower()
        if ans != 'y':
            return

    try:
        sys.path.insert(0, script_dir)
        import viutv_image_uploader as _img
    except ImportError as e:
        print(f"無法載入 viutv_image_uploader（可能缺少 Pillow 套件: {e}），跳過劇照上傳。")
        return

    if cached_images is not None:
        episode_images = cached_images
        print(f"\n使用已抓取的集數劇照 ({len(episode_images)} 集)，無需重新載入頁面。")
    else:
        print("\n正在提取集數劇照連結...")
        if is_viu_com:
            episode_images = _img.extract_episode_images_viu(driver, viutv_url)
        else:
            episode_images = _img.extract_episode_images(driver, viutv_url)

    if not episode_images:
        print("未找到集數劇照，跳過。")
        return

    viu_to_tmdb = {viu_ep: (tmdb_s, tmdb_ep) for viu_ep, tmdb_s, tmdb_ep in ep_mapping}
    target_eps = [ep for ep in sorted(episode_images) if ep in viu_to_tmdb]
    print(f"共找到 {len(episode_images)} 集劇照，其中 {len(target_eps)} 集對應 TMDB 集數。")

    success = 0
    upload_all_state = [False]
    auto_crop = None  # None=not decided, True=always crop

    def _upload_one(viu_ep: int) -> bool:
        """Download, crop and upload one episode backdrop. Returns True on success."""
        nonlocal auto_crop
        tmdb_season, tmdb_ep = viu_to_tmdb[viu_ep]
        img_url = episode_images[viu_ep]

        ext = os.path.splitext(img_url.split('?')[0])[1] or '.jpg'
        dest = os.path.join(temp_dir, f"ep_bd_{viu_ep:04d}{ext}")

        if not download_image(img_url, dest):
            log.warning(f"第 {viu_ep} 集劇照下載失敗。")
            return False

        dest = _img._ensure_jpg(dest)

        # Check aspect ratio and offer to crop to 16:9 if needed
        try:
            from PIL import Image as _PILImage
            with _PILImage.open(dest) as _im:
                _w, _h = _im.size
            _ratio = _w / _h
            _is_16_9 = abs(_ratio - 16 / 9) <= 0
        except Exception:
            _is_16_9 = True

        if not _is_16_9:
            _do_crop = False
            if auto_crop is True:
                _do_crop = True
            else:
                print(f"\n第 {viu_ep} 集圖片尺寸: {_w}x{_h}（比例 {_ratio:.3f}，非 16:9）。")
                _ans = input("是否裁切成 16:9? (y=裁切此集, a=全部裁切, n=跳過): ").strip().lower()
                if _ans in ('y', 'a'):
                    _do_crop = True
                if _ans == 'a':
                    auto_crop = True
            if _do_crop:
                _img.crop_to_16_9(dest)

        log.info(f"上傳第 {viu_ep} 集劇照 → S{tmdb_season:02d}E{tmdb_ep:02d}")
        for _attempt in range(3):
            ok = upload_episode_backdrop(
                driver, show_id, tmdb_season, tmdb_ep, dest,
                language=tmdb_lang, skip_if_exists=True, upload_all_state=upload_all_state,
            )
            if ok:
                time.sleep(1)
                return True
            if _attempt < 2:
                log.warning(f"第 {viu_ep} 集上傳失敗（第 {_attempt + 1} 次），2 秒後重試...")
                time.sleep(2)

        # 3 attempts failed → re-encode JPEG and try once more
        log.warning(f"第 {viu_ep} 集連續 3 次上傳失敗，重新編碼 JPEG 後再試一次...")
        try:
            reenc = _reencode_jpeg(dest)
            ok = upload_episode_backdrop(
                driver, show_id, tmdb_season, tmdb_ep, reenc,
                language=tmdb_lang, skip_if_exists=False, upload_all_state=upload_all_state,
            )
        except Exception as _e:
            log.error(f"第 {viu_ep} 集重新編碼上傳失敗: {_e}")
            ok = False
        time.sleep(1)
        return ok

    # Main upload pass
    failed: list[int] = []
    for viu_ep in target_eps:
        if _upload_one(viu_ep):
            success += 1
        else:
            failed.append(viu_ep)

    # Offer retry until user declines or all succeed
    while failed:
        print(f"\n劇照上傳完成！成功 {success}/{len(target_eps)} 集。")
        print(f"失敗集數 ({len(failed)} 集): {failed}")
        retry_ans = input("是否重試失敗的集數? (y=重試, n=跳過): ").strip().lower()
        if retry_ans != 'y':
            break
        retry_list, failed = list(failed), []
        for viu_ep in retry_list:
            if _upload_one(viu_ep):
                success += 1
            else:
                failed.append(viu_ep)

    print(f"\n劇照上傳完成！成功 {success}/{len(target_eps)} 集。")
    if failed:
        print(f"失敗集數 (viu 編號): {failed}")


def _upload_netflix_episode_backdrops(
    driver,
    show_id: str,
    season_num: int,
    episodes: list,
    tmdb_lang: str,
    temp_dir: str,
) -> None:
    """
    For each Netflix episode that has runtime > 0 and an image URL, upload the
    thumbnail to TMDB as an episode backdrop if TMDB doesn't already have one.

    Episodes with runtime == 0 are skipped (unreleased episodes show a generic
    placeholder that shouldn't be uploaded).

    If the downloaded image is smaller than 400×225, it is upscaled to exactly
    400×225 using LANCZOS resampling before upload.
    """
    ans = input("\n是否上傳 Netflix 集數劇照至 TMDB? (y/n，默認 n): ").strip().lower()
    if ans != 'y':
        return

    try:
        sys.path.insert(0, script_dir)
        import viutv_image_uploader as _img
    except ImportError as e:
        print(f"無法載入 viutv_image_uploader（可能缺少 Pillow 套件: {e}），跳過劇照上傳。")
        return

    # Only episodes that are confirmed released (runtime > 0) and have a thumbnail
    target_eps = [ep for ep in episodes if ep.get('runtime', 0) > 0 and ep.get('image_url')]
    skipped_no_runtime = len(episodes) - len(target_eps)

    if not target_eps:
        print("沒有符合條件的集數劇照可上傳（runtime=0 的集數已跳過）。")
        return

    print(f"共 {len(target_eps)} 集有劇照可上傳"
          + (f"（已跳過 {skipped_no_runtime} 集 runtime=0 的集數）。" if skipped_no_runtime else "。"))

    success = 0
    upload_all_state = [False]
    failed: list[int] = []
    ep_by_num = {ep['ep_num']: ep for ep in target_eps}

    def _upload_one_nf(ep: dict) -> bool:
        ep_num = ep['ep_num']
        img_url = ep['image_url']

        ext = os.path.splitext(img_url.split('?')[0])[1] or '.webp'
        dest = os.path.join(temp_dir, f"nf_ep_bd_{season_num:02d}_{ep_num:04d}{ext}")

        if not download_image(img_url, dest):
            log.warning(f"第 {ep_num} 集劇照下載失敗。")
            return False

        # Convert to JPEG (handles .webp and other formats)
        dest = _img._ensure_jpg(dest)

        # Upscale to 400×225 if the image is smaller
        try:
            from PIL import Image as _PILImage
            with _PILImage.open(dest) as _im:
                _w, _h = _im.size
            if _w < 400 or _h < 225:
                log.info(f"第 {ep_num} 集圖片 {_w}×{_h} 小於 400×225，放大中...")
                with _PILImage.open(dest) as _im:
                    _im_up = _im.resize((400, 225), _PILImage.LANCZOS)
                    _im_up.save(dest, 'JPEG', quality=90)
                log.info(f"第 {ep_num} 集圖片已放大至 400×225。")
        except Exception as exc:
            log.warning(f"第 {ep_num} 集圖片尺寸處理失敗: {exc}")

        log.info(f"上傳第 {ep_num} 集劇照 → S{season_num:02d}E{ep_num:02d}")
        for _attempt in range(3):
            ok = upload_episode_backdrop(
                driver, show_id, season_num, ep_num, dest,
                language=tmdb_lang, skip_if_exists=True, upload_all_state=upload_all_state,
            )
            if ok:
                time.sleep(1)
                return True
            if _attempt < 2:
                log.warning(f"第 {ep_num} 集上傳失敗（第 {_attempt + 1} 次），2 秒後重試...")
                time.sleep(2)

        # 3 attempts failed → re-encode JPEG and try once more
        log.warning(f"第 {ep_num} 集連續 3 次上傳失敗，重新編碼 JPEG 後再試一次...")
        try:
            reenc = _reencode_jpeg(dest)
            ok = upload_episode_backdrop(
                driver, show_id, season_num, ep_num, reenc,
                language=tmdb_lang, skip_if_exists=False, upload_all_state=upload_all_state,
            )
        except Exception as _e:
            log.error(f"第 {ep_num} 集重新編碼上傳失敗: {_e}")
            ok = False
        time.sleep(1)
        return ok

    # Main upload pass
    for ep in target_eps:
        if _upload_one_nf(ep):
            success += 1
        else:
            failed.append(ep['ep_num'])

    # Retry loop
    while failed:
        print(f"\n劇照上傳完成！成功 {success}/{len(target_eps)} 集。")
        print(f"失敗集數 ({len(failed)} 集): {failed}")
        retry_ans = input("是否重試失敗的集數? (y=重試, n=跳過): ").strip().lower()
        if retry_ans != 'y':
            break
        retry_list, failed = list(failed), []
        for ep_num in retry_list:
            ep = ep_by_num.get(ep_num)
            if ep and _upload_one_nf(ep):
                success += 1
            else:
                failed.append(ep_num)

    print(f"\n劇照上傳完成！成功 {success}/{len(target_eps)} 集。")
    if failed:
        print(f"失敗集數: {failed}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _parse_show_id(raw: str) -> str | None:
    """Accept a bare numeric ID or any TMDB /tv/{id} URL. Returns the numeric ID or None."""
    raw = raw.strip()
    m = re.search(r'themoviedb\.org/tv/(\d+)', raw)
    if m:
        return m.group(1)
    if re.fullmatch(r'\d+', raw):
        return raw
    return None


def _print_show_summary(show_info: dict, source: str = 'viu.tv'):
    """Print a summary of scraped show data."""
    ep1 = next((e for e in show_info.get('episodes', []) if e['ep_num'] == 1), None)
    print(f"\n=== 從 {source} 取得的資料 ===")
    print(f"標題     : {show_info['title']}")
    desc = show_info['description']
    print(f"描述     : {desc[:80]}..." if len(desc) > 80 else f"描述     : {desc}")
    print(f"集數     : {show_info['episode_count']}")
    print(f"類型     : {', '.join(show_info['genres'])}")
    print(f"首播日期 : {show_info['first_air_date']}")
    if ep1:
        ep1_desc = ep1.get('description', '')
        print(f"第 1 集   : {ep1['title']}")
        print(f"第 1 集描述: {ep1_desc[:60]}..." if len(ep1_desc) > 60 else f"第 1 集描述: {ep1_desc}")
    print(f"封面 URL : {show_info['poster_url']}")
    print(f"橫幅 URL : {show_info['banner_url']}")


def _upload_images(driver, show_id: str, show_info: dict, temp_dir: str, tmdb_lang: str,
                   auto_poster: bool = False):
    """Download and upload poster/backdrop. auto_poster skips the y/n prompt for poster."""
    poster_url = show_info['poster_url']
    banner_url  = show_info['banner_url']

    if poster_url:
        do_poster = True if auto_poster else (
            input("\n是否上傳封面圖片至 Posters? (y/n): ").strip().lower() == 'y'
        )
        if do_poster:
            ext = os.path.splitext(poster_url.split('?')[0])[1] or '.jpg'
            poster_path = os.path.join(temp_dir, f"poster{ext}")
            if download_image(poster_url, poster_path):
                if auto_poster:
                    print("\n正在上傳封面圖片...")
                upload_show_poster(driver, show_id, poster_path, tmdb_lang)
            else:
                print("封面圖片下載失敗，請手動上傳。")
    else:
        print("未找到封面圖片，請手動上傳。")

    if banner_url and banner_url != poster_url:
        if input("\n是否上傳橫幅圖片至 Backdrops? (y/n): ").strip().lower() == 'y':
            ext = os.path.splitext(banner_url.split('?')[0])[1] or '.jpg'
            banner_path = os.path.join(temp_dir, f"banner{ext}")
            if download_image(banner_url, banner_path):
                upload_show_backdrop(driver, show_id, banner_path, tmdb_lang)
            else:
                print("橫幅圖片下載失敗，請手動上傳。")


def _check_season_exists(driver, show_id: str, season: int) -> bool:
    """
    Navigate to the season edit page and return True if the season exists
    (i.e. the Kendo episode grid is present).
    """
    url = (f"https://www.themoviedb.org/tv/{show_id}/season/{season}/edit"
           f"?active_nav_item=episodes")
    driver.get(url)
    time.sleep(2)
    if "Oops" in driver.page_source:
        return False
    return bool(driver.find_elements(By.CSS_SELECTOR, '.k-grid, button.k-grid-add'))


def _create_tmdb_season(driver, show_id: str, season_num: int,
                         name: str = '', overview: str = '') -> bool:
    """
    Create a new season via the TMDB seasons management grid.

    Navigates to /tv/{show_id}/edit?active_nav_item=seasons, clicks
    "加入新的一季", sets the season number (and optionally name/overview),
    then clicks 儲存. Returns True on success.
    """
    seasons_url = f"https://www.themoviedb.org/tv/{show_id}/edit?active_nav_item=seasons"
    log.info(f"前往季數管理頁面以建立第 {season_num} 季: {seasons_url}")
    driver.get(seasons_url)
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'button.k-grid-add'))
        )
    except Exception:
        log.error("季數管理頁面載入失敗，找不到「加入新的一季」按鈕。")
        return False
    time.sleep(1)

    try:
        add_btn = driver.find_element(By.CSS_SELECTOR, 'button.k-grid-add')
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", add_btn)
        driver.execute_script("arguments[0].click();", add_btn)
    except Exception as e:
        log.error(f"點擊「加入新的一季」失敗: {e}")
        return False

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.k-window'))
        )
    except Exception:
        log.error("季數編輯彈窗未出現。")
        return False
    time.sleep(0.5)

    _set_kendo_episode_number(driver, season_num)

    if name:
        _fill_form_field(driver, 'zh-HK_name_text_input_field', add_spaces_between_en_zh_num(name))
    if overview:
        _fill_form_field(driver, 'zh-HK_overview_text_box_field', add_spaces_between_en_zh_num(overview))

    if _click_save_button(driver):
        log.info(f"第 {season_num} 季已成功建立。")
        time.sleep(2)
        return True

    log.error(f"第 {season_num} 季建立失敗（儲存按鈕點擊失敗）。")
    return False


def _ensure_season_exists(driver, show_id: str, season_num: int,
                           name: str = '', overview: str = '') -> None:
    """
    Check whether a season exists on TMDB; auto-create it if not.
    Falls back to a manual prompt if automatic creation fails.
    """
    if _check_season_exists(driver, show_id, season_num):
        return
    log.info(f"第 {season_num} 季不存在，嘗試自動建立...")
    if _create_tmdb_season(driver, show_id, season_num, name, overview):
        return
    print(f"\n自動建立第 {season_num} 季失敗，請在瀏覽器中手動建立：")
    print(f"  https://www.themoviedb.org/tv/{show_id}/edit?active_nav_item=seasons")
    input("建立後按 Enter 繼續...")


def _ask_season_ranges(episodes: list, num_seasons: int | None = None) -> tuple[list[tuple[int, int, int]], bool]:
    """
    Prompt the user to define which viu.tv episode numbers belong to each TMDB season.
    Returns (ranges, renumber) where:
      ranges   = list of (season_num, viu_start, viu_end)
      renumber = True  → reset to E1 each season (standard)
                 False → keep original viu.tv episode numbers (continuous across seasons)
    If num_seasons is provided, the "共有幾季?" prompt is skipped.
    """
    all_nums = sorted(ep['ep_num'] for ep in episodes)
    print(f"\n共抓取到 {len(all_nums)} 集 (viu.tv 第 {all_nums[0]}–{all_nums[-1]} 集)")

    print("\n集數編號方式:")
    print("  1. 每季重新從第 1 集開始 (標準做法，例如 S02E01)")
    print("  2. 保留 viu.tv 原有編號，跨季連續 (例如 S02E43)")
    while True:
        ans = input("選擇 (1/2，默認 1): ").strip() or '1'
        if ans in ('1', '2'):
            renumber = (ans == '1')
            break
        print("請輸入 1 或 2。")

    if num_seasons is None:
        while True:
            raw_n = input("共有幾季? ").strip()
            if raw_n.isdigit() and int(raw_n) >= 1:
                num_seasons = int(raw_n)
                break
            print("請輸入有效的季數。")

    ranges = []
    for s in range(1, num_seasons + 1):
        while True:
            raw = input(f"第 {s} 季: viu.tv 第幾集到第幾集? (例如 1-54): ").strip()
            m = re.match(r'^(\d+)-(\d+)$', raw)
            if m:
                start, end = int(m.group(1)), int(m.group(2))
                if start <= end:
                    ranges.append((s, start, end))
                    break
            print("格式錯誤，請重新輸入 (例如 1-54)。")

    return ranges, renumber


# ---------------------------------------------------------------------------
# viu.com page scraping (separate site from viu.tv)
# ---------------------------------------------------------------------------

def _expand_all_episode_descriptions_viu(driver, panel_id: str):
    """Click all collapsed 'show more' buttons inside a viu.com tab panel."""
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
        log.info(f"viu.com: 已展開 {len(buttons)} 個集數描述")


def _extract_episodes_from_panel_viu(panel, season_number: int) -> list:
    """Parse episode cards from a BeautifulSoup viu.com tab-panel element."""
    episodes = []
    for item in panel.find_all(class_='series_ep'):
        ep_title_el = item.find(id='series_ep_title')
        if not ep_title_el:
            continue
        full_title = ep_title_el.get_text(strip=True)
        match = re.match(r'^(\d+)\.\s*(.*)', full_title)
        if not match:
            log.warning(f"viu.com: 無法解析集數標題格式: {full_title!r}，跳過。")
            continue

        episode_number = int(match.group(1))
        episode_title = add_spaces_between_en_zh_num(match.group(2).strip())

        dur_el = item.find(class_='Typography-bodySm')
        duration_text = dur_el.get_text(strip=True) if dur_el else ''
        episode_runtime = _parse_duration_to_minutes(duration_text)

        desc_el = item.find(class_='Typography-bodyMd')
        episode_description = add_spaces_between_en_zh_num(
            desc_el.get_text(strip=True)
        ) if desc_el else ''

        episodes.append({
            'ep_num': episode_number,
            'season': season_number,
            'title': episode_title,
            'description': episode_description,
            'runtime': episode_runtime,
        })
        log.info(f"  viu.com S{season_number:02d}E{episode_number:02d}: {episode_title} ({episode_runtime} 分鐘)")
    return episodes


def scrape_viu_show(driver, viu_url: str) -> dict:
    """
    Load a viu.com show page and return a dict with:
      title, description, poster_url, banner_url, episode_count,
      num_seasons, episodes, genres, first_air_date.
    episodes is a flat list; each entry has a 'season' key.
    """
    log.info(f"載入 viu.com 頁面: {viu_url}")
    driver.get(viu_url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'series_title'))
    )
    time.sleep(2)

    og_image = driver.execute_script(
        "var el = document.querySelector('meta[property=\"og:image\"]');"
        "return el ? el.getAttribute('content') : null;"
    ) or ''
    log.info(f"og:image: {og_image}")

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

    log.info(f"viu.com 標題: {title}")

    # Only select episode-list tabs (ep_list_tab), ignoring recommend tabs
    ep_tabs = driver.find_elements(By.CSS_SELECTOR, 'button.ep_list_tab[role="tab"]')
    num_seasons = max(len(ep_tabs), 1)
    log.info(f"偵測到 {num_seasons} 個季度分頁。")

    # Wait for at least one episode card
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.series_ep'))
        )
    except Exception:
        log.warning("等待 .series_ep 超時，繼續嘗試解析...")

    all_episodes: list = []

    if len(ep_tabs) > 1:
        for i, tab in enumerate(ep_tabs):
            season_num = i + 1
            log.info(f"切換至第 {season_num} 季分頁...")
            driver.execute_script("arguments[0].click();", tab)
            time.sleep(1.5)
            panel_id = tab.get_attribute('aria-controls')
            if panel_id:
                _expand_all_episode_descriptions_viu(driver, panel_id)
            panel_soup = BeautifulSoup(driver.page_source, 'html.parser')
            active_panel = (
                panel_soup.find(id=panel_id) if panel_id
                else panel_soup.find('div', attrs={'role': 'tabpanel'})
            )
            if active_panel:
                all_episodes.extend(_extract_episodes_from_panel_viu(active_panel, season_num))
    else:
        # Single season: try tab button's aria-controls, or derive from the panel element
        panel_id = ep_tabs[0].get_attribute('aria-controls') if ep_tabs else None
        if not panel_id:
            current_soup = BeautifulSoup(driver.page_source, 'html.parser')
            panel_el = current_soup.find('div', attrs={'role': 'tabpanel'})
            panel_id = panel_el.get('id') if panel_el else None
            log.info(f"從 tabpanel 元素取得 panel_id: {panel_id!r}")

        if panel_id:
            _expand_all_episode_descriptions_viu(driver, panel_id)
            panel_soup = BeautifulSoup(driver.page_source, 'html.parser')
            panel = panel_soup.find(id=panel_id)
        else:
            panel_soup = BeautifulSoup(driver.page_source, 'html.parser')
            panel = panel_soup.find('div', attrs={'role': 'tabpanel'})

        if panel:
            all_episodes.extend(_extract_episodes_from_panel_viu(panel, 1))
        else:
            log.warning("viu.com: 找不到 tabpanel 元素，無法解析集數。")

    log.info(f"viu.com 共解析 {len(all_episodes)} 集 (共 {num_seasons} 季)")

    # viu.com shows no separate portrait poster — the in-page horizontal banner is the cover.
    # og:image may reflect a currently featured (different) show, so prefer the in-page image.
    soup_for_banner = BeautifulSoup(driver.page_source, 'html.parser')
    for ep_el in soup_for_banner.find_all(class_='series_ep'):
        ep_el.decompose()
    banner_url = ''
    for img in soup_for_banner.find_all('img'):
        src = img.get('src', '') or img.get('data-src', '')
        if 'original_horizontal' in src:
            banner_url = src
            break
        srcset = img.get('srcset', '')
        if 'original_horizontal' in srcset:
            parts = [p.strip().split()[0] for p in srcset.split(',') if p.strip()]
            if parts:
                banner_url = parts[-1]
            break
    poster_url = banner_url or og_image
    log.info(f"viu.com 封面/橫幅: {poster_url}")

    return {
        'title': title,
        'description': description,
        'poster_url': poster_url,
        'banner_url': banner_url,
        'episode_count': len(all_episodes),
        'num_seasons': num_seasons,
        'episodes': all_episodes,
        'genres': [],
        'first_air_date': '',
    }


def _update_season_translation(
    driver, show_id: str, season_num: int,
    title: str, description: str, language: str = 'zh-HK'
) -> bool:
    """
    Write season name and overview to the TMDB season edit page.
    Creates the zh-HK translation row if it does not yet exist.
    """
    field_id = f'{language.replace("-", "_")}_name'
    overview_id = f'{language.replace("-", "_")}_overview'
    edit_url = (
        f'https://www.themoviedb.org/tv/{show_id}'
        f'/season/{season_num}/edit?language={language}'
    )
    log.info(f"更新 TMDB 第 {season_num} 季翻譯: {edit_url}")
    try:
        driver.get(edit_url)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, 'body'))
        )
        time.sleep(1)

        try:
            driver.find_element(By.ID, field_id)
            log.info(f"欄位 '{field_id}' 已存在。")
        except Exception:
            log.info(f"欄位 '{field_id}' 不存在，嘗試點擊「新增翻譯」...")
            try:
                btn = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(text(), '新增翻譯')]")
                    )
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                btn.click()
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, field_id))
                )
                log.info("已成功新增翻譯。")
            except Exception as e:
                log.error(f"新增翻譯失敗: {e}")
                return False

        if title:
            f = driver.find_element(By.ID, field_id)
            f.clear()
            f.send_keys(title)

        try:
            f = driver.find_element(By.ID, overview_id)
            f.clear()
            if description:
                f.send_keys(description)
        except Exception:
            log.warning(f"找不到欄位 '{overview_id}'，跳過描述。")

        submit = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'submit'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", submit)
        submit.click()
        time.sleep(2)
        log.info(f"第 {season_num} 季翻譯已提交。")
        return True

    except Exception as e:
        log.error(f"_update_season_translation 失敗: {e}")
        return False


def _input_tmdb_lang() -> str:
    """Prompt for TMDB language tag; reject obvious mis-pastes (URLs, long strings)."""
    _lang_re = re.compile(r'^[a-zA-Z]{2}(-[a-zA-Z0-9]{2,8})*$')
    while True:
        raw = input("TMDB 上傳語言標籤 (默認 zh-HK): ").strip()
        if not raw:
            return 'zh-HK'
        if _lang_re.match(raw) and len(raw) <= 20:
            return raw
        print(f"  ⚠ 格式錯誤: {raw!r:.60}  (應為語言標籤如 zh-HK、en-US，按 Enter 使用默認值)")


def main():
    print("=== viu.tv / viu.com / Netflix 節目 TMDB 管理工具 ===\n")
    print("模式選擇:")
    print("  1. 建立全新節目 (從 viu.tv / viu.com / Netflix 抓取並在 TMDB 建立新條目)")
    print("  2. 為現有節目加入新一季集數")
    print("  3. 為現有節目分多季加入集數 (來源無分季，TMDB 有多季)")
    print("  4. 為現有節目更新最新集數 (只新增未有的集數，不修改已有資料)")
    print("  5. 按年份自動分季新增集數 (集數標題=播出日期，每年一季，集數連續編號)")
    mode = input("選擇 (1-5，默認 1): ").strip() or '1'

    if mode == '1':
        viutv_url = input("節目 URL (viu.tv、viu.com 或 Netflix): ").strip()
        tmdb_lang = _input_tmdb_lang()
        existing_show_id = None
        season_num = 1

    elif mode == '2':
        viutv_url = input("節目 URL (viu.tv、viu.com 或 Netflix 新一季頁面): ").strip()
        raw_id = input("TMDB 節目 ID 或 URL (例如 216055 或 https://www.themoviedb.org/tv/216055): ").strip()
        existing_show_id = _parse_show_id(raw_id)
        if not existing_show_id:
            print(f"無法解析 TMDB 節目 ID: {raw_id!r}，程式終止。")
            return
        season_num = int(input("季數 (默認 2): ").strip() or '2')
        tmdb_lang = _input_tmdb_lang()

    elif mode == '3':
        viutv_url = input("節目 URL (viu.tv、viu.com 或 Netflix): ").strip()
        raw_id = input("TMDB 節目 ID 或 URL (例如 317017 或 https://www.themoviedb.org/tv/317017): ").strip()
        existing_show_id = _parse_show_id(raw_id)
        if not existing_show_id:
            print(f"無法解析 TMDB 節目 ID: {raw_id!r}，程式終止。")
            return
        tmdb_lang = _input_tmdb_lang()

    elif mode == '4':
        viutv_url = input("節目 URL (viu.tv、viu.com 或 Netflix): ").strip()
        raw_id = input("TMDB 節目 ID 或 URL (例如 321747 或 https://www.themoviedb.org/tv/321747): ").strip()
        existing_show_id = _parse_show_id(raw_id)
        if not existing_show_id:
            print(f"無法解析 TMDB 節目 ID: {raw_id!r}，程式終止。")
            return
        season_num = int(input("季數 (默認 1): ").strip() or '1')
        tmdb_lang = _input_tmdb_lang()

    elif mode == '5':
        viutv_url = input("節目 URL (viu.tv): ").strip()
        raw_id = input("TMDB 節目 ID 或 URL (留空則建立全新節目): ").strip()
        if raw_id:
            existing_show_id = _parse_show_id(raw_id)
            if not existing_show_id:
                print(f"無法解析 TMDB 節目 ID: {raw_id!r}，程式終止。")
                return
        else:
            existing_show_id = None
        tmdb_lang = _input_tmdb_lang()
        season_num = 1  # placeholder; mode 5 sets this per year-group

    else:
        print(f"不認識的模式: {mode}")
        return

    is_viu_com = 'viu.com' in viutv_url and 'netflix.com' not in viutv_url
    is_netflix = 'netflix.com' in viutv_url

    temp_dir = tempfile.mkdtemp(prefix='viutv_creator_')
    log.info(f"臨時目錄: {temp_dir}")

    driver = None
    try:
        configs = load_configs()
        driver_tuple = get_browser_driver(profile_suffix='_viutv_show', remote_port=9225, base_dir=project_root)
        driver = driver_tuple[0]

        # Ensure TMDB login
        if not is_tmdb_logged_in(driver):
            username, password = get_tmdb_credentials(configs)
            if username and password:
                login(driver, username, password)
                log.info("已登入 TMDB。")
            else:
                log.warning("未找到 TMDB 憑證，請先在瀏覽器中手動登入。")
                input("登入後請按 Enter 繼續...")

        # Scrape the show page (viu.com, Netflix, or viu.tv)
        if is_netflix:
            netflix_data = scrape_netflix_show(driver, viutv_url)
            if not netflix_data.get('title'):
                log.error("無法取得 Netflix 節目標題，程式終止。")
                return

            available_nf_seasons = sorted(netflix_data.get('_ep_by_season', {}).keys())
            if not available_nf_seasons:
                log.error("Netflix 未抓取到任何集數，程式終止。")
                return

            if len(available_nf_seasons) > 1:
                season_labels = {sn: name for name, sn in netflix_data.get('_seasons', [])}
                print(f"\nNetflix 共有 {len(available_nf_seasons)} 季:")
                for sn in available_nf_seasons:
                    eps = netflix_data['_ep_by_season'].get(sn, [])
                    print(f"  第 {sn} 季 ({season_labels.get(sn, f'第 {sn} 季')}): {len(eps)} 集")
                default_s = available_nf_seasons[0]
                raw_s = input(f"使用哪一季的集數? (默認 {default_s}): ").strip()
                nf_season_to_use = int(raw_s) if raw_s.isdigit() else default_s
            else:
                nf_season_to_use = available_nf_seasons[0]

            flat_eps = netflix_data['_ep_by_season'].get(nf_season_to_use, [])
            show_info = {
                'title': netflix_data['title'],
                'description': netflix_data['description'],
                'poster_url': netflix_data['poster_url'],
                'banner_url': netflix_data['banner_url'],
                'episode_count': len(flat_eps),
                'episodes': flat_eps,
                'genres': [],
                'first_air_date': '',
            }

        elif is_viu_com:
            viu_com_data = scrape_viu_show(driver, viutv_url)
            if not viu_com_data['title']:
                log.error("無法取得節目標題，程式終止。")
                return

            available_viu_seasons = sorted({ep['season'] for ep in viu_com_data['episodes']})

            if len(available_viu_seasons) > 1:
                print(f"\nviu.com 共有 {len(available_viu_seasons)} 季: {available_viu_seasons}")
                default_s = available_viu_seasons[0]
                raw_s = input(f"使用哪一季的集數? (默認 {default_s}): ").strip()
                viu_season_to_use = int(raw_s) if raw_s.isdigit() else default_s
            else:
                viu_season_to_use = 1

            flat_eps = [ep for ep in viu_com_data['episodes'] if ep['season'] == viu_season_to_use]
            show_info = {
                'title': viu_com_data['title'],
                'description': viu_com_data['description'],
                'poster_url': viu_com_data['poster_url'],
                'banner_url': viu_com_data['banner_url'],
                'episode_count': len(flat_eps),
                'episodes': flat_eps,
                'genres': [],
                'first_air_date': '',
            }
        else:
            show_info = scrape_viutv_show(driver, viutv_url)
            if not show_info['title']:
                log.error("無法取得節目標題，程式終止。")
                return

        _source_label = 'Netflix' if is_netflix else ('viu.com' if is_viu_com else 'viu.tv')
        _print_show_summary(show_info, source=_source_label)

        # Mode 2 uses a single air-date schedule; modes 1 & 3 handle dates internally
        air_dates: dict = {}
        if mode == '2':
            air_dates = _ask_air_dates(
                show_info['episodes'], default_start=show_info.get('first_air_date', '')
            )

        # ── Mode 2: Add new season to existing show ──────────────────────────
        if mode == '2':
            confirm = input(
                f"\n確認為 TMDB 節目 {existing_show_id} 新增第 {season_num} 季"
                f" ({show_info['episode_count']} 集)? (y/n): "
            ).strip().lower()
            if confirm != 'y':
                print("已取消。")
                return

            # Verify the season exists on TMDB; auto-create it if not
            _ensure_season_exists(driver, existing_show_id, season_num)

            # For viu.com season 2+: show title/description → TMDB season name/description
            if is_viu_com and season_num >= 2 and (show_info['title'] or show_info['description']):
                _update_season_translation(
                    driver, existing_show_id, season_num,
                    show_info['title'], show_info['description'], tmdb_lang
                )

            add_episodes_to_season(
                driver, existing_show_id, season_num,
                show_info['episodes'], air_dates, tmdb_lang
            )

            _upload_images(driver, existing_show_id, show_info, temp_dir, tmdb_lang)

            if is_netflix:
                _upload_netflix_episode_backdrops(
                    driver, existing_show_id, season_num,
                    show_info['episodes'], tmdb_lang, temp_dir,
                )

            print(f"\n完成！請前往以下頁面確認資料:")
            print(f"  季數頁面: https://www.themoviedb.org/tv/{existing_show_id}/season/{season_num}")
            return

        # ── Mode 3: Split episodes across multiple seasons ────────────────────
        if mode == '3':
            season_ranges, renumber = _ask_season_ranges(show_info['episodes'])
            episodes_by_num = {ep['ep_num']: ep for ep in show_info['episodes']}

            # Ask for Wikipedia URL once; dates are reused (filtered) per season
            print("\n=== 播映日期 (所有季) ===")
            wiki_url_m3 = input("Wikipedia 頁面 URL (留空逐季手動設定): ").strip()
            wiki_dates_m3 = _fetch_wiki_dates(wiki_url_m3) if wiki_url_m3 else None
            if wiki_dates_m3:
                print(f"從 Wikipedia 成功抓取 {len(wiki_dates_m3)} 集播映日期。")

            for season_num, viu_start, viu_end in season_ranges:
                season_eps_raw = [
                    episodes_by_num[n]
                    for n in sorted(episodes_by_num)
                    if viu_start <= n <= viu_end
                ]
                if not season_eps_raw:
                    log.warning(f"第 {season_num} 季: 在 viu.tv 集數 {viu_start}–{viu_end} 範圍內找不到集數，跳過。")
                    continue

                if renumber:
                    # Reset episode numbers to 1-based within the season
                    offset = viu_start - 1
                    remapped = [{**ep, 'ep_num': ep['ep_num'] - offset} for ep in season_eps_raw]
                    ep_range_label = f"TMDB E1–E{len(remapped)}"
                else:
                    # Keep original viu.tv episode numbers (continuous across seasons)
                    remapped = list(season_eps_raw)
                    ep_range_label = f"TMDB E{viu_start}–E{viu_end}"

                print(f"\n=== 第 {season_num} 季 ({len(remapped)} 集: viu.tv {viu_start}–{viu_end} → {ep_range_label}) ===")
                season_air_dates = _ask_air_dates(remapped, cached_wiki_dates=wiki_dates_m3)

                _ensure_season_exists(driver, existing_show_id, season_num)

                add_episodes_to_season(
                    driver, existing_show_id, season_num,
                    remapped, season_air_dates, tmdb_lang
                )

            print(f"\n完成！請前往節目頁面確認資料:")
            print(f"  https://www.themoviedb.org/tv/{existing_show_id}")
            return

        # ── Mode 4: Update new episodes in existing season ───────────────────
        if mode == '4':
            # show_info is already populated by the common scraping block above
            episodes = show_info['episodes']
            ep_nums_all = sorted(ep['ep_num'] for ep in episodes)
            print(f"\n從 {_source_label} 抓取到 {len(episodes)} 集 (第 {ep_nums_all[0]}–{ep_nums_all[-1]} 集)")

            # Read the TMDB grid to learn the episode range already in this season.
            # viu.tv returns all seasons' episodes in one list, so we must exclude
            # episodes that belong to earlier seasons (ep_num < min of existing grid).
            season_url = (
                f"https://www.themoviedb.org/tv/{existing_show_id}/season/{season_num}/edit"
                f"?active_nav_item=episodes"
            )
            log.info(f"預讀 TMDB 第 {season_num} 季 grid 以確定集數範圍...")
            driver.get(season_url)
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'button.k-grid-add, .k-grid'))
                )
            except Exception:
                pass
            time.sleep(2)
            existing_grid = _read_grid_episodes(driver)

            if existing_grid:
                min_grid_ep = min(existing_grid.keys())
                before = len(episodes)
                episodes = [ep for ep in episodes if ep['ep_num'] >= min_grid_ep]
                print(f"TMDB 第 {season_num} 季集數從 E{min_grid_ep} 開始，"
                      f"已過濾掉 {before - len(episodes)} 集屬於其他季的集數。")

            ep_nums = sorted(ep['ep_num'] for ep in episodes)
            print(f"將處理第 {ep_nums[0]}–{ep_nums[-1]} 集 (共 {len(episodes)} 集)")

            # Get air dates (Wikipedia or manual schedule)
            air_dates = _ask_air_dates(episodes)

            confirm = input(
                f"\n確認更新 TMDB 節目 {existing_show_id} 第 {season_num} 季新集數"
                f" (已有資料的集數將自動跳過)? (y/n): "
            ).strip().lower()
            if confirm != 'y':
                print("已取消。")
                return

            newly_added, edited = add_episodes_to_season(
                driver, existing_show_id, season_num, episodes, air_dates, tmdb_lang
            )
            modified = newly_added | edited

            if newly_added:
                print(f"\n已新增 {len(newly_added)} 集: {sorted(newly_added)}")
            if edited:
                print(f"已補填 {len(edited)} 集: {sorted(edited)}")

            # Upload backdrops for all modified episodes (newly added + filled placeholders)
            if is_netflix:
                _upload_netflix_episode_backdrops(
                    driver, existing_show_id, season_num,
                    episodes, tmdb_lang, temp_dir,
                )
            elif modified:
                ep_mapping = [(ep_num, season_num, ep_num) for ep_num in sorted(modified)]
                cached = {ep['ep_num']: ep['image_url']
                          for ep in episodes if ep.get('image_url')}
                _upload_episode_backdrops(
                    driver, existing_show_id, viutv_url, is_viu_com,
                    ep_mapping, tmdb_lang, temp_dir,
                    cached_images=cached or None,
                )

            print(f"\n完成！請前往季數頁面確認資料:")
            print(f"  https://www.themoviedb.org/tv/{existing_show_id}/season/{season_num}")
            return

        # ── Mode 5: Year-based season split ──────────────────────────────────
        if mode == '5':
            episodes = show_info['episodes']
            year_groups = _group_episodes_by_year(episodes)

            if not year_groups:
                print("未能從集數標題提取年份（標題應為 YYYY-MM-DD 格式），程式終止。")
                return

            years_sorted = sorted(year_groups.keys())
            print(f"\n從 {_source_label} 抓取到 {len(episodes)} 集，偵測到年份: {years_sorted}")

            # Ask user which year is Season 1 so we can compute season numbers
            raw_sy = input(f"第 1 季對應哪一年? (默認 {years_sorted[0]}): ").strip()
            start_year = int(raw_sy) if raw_sy.isdigit() and int(raw_sy) > 1900 else years_sorted[0]

            year_to_season = {y: y - start_year + 1 for y in years_sorted}
            print("\n年份 → 季數映射 (集數連續，不重新編號):")
            for y, s in sorted(year_to_season.items()):
                eps_y = year_groups[y]
                ep_nums_y = sorted(e['ep_num'] for e in eps_y)
                print(f"  {y} 年 → 第 {s} 季: E{ep_nums_y[0]}–E{ep_nums_y[-1]} ({len(eps_y)} 集)")

            if existing_show_id is None:
                # Create the show using the first episode's year-season data
                first_year = years_sorted[0]
                s1_eps = sorted(year_groups[first_year], key=lambda e: e['ep_num'])
                s1_air_dates = _ask_air_dates(s1_eps)
                ep1_date = min(s1_air_dates.values()) if s1_air_dates else ''
                if ep1_date:
                    show_info['first_air_date'] = ep1_date
                # Wizard uses the first year's episodes
                show_info['episodes'] = s1_eps

                confirm = input(f"\n確認在 TMDB 建立此節目並按年份分季? (y/n): ").strip().lower()
                if confirm != 'y':
                    print("已取消。")
                    return

                show_id = create_tmdb_show(driver, show_info)
                if not show_id:
                    manual_id = input("\n無法自動取得節目 ID。請手動填寫 TMDB 節目 ID (或按 Enter 跳過): ").strip()
                    if not manual_id:
                        return
                    show_id = manual_id
                print(f"\n節目已建立，TMDB ID: {show_id}")

                # Season 1 episodes are already added by the wizard; add remaining seasons
                seasons_to_process = [(y, year_to_season[y]) for y in years_sorted]
            else:
                show_id = existing_show_id
                s1_air_dates = {}

                confirm = input(f"\n確認為 TMDB 節目 {show_id} 按年份分季新增集數? (y/n): ").strip().lower()
                if confirm != 'y':
                    print("已取消。")
                    return

                seasons_to_process = [(y, year_to_season[y]) for y in years_sorted]

            ep_mapping = []
            for year, season_num in seasons_to_process:
                eps_year = sorted(year_groups[year], key=lambda e: e['ep_num'])
                ep_nums_y = sorted(e['ep_num'] for e in eps_year)
                print(f"\n=== {year} 年 → 第 {season_num} 季 "
                      f"(E{ep_nums_y[0]}–E{ep_nums_y[-1]}，{len(eps_year)} 集) ===")

                if year == years_sorted[0] and existing_show_id is None:
                    # Season 1 was created by the wizard; air dates already fetched
                    season_air_dates = s1_air_dates
                else:
                    season_air_dates = _ask_air_dates(eps_year)
                    _ensure_season_exists(driver, show_id, season_num)

                add_episodes_to_season(
                    driver, show_id, season_num,
                    eps_year, season_air_dates, tmdb_lang
                )

                for ep in eps_year:
                    ep_mapping.append((ep['ep_num'], season_num, ep['ep_num']))

            # Upload episode backdrops
            cached = {ep['ep_num']: ep['image_url']
                      for ep in episodes if ep.get('image_url')}
            _upload_episode_backdrops(
                driver, show_id, viutv_url, is_viu_com,
                ep_mapping, tmdb_lang, temp_dir,
                cached_images=cached or None,
            )

            _upload_images(driver, show_id, show_info, temp_dir, tmdb_lang,
                           auto_poster=(existing_show_id is None))

            print(f"\n完成！請前往節目頁面確認資料:")
            print(f"  https://www.themoviedb.org/tv/{show_id}")
            return

        # ── Mode 1: Create new show ──────────────────────────────────────────
        print("\n=== 季數設定 ===")
        raw_ns = input("共有幾季? (默認 1): ").strip() or '1'
        num_seasons_m1 = int(raw_ns) if raw_ns.isdigit() and int(raw_ns) >= 1 else 1
        all_episodes_m1 = list(show_info['episodes'])  # preserve before wizard may truncate

        if num_seasons_m1 == 1:
            # ── Single season ────────────────────────────────────────────────
            air_dates = _ask_air_dates(
                show_info['episodes'], default_start=show_info.get('first_air_date', '')
            )
            if air_dates:
                ep1_date = air_dates.get(1, '')
                if ep1_date:
                    show_info['first_air_date'] = ep1_date

            confirm = input("\n確認在 TMDB 建立此節目? (y/n): ").strip().lower()
            if confirm != 'y':
                print("已取消。")
                return

            show_id = create_tmdb_show(driver, show_info)
            if not show_id:
                manual_id = input("\n無法自動取得節目 ID。請手動填寫 TMDB 節目 ID (或按 Enter 跳過): ").strip()
                if not manual_id:
                    return
                show_id = manual_id

            print(f"\n節目已建立，TMDB ID: {show_id}")
            print(f"節目頁面: https://www.themoviedb.org/tv/{show_id}")

            episodes = show_info.get('episodes', [])
            episodes_to_add = [e for e in episodes if e['ep_num'] > 1]
            if episodes_to_add:
                print(f"\n正在新增第 2–{max(e['ep_num'] for e in episodes_to_add)} 集...")
                add_episodes_to_season(driver, show_id, 1, episodes, air_dates, tmdb_lang)
            else:
                log.info("只有 1 集，無需批量新增。")

            # All episodes are in season 1; viu ep numbers map 1:1 to TMDB ep numbers
            ep_mapping = [(ep['ep_num'], 1, ep['ep_num']) for ep in episodes]

        else:
            # ── Multi-season ─────────────────────────────────────────────────
            season_ranges, renumber = _ask_season_ranges(
                show_info['episodes'], num_seasons=num_seasons_m1
            )
            episodes_by_num = {ep['ep_num']: ep for ep in show_info['episodes']}

            # Ask for Wikipedia air dates once; reused (filtered) per season
            print("\n=== 播映日期 (所有季) ===")
            wiki_url_m1 = input("Wikipedia 頁面 URL (留空逐季手動設定): ").strip()
            wiki_dates_m1 = _fetch_wiki_dates(wiki_url_m1) if wiki_url_m1 else None
            if wiki_dates_m1:
                print(f"從 Wikipedia 成功抓取 {len(wiki_dates_m1)} 集播映日期。")

            # Season 1 episodes & air dates are needed by the wizard before show creation
            _s1_num, s1_viu_start, s1_viu_end = season_ranges[0]
            s1_eps_raw = [
                episodes_by_num[n]
                for n in sorted(episodes_by_num)
                if s1_viu_start <= n <= s1_viu_end
            ]
            if renumber:
                s1_eps = [{**ep, 'ep_num': ep['ep_num'] - (s1_viu_start - 1)} for ep in s1_eps_raw]
            else:
                s1_eps = list(s1_eps_raw)

            print(f"\n=== 第 1 季 ({len(s1_eps)} 集: viu.tv {s1_viu_start}–{s1_viu_end}) ===")
            s1_air_dates = _ask_air_dates(s1_eps, cached_wiki_dates=wiki_dates_m1)
            ep1_date = s1_air_dates.get(1, '')
            if ep1_date:
                show_info['first_air_date'] = ep1_date

            # Wizard only needs S1 episode data
            show_info['episodes'] = s1_eps

            confirm = input("\n確認在 TMDB 建立此節目? (y/n): ").strip().lower()
            if confirm != 'y':
                print("已取消。")
                return

            show_id = create_tmdb_show(driver, show_info)
            if not show_id:
                manual_id = input("\n無法自動取得節目 ID。請手動填寫 TMDB 節目 ID (或按 Enter 跳過): ").strip()
                if not manual_id:
                    return
                show_id = manual_id

            print(f"\n節目已建立，TMDB ID: {show_id}")
            print(f"節目頁面: https://www.themoviedb.org/tv/{show_id}")

            # Build viu_ep → (tmdb_season, tmdb_ep) mapping while adding episodes
            ep_mapping = []
            for season_num, viu_start, viu_end in season_ranges:
                season_eps_raw = [
                    episodes_by_num[n]
                    for n in sorted(episodes_by_num)
                    if viu_start <= n <= viu_end
                ]
                if not season_eps_raw:
                    log.warning(f"第 {season_num} 季: 找不到 viu.tv 集數 {viu_start}–{viu_end}，跳過。")
                    continue

                if renumber:
                    offset = viu_start - 1
                    remapped = [{**ep, 'ep_num': ep['ep_num'] - offset} for ep in season_eps_raw]
                    ep_range_label = f"TMDB E1–E{len(remapped)}"
                else:
                    offset = 0
                    remapped = list(season_eps_raw)
                    ep_range_label = f"TMDB E{viu_start}–E{viu_end}"

                print(f"\n=== 第 {season_num} 季 ({len(remapped)} 集: viu.tv {viu_start}–{viu_end} → {ep_range_label}) ===")
                if season_num == 1:
                    season_air_dates = s1_air_dates
                else:
                    season_air_dates = _ask_air_dates(remapped, cached_wiki_dates=wiki_dates_m1)
                    _ensure_season_exists(driver, show_id, season_num)

                add_episodes_to_season(driver, show_id, season_num, remapped, season_air_dates, tmdb_lang)

                for viu_ep in range(viu_start, viu_end + 1):
                    ep_mapping.append((viu_ep, season_num, viu_ep - offset))

        # Upload episode backdrops (only for viu.tv; viu.com uses a different page structure)
        if not is_viu_com:
            cached = {ep['ep_num']: ep['image_url']
                      for ep in all_episodes_m1 if ep.get('image_url')}
            _upload_episode_backdrops(
                driver, show_id, viutv_url, is_viu_com, ep_mapping, tmdb_lang, temp_dir,
                cached_images=cached or None,
            )

        _upload_images(driver, show_id, show_info, temp_dir, tmdb_lang, auto_poster=True)

        print(f"\n完成！請前往以下頁面補充其他資料:")
        print(f"  編輯頁面 : https://www.themoviedb.org/tv/{show_id}/edit")
        print(f"  圖片頁面 : https://www.themoviedb.org/tv/{show_id}/images/posters")

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if driver:
            cleanup_browser(driver)


if __name__ == '__main__':
    main()
