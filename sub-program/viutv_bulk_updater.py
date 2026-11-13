"""
viutv_bulk_updater.py
批量更新 viu.tv 節目的 TMDB 集數，從 Excel 設定檔讀取配置。
用法:
  python sub-program/viutv_bulk_updater.py [excel_file]
  python sub-program/viutv_bulk_updater.py --template [output_path]

Excel 欄位說明:
  A  show_name         - 節目名稱 (僅供參考)
  B  viutv_url         - viu.tv 節目頁面 URL
  C  tmdb_id           - TMDB 節目 ID (數字或完整 URL)
  D  season_num        - TMDB 季數 (默認 1)
  E  tmdb_lang         - TMDB 上傳語言 (默認 zh-HK)
  F  last_ep           - 已更新到的集數 (程式更新後自動記錄)
  G  air_date_mode     - 播出模式: daily / weekdays / mon-sat / weekly / wiki / none
  H  air_date_weekdays - 播出星期 (週三週五=3,5；僅 weekly 模式使用)
  I  air_date_start    - 下一集播出日期 YYYY-MM-DD (更新後自動推進)
  J  wiki_url          - Wikipedia 集數頁面 URL (僅 wiki 模式使用)
  K  enabled           - 是否啟用此節目 (TRUE/FALSE，默認 TRUE)
"""

import os
import re
import sys
import time
import shutil
import logging
import tempfile

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
sys.path.insert(0, project_root)
sys.path.insert(0, script_dir)

from bs4 import BeautifulSoup
from datetime import date as _date, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
except ImportError:
    print("請先安裝 openpyxl: pip install openpyxl")
    sys.exit(1)

from others.chrome_process import get_browser_driver, cleanup_browser
from others.credentials import load_configs, get_tmdb_credentials, is_tmdb_logged_in, login

from viutv_show_creator import (
    add_episodes_to_season,
    _parse_show_id,
    _scrape_wiki_air_dates,
    _dates_from_titles,
    _collect_viutv_episode_cards,
    _parse_episode_cards,
    _read_grid_episodes,
    _upload_episode_backdrops,
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Excel column indices (1-based for openpyxl)
# ---------------------------------------------------------------------------
_COL_SHOW_NAME = 1
_COL_VIUTV_URL = 2
_COL_TMDB_ID = 3
_COL_SEASON_NUM = 4
_COL_TMDB_LANG = 5
_COL_LAST_EP = 6
_COL_AIR_MODE = 7
_COL_WEEKDAYS = 8
_COL_AIR_START = 9
_COL_WIKI_URL = 10
_COL_ENABLED = 11
_COL_BACKDROP_MODE = 12  # '' / 'normal' = upload all; 'skip_latest' = exclude newest ep

_HEADERS = [
    'show_name', 'viutv_url', 'tmdb_id', 'season_num', 'tmdb_lang',
    'last_ep', 'air_date_mode', 'air_date_weekdays', 'air_date_start',
    'wiki_url', 'enabled', 'backdrop_mode',
]
_VALID_AIR_MODES = ('daily', 'weekdays', 'mon-sat', 'weekly', 'wiki', 'title', 'none')


# ---------------------------------------------------------------------------
# Excel template creation
# ---------------------------------------------------------------------------

def _create_bulk_template(output_path: str) -> None:
    """Generate an empty Excel template with column headers and one example row."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'shows'

    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    center = Alignment(horizontal='center')

    for col_idx, header in enumerate(_HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    # Example data row
    ws.cell(row=2, column=_COL_SHOW_NAME, value='節目名稱')
    ws.cell(row=2, column=_COL_VIUTV_URL, value='https://www.viu.tv/encore/XXXXX')
    ws.cell(row=2, column=_COL_TMDB_ID, value='123456')
    ws.cell(row=2, column=_COL_SEASON_NUM, value=1)
    ws.cell(row=2, column=_COL_TMDB_LANG, value='zh-HK')
    ws.cell(row=2, column=_COL_LAST_EP, value=0)
    ws.cell(row=2, column=_COL_AIR_MODE, value='weekdays')
    ws.cell(row=2, column=_COL_WEEKDAYS, value='')
    ws.cell(row=2, column=_COL_AIR_START, value=_date(2026, 1, 1))
    ws.cell(row=2, column=_COL_WIKI_URL, value='')
    ws.cell(row=2, column=_COL_ENABLED, value=True)
    ws.cell(row=2, column=_COL_BACKDROP_MODE, value='')

    # Notes row
    ws.cell(row=3, column=_COL_AIR_MODE, value='daily/weekdays/mon-sat/weekly/wiki/none')
    ws.cell(row=3, column=_COL_WEEKDAYS, value='週三週五=3,5 (1=週一, 7=週日)')
    ws.cell(row=3, column=_COL_AIR_START, value='YYYY-MM-DD')
    ws.cell(row=3, column=_COL_BACKDROP_MODE, value='normal / skip_latest')

    col_widths = [20, 42, 12, 10, 10, 10, 14, 20, 14, 42, 10, 14]
    for col_idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[
            ws.cell(row=1, column=col_idx).column_letter
        ].width = width

    wb.save(output_path)
    print(f"已建立範本: {output_path}")


# ---------------------------------------------------------------------------
# Air-date computation
# ---------------------------------------------------------------------------

def _parse_weekdays_str(weekdays_str: str | None) -> set[int]:
    """Parse '3,5' to {2, 4} (0-indexed Monday)."""
    if not weekdays_str:
        return set()
    result = set()
    for part in str(weekdays_str).split(','):
        part = part.strip()
        if part.isdigit() and 1 <= int(part) <= 7:
            result.add(int(part) - 1)
    return result


def _mode_to_weekdays(mode: str, weekdays_str: str | None, start: _date) -> set[int] | None:
    """
    Return the broadcast weekday set for the given mode, or None if dates should not be set.
    """
    mode = (mode or 'none').lower().strip()
    if mode == 'daily':
        return set(range(7))
    if mode == 'weekdays':
        return {0, 1, 2, 3, 4}
    if mode == 'mon-sat':
        return {0, 1, 2, 3, 4, 5}
    if mode == 'weekly':
        wd = _parse_weekdays_str(weekdays_str)
        return wd if wd else {start.weekday()}
    return None  # 'wiki' (handled separately) or 'none'


def _compute_air_dates(episodes: list, mode: str, weekdays_str: str | None, start_str: str) -> dict[int, str]:
    """
    Non-interactive air date computation. Returns {ep_num: 'YYYY-MM-DD'}.
    Returns {} if mode is 'none'/'wiki' or start_str is missing.
    """
    if not start_str or not episodes:
        return {}
    try:
        start = _date.fromisoformat(str(start_str)[:10])
    except ValueError:
        log.warning(f"無效的 air_date_start: {start_str!r}")
        return {}

    weekdays = _mode_to_weekdays(mode, weekdays_str, start)
    if weekdays is None:
        return {}

    ep_nums = sorted(e['ep_num'] for e in episodes)
    air_dates: dict[int, str] = {}
    current = start
    while current.weekday() not in weekdays:
        current += timedelta(days=1)
    for ep_num in ep_nums:
        air_dates[ep_num] = current.isoformat()
        current += timedelta(days=1)
        while current.weekday() not in weekdays:
            current += timedelta(days=1)
    return air_dates


def _next_schedule_date(last_date_str: str, mode: str, weekdays_str: str | None) -> str:
    """Return the next scheduled broadcast date after last_date_str."""
    if not last_date_str:
        return last_date_str
    try:
        d = _date.fromisoformat(str(last_date_str)[:10])
    except ValueError:
        return last_date_str

    weekdays = _mode_to_weekdays(mode, weekdays_str, d)
    if weekdays is None:
        return last_date_str

    d += timedelta(days=1)
    while d.weekday() not in weekdays:
        d += timedelta(days=1)
    return d.isoformat()


# ---------------------------------------------------------------------------
# viu.tv episode scraping (accumulate while scrolling — list is virtualized)
# ---------------------------------------------------------------------------

def _check_and_scrape_new_episodes(driver, viutv_url: str, since_ep: int) -> list[dict] | None:
    """
    Load the viu.tv show page ONCE and return new episodes (ep_num > since_ep).

    - Returns None if the page failed to load.
    - Returns [] if the latest episode == since_ep (no update needed; TMDB is NOT visited).
    - Returns non-empty list of ep dicts if new episodes are found.

    Newest episodes are at the top, so the first paint already gives latest_ep.
    The list is virtualized: cards unmount as you scroll, and a 第1集 promo card
    can sit in the DOM the whole time. Accumulate episode numbers across scrolls
    until since_ep+1..latest_ep is complete — do not stop on min(visible) <= since_ep.
    """
    log.info(f"載入 viu.tv 頁面: {viutv_url}")
    driver.get(viutv_url)
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'h1')))
    except Exception:
        pass
    time.sleep(3)
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a.card--episode'))
        )
    except Exception:
        log.warning("等待集數卡片超時。")
    time.sleep(1)

    soup_init = BeautifulSoup(driver.page_source, 'html.parser')
    init_eps = _parse_episode_cards(soup_init)
    if not init_eps:
        log.warning("頁面未找到任何集數卡片。")
        return None

    latest_ep = max(init_eps)
    log.info(f"viu.tv 最新集數 (初始載入): {latest_ep}，last_ep: {since_ep}")

    if latest_ep <= since_ep:
        return []

    collected = _collect_viutv_episode_cards(
        driver, needed_from=since_ep + 1, needed_to=latest_ep,
    )
    new_eps = {k: v for k, v in collected.items() if k > since_ep}
    missing = [n for n in range(since_ep + 1, latest_ep + 1) if n not in new_eps]
    log.info(f"發現新集數: {sorted(new_eps.keys())}")
    if missing:
        log.warning(f"目標範圍內未抓到的集數: {missing}")
    return [new_eps[n] for n in sorted(new_eps)]


# ---------------------------------------------------------------------------
# Excel config loading / saving
# ---------------------------------------------------------------------------

def _cell_str(row, col_idx: int) -> str:
    val = row[col_idx - 1].value
    return str(val).strip() if val is not None else ''


def _cell_int(row, col_idx: int, default: int = 0) -> int:
    val = row[col_idx - 1].value
    try:
        return int(val) if val is not None else default
    except (ValueError, TypeError):
        return default


def _load_config_rows(excel_path: str):
    """
    Load the Excel config file.
    Returns (wb, ws, list_of_config_dicts).
    Skips rows where viutv_url or tmdb_id is empty, or enabled=FALSE.
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        if len(row) < _COL_ENABLED:
            continue

        viutv_url = _cell_str(row, _COL_VIUTV_URL)
        tmdb_id_raw = _cell_str(row, _COL_TMDB_ID)
        if not viutv_url or not tmdb_id_raw:
            continue

        enabled_val = _cell_str(row, _COL_ENABLED)
        if enabled_val.upper() == 'FALSE':
            show_name = _cell_str(row, _COL_SHOW_NAME)
            log.info(f"[{show_name}] 已停用，跳過。")
            continue

        tmdb_id = _parse_show_id(tmdb_id_raw)
        show_name = _cell_str(row, _COL_SHOW_NAME)
        if not tmdb_id:
            log.warning(f"[{show_name}] 無法解析 TMDB ID: {tmdb_id_raw!r}，跳過。")
            continue

        air_mode = _cell_str(row, _COL_AIR_MODE) or 'none'
        weekdays_str = _cell_str(row, _COL_WEEKDAYS) or None
        air_start = _cell_str(row, _COL_AIR_START)

        # air_date_start may be stored as a date object by openpyxl
        raw_start_cell = row[_COL_AIR_START - 1].value
        if raw_start_cell and hasattr(raw_start_cell, 'strftime'):
            air_start = raw_start_cell.strftime('%Y-%m-%d')

        rows.append({
            'row_idx': row_idx,
            'show_name': show_name,
            'viutv_url': viutv_url,
            'tmdb_id': tmdb_id,
            'season_num': _cell_int(row, _COL_SEASON_NUM, default=1) or 1,
            'tmdb_lang': _cell_str(row, _COL_TMDB_LANG) or 'zh-HK',
            'last_ep': _cell_int(row, _COL_LAST_EP, default=0),
            'air_mode': air_mode.lower().strip(),
            'weekdays_str': weekdays_str,
            'air_start': air_start,
            'wiki_url': _cell_str(row, _COL_WIKI_URL),
            'backdrop_mode': (
                _cell_str(row, _COL_BACKDROP_MODE).lower().strip()
                if len(row) >= _COL_BACKDROP_MODE else ''
            ),
        })
    return wb, ws, rows


def _save_row(ws, row_idx: int, new_last_ep: int, new_air_start: str) -> None:
    """Update last_ep and air_date_start columns for a specific row."""
    ws.cell(row=row_idx, column=_COL_LAST_EP, value=new_last_ep)
    # Write as a date object so Excel always stores it as a date type,
    # preventing display-format strings like 09/06/2026 from being misread.
    try:
        date_val = _date.fromisoformat(new_air_start)
    except (ValueError, TypeError):
        date_val = new_air_start  # fallback: write as-is
    ws.cell(row=row_idx, column=_COL_AIR_START, value=date_val)


def _migrate_columns(wb, ws, excel_path: str) -> bool:
    """
    Check the header row and add any columns that are defined in _HEADERS but
    missing from the sheet. Returns True if the file was modified (and saved).
    Existing data rows are left untouched; new columns get an empty value.
    """
    from openpyxl.styles import PatternFill, Font, Alignment
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    center = Alignment(horizontal='center')

    # Read existing headers (row 1)
    existing = {ws.cell(row=1, column=c).value: c
                for c in range(1, ws.max_column + 1)}
    added = []
    for col_idx, name in enumerate(_HEADERS, 1):
        if name not in existing:
            cell = ws.cell(row=1, column=col_idx, value=name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            added.append(name)

    if added:
        log.info(f"設定檔補充欄位: {added}")
        wb.save(excel_path)
        print(f"已自動補充缺少的欄位: {added}")
        return True
    return False


# ---------------------------------------------------------------------------
# Main bulk update loop
# ---------------------------------------------------------------------------

def run_bulk_update(excel_path: str) -> None:
    if not os.path.exists(excel_path):
        print(f"找不到設定檔: {excel_path}")
        return

    wb, ws, config_rows = _load_config_rows(excel_path)
    _migrate_columns(wb, ws, excel_path)
    if not config_rows:
        print("設定檔中沒有有效的節目設定。")
        return

    print(f"\n已載入 {len(config_rows)} 個節目設定")
    temp_dir = tempfile.mkdtemp(prefix='viutv_bulk_')
    log.info(f"臨時目錄: {temp_dir}")

    any_updated = False
    driver = None
    try:
        configs = load_configs()
        driver_tuple = get_browser_driver(profile_suffix='_viutv_show', remote_port=9225, base_dir=project_root)
        driver = driver_tuple[0]

        if not is_tmdb_logged_in(driver):
            username, password = get_tmdb_credentials(configs)
            if username and password:
                login(driver, username, password)
                log.info("已登入 TMDB。")
            else:
                log.warning("未找到 TMDB 憑證，請先在瀏覽器中手動登入。")
                input("登入後請按 Enter 繼續...")

        for config in config_rows:
            show_name = config['show_name']
            viutv_url = config['viutv_url']
            tmdb_id = config['tmdb_id']
            season_num = config['season_num']
            tmdb_lang = config['tmdb_lang']
            last_ep = config['last_ep']
            air_mode = config['air_mode']
            weekdays_str = config['weekdays_str']
            air_start = config['air_start']
            wiki_url = config['wiki_url']
            backdrop_mode = config.get('backdrop_mode', '') or 'normal'
            row_idx = config['row_idx']

            print(f"\n{'='*60}")
            print(f"節目: {show_name}  (TMDB {tmdb_id}, 第 {season_num} 季)")
            print(f"最後更新集數: {last_ep}  播出模式: {air_mode}  下次起始日: {air_start}")

            # ── Step 1 & 2: Load viu.tv once, check & scrape new episodes ────
            new_episodes = _check_and_scrape_new_episodes(driver, viutv_url, since_ep=last_ep)
            if new_episodes is None:
                log.warning(f"[{show_name}] 無法載入 viu.tv 頁面，跳過。")
                continue
            if not new_episodes:
                print(f"無新集數 (last_ep={last_ep})，跳過。")
                continue

            actual_new_nums = sorted(ep['ep_num'] for ep in new_episodes)
            print(f"發現新集數: {actual_new_nums}")

            # ── Step 3: Air dates ────────────────────────────────────────────
            ep_num_set = {ep['ep_num'] for ep in new_episodes}
            air_dates: dict[int, str] = {}

            if air_mode == 'title':
                title_dates = _dates_from_titles(new_episodes)
                if title_dates:
                    air_dates = title_dates
                    print(f"從集數標題取得 {len(air_dates)}/{len(ep_num_set)} 集播映日期。")
                else:
                    log.warning(f"[{show_name}] title 模式：集數標題不是日期格式，無法取得播映日期。")
            elif air_mode == 'wiki' and wiki_url:
                try:
                    wiki_dates, year_missing = _scrape_wiki_air_dates(wiki_url)
                    if year_missing and air_start:
                        try:
                            fallback_year = int(air_start[:4])
                            wiki_dates, _ = _scrape_wiki_air_dates(wiki_url, fallback_year=fallback_year)
                            log.info(f"[{show_name}] Wiki 缺少年份，以 air_date_start 年份 {fallback_year} 補填。")
                        except Exception as ye:
                            log.warning(f"[{show_name}] Wiki 年份補填失敗: {ye}")
                    elif year_missing:
                        log.warning(f"[{show_name}] Wiki 缺少年份且 air_date_start 未填，播映日期可能不完整。")
                    air_dates = {k: v for k, v in wiki_dates.items() if k in ep_num_set}
                    print(f"從 Wikipedia 取得 {len(air_dates)}/{len(ep_num_set)} 集播映日期。")
                except Exception as e:
                    log.warning(f"[{show_name}] Wikipedia 抓取失敗: {e}")
            else:
                air_dates = _compute_air_dates(new_episodes, air_mode, weekdays_str, air_start)

            if air_dates:
                dates_sorted = sorted(air_dates.values())
                print(f"播映日期: {dates_sorted[0]} – {dates_sorted[-1]}")
            else:
                print("未設定播映日期。")

            # ── Step 4: Filter by TMDB season range ──────────────────────────
            season_url = (
                f"https://www.themoviedb.org/tv/{tmdb_id}/season/{season_num}/edit"
                f"?active_nav_item=episodes"
            )
            log.info(f"預讀 TMDB grid: {season_url}")
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
                before = len(new_episodes)
                new_episodes = [ep for ep in new_episodes if ep['ep_num'] >= min_grid_ep]
                skipped = before - len(new_episodes)
                if skipped:
                    print(f"過濾掉 {skipped} 集屬於其他季的集數 (第 {season_num} 季從 E{min_grid_ep} 開始)。")

            if not new_episodes:
                log.warning(f"[{show_name}] 過濾後無集數需要更新，跳過。")
                continue

            # ── Step 5: Update TMDB ──────────────────────────────────────────
            newly_added, edited = add_episodes_to_season(
                driver, tmdb_id, season_num, new_episodes, air_dates, tmdb_lang
            )
            modified = newly_added | edited
            if newly_added:
                print(f"已新增 {len(newly_added)} 集: {sorted(newly_added)}")
            if edited:
                print(f"已補填 {len(edited)} 集: {sorted(edited)}")

            # ── Step 5b: Upload episode backdrops (newly added + filled placeholders)
            if modified:
                if backdrop_mode == 'skip_latest' and new_episodes:
                    # viu.tv only has the actual episode image once the NEXT ep airs.
                    # Exclude the highest scraped episode number from backdrop upload.
                    max_scraped = max(ep['ep_num'] for ep in new_episodes)
                    backdrop_nums = {n for n in modified if n < max_scraped}
                    if backdrop_nums:
                        log.info(f"skip_latest 模式: 上傳 {sorted(backdrop_nums)} 集劇照 (跳過最新 E{max_scraped})")
                    else:
                        log.info(f"skip_latest 模式: 最新集數 E{max_scraped} 為唯一修改集數，不上傳劇照")
                else:
                    backdrop_nums = modified

                if backdrop_nums:
                    ep_mapping = [(n, season_num, n) for n in sorted(backdrop_nums)]
                    cached = {ep['ep_num']: ep['image_url']
                              for ep in new_episodes if ep.get('image_url')}
                    _upload_episode_backdrops(
                        driver, tmdb_id, viutv_url, False,
                        ep_mapping, tmdb_lang, temp_dir,
                        cached_images=cached or None,
                        auto_upload=True,
                    )

            # ── Step 6: Update Excel ─────────────────────────────────────────
            if modified:
                max_added_ep = max(modified)
                last_air_date = air_dates.get(max_added_ep, air_start)
                new_air_start = _next_schedule_date(last_air_date, air_mode, weekdays_str)
                _save_row(ws, row_idx, max_added_ep, new_air_start)
                print(f"已更新設定: last_ep={max_added_ep}, 下次起始日={new_air_start}")
                any_updated = True
            else:
                print("所有集數已存在，無需更新。")

        if any_updated:
            wb.save(excel_path)
            print(f"\n設定檔已儲存: {excel_path}")
        else:
            print("\n所有節目均無需更新，設定檔未修改。")

    finally:
        if driver:
            cleanup_browser(driver)
        shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Excel validation (--test mode)
# ---------------------------------------------------------------------------

def _fmt_row(row_idx: int, show_name: str) -> str:
    return f"第 {row_idx} 行 [{show_name or '未命名'}]"


def validate_excel(excel_path: str) -> bool:
    """
    Validate the Excel config file without starting the browser.
    Checks all rows for missing / invalid fields and prints a report.
    Returns True if no errors were found (warnings are still OK).
    """
    if not os.path.exists(excel_path):
        print(f"[錯誤] 找不到檔案: {excel_path}")
        return False

    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as e:
        print(f"[錯誤] 無法開啟 Excel 檔案: {e}")
        return False

    ws = wb.active
    errors: list[str] = []
    warnings: list[str] = []
    enabled_count = 0
    disabled_count = 0

    # ── Header check ────────────────────────────────────────────────────────
    header_row = [ws.cell(row=1, column=c).value for c in range(1, _COL_ENABLED + 1)]
    expected = _HEADERS
    header_ok = True
    for i, (got, exp) in enumerate(zip(header_row, expected), 1):
        if str(got or '').strip().lower() != exp.lower():
            errors.append(f"欄 {i} 標題應為 '{exp}'，實際為 '{got}'")
            header_ok = False
    if not header_ok:
        print("標題列有誤，以下驗證結果僅供參考。\n")

    # ── Data rows ───────────────────────────────────────────────────────────
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        if len(row) < _COL_ENABLED:
            continue

        # Skip completely empty rows
        if all(row[c - 1].value is None for c in range(1, _COL_ENABLED + 1)):
            continue

        show_name = _cell_str(row, _COL_SHOW_NAME)
        label = _fmt_row(row_idx, show_name)

        # enabled
        enabled_val = _cell_str(row, _COL_ENABLED)
        if enabled_val.upper() == 'FALSE':
            disabled_count += 1
            continue
        enabled_count += 1

        # viutv_url
        viutv_url = _cell_str(row, _COL_VIUTV_URL)
        if not viutv_url:
            errors.append(f"{label}: viutv_url 為空")
        elif not re.search(r'viu\.(tv|com)', viutv_url):
            warnings.append(f"{label}: viutv_url 不像 viu.tv/viu.com URL → {viutv_url!r}")

        # tmdb_id
        tmdb_id_raw = _cell_str(row, _COL_TMDB_ID)
        if not tmdb_id_raw:
            errors.append(f"{label}: tmdb_id 為空")
        elif not _parse_show_id(tmdb_id_raw):
            errors.append(f"{label}: tmdb_id 無法解析 → {tmdb_id_raw!r}")

        # season_num
        season_raw = row[_COL_SEASON_NUM - 1].value
        try:
            season_num = int(season_raw) if season_raw is not None else 1
            if season_num < 1:
                errors.append(f"{label}: season_num 必須 >= 1 (現為 {season_num})")
        except (ValueError, TypeError):
            errors.append(f"{label}: season_num 不是整數 → {season_raw!r}")

        # last_ep
        last_ep_raw = row[_COL_LAST_EP - 1].value
        try:
            last_ep = int(last_ep_raw) if last_ep_raw is not None else 0
            if last_ep < 0:
                errors.append(f"{label}: last_ep 必須 >= 0 (現為 {last_ep})")
        except (ValueError, TypeError):
            errors.append(f"{label}: last_ep 不是整數 → {last_ep_raw!r}")

        # air_date_mode
        air_mode = _cell_str(row, _COL_AIR_MODE).lower() or 'none'
        if air_mode not in _VALID_AIR_MODES:
            errors.append(
                f"{label}: air_date_mode 無效 → {air_mode!r}  "
                f"(有效值: {', '.join(_VALID_AIR_MODES)})"
            )

        # air_date_weekdays — required when mode == 'weekly'
        weekdays_str = _cell_str(row, _COL_WEEKDAYS) or None
        if air_mode == 'weekly':
            wd = _parse_weekdays_str(weekdays_str)
            if not wd:
                errors.append(
                    f"{label}: weekly 模式須填寫 air_date_weekdays "
                    f"(例如 '1,3,5')，現為 {weekdays_str!r}"
                )

        # air_date_start — required for schedule-based modes
        air_start_cell = row[_COL_AIR_START - 1].value
        if hasattr(air_start_cell, 'strftime'):
            air_start = air_start_cell.strftime('%Y-%m-%d')
        else:
            air_start = str(air_start_cell).strip() if air_start_cell is not None else ''

        needs_start = air_mode in ('daily', 'weekdays', 'mon-sat', 'weekly')
        if needs_start:
            if not air_start:
                warnings.append(f"{label}: {air_mode} 模式建議填寫 air_date_start")
            else:
                try:
                    _date.fromisoformat(air_start[:10])
                except ValueError:
                    errors.append(
                        f"{label}: air_date_start 格式錯誤 → {air_start!r}  (須為 YYYY-MM-DD)"
                    )

        # wiki_url — required when mode == 'wiki'
        wiki_url = _cell_str(row, _COL_WIKI_URL)
        if air_mode == 'wiki' and not wiki_url:
            errors.append(f"{label}: wiki 模式須填寫 wiki_url")

    # ── Report ───────────────────────────────────────────────────────────────
    print(f"檔案: {excel_path}")
    print(f"啟用節目: {enabled_count}  停用節目: {disabled_count}")
    print()

    if warnings:
        print(f"⚠  警告 ({len(warnings)} 項):")
        for w in warnings:
            print(f"   {w}")
        print()

    if errors:
        print(f"✗  錯誤 ({len(errors)} 項):")
        for e in errors:
            print(f"   {e}")
        print()
        print("請修正以上錯誤後再執行更新。")
        return False

    print("✓  驗證通過，無錯誤。")
    return True


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='批量更新 viu.tv 節目的 TMDB 集數',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('excel_path', nargs='?', help='Excel 設定檔路徑')
    parser.add_argument('--template', metavar='OUTPUT_PATH',
                        help='建立空白 Excel 範本並退出')
    parser.add_argument('--test', action='store_true',
                        help='驗證 Excel 設定檔有無錯誤，不執行更新')
    args = parser.parse_args()

    if args.template:
        _create_bulk_template(args.template)
        return

    if args.excel_path:
        excel_path = args.excel_path
    else:
        default = os.path.join(script_dir, 'viutv_bulk_config.xlsx')
        if os.path.exists(default):
            excel_path = default
        else:
            excel_path = input(f"Excel 設定檔路徑 (默認 {default}): ").strip() or default

    if args.test:
        validate_excel(excel_path)
        return

    run_bulk_update(excel_path)


if __name__ == '__main__':
    main()
