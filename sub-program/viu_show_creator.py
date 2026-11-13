"""
viu_show_creator.py
從 viu.com 節目頁面抓取資料，在 TMDB 建立新的電視劇條目，並上傳封面圖片。
用法: python sub-program/viu_show_creator.py
"""

import os
import sys
import shutil
import logging
import tempfile

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
sys.path.insert(0, project_root)
sys.path.insert(0, script_dir)  # allow importing viutv_show_creator from same directory

from others.chrome_process import get_browser_driver, cleanup_browser
from others.credentials import load_configs, get_tmdb_credentials, is_tmdb_logged_in, login

# All scraping and TMDB helpers (including viu.com functions) live in viutv_show_creator
from viutv_show_creator import (
    download_image,
    _wizard_click_next, _wizard_wait_for_step, _wizard_wait_for_any_step,
    _fill_visible_field, _fill_runtime_field, _fill_form_field, _fill_kendo_date_picker,
    _set_kendo_episode_number, _click_save_button,
    _scrape_wiki_episodes, _scrape_wiki_air_dates, _ask_air_dates, _ask_broadcast_schedule,
    _read_grid_episodes, _is_placeholder_name, _edit_episode_in_grid,
    add_episodes_to_season, create_tmdb_show, _extract_show_id_from_url,
    _pause_for_manual, _has_existing_images, _upload_to_tmdb_image_page,
    upload_show_poster, upload_show_backdrop, _parse_show_id,
    _upload_images, _check_season_exists, _ask_season_ranges,
    scrape_viu_show, _update_season_translation,
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

def _print_show_summary(show_info: dict):
    """Print a summary of scraped viu.com data."""
    seasons = sorted({ep['season'] for ep in show_info['episodes']})
    ep1 = next(
        (e for e in show_info['episodes'] if e['season'] == 1 and e['ep_num'] == 1),
        None
    )
    print(f"\n=== 從 viu.com 取得的資料 ===")
    print(f"標題     : {show_info['title']}")
    desc = show_info['description']
    print(f"描述     : {desc[:80]}..." if len(desc) > 80 else f"描述     : {desc}")
    print(f"季數     : {show_info.get('num_seasons', len(seasons))}")
    for s in seasons:
        s_eps = [ep for ep in show_info['episodes'] if ep['season'] == s]
        print(f"  第 {s} 季: {len(s_eps)} 集 (E{s_eps[0]['ep_num']}–E{s_eps[-1]['ep_num']})")
    if ep1:
        ep1_desc = ep1.get('description', '')
        print(f"S01E01   : {ep1['title']}")
        print(f"S01E01描述: {ep1_desc[:60]}..." if len(ep1_desc) > 60 else f"S01E01描述: {ep1_desc}")
    print(f"封面 URL : {show_info['poster_url']}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=== viu.com 節目 TMDB 管理工具 ===\n")
    print("模式選擇:")
    print("  1. 建立全新節目 (從 viu.com 抓取並在 TMDB 建立新條目)")
    print("  2. 為現有節目加入新一季集數")
    print("  3. 為現有節目分多季加入集數 (viu.com 單季 → TMDB 多季)")
    mode = input("選擇 (1-3，默認 1): ").strip() or '1'

    if mode == '1':
        viu_url = input("viu.com 節目 URL: ").strip()
        tmdb_lang = input("TMDB 上傳語言標籤 (默認 zh-HK): ").strip() or 'zh-HK'
        existing_show_id = None

    elif mode == '2':
        viu_url = input("viu.com 節目 URL: ").strip()
        raw_id = input("TMDB 節目 ID 或 URL: ").strip()
        existing_show_id = _parse_show_id(raw_id)
        if not existing_show_id:
            print(f"無法解析 TMDB 節目 ID: {raw_id!r}，程式終止。")
            return
        tmdb_lang = input("TMDB 上傳語言標籤 (默認 zh-HK): ").strip() or 'zh-HK'

    elif mode == '3':
        viu_url = input("viu.com 節目 URL: ").strip()
        raw_id = input("TMDB 節目 ID 或 URL: ").strip()
        existing_show_id = _parse_show_id(raw_id)
        if not existing_show_id:
            print(f"無法解析 TMDB 節目 ID: {raw_id!r}，程式終止。")
            return
        tmdb_lang = input("TMDB 上傳語言標籤 (默認 zh-HK): ").strip() or 'zh-HK'

    else:
        print(f"不認識的模式: {mode}")
        return

    temp_dir = tempfile.mkdtemp(prefix='viu_creator_')
    log.info(f"臨時目錄: {temp_dir}")

    driver = None
    try:
        configs = load_configs()
        driver_tuple = get_browser_driver(profile_suffix='_viu_show', remote_port=9227, base_dir=project_root)
        driver = driver_tuple[0]

        if not is_tmdb_logged_in(driver):
            username, password = get_tmdb_credentials(configs)
            if username and password:
                login(driver, username, password)
                log.info("已登入 TMDB。")
            else:
                log.warning("未找到 TMDB 憑證，請先在瀏覽器中手動登入。")
                input("登入後請按 Enter 繼續...")

        show_info = scrape_viu_show(driver, viu_url)
        if not show_info['title']:
            log.error("無法取得節目標題，程式終止。")
            return

        _print_show_summary(show_info)

        # ── Mode 2: Add a specific viu.com season to an existing TMDB show ──
        if mode == '2':
            available_seasons = sorted({ep['season'] for ep in show_info['episodes']})
            if len(available_seasons) > 1:
                print(f"\nviu.com 共有 {len(available_seasons)} 季: {available_seasons}")
                raw_s = input("使用哪一季的集數? (默認 1): ").strip() or '1'
                viu_season = int(raw_s) if raw_s.isdigit() else 1
            else:
                viu_season = 1

            season_episodes = [ep for ep in show_info['episodes'] if ep['season'] == viu_season]
            if not season_episodes:
                print(f"viu.com 第 {viu_season} 季找不到任何集數，程式終止。")
                return

            tmdb_season = int(input("上傳至 TMDB 哪一季? (默認 2): ").strip() or '2')

            air_dates = _ask_air_dates(season_episodes)

            confirm = input(
                f"\n確認為 TMDB 節目 {existing_show_id} 新增第 {tmdb_season} 季"
                f" (來自 viu.com 第 {viu_season} 季，共 {len(season_episodes)} 集)? (y/n): "
            ).strip().lower()
            if confirm != 'y':
                print("已取消。")
                return

            if not _check_season_exists(driver, existing_show_id, tmdb_season):
                print(f"\n第 {tmdb_season} 季在 TMDB 上尚未存在，或頁面無法載入。")
                print(f"請先在 TMDB 手動建立第 {tmdb_season} 季：")
                print(f"  https://www.themoviedb.org/tv/{existing_show_id}/season/new")
                input("建立後按 Enter 繼續...")

            # For season 2+: viu.com show title/description → TMDB season name/description
            if tmdb_season >= 2 and (show_info['title'] or show_info['description']):
                _update_season_translation(
                    driver, existing_show_id, tmdb_season,
                    show_info['title'], show_info['description'], tmdb_lang
                )

            add_episodes_to_season(
                driver, existing_show_id, tmdb_season,
                season_episodes, air_dates, tmdb_lang
            )
            _upload_images(driver, existing_show_id, show_info, temp_dir, tmdb_lang)

            print(f"\n完成！請前往以下頁面確認資料:")
            print(f"  https://www.themoviedb.org/tv/{existing_show_id}/season/{tmdb_season}")
            return

        # ── Mode 3: Single viu.com season → multiple TMDB seasons ────────────
        if mode == '3':
            all_eps = show_info['episodes']
            available_seasons = sorted({ep['season'] for ep in all_eps})
            if len(available_seasons) > 1:
                print(f"\nviu.com 共有 {len(available_seasons)} 季: {available_seasons}")
                raw_s = input("使用哪一季的集數進行分割? (默認 1): ").strip() or '1'
                viu_season = int(raw_s) if raw_s.isdigit() else 1
                flat_episodes = [ep for ep in all_eps if ep['season'] == viu_season]
            else:
                flat_episodes = all_eps

            season_ranges, renumber = _ask_season_ranges(flat_episodes)
            episodes_by_num = {ep['ep_num']: ep for ep in flat_episodes}

            print("\n=== 播映日期 (所有季) ===")
            wiki_url_m3 = input("Wikipedia 頁面 URL (留空逐季手動設定): ").strip()
            wiki_dates_m3 = _scrape_wiki_air_dates(wiki_url_m3) if wiki_url_m3 else None
            if wiki_dates_m3:
                print(f"從 Wikipedia 成功抓取 {len(wiki_dates_m3)} 集播映日期。")

            for season_num, viu_start, viu_end in season_ranges:
                season_eps_raw = [
                    episodes_by_num[n]
                    for n in sorted(episodes_by_num)
                    if viu_start <= n <= viu_end
                ]
                if not season_eps_raw:
                    log.warning(f"第 {season_num} 季: 集數 {viu_start}–{viu_end} 範圍內找不到集數，跳過。")
                    continue

                if renumber:
                    offset = viu_start - 1
                    remapped = [{**ep, 'ep_num': ep['ep_num'] - offset} for ep in season_eps_raw]
                    ep_range_label = f"TMDB E1–E{len(remapped)}"
                else:
                    remapped = list(season_eps_raw)
                    ep_range_label = f"TMDB E{viu_start}–E{viu_end}"

                print(f"\n=== 第 {season_num} 季 ({len(remapped)} 集: viu.com {viu_start}–{viu_end} → {ep_range_label}) ===")
                season_air_dates = _ask_air_dates(remapped, cached_wiki_dates=wiki_dates_m3)

                if not _check_season_exists(driver, existing_show_id, season_num):
                    print(f"\n第 {season_num} 季在 TMDB 上尚未存在，請先建立：")
                    print(f"  https://www.themoviedb.org/tv/{existing_show_id}/season/new")
                    input("建立後按 Enter 繼續...")

                add_episodes_to_season(
                    driver, existing_show_id, season_num,
                    remapped, season_air_dates, tmdb_lang
                )

            print(f"\n完成！請前往節目頁面確認資料:")
            print(f"  https://www.themoviedb.org/tv/{existing_show_id}")
            return

        # ── Mode 1: Create new show ──────────────────────────────────────────
        available_seasons = sorted({ep['season'] for ep in show_info['episodes']})
        s1_episodes = [ep for ep in show_info['episodes'] if ep['season'] == 1]
        if not s1_episodes:
            log.error("找不到第 1 季集數，程式終止。")
            return

        air_dates_by_season: dict = {}
        if show_info.get('num_seasons', 1) > 1:
            print(f"\nviu.com 共有 {show_info['num_seasons']} 季，建立節目後將依序新增各季集數。")

        print(f"\n=== 第 1 季播映日期 ===")
        air_dates_by_season[1] = _ask_air_dates(
            s1_episodes, default_start=show_info.get('first_air_date', '')
        )
        if air_dates_by_season[1]:
            ep1_date = air_dates_by_season[1].get(s1_episodes[0]['ep_num'], '')
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

        add_episodes_to_season(driver, show_id, 1, s1_episodes, air_dates_by_season[1], tmdb_lang)

        for s in available_seasons:
            if s == 1:
                continue
            s_episodes = [ep for ep in show_info['episodes'] if ep['season'] == s]
            print(f"\n=== 第 {s} 季 ({len(s_episodes)} 集) ===")
            print(f"\n請在 TMDB 手動建立第 {s} 季：")
            print(f"  https://www.themoviedb.org/tv/{show_id}/season/new")
            input("建立後按 Enter 繼續...")

            if show_info['title'] or show_info['description']:
                _update_season_translation(
                    driver, show_id, s,
                    show_info['title'], show_info['description'], tmdb_lang
                )

            air_dates_by_season[s] = _ask_air_dates(s_episodes)
            add_episodes_to_season(driver, show_id, s, s_episodes, air_dates_by_season[s], tmdb_lang)

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
