"""
viutv_image_uploader.py
從 viu.tv 提取每集劇照並上傳至 TMDB 對應集數的 Backdrops 頁面。
用法: python sub-program/viutv_image_uploader.py
"""

import os
import re
import sys
import time
import shutil
import logging
import tempfile
import requests
from PIL import Image

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
sys.path.insert(0, project_root)

from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from others.chrome_process import get_browser_driver, cleanup_browser
from others.credentials import load_configs, get_tmdb_credentials, is_tmdb_logged_in, login

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Image extraction from viu.com
# ---------------------------------------------------------------------------

def extract_episode_images_viu(driver, viu_url: str) -> dict[int, str]:
    """
    Load a viu.com show page and return {episode_number: image_url}.
    For multi-season pages the user is asked which season's images to extract.
    Image URLs have the resize query-string stripped for full resolution.
    """
    log.info(f"正在載入 viu.com 頁面: {viu_url}")
    driver.get(viu_url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'series_title'))
    )
    time.sleep(2)

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.series_ep'))
        )
    except Exception:
        log.warning("等待 .series_ep 超時，繼續嘗試解析...")

    ep_tabs = driver.find_elements(By.CSS_SELECTOR, 'button.ep_list_tab[role="tab"]')
    num_seasons = len(ep_tabs)

    def _parse_panel_images(panel) -> dict[int, str]:
        images: dict[int, str] = {}
        for item in panel.find_all(class_='series_ep'):
            ep_title_el = item.find(id='series_ep_title')
            if not ep_title_el:
                continue
            m = re.match(r'^(\d+)\.', ep_title_el.get_text(strip=True))
            if not m:
                continue
            ep_num = int(m.group(1))
            thumb = item.find('div', class_='grid_thumbnail')
            img_el = thumb.find('img') if thumb else item.find('img')
            if not img_el:
                continue
            src = img_el.get('src', '')
            if not src or 'placeholder' in src:
                continue
            # Strip resize query-string and CDN resize path segment for full resolution
            images[ep_num] = _strip_resize_path(src.split('?')[0])
        return images

    if num_seasons > 1:
        print(f"\nviu.com 共有 {num_seasons} 個季度分頁。")
        while True:
            raw = input(f"提取哪一季的圖片? (1-{num_seasons}，默認 1): ").strip() or '1'
            if raw.isdigit() and 1 <= int(raw) <= num_seasons:
                break
            print(f"請輸入 1 至 {num_seasons} 之間的數字。")
        tab = ep_tabs[int(raw) - 1]
        driver.execute_script("arguments[0].click();", tab)
        time.sleep(1.5)
        panel_id = tab.get_attribute('aria-controls')
        panel_soup = BeautifulSoup(driver.page_source, 'html.parser')
        panel = panel_soup.find(id=panel_id) if panel_id else panel_soup.find('div', attrs={'role': 'tabpanel'})
    else:
        panel_id = ep_tabs[0].get_attribute('aria-controls') if ep_tabs else None
        if not panel_id:
            tmp = BeautifulSoup(driver.page_source, 'html.parser')
            el = tmp.find('div', attrs={'role': 'tabpanel'})
            panel_id = el.get('id') if el else None
        panel_soup = BeautifulSoup(driver.page_source, 'html.parser')
        panel = panel_soup.find(id=panel_id) if panel_id else panel_soup.find('div', attrs={'role': 'tabpanel'})

    result: dict[int, str] = {}
    if panel:
        result = _parse_panel_images(panel)

    log.info(f"共找到 {len(result)} 集的圖片連結。")
    print(f"\n共找到 {len(result)} 集的圖片連結:")
    for ep, url in sorted(result.items()):
        log.info(f"  第 {ep:>3} 集: {url}")
        print(f"  第 {ep:>3} 集: {url}")
    return result


def _extract_show_backdrop_viu(driver, viu_url: str) -> str | None:
    """Extract the show's main image from a viu.com show page via og:image."""
    log.info(f"正在提取 viu.com 節目圖片: {viu_url}")
    driver.get(viu_url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'series_title'))
    )
    time.sleep(1)
    url = driver.execute_script(
        "var el = document.querySelector('meta[property=\"og:image\"]');"
        "return el ? el.getAttribute('content') : null;"
    )
    if url:
        url = _strip_resize_path(url)
        log.info(f"og:image: {url}")
    return url or None


# ---------------------------------------------------------------------------
# Image extraction from viu.tv
# ---------------------------------------------------------------------------

def _strip_resize_path(url: str) -> str:
    """Remove viu CDN resize segment (e.g. /resized-800/) to get full-resolution URL."""
    return re.sub(r'/resized-\d+/', '/', url)


def _largest_url_from_srcset(srcset: str) -> str | None:
    """Return the URL of the highest-resolution entry in a srcset attribute."""
    best_width, best_url = 0, None
    for part in srcset.split(','):
        part = part.strip()
        pieces = part.rsplit(' ', 1)
        if len(pieces) != 2:
            continue
        url = pieces[0].strip()
        try:
            width = int(pieces[1].rstrip('w'))
        except ValueError:
            continue
        if width > best_width:
            best_width, best_url = width, url
    return best_url


def _extract_show_backdrop(driver, viutv_url: str) -> str | None:
    """
    Extract the show's main backdrop/banner image URL from a viu.tv show page.
    Returns the highest-resolution URL found, or None.
    """
    log.info(f"正在提取節目橫幅圖片: {viutv_url}")
    driver.get(viutv_url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.TAG_NAME, 'h1'))
    )
    time.sleep(2)
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Try known intro/banner image elements
    for cls in ('intro__thumbnail', 'series-cover', 'series__thumbnail',
                'show-cover', 'banner__image', 'hero__image'):
        el = soup.find(class_=cls)
        if not el:
            continue
        tag = el if el.name == 'img' else el.find('img')
        if tag:
            srcset = tag.get('srcset', '')
            url = _largest_url_from_srcset(srcset) if srcset else tag.get('src', '')
            if url and 'placeholder' not in url:
                return url if not url.startswith('/') else 'https://www.viu.tv' + url
        # background-image style
        style = el.get('style', '')
        m = re.search(r'url\(["\']?([^"\')\s]+)["\']?\)', style)
        if m:
            return m.group(1)

    # Fallback: largest img inside the intro section
    intro = soup.find('div', class_='intro') or soup.find('section', class_='intro')
    if intro:
        best_width, best_url = 0, None
        for img in intro.find_all('img'):
            srcset = img.get('srcset', '')
            url = _largest_url_from_srcset(srcset) if srcset else img.get('src', '')
            if not url or 'placeholder' in url:
                continue
            if not url.startswith('http'):
                url = 'https://www.viu.tv' + url
            if srcset:
                for part in srcset.split(','):
                    pieces = part.strip().rsplit(' ', 1)
                    if len(pieces) == 2:
                        try:
                            w = int(pieces[1].rstrip('w'))
                            if w > best_width:
                                best_width, best_url = w, url
                        except ValueError:
                            pass
            elif not best_url:
                best_url = url
        if best_url:
            return best_url

    # Final fallback: og:image meta tag (strip resize segment for full resolution)
    og_url = driver.execute_script(
        "var el = document.querySelector('meta[property=\"og:image\"]');"
        "return el ? el.getAttribute('content') : null;"
    )
    if og_url:
        og_url = _strip_resize_path(og_url)
        log.info(f"og:image (fallback): {og_url}")
        return og_url

    return None


def extract_episode_images(driver, viutv_url: str) -> dict[int, str]:
    """
    Load a viu.tv show page and return {episode_number: image_url}.
    Deduplicates: regular episodes are preferred over data-prevent-progress
    special cuts (足本版 etc.).
    """
    log.info(f"正在載入 viu.tv 頁面: {viutv_url}")
    driver.get(viutv_url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.TAG_NAME, 'h1'))
    )
    time.sleep(2)

    sys.path.insert(0, script_dir)
    from viutv_show_creator import _collect_viutv_episode_cards
    collected = _collect_viutv_episode_cards(driver)
    result = {
        ep: d['image_url']
        for ep, d in sorted(collected.items())
        if d.get('image_url')
    }
    log.info(f"共找到 {len(result)} 集的圖片連結。")
    print(f"\n共找到 {len(result)} 集的圖片連結:")
    for ep, url in result.items():
        log.info(f"  第 {ep:>3} 集: {url}")
        print(f"  第 {ep:>3} 集: {url}")
    return result


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
# Image processing
# ---------------------------------------------------------------------------

_RATIO_16_9 = 16 / 9
_RATIO_TOLERANCE = 0  # TMDB strictly requires exactly 16:9; crop anything that deviates

def _ensure_jpg(image_path: str) -> str:
    """
    If image_path is not a JPEG, convert it to JPEG and return the new path.
    PNG images with transparency are flattened onto a white background.
    Returns the original path unchanged if it is already JPEG.
    """
    try:
        with Image.open(image_path) as img:
            if img.format == 'JPEG':
                return image_path
            rgb = img.convert('RGB') if img.mode != 'RGB' else img
            jpg_path = os.path.splitext(image_path)[0] + '.jpg'
            rgb.save(jpg_path, format='JPEG', quality=95)
            log.info(f"已將 {img.format} 轉換為 JPEG: {jpg_path}")
            if jpg_path != image_path:
                os.remove(image_path)
            return jpg_path
    except Exception as e:
        log.warning(f"圖片格式轉換失敗 ({image_path}): {e}")
        return image_path


def crop_to_16_9(image_path: str) -> bool:
    """
    Centre-crop the image to 16:9 if its aspect ratio deviates by more than
    _RATIO_TOLERANCE. Overwrites the file in place.
    Returns True if the image was cropped, False if it was already 16:9.
    """
    try:
        with Image.open(image_path) as img:
            w, h = img.size
            ratio = w / h
            if abs(ratio - _RATIO_16_9) <= _RATIO_TOLERANCE:
                log.info(f"圖片已是 16:9 ({w}x{h})，無需裁切。")
                return False

            if ratio > _RATIO_16_9:
                # Too wide — crop horizontally from centre
                new_w = int(h * _RATIO_16_9)
                left = (w - new_w) // 2
                box = (left, 0, left + new_w, h)
            else:
                # Too tall — crop vertically from centre
                new_h = int(w / _RATIO_16_9)
                top = (h - new_h) // 2
                box = (0, top, w, top + new_h)

            cropped = img.crop(box)
            fmt = img.format or 'JPEG'
            cropped.save(image_path, format=fmt)
            log.info(f"已裁切圖片至 16:9: {w}x{h} → {cropped.size[0]}x{cropped.size[1]} ({image_path})")
            return True
    except Exception as e:
        log.warning(f"裁切圖片失敗 ({image_path}): {e}")
        return False


# ---------------------------------------------------------------------------
# TMDB upload helpers (same flow as fzlins/TMDB-Import importors/episode.py)
# ---------------------------------------------------------------------------

_TMDB_IMAGE_PAGE_READY = (
    'h2, h3, div.image_content, li#no_results, a.add_image.upload, #upload_files'
)


def _has_existing_images(driver) -> bool:
    """Return True if the current TMDB image page already contains images.

    TMDB-Import skips when li#no_results is absent (i.e. images already exist).
    """
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, _TMDB_IMAGE_PAGE_READY))
        )
    except Exception:
        pass
    if driver.find_elements(By.CSS_SELECTOR, 'li#no_results'):
        return False
    return len(driver.find_elements(By.CSS_SELECTOR, 'div.image_content, li.card img')) > 0


def _dismiss_upload_dialog(driver) -> None:
    """TMDB-Import closes the cropper with Escape after a successful upload."""
    try:
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        time.sleep(0.5)
    except Exception:
        pass


def _upload_via_tmdb_cropper(driver, image_path: str, label: str) -> bool:
    """
    Upload one image using TMDB-Import's cropper flow.
    Assumes the driver is already on the images/backdrops or posters page.

      page.locator("a.add_image.upload").click()
      page.locator("input[id='upload_files'][accept='image/*']").set_input_files(path)
      page.wait_for_function("() => document.querySelector('span.file_status.success')")
      page.keyboard.press('Escape')
    """
    try:
        btn = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a.add_image.upload'))
        )
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(1.5)

        file_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input#upload_files[accept='image/*']")
            )
        )
        abs_path = os.path.abspath(image_path)
        try:
            file_input.send_keys(abs_path)
        except Exception:
            driver.execute_script(
                "var el=arguments[0]; el.style.display='block'; el.style.visibility='visible';"
                "el.style.opacity='1'; el.removeAttribute('tabindex');",
                file_input,
            )
            file_input.send_keys(abs_path)
        log.info(f"已選擇圖片: {abs_path}")

        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script(
                "return !!document.querySelector('span.file_status.success')"
            )
        )
        log.info(f"{label} 圖片上傳成功。")
        _dismiss_upload_dialog(driver)
        return True
    except Exception as e:
        msg = str(e).strip() or type(e).__name__
        err_text = ''
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, 'span.file_status'):
                if 'error' in (el.get_attribute('class') or ''):
                    err_text = (el.text or '').strip()
                    break
        except Exception:
            pass
        if err_text and err_text not in msg:
            msg = f"{msg} ({err_text})"
        log.error(f"{label} 圖片上傳失敗: {msg}")
        _dismiss_upload_dialog(driver)
        return False


def _upload_to_tmdb_image_page(
    driver,
    page_url: str,
    image_path: str,
    label: str,
    skip_if_exists: bool = True,
    upload_all_state: list | None = None,
) -> bool:
    """
    Generic TMDB image-page uploader (cropper UI).
    Returns True on success or when upload was skipped because images already exist.
    upload_all_state is a one-element list [bool] shared across calls; set to True
    when the user chooses 'a' (all) to suppress future prompts for this batch.
    """
    log.info(f"前往 TMDB 圖片頁面: {page_url}")
    driver.get(page_url)
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, _TMDB_IMAGE_PAGE_READY))
        )
    except Exception:
        pass
    time.sleep(1)

    if skip_if_exists and _has_existing_images(driver):
        if upload_all_state is not None and upload_all_state[0]:
            log.info(f"{label} 已有圖片，自動上載（all 模式）。")
        else:
            ans = input(f"{label} 已有圖片，是否仍要上載? (y=上載, a=全部上載, n=跳過，默認 n): ").strip().lower()
            if ans == 'a':
                if upload_all_state is not None:
                    upload_all_state[0] = True
            elif ans != 'y':
                log.info(f"{label} 已有圖片，用戶選擇跳過。")
                return True
        log.info(f"{label} 已有圖片，用戶選擇繼續上載。")

    return _upload_via_tmdb_cropper(driver, image_path, label)


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
    """Upload an episode backdrop; skips if the episode already has images."""
    url = (
        f"https://www.themoviedb.org/tv/{show_id}"
        f"/season/{season}/episode/{episode}"
        f"/images/backdrops?language={language}"
    )
    return _upload_to_tmdb_image_page(driver, url, image_path, f"第 {episode} 集劇照", skip_if_exists, upload_all_state)


def upload_show_poster(
    driver,
    show_id: str,
    image_path: str,
    language: str = 'zh-HK',
    skip_if_exists: bool = True,
) -> bool:
    """Upload a show poster; skips if the show already has a poster."""
    url = f"https://www.themoviedb.org/tv/{show_id}/images/posters?language={language}"
    return _upload_to_tmdb_image_page(driver, url, image_path, f"節目 {show_id} 封面", skip_if_exists)


def upload_show_backdrop(
    driver,
    show_id: str,
    image_path: str,
    language: str = 'zh-HK',
    skip_if_exists: bool = True,
) -> bool:
    """Upload a show backdrop/banner; skips if the show already has a backdrop."""
    url = f"https://www.themoviedb.org/tv/{show_id}/images/backdrops?language={language}"
    return _upload_to_tmdb_image_page(driver, url, image_path, f"節目 {show_id} 橫幅", skip_if_exists)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _extract_show_id(raw: str) -> str:
    """Accept a bare numeric ID or any TMDB URL containing /tv/{id} and return the numeric ID."""
    raw = raw.strip()
    m = re.search(r'themoviedb\.org/tv/(\d+)', raw)
    if m:
        return m.group(1)
    if re.fullmatch(r'\d+', raw):
        return raw
    raise ValueError(f"無法從輸入中解析 TMDB 劇集 ID: {raw!r}")


def _parse_episode_range(text: str, available: list[int]) -> list[int]:
    """
    Parse a comma-separated list of ranges/numbers into a filtered episode list.
    Examples: '5', '3-7', '1-27, 34-41, 49-59'
    Returns available (all) if the input is empty or unparseable.
    """
    text = text.strip()
    if not text:
        return available
    selected = set()
    for segment in text.split(','):
        segment = segment.strip()
        m = re.match(r'^(\d+)(?:-(\d+))?$', segment)
        if not m:
            return available  # unrecognised format → return all
        start = int(m.group(1))
        end = int(m.group(2)) if m.group(2) else start
        selected.update(ep for ep in available if start <= ep <= end)
    return sorted(selected)


def _prompt_image_source(label: str, temp_dir: str) -> str | None:
    """Ask user for a local path or URL; download URL to temp_dir. Returns local path or None."""
    raw = input(f"{label} (本地路徑或 http URL，留空跳過): ").strip()
    if not raw:
        return None
    if raw.startswith('http'):
        ext = os.path.splitext(raw.split('?')[0])[1] or '.jpg'
        dest = os.path.join(temp_dir, f"download{ext}")
        return dest if download_image(raw, dest) else None
    return raw if os.path.isfile(raw) else None


def _ask_season_config(available: list[int] | None = None, source_name: str = '來源') -> dict:
    """
    Prompt the user for season mapping configuration.
    Returns either:
      {'type': 'single', 'season': int, 'offset': int}
      {'type': 'multi',  'renumber': bool, 'ranges': [(season_num, src_start, src_end), ...]}
    """
    if available:
        print(f"\n共找到 {len(available)} 集 ({source_name} 第 {available[0]}–{available[-1]} 集)")
    print("\n季數設定:")
    print("  1. 單一季數")
    print(f"  2. 多季 ({source_name} 無分季，自行設定範圍)")
    while True:
        ans = input("選擇 (1/2，默認 1): ").strip() or '1'
        if ans in ('1', '2'):
            break
        print("請輸入 1 或 2。")

    if ans == '2':
        print("\n集數編號方式:")
        print(f"  1. 保留{source_name}原有編號，跨季連續 (例如 S02E43)")
        print("  2. 每季重新從第 1 集開始 (例如 S02E01)")
        while True:
            num_ans = input("選擇 (1/2，默認 1): ").strip() or '1'
            if num_ans in ('1', '2'):
                renumber = (num_ans == '2')
                break
            print("請輸入 1 或 2。")

        while True:
            raw_n = input("共有幾季? ").strip()
            if raw_n.isdigit() and int(raw_n) >= 1:
                num_seasons = int(raw_n)
                break
            print("請輸入有效的季數。")

        ranges = []
        for s in range(1, num_seasons + 1):
            while True:
                raw = input(f"第 {s} 季: {source_name}第幾集到第幾集? (例如 1-42): ").strip()
                m = re.match(r'^(\d+)-(\d+)$', raw)
                if m:
                    start, end = int(m.group(1)), int(m.group(2))
                    if start <= end:
                        ranges.append((s, start, end))
                        break
                print("格式錯誤，請重新輸入 (例如 1-42)。")

        return {'type': 'multi', 'renumber': renumber, 'ranges': ranges}

    # Single season
    season = int(input("季數 (默認 1): ").strip() or '1')
    print("\n集數編號方式:")
    print(f"  1. 保留{source_name}原有編號，跨季連續 (例如 S02E43)")
    print("  2. 重新從第 1 集開始 (例如 S02E01)")
    while True:
        num_ans = input("選擇 (1/2，默認 1): ").strip() or '1'
        if num_ans in ('1', '2'):
            break
        print("請輸入 1 或 2。")
    if num_ans == '2':
        while True:
            raw_start = input(f"此季{source_name}起始集數 (例如 43): ").strip()
            if raw_start.isdigit() and int(raw_start) >= 1:
                ep_offset = int(raw_start) - 1
                break
            print("請輸入有效的起始集數。")
    else:
        ep_offset = 0
    return {'type': 'single', 'season': season, 'offset': ep_offset}


def main():
    print("=== Viu 圖片上傳至 TMDB (支援 viu.tv / viu.com) ===\n")
    print("上傳類型:")
    print("  1. 集數劇照 (Episode Backdrops) — 從 viu.tv / viu.com 提取")
    print("  2. 節目封面 (Show Poster)")
    print("  3. 節目橫幅 (Show Backdrop)")
    print("  4. 封面 + 橫幅 (Poster & Backdrop)")
    mode = input("選擇 (1-4，默認 1): ").strip() or '1'

    # Modes 1 and 3 ask for the source URL upfront; other modes only need TMDB ID.
    if mode in ('1', '3'):
        viutv_url = input("節目 URL (viu.tv 或 viu.com): ").strip()
        show_id   = _extract_show_id(input("TMDB 劇集 ID 或 URL (例如 216055 或 https://www.themoviedb.org/tv/216055): "))
    else:
        viutv_url = None
        show_id = _extract_show_id(input("TMDB 劇集 ID 或 URL (例如 216055 或 https://www.themoviedb.org/tv/216055): "))

    is_viu_com = viutv_url is not None and 'viu.com' in viutv_url

    language = input("語言代碼 (默認 zh-HK): ").strip() or 'zh-HK'

    temp_dir = tempfile.mkdtemp(prefix='viutv_img_')
    log.info(f"臨時圖片目錄: {temp_dir}")

    driver = None
    try:
        configs = load_configs()
        driver_tuple = get_browser_driver(
            profile_suffix='_viutv_img', remote_port=9226, base_dir=project_root
        )
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

        # ── Mode 1: Episode backdrops ────────────────────────────────────────
        if mode == '1':
            if is_viu_com:
                episode_images = extract_episode_images_viu(driver, viutv_url)
                source_name = 'viu.com '
            else:
                episode_images = extract_episode_images(driver, viutv_url)
                source_name = 'viu.tv '

            if not episode_images:
                log.error("未找到任何集數圖片，程式終止。")
                return

            available = sorted(episode_images.keys())
            season_config = _ask_season_config(available, source_name=source_name)

            range_str  = input(f"上傳範圍 ({source_name}集數，例如 1-15，直接按 Enter 上傳全部): ").strip()
            target_eps = _parse_episode_range(range_str, available) if range_str else available

            if not target_eps:
                log.warning("沒有符合範圍的集數，程式終止。")
                return

            # Build viu_ep → (tmdb_season, tmdb_ep) mapping function
            if season_config['type'] == 'single':
                _s, _off = season_config['season'], season_config['offset']
                def _get_mapping(viu_ep):
                    return (_s, viu_ep - _off)
            else:
                _ep_map: dict[int, tuple[int, int]] = {}
                for (s, viu_start, viu_end) in season_config['ranges']:
                    _off = (viu_start - 1) if season_config['renumber'] else 0
                    for viu_ep in range(viu_start, viu_end + 1):
                        _ep_map[viu_ep] = (s, viu_ep - _off)
                def _get_mapping(viu_ep):
                    return _ep_map.get(viu_ep)

            success, failed = 0, []
            auto_crop = None  # None=not decided, True=always crop, False=never crop
            upload_all_state = [False]  # becomes True when user picks 'a' (all)

            def _upload_one(ep) -> bool | None:
                """Download, convert, crop and upload one episode.
                Returns True=success, False=upload failed, None=skipped (no mapping)."""
                nonlocal auto_crop
                mapping = _get_mapping(ep)
                if mapping is None:
                    log.warning(f"第 {ep} 集不在任何季數範圍內，跳過。")
                    return None
                tmdb_season, tmdb_ep = mapping

                img_url = episode_images[ep]
                print(f"  第 {ep} 集圖片 URL: {img_url}")
                ext  = os.path.splitext(img_url.split('?')[0])[1] or '.jpg'
                dest = os.path.join(temp_dir, f"ep{ep:03d}{ext}")

                if not download_image(img_url, dest):
                    log.warning(f"第 {ep} 集圖片下載失敗。")
                    return False

                dest_final = _ensure_jpg(dest)

                try:
                    with Image.open(dest_final) as _img:
                        _w, _h = _img.size
                    _ratio = _w / _h
                    _is_16_9 = abs(_ratio - _RATIO_16_9) <= _RATIO_TOLERANCE
                except Exception:
                    _is_16_9 = True

                if not _is_16_9:
                    _do_crop = False
                    if auto_crop is True:
                        _do_crop = True
                    elif auto_crop is None:
                        print(f"\n第 {ep} 集圖片尺寸: {_w}x{_h}（比例 {_ratio:.3f}，非 16:9）。")
                        _ans = input("是否裁切成 16:9? (y=裁切此集, a=全部裁切, n=跳過): ").strip().lower()
                        if _ans in ('y', 'a'):
                            _do_crop = True
                        if _ans == 'a':
                            auto_crop = True
                    if _do_crop:
                        crop_to_16_9(dest_final)

                log.info(f"上傳 {source_name}第 {ep} 集 → S{tmdb_season:02d}E{tmdb_ep:02d}")
                return upload_episode_backdrop(driver, show_id, tmdb_season, tmdb_ep, dest_final, language, upload_all_state=upload_all_state)

            for ep in target_eps:
                ok = _upload_one(ep)
                if ok is None:
                    failed.append(ep)
                elif ok:
                    success += 1
                else:
                    failed.append(ep)
                time.sleep(1.5)

            # After main loop: offer to retry all remaining failures
            while failed:
                print(f"\n目前成功: {success}/{len(target_eps)} 集。失敗集數: {failed}")
                retry_all = input("是否重試所有失敗的集數? (y/n，默認 n): ").strip().lower()
                if retry_all != 'y':
                    break
                retry_list, failed = list(failed), []
                for ep in retry_list:
                    ok = _upload_one(ep)
                    if ok:
                        success += 1
                    else:
                        failed.append(ep)
                    time.sleep(1.5)

            print(f"\n完成！成功上傳 {success}/{len(target_eps)} 集。")
            if failed:
                hint = ','.join(str(e) for e in failed)
                print(f"失敗的集數 ({source_name}編號): {failed}")
                print(f"提示: 重新執行程式時，在上傳範圍輸入 [{hint}] 可只重試這些集數。")

        # ── Mode 2: Show poster ──────────────────────────────────────────────
        elif mode == '2':
            path = _prompt_image_source("封面圖片", temp_dir)
            if path:
                upload_show_poster(driver, show_id, path, language)
            else:
                print("已跳過封面上傳。")

        # ── Mode 3: Show backdrop ────────────────────────────────────────────
        elif mode == '3':
            path = None
            if viutv_url:
                if is_viu_com:
                    backdrop_url = _extract_show_backdrop_viu(driver, viutv_url)
                    site_label = 'viu.com'
                else:
                    backdrop_url = _extract_show_backdrop(driver, viutv_url)
                    site_label = 'viu.tv'
                if backdrop_url:
                    print(f"\n  節目橫幅圖片 URL: {backdrop_url}")
                    use_it = input("是否使用此圖片上傳? (y/n，默認 y): ").strip().lower() or 'y'
                    if use_it == 'y':
                        ext = os.path.splitext(backdrop_url.split('?')[0])[1] or '.jpg'
                        dest = os.path.join(temp_dir, f"show_backdrop{ext}")
                        if download_image(backdrop_url, dest):
                            path = _ensure_jpg(dest)
                        else:
                            log.warning("節目橫幅圖片下載失敗。")
                else:
                    log.warning(f"未能從 {site_label} 頁面提取橫幅圖片。")
            if not path:
                path = _prompt_image_source("橫幅圖片", temp_dir)
            if path:
                upload_show_backdrop(driver, show_id, path, language)
            else:
                print("已跳過橫幅上傳。")

        # ── Mode 4: Poster + Backdrop ────────────────────────────────────────
        elif mode == '4':
            poster_path  = _prompt_image_source("封面圖片 (Poster)", temp_dir)
            backdrop_path = _prompt_image_source("橫幅圖片 (Backdrop)", temp_dir)
            if poster_path:
                upload_show_poster(driver, show_id, poster_path, language)
            if backdrop_path:
                upload_show_backdrop(driver, show_id, backdrop_path, language)

        else:
            print(f"不認識的模式: {mode}")

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if driver:
            cleanup_browser(driver)


if __name__ == '__main__':
    main()
