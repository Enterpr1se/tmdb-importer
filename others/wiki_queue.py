import threading
import queue
import logging
import time

_request_queue = queue.Queue()
_worker_started = False
_start_lock = threading.Lock()

# Adaptive delay: shrinks when requests succeed, grows on 429.
_current_delay = 5.0
_MIN_DELAY = 3.0   # floor — enough spacing to be polite
_MAX_DELAY = 60.0  # ceiling — hit when repeated 429s occur
MAX_SEARCH_SECONDS = 300  # 單次搜索最長允許 5 分鐘，超時放棄並繼續下一個


def _worker():
    from others import utils
    global _current_delay
    log = logging.getLogger()
    log.info("[WikiQueue] Worker thread started, initial delay: %.0fs.", _current_delay)
    while True:
        item = _request_queue.get()
        if item is None:
            log.info("[WikiQueue] Received shutdown signal, stopping.")
            break

        # 【2026-11】隊列項目帶埋要執行嘅函數（舊名稱搜尋 / Wikidata ID 查詢共用同一條隊同限速）
        func, args, kwargs, result_box, done_event = item

        if utils._PROGRAM_INTERRUPTED:
            result_box[0] = None
            done_event.set()
            continue

        def _run():
            try:
                result_box[0] = func(*args, **kwargs)
            except Exception as e:
                log.error("[WikiQueue] Error processing wiki request: %s", e)
                result_box[0] = None

        worker_thread = threading.Thread(target=_run, daemon=True)
        worker_thread.start()
        worker_thread.join(timeout=MAX_SEARCH_SECONDS)
        if worker_thread.is_alive():
            name_hint = args[1] if len(args) > 1 else '?'
            log.warning("[WikiQueue] 搜索 '%s' 超過 %d 秒，放棄並繼續下一個。", name_hint, MAX_SEARCH_SECONDS)
        done_event.set()

        # Adaptive delay: back off on 429, recover on clean runs.
        # Re-import to get the current flag value after the search completed.
        from others.wiki_search import _search_hit_rate_limit as _hit_limit
        if _hit_limit:
            _current_delay = min(_current_delay * 2.0, _MAX_DELAY)
            log.info("[WikiQueue] 偵測到限速 (429)，延長下次延遲至 %.0fs。", _current_delay)
        else:
            _current_delay = max(_current_delay * 0.75, _MIN_DELAY)
            log.debug("[WikiQueue] 搜索順暢，縮短下次延遲至 %.1fs。", _current_delay)

        # Wait in 1s chunks so Ctrl+C responds immediately
        delay_secs = int(_current_delay)
        for _ in range(delay_secs):
            if utils._PROGRAM_INTERRUPTED:
                break
            time.sleep(1)


def _ensure_started():
    global _worker_started
    if _worker_started:
        return
    with _start_lock:
        if not _worker_started:
            t = threading.Thread(target=_worker, name="WikiQueueWorker", daemon=True)
            t.start()
            _worker_started = True


def search_wiki_for_hk_name_queued(item_id, original_name, media_type, year=None,
                                    search_targets_by_lang=None, original_language=None,
                                    tmdb_chinese_names=None):
    """
    Drop-in replacement for search_wiki_for_hk_name that routes all network requests
    through a single serialized worker thread with adaptive rate-limit delay.

    Cache hits are resolved before entering the queue so callers are never blocked
    unnecessarily while waiting for unrelated network work to complete.
    """
    from others.wiki_search import _init_db, _query_wiki_cache, search_wiki_for_hk_name

    _init_db()
    if item_id:
        cached = _query_wiki_cache(item_id, media_type)
        if cached is not None:
            if cached['status'] == 'FOUND':
                return cached['hk_name']
            if cached['status'] in ['NOT_FOUND', 'FAILED']:
                return None

    return _run_queued(search_wiki_for_hk_name, (item_id, original_name, media_type), {
        'year': year,
        'search_targets_by_lang': search_targets_by_lang,
        'original_language': original_language,
        'tmdb_chinese_names': tmdb_chinese_names,
    })


def lookup_hk_name_by_wikidata_queued(item_id, wikidata_id, original_name=None):
    """【2026-11】經隊列執行 wiki_search.lookup_hk_name_by_wikidata；快取命中唔使排隊。"""
    from others.wiki_search import _init_db, _wikidata_cache_lookup, lookup_hk_name_by_wikidata
    _init_db()
    if item_id:
        cached, need_query = _wikidata_cache_lookup(item_id)
        if cached or not need_query:
            return cached
    return _run_queued(lookup_hk_name_by_wikidata, (item_id, wikidata_id, original_name), {})


def _run_queued(func, args, kwargs):
    from others import utils
    if utils._PROGRAM_INTERRUPTED:
        return None

    _ensure_started()

    result_box = [None]
    done_event = threading.Event()
    _request_queue.put((func, args, kwargs, result_box, done_event))

    # Wait in 1s chunks so Ctrl+C can unblock this caller immediately
    while not done_event.wait(timeout=1.0):
        if utils._PROGRAM_INTERRUPTED:
            return None

    return result_box[0]
