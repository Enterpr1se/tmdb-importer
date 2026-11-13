"""
【2026-11 L2】記錄本程式自己開啟的瀏覽器行程（driver + 瀏覽器主行程），結束時一定關閉。

背景：監控面板 / 命令視窗的 Ctrl+C 會送到同一 console 的所有行程，msedgedriver 往往先結束，
Edge 主行程因此失去父行程，不再是本程式的「子孫行程」，舊的強制結束邏輯找不到它 → 瀏覽器殘留。
這裡在啟動時就記下 PID（連同建立時間，避免 PID 被系統重用後誤殺別的程式），
結束時關閉這些行程及其子行程。只處理本程式記錄過的行程，不會影響使用者自行開啟的瀏覽器。

此模組不可 import 專案內其他模組（避免循環匯入），可在任何執行緒呼叫。
"""
import threading

try:
    import psutil
except Exception:  # psutil 是必要依賴；萬一缺少，所有函數變成 no-op
    psutil = None

_lock = threading.Lock()
_owned = {}      # pid -> create_time
_drivers = {}    # id(driver) -> (driver, [pid, ...])


def _snapshot(pid):
    """回傳 pid 及其所有子孫行程的 {pid: create_time}。"""
    out = {}
    if psutil is None or not pid:
        return out
    try:
        p = psutil.Process(pid)
        out[p.pid] = p.create_time()
        for c in p.children(recursive=True):
            try:
                out[c.pid] = c.create_time()
            except Exception:
                pass
    except Exception:
        pass
    return out


def register_driver(driver):
    """瀏覽器啟動成功後呼叫：記錄 driver 行程及其子孫（瀏覽器主行程等）。"""
    try:
        pid = driver.service.process.pid
    except Exception:
        return
    snap = _snapshot(pid)
    if not snap:
        return
    with _lock:
        _owned.update(snap)
        _drivers[id(driver)] = (driver, list(snap))


def refresh_driver(driver):
    """補記 driver 之後新增的子行程（例如瀏覽器後來才開的 renderer）。"""
    with _lock:
        known = id(driver) in _drivers
    if known:
        register_driver(driver)


def unregister_driver(driver):
    """driver 已正常關閉後呼叫。"""
    with _lock:
        entry = _drivers.pop(id(driver), None)
        if entry:
            for pid in entry[1]:
                _owned.pop(pid, None)


def kill_driver_procs(driver):
    """cleanup_browser 之後呼叫：結束這個 driver 記錄過、但仍殘留的行程（例如失去父行程的 Edge），然後取消記錄。"""
    with _lock:
        entry = _drivers.get(id(driver))
        pids = {pid: _owned.get(pid) for pid in (entry[1] if entry else [])}
    procs = _alive_owned({p: t for p, t in pids.items() if t is not None})
    for p in procs:
        try:
            p.kill()
        except Exception:
            pass
    if psutil is not None and procs:
        try:
            psutil.wait_procs(procs, timeout=3)
        except Exception:
            pass
    unregister_driver(driver)
    return len(procs)


def _alive_owned(items=None):
    """回傳仍存活、且 PID 未被重用的已記錄行程（含它們現時的子孫）。"""
    if psutil is None:
        return []
    if items is None:
        with _lock:
            items = dict(_owned)
    items = list(items.items())
    procs = {}
    for pid, ctime in items:
        try:
            p = psutil.Process(pid)
            if abs(p.create_time() - ctime) > 1:
                continue  # PID 已被其他程式重用
            procs[p.pid] = p
            for c in p.children(recursive=True):
                procs[c.pid] = c
        except Exception:
            continue
    return list(procs.values())


def kill_owned(quit_drivers=True, quit_timeout=8):
    """關閉所有已記錄的瀏覽器：先嘗試 driver.quit()（有逾時），再強制結束仍存活的行程。
    回傳被強制結束的行程數目。"""
    if quit_drivers:
        with _lock:
            drivers = [d for d, _ in _drivers.values()]
        for d in drivers:
            t = threading.Thread(target=lambda d=d: _safe_quit(d), daemon=True)
            t.start()
            t.join(timeout=quit_timeout)
    procs = _alive_owned()
    killed = 0
    for p in procs:
        try:
            p.kill()
            killed += 1
        except Exception:
            pass
    if psutil is not None and procs:
        try:
            psutil.wait_procs(procs, timeout=3)
        except Exception:
            pass
    with _lock:
        _owned.clear()
        _drivers.clear()
    return killed


def _safe_quit(driver):
    try:
        driver.quit()
    except Exception:
        pass
