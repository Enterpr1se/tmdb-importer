"""
tools_gui.py：sub-program 小工具總面板。

每個工具喺新嘅命令視窗運行（工具本身會問問題，答案照舊喺嗰個視窗輸入），完成後視窗唔會自動關閉。
下面列出各工具輸出嘅 *_urls.txt，可以開啟、複製全部（貼入 bulk_videos.xlsx）或者封存。

用法：雙擊 sub-program/tools.bat，或者 python sub-program/tools_gui.py
"""
import datetime as dt
import glob
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import bulk_txt  # noqa: E402

# (分組, [(檔名, 名稱, 說明, 會唔會直接改 TMDB)])
TOOLS = [
    ('搵 TMDB 網址（結果寫入 <工具>_urls.txt，唔會改 bulk_videos.xlsx）', [
        ('tmdb_search.py', 'TMDB 關鍵字搜尋',
         '輸入關鍵字、TMDB 網址或 ID（用逗號分隔），搵出有中文翻譯嘅劇集 / 電影。', False),
        ('manage_video_urls.py', '檢查影片網址',
         '貼上 TMDB 影片網址，保留有中文內容嘅。', False),
        ('bulk_video_person.py', '影片 / 演員擴展',
         '貼上影片或演員網址：搵出相關演員，再搵佢哋其他有中文嘅作品。', False),
        ('tmdb_person_search_web.py', '演員名搜尋',
         '貼上演員名列表，喺 TMDB 搵演員，再搵佢哋有中文嘅作品。', False),
        ('wiki_bulk.py', '維基百科列表',
         '輸入維基百科頁面網址（例如劇集列表），抽出標題再搜尋 TMDB；搵唔到嘅寫入 data/unfound_titles.xlsx。', False),
        ('tmdb_merged_runner.py', 'TMDB 貢獻 / JAVHD / MissAV',
         'TMDB 首頁、排行榜、指定用家嘅貢獻、JAVHD、MissAV 名單 → 演員作品 → 有中文嘅網址。'
         'MissAV 名單要先用下面「MissAV 書籤小程式」下載。', False),
    ]),
    ('viu.tv / viu.com（會直接修改 TMDB）', [
        ('viutv_show_creator.py', 'viu.tv 節目建立 / 加集數',
         '由 viu.tv、viu.com 或 Netflix 節目頁，喺 TMDB 建立新節目、加新一季或者新集數。', True),
        ('viutv_bulk_updater.py', 'viu.tv 批量更新集數',
         '按 viutv_bulk_config.xlsx 設定，批量為 viu.tv 節目加最新集數。', True),
        ('viutv_image_uploader.py', 'viu.tv 劇照上載',
         '由 viu.tv 抓每集劇照，上載到 TMDB 對應集數。', True),
    ]),
]


def console_python():
    """有命令視窗嘅 python.exe（面板本身可能用 pythonw 開）。"""
    exe = sys.executable
    d, name = os.path.split(exe)
    if name.lower() == 'pythonw.exe' and os.path.exists(os.path.join(d, 'python.exe')):
        return os.path.join(d, 'python.exe')
    return exe


def launch(script, title):
    py = console_python()
    path = os.path.join(HERE, script)
    # cmd /k：工具完成後保留視窗睇結果。成條命令外面多一對引號，cmd 先會正確處理有空格嘅路徑。
    cmd = f'cmd /k "title {title} & "{py}" -X utf8 "{path}""'
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    subprocess.Popen(cmd, cwd=HERE, env=env, creationflags=subprocess.CREATE_NEW_CONSOLE)


def open_path(path):
    try:
        os.startfile(path)
    except OSError as e:
        messagebox.showerror('開啟失敗', f'{path}\n\n{e}')


class ToolsApp:
    def __init__(self, root):
        self.root = root
        root.title('TMDB 小工具')
        root.geometry('900x800')
        root.minsize(640, 480)
        outer = ttk.Frame(root, padding=10)
        outer.pack(fill='both', expand=True)

        canvas = tk.Canvas(outer, highlightthickness=0)
        bar = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        body = ttk.Frame(canvas)
        body.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        win = canvas.create_window((0, 0), window=body, anchor='nw')
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=bar.set)
        canvas.pack(side='left', fill='both', expand=True)
        bar.pack(side='right', fill='y')
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-e.delta / 120), 'units'))

        for group, tools in TOOLS:
            box = ttk.LabelFrame(body, text=group, padding=8)
            box.pack(fill='x', pady=(0, 10))
            box.columnconfigure(1, weight=1)
            for r, (script, name, desc, writes_tmdb) in enumerate(tools):
                exists = os.path.exists(os.path.join(HERE, script))
                btn = ttk.Button(box, text=f'▶ {name}', width=30,
                                 command=lambda s=script, n=name, w=writes_tmdb: self.run_tool(s, n, w))
                btn.grid(row=r, column=0, sticky='nw', padx=(0, 8), pady=3)
                if not exists:
                    btn.state(['disabled'])
                lbl = ttk.Label(box, text=desc + ('' if exists else f'（搵唔到 {script}）'), wraplength=520,
                                justify='left')
                lbl.grid(row=r, column=1, sticky='w', pady=3)
                box.bind('<Configure>', lambda e, l=lbl: l.configure(wraplength=max(200, e.width - 290)), add='+')

        other = ttk.LabelFrame(body, text='其他', padding=8)
        other.pack(fill='x', pady=(0, 10))
        ttk.Button(other, text='MissAV 書籤小程式', width=30,
                   command=lambda: open_path(os.path.join(HERE, 'missav_bookmarklet.html'))).grid(row=0, column=0, sticky='w', padx=(0, 8), pady=3)
        ttk.Label(other, text='喺瀏覽器開啟安裝 / 使用說明。下載嘅 missav_names.txt 交畀「TMDB 貢獻 / JAVHD / MissAV」選項 4。').grid(row=0, column=1, sticky='w')
        ttk.Button(other, text='viu.tv 批量設定檔', width=30,
                   command=lambda: open_path(os.path.join(HERE, 'viutv_bulk_config.xlsx'))).grid(row=1, column=0, sticky='w', padx=(0, 8), pady=3)
        ttk.Label(other, text='開啟 viutv_bulk_config.xlsx（「viu.tv 批量更新集數」用）。').grid(row=1, column=1, sticky='w')
        ttk.Button(other, text='開啟 sub-program 資料夾', width=30,
                   command=lambda: open_path(HERE)).grid(row=2, column=0, sticky='w', padx=(0, 8), pady=3)

        out = ttk.LabelFrame(body, text='輸出網址（*_urls.txt）', padding=8)
        out.pack(fill='both', expand=True)
        cols = ('file', 'count', 'mtime')
        self.tree = ttk.Treeview(out, columns=cols, show='headings', height=6, selectmode='browse')
        for c, t, w in (('file', '檔案', 330), ('count', '網址數', 80), ('mtime', '最後修改', 150)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor='w' if c == 'file' else 'center')
        self.tree.pack(fill='both', expand=True)
        self.tree.bind('<Double-1>', lambda e: self.open_selected())
        btns = ttk.Frame(out)
        btns.pack(fill='x', pady=(6, 0))
        ttk.Button(btns, text='開啟', command=self.open_selected).pack(side='left')
        ttk.Button(btns, text='複製全部網址', command=self.copy_selected).pack(side='left', padx=6)
        ttk.Button(btns, text='匯入 Excel…', command=self.import_selected).pack(side='left', padx=(0, 6))
        ttk.Button(btns, text='封存（已貼入 Excel）', command=self.archive_selected).pack(side='left')
        ttk.Button(btns, text='重新整理', command=self.refresh).pack(side='right')
        self.status = ttk.Label(out, text='', foreground='#555')
        self.status.pack(fill='x', pady=(4, 0))

        self.refresh()
        root.after(5000, self._auto_refresh)

    # ---------- tools
    def run_tool(self, script, name, writes_tmdb):
        if writes_tmdb and not messagebox.askyesno(
                '會修改 TMDB', f'「{name}」會用你嘅 TMDB 帳戶直接修改 TMDB 上面嘅資料。\n\n確定開始？'):
            return
        try:
            launch(script, name)
            self.status.configure(text=f'已喺新視窗開啟：{name}')
        except Exception as e:
            messagebox.showerror('開啟失敗', str(e))

    # ---------- outputs
    def _files(self):
        return sorted(glob.glob(os.path.join(HERE, '*_urls.txt')), key=os.path.getmtime, reverse=True)

    def refresh(self):
        sel = self._selected()
        self.tree.delete(*self.tree.get_children())
        for p in self._files():
            try:
                with open(p, encoding='utf-8') as f:
                    n = sum(1 for line in f if line.strip())
                mtime = dt.datetime.fromtimestamp(os.path.getmtime(p)).strftime('%Y-%m-%d %H:%M')
            except OSError:
                continue
            self.tree.insert('', 'end', iid=p, values=(os.path.basename(p), n, mtime))
        if sel and self.tree.exists(sel):
            self.tree.selection_set(sel)
        if not self.tree.get_children():
            self.status.configure(text='未有輸出檔案。工具搵到網址之後會喺呢度出現。')

    def _auto_refresh(self):
        self.refresh()
        self.root.after(5000, self._auto_refresh)

    def _selected(self):
        s = self.tree.selection()
        return s[0] if s else None

    def _need_selected(self):
        p = self._selected()
        if not p:
            messagebox.showinfo('未揀檔案', '請先喺列表揀一個檔案。')
        return p

    def open_selected(self):
        p = self._need_selected()
        if p:
            open_path(p)

    def copy_selected(self):
        p = self._need_selected()
        if not p:
            return
        with open(p, encoding='utf-8') as f:
            urls = [line.strip() for line in f if line.strip()]
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(urls))
        self.status.configure(text=f'已複製 {len(urls)} 個網址，可以貼入 bulk_videos.xlsx 嘅 Video URL 欄。')

    def import_selected(self):
        p = self._need_selected()
        if not p:
            return
        with open(p, encoding='utf-8') as f:
            urls = [line.strip() for line in f if line.strip()]
        if not urls:
            messagebox.showinfo('冇網址', '呢個檔案冇網址。')
            return
        try:
            sheets = bulk_txt.list_bulk_sheets()
        except Exception as e:
            messagebox.showerror('讀取失敗', f'讀唔到 bulk_videos.xlsx：{e}')
            return
        ImportDialog(self, p, urls, sheets)

    def archive_selected(self, path=None, ask=True):
        p = path or self._need_selected()
        if not p:
            return
        stamp = dt.datetime.now().strftime('%Y%m%d-%H%M%S')
        new = p[:-len('_urls.txt')] + f'_已封存_{stamp}.txt'
        if ask and not messagebox.askyesno('封存', f'將 {os.path.basename(p)} 改名做\n{os.path.basename(new)}？\n\n'
                                         '下次工具會開一個新嘅 txt（已經喺 bulk_videos.xlsx 嘅網址唔會再寫入）。'):
            return
        try:
            os.rename(p, new)
        except OSError as e:
            messagebox.showerror('封存失敗', str(e))
            return
        self.status.configure(text=f'已封存：{os.path.basename(new)}')
        self.refresh()


class ImportDialog:
    """揀工作表（或者新增），將 txt 嘅網址加到 bulk_videos.xlsx。"""
    NEW = '（新增工作表）'

    def __init__(self, app, path, urls, sheets):
        self.app, self.path, self.urls = app, path, urls
        self.top = top = tk.Toplevel(app.root)
        top.title('匯入 bulk_videos.xlsx')
        top.transient(app.root)
        top.grab_set()
        top.resizable(False, False)
        f = ttk.Frame(top, padding=14)
        f.pack(fill='both', expand=True)
        ttk.Label(f, text=f'{os.path.basename(path)}：{len(urls)} 個網址').grid(row=0, column=0, columnspan=2, sticky='w')
        ttk.Label(f, text='加到工作表：').grid(row=1, column=0, sticky='w', pady=(10, 0))
        self.sheet = tk.StringVar(value='Bulk' if 'Bulk' in sheets else (sheets[0] if sheets else self.NEW))
        cb = ttk.Combobox(f, textvariable=self.sheet, values=sheets + [self.NEW], state='readonly', width=30)
        cb.grid(row=1, column=1, sticky='w', pady=(10, 0))
        cb.bind('<<ComboboxSelected>>', lambda e: self._toggle())
        ttk.Label(f, text='新工作表名稱：').grid(row=2, column=0, sticky='w', pady=(6, 0))
        self.new_name = tk.StringVar(value=os.path.basename(path)[:-len('_urls.txt')][:31])
        self.entry = ttk.Entry(f, textvariable=self.new_name, width=33)
        self.entry.grid(row=2, column=1, sticky='w', pady=(6, 0))
        ttk.Label(f, text='已經喺任何工作表嘅網址會略過；新網址加喺工作表最底。\n'
                          '主程式運行緊都可以匯入（會等佢寫完 Done 先寫）。',
                  foreground='#555', justify='left').grid(row=3, column=0, columnspan=2, sticky='w', pady=(10, 0))
        self.archive = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text='匯入成功後封存呢個 txt', variable=self.archive).grid(
            row=4, column=0, columnspan=2, sticky='w', pady=(8, 0))
        self.msg = ttk.Label(f, text='', foreground='#0b6bcb')
        self.msg.grid(row=5, column=0, columnspan=2, sticky='w', pady=(8, 0))
        b = ttk.Frame(f)
        b.grid(row=6, column=0, columnspan=2, sticky='e', pady=(10, 0))
        self.ok = ttk.Button(b, text='匯入', command=self.start)
        self.ok.pack(side='left', padx=(0, 6))
        self.cancel = ttk.Button(b, text='取消', command=top.destroy)
        self.cancel.pack(side='left')
        self._toggle()

    def _toggle(self):
        self.entry.state(['!disabled'] if self.sheet.get() == self.NEW else ['disabled'])

    def start(self):
        create = self.sheet.get() == self.NEW
        name = (self.new_name.get() if create else self.sheet.get()).strip()
        if create and (not name or len(name) > 31 or any(c in name for c in '[]:*?/\\')):
            messagebox.showerror('名稱唔啱', '工作表名稱唔可以空白、超過 31 個字，或者有 [ ] : * ? / \\。',
                                 parent=self.top)
            return
        self.ok.state(['disabled'])
        self.cancel.state(['disabled'])
        self.top.protocol('WM_DELETE_WINDOW', lambda: None)

        # 背景執行緒唔可以直接改 tkinter 元件：經 queue 傳返 UI 執行緒處理
        q = queue.Queue()

        def work():
            try:
                res = bulk_txt.import_urls_to_bulk(self.urls, name, create_sheet=create,
                                                   progress=lambda m: q.put(('msg', m)))
                q.put(('done', (res, None)))
            except Exception as e:
                q.put(('done', (None, e)))

        def poll():
            try:
                while True:
                    kind, val = q.get_nowait()
                    if kind == 'msg':
                        self.msg.configure(text=val)
                    else:
                        self.done(name, *val)
                        return
            except queue.Empty:
                pass
            self.top.after(200, poll)

        threading.Thread(target=work, daemon=True).start()
        poll()

    def done(self, name, res, err):
        self.top.protocol('WM_DELETE_WINDOW', self.top.destroy)
        if err:
            self.ok.state(['!disabled'])
            self.cancel.state(['!disabled'])
            self.msg.configure(text='')
            messagebox.showerror('匯入失敗', str(err), parent=self.top)
            return
        added, skipped = res
        self.top.destroy()
        text = f'已加 {added} 個網址到工作表「{name}」。' + (f'\n{skipped} 個已經喺 Excel，略過。' if skipped else '')
        if self.archive.get():
            self.app.archive_selected(self.path, ask=False)
        self.app.status.configure(text=text.replace('\n', ' '))
        messagebox.showinfo('匯入完成', text)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use('vista')
    except tk.TclError:
        pass
    ToolsApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
