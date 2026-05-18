#!/usr/bin/env python3
"""
KDHE KEIMS Compliance Document Batch Downloader
================================================
Reads the CSV export from KDHE KEIMS, shows all documents in a table,
and downloads selected files to a folder you choose.

SETUP (one-time):
    pip install requests

Python 3.8+ required. tkinter is included with standard Python installations.

USAGE:
    python kdhe_downloader.py
"""

import csv
import random
import re
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from html.parser import HTMLParser

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

MAX_RETRIES = 5
RETRIABLE_STATUS_CODES = {429, 500, 502, 503, 504}
RETRY_BACKOFF_BASE = 0.5
RETRY_BACKOFF_CAP = 8.0

# ── Helpers ───────────────────────────────────────────────────────────────────

class AnchorParser(HTMLParser):
    """Extract the href from an HTML <a> tag string."""
    def __init__(self):
        super().__init__()
        self.href = None

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            for attr, val in attrs:
                if attr == 'href':
                    self.href = val


def extract_url(html_fragment: str) -> str | None:
    parser = AnchorParser()
    parser.feed(html_fragment)
    return parser.href


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name)
    name = re.sub(r'_+', '_', name).strip('_ ')
    return name[:200]


def load_csv(path: str) -> list[dict]:
    """Parse the KDHE documents CSV and return a list of document records."""
    documents = []
    with open(path, newline='', encoding='utf-8-sig') as fh:
        reader = csv.DictReader(fh)
        for i, row in enumerate(reader, start=1):
            url = extract_url(row.get('DocMgmtDocurl', ''))
            name = row.get('DocMgmtDocName', '').strip()
            if not url:
                continue
            documents.append({
                'index': i,
                'url': url,
                'name': name or f'document_{i}.pdf',
                'description': row.get('DocMgmtDocDescr', '').strip(),
                'category': row.get('DocMgmtCategory', '').strip(),
                'date': row.get('DocMgmtDocRvcdCreatedDate', '').strip(),
                'source_number': row.get('DocMgmtSourcenumber', '').strip(),
                'source_type': row.get('DocMgmtSourcetype', '').strip(),
                'status': row.get('DocMgmtRefDocStatTypeDescr', '').strip(),
            })
    return documents


def is_retriable_reason(reason: str) -> bool:
    return reason in {'timeout', 'http_429', 'http_5xx', 'request_error'}


def classify_exception(exc: Exception) -> tuple[str, str]:
    if HAS_REQUESTS and isinstance(exc, requests.Timeout):
        return ('timeout', str(exc))
    if HAS_REQUESTS and isinstance(exc, requests.HTTPError):
        code = exc.response.status_code if exc.response is not None else None
        if code == 429:
            return ('http_429', f'HTTP {code}: {exc}')
        if code is not None and 500 <= code < 600:
            return ('http_5xx', f'HTTP {code}: {exc}')
        return ('http_error', f'HTTP {code}: {exc}')
    if HAS_REQUESTS and isinstance(exc, requests.RequestException):
        return ('request_error', str(exc))
    return ('unexpected_error', str(exc))


def backoff_seconds(attempt: int) -> float:
    raw = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
    jitter = random.uniform(0, RETRY_BACKOFF_BASE)
    return min(RETRY_BACKOFF_CAP, raw + jitter)

# ── Main application ──────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('KDHE Document Batch Downloader')
        self.geometry('1060x720')
        self.resizable(True, True)
        self.minsize(900, 580)
        self.configure(bg='#f2f2f2')

        self.documents: list[dict] = []
        self._cancel_flag = False
        self._download_thread: threading.Thread | None = None

        self._build_ui()
        self._check_requests()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {'padx': 8, 'pady': 4}

        # ── Step 1: CSV load ──────────────────────────────────────────────────
        step1 = ttk.LabelFrame(self, text='Step 1 - Load KDHE documents CSV', padding=8)
        step1.pack(fill='x', padx=12, pady=(12, 4))

        ttk.Button(step1, text='Browse CSV...', command=self._browse_csv).pack(side='left', **pad)
        self.csv_var = tk.StringVar()
        csv_entry = ttk.Entry(step1, textvariable=self.csv_var, width=72)
        csv_entry.pack(side='left', fill='x', expand=True, **pad)
        csv_entry.bind('<Return>', lambda _: self._load_documents())

        # ── Step 2: Document table ────────────────────────────────────────────
        step2 = ttk.LabelFrame(self, text='Step 2 - Select documents to download', padding=8)
        step2.pack(fill='both', expand=True, padx=12, pady=4)

        toolbar = ttk.Frame(step2)
        toolbar.pack(fill='x', pady=(0, 4))

        ttk.Button(toolbar, text='Select All', command=self._select_all).pack(side='left', padx=4)
        ttk.Button(toolbar, text='Deselect All', command=self._deselect_all).pack(side='left', padx=4)
        self.selection_label = ttk.Label(toolbar, text='0 documents loaded')
        self.selection_label.pack(side='right', padx=6)

        table_frame = ttk.Frame(step2)
        table_frame.pack(fill='both', expand=True)

        columns = ('name', 'source_type', 'description', 'date')
        headings = ('Filename', 'Source Type', 'Description', 'Date Received')
        col_widths = (300, 220, 240, 150)

        self.tree = ttk.Treeview(
            table_frame, columns=columns, show='headings',
            selectmode='extended', height=14
        )
        for col, heading, width in zip(columns, headings, col_widths):
            self.tree.heading(col, text=heading,
                              command=lambda c=col: self._sort_column(c))
            self.tree.column(col, width=width, minwidth=60)

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        # Right-click context menu
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label='Select all', command=self._select_all)
        self.context_menu.add_command(label='Deselect all', command=self._deselect_all)
        self.tree.bind('<Button-3>', self._show_context_menu)

        # ── Step 3: Output folder ─────────────────────────────────────────────
        step3 = ttk.LabelFrame(self, text='Step 3 - Output folder', padding=8)
        step3.pack(fill='x', padx=12, pady=4)

        self.out_var = tk.StringVar(value=str(Path.home() / 'Downloads'))
        ttk.Entry(step3, textvariable=self.out_var, width=72).pack(
            side='left', fill='x', expand=True, **pad
        )
        ttk.Button(step3, text='Browse...', command=self._browse_output).pack(side='left', **pad)

        # ── Download controls ─────────────────────────────────────────────────
        ctrl_frame = ttk.Frame(self)
        ctrl_frame.pack(fill='x', padx=12, pady=4)

        self.download_btn = ttk.Button(
            ctrl_frame,
            text='⬇ Download Selected',
            command=self._start_download,
            state='disabled'
        )
        self.download_btn.pack(side='left', padx=4)

        self.cancel_btn = ttk.Button(
            ctrl_frame,
            text='⏹ Cancel',
            command=self._cancel_download,
            state='disabled'
        )
        self.cancel_btn.pack(side='left', padx=4)

        self.progress = ttk.Progressbar(ctrl_frame, orient='horizontal', mode='determinate', length=380)
        self.progress.pack(side='left', padx=8)

        self.progress_label = ttk.Label(ctrl_frame, text='0 / 0')
        self.progress_label.pack(side='left')

        self.tree.bind('<<TreeviewSelect>>', lambda _: self._update_selection_label())

        # ── Status ─────────────────────────────────────────────────────────────
        prog_frame = ttk.Frame(self)
        prog_frame.pack(fill='x', padx=12, pady=(0, 8))

        self.status_var = tk.StringVar(value='Ready.')
        ttk.Label(prog_frame, textvariable=self.status_var, anchor='w').grid(
            row=1, column=0, sticky='ew')

    # ── Checks ────────────────────────────────────────────────────────────────

    def _check_requests(self):
        if not HAS_REQUESTS:
            messagebox.showerror(
                'Missing dependency',
                'The "requests" library is required.\n\n'
                'Install it by running:\n    pip install requests\n\n'
                'Then restart this application.'
            )

    # ── File browsing ─────────────────────────────────────────────────────────

    def _browse_csv(self):
        path = filedialog.askopenfilename(
            title='Select KDHE documents CSV',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')]
        )
        if path:
            self.csv_var.set(path)
            self._load_documents()

    def _browse_output(self):
        path = filedialog.askdirectory(title='Select download folder')
        if path:
            self.out_var.set(path)

    # ── Document loading ──────────────────────────────────────────────────────

    def _load_documents(self):
        csv_path = self.csv_var.get().strip()
        if not csv_path:
            messagebox.showwarning('No file selected', 'Please select a CSV file first.')
            return
        if not Path(csv_path).is_file():
            messagebox.showerror('File not found', f'Cannot find:\n{csv_path}')
            return

        try:
            docs = load_csv(csv_path)
        except Exception as exc:
            messagebox.showerror('Error reading CSV', str(exc))
            return

        if not docs:
            messagebox.showwarning('No documents', 'No download links were found in the CSV.')
            return

        self.documents = docs
        self._populate_table(docs)
        self._select_all()
        self._update_selection_label()
        self.download_btn.config(state='normal')
        self.status_var.set(f'Loaded {len(docs)} documents.')

    def _populate_table(self, docs: list[dict]):
        self.tree.delete(*self.tree.get_children())
        for doc in docs:
            self.tree.insert('', 'end', iid=str(doc['index']), values=(
                doc['name'],
                doc['source_type'],
                doc['description'],
                doc['date'],
            ))

    # ── Selection helpers ─────────────────────────────────────────────────────

    def _select_all(self):
        self.tree.selection_set(self.tree.get_children())
        self._update_selection_label()

    def _deselect_all(self):
        self.tree.selection_remove(self.tree.get_children())
        self._update_selection_label()

    def _update_selection_label(self):
        total = len(self.tree.get_children())
        selected = len(self.tree.selection())
        if total == 0:
            self.selection_label.config(text='0 documents loaded')
        else:
            self.selection_label.config(text=f'{selected} of {total} selected.')

    def _show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    # ── Table sorting ─────────────────────────────────────────────────────────

    def _sort_column(self, col):
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        data.sort()
        for i, (_, k) in enumerate(data):
            self.tree.move(k, '', i)

    # ── Download ──────────────────────────────────────────────────────────────

    def _start_download(self):
        if not HAS_REQUESTS:
            self._check_requests()
            return

        selected_iids = self.tree.selection()
        if not selected_iids:
            messagebox.showwarning('Nothing selected', 'Select at least one document to download.')
            return

        out_dir = self.out_var.get().strip()
        if not out_dir:
            messagebox.showwarning('No output folder', 'Please select a folder to save files to.')
            return

        out_path = Path(out_dir)
        try:
            out_path.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror('Cannot create folder', str(exc))
            return

        selected_indices = {int(iid) for iid in selected_iids}
        selected_docs = [d for d in self.documents if d['index'] in selected_indices]

        self._cancel_flag = False
        self.download_btn.config(state='disabled')
        self.cancel_btn.config(state='normal')
        self.progress['value'] = 0
        self.progress['maximum'] = len(selected_docs)
        self.progress_label.config(text=f'0 / {len(selected_docs)}')

        self._download_thread = threading.Thread(
            target=self._download_worker,
            args=(selected_docs, out_path),
            daemon=True
        )
        self._download_thread.start()

    def _cancel_download(self):
        self._cancel_flag = True
        self.status_var.set('Cancelling…')

    def _download_worker(self, docs: list[dict], out_path: Path):
        session = requests.Session()
        session.headers['User-Agent'] = (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0.0.0 Safari/537.36'
        )

        success = 0
        failed = 0
        skipped = 0

        for i, doc in enumerate(docs, start=1):
            if self._cancel_flag:
                skipped = len(docs) - i + 1
                break

            safe_name = sanitize_filename(doc['name'])
            if not safe_name.lower().endswith('.pdf'):
                safe_name += '.pdf'

            filepath = out_path / safe_name

            # Avoid overwriting: append _2, _3, etc.
            if filepath.exists():
                stem = filepath.stem
                counter = 2
                while filepath.exists():
                    filepath = out_path / f'{stem}_{counter}.pdf'
                    counter += 1

            self._set_status(f'({i}/{len(docs)}) Downloading: {doc["name"][:60]}')

            download_ok = False
            last_error = 'Unknown failure'
            last_reason = 'unexpected_error'

            for attempt in range(1, MAX_RETRIES + 2):
                if self._cancel_flag:
                    break

                try:
                    if attempt > 1:
                        self._set_status(
                            f'({i}/{len(docs)}) Retry {attempt - 1}/{MAX_RETRIES}: {doc["name"][:50]}'
                        )

                    response = session.get(doc['url'], stream=True, timeout=60)
                    response.raise_for_status()
                    with open(filepath, 'wb') as fh:
                        for chunk in response.iter_content(chunk_size=16_384):
                            if chunk:
                                fh.write(chunk)

                    download_ok = True
                    success += 1
                    break
                except Exception as exc:
                    last_reason, last_error = classify_exception(exc)
                    if attempt <= MAX_RETRIES and is_retriable_reason(last_reason):
                        delay = backoff_seconds(attempt)
                        self._set_status(
                            f'({i}/{len(docs)}) RETRYING in {delay:.2f}s: '
                            f'{doc["name"][:40]} ({last_reason})'
                        )
                        time.sleep(delay)
                        continue
                    break

            if not download_ok and not self._cancel_flag:
                failed += 1
                self._set_status(f'({i}/{len(docs)}) FAILED: {doc["name"][:50]} — {last_error}')
                time.sleep(0.5)

            self._set_progress(i)
            time.sleep(0.15)  # small courtesy delay

        self.after(0, self._download_finished, success, failed, skipped, out_path)

    def _set_status(self, msg: str):
        self.after(0, lambda: self.status_var.set(msg))

    def _set_progress(self, value: int):
        def update_progress():
            self.progress.configure(value=value)
            total = int(float(self.progress.cget('maximum')))
            self.progress_label.config(text=f'{value} / {total}')

        self.after(0, update_progress)

    def _download_finished(self, success: int, failed: int, skipped: int, out_path: Path):
        self.download_btn.config(state='normal')
        self.cancel_btn.config(state='disabled')
        completed = success + failed
        total = int(float(self.progress.cget('maximum')))
        self.progress.configure(value=completed)
        self.progress_label.config(text=f'{completed} / {total}')

        parts = [f'{success} downloaded']
        if failed:
            parts.append(f'{failed} failed')
        if skipped:
            parts.append(f'{skipped} cancelled')
        summary = ', '.join(parts) + f'.  Saved to: {out_path}'
        self.status_var.set(summary)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = App()
    app.mainloop()
