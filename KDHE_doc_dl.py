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
import os
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Empty, Queue
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
DEFAULT_WORKERS = 4

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


class DownloadCancelled(Exception):
    pass

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
        self._cancel_event = threading.Event()
        self._download_thread: threading.Thread | None = None
        self._status_q: Queue = Queue()
        self._thread_local = threading.local()
        self._name_lock = threading.Lock()
        self._reserved_names: set[str] = set()
        self._counter_lock = threading.Lock()
        self._total = 0
        self._completed = 0
        self._success = 0
        self._failed = 0
        self._cancelled = 0

        self._build_ui()
        self._check_requests()
        self._poll_ui_queue()

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

        self._cancel_event.clear()
        with self._name_lock:
            self._reserved_names.clear()
        with self._counter_lock:
            self._total = len(selected_docs)
            self._completed = 0
            self._success = 0
            self._failed = 0
            self._cancelled = 0

        self.download_btn.config(state='disabled')
        self.cancel_btn.config(state='normal')
        self.progress['value'] = 0
        self.progress['maximum'] = len(selected_docs)
        self.progress_label.config(text=f'0 / {len(selected_docs)}')

        workers = min(DEFAULT_WORKERS, len(selected_docs))
        self._enqueue_ui('status', f'Starting {workers} worker(s) for {len(selected_docs)} file(s)...')

        self._download_thread = threading.Thread(
            target=self._download_worker,
            args=(selected_docs, out_path),
            daemon=True
        )
        self._download_thread.start()

    def _cancel_download(self):
        self._cancel_event.set()
        self.cancel_btn.config(state='disabled')
        self._enqueue_ui('status', 'Cancelling...')

    def _enqueue_ui(self, kind: str, value=None):
        self._status_q.put((kind, value))

    def _poll_ui_queue(self):
        try:
            while True:
                kind, value = self._status_q.get_nowait()
                if kind == 'status':
                    self.status_var.set(str(value))
                elif kind == 'progress':
                    completed, total = value
                    self.progress.configure(value=completed, maximum=total)
                    self.progress_label.config(text=f'{completed} / {total}')
                elif kind == 'done':
                    self._download_finished(value['success'], value['failed'], value['cancelled'], value['out_path'])
        except Empty:
            pass
        self.after(100, self._poll_ui_queue)

    def _get_thread_session(self):
        session = getattr(self._thread_local, 'session', None)
        if session is None:
            session = requests.Session()
            session.headers['User-Agent'] = (
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                'AppleWebKit/537.36 (KHTML, like Gecko) '
                'Chrome/120.0.0.0 Safari/537.36'
            )
            self._thread_local.session = session
        return session

    def _reserve_output_path(self, out_path: Path, original_name: str) -> Path:
        safe_name = sanitize_filename(original_name)
        if not safe_name.lower().endswith('.pdf'):
            safe_name += '.pdf'

        stem = Path(safe_name).stem
        suffix = Path(safe_name).suffix
        counter = 1

        with self._name_lock:
            while True:
                candidate = safe_name if counter == 1 else f'{stem}_{counter}.{suffix.lstrip(".")}'
                candidate_key = candidate.lower()
                candidate_path = out_path / candidate
                if candidate_key not in self._reserved_names and not candidate_path.exists():
                    self._reserved_names.add(candidate_key)
                    return candidate_path
                counter += 1

    def _release_output_path(self, filepath: Path):
        with self._name_lock:
            self._reserved_names.discard(filepath.name.lower())

    @staticmethod
    def _cleanup_partial_file(filepath: Path):
        tmp_path = filepath.with_suffix(filepath.suffix + '.part')
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except Exception:
            pass

    def _download_one(self, doc: dict, out_path: Path, total: int) -> str:
        if self._cancel_event.is_set():
            return 'cancelled'

        filepath = self._reserve_output_path(out_path, doc['name'])
        tmp_path = filepath.with_suffix(filepath.suffix + '.part')
        display_name = doc['name'][:60]
        self._enqueue_ui('status', f'Downloading: {display_name}')

        session = self._get_thread_session()
        last_reason = 'unexpected_error'
        last_error = 'Unknown failure'

        try:
            for attempt in range(1, MAX_RETRIES + 2):
                if self._cancel_event.is_set():
                    raise DownloadCancelled('Cancelled by user')

                try:
                    if attempt > 1:
                        self._enqueue_ui(
                            'status',
                            f'Retry {attempt - 1}/{MAX_RETRIES}: {doc["name"][:50]}'
                        )

                    response = session.get(doc['url'], stream=True, timeout=60)
                    response.raise_for_status()
                    with open(tmp_path, 'wb') as fh:
                        for chunk in response.iter_content(chunk_size=16_384):
                            if self._cancel_event.is_set():
                                raise DownloadCancelled('Cancelled during download stream')
                            if chunk:
                                fh.write(chunk)

                    os.replace(tmp_path, filepath)
                    return 'success'
                except DownloadCancelled:
                    raise
                except Exception as exc:
                    self._cleanup_partial_file(filepath)
                    last_reason, last_error = classify_exception(exc)
                    if attempt <= MAX_RETRIES and is_retriable_reason(last_reason):
                        delay = backoff_seconds(attempt)
                        self._enqueue_ui(
                            'status',
                            f'RETRYING in {delay:.2f}s: {doc["name"][:40]} ({last_reason})'
                        )
                        if self._cancel_event.wait(delay):
                            raise DownloadCancelled('Cancelled during retry backoff')
                        continue
                    break
        except DownloadCancelled:
            self._cleanup_partial_file(filepath)
            return 'cancelled'
        finally:
            self._release_output_path(filepath)

        self._enqueue_ui('status', f'FAILED: {doc["name"][:50]} - {last_error}')
        return 'failed'

    def _download_worker(self, docs: list[dict], out_path: Path):
        total = len(docs)
        workers = min(DEFAULT_WORKERS, total)

        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix='kdhe-dl') as pool:
            futures = [pool.submit(self._download_one, doc, out_path, total) for doc in docs]
            cancelled_pending = False

            for fut in as_completed(futures):
                if self._cancel_event.is_set() and not cancelled_pending:
                    cancelled_pending = True
                    for pending in futures:
                        pending.cancel()

                if fut.cancelled():
                    result = 'cancelled'
                else:
                    try:
                        result = fut.result()
                    except Exception:
                        result = 'failed'

                with self._counter_lock:
                    self._completed += 1
                    if result == 'success':
                        self._success += 1
                    elif result == 'cancelled':
                        self._cancelled += 1
                    else:
                        self._failed += 1
                    completed = self._completed

                self._enqueue_ui('progress', (completed, total))

        with self._counter_lock:
            done_payload = {
                'success': self._success,
                'failed': self._failed,
                'cancelled': self._cancelled,
                'out_path': out_path,
            }
        self._enqueue_ui('done', done_payload)

    def _set_status(self, msg: str):
        self._enqueue_ui('status', msg)

    def _set_progress(self, value: int):
        total = int(float(self.progress.cget('maximum')))
        self._enqueue_ui('progress', (value, total))

    def _download_finished(self, success: int, failed: int, skipped: int, out_path: Path):
        self.download_btn.config(state='normal')
        self.cancel_btn.config(state='disabled')
        completed = success + failed + skipped
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
