"""
ECMP Nebraska Public Access Bulk Downloader
--------------------------------------------
GUI tool for downloading documents from the Nebraska ECMP Public Access system.

Usage:
    python ecmp_bulk_downloader.py

Requirements:
    pip install requests
    (tkinter is included with standard Python installations)
"""

import os
import re
import csv
import random
import time
import queue
import threading
import requests
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timezone, timedelta
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ── Constants ─────────────────────────────────────────────────────────────────

BASE_URL   = "https://ecmp.nebraska.gov/PublicAccess"
SEARCH_URL = f"{BASE_URL}/api/CustomQuery/KeywordSearch"
DOC_URL    = f"{BASE_URL}/api/Document/{{token}}/"
QUERY_ID   = 340
DEQ_PROGRAM = "AIR"   # Fixed — not exposed in the GUI

KW_UNKNOWN    = 113
KW_FACILITY   = 114
KW_PROGRAM    = 115
KW_PROGRAM_ID = 116

DELAY_BETWEEN_DOWNLOADS = 0.5
MAX_WORKERS_CAP = 8

CONNECT_TIMEOUT = 10
DOWNLOAD_TIMEOUT = 60
DOWNLOAD_STREAM_TIMEOUT = (CONNECT_TIMEOUT, DOWNLOAD_TIMEOUT)

MAX_RETRIES = 3
RETRIABLE_STATUS_CODES = {429, 500, 502, 503, 504}
RETRY_BACKOFF_BASE = 0.5
RETRY_BACKOFF_CAP = 8.0

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "Referer": f"{BASE_URL}/index.html",
}


# ── API helpers ───────────────────────────────────────────────────────────────

def date_to_utc_iso(date_str: str) -> str:
    """Convert YYYY-MM-DD to a UTC ISO timestamp (treats input as midnight CST)."""
    dt  = datetime.strptime(date_str, "%Y-%m-%d")
    cst = timezone(timedelta(hours=-6))
    return dt.replace(tzinfo=cst).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def build_request_body(facility: str, from_date: str, to_date: str) -> dict:
    body = {
        "QueryID":  QUERY_ID,
        "Keywords": [
            {"ID": KW_UNKNOWN,    "Value": "",         "KeywordOperator": "="},
            {"ID": KW_FACILITY,   "Value": facility,   "KeywordOperator": "="},
            {"ID": KW_PROGRAM,    "Value": DEQ_PROGRAM,"KeywordOperator": "="},
            {"ID": KW_PROGRAM_ID, "Value": "",         "KeywordOperator": "="},
        ],
        "QueryLimit": 0,
    }
    if from_date:
        body["FromDate"] = date_to_utc_iso(from_date)
    if to_date:
        body["ToDate"] = date_to_utc_iso(to_date)
    return body


def fetch_document_list(facility: str, from_date: str, to_date: str) -> list[dict]:
    body     = build_request_body(facility, from_date, to_date)
    response = requests.post(SEARCH_URL, json=body, headers=HEADERS, timeout=30)
    response.raise_for_status()
    data = response.json()
    return data.get("Data", []), data.get("Truncated", False)


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name.strip(". ") or "document"


def get_server_filename(response: requests.Response) -> str:
    cd    = response.headers.get("Content-Disposition", "")
    match = re.search(r'filename[^;=\n]*=\s*["\']?([^"\';\n]+)', cd, re.IGNORECASE)
    return match.group(1).strip() if match else ""


class NDWEEDownloader:
    """Per-thread downloader with its own HTTP session."""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self._configure_retries()

    def _configure_retries(self):
        retry_cfg = Retry(
            total=2,
            connect=2,
            read=2,
            status=2,
            backoff_factor=0.4,
            status_forcelist=sorted(RETRIABLE_STATUS_CODES),
            allowed_methods=frozenset({"GET"}),
            respect_retry_after_header=True,
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry_cfg)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

    def fetch_document(self, token: str) -> requests.Response:
        url = DOC_URL.format(token=token)
        response = self.session.get(url, headers=HEADERS, timeout=DOWNLOAD_STREAM_TIMEOUT, stream=True)
        response.raise_for_status()
        return response

    @staticmethod
    def write_atomic_stream(response: requests.Response, out_path: str):
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        tmp_path = f"{out_path}.part"
        try:
            with open(tmp_path, "wb") as fh:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        fh.write(chunk)
            os.replace(tmp_path, out_path)
        except Exception:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
            raise


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ECMP Nebraska Bulk Downloader")
        self.geometry("980x760")
        self.resizable(False, False)
        self._msg_queue = queue.Queue()
        self._stop_event = threading.Event()
        self._dl_local = threading.local()
        self._rate_lock = threading.Lock()
        self._next_request_ts = 0.0
        self._name_lock = threading.Lock()
        self._reserved_names = set()
        self._result_records: list[dict] = []
        self._results_columns: list[str] = []
        self._results_truncated = False
        self._build_ui()
        self._poll_queue()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── Header
        header = tk.Frame(self, bg="#003366")
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(
            header,
            text="ECMP Nebraska  •  Bulk Document Downloader",
            bg="#003366", fg="white",
            font=("Helvetica", 13, "bold"),
            pady=10,
        ).pack()

        # ── Input frame
        frm = ttk.LabelFrame(self, text="Search Parameters", padding=12)
        frm.grid(row=1, column=0, sticky="ew", padx=16, pady=(14, 6))
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Facility Number:").grid(row=0, column=0, sticky="w", **pad)
        self.facility_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.facility_var, width=20).grid(
            row=0, column=1, sticky="w", **pad)

        ttk.Label(frm, text="From Date:").grid(row=1, column=0, sticky="w", **pad)
        self.from_var = tk.StringVar()
        from_entry = tk.Entry(frm, textvariable=self.from_var, width=20,
                              relief="flat", highlightthickness=1,
                              highlightbackground="#cccccc", highlightcolor="#0055a5")
        from_entry.grid(row=1, column=1, sticky="w", **pad)
        self._add_placeholder(from_entry, self.from_var, "YYYY-MM-DD  (optional)")

        ttk.Label(frm, text="To Date:").grid(row=2, column=0, sticky="w", **pad)
        self.to_var = tk.StringVar()
        to_entry = tk.Entry(frm, textvariable=self.to_var, width=20,
                            relief="flat", highlightthickness=1,
                            highlightbackground="#cccccc", highlightcolor="#0055a5")
        to_entry.grid(row=2, column=1, sticky="w", **pad)
        self._add_placeholder(to_entry, self.to_var, "YYYY-MM-DD  (optional)")

        ttk.Label(frm, text="Save to:").grid(row=3, column=0, sticky="w", **pad)
        dest_row = ttk.Frame(frm)
        dest_row.grid(row=3, column=1, sticky="ew", **pad)
        self.dest_var = tk.StringVar(value=os.path.expanduser("~/Downloads/ecmp_downloads"))
        ttk.Entry(dest_row, textvariable=self.dest_var, width=34).pack(side="left")
        ttk.Button(dest_row, text="Browse…", command=self._browse).pack(side="left", padx=(6, 0))

        # ── Action buttons
        action_row = ttk.Frame(self)
        action_row.grid(row=2, column=0, pady=(6, 4))
        self.search_btn = ttk.Button(
            action_row, text="Search", command=self._start_search, width=18)
        self.search_btn.pack(side="left", padx=(0, 6))
        self.download_selected_btn = ttk.Button(
            action_row, text="Download Selected", command=self._start_download_selected,
            width=18, state="disabled")
        self.download_selected_btn.pack(side="left", padx=(0, 6))
        self.stop_btn = ttk.Button(
            action_row, text="Stop", command=self._stop_download, width=10, state="disabled")
        self.stop_btn.pack(side="left")

        # ── Search results frame
        results_frm = ttk.LabelFrame(self, text="Search Results", padding=12)
        results_frm.grid(row=3, column=0, sticky="ew", padx=16, pady=(2, 4))

        results_actions = ttk.Frame(results_frm)
        results_actions.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(results_actions, text="Select All", command=self._select_all_results).pack(
            side="left", padx=(0, 6))
        ttk.Button(results_actions, text="Deselect All", command=self._deselect_all_results).pack(
            side="left")
        self.results_count_label = ttk.Label(results_actions, text="No results loaded")
        self.results_count_label.pack(side="right")

        tree_wrap = ttk.Frame(results_frm)
        tree_wrap.grid(row=1, column=0, sticky="nsew")
        results_frm.rowconfigure(1, weight=1)
        results_frm.columnconfigure(0, weight=1)

        self.results_tree = ttk.Treeview(tree_wrap, show="headings", selectmode="extended", height=12)
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        self.results_tree.bind("<<TreeviewSelect>>", lambda _: self._update_selection_count())
        self.results_tree.bind("<Configure>", self._on_results_tree_resize)

        tree_vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.results_tree.yview)
        tree_hsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)
        tree_vsb.grid(row=0, column=1, sticky="ns")
        tree_hsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        # ── Progress frame
        prog_frm = ttk.LabelFrame(self, text="Progress", padding=12)
        prog_frm.grid(row=4, column=0, sticky="ew", padx=16, pady=(4, 14))
        prog_frm.columnconfigure(0, weight=1)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            prog_frm, variable=self.progress_var,
            maximum=100, length=380, mode="determinate")
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))

        self.count_label = ttk.Label(prog_frm, text="")
        self.count_label.grid(row=1, column=0, sticky="w")

        self.status_label = ttk.Label(prog_frm, text="Ready.", foreground="#555555",
                                      wraplength=380, justify="left")
        self.status_label.grid(row=2, column=0, sticky="w", pady=(2, 0))

        self.columnconfigure(0, weight=1)

    # ── Placeholder text helper ───────────────────────────────────────────────

    def _add_placeholder(self, entry: tk.Entry, var: tk.StringVar, text: str):
        var.set(text)
        entry.config(foreground="white", background="#d0d0d0")

        def on_focus_in(e):
            if var.get() == text:
                var.set("")
            entry.config(foreground="black", background="white")

        def on_focus_out(e):
            if not var.get():
                var.set(text)
                entry.config(foreground="white", background="#d0d0d0")
            else:
                entry.config(foreground="black", background="white")

        entry.bind("<FocusIn>",  on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
        entry._placeholder = text

    def _get_date(self, var: tk.StringVar, entry_widget=None) -> str:
        """Return the date string, or empty string if it still shows placeholder."""
        val = var.get().strip()
        if not val or "optional" in val.lower():
            return ""
        return val

    # ── Browse ────────────────────────────────────────────────────────────────

    def _browse(self):
        folder = filedialog.askdirectory(title="Select download folder")
        if folder:
            self.dest_var.set(folder)

    # ── Results list helpers ────────────────────────────────────────────────

    @staticmethod
    def _find_record_key(records: list[dict], candidates: list[str]) -> str | None:
        for record in records:
            by_lower = {str(k).lower(): k for k in record.keys()}
            for candidate in candidates:
                found = by_lower.get(candidate.lower())
                if found:
                    return found
        return None

    def _derive_results_columns(self, records: list[dict]) -> list[tuple[str, str, int]]:
        name_key = self._find_record_key(records, ["Name"]) or "Name"
        return [(name_key, "Document Name", 900)]

    def _clear_results(self):
        self._result_records = []
        self._results_columns = []
        for iid in self.results_tree.get_children():
            self.results_tree.delete(iid)
        self.results_tree.configure(columns=())
        self.results_count_label.config(text="No results loaded")
        self.download_selected_btn.config(state="disabled")

    def _on_results_tree_resize(self, _event=None):
        if not self._results_columns:
            return
        name_col = self._results_columns[0]
        # Keep one-column table filling the visible widget width.
        width = max(200, self.results_tree.winfo_width() - 6)
        self.results_tree.column(name_col, width=width, stretch=True)

    def _populate_results(self, records: list[dict]):
        self._clear_results()
        self._result_records = list(records)

        column_defs = self._derive_results_columns(self._result_records)
        column_keys = [key for key, _, _ in column_defs]
        self._results_columns = column_keys

        self.results_tree.configure(columns=tuple(column_keys))
        for key, heading, width in column_defs:
            self.results_tree.heading(key, text=heading)
            self.results_tree.column(key, width=width, anchor="w", stretch=(key == column_keys[0]))

        for idx, record in enumerate(self._result_records):
            values = [str(record.get(key, "")) for key in column_keys]
            self.results_tree.insert("", "end", iid=str(idx), values=values)

        self._on_results_tree_resize()

        children = self.results_tree.get_children()
        if children:
            self.results_tree.selection_set(children)
        self._update_selection_count()
        self.download_selected_btn.config(state="normal" if children else "disabled")

    def _selected_records(self) -> list[dict]:
        selected = []
        for iid in self.results_tree.selection():
            try:
                selected.append(self._result_records[int(iid)])
            except (ValueError, IndexError):
                continue
        return selected

    def _select_all_results(self):
        children = self.results_tree.get_children()
        if children:
            self.results_tree.selection_set(children)
            self._update_selection_count()

    def _deselect_all_results(self):
        self.results_tree.selection_remove(self.results_tree.get_children())
        self._update_selection_count()

    def _update_selection_count(self):
        total = len(self._result_records)
        selected = len(self.results_tree.selection())
        if total == 0:
            self.results_count_label.config(text="No results loaded")
        else:
            self.results_count_label.config(text=f"{selected} of {total} selected")

    # ── Validation ────────────────────────────────────────────────────────────

    def _validate(self) -> bool:
        if not self.facility_var.get().strip():
            messagebox.showwarning("Missing input", "Please enter a Facility Number.")
            return False
        for label, var in [("From Date", self.from_var), ("To Date", self.to_var)]:
            val = self._get_date(var)
            if val:
                try:
                    datetime.strptime(val, "%Y-%m-%d")
                except ValueError:
                    messagebox.showerror(
                        "Invalid date",
                        f"{label} must be in YYYY-MM-DD format (e.g. 2024-01-15)."
                    )
                    return False
        if not self.dest_var.get().strip():
            messagebox.showwarning("Missing input", "Please select a download folder.")
            return False
        return True

    # ── Search / download flow ───────────────────────────────────────────────

    def _start_search(self):
        if not self._validate():
            return

        facility  = self.facility_var.get().strip()
        from_date = self._get_date(self.from_var)
        to_date   = self._get_date(self.to_var)

        self._stop_event.clear()
        self.stop_btn.config(state="disabled")
        self._next_request_ts = 0.0
        with self._name_lock:
            self._reserved_names.clear()

        self.search_btn.config(state="disabled")
        self.download_selected_btn.config(state="disabled")
        self._set_status("Querying ECMP API…")
        self.progress_var.set(0)
        self.count_label.config(text="")
        self._clear_results()

        threading.Thread(
            target=self._worker,
            args=(facility, from_date, to_date),
            daemon=True,
        ).start()

    def _start_download_selected(self):
        records = self._selected_records()
        if not records:
            messagebox.showwarning("Nothing selected", "Select at least one document to download.")
            return

        dest = self.dest_var.get().strip()
        if not dest:
            messagebox.showwarning("Missing input", "Please select a download folder.")
            return

        try:
            os.makedirs(dest, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("Output folder error", str(exc))
            return

        self._stop_event.clear()
        self._next_request_ts = 0.0
        with self._name_lock:
            self._reserved_names.clear()

        self.search_btn.config(state="disabled")
        self.download_selected_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progress_var.set(0)
        self.count_label.config(text=f"0 / {len(records)} files")
        self._set_status("Starting selected download…")
        self._do_download(records, dest)

    # ── Background worker ─────────────────────────────────────────────────────

    def _worker(self, facility, from_date, to_date):
        try:
            self._post(("status", "Querying ECMP — please wait…"))
            records, truncated = fetch_document_list(facility, from_date, to_date)

            if not records:
                self._post(("done_none",))
                return

            self._post(("results_loaded", records, truncated))

        except requests.RequestException as exc:
            self._post(("error", f"Network error:\n{exc}"))

    def _do_download(self, records, dest):
        """Called from the main thread. Spawns download thread."""
        threading.Thread(
            target=self._download_worker,
            args=(records, dest),
            daemon=True,
        ).start()

    def _stop_download(self):
        self._stop_event.set()
        self.stop_btn.config(state="disabled")
        self._set_status("Stop requested. Finishing in-flight downloads…")

    def _compute_worker_count(self, total_docs: int) -> int:
        cpus = os.cpu_count() or 2
        return max(1, min(MAX_WORKERS_CAP, total_docs, max(2, cpus * 2)))

    def _get_thread_downloader(self) -> NDWEEDownloader:
        downloader = getattr(self._dl_local, "downloader", None)
        if downloader is None:
            downloader = NDWEEDownloader()
            self._dl_local.downloader = downloader
        return downloader

    @staticmethod
    def _is_retriable_reason(reason: str) -> bool:
        return reason in {"timeout", "http_429", "http_5xx", "request_error"}

    @staticmethod
    def _classify_exception(exc: Exception) -> tuple[str, str]:
        if isinstance(exc, requests.Timeout):
            return ("timeout", str(exc))
        if isinstance(exc, requests.HTTPError):
            code = exc.response.status_code if exc.response is not None else None
            if code == 429:
                return ("http_429", f"HTTP {code}: {exc}")
            if code is not None and 500 <= code < 600:
                return ("http_5xx", f"HTTP {code}: {exc}")
            return ("http_error", f"HTTP {code}: {exc}")
        if isinstance(exc, requests.RequestException):
            return ("request_error", str(exc))
        return ("unexpected_error", str(exc))

    @staticmethod
    def _backoff_seconds(attempt: int) -> float:
        raw = RETRY_BACKOFF_BASE * (2 ** (attempt - 1))
        jitter = random.uniform(0, RETRY_BACKOFF_BASE)
        return min(RETRY_BACKOFF_CAP, raw + jitter)

    def _respect_global_delay(self):
        if DELAY_BETWEEN_DOWNLOADS <= 0:
            return
        with self._rate_lock:
            now = time.monotonic()
            wait_for = self._next_request_ts - now
            if wait_for > 0:
                time.sleep(wait_for)
            self._next_request_ts = time.monotonic() + DELAY_BETWEEN_DOWNLOADS

    def _reserve_output_name(self, dest: str, filename: str) -> str:
        root, ext = os.path.splitext(filename)
        n = 0
        with self._name_lock:
            while True:
                candidate = filename if n == 0 else f"{root} ({n}){ext}"
                candidate_path = os.path.join(dest, candidate)
                key = candidate.lower()
                if key not in self._reserved_names and not os.path.exists(candidate_path):
                    self._reserved_names.add(key)
                    return candidate
                n += 1

    def _release_output_name(self, filename: str):
        with self._name_lock:
            self._reserved_names.discard(filename.lower())

    def _build_base_name(self, record: dict, response: requests.Response, index: int) -> str:
        doc_name = record.get("Name", "")
        if doc_name:
            base_name = sanitize_filename(doc_name)
            server_name = get_server_filename(response)
            if server_name:
                _, s_ext = os.path.splitext(server_name)
                _, l_ext = os.path.splitext(base_name)
                if s_ext and not l_ext:
                    base_name += s_ext
            return base_name

        server_name = get_server_filename(response)
        if server_name:
            return sanitize_filename(server_name)
        return f"document_{index:04d}.bin"

    def _download_one(self, record: dict, dest: str, index: int, total: int) -> dict:
        token = str(record.get("ID", "")).strip()
        doc_name = record.get("Name", "")

        if self._stop_event.is_set():
            return {
                "status": "stopped",
                "reason": "stopped",
                "error": "Stopped before start",
                "attempts": 0,
                "row": record,
                "local_name": "",
            }

        if not token:
            return {
                "status": "fail",
                "reason": "missing_token",
                "error": "Document record did not include ID token",
                "attempts": 0,
                "row": record,
                "local_name": "",
            }

        last_reason = "unexpected_error"
        last_error = "Unknown error"
        last_attempt = 0
        local_name = ""

        for attempt in range(1, MAX_RETRIES + 2):
            last_attempt = attempt
            if self._stop_event.is_set():
                if local_name:
                    self._release_output_name(local_name)
                return {
                    "status": "stopped",
                    "reason": "stopped",
                    "error": "Stopped by user",
                    "attempts": attempt - 1,
                    "row": record,
                    "local_name": local_name,
                }

            response = None
            try:
                self._respect_global_delay()
                downloader = self._get_thread_downloader()
                response = downloader.fetch_document(token)
                base_name = self._build_base_name(record, response, index)
                local_name = self._reserve_output_name(dest, base_name)
                out_path = os.path.join(dest, local_name)
                downloader.write_atomic_stream(response, out_path)

                self._post(("status", f"Downloaded {index} of {total}:  {doc_name or local_name}"))
                return {
                    "status": "ok",
                    "reason": "success",
                    "error": "",
                    "attempts": attempt,
                    "row": record,
                    "local_name": local_name,
                }

            except Exception as exc:
                if local_name:
                    self._release_output_name(local_name)
                    local_name = ""
                last_reason, last_error = self._classify_exception(exc)

            finally:
                if response is not None:
                    response.close()

            if attempt <= MAX_RETRIES and self._is_retriable_reason(last_reason):
                delay = self._backoff_seconds(attempt)
                self._post((
                    "status",
                    f"Retrying {doc_name or token} ({attempt}/{MAX_RETRIES}) after {delay:.1f}s: {last_reason}",
                ))
                time.sleep(delay)
                continue

            break

        return {
            "status": "fail",
            "reason": last_reason,
            "error": last_error,
            "attempts": last_attempt,
            "row": record,
            "local_name": local_name,
        }

    def _run_download_pass(self, records: list[dict], dest: str) -> dict:
        total = len(records)
        workers = self._compute_worker_count(total)
        completed = 0
        ok = 0
        fail = 0
        stopped = 0
        failed_records = []

        self._post(("progress_reset", total))
        self._post(("status", f"Downloading {total} file(s) with {workers} worker(s)…"))

        cancelled_futures = False
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="ndwee-dl") as pool:
            futures = [
                pool.submit(self._download_one, record, dest, i, total)
                for i, record in enumerate(records, start=1)
            ]

            for future in as_completed(futures):
                if self._stop_event.is_set() and not cancelled_futures:
                    cancelled_futures = True
                    for pending in futures:
                        pending.cancel()

                if future.cancelled():
                    stopped += 1
                else:
                    try:
                        result = future.result()
                        status = result.get("status", "fail")
                        if status == "ok":
                            ok += 1
                        elif status == "stopped":
                            stopped += 1
                            failed_records.append(result)
                        else:
                            fail += 1
                            failed_records.append(result)
                    except Exception as exc:
                        fail += 1
                        failed_records.append({
                            "status": "fail",
                            "reason": "worker_crash",
                            "error": str(exc),
                            "attempts": 0,
                            "row": {},
                            "local_name": "",
                        })

                completed += 1
                self._post(("progress", completed, total))

        return {
            "total": total,
            "ok": ok,
            "fail": fail,
            "stopped": stopped,
            "failed_records": failed_records,
        }

    def _write_failed_rows(self, dest: str, failed_records: list[dict]) -> str | None:
        if not failed_records:
            return None

        out_path = os.path.join(dest, "failed_rows.csv")
        all_fields = []
        seen = set()
        for rec in failed_records:
            row = rec.get("row", {})
            for key in row.keys():
                if key not in seen:
                    seen.add(key)
                    all_fields.append(key)

        all_fields += ["_reason", "_error", "_attempts", "_local_name"]
        with open(out_path, "w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=all_fields)
            writer.writeheader()
            for rec in failed_records:
                out_row = dict(rec.get("row", {}))
                out_row["_reason"] = rec.get("reason", "")
                out_row["_error"] = rec.get("error", "")
                out_row["_attempts"] = rec.get("attempts", 0)
                out_row["_local_name"] = rec.get("local_name", "")
                writer.writerow(out_row)

        return out_path

    def _download_worker(self, records, dest):
        os.makedirs(dest, exist_ok=True)
        summary = self._run_download_pass(records, dest)
        failed_csv = self._write_failed_rows(dest, summary["failed_records"])
        summary["dest"] = dest
        summary["failed_csv"] = failed_csv
        self._post(("done", summary))

    # ── Queue / thread-safe UI updates ────────────────────────────────────────

    def _post(self, msg):
        """Send a message from the worker thread to the main thread."""
        self._msg_queue.put(msg)

    def _poll_queue(self):
        """Process messages posted by worker threads. Runs on the main thread."""
        try:
            while True:
                msg = self._msg_queue.get_nowait()
                kind = msg[0]

                if kind == "status":
                    self._set_status(msg[1])

                elif kind == "progress":
                    i, total = msg[1], msg[2]
                    pct = (i / total) * 100
                    self.progress_var.set(pct)
                    self.count_label.config(text=f"{i} / {total} files")

                elif kind == "progress_reset":
                    total = msg[1]
                    self.progress_var.set(0)
                    self.count_label.config(text=f"0 / {total} files")

                elif kind == "results_loaded":
                    _, records, truncated = msg
                    self._results_truncated = bool(truncated)
                    self._populate_results(records)
                    self.search_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    self._set_status(f"Found {len(records)} document(s). Select files to download.")
                    if self._results_truncated:
                        messagebox.showwarning(
                            "Results Truncated",
                            "Search results were truncated, so not all documents were returned.\n"
                            "Consider narrowing your date range."
                        )

                elif kind == "done":
                    summary = msg[1]
                    total = summary["total"]
                    ok = summary["ok"]
                    fail = summary["fail"]
                    stopped = summary["stopped"]
                    dest = summary["dest"]
                    failed_csv = summary.get("failed_csv")

                    self.progress_var.set(100 if total else 0)
                    self.count_label.config(text=f"{total} / {total} files")
                    self.search_btn.config(state="normal")
                    self.download_selected_btn.config(
                        state="normal" if self.results_tree.get_children() else "disabled"
                    )
                    self.stop_btn.config(state="disabled")

                    if stopped and not fail:
                        self._set_status(
                            f"Stopped. {ok} succeeded, {stopped} stopped.\nSaved to:\n{dest}"
                        )
                    elif fail or stopped:
                        extra = f"\nFailed rows file:\n{failed_csv}" if failed_csv else ""
                        self._set_status(
                            f"Finished with issues. {ok} succeeded, {fail} failed, {stopped} stopped.\nSaved to:\n{dest}{extra}"
                        )
                    else:
                        self._set_status(f"✓  Done. {ok} file(s) saved to:\n{dest}")

                elif kind == "done_none":
                    self._clear_results()
                    self._set_status("No documents found for the given filters.")
                    self.search_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")

                elif kind == "error":
                    self._set_status(f"Error: {msg[1]}")
                    self.search_btn.config(state="normal")
                    self.download_selected_btn.config(
                        state="normal" if self.results_tree.get_children() else "disabled"
                    )
                    self.stop_btn.config(state="disabled")
                    messagebox.showerror("Error", msg[1])

        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    def _set_status(self, text: str):
        self.status_label.config(text=text)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
