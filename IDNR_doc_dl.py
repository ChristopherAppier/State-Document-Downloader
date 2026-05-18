#!/usr/bin/env python3
"""
Iowa DNR Document Search – Batch Downloader
============================================
Loads a CSV export from https://programs.iowadnr.gov/docsearch/Home/Search,
resolves each document's OTCS objectID via the search API, and downloads
the files to a folder of your choice.

Dependencies (install once):
    pip install requests beautifulsoup4

Usage:
    python idnr_batch_downloader.py
"""

import csv
import datetime as dt
import os
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from queue import Empty, Queue

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ─────────────────────────────────────────────────────────────────────────────
#  Configuration — adjust if the site URL or endpoints change
# ─────────────────────────────────────────────────────────────────────────────
BASE_URL        = "https://programs.iowadnr.gov/docsearch"
SEARCH_PAGE_URL = f"{BASE_URL}/Home/Search"    # GET — loads form + cookies
SEARCH_URL      = f"{BASE_URL}/Home/OTCSSearch" # POST — executes search, returns results
DOWNLOAD_URL    = f"{BASE_URL}/Home/OTCSDownload"

# Maps CSV "Program" column values to the site's programfilter / aqsubfilter values.
# Add entries here if you work with programs outside Air Quality.
PROGRAM_FILTER_MAP: dict[str, dict[str, str]] = {
    "Compliance":           {"programfilter": "AirQuality", "aqsubfilter": "AQB - Compliance"},
    "Construction Permits": {"programfilter": "AirQuality", "aqsubfilter": "AQB - Construction Permits"},
    "Stack Testing":        {"programfilter": "AirQuality", "aqsubfilter": "AQB - Stack Testing"},
    "Miscellaneous":        {"programfilter": "AirQuality", "aqsubfilter": "AQB - Miscellaneous"},
    "Planning":             {"programfilter": "AirQuality", "aqsubfilter": "AQB - Planning"},
    "Title V Operating Permit": {"programfilter": "AirQuality", "aqsubfilter": "AQB - Title V Operating Permit"},
    "Small Source Operating Permit": {"programfilter": "AirQuality", "aqsubfilter": "AQB - Small Source Operating Permit"},
    "Voluntary Operating Permit":    {"programfilter": "AirQuality", "aqsubfilter": "AQB - Voluntary Operating Permit"},
}

REQUEST_TIMEOUT   = 30   # seconds per HTTP request
DOWNLOAD_TIMEOUT  = 120  # seconds for file download
DELAY_BETWEEN     = 0.5  # seconds between downloads (be polite)

CONNECT_TIMEOUT = 10
SEARCH_TIMEOUT = (CONNECT_TIMEOUT, REQUEST_TIMEOUT)
DOWNLOAD_STREAM_TIMEOUT = (CONNECT_TIMEOUT, DOWNLOAD_TIMEOUT)

MAX_RETRIES = 5
RETRIABLE_STATUS_CODES = {429, 500, 502, 503, 504}
RETRY_BACKOFF_BASE = 0.5
RETRY_BACKOFF_CAP = 8.0

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": f"{BASE_URL}/Home/Search",
}

# CSV column names as exported by the site
COL_VIEW   = "View Doc"
COL_PROG   = "Program"
COL_DOCID  = "Document ID"
COL_DATE   = "Document Date"
COL_TYPE   = "Document Type"
COL_NOTES  = "Notes"
COL_FAC    = "Facility ID"
COL_PERMIT = "Permit Number"
COL_NAME   = "Facility Name"
COL_PROJ   = "Project Number"
COL_CITY   = "City"
COL_COUNTY = "County"


# ─────────────────────────────────────────────────────────────────────────────
#  Downloader — handles HTTP session, objectID resolution, and file download
# ─────────────────────────────────────────────────────────────────────────────
class Downloader:
    def __init__(self, log_fn=None):
        self._log = log_fn or (lambda msg: None)
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self._configure_retries()
        self._csrf_token: str | None = None
        self._init_session()

    def _configure_retries(self):
        retry_cfg = Retry(
            total=2,
            connect=2,
            read=2,
            status=2,
            backoff_factor=0.4,
            status_forcelist=sorted(RETRIABLE_STATUS_CODES),
            allowed_methods=frozenset({"GET", "POST"}),
            respect_retry_after_header=True,
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry_cfg)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

    # ── Session bootstrap ─────────────────────────────────────────────────────
    def _init_session(self):
        """
        Fetch the search page once to pick up session cookies and any
        ASP.NET anti-forgery (CSRF) token.
        """
        try:
            resp = self.session.get(SEARCH_PAGE_URL, timeout=SEARCH_TIMEOUT)
            self._csrf_token = self._extract_csrf(resp.text)
            if self._csrf_token:
                self._log(f"[session] CSRF token acquired.")
            else:
                self._log(f"[session] No CSRF token found (may not be required).")
        except Exception as e:
            self._log(f"[session] Warning – could not initialise session: {e}")

    @staticmethod
    def _extract_csrf(html: str) -> str | None:
        """Return the ASP.NET RequestVerificationToken value, if present."""
        soup = BeautifulSoup(html, "html.parser")
        tag = soup.find("input", {"name": "__RequestVerificationToken"})
        return tag["value"] if tag else None

    # ── Filename helper ───────────────────────────────────────────────────────
    @staticmethod
    def filename_from_view(view_doc: str) -> str:
        """
        The CSV 'View Doc' cell reads 'View 032326-026.pdf'.
        Strip the leading 'View ' to get the bare filename.
        """
        return re.sub(r"^View\s+", "", view_doc).strip()

    # ── objectID resolution ───────────────────────────────────────────────────
    @staticmethod
    def _datatables_columns() -> list[tuple[str, str]]:
        """
        Build the fixed DataTables columns[] payload that OTCSSearch requires.
        Matches the 12 columns the browser sends (captured via DevTools).
        """
        col_defs = [
            ("name",          "false"),   # col 0 — filename link, not orderable
            ("program",       "true"),
            ("documentID",    "true"),
            ("documentDate",  "true"),
            ("documentType",  "true"),
            ("notes",         "true"),
            ("facilityID",    "true"),
            ("permitNumber",  "true"),
            ("facilityName",  "true"),
            ("projectNumber", "true"),
            ("facilityCity",  "true"),
            ("facilityCounty","true"),
        ]
        params: list[tuple[str, str]] = []
        for i, (data_field, orderable) in enumerate(col_defs):
            p = f"columns[{i}]"
            params += [
                (f"{p}[data]",          data_field),
                (f"{p}[name]",          ""),
                (f"{p}[searchable]",    "true"),
                (f"{p}[orderable]",     orderable),
                (f"{p}[search][value]", ""),
                (f"{p}[search][regex]", "false"),
            ]
        return params

    def find_object_id(
        self,
        filename: str,
        facility_id: str = "",
        doc_id: str = "",
        doc_type: str = "",
        program: str = "",
        notes: str = "",
        permit_number: str = "",
    ) -> str | None:
        """
        POST to OTCSSearch using the exact DataTables + viewModel payload the
        browser sends.  Returns the OTCS objectID string, or None.

        Strategy (most → least specific):
          1. Search by Document ID
          2. Filename stem as Document ID (for PDFs with no CSV doc ID)
          3. Notes value (e.g. variance request #) — very specific, small result set
          4. Facility ID + program + doc type
          5. Facility ID + program (no type filter)
        """
        # viewModel search field combinations to try, most → least specific
        queries = []
        if doc_id:
            queries.append({"viewModel[docidfilter]": doc_id})
        # For files without a Document ID, try the filename stem as the doc ID
        if not doc_id:
            stem = Path(filename).stem
            queries.append({"viewModel[docidfilter]": stem})

        # Notes-based search — very narrow when a unique notes value is available
        # (e.g. a variance request number shared by only a handful of documents)
        if notes:
            queries.append({"viewModel[notesfilter]": notes})

        # Facility ID fallback — narrow by program + doc type to avoid 500 errors
        # on large facilities that have hundreds of documents across all programs.
        if facility_id:
            prog_overrides = PROGRAM_FILTER_MAP.get(program, {})
            fac_base: dict[str, str] = {"viewModel[facilityIDfilter]": facility_id}
            if prog_overrides:
                fac_base["viewModel[programfilter]"] = prog_overrides["programfilter"]
                if "aqsubfilter" in prog_overrides:
                    fac_base["viewModel[aqsubfilter]"] = prog_overrides["aqsubfilter"]

            # Try with typefilter first for a narrower result set
            if doc_type:
                fac_with_type = {**fac_base, "viewModel[typefilter]": doc_type}
                queries.append(fac_with_type)

            # Fallback without typefilter — some type values trigger a 500
            queries.append(fac_base)

        if not queries:
            return None

        # Fixed DataTables scaffolding — same for every request
        dt_cols = self._datatables_columns()

        dt_base: list[tuple[str, str]] = [
            ("draw",               "1"),
            *dt_cols,
            ("order[0][column]",   "3"),
            ("order[0][dir]",      "desc"),
            ("order[0][name]",     ""),
            ("start",              "0"),
            ("length",             "3200"),   # fetch all results at once
            ("search[value]",      ""),
            ("search[regex]",      "false"),
        ]

        # viewModel defaults — all filters blank / wildcard
        vm_defaults: list[tuple[str, str]] = [
            ("viewModel[programfilter]",              "*"),
            ("viewModel[aqsubfilter]",                "*AQB*"),
            ("viewModel[damsafetysubfilter]",         "*LQB - Dam Safety*"),
            ("viewModel[floodplainsubfilter]",        "*LQB - Flood*"),
            ("viewModel[docidfilter]",                ""),
            ("viewModel[facilityIDfilter]",           ""),
            ("viewModel[facilityNamefilter]",         ""),
            ("viewModel[permitNumberfilter]",         ""),
            ("viewModel[projectNumberfilter]",        ""),
            ("viewModel[facilityCityfilter]",         ""),
            ("viewModel[facilityCountyfilter]",       ""),
            ("viewModel[typefilter]",                 ""),
            ("viewModel[notesfilter]",                ""),
            ("viewModel[documentDateFilter]",         ""),
            ("viewModel[documentDateEndFilter]",      ""),
            ("viewModel[documentTypeDescriptionFilter]", ""),
            ("viewModel[legalAttorneyFilter]",        ""),
            ("viewModel[legalAreaFilter]",            ""),
            ("viewModel[section]",                    ""),
            ("viewModel[township]",                   ""),
            ("viewModel[range]",                      ""),
            ("viewModel[limitFilter]",                "3200"),
        ]

        ajax_headers = {
            "X-Requested-With": "XMLHttpRequest",
            "Accept": "application/json, text/javascript, */*; q=0.01",
        }

        for search_fields in queries:
            # Merge: replace matching defaults with search-specific values
            vm = list(vm_defaults)
            for k, v in search_fields.items():
                for i, (pk, _) in enumerate(vm):
                    if pk == k:
                        vm[i] = (k, v)
                        break

            payload = dt_base + vm

            try:
                resp = self.session.post(
                    SEARCH_URL,
                    data=payload,
                    headers=ajax_headers,
                    timeout=SEARCH_TIMEOUT,
                )
                if resp.status_code != 200:
                    self._log(f"          [POST OTCSSearch] HTTP {resp.status_code}")
                    continue

                oid = self._parse_datatables_json(resp.text, filename)
                if oid:
                    return oid

                # Fallback: scan raw text for any OTCSDownload link with this filename
                oid = self._parse_object_id(resp.text, filename)
                if oid:
                    return oid

            except requests.RequestException as e:
                self._log(f"          [POST OTCSSearch] request error: {e}")
                continue

        return None

    @staticmethod
    def _parse_object_id(html: str, filename: str) -> str | None:
        """
        Look for OTCSDownload hrefs that contain *filename* and extract
        the objectID query parameter.
        """
        soup = BeautifulSoup(html, "html.parser")

        # 1. Anchor tags with href containing OTCSDownload
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if "OTCSDownload" in href and filename.lower() in href.lower():
                m = re.search(r"[?&]objectID=(\d+)", href, re.IGNORECASE)
                if m:
                    return m.group(1)

        # 2. Any attribute in the page that looks like an OTCSDownload URL
        pattern = (
            r"OTCSDownload[^\"']*objectID=(\d+)[^\"']*"
            + re.escape(filename)
        )
        m = re.search(pattern, html, re.IGNORECASE)
        if m:
            return m.group(1)

        # Reversed order (objectID before filename)
        pattern2 = (
            r"objectID=(\d+)[^\"'<>]*"
            + re.escape(filename)
        )
        m2 = re.search(pattern2, html, re.IGNORECASE)
        if m2:
            return m2.group(1)

        return None

    @staticmethod
    def _parse_datatables_json(text: str, filename: str) -> str | None:
        """
        Parse a DataTables server-side JSON response from OTCSSearch.

        Actual response shape (confirmed via DevTools):
          {"draw": N, "recordsTotal": N, "recordsFiltered": N,
           "data": [{"object_ID": "10461646", "name": "03272026ACC7001005.zip",
                     "program": "...", ...}, ...]}

        The 'name' field is the bare filename string.
        The objectID is in the 'object_ID' field (note: underscore + capital ID).
        """
        import json as _json
        try:
            payload = _json.loads(text)
        except Exception:
            return None

        rows = payload.get("data", [])
        if not isinstance(rows, list):
            return None

        for row in rows:
            name_cell = str(row.get("name", ""))

            # Match rows where name equals our target filename (case-insensitive)
            if name_cell.lower() != filename.lower():
                continue

            # Primary key as observed in live responses
            for key in ("object_ID", "objectID", "objectId", "object_id", "id"):
                val = row.get(key)
                if val:
                    return str(val)

            # If objectID is embedded in an OTCSDownload URL within name_cell
            m = re.search(r"objectID=(\d+)", name_cell, re.IGNORECASE)
            if m:
                return m.group(1)

        return None

    # ── File download ─────────────────────────────────────────────────────────
    def download_by_id(
        self,
        object_id: str,
        filename: str,
        output_dir: Path,
        save_name: str | None = None,
    ) -> Path:
        """Download using a known objectID."""
        url = f"{DOWNLOAD_URL}?objectID={object_id}&name={filename}"
        local_name = save_name or filename
        return self._stream_download(url, local_name, output_dir)

    def download_by_name(
        self,
        filename: str,
        output_dir: Path,
        save_name: str | None = None,
    ) -> Path | None:
        """
        Fallback: attempt download without an objectID.
        Some OTCS deployments accept a name-only request.
        Returns the saved path, or None if the response looks like an error page.
        """
        url = f"{DOWNLOAD_URL}?name={filename}"
        resp = self.session.get(url, stream=True, timeout=DOWNLOAD_STREAM_TIMEOUT)
        resp.raise_for_status()
        ct = resp.headers.get("Content-Type", "")
        # Reject HTML error pages
        if "text/html" not in ct:
            out = output_dir / (save_name or filename)
            self._write_atomic_stream(resp, out)
            return out
        return None

    def _stream_download(self, url: str, filename: str, output_dir: Path) -> Path:
        resp = self.session.get(url, stream=True, timeout=DOWNLOAD_STREAM_TIMEOUT)
        resp.raise_for_status()
        out = output_dir / filename
        self._write_atomic_stream(resp, out)
        return out

    @staticmethod
    def _write_atomic_stream(resp: requests.Response, out: Path):
        out.parent.mkdir(parents=True, exist_ok=True)
        tmp = out.with_suffix(out.suffix + ".part")
        try:
            with open(tmp, "wb") as f:
                for chunk in resp.iter_content(65536):
                    if chunk:
                        f.write(chunk)
            os.replace(tmp, out)
        except Exception:
            try:
                if tmp.exists():
                    tmp.unlink()
            except Exception:
                pass
            raise


# ─────────────────────────────────────────────────────────────────────────────
#  Tkinter UI
# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    # Treeview columns to display (must be a subset of CSV columns)
    DISPLAY_COLS = (
        COL_VIEW, COL_PROG, COL_DATE, COL_TYPE,
        COL_NAME, COL_FAC, COL_COUNTY,
    )
    COL_WIDTHS = {
        COL_VIEW: 220, COL_PROG: 115, COL_DATE: 85,
        COL_TYPE: 165, COL_NAME: 200, COL_FAC: 75, COL_COUNTY: 80,
    }

    def __init__(self):
        super().__init__()
        self.title("Iowa DNR – Batch Document Downloader")
        self.geometry("1060x720")
        self.minsize(800, 550)
        self.configure(bg="#f2f2f2")

        self.rows: list[dict]            = []
        self.output_dir  = tk.StringVar(value=str(Path.home() / "Downloads"))
        self.status_q: Queue             = Queue()
        self._stop_event = threading.Event()

        self._dl_local = threading.local()  # one Downloader per worker thread
        self._rate_lock = threading.Lock()
        self._next_request_ts = 0.0
        self._name_lock = threading.Lock()
        self._reserved_names: set[str] = set()
        self._log_lock = threading.Lock()
        self._run_log_path: Path | None = None

        self._build_ui()
        self._poll_queue()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # ── Step 1: CSV load ──────────────────────────────────────────────────
        top = ttk.LabelFrame(self, text="Step 1 – Load CSV export from Document Search", padding=8)
        top.pack(fill="x", padx=12, pady=(12, 4))

        ttk.Button(top, text="Browse CSV…", command=self._load_csv).pack(side="left", **pad)
        self.csv_label = ttk.Label(top, text="No file loaded", foreground="gray")
        self.csv_label.pack(side="left", **pad)

        # ── Step 2: Document table ────────────────────────────────────────────
        mid = ttk.LabelFrame(self, text="Step 2 – Select documents to download", padding=8)
        mid.pack(fill="both", expand=True, padx=12, pady=4)

        btn_row = ttk.Frame(mid)
        btn_row.pack(fill="x", pady=(0, 4))
        ttk.Button(btn_row, text="Select All",   command=self._select_all  ).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Deselect All", command=self._deselect_all).pack(side="left", padx=4)
        self.count_lbl = ttk.Label(btn_row, text="0 documents loaded")
        self.count_lbl.pack(side="right", padx=6)

        # Inner frame keeps grid geometry isolated from the pack-managed btn_row above
        tree_frame = ttk.Frame(mid)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_frame, columns=self.DISPLAY_COLS, show="headings", selectmode="extended"
        )
        hdrs = {COL_VIEW: "Filename"}
        for c in self.DISPLAY_COLS:
            self.tree.heading(c, text=hdrs.get(c, c))
            self.tree.column(c, width=self.COL_WIDTHS.get(c, 100), anchor="w", stretch=(c == COL_NAME))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", lambda _: self._update_selection_count())

        # ── Step 3: Output folder ─────────────────────────────────────────────
        out_frame = ttk.LabelFrame(self, text="Step 3 – Output folder", padding=8)
        out_frame.pack(fill="x", padx=12, pady=4)

        self.out_entry = ttk.Entry(out_frame, textvariable=self.output_dir, width=72)
        self.out_entry.pack(side="left", **pad)
        ttk.Button(out_frame, text="Browse…", command=self._pick_output).pack(side="left", **pad)

        # ── Download controls ─────────────────────────────────────────────────
        ctrl = ttk.Frame(self)
        ctrl.pack(fill="x", padx=12, pady=4)

        self.dl_btn = ttk.Button(
            ctrl, text="⬇  Download Selected", command=self._start_download
        )
        self.dl_btn.pack(side="left", padx=4)

        self.stop_btn = ttk.Button(
            ctrl, text="⏹  Stop", command=self._stop_download, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=4)

        self.progress = ttk.Progressbar(ctrl, length=380, mode="determinate")
        self.progress.pack(side="left", padx=8)

        self.prog_lbl = ttk.Label(ctrl, text="")
        self.prog_lbl.pack(side="left")

    # ── CSV loading ───────────────────────────────────────────────────────────
    def _load_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV export from Iowa DNR Document Search",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return

        self.rows.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if not any(v.strip() for v in row.values()):
                        continue
                    if not row.get(COL_VIEW, "").strip():
                        continue
                    self.rows.append(dict(row))
        except Exception as e:
            messagebox.showerror("CSV Error", f"Could not read file:\n{e}")
            return

        n = len(self.rows)
        self.csv_label.configure(
            text=f"{Path(path).name}  —  {n} document(s)", foreground="black"
        )
        self._populate_tree()
        self._select_all()
        self._write_log(f"Loaded {n} rows from {Path(path).name}\n")

    def _populate_tree(self):
        for idx, row in enumerate(self.rows):
            filename = Downloader.filename_from_view(row.get(COL_VIEW, ""))
            values = (
                filename,
                row.get(COL_PROG,   ""),
                row.get(COL_DATE,   ""),
                row.get(COL_TYPE,   ""),
                row.get(COL_NAME,   ""),
                row.get(COL_FAC,    ""),
                row.get(COL_COUNTY, ""),
            )
            self.tree.insert("", "end", iid=str(idx), values=values)

    # ── Tree interaction ──────────────────────────────────────────────────────
    def _update_selection_count(self):
        total = len(self.rows)
        selected = len(self.tree.selection())
        if total == 0:
            self.count_lbl.configure(text="0 documents loaded")
        else:
            self.count_lbl.configure(text=f"{selected} of {total} selected")

    def _select_all(self):
        self.tree.selection_set(self.tree.get_children())
        self._update_selection_count()

    def _deselect_all(self):
        self.tree.selection_remove(self.tree.get_children())
        self._update_selection_count()

    # ── Output folder ─────────────────────────────────────────────────────────
    def _pick_output(self):
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.output_dir.set(d)

    # ── Download ──────────────────────────────────────────────────────────────
    def _start_download(self):
        selected = [self.rows[int(iid)] for iid in self.tree.selection()]
        if not selected:
            messagebox.showwarning("Nothing selected", "Check at least one document.")
            return

        out = Path(self.output_dir.get())
        try:
            out.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Output folder error", str(e))
            return

        self._stop_event.clear()
        self.dl_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress["value"] = 0
        self.progress["maximum"] = len(selected)
        self.prog_lbl.configure(text="")
        self._next_request_ts = 0.0
        with self._name_lock:
            self._reserved_names.clear()
        self._start_run_log(out, len(selected))

        self._write_log("Starting parallel downloader…\n")

        threading.Thread(
            target=self._download_worker,
            args=(selected, out),
            daemon=True,
        ).start()

    def _stop_download(self):
        self._stop_event.set()
        self.stop_btn.configure(state="disabled")
        self._write_log("Stop requested — finishing current file…\n")

    def _compute_worker_count(self, selected_count: int) -> int:
        cpus = os.cpu_count() or 2
        return max(1, min(8, selected_count, max(2, cpus * 2)))

    def _get_thread_downloader(self) -> Downloader:
        dl = getattr(self._dl_local, "downloader", None)
        if dl is None:
            dl = Downloader(log_fn=self._q_log)
            self._dl_local.downloader = dl
        return dl

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

    def _start_run_log(self, out: Path, total: int):
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        self._run_log_path = out / f"idnr_run_{stamp}.log"
        with self._log_lock:
            with open(self._run_log_path, "w", encoding="utf-8") as f:
                f.write(f"Run started: {dt.datetime.now().isoformat()}\n")
                f.write(f"Selected rows: {total}\n")
                f.write(f"Max retries per file: {MAX_RETRIES}\n\n")

    def _write_failed_rows(self, out: Path, failed_records: list[dict]) -> Path | None:
        if not failed_records:
            return None

        path = out / "failed_rows.csv"
        all_fields: list[str] = []
        seen: set[str] = set()
        for rec in failed_records:
            for key in rec["row"].keys():
                if key not in seen:
                    seen.add(key)
                    all_fields.append(key)
        all_fields += ["_reason", "_error", "_attempts", "_local_name"]

        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=all_fields)
            writer.writeheader()
            for rec in failed_records:
                out_row = dict(rec["row"])
                out_row["_reason"] = rec["reason"]
                out_row["_error"] = rec["error"]
                out_row["_attempts"] = rec["attempts"]
                out_row["_local_name"] = rec["local_name"]
                writer.writerow(out_row)
        return path

    def _load_failed_rows(self, failed_path: Path) -> list[dict]:
        if not failed_path.exists():
            return []

        rows: list[dict] = []
        with open(failed_path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                cleaned = {k: v for k, v in row.items() if not k.startswith("_")}
                if cleaned.get(COL_VIEW, "").strip():
                    rows.append(cleaned)
        return rows

    def _respect_global_delay(self):
        if DELAY_BETWEEN <= 0:
            return
        with self._rate_lock:
            now = time.monotonic()
            wait_for = self._next_request_ts - now
            if wait_for > 0:
                time.sleep(wait_for)
            self._next_request_ts = time.monotonic() + DELAY_BETWEEN

    def _reserve_output_name(self, output_dir: Path, filename: str) -> str:
        stem = Path(filename).stem
        suffix = Path(filename).suffix
        n = 0

        with self._name_lock:
            while True:
                candidate = filename if n == 0 else f"{stem} ({n}){suffix}"
                candidate_key = candidate.lower()
                candidate_path = output_dir / candidate
                if candidate_key not in self._reserved_names and not candidate_path.exists():
                    self._reserved_names.add(candidate_key)
                    return candidate
                n += 1

    def _release_output_name(self, filename: str):
        with self._name_lock:
            self._reserved_names.discard(filename.lower())

    def _download_one(self, row: dict, out: Path, index: int, total: int) -> dict:
        if self._stop_event.is_set():
            return {
                "status": "stopped",
                "reason": "stopped",
                "filename": "",
                "row": row,
                "error": "Stopped before start",
                "attempts": 0,
                "local_name": "",
            }

        filename      = Downloader.filename_from_view(row.get(COL_VIEW,   ""))
        doc_id        = row.get(COL_DOCID,  "").strip()
        fac_id        = row.get(COL_FAC,    "").strip()
        doc_type      = row.get(COL_TYPE,   "").strip()
        program       = row.get(COL_PROG,   "").strip()
        notes         = row.get(COL_NOTES,  "").strip()
        permit_number = row.get(COL_PERMIT, "").strip()
        local_name = self._reserve_output_name(out, filename)

        self._q_log(f"[{index}/{total}] {filename}")
        self._q_log(f"          Facility={fac_id or '—'}  DocID={doc_id or '—'}  Type={doc_type or '—'}")

        last_reason = "unexpected_error"
        last_error = "Unknown failure"
        last_attempt = 0

        for attempt in range(1, MAX_RETRIES + 2):
            last_attempt = attempt
            if self._stop_event.is_set():
                self._release_output_name(local_name)
                return {
                    "status": "stopped",
                    "reason": "stopped",
                    "filename": filename,
                    "row": row,
                    "error": "Stopped by user",
                    "attempts": attempt - 1,
                    "local_name": local_name,
                }

            try:
                self._q_log(f"          Attempt {attempt}/{MAX_RETRIES + 1}: resolving objectID…")
                downloader = self._get_thread_downloader()
                self._respect_global_delay()
                obj_id = downloader.find_object_id(
                    filename, fac_id, doc_id, doc_type, program, notes, permit_number
                )

                if obj_id:
                    self._q_log(f"          objectID={obj_id}  →  downloading…")
                    self._respect_global_delay()
                    path = downloader.download_by_id(obj_id, filename, out, save_name=local_name)
                    if local_name != filename:
                        self._q_log(f"          ↳ Saved as '{local_name}' to avoid overwrite")
                    self._q_log(f"          ✓ Saved: {path.name}  ({path.stat().st_size:,} bytes)")
                    return {
                        "status": "ok",
                        "reason": "success",
                        "filename": filename,
                        "row": row,
                        "error": "",
                        "attempts": attempt,
                        "local_name": local_name,
                    }

                self._q_log("          objectID not found — trying name-only download…")
                self._respect_global_delay()
                path = downloader.download_by_name(filename, out, save_name=local_name)
                if path:
                    if local_name != filename:
                        self._q_log(f"          ↳ Saved as '{local_name}' to avoid overwrite")
                    self._q_log(f"          ✓ Saved (fallback): {path.name}  ({path.stat().st_size:,} bytes)")
                    return {
                        "status": "ok",
                        "reason": "success",
                        "filename": filename,
                        "row": row,
                        "error": "",
                        "attempts": attempt,
                        "local_name": local_name,
                    }

                last_reason = "objectid_not_found"
                last_error = "Could not resolve objectID and name-only fallback returned no file"

            except Exception as e:
                last_reason, last_error = self._classify_exception(e)
                self._q_log(f"          ✗ Attempt {attempt} failed ({last_reason}): {last_error}")

            if attempt <= MAX_RETRIES and self._is_retriable_reason(last_reason):
                delay = self._backoff_seconds(attempt)
                self._q_log(f"          Retrying in {delay:.2f}s…")
                time.sleep(delay)
                continue

            break

        self._release_output_name(local_name)
        return {
            "status": "fail",
            "reason": last_reason,
            "filename": filename,
            "row": row,
            "error": last_error,
            "attempts": last_attempt,
            "local_name": local_name,
        }

    def _run_download_pass(self, rows: list[dict], out: Path, pass_label: str) -> dict:
        total = len(rows)
        self.status_q.put(("progress_reset", total))

        ok = fail = stopped = completed = 0
        failed_records: list[dict] = []
        workers = self._compute_worker_count(total)
        self._q_log(f"{pass_label}: using {workers} worker thread(s) for {total} row(s).")

        cancelled_futures = False

        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="idnr-dl") as pool:
            futures = [
                pool.submit(self._download_one, row, out, i, total)
                for i, row in enumerate(rows, 1)
            ]

            for fut in as_completed(futures):
                if self._stop_event.is_set() and not cancelled_futures:
                    cancelled_futures = True
                    for pending in futures:
                        pending.cancel()

                if fut.cancelled():
                    stopped += 1
                else:
                    try:
                        result = fut.result()
                        status = result.get("status", "fail")
                        if status == "ok":
                            ok += 1
                        elif status == "stopped":
                            stopped += 1
                            failed_records.append(result)
                        else:
                            fail += 1
                            failed_records.append(result)
                    except Exception as e:
                        self._q_log(f"          ✗ ERROR – worker crashed: {e}")
                        fail += 1

                completed += 1
                self.status_q.put(("progress", completed))

        return {
            "ok": ok,
            "fail": fail,
            "stopped": stopped,
            "total": total,
            "failed_records": failed_records,
        }

    def _download_worker(self, rows: list[dict], out: Path):
        self._q_log("PASS 1: starting selected rows…")
        pass1 = self._run_download_pass(rows, out, "PASS 1")

        pass2 = {
            "ok": 0,
            "fail": 0,
            "stopped": 0,
            "total": 0,
            "failed_records": [],
        }

        failed_path = self._write_failed_rows(out, pass1["failed_records"])
        remaining_failed = pass1["failed_records"]

        if failed_path and not self._stop_event.is_set():
            try:
                retry_rows = self._load_failed_rows(failed_path)
            except Exception as e:
                self._q_log(f"Could not reload failed_rows.csv for pass 2: {e}")
                retry_rows = []

            if retry_rows:
                with self._name_lock:
                    self._reserved_names.clear()

                self._q_log(f"PASS 2: retrying {len(retry_rows)} failed row(s) from failed_rows.csv…")
                pass2 = self._run_download_pass(retry_rows, out, "PASS 2")
                remaining_failed = pass2["failed_records"]

                if remaining_failed:
                    self._write_failed_rows(out, remaining_failed)
                    self._q_log(f"Failed rows after pass 2 saved to: {failed_path}")
                else:
                    try:
                        failed_path.unlink()
                        self._q_log("PASS 2 recovered all failed rows; removed failed_rows.csv")
                    except Exception as e:
                        self._q_log(f"PASS 2 recovered all rows, but could not remove failed_rows.csv: {e}")
            else:
                self._q_log("PASS 2: no retryable rows found in failed_rows.csv")

        elif failed_path:
            self._q_log(f"Failed rows saved to: {failed_path}")

        if self._stop_event.is_set():
            self._q_log(f"\n{'─' * 55}\nStopped by user.\n")

        final_fail = len(remaining_failed)
        total_ok = pass1["ok"] + pass2["ok"]
        total_stopped = pass1["stopped"] + pass2["stopped"]
        initial_total = pass1["total"]

        if self._run_log_path:
            self._q_log(f"Run log saved to: {self._run_log_path}")

        self._q_log(
            f"\n{'─' * 55}\n"
            f"Pass 1:    {pass1['ok']} succeeded,  {pass1['fail']} failed,  {pass1['stopped']} stopped  (of {pass1['total']})\n"
            f"Pass 2:    {pass2['ok']} recovered,   {pass2['fail']} failed,  {pass2['stopped']} stopped  (of {pass2['total']} retried)\n"
            f"Final:     {total_ok} succeeded total,  {final_fail} remaining failed,  {total_stopped} stopped  (of {initial_total} selected)\n"
            f"Output:    {out}\n"
        )
        self.status_q.put(("done", None))

    # ── Queue polling (runs on main thread) ───────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                kind, value = self.status_q.get_nowait()
                if kind == "log":
                    self._write_log(value + "\n")
                elif kind == "progress_reset":
                    total = int(value)
                    self.progress["value"] = 0
                    self.progress["maximum"] = max(1, total)
                    self.prog_lbl.configure(text=f"0 / {total}")
                elif kind == "progress":
                    self.progress["value"] = value
                    total = int(self.progress["maximum"])
                    self.prog_lbl.configure(text=f"{value} / {total}")
                elif kind == "done":
                    self.dl_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
        except Empty:
            pass
        self.after(100, self._poll_queue)

    def _q_log(self, msg: str):
        self.status_q.put(("log", msg))

    def _write_log(self, msg: str):
        text = msg.rstrip("\n")
        if self._run_log_path is None or not text:
            return
        stamp = dt.datetime.now().strftime("%H:%M:%S")
        with self._log_lock:
            with open(self._run_log_path, "a", encoding="utf-8") as f:
                f.write(f"[{stamp}] {text}\n")


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
