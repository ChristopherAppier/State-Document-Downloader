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
import time
import queue
import threading
import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timezone, timedelta


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


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ECMP Nebraska Bulk Downloader")
        self.resizable(False, False)
        self._msg_queue = queue.Queue()
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

        # ── Action button
        self.download_btn = ttk.Button(
            self, text="Search & Download", command=self._start, width=28)
        self.download_btn.grid(row=2, column=0, pady=(6, 4))

        # ── Progress frame
        prog_frm = ttk.LabelFrame(self, text="Progress", padding=12)
        prog_frm.grid(row=3, column=0, sticky="ew", padx=16, pady=(4, 14))
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

    # ── Start download flow ───────────────────────────────────────────────────

    def _start(self):
        if not self._validate():
            return

        facility  = self.facility_var.get().strip()
        from_date = self._get_date(self.from_var)
        to_date   = self._get_date(self.to_var)
        dest      = self.dest_var.get().strip()

        self.download_btn.config(state="disabled")
        self._set_status("Querying ECMP API…")
        self.progress_var.set(0)
        self.count_label.config(text="")

        threading.Thread(
            target=self._worker,
            args=(facility, from_date, to_date, dest),
            daemon=True,
        ).start()

    # ── Background worker ─────────────────────────────────────────────────────

    def _worker(self, facility, from_date, to_date, dest):
        try:
            # Step 1: fetch document list
            self._post(("status", "Querying ECMP — please wait…"))
            records, truncated = fetch_document_list(facility, from_date, to_date)

            if not records:
                self._post(("done_none",))
                return

            # Step 2: confirm with user before downloading
            self._post(("confirm", len(records), truncated, records, dest))

        except requests.RequestException as exc:
            self._post(("error", f"Network error:\n{exc}"))

    def _do_download(self, records, dest):
        """Called from the main thread after user confirms. Spawns download thread."""
        threading.Thread(
            target=self._download_worker,
            args=(records, dest),
            daemon=True,
        ).start()

    def _download_worker(self, records, dest):
        os.makedirs(dest, exist_ok=True)
        total = len(records)

        with requests.Session() as session:
            for i, record in enumerate(records, start=1):
                token    = record["ID"]
                doc_name = record.get("Name", "")
                url      = DOC_URL.format(token=token)

                self._post(("status", f"Downloading {i} of {total}:  {doc_name or '…'}"))
                self._post(("progress", i, total))

                try:
                    response = session.get(url, headers=HEADERS, timeout=60, stream=True)
                    response.raise_for_status()
                except requests.RequestException as exc:
                    self._post(("status", f"ERROR on file {i}: {exc}"))
                    time.sleep(DELAY_BETWEEN_DOWNLOADS)
                    continue

                # Build filename
                if doc_name:
                    base_name   = sanitize_filename(doc_name)
                    server_name = get_server_filename(response)
                    if server_name:
                        _, s_ext = os.path.splitext(server_name)
                        _, l_ext = os.path.splitext(base_name)
                        if s_ext and not l_ext:
                            base_name += s_ext
                else:
                    server_name = get_server_filename(response)
                    base_name   = sanitize_filename(server_name) if server_name else f"document_{i:04d}.bin"

                filepath = os.path.join(dest, base_name)
                if os.path.exists(filepath):
                    root, ext = os.path.splitext(base_name)
                    filepath  = os.path.join(dest, f"{root}_{i:04d}{ext}")

                with open(filepath, "wb") as fh:
                    for chunk in response.iter_content(chunk_size=8192):
                        fh.write(chunk)

                if i < total:
                    time.sleep(DELAY_BETWEEN_DOWNLOADS)

        self._post(("done", total, dest))

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

                elif kind == "confirm":
                    _, count, truncated, records, dest = msg
                    self.download_btn.config(state="normal")
                    warn = (
                        "\n\n⚠️  Results were truncated — not all documents returned.\n"
                        "Consider narrowing your date range."
                        if truncated else ""
                    )
                    proceed = messagebox.askyesno(
                        "Confirm Download",
                        f"Found {count} document(s) for the selected filters.{warn}\n\n"
                        f"Download all {count} file(s) to:\n{dest}\n\nProceed?",
                    )
                    if proceed:
                        self.download_btn.config(state="disabled")
                        self._set_status("Starting download…")
                        self.count_label.config(text=f"0 / {count} files")
                        self._do_download(records, dest)
                    else:
                        self._set_status("Download cancelled.")
                        self.progress_var.set(0)
                        self.count_label.config(text="")

                elif kind == "done":
                    _, total, dest = msg
                    self.progress_var.set(100)
                    self.count_label.config(text=f"{total} / {total} files")
                    self._set_status(f"✓  Done. {total} file(s) saved to:\n{dest}")
                    self.download_btn.config(state="normal")
                    messagebox.showinfo("Download Complete",
                                        f"Successfully downloaded {total} file(s).\n\nSaved to:\n{dest}")

                elif kind == "done_none":
                    self._set_status("No documents found for the given filters.")
                    self.download_btn.config(state="normal")

                elif kind == "error":
                    self._set_status(f"Error: {msg[1]}")
                    self.download_btn.config(state="normal")
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
