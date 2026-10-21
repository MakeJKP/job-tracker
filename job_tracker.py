#!/usr/bin/env python3
"""
Job Application Tracker
=======================
A desktop app to track job applications during your search.
Uses SQLite (no Access required) and Python's built-in tkinter GUI.

Author: Your Name
GitHub: https://github.com/yourusername/job-tracker
"""

import os
import sys
import csv
import sqlite3
import subprocess
import webbrowser
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

# Optional: python-docx for Word document creation
try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# Windows-specific enhancements (gracefully skipped on Mac/Linux)
try:
    import windows_utils as wu
    HAS_WIN_UTILS = True
except ImportError:
    HAS_WIN_UTILS = False

# ============================================================
# CONFIGURATION
# ============================================================

# Database stored in user's home directory
DB_FILE = os.path.join(os.path.expanduser("~"), "job_tracker.db")

# Default folder path — use Win32 API on Windows for reliable path detection
if HAS_WIN_UTILS and sys.platform == "win32":
    DEFAULT_JOBS_FOLDER = wu.find_jobs_folder()
else:
    _onedrive = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "jobapplications")
    _desktop  = os.path.join(os.path.expanduser("~"), "Desktop", "jobapplications")
    DEFAULT_JOBS_FOLDER = _onedrive if os.path.exists(_onedrive) else _desktop

STATUS_OPTIONS = [
    "Applied",
    "Phone Screen Scheduled",
    "Interview Scheduled",
    "Final Round",
    "Offer Received",
    "Rejected",
    "Withdrawn",
    "Ghosted",
]

STATUS_COLORS = {
    "Applied":                  "#007200",
    "Phone Screen Scheduled":   "#0055A5",
    "Interview Scheduled":      "#0055A5",
    "Final Round":              "#C46000",
    "Offer Received":           "#00843D",
    "Rejected":                 "#A30000",
    "Withdrawn":                "#A30000",
    "Ghosted":                  "#606060",
}

APP_TITLE  = "Job Application Tracker"
WIN_WIDTH  = 900
WIN_HEIGHT = 780
FONT_BODY  = ("Segoe UI", 10)
FONT_LABEL = ("Segoe UI", 9)
FONT_HEAD  = ("Segoe UI", 11, "bold")
FONT_SMALL = ("Segoe UI", 9)


# ============================================================
# DATABASE
# ============================================================

class Database:
    """All SQLite operations for the job tracker."""

    def __init__(self, db_file=DB_FILE):
        self.db_file = db_file
        self._initialize()

    def _conn(self):
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        return conn

    def _initialize(self):
        conn = self._conn()
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                company_name    TEXT,
                position_title  TEXT,
                location        TEXT,
                job_url         TEXT,
                date_posted     TEXT,
                date_applied    TEXT,
                source          TEXT,
                status          TEXT DEFAULT 'Applied',
                status_date     TEXT,
                follow_up_date  TEXT,
                salary_range    TEXT,
                notes           TEXT,
                documents       TEXT,
                job_description TEXT,
                created_at      TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS archive (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                company_name    TEXT,
                position_title  TEXT,
                location        TEXT,
                job_url         TEXT,
                date_posted     TEXT,
                date_applied    TEXT,
                source          TEXT,
                status          TEXT,
                status_date     TEXT,
                follow_up_date  TEXT,
                salary_range    TEXT,
                notes           TEXT,
                documents       TEXT,
                job_description TEXT,
                date_archived   TEXT,
                created_at      TEXT
            )
        """)
        conn.commit()
        conn.close()

    # ---- JOBS ----

    def get_all_jobs(self):
        conn = self._conn()
        c = conn.cursor()
        c.execute("""
            SELECT * FROM jobs
            ORDER BY
                CASE status WHEN 'Applied' THEN 1 ELSE 2 END,
                CASE WHEN date_applied IS NULL OR date_applied='' THEN 1 ELSE 0 END,
                date_applied DESC
        """)
        rows = [dict(r) for r in c.fetchall()]
        conn.close()
        return rows

    def get_job(self, job_id):
        conn = self._conn()
        c = conn.cursor()
        c.execute("SELECT * FROM jobs WHERE id=?", (job_id,))
        row = c.fetchone()
        conn.close()
        return dict(row) if row else None

    def insert_job(self, data):
        conn = self._conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO jobs
                (company_name, position_title, location, job_url, date_posted,
                 date_applied, source, status, status_date, follow_up_date,
                 salary_range, notes, documents, job_description)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            data.get("company_name", ""),  data.get("position_title", ""),
            data.get("location", ""),      data.get("job_url", ""),
            data.get("date_posted", ""),   data.get("date_applied", ""),
            data.get("source", ""),        data.get("status", "Applied"),
            data.get("status_date", ""),   data.get("follow_up_date", ""),
            data.get("salary_range", ""),  data.get("notes", ""),
            data.get("documents", ""),     data.get("job_description", ""),
        ))
        new_id = c.lastrowid
        conn.commit()
        conn.close()
        return new_id

    def update_job(self, job_id, data):
        conn = self._conn()
        c = conn.cursor()
        c.execute("""
            UPDATE jobs SET
                company_name=?, position_title=?, location=?, job_url=?,
                date_posted=?, date_applied=?, source=?, status=?,
                status_date=?, follow_up_date=?, salary_range=?,
                notes=?, documents=?, job_description=?
            WHERE id=?
        """, (
            data.get("company_name", ""),  data.get("position_title", ""),
            data.get("location", ""),      data.get("job_url", ""),
            data.get("date_posted", ""),   data.get("date_applied", ""),
            data.get("source", ""),        data.get("status", "Applied"),
            data.get("status_date", ""),   data.get("follow_up_date", ""),
            data.get("salary_range", ""),  data.get("notes", ""),
            data.get("documents", ""),     data.get("job_description", ""),
            job_id,
        ))
        conn.commit()
        conn.close()

    def archive_job(self, job_id):
        conn = self._conn()
        c = conn.cursor()
        c.execute("SELECT * FROM jobs WHERE id=?", (job_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return False
        job = dict(row)
        c.execute("""
            INSERT INTO archive
                (company_name, position_title, location, job_url, date_posted,
                 date_applied, source, status, status_date, follow_up_date,
                 salary_range, notes, documents, job_description,
                 date_archived, created_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            job["company_name"], job["position_title"], job["location"],
            job["job_url"],      job["date_posted"],    job["date_applied"],
            job["source"],       job["status"],         job["status_date"],
            job["follow_up_date"], job["salary_range"], job["notes"],
            job["documents"],    job["job_description"],
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            job.get("created_at", ""),
        ))
        c.execute("DELETE FROM jobs WHERE id=?", (job_id,))
        conn.commit()
        conn.close()
        return True

    def restore_from_archive(self, archive_row):
        """Re-insert an archived job as a fresh Applied record."""
        data = dict(archive_row)
        data["status"]       = "Applied"
        data["status_date"]  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data["date_applied"] = ""
        return self.insert_job(data)

    # ---- ARCHIVE ----

    def get_all_archive(self):
        conn = self._conn()
        c = conn.cursor()
        c.execute("SELECT * FROM archive ORDER BY date_archived DESC")
        rows = [dict(r) for r in c.fetchall()]
        conn.close()
        return rows

    # ---- LOOKUP ----

    def find_active(self, company, position):
        conn = self._conn()
        c = conn.cursor()
        c.execute("SELECT id FROM jobs WHERE company_name=? AND position_title=?",
                  (company, position))
        row = c.fetchone()
        conn.close()
        return row["id"] if row else None

    def find_archived(self, company, position):
        conn = self._conn()
        c = conn.cursor()
        c.execute("""SELECT * FROM archive
                     WHERE company_name=? AND position_title=?
                     ORDER BY date_archived DESC LIMIT 1""",
                  (company, position))
        row = c.fetchone()
        conn.close()
        return dict(row) if row else None

    def path_exists(self, path):
        conn = self._conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM jobs WHERE documents=?",    (path,))
        a = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM archive WHERE documents=?", (path,))
        b = c.fetchone()[0]
        conn.close()
        return (a + b) > 0

    # ---- STATS ----

    def get_stats(self):
        conn = self._conn()
        c = conn.cursor()
        s = {}
        c.execute("SELECT COUNT(*) FROM jobs");    s["total_active"]   = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM archive"); s["total_archived"] = c.fetchone()[0]
        mo = date.today().strftime("%Y-%m")
        c.execute("SELECT COUNT(*) FROM jobs WHERE strftime('%Y-%m', date_applied)=?", (mo,))
        s["this_month"] = c.fetchone()[0]
        for st in STATUS_OPTIONS:
            c.execute("SELECT COUNT(*) FROM jobs WHERE status=?", (st,))
            s[st] = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM jobs WHERE date_applied IS NULL OR date_applied=''")
        s["no_date"] = c.fetchone()[0]
        ago30 = (date.today() - timedelta(days=30)).strftime("%Y-%m-%d")
        c.execute("SELECT COUNT(*) FROM jobs WHERE date_applied<=? AND status='Applied' AND date_applied!=''",
                  (ago30,))
        s["stale"] = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM jobs WHERE follow_up_date<=? AND follow_up_date!=''",
                  (date.today().strftime("%Y-%m-%d"),))
        s["overdue"] = c.fetchone()[0]
        c.execute("SELECT company_name, position_title, status, date_applied FROM jobs ORDER BY id DESC LIMIT 5")
        s["recent"] = [dict(r) for r in c.fetchall()]
        conn.close()
        return s

    # ---- FILTER ----

    def get_filtered(self, status=None, company=None):
        conn = self._conn()
        c = conn.cursor()
        q = "SELECT id, company_name, position_title, status, date_applied FROM jobs WHERE 1=1"
        p = []
        if status and status != "All":
            q += " AND status=?"; p.append(status)
        if company and company != "All":
            q += " AND company_name=?"; p.append(company)
        q += " ORDER BY date_applied DESC"
        c.execute(q, p)
        rows = [dict(r) for r in c.fetchall()]
        conn.close()
        return rows

    def get_companies(self):
        conn = self._conn()
        c = conn.cursor()
        c.execute("SELECT DISTINCT company_name FROM jobs ORDER BY company_name")
        rows = [r[0] for r in c.fetchall()]
        conn.close()
        return rows

    # ---- SCAN ----

    def scan_folders(self, base_path):
        added, skipped, errors = 0, 0, []
        if not os.path.isdir(base_path):
            return 0, 0, [f"Folder not found: {base_path}"]
        for company in os.listdir(base_path):
            cp = os.path.join(base_path, company)
            if not os.path.isdir(cp):
                continue
            for position in os.listdir(cp):
                pp = os.path.join(cp, position)
                if not os.path.isdir(pp):
                    continue
                if self.path_exists(pp):
                    skipped += 1
                else:
                    try:
                        self.insert_job({
                            "company_name":   company,
                            "position_title": position,
                            "status":         "Applied",
                            "notes":          "Imported from folder structure",
                            "documents":      pp,
                        })
                        added += 1
                    except Exception as e:
                        errors.append(f"{company}/{position}: {e}")
        return added, skipped, errors

    # ---- EXPORT ----

    def export_csv(self, filepath):
        conn = self._conn()
        c = conn.cursor()
        c.execute("SELECT * FROM jobs ORDER BY date_applied DESC")
        rows = c.fetchall()
        conn.close()
        if not rows:
            return False
        with open(filepath, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=rows[0].keys())
            w.writeheader()
            w.writerows([dict(r) for r in rows])
        return True


# ============================================================
# HELPERS
# ============================================================

def open_path(path):
    """Open a folder or URL in the OS default handler."""
    if not path:
        return
    if path.startswith("http"):
        webbrowser.open(path)
    elif os.path.isdir(path):
        if sys.platform == "win32":
            if HAS_WIN_UTILS:
                wu.shell_open(path)      # uses ShellExecuteW (Win32 C API)
            else:
                os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    else:
        messagebox.showerror("Not Found", f"Path not found:\n{path}")


def sanitize_filename(name):
    for ch in r'/\:*?"<>|':
        name = name.replace(ch, "-")
    return name.strip()


def make_word_doc(folder_path, company, position):
    """Create a job description .docx in the folder. Returns (ok, message)."""
    if not HAS_DOCX:
        return False, "python-docx is not installed.\nRun: pip install python-docx"
    filename  = sanitize_filename(position) + " - Job Description.docx"
    filepath  = os.path.join(folder_path, filename)
    if os.path.exists(filepath):
        open_path(filepath)
        return True, f"Opened existing document:\n{filename}"
    doc = Document()
    h = doc.add_heading(f"{company} — {position}", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_heading("Job Description", level=2)
    doc.add_paragraph("Paste the job description here...")
    doc.save(filepath)
    open_path(filepath)
    return True, f"Created and opened:\n{filename}"


# ============================================================
# MAIN TRACKER WINDOW
# ============================================================

class TrackerWindow(tk.Toplevel):
    """Single-record form for browsing and editing job applications."""

    def __init__(self, master, db, start_id=None):
        super().__init__(master)
        self.db        = db
        self.jobs      = []      # list of dicts
        self.index     = 0       # current position
        self._dirty    = False   # unsaved changes flag
        self._old_status = ""

        self.title(APP_TITLE)
        self.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.resizable(True, True)
        self.configure(bg="#F0F0F0")

        self._build_ui()
        self.reload(start_id=start_id)

        # Windows: dark title bar if system is in dark mode
        if HAS_WIN_UTILS and wu.is_windows_11():
            self.after(100, lambda: wu.apply_dark_titlebar(
                self.winfo_id(), wu.is_system_dark_mode()))

    # ---- BUILD UI ----

    def _build_ui(self):
        # ---- TOOLBAR ----
        tb = tk.Frame(self, bg="#2B579A", pady=6)
        tb.pack(fill="x")
        btn_style = dict(bg="#2B579A", fg="white", relief="flat",
                         font=FONT_BODY, padx=12, pady=4, cursor="hand2",
                         activebackground="#1E3F7A", activeforeground="white")
        tk.Button(tb, text="⟳ Scan Folders",   command=self._scan,       **btn_style).pack(side="left", padx=4)
        tk.Button(tb, text="↻ Refresh",         command=self._refresh,    **btn_style).pack(side="left", padx=4)
        tk.Button(tb, text="＋ New Job",         command=self._new_job,    **btn_style).pack(side="left", padx=4)
        tk.Button(tb, text="⬛ Dashboard",       command=self._dashboard,  **btn_style).pack(side="left", padx=4)
        tk.Button(tb, text="🗃 Archive",         command=self._open_archive,**btn_style).pack(side="left", padx=4)
        tk.Button(tb, text="📤 Export CSV",      command=self._export_csv, **btn_style).pack(side="right", padx=4)

        # ---- NAV BAR ----
        nav = tk.Frame(self, bg="#E0E7F3", pady=5)
        nav.pack(fill="x")
        tk.Button(nav, text="◀ Prev", command=self._prev,
                  font=FONT_BODY, relief="groove").pack(side="left", padx=8)
        self.lbl_nav = tk.Label(nav, text="", font=FONT_BODY, bg="#E0E7F3")
        self.lbl_nav.pack(side="left")
        tk.Button(nav, text="Next ▶", command=self._next,
                  font=FONT_BODY, relief="groove").pack(side="left", padx=8)
        tk.Button(nav, text="💾 Save", command=self._save,
                  font=FONT_BODY, relief="groove", bg="#D9EAD3").pack(side="right", padx=8)

        # ---- FORM FRAME (scrollable) ----
        canvas   = tk.Canvas(self, bg="#F0F0F0", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.form = tk.Frame(canvas, bg="#F0F0F0", padx=16, pady=12)
        self.form.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.form, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # ---- FIELD VARIABLES ----
        self.v = {
            "company_name":    tk.StringVar(),
            "position_title":  tk.StringVar(),
            "location":        tk.StringVar(),
            "date_posted":     tk.StringVar(),
            "date_applied":    tk.StringVar(),
            "source":          tk.StringVar(),
            "status":          tk.StringVar(value="Applied"),
            "status_date":     tk.StringVar(),
            "follow_up_date":  tk.StringVar(),
            "salary_range":    tk.StringVar(),
            "job_url":         tk.StringVar(),
            "documents":       tk.StringVar(),
        }

        f = self.form
        pad = dict(padx=6, pady=4)

        def lbl(text, row, col, **kw):
            tk.Label(f, text=text, font=FONT_LABEL, bg="#F0F0F0",
                     anchor="w", fg="#444").grid(row=row, column=col,
                     sticky="w", **pad, **kw)

        def entry(var, row, col, width=30, **kw):
            e = ttk.Entry(f, textvariable=var, width=width, font=FONT_BODY)
            e.grid(row=row, column=col, sticky="ew", **pad, **kw)
            return e

        # Row 0: ID / Company
        lbl("App ID:",        0, 0)
        self.lbl_id = tk.Label(f, text="—", font=FONT_BODY, bg="#F0F0F0", fg="#666")
        self.lbl_id.grid(row=0, column=1, sticky="w", **pad)

        lbl("Company Name:",  0, 2)
        entry(self.v["company_name"], 0, 3, width=36)

        # Row 1: Position / Location
        lbl("Position Title:", 1, 0)
        entry(self.v["position_title"], 1, 1, width=30, columnspan=1)
        lbl("Location:",       1, 2)
        entry(self.v["location"], 1, 3, width=36)

        # Row 2: Date Posted / Date Applied / Source
        lbl("Date Posted:",   2, 0)
        entry(self.v["date_posted"],  2, 1, width=14)
        lbl("Date Applied:",  2, 2)
        entry(self.v["date_applied"], 2, 3, width=14)

        # Row 3: Source / Salary
        lbl("Source:",        3, 0)
        entry(self.v["source"],       3, 1, width=20)
        lbl("Salary Range:",  3, 2)
        entry(self.v["salary_range"], 3, 3, width=24)

        # Row 4: Status / Status Date / Follow Up
        lbl("Status:",        4, 0)
        self.cmb_status = ttk.Combobox(
            f, textvariable=self.v["status"],
            values=STATUS_OPTIONS, state="readonly", width=24, font=FONT_BODY)
        self.cmb_status.grid(row=4, column=1, sticky="ew", **pad)
        self.cmb_status.bind("<<ComboboxSelected>>", self._on_status_change)
        self.lbl_status_color = tk.Label(f, text="●", font=("Segoe UI", 14),
                                          bg="#F0F0F0")
        self.lbl_status_color.grid(row=4, column=1, sticky="e", padx=2)

        lbl("Status Date:",   4, 2)
        entry(self.v["status_date"],   4, 3, width=20)

        # Row 5: Follow Up Date
        lbl("Follow-Up Date:", 5, 0)
        entry(self.v["follow_up_date"], 5, 1, width=14)

        # ---- NOTES ----
        lbl("Notes:", 6, 0)
        self.txt_notes = tk.Text(f, height=5, width=80, font=FONT_BODY,
                                  wrap="word", relief="solid", bd=1)
        self.txt_notes.grid(row=7, column=0, columnspan=4, sticky="ew",
                            padx=6, pady=2)

        # ---- JOB DESCRIPTION ----
        lbl("Job Description:", 8, 0)
        self.txt_desc = tk.Text(f, height=6, width=80, font=FONT_BODY,
                                 wrap="word", relief="solid", bd=1)
        self.txt_desc.grid(row=9, column=0, columnspan=4, sticky="ew",
                           padx=6, pady=2)

        # ---- ACTION BUTTONS ----
        btn_frame = tk.Frame(f, bg="#F0F0F0")
        btn_frame.grid(row=10, column=0, columnspan=4, pady=10, sticky="ew")

        action_style = dict(font=FONT_BODY, relief="groove",
                            padx=10, pady=6, cursor="hand2")

        self.btn_url = tk.Button(btn_frame, text="Add Job URL",
                                  command=self._job_url, bg="#CFE2FF",
                                  **action_style)
        self.btn_url.pack(side="left", padx=4)

        self.btn_docs = tk.Button(btn_frame, text="Open Folder",
                                   command=self._open_docs, bg="#D1ECE1",
                                   **action_style)
        self.btn_docs.pack(side="left", padx=4)

        tk.Button(btn_frame, text="Job Description Doc",
                  command=self._job_doc, bg="#FFF3CD",
                  **action_style).pack(side="left", padx=4)

        tk.Button(btn_frame, text="Archive This Job",
                  command=self._archive, bg="#FADADD",
                  **action_style).pack(side="right", padx=4)

        f.columnconfigure(1, weight=1)
        f.columnconfigure(3, weight=2)

        # dirty tracking
        for var in self.v.values():
            var.trace_add("write", self._mark_dirty)

    # ---- DATA ----

    def reload(self, start_id=None):
        self.jobs = self.db.get_all_jobs()
        if not self.jobs:
            self._clear_form()
            self.lbl_nav.config(text="No records")
            return
        if start_id:
            ids = [j["id"] for j in self.jobs]
            if start_id in ids:
                self.index = ids.index(start_id)
        self.index = max(0, min(self.index, len(self.jobs)-1))
        self._load_record()

    def _load_record(self):
        if not self.jobs:
            return
        self._dirty = False
        job = self.jobs[self.index]
        self.current_id = job["id"]
        self._old_status = job.get("status", "Applied")

        self.lbl_id.config(text=str(job["id"]))
        self.lbl_nav.config(text=f"Record {self.index+1} of {len(self.jobs)}")

        # Set variables (suppress dirty flag during load)
        self._loading = True
        for key, var in self.v.items():
            var.set(job.get(key) or "")

        self.txt_notes.delete("1.0", "end")
        self.txt_notes.insert("1.0", job.get("notes") or "")
        self.txt_desc.delete("1.0", "end")
        self.txt_desc.insert("1.0", job.get("job_description") or "")
        self._loading = False
        self._dirty   = False

        self._update_status_color()
        self._update_url_btn()
        self._update_docs_btn()

    def _clear_form(self):
        self.current_id = None
        self.lbl_id.config(text="—")
        for var in self.v.values():
            var.set("")
        self.txt_notes.delete("1.0", "end")
        self.txt_desc.delete("1.0", "end")

    def _collect(self):
        return {
            "company_name":    self.v["company_name"].get(),
            "position_title":  self.v["position_title"].get(),
            "location":        self.v["location"].get(),
            "job_url":         self.v["job_url"].get(),
            "date_posted":     self.v["date_posted"].get(),
            "date_applied":    self.v["date_applied"].get(),
            "source":          self.v["source"].get(),
            "status":          self.v["status"].get(),
            "status_date":     self.v["status_date"].get(),
            "follow_up_date":  self.v["follow_up_date"].get(),
            "salary_range":    self.v["salary_range"].get(),
            "notes":           self.txt_notes.get("1.0", "end-1c"),
            "documents":       self.v["documents"].get(),
            "job_description": self.txt_desc.get("1.0", "end-1c"),
        }

    def _mark_dirty(self, *_):
        if not getattr(self, "_loading", False):
            self._dirty = True

    # ---- NAV ----

    def _prev(self):
        if self._dirty and not self._confirm_discard():
            return
        if self.index > 0:
            self.index -= 1
            self._load_record()

    def _next(self):
        if self._dirty and not self._confirm_discard():
            return
        if self.index < len(self.jobs)-1:
            self.index += 1
            self._load_record()

    def _confirm_discard(self):
        return messagebox.askyesno("Unsaved Changes",
            "You have unsaved changes.\nDiscard and continue?")

    # ---- ACTIONS ----

    def _save(self):
        if not self.current_id:
            return
        data = self._collect()
        self.db.update_job(self.current_id, data)
        self._dirty = False
        # Refresh local list entry
        for j in self.jobs:
            if j["id"] == self.current_id:
                j.update(data)
                break
        messagebox.showinfo("Saved", "Record saved successfully.")

    def _refresh(self):
        if self._dirty and not self._confirm_discard():
            return
        self.reload(start_id=getattr(self, "current_id", None))
        messagebox.showinfo("Refreshed", "Records reloaded.")

    def _scan(self):
        folder = DEFAULT_JOBS_FOLDER
        if not os.path.isdir(folder):
            folder = filedialog.askdirectory(
                title="Select your jobapplications folder",
                initialdir=os.path.expanduser("~"))
            if not folder:
                return
        added, skipped, errors = self.db.scan_folders(folder)
        msg = f"{added} new job(s) added.\n{skipped} already existed and were skipped."
        if errors:
            msg += "\n\nErrors:\n" + "\n".join(errors[:5])
        messagebox.showinfo("Scan Complete", msg)
        # Windows toast notification
        if HAS_WIN_UTILS and added > 0:
            wu.show_notification("Job Tracker — Scan Complete",
                                 f"{added} new application(s) added.")
        self.reload(start_id=getattr(self, "current_id", None))

    def _archive(self):
        if not self.current_id:
            return
        company  = self.v["company_name"].get()
        position = self.v["position_title"].get()
        if not messagebox.askyesno("Archive",
            f"Archive this application?\n\n{company} — {position}\n\n"
            f"It will be removed from your active tracker."):
            return
        self.db.archive_job(self.current_id)
        if HAS_WIN_UTILS:
            wu.play_sound(wu.MB_ICONINFO)
        messagebox.showinfo("Archived", f"{company} — {position} archived.")
        self.reload()

    def _on_status_change(self, event=None):
        new_status = self.v["status"].get()
        if self._old_status and self._old_status != new_status:
            if not messagebox.askyesno("Confirm Status Change",
                f"Change status?\n\nFrom: {self._old_status}\nTo:   {new_status}"):
                self.v["status"].set(self._old_status)
                return
        self._old_status = new_status
        self.v["status_date"].set(datetime.now().strftime("%Y-%m-%d %H:%M"))
        self._update_status_color()

    def _update_status_color(self):
        color = STATUS_COLORS.get(self.v["status"].get(), "#000000")
        self.lbl_status_color.config(fg=color)

    def _job_url(self):
        url = self.v["job_url"].get().strip()
        if url:
            choice = messagebox.askyesnocancel(
                "Job URL", f"URL saved:\n{url}\n\nYes = Open  /  No = Edit  /  Cancel = Close")
            if choice is True:
                webbrowser.open(url)
            elif choice is False:
                new = simpledialog.askstring("Edit URL", "Edit the URL:", initialvalue=url)
                if new:
                    self.v["job_url"].set(new)
        else:
            new = simpledialog.askstring("Add Job URL", "Paste the job posting URL:")
            if new:
                self.v["job_url"].set(new)
        self._update_url_btn()

    def _update_url_btn(self):
        has = bool(self.v["job_url"].get().strip())
        self.btn_url.config(text="Open Job Posting" if has else "Add Job URL")

    def _open_docs(self):
        path = self.v["documents"].get().strip()
        if path:
            open_path(path)
        else:
            folder = filedialog.askdirectory(title="Link a folder to this application")
            if folder:
                self.v["documents"].set(folder)
                self._update_docs_btn()

    def _update_docs_btn(self):
        has  = bool(self.v["documents"].get().strip())
        self.btn_docs.config(text="Open Application Folder" if has else "Link Folder")

    def _job_doc(self):
        folder = self.v["documents"].get().strip()
        if not folder:
            messagebox.showwarning("No Folder", "No folder linked to this record.")
            return
        ok, msg = make_word_doc(folder,
                                 self.v["company_name"].get(),
                                 self.v["position_title"].get())
        if not ok:
            messagebox.showerror("Error", msg)
        else:
            messagebox.showinfo("Job Description", msg)

    def _new_job(self):
        NewJobDialog(self, self.db, callback=self.reload)

    def _dashboard(self):
        DashboardWindow(self.master, self.db,
                        open_tracker_cb=lambda jid: self.reload(start_id=jid))

    def _open_archive(self):
        ArchiveWindow(self.master, self.db)

    def _export_csv(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV", "*.csv")],
            title="Export active jobs to CSV")
        if path:
            ok = self.db.export_csv(path)
            if ok:
                messagebox.showinfo("Exported", f"Saved to:\n{path}")
            else:
                messagebox.showwarning("Empty", "No active jobs to export.")


# ============================================================
# NEW JOB DIALOG
# ============================================================

class NewJobDialog(tk.Toplevel):
    """Dialog for creating a new job application record (and folder)."""

    def __init__(self, master, db, callback=None):
        super().__init__(master)
        self.db       = db
        self.callback = callback
        self.title("New Job Application")
        self.geometry("420x260")
        self.resizable(False, False)
        self.grab_set()
        self._build()

    def _build(self):
        pad = dict(padx=12, pady=6)
        tk.Label(self, text="New Job Application", font=FONT_HEAD).pack(**pad)

        frm = tk.Frame(self); frm.pack(fill="x", padx=16)
        self.v_company  = tk.StringVar()
        self.v_position = tk.StringVar()
        self.v_folder   = tk.StringVar(value=DEFAULT_JOBS_FOLDER)

        for row, (lbl, var) in enumerate([
            ("Company Name:",   self.v_company),
            ("Position Title:", self.v_position),
            ("Base Folder:",    self.v_folder),
        ]):
            tk.Label(frm, text=lbl, anchor="w", font=FONT_LABEL).grid(
                row=row, column=0, sticky="w", padx=4, pady=4)
            ttk.Entry(frm, textvariable=var, width=34, font=FONT_BODY).grid(
                row=row, column=1, padx=4, pady=4)

        frm.columnconfigure(1, weight=1)

        btn_f = tk.Frame(self); btn_f.pack(pady=10)
        tk.Button(btn_f, text="Create",  command=self._create,
                  font=FONT_BODY, bg="#D9EAD3", padx=12).pack(side="left", padx=6)
        tk.Button(btn_f, text="Cancel",  command=self.destroy,
                  font=FONT_BODY, padx=12).pack(side="left", padx=6)

    def _create(self):
        company  = self.v_company.get().strip()
        position = self.v_position.get().strip()
        base     = self.v_folder.get().strip()

        if not company or not position:
            messagebox.showwarning("Missing Info", "Company and position are required.")
            return

        # Check active
        existing_id = self.db.find_active(company, position)
        if existing_id:
            messagebox.showinfo("Already Exists",
                f"{company} — {position} is already in your active tracker.")
            if self.callback:
                self.callback(start_id=existing_id)
            self.destroy()
            return

        # Check archive
        archived = self.db.find_archived(company, position)
        if archived:
            if messagebox.askyesno("Found in Archive",
                f"{company} — {position} is in your archive.\n\n"
                f"Restore as a new application (job reposted)?"):
                new_id = self.db.restore_from_archive(archived)
                messagebox.showinfo("Restored", "Job restored from archive. Update Date Applied.")
                if self.callback:
                    self.callback(start_id=new_id)
            self.destroy()
            return

        # Build folder path
        company_path  = os.path.join(base, company)
        position_path = os.path.join(company_path, position)
        create_folder = False
        if not os.path.isdir(position_path):
            if messagebox.askyesno("Create Folders",
                f"Create these folders?\n\n{position_path}"):
                create_folder = True
            else:
                self.destroy()
                return

        if create_folder:
            os.makedirs(position_path, exist_ok=True)

        new_id = self.db.insert_job({
            "company_name":   company,
            "position_title": position,
            "status":         "Applied",
            "documents":      position_path if os.path.isdir(position_path) else "",
        })

        messagebox.showinfo("Created", f"New application created:\n{company} — {position}")
        if create_folder and os.path.isdir(position_path):
            open_path(position_path)
        if self.callback:
            self.callback(start_id=new_id)
        self.destroy()


# ============================================================
# DASHBOARD WINDOW
# ============================================================

class DashboardWindow(tk.Toplevel):
    """Summary stats, flags, filter, and quick-open of any record."""

    def __init__(self, master, db, open_tracker_cb=None):
        super().__init__(master)
        self.db              = db
        self.open_tracker_cb = open_tracker_cb
        self.title("Dashboard — " + APP_TITLE)
        self.geometry("860x700")
        self.configure(bg="#F0F0F0")
        self._build()
        self._refresh()

    # ---- BUILD ----

    def _build(self):
        # Toolbar
        tb = tk.Frame(self, bg="#2B579A", pady=6)
        tb.pack(fill="x")
        btn = dict(bg="#2B579A", fg="white", relief="flat", font=FONT_BODY,
                   padx=10, pady=4, cursor="hand2",
                   activebackground="#1E3F7A", activeforeground="white")
        tk.Button(tb, text="↻ Refresh",    command=self._refresh,  **btn).pack(side="left", padx=4)
        tk.Button(tb, text="⟳ Scan",       command=self._scan,     **btn).pack(side="left", padx=4)
        tk.Button(tb, text="＋ New Job",    command=self._new_job,  **btn).pack(side="left", padx=4)
        tk.Button(tb, text="📂 Tracker",   command=self._tracker,  **btn).pack(side="left", padx=4)
        tk.Button(tb, text="🗃 Archive",   command=self._archive,  **btn).pack(side="left", padx=4)
        if HAS_WIN_UTILS and sys.platform == "win32":
            tk.Button(tb, text="⚙ Windows Settings",
                      command=self._win_settings, **btn).pack(side="right", padx=4)

        # Main split: left = stats, right = filter
        body = tk.Frame(self, bg="#F0F0F0")
        body.pack(fill="both", expand=True, padx=12, pady=8)

        left  = tk.Frame(body, bg="#F0F0F0"); left.pack(side="left", fill="y", padx=4)
        right = tk.Frame(body, bg="#F0F0F0"); right.pack(side="left", fill="both", expand=True, padx=4)

        # ---- LEFT: stats ----
        self._stat_tiles = {}

        def section(parent, title):
            tk.Label(parent, text=title, font=FONT_HEAD, bg="#F0F0F0",
                     fg="#2B579A").pack(anchor="w", pady=(8,2))

        def stat_row(parent, label, key, color="#222"):
            f = tk.Frame(parent, bg="#FFFFFF", relief="groove", bd=1)
            f.pack(fill="x", pady=2, ipadx=6, ipady=3)
            tk.Label(f, text=label, font=FONT_LABEL, bg="#FFFFFF",
                     anchor="w", width=24).pack(side="left")
            lbl = tk.Label(f, text="—", font=("Segoe UI", 11, "bold"),
                           bg="#FFFFFF", fg=color, width=6)
            lbl.pack(side="right")
            self._stat_tiles[key] = lbl

        section(left, "📊 Overview")
        stat_row(left, "Total Active",    "total_active")
        stat_row(left, "Total Archived",  "total_archived", "#555")
        stat_row(left, "Added This Month","this_month",     "#0055A5")

        section(left, "📋 By Status")
        for st in STATUS_OPTIONS:
            stat_row(left, st, st, STATUS_COLORS.get(st, "#222"))

        section(left, "⚠ Flags")
        stat_row(left, "Missing Date Applied", "no_date",  "#A30000")
        stat_row(left, "Overdue Follow-Ups",   "overdue",  "#A30000")
        stat_row(left, "Stale 30+ Days",       "stale",    "#C46000")

        # ---- RIGHT: filter + list ----
        section(right, "🔍 Filter Jobs")
        ff = tk.Frame(right, bg="#F0F0F0"); ff.pack(fill="x", pady=4)

        tk.Label(ff, text="Status:", font=FONT_LABEL, bg="#F0F0F0").grid(
            row=0, column=0, sticky="w", padx=4)
        self.v_filter_status = tk.StringVar(value="All")
        cmb_s = ttk.Combobox(ff, textvariable=self.v_filter_status,
                              values=["All"] + STATUS_OPTIONS,
                              state="readonly", width=22, font=FONT_BODY)
        cmb_s.grid(row=0, column=1, padx=4)

        tk.Label(ff, text="Company:", font=FONT_LABEL, bg="#F0F0F0").grid(
            row=0, column=2, sticky="w", padx=4)
        self.v_filter_company = tk.StringVar(value="All")
        self.cmb_company = ttk.Combobox(ff, textvariable=self.v_filter_company,
                                         state="readonly", width=22, font=FONT_BODY)
        self.cmb_company.grid(row=0, column=3, padx=4)

        tk.Button(ff, text="Apply Filter", command=self._apply_filter,
                  font=FONT_BODY, relief="groove").grid(row=0, column=4, padx=6)
        tk.Button(ff, text="Clear",        command=self._clear_filter,
                  font=FONT_BODY, relief="groove").grid(row=0, column=5, padx=2)

        tk.Label(right, text="Double-click a row to open in Tracker",
                 font=FONT_SMALL, bg="#F0F0F0", fg="#666").pack(anchor="w", pady=(4,0))

        cols = ("Company", "Position", "Status", "Date Applied")
        self.tree = ttk.Treeview(right, columns=cols, show="headings", height=14)
        for col, w in zip(cols, (160, 200, 160, 100)):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<Double-1>", self._on_double_click)

        sb = ttk.Scrollbar(right, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        # Recent section
        section(right, "🕐 Recent Applications")
        self.txt_recent = tk.Text(right, height=5, font=FONT_SMALL,
                                   state="disabled", relief="solid", bd=1)
        self.txt_recent.pack(fill="x", pady=4)

    # ---- LOGIC ----

    def _refresh(self):
        stats = self.db.get_stats()
        for key, lbl in self._stat_tiles.items():
            val = stats.get(key, 0)
            lbl.config(text=str(val))
            # Flag coloring
            if key in ("no_date", "overdue", "stale") and val > 0:
                lbl.config(fg="#A30000")
            elif key in ("no_date", "overdue", "stale"):
                lbl.config(fg="#007200")

        # Update company combobox
        companies = ["All"] + self.db.get_companies()
        self.cmb_company.config(values=companies)

        # Recent text
        self.txt_recent.config(state="normal")
        self.txt_recent.delete("1.0", "end")
        for r in stats.get("recent", []):
            line = (f"{r['company_name']} — {r['position_title']}  "
                    f"[{r['status']}]  {r.get('date_applied') or 'no date'}\n")
            self.txt_recent.insert("end", line)
        self.txt_recent.config(state="disabled")

        self._apply_filter()

    def _apply_filter(self):
        self.tree.delete(*self.tree.get_children())
        rows = self.db.get_filtered(
            status=self.v_filter_status.get(),
            company=self.v_filter_company.get())
        for r in rows:
            color = STATUS_COLORS.get(r["status"], "#000")
            self.tree.insert("", "end",
                iid=str(r["id"]),
                values=(r["company_name"], r["position_title"],
                        r["status"], r.get("date_applied") or "—"),
                tags=(r["status"],))
        for st, col in STATUS_COLORS.items():
            self.tree.tag_configure(st, foreground=col)

    def _clear_filter(self):
        self.v_filter_status.set("All")
        self.v_filter_company.set("All")
        self._apply_filter()

    def _on_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        job_id = int(sel[0])
        if self.open_tracker_cb:
            self.open_tracker_cb(job_id)

    def _scan(self):
        folder = DEFAULT_JOBS_FOLDER
        if not os.path.isdir(folder):
            folder = filedialog.askdirectory(title="Select jobapplications folder")
            if not folder:
                return
        added, skipped, errors = self.db.scan_folders(folder)
        messagebox.showinfo("Scan", f"{added} added, {skipped} skipped.")
        self._refresh()

    def _new_job(self):
        NewJobDialog(self, self.db, callback=lambda **_: self._refresh())

    def _tracker(self):
        TrackerWindow(self.master, self.db)

    def _archive(self):
        ArchiveWindow(self.master, self.db)

    def _win_settings(self):
        """Windows-specific settings dialog."""
        if not HAS_WIN_UTILS:
            return
        win = tk.Toplevel(self)
        win.title("Windows Settings")
        win.geometry("360x220")
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text="Windows Integration Settings",
                 font=FONT_HEAD).pack(pady=10)

        # Startup toggle
        startup_var = tk.BooleanVar(value=wu.is_in_startup())
        frm = tk.Frame(win); frm.pack(fill="x", padx=20, pady=6)
        tk.Checkbutton(frm, text="Start with Windows (adds to registry Run key)",
                       variable=startup_var, font=FONT_BODY).pack(anchor="w")

        # Dark title bar toggle
        dark_var = tk.BooleanVar(value=wu.is_system_dark_mode())
        tk.Checkbutton(frm, text="Dark title bar (Windows 11)",
                       variable=dark_var, font=FONT_BODY).pack(anchor="w", pady=4)

        # Version info
        ver = wu.get_windows_version()
        tk.Label(win, text=f"Windows {ver[0]}.{ver[1]} build {ver[2]}",
                 font=FONT_SMALL, fg="#666").pack()

        def apply():
            exe = sys.executable if not getattr(sys, "frozen", False) \
                  else sys.executable
            if startup_var.get():
                wu.add_to_startup(exe)
            else:
                wu.remove_from_startup()
            if wu.is_windows_11():
                wu.apply_dark_titlebar(self.winfo_id(), dark_var.get())
            messagebox.showinfo("Settings", "Settings applied.", parent=win)
            win.destroy()

        tk.Button(win, text="Apply", command=apply,
                  font=FONT_BODY, bg="#D9EAD3", padx=16).pack(pady=10)


# ============================================================
# ARCHIVE WINDOW
# ============================================================

class ArchiveWindow(tk.Toplevel):
    """Read-only view of archived job applications."""

    def __init__(self, master, db):
        super().__init__(master)
        self.db = db
        self.title("Archive — " + APP_TITLE)
        self.geometry("860x520")
        self._build()
        self._load()

    def _build(self):
        tb = tk.Frame(self, bg="#555", pady=6)
        tb.pack(fill="x")
        tk.Button(tb, text="↻ Refresh", command=self._load,
                  bg="#555", fg="white", relief="flat", font=FONT_BODY,
                  padx=10, pady=4).pack(side="left", padx=4)

        tk.Label(self, text="Archived Applications (read-only)",
                 font=FONT_HEAD, bg="#F0F0F0").pack(pady=6)

        cols = ("ID", "Company", "Position", "Status", "Date Applied", "Date Archived")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        for col, w in zip(cols, (50, 160, 200, 130, 100, 130)):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=8, pady=4)

        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

    def _load(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.db.get_all_archive():
            self.tree.insert("", "end", values=(
                r["id"], r["company_name"], r["position_title"],
                r["status"],
                r.get("date_applied") or "—",
                (r.get("date_archived") or "—")[:10],
            ))


# ============================================================
# APPLICATION ENTRY POINT
# ============================================================

class App(tk.Tk):
    def __init__(self):
        # ---- Windows: DPI awareness BEFORE creating any window ----
        if HAS_WIN_UTILS:
            wu.set_dpi_aware()

        super().__init__()
        self.withdraw()   # hide root window; we use Toplevels
        self.title(APP_TITLE)
        db = Database(DB_FILE)

        # Open dashboard first
        dash = DashboardWindow(
            self, db,
            open_tracker_cb=lambda jid: TrackerWindow(self, db, start_id=jid))

        # Also open tracker
        tracker = TrackerWindow(self, db)

        # If either window closes and the other is still open, keep app alive
        def on_close_tracker():
            tracker.destroy()
            if not dash.winfo_exists():
                self.quit()
        def on_close_dash():
            dash.destroy()
            if not tracker.winfo_exists():
                self.quit()

        tracker.protocol("WM_DELETE_WINDOW", on_close_tracker)
        dash.protocol("WM_DELETE_WINDOW", on_close_dash)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
