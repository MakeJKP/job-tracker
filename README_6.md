# Job Application Tracker 🗂️

A desktop app for tracking job applications during your job search.  
Built with Python + tkinter — **no Microsoft Access required**.  
Data stored locally in a SQLite database file (`~/job_tracker.db`).

---

## Features

| Feature | Description |
|---|---|
| **Main Tracker** | Browse and edit each application one at a time |
| **Status Dropdown** | Applied → Phone Screen → Interview → Final Round → Offer / Rejected / Withdrawn / Ghosted |
| **Status Colors** | Each status gets a color indicator |
| **Archive** | Move closed applications to a separate archive table |
| **Folder Scanner** | Reads your `jobapplications\{Company}\{Position}\` folder structure and imports new records |
| **New Job Creator** | Creates company + position folders and inserts a new record — restores from archive if the job is reposted |
| **URL Button** | Add, view, or edit the job posting URL |
| **Open Folder** | Open the linked application folder in Windows Explorer |
| **Job Description Doc** | Creates a `.docx` file in the application folder using `python-docx` |
| **Dashboard** | Summary stats by status, flags for stale/overdue records, filter + quick-open |
| **Export CSV** | Export all active jobs to a CSV file |

---

## Requirements

- **Python 3.8+**  
  Download: https://www.python.org/downloads/
- **tkinter** — included with Python on Windows (no install needed)
- **python-docx** *(optional)* — only needed for the "Job Description Doc" button

---

## Installation

```bash
# 1. Clone the repo
git clone https://github.com/yourusername/job-tracker.git
cd job-tracker

# 2. (Optional) Install python-docx for Word document support
pip install python-docx

# 3. Run the app
python job_tracker.py
```

No other packages required.

---

## Folder Structure for Scanning

The scanner expects this layout under your `jobapplications` folder:

```
Desktop\jobapplications\
    ├── Google\
    │       └── Software Engineer\          ← one folder per position
    │               ├── Resume.docx
    │               └── Cover Letter.docx
    └── Microsoft\
            └── IT Support Specialist\
```

The scanner checks **both** active jobs and the archive before adding a record,  
so it never creates duplicates.

**Default scan path:**  
`%USERPROFILE%\OneDrive\Desktop\jobapplications\`  
Falls back to `%USERPROFILE%\Desktop\jobapplications\` if OneDrive is not found.  
You can also pick any folder manually when you click **Scan**.

---

## Status Values & Colors

| Status | Color |
|---|---|
| Applied | 🟢 Green |
| Phone Screen Scheduled | 🔵 Blue |
| Interview Scheduled | 🔵 Blue |
| Final Round | 🟠 Orange |
| Offer Received | 🟩 Bright Green |
| Rejected | 🔴 Red |
| Withdrawn | 🔴 Red |
| Ghosted | ⚫ Gray |

---

## Database

The database is a single SQLite file stored at:

```
C:\Users\YourName\job_tracker.db
```

**Tables:**

| Table | Description |
|---|---|
| `jobs` | Active job applications |
| `archive` | Archived / closed applications |

You can open the database with [DB Browser for SQLite](https://sqlitebrowser.org/) to view or back up your data.

---

## Running on Mac or Linux

The app works on Mac and Linux with minor differences:

- Folder paths default to `~/Desktop/jobapplications/` (no OneDrive)
- "Open Folder" uses `open` (Mac) or `xdg-open` (Linux)
- tkinter must be installed separately on some Linux distros:  
  `sudo apt-get install python3-tk`

---

## License

MIT — free to use, modify, and share.
