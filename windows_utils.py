"""
windows_utils.py
================
Windows-specific enhancements for Job Application Tracker.
Uses ctypes to call Win32 C APIs directly — no extra install needed.

Imported by job_tracker.py only when running on Windows.
All functions degrade gracefully if called on Mac/Linux.
"""

import sys
import os

IS_WINDOWS = sys.platform == "win32"

# ============================================================
# LAZY IMPORTS — only load when on Windows
# ============================================================
if IS_WINDOWS:
    import ctypes
    import ctypes.wintypes
    from ctypes import windll, wintypes
    import winreg
    import subprocess


# ============================================================
# 1. DPI AWARENESS
#    Makes text and UI sharp on 4K / HiDPI monitors.
#    Without this, Windows scales the app and it looks blurry.
#    Calls: SetProcessDpiAwareness (shcore.dll) — Win8.1+
#           SetProcessDPIAware     (user32.dll)  — fallback
# ============================================================

def set_dpi_aware():
    """Call before creating the Tk root window."""
    if not IS_WINDOWS:
        return
    try:
        # Per-monitor DPI awareness (best)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            # System DPI awareness (fallback for older Windows)
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ============================================================
# 2. WINDOWS TITLE BAR — DARK MODE
#    On Windows 11 (build 22000+) we can tell the OS to give
#    the title bar a dark background to match a dark theme.
#    Calls: DwmSetWindowAttribute (dwmapi.dll)
# ============================================================

DWMWA_USE_IMMERSIVE_DARK_MODE = 20   # Windows 11 attribute ID

def apply_dark_titlebar(hwnd: int, dark: bool = True):
    """
    hwnd: window handle from root.winfo_id()
    dark: True = dark title bar, False = light
    """
    if not IS_WINDOWS:
        return
    try:
        value = ctypes.c_int(1 if dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
    except Exception:
        pass


def is_system_dark_mode() -> bool:
    """Check if the user has Windows dark mode enabled (registry)."""
    if not IS_WINDOWS:
        return False
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        )
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0   # 0 = dark mode
    except Exception:
        return False


# ============================================================
# 3. WINDOWS TOAST NOTIFICATIONS
#    Shows a pop-up notification in the system tray area.
#    Works without needing the app window to be in focus.
#
#    Uses: Shell_NotifyIcon (shell32.dll) via a small helper
#    or falls back to a simpler balloon-tip approach via
#    win10toast if installed, then ctypes balloon tip.
# ============================================================

def _balloon_tip_ctypes(title: str, message: str, icon_path: str = ""):
    """
    Show a Windows balloon tooltip notification using ctypes.
    Works on Windows 7+ without any extra libraries.
    """
    if not IS_WINDOWS:
        return

    # Constants
    NIM_ADD        = 0x00000000
    NIM_MODIFY     = 0x00000001
    NIM_DELETE     = 0x00000002
    NIF_MESSAGE    = 0x00000001
    NIF_ICON       = 0x00000002
    NIF_TIP        = 0x00000004
    NIF_INFO       = 0x00000010
    NIIF_INFO      = 0x00000001
    NIIF_WARNING   = 0x00000002
    WM_USER        = 0x0400
    IDI_INFORMATION = 0x7F

    class NOTIFYICONDATA(ctypes.Structure):
        _fields_ = [
            ("cbSize",           ctypes.c_uint32),
            ("hWnd",             ctypes.c_void_p),
            ("uID",              ctypes.c_uint32),
            ("uFlags",           ctypes.c_uint32),
            ("uCallbackMessage", ctypes.c_uint32),
            ("hIcon",            ctypes.c_void_p),
            ("szTip",            ctypes.c_wchar * 128),
            ("dwState",          ctypes.c_uint32),
            ("dwStateMask",      ctypes.c_uint32),
            ("szInfo",           ctypes.c_wchar * 256),
            ("uTimeout",         ctypes.c_uint32),
            ("szInfoTitle",      ctypes.c_wchar * 64),
            ("dwInfoFlags",      ctypes.c_uint32),
        ]

    shell32 = ctypes.windll.shell32
    user32  = ctypes.windll.user32

    # Get a hidden window handle
    hwnd = user32.GetDesktopWindow()

    # Load info icon
    hicon = user32.LoadIconW(None, ctypes.c_wchar_p(IDI_INFORMATION))

    nid               = NOTIFYICONDATA()
    nid.cbSize        = ctypes.sizeof(NOTIFYICONDATA)
    nid.hWnd          = hwnd
    nid.uID           = 1
    nid.uFlags        = NIF_ICON | NIF_TIP | NIF_INFO | NIF_MESSAGE
    nid.uCallbackMessage = WM_USER + 1
    nid.hIcon         = hicon
    nid.szTip         = "Job Tracker"
    nid.szInfo        = message[:255]
    nid.szInfoTitle   = title[:63]
    nid.dwInfoFlags   = NIIF_INFO
    nid.uTimeout      = 5000

    shell32.Shell_NotifyIconW(NIM_ADD,    ctypes.byref(nid))
    shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(nid))

    # Clean up after 6 seconds
    import threading
    def _remove():
        import time; time.sleep(6)
        try:
            shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
        except Exception:
            pass
    threading.Thread(target=_remove, daemon=True).start()


def show_notification(title: str, message: str):
    """
    Show a Windows toast/balloon notification.
    Tries win10toast first (pip install win10toast), then ctypes fallback.
    """
    if not IS_WINDOWS:
        return
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(title, message, duration=6, threaded=True)
    except ImportError:
        _balloon_tip_ctypes(title, message)
    except Exception:
        pass


# ============================================================
# 4. FLASH TASKBAR BUTTON
#    Flashes the app's taskbar button to get the user's
#    attention without stealing focus.
#    Calls: FlashWindowEx (user32.dll)
# ============================================================

def flash_taskbar(hwnd: int, times: int = 3):
    """Flash the taskbar button to alert the user."""
    if not IS_WINDOWS:
        return
    try:
        class FLASHWINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize",    ctypes.c_uint),
                ("hwnd",      ctypes.c_void_p),
                ("dwFlags",   ctypes.c_uint),
                ("uCount",    ctypes.c_uint),
                ("dwTimeout", ctypes.c_uint),
            ]
        FLASHW_ALL   = 0x00000003  # flash both caption and taskbar
        FLASHW_TIMER = 0x00000004  # keep flashing until StopFlash
        info = FLASHWINFO(
            cbSize=ctypes.sizeof(FLASHWINFO),
            hwnd=hwnd,
            dwFlags=FLASHW_ALL,
            uCount=times,
            dwTimeout=0,
        )
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
    except Exception:
        pass


# ============================================================
# 5. RELIABLE PATH DETECTION
#    Uses SHGetKnownFolderPath (shell32) to find Desktop and
#    OneDrive paths correctly even with folder redirection.
# ============================================================

# GUIDs for known folders
FOLDERID_Desktop    = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
FOLDERID_Documents  = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
FOLDERID_UserProfile= "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"

def get_known_folder(folder_guid: str) -> str:
    """
    Calls SHGetKnownFolderPath to get a Windows special folder path.
    More reliable than os.path.expanduser on machines with folder redirection.
    """
    if not IS_WINDOWS:
        return ""
    try:
        import uuid
        from ctypes import POINTER, c_wchar_p

        CoTaskMemFree = ctypes.windll.ole32.CoTaskMemFree
        SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
        SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(ctypes.c_byte * 16),
            ctypes.c_uint32,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_wchar_p),
        ]
        SHGetKnownFolderPath.restype = ctypes.c_long

        guid_bytes = (ctypes.c_byte * 16)(*uuid.UUID(folder_guid).bytes_le)
        path_ptr   = ctypes.c_wchar_p()
        hr = SHGetKnownFolderPath(
            ctypes.byref(guid_bytes), 0, None, ctypes.byref(path_ptr))
        if hr == 0:
            path = path_ptr.value
            CoTaskMemFree(path_ptr)
            return path or ""
    except Exception:
        pass
    return ""


def get_desktop_path() -> str:
    path = get_known_folder(FOLDERID_Desktop)
    return path or os.path.join(os.path.expanduser("~"), "Desktop")


def find_jobs_folder() -> str:
    """
    Find the jobapplications folder.
    Checks OneDrive Desktop, then regular Desktop.
    Uses Win32 API for accurate path — handles OneDrive folder redirection.
    """
    desktop = get_desktop_path()

    # OneDrive Desktop (common on work/school machines)
    onedrive_env = os.environ.get("OneDriveConsumer") or os.environ.get("OneDrive", "")
    if onedrive_env:
        od_path = os.path.join(onedrive_env, "Desktop", "jobapplications")
        if os.path.isdir(od_path):
            return od_path

    # Regular desktop
    regular = os.path.join(desktop, "jobapplications")
    if os.path.isdir(regular):
        return regular

    # Fallback
    return regular


# ============================================================
# 6. WINDOWS SOUND — play system sounds
#    Calls: MessageBeep (user32.dll)
# ============================================================

MB_OK          = 0x00000000   # default beep
MB_ICONWARNING = 0x00000030   # warning sound
MB_ICONERROR   = 0x00000010   # error sound
MB_ICONINFO    = 0x00000040   # info / asterisk sound

def play_sound(sound_type: int = MB_ICONINFO):
    """Play a Windows system sound."""
    if not IS_WINDOWS:
        return
    try:
        ctypes.windll.user32.MessageBeep(sound_type)
    except Exception:
        pass


# ============================================================
# 7. OPEN FILE / FOLDER WITH ShellExecute
#    More powerful than os.startfile — supports "select" mode
#    which opens Explorer with the file highlighted.
# ============================================================

def open_and_select_file(filepath: str):
    """Open Explorer and highlight a specific file."""
    if not IS_WINDOWS:
        return
    try:
        subprocess.Popen(["explorer", "/select,", os.path.normpath(filepath)])
    except Exception:
        try:
            os.startfile(os.path.dirname(filepath))
        except Exception:
            pass


def shell_open(path: str):
    """Open a file or folder using the Windows shell (handles all types)."""
    if not IS_WINDOWS:
        return
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "open", path, None, None, 1)
    except Exception:
        try:
            os.startfile(path)
        except Exception:
            pass


# ============================================================
# 8. WINDOWS VERSION CHECK
# ============================================================

def get_windows_version() -> tuple:
    """Returns (major, minor, build) e.g. (10, 0, 22000) for Win11."""
    if not IS_WINDOWS:
        return (0, 0, 0)
    try:
        ver = sys.getwindowsversion()
        return (ver.major, ver.minor, ver.build)
    except Exception:
        return (0, 0, 0)


def is_windows_11() -> bool:
    major, _, build = get_windows_version()
    return major >= 10 and build >= 22000


# ============================================================
# 9. STARTUP WITH WINDOWS (optional)
#    Adds/removes the app from HKCU Run registry key
#    so it starts when the user logs in.
# ============================================================

REGISTRY_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_NAME         = "JobApplicationTracker"

def add_to_startup(exe_path: str):
    """Register the app to start with Windows."""
    if not IS_WINDOWS:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_RUN_KEY, 0,
                             winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{exe_path}"')
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def remove_from_startup():
    """Remove the app from Windows startup."""
    if not IS_WINDOWS:
        return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_RUN_KEY, 0,
                             winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def is_in_startup() -> bool:
    if not IS_WINDOWS:
        return False
    try:
        key  = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_RUN_KEY)
        val, _ = winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return bool(val)
    except Exception:
        return False
