"""
where-file - Lightweight Tray App
Zero external dependencies. Pure Python + Windows API.

Features:
  - System tray icon (always running)
  - Ctrl+Shift+F: Copy active window's file path instantly
  - Click tray icon: Show open files popup
"""

import ctypes
import ctypes.wintypes as w
import json
import os
import subprocess
import threading
import time
import tkinter as tk
from pathlib import Path

# ─── Windows API ─────────────────────────────────────────────────────

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
shell32 = ctypes.windll.shell32
gdi32 = ctypes.windll.gdi32

# Constants
WM_USER = 0x0400
WM_TRAYICON = WM_USER + 20
WM_COMMAND = 0x0111
WM_DESTROY = 0x0002
WM_HOTKEY = 0x0312
WM_LBUTTONUP = 0x0202
WM_RBUTTONUP = 0x0205

NIM_ADD = 0x00
NIM_MODIFY = 0x01
NIM_DELETE = 0x02
NIF_MESSAGE = 0x01
NIF_ICON = 0x02
NIF_TIP = 0x04
NIF_INFO = 0x10

HOTKEY_ID = 1
MOD_CTRL_SHIFT = 0x0002 | 0x0004
VK_F = 0x46

IDM_OPEN = 1001
IDM_QUIT = 1002

WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_long, w.HWND, w.UINT, w.WPARAM, w.LPARAM)


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", w.DWORD),
        ("hWnd", w.HWND),
        ("uID", w.UINT),
        ("uFlags", w.UINT),
        ("uCallbackMessage", w.UINT),
        ("hIcon", w.HICON),
        ("szTip", w.WCHAR * 128),
        ("dwState", w.DWORD),
        ("dwStateMask", w.DWORD),
        ("szInfo", w.WCHAR * 256),
        ("uVersion", w.UINT),
        ("szInfoTitle", w.WCHAR * 64),
        ("dwInfoFlags", w.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", w.HICON),
    ]


class WNDCLASSEXW(ctypes.Structure):
    _fields_ = [
        ("cbSize", w.UINT),
        ("style", w.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", w.HINSTANCE),
        ("hIcon", w.HICON),
        ("hCursor", w.HANDLE),
        ("hbrBackground", w.HBRUSH),
        ("lpszMenuName", w.LPCWSTR),
        ("lpszClassName", w.LPCWSTR),
        ("hIconSm", w.HICON),
    ]


# ─── Clipboard (instant, no subprocess) ─────────────────────────────

def copy_to_clipboard(text):
    if not text:
        return
    user32.OpenClipboard(0)
    user32.EmptyClipboard()
    data = text.encode("utf-16-le") + b"\x00\x00"
    hmem = kernel32.GlobalAlloc(0x0042, len(data))
    ptr = kernel32.GlobalLock(hmem)
    ctypes.memmove(ptr, data, len(data))
    kernel32.GlobalUnlock(hmem)
    user32.SetClipboardData(13, hmem)  # CF_UNICODETEXT
    user32.CloseClipboard()


# ─── File Scanner (PowerShell COM, cached) ───────────────────────────

_cache = {"files": [], "ts": 0, "active": "", "active_ts": 0}
CACHE_TTL = 3

PS_ALL_FILES = r"""
$r = @()
try { $a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
  foreach($p in $a.Presentations) { if($p.FullName) { $r += [PSCustomObject]@{app='PowerPoint';name=$p.Name;path=$p.FullName;ext='.pptx'} } }
} catch {}
try { $a = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
  foreach($d in $a.Documents) { if($d.FullName) { $r += [PSCustomObject]@{app='Word';name=$d.Name;path=$d.FullName;ext='.docx'} } }
} catch {}
try { $a = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  foreach($w in $a.Workbooks) { if($w.FullName) { $r += [PSCustomObject]@{app='Excel';name=$w.Name;path=$w.FullName;ext='.xlsx'} } }
} catch {}
$x = '\.(pdf|txt|csv|hwp|hwpx|psd|ai|svg|png|jpg|jpeg|gif|mp4|mp3|zip|dwg)["''\\s]?$'
Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -and $_.Name -notin @('POWERPNT.EXE','WINWORD.EXE','EXCEL.EXE') -and $_.CommandLine -match $x } | ForEach-Object {
  $ms = [regex]::Matches($_.CommandLine, '"([^"]+\.[a-zA-Z0-9]+)"')
  foreach($m in $ms) { $fp=$m.Groups[1].Value; if((Test-Path $fp -EA 0) -and !(Test-Path $fp -PathType Container -EA 0)) {
    $r += [PSCustomObject]@{app=$_.Name -replace '\.exe$','';name=[IO.Path]::GetFileName($fp);path=$fp;ext=[IO.Path]::GetExtension($fp).ToLower()} } }
}
if($r.Count -eq 0){'[]'} elseif($r.Count -eq 1){'['+($r|ConvertTo-Json -Compress)+']'} else{$r|ConvertTo-Json -Compress}
"""


def _run_ps(script):
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return r.stdout.strip()
    except Exception:
        return ""


def get_all_files(force=False):
    now = time.time()
    if not force and now - _cache["ts"] < CACHE_TTL:
        return _cache["files"]
    out = _run_ps(PS_ALL_FILES)
    try:
        files = json.loads(out) if out and out != "[]" else []
    except json.JSONDecodeError:
        files = []
    _cache["files"] = files
    _cache["ts"] = now
    return files


def get_active_file_path():
    """Get file path of the foreground window. Uses ctypes first, PowerShell only for Office."""
    hwnd = user32.GetForegroundWindow()
    pid = w.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

    # Get process name via ctypes (instant)
    PROCESS_QUERY_INFO = 0x0400
    PROCESS_VM_READ = 0x0010
    handle = kernel32.OpenProcess(PROCESS_QUERY_INFO | PROCESS_VM_READ, False, pid.value)
    proc_name = ""
    if handle:
        buf = (ctypes.c_wchar * 260)()
        psapi = ctypes.windll.psapi
        psapi.GetModuleBaseNameW(handle, None, buf, 260)
        kernel32.CloseHandle(handle)
        proc_name = buf.value.upper()

    # Office apps: use COM via PowerShell (only way to get accurate path)
    office_map = {
        "POWERPNT.EXE": "PowerPoint.Application", "POWERPNT": "PowerPoint.Application",
        "WINWORD.EXE": "Word.Application", "WINWORD": "Word.Application",
        "EXCEL.EXE": "Excel.Application", "EXCEL": "Excel.Application",
    }
    com_prog = office_map.get(proc_name)
    if com_prog:
        prop = {"PowerPoint.Application": "ActivePresentation",
                "Word.Application": "ActiveDocument",
                "Excel.Application": "ActiveWorkbook"}[com_prog]
        script = f"try {{ ([Runtime.InteropServices.Marshal]::GetActiveObject('{com_prog}')).{prop}.FullName }} catch {{}}"
        return _run_ps(script)

    # Non-Office: try window title parsing (instant, no subprocess)
    length = user32.GetWindowTextLengthW(hwnd) + 1
    title_buf = (ctypes.c_wchar * length)()
    user32.GetWindowTextW(hwnd, title_buf, length)
    title = title_buf.value

    # Many apps show "filename - AppName" or "filepath - AppName"
    if " - " in title:
        candidate = title.split(" - ")[0].strip()
        if os.path.isfile(candidate):
            return candidate

    return ""


# ─── Tray Icon (pure ctypes, no pystray/Pillow) ─────────────────────

class TrayApp:
    def __init__(self):
        self.hwnd = None
        self.nid = None
        self.popup = None
        self._wndproc_ref = None  # prevent GC

    def _create_icon(self):
        """Load folder icon from shell32.dll (index 4)."""
        return shell32.ExtractIconW(kernel32.GetModuleHandleW(None), "shell32.dll", 4)

    def _wndproc(self, hwnd, msg, wparam, lparam):
        if msg == WM_TRAYICON:
            if lparam == WM_LBUTTONUP:
                self._show_popup()
            elif lparam == WM_RBUTTONUP:
                self._show_menu()
            return 0
        elif msg == WM_COMMAND:
            cmd_id = wparam & 0xFFFF
            if cmd_id == IDM_OPEN:
                self._show_popup()
            elif cmd_id == IDM_QUIT:
                self._cleanup()
                user32.PostQuitMessage(0)
            return 0
        elif msg == WM_HOTKEY and wparam == HOTKEY_ID:
            self._on_hotkey()
            return 0
        elif msg == WM_DESTROY:
            self._cleanup()
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    def _show_menu(self):
        menu = user32.CreatePopupMenu()
        user32.AppendMenuW(menu, 0x0000, IDM_OPEN, "Open File List")
        user32.AppendMenuW(menu, 0x0800, 0, None)  # separator
        user32.AppendMenuW(menu, 0x0000, IDM_QUIT, "Quit")
        pt = w.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        user32.SetForegroundWindow(self.hwnd)
        user32.TrackPopupMenu(menu, 0, pt.x, pt.y, 0, self.hwnd, None)
        user32.DestroyMenu(menu)

    def _show_balloon(self, title, msg):
        """Show balloon notification (instant, no subprocess)."""
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.hwnd
        nid.uID = 1
        nid.uFlags = NIF_INFO
        nid.szInfoTitle = title[:63]
        nid.szInfo = msg[:255]
        nid.dwInfoFlags = 0x01  # NIIF_INFO
        shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(nid))

    def _on_hotkey(self):
        """Ctrl+Shift+F: copy active file path."""
        def _do():
            path = get_active_file_path()
            if path:
                copy_to_clipboard(path)
                name = Path(path).name
                self._show_balloon("Path Copied", name)
            else:
                self._show_balloon("where-file", "No file path found.")
        threading.Thread(target=_do, daemon=True).start()

    def _show_popup(self):
        threading.Thread(target=self._popup_thread, daemon=True).start()

    def _popup_thread(self):
        FilePopup(get_all_files(force=True)).show()

    def _cleanup(self):
        if self.nid:
            shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self.nid))

    def run(self):
        hinstance = kernel32.GetModuleHandleW(None)
        class_name = "OpenFilesViewerClass"

        self._wndproc_ref = WNDPROC(self._wndproc)

        wc = WNDCLASSEXW()
        wc.cbSize = ctypes.sizeof(WNDCLASSEXW)
        wc.lpfnWndProc = self._wndproc_ref
        wc.hInstance = hinstance
        wc.lpszClassName = class_name
        user32.RegisterClassExW(ctypes.byref(wc))

        self.hwnd = user32.CreateWindowExW(
            0, class_name, "OpenFilesViewer", 0,
            0, 0, 0, 0, None, None, hinstance, None,
        )

        # Add tray icon
        icon = self._create_icon()
        self.nid = NOTIFYICONDATAW()
        self.nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        self.nid.hWnd = self.hwnd
        self.nid.uID = 1
        self.nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        self.nid.uCallbackMessage = WM_TRAYICON
        self.nid.hIcon = icon
        self.nid.szTip = "where-file (Ctrl+Shift+F)"
        shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(self.nid))

        # Register hotkey: Ctrl+Shift+F
        user32.RegisterHotKey(self.hwnd, HOTKEY_ID, MOD_CTRL_SHIFT, VK_F)

        print("where-file running.")
        print("  Ctrl+Shift+F : Copy active file path")
        print("  Tray icon    : Click to see all files")
        print("  Tray right-click > Quit to exit")

        # Message loop
        msg = w.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0):
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))


# ─── Popup Window (tkinter, built-in) ────────────────────────────────

class FilePopup:
    BG = "#0f172a"
    CARD = "#1e293b"
    BORDER = "#334155"
    TXT = "#f1f5f9"
    SUB = "#94a3b8"
    GREEN = "#16a34a"
    COLORS = {"PowerPoint": "#dc2626", "Word": "#2563eb", "Excel": "#16a34a"}

    def __init__(self, files):
        self.files = files

    def show(self):
        self.win = tk.Tk()
        self.win.title("Open Files")
        self.win.configure(bg=self.BG)
        self.win.geometry("680x460")
        self.win.attributes("-topmost", True)

        # Header
        hdr = tk.Frame(self.win, bg=self.BG, pady=10, padx=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"Open Files ({len(self.files)})",
                 font=("Segoe UI", 14, "bold"), bg=self.BG, fg=self.TXT).pack(side="left")
        tk.Button(hdr, text="Refresh", font=("Segoe UI", 9),
                  bg="#334155", fg=self.TXT, relief="flat", padx=10, pady=3,
                  cursor="hand2", command=self._refresh,
                  activebackground="#475569", activeforeground=self.TXT).pack(side="right")

        # Scrollable list
        container = tk.Frame(self.win, bg=self.BG)
        container.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.canvas = tk.Canvas(container, bg=self.BG, highlightthickness=0)
        sb = tk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.frame = tk.Frame(self.canvas, bg=self.BG)
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw")
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>",
                             lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        self._render()
        self.win.mainloop()

    def _refresh(self):
        self.files = get_all_files(force=True)
        self._render()

    def _render(self):
        for w_child in self.frame.winfo_children():
            w_child.destroy()

        if not self.files:
            tk.Label(self.frame, text="No open files detected",
                     font=("Segoe UI", 12), bg=self.BG, fg=self.SUB).pack(pady=40)
            return

        for f in self.files:
            card = tk.Frame(self.frame, bg=self.CARD, padx=12, pady=8,
                            highlightbackground=self.BORDER, highlightthickness=1)
            card.pack(fill="x", pady=2)

            color = self.COLORS.get(f.get("app", ""), "#475569")
            tk.Frame(card, bg=color, width=4).pack(side="left", fill="y", padx=(0, 10))

            info = tk.Frame(card, bg=self.CARD)
            info.pack(side="left", fill="both", expand=True)
            tk.Label(info, text=f.get("name", ""), font=("Segoe UI", 11, "bold"),
                     bg=self.CARD, fg=self.TXT, anchor="w").pack(fill="x")
            tk.Label(info, text=f.get("path", ""), font=("Segoe UI", 8),
                     bg=self.CARD, fg=self.SUB, anchor="w", wraplength=380).pack(fill="x")

            btns = tk.Frame(card, bg=self.CARD)
            btns.pack(side="right", padx=(6, 0))

            path = f.get("path", "")
            folder = str(Path(path).parent) if path else ""

            for label, val in [("Path", path), ("Folder", folder)]:
                b = tk.Button(btns, text=f"Copy {label}", font=("Segoe UI", 8),
                              bg="#334155", fg=self.TXT, relief="flat", padx=6, pady=2,
                              cursor="hand2", activebackground="#475569", activeforeground=self.TXT)
                b.configure(command=lambda v=val, btn=b, lbl=label: self._copy(v, btn, lbl))
                b.pack(pady=1)

    def _copy(self, text, btn, label):
        copy_to_clipboard(text)
        btn.configure(text="Copied!", bg=self.GREEN)
        self.win.after(1000, lambda: btn.configure(text=f"Copy {label}", bg="#334155")
                       if btn.winfo_exists() else None)


# ─── Entry Point ─────────────────────────────────────────────────────

if __name__ == "__main__":
    TrayApp().run()
