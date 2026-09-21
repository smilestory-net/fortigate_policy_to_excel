#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FortiGate Policy Exporter - Modern Dark Cyber-Glassmorphism Webview UI
=====================================================================
High-fidelity dark cyber-glassmorphism GUI matching user requirements:
- Authentic Fortinet Logo embedded from fortinet.ico
- Windows 11 style titlebar with functional Minimize, Maximize/Restore, Close
- Dragging restricted strictly to the top title bar
- Identical equal-width buttons for Select File & Select Folder (125px)
- Integrated cyber-blue progress bar (animated in real-time)
- Clean OUTPUT console header (removed unused dummy tabs)
- 8-directional smooth border resize handles without unwanted OS caption bar
- Explicit Taskbar AppUserModelID and Fortinet icon binding
- Powered by pywebview and Microsoft Edge WebView2.
"""

import sys
import os
import re
import json
import time
import tempfile
import threading
import subprocess
import webbrowser
import struct
import ctypes
from ctypes import wintypes

try:
    import webview
    HAS_WEBVIEW = True
except ImportError:
    HAS_WEBVIEW = False

user32 = ctypes.windll.user32 if sys.platform == 'win32' else None
kernel32 = ctypes.windll.kernel32 if sys.platform == 'win32' else None
dwmapi = ctypes.windll.dwmapi if sys.platform == 'win32' else None

if sys.platform == 'win32':
    kernel32.VirtualAlloc.restype = ctypes.c_void_p
    kernel32.VirtualAlloc.argtypes = [ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD, wintypes.DWORD]
    user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
    user32.SetWindowLongPtrW.restype = ctypes.c_void_p
    user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetWindowLongPtrW.restype = ctypes.c_void_p
    user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.CallWindowProcW.restype = ctypes.c_longlong

WM_NCLBUTTONDOWN = 0x00A1
WM_NCLBUTTONDBLCLK = 0x00A3
HTCAPTION = 2

HTLEFT = 10
HTRIGHT = 11
HTTOP = 12
HTTOPLEFT = 13
HTTOPRIGHT = 14
HTBOTTOM = 15
HTBOTTOMLEFT = 16
HTBOTTOMRIGHT = 17

SW_MINIMIZE = 6
SW_MAXIMIZE = 3
SW_RESTORE = 9

GWL_STYLE = -16
GWL_WNDPROC = -4
WS_CAPTION = 0x00C00000
WS_MINIMIZEBOX = 0x00020000
WS_MAXIMIZEBOX = 0x00010000
WS_SYSMENU = 0x00080000
WS_THICKFRAME = 0x00040000

WM_NCCALCSIZE = 0x0083
SWP_FRAMECHANGED = 0x0020
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOZORDER = 0x0004

DWMWA_TRANSITIONS_FORCEDISABLED = 3
DWMWA_BORDER_COLOR = 34
DWMWA_COLOR_DEFAULT = 0xFFFFFFFF
DWMWA_COLOR_NONE = 0xFFFFFFFE



class RECT(ctypes.Structure):
    _fields_ = [
        ('left', wintypes.LONG),
        ('top', wintypes.LONG),
        ('right', wintypes.LONG),
        ('bottom', wintypes.LONG),
    ]


class NCCALCSIZE_PARAMS(ctypes.Structure):
    _fields_ = [
        ('rgrc', RECT * 3),
        ('lppos', ctypes.c_void_p),
    ]


class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('rcMonitor', RECT),
        ('rcWork', RECT),
        ('dwFlags', wintypes.DWORD),
    ]


class MARGINS(ctypes.Structure):
    _fields_ = [
        ('cxLeftWidth', ctypes.c_int),
        ('cxRightWidth', ctypes.c_int),
        ('cyTopHeight', ctypes.c_int),
        ('cyBottomHeight', ctypes.c_int),
    ]


class WINDOWPLACEMENT(ctypes.Structure):
    _fields_ = [
        ('length', wintypes.UINT),
        ('flags', wintypes.UINT),
        ('showCmd', wintypes.UINT),
        ('ptMinPosition', wintypes.POINT),
        ('ptMaxPosition', wintypes.POINT),
        ('rcNormalPosition', RECT),
    ]


# --- Pure x64 Machine Code Thunk for WM_NCCALCSIZE (Zero-GIL, Zero-Deadlock) ---
_allocated_thunk_memory = []


def _build_nccalcsize_thunk(old_proc, pCallWindowProcW):
    """
    Pure x64 machine code thunk that intercepts WM_NCCALCSIZE (0x83) and returns 0,
    removing the Windows native titlebar and borders completely, while allowing
    WS_CAPTION | WS_THICKFRAME to remain enabled so Windows DWM animations work!
    Runs at raw CPU speed in the native message loop with ZERO Python GIL involvement.
    """
    code = bytearray()
    code.extend(b'\x81\xfa\x83\x00\x00\x00')  # cmp edx, 0x83 (WM_NCCALCSIZE)
    code.extend(b'\x75\x03')                  # jne +3 (to call_orig)
    code.extend(b'\x31\xc0')                  # xor eax, eax (return 0)
    code.extend(b'\xc3')                      # ret
    # call_orig: CallWindowProcW(old_proc, hwnd, msg, wparam, lparam)
    code.extend(b'\x48\x83\xec\x28')          # sub rsp, 0x28 (shadow space + 16-byte alignment)
    code.extend(b'\x4c\x89\x4c\x24\x20')      # mov [rsp+0x20], r9 (5th arg: lparam)
    code.extend(b'\x4d\x89\xc1')              # mov r9, r8         (4th arg: wparam)
    code.extend(b'\x49\x89\xd0')              # mov r8, rdx        (3rd arg: msg)
    code.extend(b'\x48\x89\xca')              # mov rdx, rcx       (2nd arg: hwnd)
    code.extend(b'\x48\xb9' + struct.pack('<Q', old_proc))         # mov rcx, old_proc (1st arg)
    code.extend(b'\x48\xb8' + struct.pack('<Q', pCallWindowProcW)) # mov rax, pCallWindowProcW
    code.extend(b'\xff\xd0')                  # call rax
    code.extend(b'\x48\x83\xc4\x28')          # add rsp, 0x28
    code.extend(b'\xc3')                      # ret
    return bytes(code)




_accent_color_cached = None
_accent_color_ref_cached = None


def get_system_accent_color():
    """Retrieve Windows system accent color as hex string (#rrggbb)"""
    global _accent_color_cached
    if _accent_color_cached is not None:
        return _accent_color_cached
    if sys.platform == 'win32':
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\DWM')
            val, _ = winreg.QueryValueEx(key, 'AccentColor')
            # AccentColor is stored in ABGR (0xAABBGGRR)
            r = val & 0xFF
            g = (val >> 8) & 0xFF
            b = (val >> 16) & 0xFF
            _accent_color_cached = f"#{r:02x}{g:02x}{b:02x}"
            return _accent_color_cached
        except Exception:
            pass
    _accent_color_cached = "#0078d4"
    return _accent_color_cached


def get_system_accent_color_ref():
    """Retrieve Windows system accent color as COLORREF (0x00BBGGRR) for DwmSetWindowAttribute"""
    global _accent_color_ref_cached
    if _accent_color_ref_cached is not None:
        return _accent_color_ref_cached
    if sys.platform == 'win32':
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\DWM')
            val, _ = winreg.QueryValueEx(key, 'AccentColor')
            # AccentColor is 0xAABBGGRR. For COLORREF (0x00BBGGRR), strip alpha:
            _accent_color_ref_cached = val & 0x00FFFFFF
            return _accent_color_ref_cached
        except Exception:
            pass
    _accent_color_ref_cached = 0x00D47800  # Fallback to #0078d4 in COLORREF (0x00, 0xD4, 0x78, 0x00)
    return _accent_color_ref_cached


def _get_config_dir():
    r"""
    Retrieve user application data / config directory across Windows, macOS, and Linux.
    Keeps configuration completely isolated from the executable or workspace directory.
    - Windows : %LOCALAPPDATA%\FortiGatePolicyExporter
    - macOS   : ~/Library/Application Support/FortiGatePolicyExporter
    - Linux   : ~/.config/fortigate_policy_exporter
    """
    if sys.platform == 'win32':
        base = os.environ.get('LOCALAPPDATA')
        if not base:
            base = os.path.join(os.environ.get('USERPROFILE', ''), 'AppData', 'Local')
        if not os.path.isdir(base):
            base = tempfile.gettempdir()
        path = os.path.join(base, 'FortiGatePolicyExporter')
    elif sys.platform == 'darwin':
        path = os.path.expanduser('~/Library/Application Support/FortiGatePolicyExporter')
    else:
        xdg_config = os.environ.get('XDG_CONFIG_HOME', os.path.expanduser('~/.config'))
        path = os.path.join(xdg_config, 'fortigate_policy_exporter')

    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        path = tempfile.gettempdir()
    return path


THEME_CONFIG_FILE = os.path.join(_get_config_dir(), ".theme_config.json")

VALID_THEMES = {
    "antigravity-dark": {"is_dark": True, "bg_color": "#131314", "label": "Antigravity Dark"},
    "antigravity-light": {"is_dark": False, "bg_color": "#f8f9fa", "label": "Antigravity Light"},
    "cyber-dark": {"is_dark": True, "bg_color": "#0D121B", "label": "Cyber Dark"},
    "cyber-light": {"is_dark": False, "bg_color": "#eef2f8", "label": "Cyber Light"},
}


def load_app_config():
    """Retrieve saved application configuration (theme, auto_open), defaulting to 'antigravity-dark' and True"""
    cfg = {
        "theme": "antigravity-dark",
        "auto_open": True
    }
    try:
        # 1. Check user AppData/config directory
        if os.path.isfile(THEME_CONFIG_FILE):
            with open(THEME_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    if data.get("theme") in VALID_THEMES:
                        cfg["theme"] = data["theme"]
                    if "auto_open" in data:
                        cfg["auto_open"] = bool(data["auto_open"])
                    return cfg

        # 2. Check legacy file in workspace/exe directory, migrate to AppData and clean up
        legacy_dirs = [
            os.path.dirname(os.path.abspath(__file__)),
            os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else None
        ]
        for ldir in legacy_dirs:
            if not ldir:
                continue
            legacy_file = os.path.join(ldir, ".theme_config.json")
            if os.path.isfile(legacy_file):
                try:
                    with open(legacy_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if isinstance(data, dict):
                            if data.get("theme") in VALID_THEMES:
                                cfg["theme"] = data["theme"]
                            if "auto_open" in data:
                                cfg["auto_open"] = bool(data["auto_open"])
                            save_app_config(cfg)
                    os.remove(legacy_file)
                    return cfg
                except Exception:
                    pass
    except Exception:
        pass
    return cfg


def get_saved_theme():
    """Retrieve saved theme ID, defaulting to 'antigravity-dark' on initial run"""
    return load_app_config().get("theme", "antigravity-dark")


def get_saved_auto_open():
    """Retrieve saved auto_open toggle state, defaulting to True on initial run"""
    return load_app_config().get("auto_open", True)


def save_app_config(updates):
    """Persist settings to user AppData config file without overwriting existing keys"""
    try:
        cfg = load_app_config()
        cfg.update(updates)
        cfg_dir = _get_config_dir()
        os.makedirs(cfg_dir, exist_ok=True)
        with open(THEME_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print("save_app_config error:", e)


def save_theme_to_file(theme_id):
    """Persist theme ID to user AppData config file"""
    if theme_id in VALID_THEMES:
        save_app_config({"theme": theme_id})


def save_auto_open_to_file(auto_open):
    """Persist auto_open toggle state to user AppData config file"""
    save_app_config({"auto_open": bool(auto_open)})




def _get_resource_path(filename):
    """Resolve file path for normal python execution and PyInstaller onefile bundles."""
    if getattr(sys, 'frozen', False):
        if hasattr(sys, '_MEIPASS'):
            candidate = os.path.join(sys._MEIPASS, filename)
            if os.path.isfile(candidate):
                return candidate
        exe_candidate = os.path.join(os.path.dirname(sys.executable), filename)
        if os.path.isfile(exe_candidate):
            return exe_candidate
    local_candidate = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    if os.path.isfile(local_candidate):
        return local_candidate
    return filename


def _load_fortinet_icon_b64():
    """Extract embedded PNG from fortinet.ico to ensure 100% authentic Fortinet logo"""
    ico_path = _get_resource_path('fortinet.ico')
    b64_path = _get_resource_path('fortinet_logo.b64')
    
    if os.path.isfile(b64_path):
        try:
            with open(b64_path, 'r', encoding='ascii') as f:
                return f.read().strip()
        except Exception:
            pass

    if os.path.isfile(ico_path):
        try:
            import base64
            with open(ico_path, 'rb') as f:
                data = f.read()
            # In fortinet.ico, offset 22 is PNG data
            if len(data) > 22 and data[22:26] == b'\x89PNG':
                return base64.b64encode(data[22:]).decode('ascii')
            return base64.b64encode(data).decode('ascii')
        except Exception:
            pass
    return ""


FORTINET_LOGO_B64 = _load_fortinet_icon_b64()


HTML_TEMPLATE_RAW = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FortiGate Policy Exporter v2.1</title>
<style>
  :root {
    --win-accent: __WIN_ACCENT_COLOR__;

    /* Default Theme: Antigravity Dark */
    --app-bg: #131314;
    --app-text: #e3e3e3;
    --titlebar-border: #282a2d;
    --brand-version: #8ab4f8;
    --win-btn-color: #9aa0a6;
    --win-btn-hover-bg: #282a2d;
    --win-btn-hover-color: #e3e3e3;
    --card-bg: #1e1f20;
    --card-border: #333537;
    --card-title: #c4c7c5;
    --form-label: #9aa0a6;
    --input-bg: #131314;
    --input-border: #3c4043;
    --input-focus: #8ab4f8;
    --input-icon: #9aa0a6;
    --input-text: #e3e3e3;
    --input-placeholder: #5f6368;
    --btn-select-bg: #282a2d;
    --btn-select-border: #3c4043;
    --btn-select-text: #e3e3e3;
    --btn-select-hover-bg: #333539;
    --btn-select-hover-border: #5f6368;
    --btn-select-hover-text: #ffffff;
    --btn-select-shadow: none;
    --btn-select-hover-shadow: none;
    --slider-bg: #3c4043;
    --slider-checked: #8ab4f8;
    --slider-knob: #e3e3e3;
    --slider-knob-checked: #131314;
    --slider-shadow: none;
    --toggle-label: #c4c7c5;
    --btn-primary-bg: #1a73e8;
    --btn-primary-border: #8ab4f8;
    --btn-primary-text: #ffffff;
    --btn-primary-hover-bg: #1967d2;
    --btn-primary-hover-border: #aecbfa;
    --btn-primary-shadow: none;
    --btn-primary-hover-shadow: none;
    --btn-secondary-bg: #282a2d;
    --btn-secondary-border: #3c4043;
    --btn-secondary-text: #9aa0a6;
    --btn-secondary-hover-bg: #333539;
    --btn-secondary-hover-border: #5f6368;
    --btn-secondary-hover-text: #e3e3e3;
    --progress-wrap-bg: #282a2d;
    --progress-fill-bg: #8ab4f8;
    --progress-shadow: none;
    --term-bg: #18191c;
    --term-border: #282a2e;
    --term-header-bg: #1e1f20;
    --term-header-title: #e3e3e3;
    --term-dot-bg: #8ab4f8;
    --term-dot-shadow: none;
    --term-body-text: #c4c7c5;
    --term-scroll-track: #18191c;
    --term-scroll-thumb: #282a2d;
    --term-cursor: #9aa0a6;
    --status-bar-text: #9aa0a6;
    --status-dot-ready: #81c995;
    --status-dot-running: #fdd663;
    --status-dot-done: #81c995;
    --status-dot-error: #f28b82;
    --status-dot-shadow: none;
    --status-divider: #3c4043;
    --status-link: #9aa0a6;
    --status-link-hover: #e3e3e3;
    --theme-badge-bg: #282a2d;
    --theme-badge-border: #3c4043;
    --theme-badge-text: #8ab4f8;
    --theme-badge-hover-bg: #333539;
    --theme-dropdown-bg: #1e1f20;
    --theme-dropdown-border: #3c4043;
    --theme-dropdown-item-hover: #282a2d;
    --theme-dropdown-item-text: #e3e3e3;
    --log-vdom-color: #81c995;
    --log-vdom-shadow: none;
    --log-error-fg: #f28b82;
    --log-error-bg: rgba(242, 139, 130, 0.12);
    --log-error-border: #f28b82;
    --log-warning-fg: #fdd663;
    --log-warning-bg: rgba(253, 214, 99, 0.12);
    --log-warning-border: #fdd663;
    --log-debug-fg: #8ab4f8;
    --log-debug-bg: rgba(138, 180, 248, 0.12);
    --log-debug-border: #8ab4f8;
    --log-trace-fg: #d2a8ff;
  }

  /* Theme 2: Antigravity Light Mode */
  body[data-theme="antigravity-light"] {
    --app-bg: #f8f9fa;
    --app-text: #1f1f1f;
    --titlebar-border: #e0e2e6;
    --brand-version: #1a73e8;
    --win-btn-color: #5f6368;
    --win-btn-hover-bg: #e8eaed;
    --win-btn-hover-color: #202124;
    --card-bg: #ffffff;
    --card-border: #dadce0;
    --card-title: #444746;
    --form-label: #5f6368;
    --input-bg: #f1f3f4;
    --input-border: #dadce0;
    --input-focus: #1a73e8;
    --input-icon: #5f6368;
    --input-text: #1f1f1f;
    --input-placeholder: #80868b;
    --btn-select-bg: #edf2fc;
    --btn-select-border: #dadce0;
    --btn-select-text: #1a73e8;
    --btn-select-hover-bg: #e2ecfd;
    --btn-select-hover-border: #aecbfa;
    --btn-select-hover-text: #174ea6;
    --btn-select-shadow: none;
    --btn-select-hover-shadow: none;
    --slider-bg: #dadce0;
    --slider-checked: #1a73e8;
    --slider-knob: #ffffff;
    --slider-knob-checked: #ffffff;
    --slider-shadow: none;
    --toggle-label: #444746;
    --btn-primary-bg: #1a73e8;
    --btn-primary-border: #1a73e8;
    --btn-primary-text: #ffffff;
    --btn-primary-hover-bg: #174ea6;
    --btn-primary-hover-border: #174ea6;
    --btn-primary-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
    --btn-primary-hover-shadow: 0 2px 6px rgba(26, 115, 232, 0.25);
    --btn-secondary-bg: #f1f3f4;
    --btn-secondary-border: #dadce0;
    --btn-secondary-text: #5f6368;
    --btn-secondary-hover-bg: #e8eaed;
    --btn-secondary-hover-border: #bdc1c6;
    --btn-secondary-hover-text: #202124;
    --progress-wrap-bg: #e8eaed;
    --progress-fill-bg: #1a73e8;
    --progress-shadow: none;
    --term-bg: #ffffff;
    --term-border: #dadce0;
    --term-header-bg: #f1f3f4;
    --term-header-title: #202124;
    --term-dot-bg: #1a73e8;
    --term-dot-shadow: none;
    --term-body-text: #3c4043;
    --term-scroll-track: #f1f3f4;
    --term-scroll-thumb: #dadce0;
    --term-cursor: #5f6368;
    --status-bar-text: #5f6368;
    --status-dot-ready: #34a853;
    --status-dot-running: #f9ab00;
    --status-dot-done: #34a853;
    --status-dot-error: #ea4335;
    --status-dot-shadow: none;
    --status-divider: #dadce0;
    --status-link: #5f6368;
    --status-link-hover: #1a73e8;
    --theme-badge-bg: #edf2fc;
    --theme-badge-border: #dadce0;
    --theme-badge-text: #1a73e8;
    --theme-badge-hover-bg: #e2ecfd;
    --theme-dropdown-bg: #ffffff;
    --theme-dropdown-border: #dadce0;
    --theme-dropdown-item-hover: #f1f3f4;
    --theme-dropdown-item-text: #202124;
    --log-vdom-color: #1e8e3e;
    --log-vdom-shadow: none;
    --log-error-fg: #d93025;
    --log-error-bg: #fce8e6;
    --log-error-border: #ea4335;
    --log-warning-fg: #b06000;
    --log-warning-bg: #fef7e0;
    --log-warning-border: #f9ab00;
    --log-debug-fg: #1a73e8;
    --log-debug-bg: #e8f0fe;
    --log-debug-border: #1a73e8;
    --log-trace-fg: #7627bb;
  }

  /* Theme 3: Cyberpunk / Neon Dark Mode (Original Design) */
  body[data-theme="cyber-dark"] {
    --app-bg: #0D121B;
    --app-text: #f8fafc;
    --titlebar-border: rgba(255, 255, 255, 0.08);
    --brand-version: #38bdf8;
    --win-btn-color: #64748b;
    --win-btn-hover-bg: rgba(255, 255, 255, 0.08);
    --win-btn-hover-color: #f1f5f9;
    --card-bg: rgba(13, 19, 33, 0.75);
    --card-border: rgba(255, 255, 255, 0.1);
    --card-title: #94a3b8;
    --form-label: #94a3b8;
    --input-bg: rgba(9, 13, 19, 0.85);
    --input-border: rgba(255, 255, 255, 0.07);
    --input-focus: #2563eb;
    --input-icon: #64748b;
    --input-text: #e2e8f0;
    --input-placeholder: #475569;
    --btn-select-bg: #0e141f;
    --btn-select-border: #1e5bb8;
    --btn-select-text: #f8fafc;
    --btn-select-hover-bg: #142032;
    --btn-select-hover-border: #3b82f6;
    --btn-select-hover-text: #ffffff;
    --btn-select-shadow: 0 0 8px rgba(30, 91, 184, 0.55), inset 0 0 5px rgba(30, 91, 184, 0.25);
    --btn-select-hover-shadow: 0 0 14px rgba(59, 130, 246, 0.7), inset 0 0 8px rgba(59, 130, 246, 0.35);
    --slider-bg: #334155;
    --slider-checked: #2563eb;
    --slider-knob: #ffffff;
    --slider-knob-checked: #ffffff;
    --slider-shadow: 0 0 8px rgba(37, 99, 235, 0.5);
    --toggle-label: #cbd5e1;
    --btn-primary-bg: linear-gradient(180deg, #1d456f 0%, #122e4d 50%, #0b1e33 100%);
    --btn-primary-border: rgba(255, 255, 255, 0.22);
    --btn-primary-text: #ffffff;
    --btn-primary-hover-bg: linear-gradient(180deg, #245588 0%, #173b62 50%, #0e2742 100%);
    --btn-primary-hover-border: rgba(255, 255, 255, 0.35);
    --btn-primary-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.35), 0 6px 18px rgba(0, 0, 0, 0.45);
    --btn-primary-hover-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.45), 0 8px 24px rgba(11, 30, 51, 0.6);
    --btn-secondary-bg: rgba(255, 255, 255, 0.05);
    --btn-secondary-border: rgba(255, 255, 255, 0.12);
    --btn-secondary-text: #94a3b8;
    --btn-secondary-hover-bg: rgba(255, 255, 255, 0.09);
    --btn-secondary-hover-border: rgba(255, 255, 255, 0.2);
    --btn-secondary-hover-text: #f1f5f9;
    --progress-wrap-bg: rgba(255, 255, 255, 0.06);
    --progress-fill-bg: linear-gradient(90deg, #1d4ed8 0%, #3b82f6 60%, #60a5fa 100%);
    --progress-shadow: 0 0 10px rgba(59, 130, 246, 0.85);
    --term-bg: #080a0f;
    --term-border: rgba(255, 255, 255, 0.08);
    --term-header-bg: rgba(14, 19, 28, 0.95);
    --term-header-title: #f1f5f9;
    --term-dot-bg: #3b82f6;
    --term-dot-shadow: 0 0 8px #3b82f6;
    --term-body-text: #cbd5e1;
    --term-scroll-track: #080a0f;
    --term-scroll-thumb: #1e293b;
    --term-cursor: #94a3b8;
    --status-bar-text: #94a3b8;
    --status-dot-ready: #3b82f6;
    --status-dot-running: #f59e0b;
    --status-dot-done: #10b981;
    --status-dot-error: #ef4444;
    --status-dot-shadow: 0 0 10px currentColor;
    --status-divider: rgba(255, 255, 255, 0.15);
    --status-link: #94a3b8;
    --status-link-hover: #f1f5f9;
    --theme-badge-bg: rgba(30, 91, 184, 0.25);
    --theme-badge-border: #1e5bb8;
    --theme-badge-text: #38bdf8;
    --theme-badge-hover-bg: rgba(59, 130, 246, 0.35);
    --theme-dropdown-bg: #0e141f;
    --theme-dropdown-border: #1e5bb8;
    --theme-dropdown-item-hover: #142032;
    --theme-dropdown-item-text: #f8fafc;
    --log-vdom-color: #22c55e;
    --log-vdom-shadow: 0 0 8px rgba(34, 197, 94, 0.5);
    --log-error-fg: #ff4d6d;
    --log-error-bg: rgba(255, 77, 109, 0.16);
    --log-error-border: #ff4d6d;
    --log-warning-fg: #fbbf24;
    --log-warning-bg: rgba(251, 191, 36, 0.16);
    --log-warning-border: #fbbf24;
    --log-debug-fg: #38bdf8;
    --log-debug-bg: rgba(56, 189, 248, 0.16);
    --log-debug-border: #38bdf8;
    --log-trace-fg: #c084fc;
  }

  /* Theme 4: Cyberpunk / Sci-Fi Light Mode */
  body[data-theme="cyber-light"] {
    --app-bg: #eef2f8;
    --app-text: #0f172a;
    --titlebar-border: rgba(14, 165, 233, 0.25);
    --brand-version: #0284c7;
    --win-btn-color: #64748b;
    --win-btn-hover-bg: rgba(14, 165, 233, 0.12);
    --win-btn-hover-color: #0369a1;
    --card-bg: rgba(255, 255, 255, 0.9);
    --card-border: rgba(14, 165, 233, 0.3);
    --card-title: #0284c7;
    --form-label: #475569;
    --input-bg: #ffffff;
    --input-border: #cbd5e1;
    --input-focus: #0284c7;
    --input-icon: #0284c7;
    --input-text: #0f172a;
    --input-placeholder: #94a3b8;
    --btn-select-bg: #f0f9ff;
    --btn-select-border: #0284c7;
    --btn-select-text: #0284c7;
    --btn-select-hover-bg: #e0f2fe;
    --btn-select-hover-border: #0369a1;
    --btn-select-hover-text: #0369a1;
    --btn-select-shadow: 0 0 6px rgba(2, 132, 199, 0.3);
    --btn-select-hover-shadow: 0 0 10px rgba(2, 132, 199, 0.5);
    --slider-bg: #cbd5e1;
    --slider-checked: #0284c7;
    --slider-knob: #ffffff;
    --slider-knob-checked: #ffffff;
    --slider-shadow: 0 0 6px rgba(2, 132, 199, 0.4);
    --toggle-label: #334155;
    --btn-primary-bg: linear-gradient(180deg, #0284c7 0%, #0369a1 100%);
    --btn-primary-border: #0284c7;
    --btn-primary-text: #ffffff;
    --btn-primary-hover-bg: linear-gradient(180deg, #0369a1 0%, #075985 100%);
    --btn-primary-hover-border: #075985;
    --btn-primary-shadow: 0 2px 8px rgba(2, 132, 199, 0.35);
    --btn-primary-hover-shadow: 0 4px 12px rgba(2, 132, 199, 0.5);
    --btn-secondary-bg: #ffffff;
    --btn-secondary-border: #cbd5e1;
    --btn-secondary-text: #64748b;
    --btn-secondary-hover-bg: #f8fafc;
    --btn-secondary-hover-border: #94a3b8;
    --btn-secondary-hover-text: #0f172a;
    --progress-wrap-bg: #e2e8f0;
    --progress-fill-bg: linear-gradient(90deg, #0284c7 0%, #38bdf8 100%);
    --progress-shadow: 0 0 8px rgba(2, 132, 199, 0.5);
    --term-bg: #ffffff;
    --term-border: rgba(14, 165, 233, 0.25);
    --term-header-bg: #f0f9ff;
    --term-header-title: #0369a1;
    --term-dot-bg: #0284c7;
    --term-dot-shadow: 0 0 6px #0284c7;
    --term-body-text: #1e293b;
    --term-scroll-track: #f8fafc;
    --term-scroll-thumb: #cbd5e1;
    --term-cursor: #0284c7;
    --status-bar-text: #64748b;
    --status-dot-ready: #0284c7;
    --status-dot-running: #f59e0b;
    --status-dot-done: #10b981;
    --status-dot-error: #ef4444;
    --status-dot-shadow: 0 0 8px currentColor;
    --status-divider: #cbd5e1;
    --status-link: #64748b;
    --status-link-hover: #0284c7;
    --theme-badge-bg: #f0f9ff;
    --theme-badge-border: #0284c7;
    --theme-badge-text: #0284c7;
    --theme-badge-hover-bg: #e0f2fe;
    --theme-dropdown-bg: #ffffff;
    --theme-dropdown-border: #0284c7;
    --theme-dropdown-item-hover: #f0f9ff;
    --theme-dropdown-item-text: #0f172a;
    --log-vdom-color: #16a34a;
    --log-vdom-shadow: none;
    --log-error-fg: #dc2626;
    --log-error-bg: #fef2f2;
    --log-error-border: #ef4444;
    --log-warning-fg: #d97706;
    --log-warning-bg: #fffbeb;
    --log-warning-border: #f59e0b;
    --log-debug-fg: #0284c7;
    --log-debug-bg: #f0f9ff;
    --log-debug-border: #0284c7;
    --log-trace-fg: #7c3aed;
  }

  * {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    user-select: none;
    -webkit-user-select: none;
  }

  html, body {
    width: 100%;
    height: 100%;
    margin: 0;
    padding: 0;
    box-sizing: border-box;
    overflow: hidden;
    background: var(--app-bg);
    transition: background 0.25s ease;
  }

  /* 8-Directional Native Resize Border Handles */
  .resize-edge {
    position: absolute;
    z-index: 99999;
    background: transparent;
  }
  .resize-top { top: 0; left: 12px; right: 12px; height: 6px; cursor: ns-resize; }
  .resize-bottom { bottom: 0; left: 12px; right: 12px; height: 6px; cursor: ns-resize; }
  .resize-left { left: 0; top: 12px; bottom: 12px; width: 6px; cursor: ew-resize; }
  .resize-right { right: 0; top: 12px; bottom: 12px; width: 6px; cursor: ew-resize; }
  .resize-topleft { top: 0; left: 0; width: 12px; height: 12px; cursor: nwse-resize; }
  .resize-topright { top: 0; right: 0; width: 12px; height: 12px; cursor: nesw-resize; }
  .resize-bottomleft { bottom: 0; left: 0; width: 12px; height: 12px; cursor: nesw-resize; }
  .resize-bottomright { bottom: 0; right: 0; width: 12px; height: 12px; cursor: nwse-resize; }

  /* Root Window Frame */
  .window-root {
    width: 100%;
    height: 100%;
    margin: 0;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    padding: 8px 16px 10px 16px;
    background: var(--app-bg);
    border: none;
    box-shadow: none;
    border-radius: 0;
    position: relative;
    overflow: hidden;
    transition: background 0.25s ease;
  }

  /* Active Window Border - Handled natively by Windows 11 DWM */
  .window-root.active-window {
    border: none;
    box-shadow: none;
  }

  /* Maximized state: zero border radius, border completely removed */
  body.is-maximized {
    padding: 0 !important;
  }

  body.is-maximized .resize-edge {
    display: none !important;
  }

  .window-root.maximized,
  body.is-maximized .window-root {
    border-radius: 0 !important;
    border: none !important;
    box-shadow: none !important;
    padding: 10px 20px 12px 20px;
  }

  .window-root.animating-window {
    transition: all 0.28s cubic-bezier(0.16, 1, 0.3, 1) !important;
  }

  /* Header & Title Bar - Only Draggable Region */
  .title-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 2px 4px 8px 4px;
    cursor: default;
    border-bottom: 1px solid var(--titlebar-border);
    transition: border-color 0.25s ease;
  }

  .title-left {
    display: flex;
    align-items: center;
    gap: 10px;
    pointer-events: none;
  }

  /* Authentic Fortinet Logo */
  .brand-logo-wrap {
    display: flex;
    align-items: center;
    justify-content: center;
    pointer-events: auto;
  }

  .fortinet-brand-logo {
    width: 24px;
    height: 24px;
    border-radius: 4px;
    display: block;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.25);
  }

  .brand-title-wrap {
    display: flex;
    align-items: center;
    gap: 8px;
    pointer-events: auto;
  }

  .brand-title {
    font-size: 14px;
    font-weight: 700;
    letter-spacing: 0.8px;
    color: var(--app-text);
    text-transform: uppercase;
    transition: color 0.25s ease;
  }

  .brand-version {
    font-size: 11px;
    font-weight: 600;
    color: var(--brand-version);
    margin-left: 1px;
    transition: color 0.25s ease;
  }

  /* Header Right: Theme Badge & Windows Control Buttons */
  .title-right {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  /* Oval Theme Selector Badge (Left of Minimize) */
  .theme-badge-wrapper {
    position: relative;
    display: inline-block;
  }

  .theme-badge {
    height: 24px;
    padding: 0 11px;
    display: inline-flex;
    align-items: center;
    gap: 5px;
    border-radius: 9999px;
    background: var(--theme-badge-bg);
    border: 1px solid var(--theme-badge-border);
    color: var(--theme-badge-text);
    font-size: 11.5px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s ease;
    user-select: none;
    -webkit-user-select: none;
  }

  .theme-badge:hover {
    background: var(--theme-badge-hover-bg);
    transform: translateY(-0.5px);
  }

  .theme-badge svg.palette-icon {
    width: 12px;
    height: 12px;
    stroke: currentColor;
    stroke-width: 2;
    fill: none;
  }

  .theme-badge svg.chevron-icon {
    width: 10px;
    height: 10px;
    stroke: currentColor;
    stroke-width: 2;
    fill: none;
    transition: transform 0.2s ease;
  }

  .theme-badge-wrapper.open .theme-badge svg.chevron-icon {
    transform: rotate(180deg);
  }

  /* Theme Selection Dropdown Menu */
  .theme-dropdown {
    display: none;
    position: absolute;
    top: calc(100% + 6px);
    right: 0;
    min-width: 190px;
    background: var(--theme-dropdown-bg);
    border: 1px solid var(--theme-dropdown-border);
    border-radius: 10px;
    padding: 4px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
    z-index: 100000;
    backdrop-filter: blur(10px);
  }

  .theme-badge-wrapper.open .theme-dropdown {
    display: block;
    animation: fadeInDropdown 0.15s ease-out;
  }

  @keyframes fadeInDropdown {
    from { opacity: 0; transform: translateY(-4px); }
    to { opacity: 1; transform: translateY(0); }
  }

  .theme-item {
    display: flex;
    align-items: center;
    gap: 9px;
    padding: 7px 10px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 500;
    color: var(--theme-dropdown-item-text);
    cursor: pointer;
    transition: background 0.15s ease;
  }

  .theme-item:hover {
    background: var(--theme-dropdown-item-hover);
  }

  .theme-item.active {
    font-weight: 700;
    color: var(--brand-version);
  }

  .theme-preview-dot {
    width: 12px;
    height: 12px;
    border-radius: 50%;
    flex-shrink: 0;
    border: 1.5px solid rgba(255, 255, 255, 0.25);
  }

  /* Windows 11 Style Window Controls */
  .win-controls {
    display: flex;
    align-items: center;
    margin-left: 2px;
  }

  .win-btn {
    width: 36px;
    height: 28px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: transparent;
    border: none;
    cursor: pointer;
    color: var(--win-btn-color);
    border-radius: 4px;
    transition: all 0.15s ease;
  }

  .win-btn svg {
    width: 11px;
    height: 11px;
    stroke: currentColor;
    stroke-width: 1.5;
    fill: none;
  }

  .win-btn:hover {
    background: var(--win-btn-hover-bg);
    color: var(--win-btn-hover-color);
  }

  .win-btn.close:hover {
    background: #e81123;
    color: #ffffff;
  }

  /* Configuration Card */
  .config-card {
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
    padding: 13px 18px 14px 18px;
    margin-top: 8px;
    margin-bottom: 8px;
    transition: all 0.25s ease;
  }

  .config-title {
    font-size: 11.5px;
    font-weight: 700;
    letter-spacing: 0.8px;
    color: var(--card-title);
    text-transform: uppercase;
    margin-bottom: 10px;
    transition: color 0.25s ease;
  }

  /* Form Row */
  .form-row {
    display: flex;
    align-items: center;
    margin-bottom: 10px;
  }

  .form-label {
    width: 125px;
    font-size: 13px;
    font-weight: 500;
    color: var(--form-label);
    flex-shrink: 0;
    transition: color 0.25s ease;
  }

  .input-pill {
    flex: 1;
    height: 36px;
    border-radius: 9999px;
    background: var(--input-bg);
    border: 1px solid var(--input-border);
    display: flex;
    align-items: center;
    padding: 0 15px;
    transition: all 0.2s ease;
  }

  .input-pill:focus-within {
    border-color: var(--input-focus);
  }

  .input-icon {
    display: flex;
    align-items: center;
    color: var(--input-icon);
    flex-shrink: 0;
    transition: color 0.25s ease;
  }

  .input-icon svg {
    width: 16px;
    height: 16px;
    stroke: currentColor;
    stroke-width: 1.8;
    fill: none;
  }

  .input-divider {
    width: 1px;
    height: 17px;
    background: var(--input-border);
    margin: 0 11px;
    flex-shrink: 0;
    transition: background 0.25s ease;
  }

  .input-field {
    flex: 1;
    border: none;
    background: transparent;
    outline: none;
    font-size: 13px;
    color: var(--input-text);
    user-select: text;
    -webkit-user-select: text;
    transition: color 0.25s ease;
  }

  .input-field::placeholder {
    color: var(--input-placeholder);
    font-weight: 400;
  }

  /* Select Pill Buttons */
  .btn-select {
    height: 36px;
    width: 125px;
    padding: 0;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 9999px;
    background: var(--btn-select-bg);
    border: 1px solid var(--btn-select-border);
    box-shadow: var(--btn-select-shadow);
    font-size: 13px;
    font-weight: 600;
    color: var(--btn-select-text);
    cursor: pointer;
    margin-left: 12px;
    flex-shrink: 0;
    transition: all 0.2s ease;
  }

  .btn-select:hover {
    background: var(--btn-select-hover-bg);
    border-color: var(--btn-select-hover-border);
    color: var(--btn-select-hover-text);
    box-shadow: var(--btn-select-hover-shadow);
  }

  .btn-select:active {
    transform: translateY(0.5px);
  }

  /* Option Toggle Row */
  .toggle-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 10px 0 12px 2px;
    cursor: pointer;
    width: fit-content;
    user-select: none;
    -webkit-user-select: none;
  }

  .switch {
    position: relative;
    display: inline-block;
    width: 38px;
    height: 20px;
    flex-shrink: 0;
    pointer-events: none;
  }

  .switch input {
    opacity: 0;
    width: 0;
    height: 0;
  }

  .slider {
    position: absolute;
    cursor: pointer;
    top: 0; left: 0; right: 0; bottom: 0;
    background-color: var(--slider-bg);
    transition: 0.25s ease;
    border-radius: 9999px;
  }

  .slider:before {
    position: absolute;
    content: "";
    height: 16px;
    width: 16px;
    left: 2px;
    bottom: 2px;
    background-color: var(--slider-knob);
    transition: 0.25s ease;
    border-radius: 50%;
  }

  input:checked + .slider {
    background-color: var(--slider-checked);
    box-shadow: var(--slider-shadow);
  }

  input:checked + .slider:before {
    background-color: var(--slider-knob-checked);
    transform: translateX(18px);
  }

  .toggle-label {
    font-size: 13px;
    font-weight: 500;
    color: var(--toggle-label);
    transition: color 0.25s ease;
  }

  /* Action Buttons Bar */
  .action-bar {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-top: 2px;
  }

  /* Primary Start & Open Folder Button Styling */
  .btn-start, .btn-open-folder {
    height: 44px;
    border-radius: 9999px;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 10px;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 0.6px;
    cursor: pointer;
    transition: all 0.2s ease;
  }

  .btn-start {
    flex: 1.6;
  }

  .btn-open-folder {
    flex: 1;
  }

  /* Primary Accent State */
  .btn-primary-style {
    border: 1px solid var(--btn-primary-border);
    background: var(--btn-primary-bg);
    box-shadow: var(--btn-primary-shadow);
    color: var(--btn-primary-text);
    cursor: pointer;
  }

  .btn-primary-style:hover:not(:disabled) {
    background: var(--btn-primary-hover-bg);
    border-color: var(--btn-primary-hover-border);
    box-shadow: var(--btn-primary-hover-shadow);
  }

  .btn-primary-style:active:not(:disabled) {
    transform: translateY(0.5px);
  }

  /* Secondary / Deactivated Style */
  .btn-secondary-style {
    background: var(--btn-secondary-bg);
    border: 1px solid var(--btn-secondary-border);
    color: var(--btn-secondary-text);
  }

  .btn-secondary-style:hover:not(:disabled) {
    background: var(--btn-secondary-hover-bg);
    border-color: var(--btn-secondary-hover-border);
    color: var(--btn-secondary-hover-text);
  }

  .btn-start:disabled, .btn-open-folder:disabled {
    opacity: 0.38;
    cursor: not-allowed;
    transform: none !important;
  }

  .excel-badge {
    display: flex;
    align-items: center;
    justify-content: center;
    background: #107c41;
    border-radius: 4px;
    padding: 3px 5px;
  }

  .excel-badge svg {
    width: 13px;
    height: 13px;
    fill: #ffffff;
  }

  .btn-open-folder svg {
    width: 17px;
    height: 17px;
    stroke: currentColor;
    stroke-width: 1.8;
    fill: none;
  }

  /* Progress Bar */
  .progress-bar-wrap {
    width: 100%;
    height: 4px;
    background: var(--progress-wrap-bg);
    border-radius: 9999px;
    margin: 4px 0 8px 0;
    overflow: hidden;
    position: relative;
    transition: background 0.25s ease;
  }

  .progress-bar-fill {
    height: 100%;
    width: 0%;
    background: var(--progress-fill-bg);
    box-shadow: var(--progress-shadow);
    border-radius: 9999px;
    transition: width 0.25s ease;
  }

  /* Terminal Console Card */
  .terminal-card {
    flex: 1;
    min-height: 120px;
    border-radius: 10px;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    background: var(--term-bg);
    border: 1px solid var(--term-border);
    transition: all 0.25s ease;
  }

  .terminal-tab-bar {
    background: var(--term-header-bg);
    border-bottom: 1px solid var(--term-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 16px;
    height: 32px;
    transition: all 0.25s ease;
  }

  .terminal-header-title {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 12px;
    font-weight: 700;
    color: var(--term-header-title);
    letter-spacing: 0.6px;
    transition: color 0.25s ease;
  }

  .terminal-header-title .active-dot {
    width: 7px;
    height: 7px;
    border-radius: 50%;
    background: var(--term-dot-bg);
    box-shadow: var(--term-dot-shadow);
    transition: all 0.25s ease;
  }

  .btn-clear {
    display: flex;
    align-items: center;
    gap: 5px;
    font-size: 11.5px;
    font-weight: 600;
    color: var(--form-label);
    background: transparent;
    border: none;
    cursor: pointer;
    padding: 3px 7px;
    border-radius: 5px;
    transition: all 0.15s ease;
  }

  .btn-clear:hover {
    background: var(--win-btn-hover-bg);
    color: var(--app-text);
  }

  .btn-clear svg {
    width: 12px;
    height: 12px;
    stroke: currentColor;
    stroke-width: 2;
    fill: none;
  }

  /* Console Body */
  .terminal-body {
    flex: 1;
    overflow-y: auto;
    padding: 12px 16px;
    font-family: Consolas, "Cascadia Code", "Courier New", monospace;
    font-size: 12.5px;
    line-height: 1.6;
    color: var(--term-body-text);
    user-select: text;
    -webkit-user-select: text;
    white-space: pre-wrap;
    word-break: break-all;
    transition: color 0.25s ease;
  }

  .terminal-body::-webkit-scrollbar {
    width: 7px;
  }
  .terminal-body::-webkit-scrollbar-track {
    background: var(--term-scroll-track);
  }
  .terminal-body::-webkit-scrollbar-thumb {
    background: var(--term-scroll-thumb);
    border-radius: 4px;
  }
  .terminal-body::-webkit-scrollbar-thumb:hover {
    background: var(--card-border);
  }

  .log-vdom-name {
    color: var(--log-vdom-color, #81c995);
    text-shadow: var(--log-vdom-shadow, none);
    font-weight: 700;
  }

  /* Error Highlight Box */
  .log-line-error {
    display: inline-block;
    width: 100%;
    box-sizing: border-box;
    background: var(--log-error-bg);
    border-left: 3.5px solid var(--log-error-border);
    border-radius: 0 4px 4px 0;
    padding: 3px 8px;
    margin: 3px 0;
    color: var(--log-error-fg);
    font-weight: 600;
  }
  .log-badge-error {
    display: inline-block;
    background: var(--log-error-fg);
    color: #131314;
    font-size: 10px;
    font-weight: 800;
    padding: 1px 6px;
    border-radius: 3px;
    margin-right: 6px;
    letter-spacing: 0.5px;
    vertical-align: 1px;
  }

  /* Warning Highlight Box */
  .log-line-warning {
    display: inline-block;
    width: 100%;
    box-sizing: border-box;
    background: var(--log-warning-bg);
    border-left: 3.5px solid var(--log-warning-border);
    border-radius: 0 4px 4px 0;
    padding: 3px 8px;
    margin: 3px 0;
    color: var(--log-warning-fg);
    font-weight: 600;
  }
  .log-badge-warning {
    display: inline-block;
    background: var(--log-warning-fg);
    color: #131314;
    font-size: 10px;
    font-weight: 800;
    padding: 1px 6px;
    border-radius: 3px;
    margin-right: 6px;
    letter-spacing: 0.5px;
    vertical-align: 1px;
  }

  /* Debug Highlight Box */
  .log-line-debug {
    display: inline-block;
    width: 100%;
    box-sizing: border-box;
    background: var(--log-debug-bg);
    border-left: 3.5px solid var(--log-debug-border);
    border-radius: 0 4px 4px 0;
    padding: 2px 8px;
    margin: 2px 0;
    color: var(--log-debug-fg);
    font-weight: 500;
  }
  .log-badge-debug {
    display: inline-block;
    background: var(--log-debug-fg);
    color: #131314;
    font-size: 10px;
    font-weight: 800;
    padding: 1px 6px;
    border-radius: 3px;
    margin-right: 6px;
    letter-spacing: 0.5px;
    vertical-align: 1px;
  }

  /* Traceback Details Highlight */
  .log-line-trace {
    color: var(--log-trace-fg);
    font-weight: 500;
  }

  .cursor-block {
    display: inline-block;
    width: 8px;
    height: 15px;
    background: var(--term-cursor);
    vertical-align: middle;
    margin-left: 4px;
    animation: blink 1s step-end infinite;
  }

  @keyframes blink {
    0%, 100% { opacity: 1; }
    50% { opacity: 0; }
  }

  /* Status Bar */
  .status-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 8px 4px 0 4px;
    font-size: 11.5px;
    font-weight: 600;
    color: var(--status-bar-text);
    flex-shrink: 0;
    transition: color 0.25s ease;
  }

  .status-left {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .status-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: var(--status-dot-ready);
    box-shadow: var(--status-dot-shadow);
    transition: background 0.3s ease, box-shadow 0.3s ease;
  }

  .status-dot.running {
    background: var(--status-dot-running);
    animation: pulse 1.2s infinite;
  }

  .status-dot.done {
    background: var(--status-dot-done);
  }

  .status-dot.error {
    background: var(--status-dot-error);
  }

  @keyframes pulse {
    0% { transform: scale(0.95); opacity: 0.8; }
    50% { transform: scale(1.15); opacity: 1; }
    100% { transform: scale(0.95); opacity: 0.8; }
  }

  .status-text {
    letter-spacing: 0.5px;
    text-transform: uppercase;
    font-size: 11px;
    color: var(--status-bar-text);
    transition: color 0.25s ease;
  }

  .status-divider {
    width: 1px;
    height: 13px;
    background: var(--status-divider);
    margin: 0 4px;
    transition: background 0.25s ease;
  }

  .status-right a {
    display: flex;
    align-items: center;
    gap: 6px;
    color: var(--status-link);
    text-decoration: none;
    cursor: pointer;
    transition: color 0.15s ease;
  }

  .status-right a:hover {
    color: var(--status-link-hover);
  }

  .status-right svg {
    width: 14.5px;
    height: 14.5px;
    fill: currentColor;
  }
</style>
</head>
<body data-theme="__INITIAL_THEME__">

<!-- 8 Native Resize Borders & Corners -->
<div class="resize-edge resize-top" onmousedown="handleBorderResize(event, 'top')"></div>
<div class="resize-edge resize-bottom" onmousedown="handleBorderResize(event, 'bottom')"></div>
<div class="resize-edge resize-left" onmousedown="handleBorderResize(event, 'left')"></div>
<div class="resize-edge resize-right" onmousedown="handleBorderResize(event, 'right')"></div>
<div class="resize-edge resize-topleft" onmousedown="handleBorderResize(event, 'topleft')"></div>
<div class="resize-edge resize-topright" onmousedown="handleBorderResize(event, 'topright')"></div>
<div class="resize-edge resize-bottomleft" onmousedown="handleBorderResize(event, 'bottomleft')"></div>
<div class="resize-edge resize-bottomright" onmousedown="handleBorderResize(event, 'bottomright')"></div>

<div id="windowRoot" class="window-root active-window">

  <!-- Header / Titlebar - Only this area initiates window drag -->
  <div class="title-bar" onmousedown="handleTitlebarMouseDown(event)" ondblclick="handleTitlebarDblClick(event)">
    <div class="title-left">
      <!-- Authentic Fortinet Logo -->
      <div class="brand-logo-wrap">
        <img src="data:image/png;base64,__FORTINET_LOGO_B64__" class="fortinet-brand-logo" alt="Fortinet Logo">
      </div>

      <!-- Brand Title -->
      <div class="brand-title-wrap">
        <span class="brand-title">FORTIGATE POLICY EXPORTER</span>
        <span class="brand-version">v2.1</span>
      </div>
    </div>

    <!-- Right: Theme Badge & Windows Control Buttons -->
    <div class="title-right" onmousedown="event.stopPropagation()">
      <!-- Oval Theme Selector Badge -->
      <div id="themeBadgeWrapper" class="theme-badge-wrapper">
        <button id="themeBadgeBtn" class="theme-badge" onclick="toggleThemeDropdown(event)" title="Change Color Theme">
          <svg class="palette-icon" viewBox="0 0 24 24">
            <circle cx="12" cy="12" r="10"/>
            <path d="M12 2a10 10 0 0 1 10 10c0 2.5-2 4.5-4.5 4.5H16a2 2 0 0 0-2 2v.5c0 1.5-1 2.5-2.5 2.5C6.5 21.5 2 17 2 12A10 10 0 0 1 12 2z"/>
            <circle cx="7.5" cy="10.5" r="1.5" fill="currentColor"/>
            <circle cx="12" cy="7.5" r="1.5" fill="currentColor"/>
            <circle cx="16.5" cy="10.5" r="1.5" fill="currentColor"/>
          </svg>
          <span id="themeBadgeLabel">__INITIAL_THEME_LABEL__</span>
          <svg class="chevron-icon" viewBox="0 0 16 16">
            <polyline points="4 6 8 10 12 6"/>
          </svg>
        </button>
        <!-- Dropdown Menu with 4 Options -->
        <div class="theme-dropdown">
          <div class="theme-item" onclick="selectTheme('antigravity-dark')" data-theme-id="antigravity-dark">
            <span class="theme-preview-dot" style="background: #131314; border-color: #8ab4f8;"></span>
            <span>Antigravity Dark</span>
          </div>
          <div class="theme-item" onclick="selectTheme('antigravity-light')" data-theme-id="antigravity-light">
            <span class="theme-preview-dot" style="background: #ffffff; border-color: #1a73e8;"></span>
            <span>Antigravity Light</span>
          </div>
          <div class="theme-item" onclick="selectTheme('cyber-dark')" data-theme-id="cyber-dark">
            <span class="theme-preview-dot" style="background: #0D121B; border-color: #38bdf8; box-shadow: 0 0 6px #38bdf8;"></span>
            <span>Cyber Dark (Neon)</span>
          </div>
          <div class="theme-item" onclick="selectTheme('cyber-light')" data-theme-id="cyber-light">
            <span class="theme-preview-dot" style="background: #eef2f8; border-color: #0284c7; box-shadow: 0 0 6px #0284c7;"></span>
            <span>Cyber Light (Sci-Fi)</span>
          </div>
        </div>
      </div>

      <!-- Windows Controls -->
      <div class="win-controls" onmousedown="event.stopPropagation()">
        <button class="win-btn" onclick="handleMinimizeWindow()" title="Minimize">
          <svg viewBox="0 0 16 16"><line x1="2" y1="8" x2="14" y2="8"/></svg>
        </button>
        <button id="btnMax" class="win-btn" onclick="handleToggleMaximize()" title="Maximize">
          <svg viewBox="0 0 16 16"><rect x="3" y="3" width="10" height="10" rx="1"/></svg>
        </button>
        <button class="win-btn close" onclick="handleCloseWindow()" title="Close">
          <svg viewBox="0 0 16 16"><line x1="3" y1="3" x2="13" y2="13"/><line x1="13" y1="3" x2="3" y2="13"/></svg>
        </button>
      </div>
    </div>
  </div>

  <!-- Configuration Card -->
  <div class="config-card">
    <div class="config-title">CONFIGURATION</div>

    <!-- Row 1: Config File -->
    <div class="form-row">
      <div class="form-label">Config File</div>
      <div class="input-pill">
        <div class="input-icon">
          <svg viewBox="0 0 24 24">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <line x1="16" y1="13" x2="8" y2="13"/>
            <line x1="16" y1="17" x2="8" y2="17"/>
            <polyline points="10 9 9 9 8 9"/>
          </svg>
        </div>
        <div class="input-divider"></div>
        <input type="text" id="configFile" class="input-field" placeholder="Select FortiGate .conf or .txt file..." readonly>
      </div>
      <button class="btn-select" onclick="handleSelectFile()">Select File</button>
    </div>

    <!-- Row 2: Output Directory -->
    <div class="form-row">
      <div class="form-label">Output Directory</div>
      <div class="input-pill">
        <div class="input-icon">
          <svg viewBox="0 0 24 24">
            <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
          </svg>
        </div>
        <div class="input-divider"></div>
        <input type="text" id="outputDir" class="input-field" placeholder="Default: Same folder as config file">
      </div>
      <button class="btn-select" onclick="handleSelectFolder()">Select Folder</button>
    </div>

    <!-- Row 3: Option Toggle -->
    <label class="toggle-row" for="autoOpenFolder">
      <span class="switch">
        <input type="checkbox" id="autoOpenFolder" __INITIAL_AUTO_OPEN_CHECKED__ onchange="handleAutoOpenChange(this.checked)">
        <span class="slider"></span>
      </span>
      <span class="toggle-label">Automatically open result folder upon completion</span>
    </label>

    <!-- Row 4: Action Buttons -->
    <div class="action-bar">
      <!-- Main Start Export Button: disabled initially until file selected -->
      <button id="btnStart" class="btn-start btn-primary-style" onclick="handleStartExport()" disabled>
        <div class="excel-badge">
          <svg viewBox="0 0 24 24">
            <path d="M7 6L11 12L7 18H9.5L12 14L14.5 18H17L13 12L17 6H14.5L12 10L9.5 6H7Z"/>
          </svg>
        </div>
        <span>START EXPORT TO EXCEL</span>
      </button>

      <!-- Secondary Open Result Button: disabled initially -->
      <button id="btnOpenResult" class="btn-open-folder btn-secondary-style" onclick="handleOpenResultFolder()" disabled>
        <svg viewBox="0 0 24 24">
          <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
        </svg>
        <span>Open Result Folder</span>
      </button>
    </div>
  </div>

  <!-- Cyber Blue Progress Bar -->
  <div class="progress-bar-wrap">
    <div id="progressBar" class="progress-bar-fill"></div>
  </div>

  <!-- Terminal Console Card -->
  <div class="terminal-card">
    <div class="terminal-tab-bar">
      <div class="terminal-header-title">
        <span class="active-dot"></span>
        <span>OUTPUT</span>
      </div>
      <button class="btn-clear" onclick="handleClearLog()" title="Clear Console Log">
        <svg viewBox="0 0 24 24">
          <path d="M23 4v6h-6"/>
          <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
        </svg>
        <span>Clear</span>
      </button>
    </div>

    <!-- Terminal Text Area -->
    <div id="terminalBody" class="terminal-body">Fortigate Config -> Excel Converter initialized.
Select a .conf file and click 'Start Conversion' to begin.

Fortigate Config -> <span class="cursor-block"></span></div>
  </div>

  <!-- Bottom Status Bar -->
  <div class="status-bar">
    <div class="status-left">
      <div id="statusDot" class="status-dot"></div>
      <span id="statusText" class="status-text">READY</span>
      <div class="status-divider"></div>
    </div>
    <div class="status-right">
      <a href="#" onclick="handleOpenGithub(); return false;">
        <svg viewBox="0 0 24 24">
          <path d="M12 0C5.37 0 0 5.37 0 12c0 5.31 3.435 9.795 8.205 11.385.6.105.825-.255.825-.57 0-.285-.015-1.23-.015-2.235-3.015.555-3.795-.735-4.035-1.41-.135-.345-.72-1.41-1.23-1.695-.42-.225-1.02-.78-.015-.795.945-.015 1.62.87 1.845 1.23 1.08 1.815 2.805 1.305 3.495.99.105-.78.42-1.305.765-1.605-2.67-.3-5.46-1.335-5.46-5.925 0-1.305.465-2.385 1.23-3.225-.12-.3-.54-1.53.12-3.18 0 0 1.005-.315 3.3 1.23.96-.27 1.98-.405 3-.405s2.04.135 3 .405c2.295-1.56 3.3-1.23 3.3-1.23.66 1.65.24 2.88.12 3.18.765.84 1.23 1.905 1.23 3.225 0 4.605-2.805 5.625-5.475 5.925.435.375.81 1.095.81 2.22 0 1.605-.015 2.895-.015 3.3 0 .315.225.69.825.57A12.02 12.02 0 0 0 24 12c0-6.63-5.37-12-12-12z"/>
        </svg>
        <span>github.com/smilestory-net</span>
      </a>
    </div>
  </div>

</div>

<script>
  let logContent = "Fortigate Config -> Excel Converter initialized.\nSelect a .conf file and click 'Start Conversion' to begin.\n\nFortigate Config -> ";

  // Theme Definition Map
  const THEMES = {
    'antigravity-dark': {
      name: 'Antigravity Dark',
      shortName: 'Antigravity Dark',
      isDark: true
    },
    'antigravity-light': {
      name: 'Antigravity Light',
      shortName: 'Antigravity Light',
      isDark: false
    },
    'cyber-dark': {
      name: 'Cyber Dark (Neon)',
      shortName: 'Cyber Dark',
      isDark: true
    },
    'cyber-light': {
      name: 'Cyber Light (Sci-Fi)',
      shortName: 'Cyber Light',
      isDark: false
    }
  };

  const INITIAL_THEME = "__INITIAL_THEME__";
  let currentTheme = (INITIAL_THEME && THEMES[INITIAL_THEME]) ? INITIAL_THEME : 'antigravity-dark';

  function initTheme() {
    applyTheme(currentTheme, false);
  }

  function applyTheme(themeId, notifyPython = true) {
    if (!THEMES[themeId]) return;
    currentTheme = themeId;
    document.body.setAttribute('data-theme', themeId);
    const label = document.getElementById('themeBadgeLabel');
    if (label) {
      label.innerText = THEMES[themeId].shortName;
    }

    // Update active class in dropdown items
    document.querySelectorAll('.theme-item').forEach(el => {
      if (el.getAttribute('data-theme-id') === themeId) {
        el.classList.add('active');
      } else {
        el.classList.remove('active');
      }
    });

    // Inform native Windows DWM and Python backend to persist theme to disk
    if (notifyPython && window.pywebview && window.pywebview.api && window.pywebview.api.notify_theme_changed) {
      window.pywebview.api.notify_theme_changed(themeId, THEMES[themeId].isDark);
    }
  }

  function toggleThemeDropdown(event) {
    event.stopPropagation();
    const wrapper = document.getElementById('themeBadgeWrapper');
    if (wrapper) {
      wrapper.classList.toggle('open');
    }
  }

  function selectTheme(themeId) {
    applyTheme(themeId, true);
    const wrapper = document.getElementById('themeBadgeWrapper');
    if (wrapper) {
      wrapper.classList.remove('open');
    }
  }

  // Close dropdown when clicking anywhere outside
  window.addEventListener('click', (e) => {
    const wrapper = document.getElementById('themeBadgeWrapper');
    if (wrapper && !wrapper.contains(e.target)) {
      wrapper.classList.remove('open');
    }
  });

  // Window Focus & Blur Accent Border Toggle + Smooth Restore Entrance
  window.addEventListener('focus', () => {
    const root = document.getElementById('windowRoot');
    if (root) {
      root.classList.add('active-window');
      root.style.transition = 'transform 0.25s cubic-bezier(0, 0, 0.2, 1), opacity 0.22s ease-out';
      root.style.transform = 'scale(1) translateY(0)';
      root.style.opacity = '1';
      setTimeout(() => {
        root.style.transition = '';
        root.style.transform = '';
        root.style.opacity = '';
      }, 260);
    }
    if (window.pywebview && window.pywebview.api) {
      window.pywebview.api.set_window_active(true);
    }
  });

  window.addEventListener('blur', () => {
    const root = document.getElementById('windowRoot');
    if (root) root.classList.remove('active-window');
    if (window.pywebview && window.pywebview.api) {
      window.pywebview.api.set_window_active(false);
    }
  });

  initTheme();

  window.addEventListener('DOMContentLoaded', () => {
    initTheme();
    setTimeout(async () => {
      if (window.pywebview && window.pywebview.api) {
        try {
          const bounds = await window.pywebview.api.get_window_bounds();
          if (bounds && bounds.length >= 4) {
            savedNormalBounds = [...bounds];
          }
        } catch (e) {}
      }
    }, 150);
  });

  function handleClearLog() {
    logContent = "";
    renderTerminal();
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.innerText = text;
    return div.innerHTML;
  }

  function formatLogHtml(rawText) {
    if (!rawText) return '';
    const lines = rawText.split('\n');
    const formattedLines = lines.map(line => {
      let escaped = escapeHtml(line);

      // 1. Highlight vDOM name in green for "Parsing vDOM '<name>'"
      escaped = escaped.replace(/(Parsing vDOM\s+(&#39;|&quot;|['"]))(.*?)(&quot;|&#39;|['"])/g, (match, prefix, qOpen, vdom, qClose) => {
        return `${prefix}<span class="log-vdom-name">${vdom}</span>${qClose}`;
      });

      // 2. Highlight [ERROR] / ERROR: lines
      if (/^\s*(\[ERROR\]|ERROR:)/.test(line)) {
        return escaped.replace(/^(\s*)(\[ERROR\]|ERROR:)(.*)$/, (m, indent, badge, rest) => {
          return `${indent}<span class="log-line-error"><span class="log-badge-error">[ERROR]</span>${rest}</span>`;
        });
      }

      // 3. Highlight [WARNING] / WARNING: lines
      if (/^\s*(\[WARNING\]|WARNING:)/.test(line)) {
        return escaped.replace(/^(\s*)(\[WARNING\]|WARNING:)(.*)$/, (m, indent, badge, rest) => {
          return `${indent}<span class="log-line-warning"><span class="log-badge-warning">[WARNING]</span>${rest}</span>`;
        });
      }

      // 4. Highlight [DEBUG] / DEBUG: lines
      if (/^\s*(\[DEBUG\]|DEBUG:)/.test(line)) {
        return escaped.replace(/^(\s*)(\[DEBUG\]|DEBUG:)(.*)$/, (m, indent, badge, rest) => {
          return `${indent}<span class="log-line-debug"><span class="log-badge-debug">[DEBUG]</span>${rest}</span>`;
        });
      }

      // 5. Highlight Traceback frames and exception details
      if (/^\s*(Traceback \(most recent call last\):|File\s+(&quot;|&#39;|['"]).*?(&quot;|&#39;|['"]),\s+line\s+\d+|During handling of the above exception|The above exception was the direct cause|[A-Za-z0-9_]+Error:|[A-Za-z0-9_]+Exception:)/.test(line)) {
        return `<span class="log-line-trace">${escaped}</span>`;
      }

      return escaped;
    });

    return formattedLines.join('\n');
  }

  function renderTerminal() {
    const term = document.getElementById('terminalBody');
    term.innerHTML = formatLogHtml(logContent) + '<span class="cursor-block"></span>';
    term.scrollTop = term.scrollHeight;
  }

  function appendLog(text) {
    logContent += (logContent ? "\n" : "") + text;
    renderTerminal();
  }

  function setStatus(text, state = 'ready') {
    const sText = document.getElementById('statusText');
    const sDot = document.getElementById('statusDot');
    sText.innerText = text;
    sDot.className = 'status-dot ' + state;
  }

  function setProgress(pct) {
    const pBar = document.getElementById('progressBar');
    if (pBar) {
      pBar.style.width = Math.min(100, Math.max(0, pct)) + '%';
    }
  }

  /* Universal Titlebar Drag & Restore Engine */
  let isTitlebarDragging = false;
  let dragStartScreenX = 0, dragStartScreenY = 0;
  let dragStartWinX = 0, dragStartWinY = 0;
  let dragStartWinW = 0, dragStartWinH = 0;
  let dragHasMoved = false;
  let dragWasMaximized = false;
  let dragRatioX = 0.5;
  let dragRafId = null;

  async function handleTitlebarMouseDown(event) {
    if (event.button !== 0) return;
    if (event.target.closest('.win-controls') || event.target.closest('button') || event.target.closest('input')) {
      return;
    }
    if (!window.pywebview || !window.pywebview.api) return;

    event.preventDefault();

    const isMax = await window.pywebview.api.is_maximized();
    isTitlebarDragging = true;
    dragHasMoved = false;
    dragWasMaximized = isMax;
    dragStartScreenX = event.screenX;
    dragStartScreenY = event.screenY;

    if (isMax) {
      dragRatioX = event.clientX / window.innerWidth;
      dragRatioX = Math.max(0.05, Math.min(0.85, dragRatioX));
      const dims = await window.pywebview.api.get_normal_dimensions();
      dragStartWinW = dims ? dims[0] : 1040;
      dragStartWinH = dims ? dims[1] : 720;
    } else {
      const bounds = await window.pywebview.api.get_window_bounds();
      if (bounds && bounds.length >= 4) {
        dragStartWinX = bounds[0];
        dragStartWinY = bounds[1];
        dragStartWinW = bounds[2];
        dragStartWinH = bounds[3];
      } else {
        dragStartWinW = 1040;
        dragStartWinH = 720;
      }
    }

    window.addEventListener('mousemove', onTitlebarMouseMove, { passive: true });
    window.addEventListener('mouseup', onTitlebarMouseUp);
  }

  function onTitlebarMouseMove(event) {
    if (!isTitlebarDragging) return;
    if (dragRafId) cancelAnimationFrame(dragRafId);

    const clientScreenX = event.screenX;
    const clientScreenY = event.screenY;

    dragRafId = requestAnimationFrame(async () => {
      if (!isTitlebarDragging) return;

      const dpr = window.devicePixelRatio || 1;
      const dx = Math.round((clientScreenX - dragStartScreenX) * dpr);
      const dy = Math.round((clientScreenY - dragStartScreenY) * dpr);

      if (!dragHasMoved) {
        if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
          dragHasMoved = true;
          if (dragWasMaximized) {
            const normalW = dragStartWinW;
            const normalH = dragStartWinH;
            dragStartWinX = Math.round(clientScreenX * dpr - (normalW * dragRatioX));
            dragStartWinY = Math.round(clientScreenY * dpr - (15 * dpr));
            dragStartScreenX = clientScreenX;
            dragStartScreenY = clientScreenY;
            dragWasMaximized = false;

            const root = document.getElementById('windowRoot');
            const btn = document.getElementById('btnMax');
            if (root) root.classList.remove('maximized');
            document.body.classList.remove('is-maximized');
            if (btn) {
              btn.innerHTML = '<svg viewBox="0 0 16 16"><rect x="3" y="3" width="10" height="10" rx="1"/></svg>';
              btn.title = "Maximize";
            }

            await window.pywebview.api.set_window_bounds(dragStartWinX, dragStartWinY, normalW, normalH);
            return;
          }
        } else {
          return;
        }
      }

      if (dragWasMaximized) {
        const normalW = dragStartWinW;
        const normalH = dragStartWinH;
        dragStartWinX = Math.round(clientScreenX * dpr - (normalW * dragRatioX));
        dragStartWinY = Math.round(clientScreenY * dpr - (15 * dpr));
        dragStartScreenX = clientScreenX;
        dragStartScreenY = clientScreenY;
        dragWasMaximized = false;
        await window.pywebview.api.set_window_bounds(dragStartWinX, dragStartWinY, normalW, normalH);
        return;
      }

      const newX = dragStartWinX + Math.round((clientScreenX - dragStartScreenX) * dpr);
      const newY = dragStartWinY + Math.round((clientScreenY - dragStartScreenY) * dpr);

      window.pywebview.api.move_window(newX, newY);
    });
  }

  async function onTitlebarMouseUp(event) {
    if (!isTitlebarDragging) return;
    isTitlebarDragging = false;
    if (dragRafId) {
      cancelAnimationFrame(dragRafId);
      dragRafId = null;
    }
    window.removeEventListener('mousemove', onTitlebarMouseMove);
    window.removeEventListener('mouseup', onTitlebarMouseUp);

    // If dragged to the very top edge of the screen, auto-maximize (Aero Snap to top)
    if (dragHasMoved && event.screenY <= 4) {
      window.pywebview.api.toggle_maximize();
    }
  }

  /* Titlebar Double-Click - Maximize / Restore Toggle */
  function handleTitlebarDblClick(event) {
    if (event.button !== 0) return;
    if (event.target.closest('.win-controls') || event.target.closest('button') || event.target.closest('input')) {
      return;
    }
    if (dragHasMoved) {
      return;
    }
    handleToggleMaximize();
  }

  /* Maximize & Restore Toggle via Native Win32 ShowWindow (Genuine DWM Animation) */
  async function handleToggleMaximize() {
    if (!window.pywebview || !window.pywebview.api) return;
    window.pywebview.api.toggle_maximize();
  }

  /* Native Minimize Window with DWM Animation */
  function handleMinimizeWindow() {
    if (window.pywebview && window.pywebview.api) {
      window.pywebview.api.minimize_window();
    }
  }

  /* Auto-open Toggle State Handler */
  function handleAutoOpenChange(checked) {
    if (window.pywebview && window.pywebview.api) {
      window.pywebview.api.save_auto_open(checked);
    }
  }

  /* Close Window with Auto-Open State Preservation */
  async function handleCloseWindow() {
    try {
      const toggle = document.getElementById('autoOpenFolder');
      const isChecked = toggle ? toggle.checked : true;
      if (window.pywebview && window.pywebview.api) {
        await window.pywebview.api.save_auto_open(isChecked);
        window.pywebview.api.close_window(isChecked);
        return;
      }
    } catch (err) {
      console.error(err);
    }
    if (window.pywebview && window.pywebview.api) {
      window.pywebview.api.close_window();
    }
  }

  // Synchronize maximize state and button icon when window state changes
  window.addEventListener('resize', async () => {
    if (isSizing) return;
    if (!window.pywebview || !window.pywebview.api) return;
    try {
      const isMax = await window.pywebview.api.is_maximized();
      const root = document.getElementById('windowRoot');
      const btn = document.getElementById('btnMax');
      if (isMax) {
        if (root) root.classList.add('maximized');
        document.body.classList.add('is-maximized');
        if (btn) {
          btn.innerHTML = '<svg viewBox="0 0 16 16"><rect x="4" y="2" width="9" height="9" rx="1"/><rect x="2" y="5" width="9" height="9" rx="1" fill="#111620"/></svg>';
          btn.title = "Restore";
        }
      } else {
        if (root) root.classList.remove('maximized');
        document.body.classList.remove('is-maximized');
        if (btn) {
          btn.innerHTML = '<svg viewBox="0 0 16 16"><rect x="3" y="3" width="10" height="10" rx="1"/></svg>';
          btn.title = "Maximize";
        }
      }
    } catch (e) {}
  });

  /* 8-Directional Window Resize Engine (Smooth, Real-Time, Anchored, Zero Jitter) */
  let isSizing = false;
  let sizeDir = '';
  let startMouseX = 0, startMouseY = 0;
  let startWinX = 0, startWinY = 0;
  let startWinW = 0, startWinH = 0;
  let fixedRight = 0, fixedBottom = 0;
  let resizeRafId = null;

  async function handleBorderResize(e, edge) {
    e.preventDefault();
    e.stopPropagation();
    if (!window.pywebview || !window.pywebview.api) return;

    const bounds = await window.pywebview.api.get_window_bounds();
    if (!bounds || bounds.length < 4) return;

    isSizing = true;
    sizeDir = edge;
    startMouseX = e.screenX;
    startMouseY = e.screenY;
    startWinX = bounds[0];
    startWinY = bounds[1];
    startWinW = bounds[2];
    startWinH = bounds[3];
    fixedRight = startWinX + startWinW;
    fixedBottom = startWinY + startWinH;

    window.addEventListener('mousemove', onBorderMouseMove, { passive: true });
    window.addEventListener('mouseup', onBorderMouseUp);
  }

  function onBorderMouseMove(e) {
    if (!isSizing) return;
    if (resizeRafId) cancelAnimationFrame(resizeRafId);

    const clientScreenX = e.screenX;
    const clientScreenY = e.screenY;

    resizeRafId = requestAnimationFrame(() => {
      if (!isSizing) return;
      const dpr = window.devicePixelRatio || 1;
      const dx = Math.round((clientScreenX - startMouseX) * dpr);
      const dy = Math.round((clientScreenY - startMouseY) * dpr);

      let newX = startWinX;
      let newY = startWinY;
      let newW = startWinW;
      let newH = startWinH;

      const MIN_W = Math.round(920 * dpr);
      const MIN_H = Math.round(650 * dpr);

      // Horizontal sizing with opposite edge strictly anchored
      if (sizeDir.includes('right')) {
        newW = Math.max(MIN_W, startWinW + dx);
        newX = startWinX;
      } else if (sizeDir.includes('left')) {
        newW = Math.max(MIN_W, startWinW - dx);
        newX = fixedRight - newW;
      }

      // Vertical sizing with opposite edge strictly anchored (ZERO jitter)
      if (sizeDir.includes('bottom')) {
        newH = Math.max(MIN_H, startWinH + dy);
        newY = startWinY;
      } else if (sizeDir.includes('top')) {
        newH = Math.max(MIN_H, startWinH - dy);
        newY = fixedBottom - newH;
      }

      window.pywebview.api.set_window_bounds(newX, newY, newW, newH);
    });
  }

  function onBorderMouseUp() {
    isSizing = false;
    if (resizeRafId) {
      cancelAnimationFrame(resizeRafId);
      resizeRafId = null;
    }
    window.removeEventListener('mousemove', onBorderMouseMove);
    window.removeEventListener('mouseup', onBorderMouseUp);
  }

  async function handleSelectFile() {
    try {
      const filePath = await window.pywebview.api.select_file();
      if (filePath) {
        document.getElementById('configFile').value = filePath;
        const currentOut = document.getElementById('outputDir').value.trim();
        if (!currentOut) {
          const sep = filePath.lastIndexOf('\\') !== -1 ? '\\' : '/';
          const dir = filePath.substring(0, filePath.lastIndexOf(sep));
          document.getElementById('outputDir').value = dir;
        }

        // Button State Flow: Enable Start Export, deactivate & disable Open Result Folder
        const btnStart = document.getElementById('btnStart');
        const btnOpenResult = document.getElementById('btnOpenResult');

        btnStart.disabled = false;
        btnStart.className = 'btn-start btn-primary-style';

        btnOpenResult.disabled = true;
        btnOpenResult.className = 'btn-open-folder btn-secondary-style';
      }
    } catch (err) {
      console.error(err);
    }
  }

  async function handleSelectFolder() {
    try {
      const folderPath = await window.pywebview.api.select_folder();
      if (folderPath) {
        document.getElementById('outputDir').value = folderPath;
      }
    } catch (err) {
      console.error(err);
    }
  }

  async function handleStartExport() {
    const confFile = document.getElementById('configFile').value.trim();
    const outDir = document.getElementById('outputDir').value.trim();
    const autoOpen = document.getElementById('autoOpenFolder').checked;

    if (!confFile) {
      alert("Please select a FortiGate configuration file (.conf).");
      return;
    }

    const btnStart = document.getElementById('btnStart');
    const btnOpenResult = document.getElementById('btnOpenResult');
    btnStart.disabled = true;
    btnOpenResult.disabled = true;
    setProgress(0);
    setStatus('PROCESSING...', 'running');

    try {
      await window.pywebview.api.start_conversion(confFile, outDir, autoOpen);
    } catch (err) {
      console.error(err);
      setStatus('ERROR', 'error');
      btnStart.disabled = false;
    }
  }

  function handleOpenResultFolder() {
    window.pywebview.api.open_result_folder();
  }

  function handleOpenGithub() {
    window.pywebview.api.open_external_link('https://github.com/smilestory-net');
  }

  function onExportFinished(success, message) {
    const btnStart = document.getElementById('btnStart');
    const btnOpenResult = document.getElementById('btnOpenResult');

    if (success) {
      setProgress(100);
      setStatus('COMPLETED', 'done');
      // Button State Flow: Deactivate Start Export, highlight Open Result Folder
      btnStart.disabled = true;
      btnStart.className = 'btn-start btn-secondary-style';

      btnOpenResult.disabled = false;
      btnOpenResult.className = 'btn-open-folder btn-primary-style';
    } else {
      setStatus('ERROR', 'error');
      btnStart.disabled = false;
      btnStart.className = 'btn-start btn-primary-style';
      alert("Export failed: " + message);
    }
  }
</script>

</body>
"""


def get_html_template(theme_id=None, auto_open=None):
    """Generate dynamic HTML template with the specified or saved theme and auto_open state."""
    cfg = load_app_config()
    if not theme_id or theme_id not in VALID_THEMES:
        theme_id = cfg.get("theme", "antigravity-dark")
    if auto_open is None:
        auto_open = cfg.get("auto_open", True)
    theme_info = VALID_THEMES.get(theme_id, VALID_THEMES["antigravity-dark"])
    theme_label = theme_info.get("label", "Antigravity Dark")

    html = HTML_TEMPLATE_RAW.replace("__FORTINET_LOGO_B64__", FORTINET_LOGO_B64)
    html = html.replace("__WIN_ACCENT_COLOR__", get_system_accent_color())
    html = html.replace("__INITIAL_THEME__", theme_id)
    html = html.replace("__INITIAL_THEME_LABEL__", theme_label)
    html = html.replace("__INITIAL_AUTO_OPEN_CHECKED__", "checked" if auto_open else "")
    return html


def __getattr__(name):
    if name == 'HTML_TEMPLATE':
        return get_html_template()
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")



class WebviewApi:
    def __init__(self, run_converter_callback, initial_theme="antigravity-dark", initial_auto_open=True):
        self._window = None
        self.hwnd = None
        self._run_converter_callback = run_converter_callback
        self.last_target_dir = ""
        self.is_running = False
        self.current_theme = initial_theme
        self.is_dark_theme = VALID_THEMES.get(initial_theme, {}).get("is_dark", True)
        self.auto_open_state = initial_auto_open

    def set_window(self, window):
        self._window = window

    def set_hwnd(self, hwnd):
        self.hwnd = hwnd

    def notify_theme_changed(self, theme_id, is_dark):
        """Called from frontend when user selects a theme to persist both to disk and native window"""
        self.current_theme = str(theme_id)
        self.is_dark_theme = bool(is_dark)
        save_theme_to_file(self.current_theme)
        if self.hwnd and dwmapi:
            try:
                # DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                dark_flag = ctypes.c_int(1 if self.is_dark_theme else 0)
                dwmapi.DwmSetWindowAttribute(self.hwnd, 20, ctypes.byref(dark_flag), ctypes.sizeof(dark_flag))
                # Refresh current border
                self.set_window_active(True)
            except Exception as e:
                print("notify_theme_changed error:", e)

    def set_window_active(self, is_active):
        """Update window border color to Windows accent color on focus, or subtle dark/light border on blur"""
        if self.hwnd and dwmapi:
            try:
                if is_active:
                    color = get_system_accent_color_ref()
                else:
                    # Inactive border: dark gray for dark mode, subtle silver for light mode
                    color = 0x00333333 if self.is_dark_theme else 0x00D0D0D0
                dwmapi.DwmSetWindowAttribute(self.hwnd, 34, ctypes.byref(ctypes.c_uint(color)), 4)
            except Exception as e:
                print("set_window_active error:", e)

    def start_window_drag(self):
        """Native Windows window dragging loop triggered ONLY from the title bar"""
        if self.hwnd:
            try:
                pt = wintypes.POINT()
                user32.GetCursorPos(ctypes.byref(pt))

                if user32.IsZoomed(self.hwnd):
                    # Pre-update rcNormalPosition so Windows Shell's built-in DragFromMaximize
                    # restores the window anchored proportionally under the mouse cursor!
                    rect_win = wintypes.RECT()
                    user32.GetWindowRect(self.hwnd, ctypes.byref(rect_win))
                    win_w = rect_win.right - rect_win.left

                    wp = WINDOWPLACEMENT()
                    wp.length = ctypes.sizeof(WINDOWPLACEMENT)
                    user32.GetWindowPlacement(self.hwnd, ctypes.byref(wp))
                    normal_w = wp.rcNormalPosition.right - wp.rcNormalPosition.left
                    normal_h = wp.rcNormalPosition.bottom - wp.rcNormalPosition.top

                    if normal_w <= 0 or normal_h <= 0:
                        normal_w = 1040
                        normal_h = 720

                    if win_w > 0:
                        computed_ratio = (pt.x - rect_win.left) / float(win_w)
                    else:
                        computed_ratio = 0.5

                    # Prevent cursor from landing on top of right controls (130px)
                    max_ratio = max(0.5, (normal_w - 130.0) / float(normal_w))
                    computed_ratio = max(0.05, min(max_ratio, computed_ratio))

                    new_x = int(pt.x - (normal_w * computed_ratio))
                    new_y = int(pt.y - 15)

                    wp.rcNormalPosition.left = new_x
                    wp.rcNormalPosition.top = new_y
                    wp.rcNormalPosition.right = new_x + normal_w
                    wp.rcNormalPosition.bottom = new_y + normal_h
                    user32.SetWindowPlacement(self.hwnd, ctypes.byref(wp))

                lparam = ctypes.c_int32(((pt.y & 0xFFFF) << 16) | (pt.x & 0xFFFF)).value
                user32.ReleaseCapture()
                user32.PostMessageW(self.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, lparam)
            except Exception as e:
                print("start_window_drag error:", e)

    def toggle_maximize(self):
        """Toggle maximize and restore with genuine native DWM animation"""
        if self.hwnd:
            try:
                if user32.IsZoomed(self.hwnd):
                    user32.ShowWindow(self.hwnd, SW_RESTORE)
                    return False
                else:
                    user32.ShowWindow(self.hwnd, SW_MAXIMIZE)
                    return True
            except Exception as e:
                print("toggle_maximize error:", e)
        if self._window:
            if getattr(self._window, 'maximized', False):
                self._window.restore()
                return False
            else:
                self._window.maximize()
                return True
        return False

    def is_maximized(self):
        """Check if native window is currently maximized (Zoomed)"""
        if self.hwnd:
            try:
                return bool(user32.IsZoomed(self.hwnd))
            except Exception:
                pass
        if self._window:
            return bool(getattr(self._window, 'maximized', False))
        return False

    def restore_and_drag(self, ratio_x=0.5):
        """Restore and drag fallback delegating to start_window_drag"""
        self.start_window_drag()

    def get_work_area(self):
        """Get work area of the monitor where window currently resides [x, y, w, h] in physical pixels"""
        if self.hwnd:
            try:
                mi = MONITORINFO()
                mi.cbSize = ctypes.sizeof(MONITORINFO)
                hMon = user32.MonitorFromWindow(self.hwnd, 2)  # MONITOR_DEFAULTTONEAREST
                if user32.GetMonitorInfoW(hMon, ctypes.byref(mi)):
                    w = mi.rcWork.right - mi.rcWork.left
                    h = mi.rcWork.bottom - mi.rcWork.top
                    return [mi.rcWork.left, mi.rcWork.top, w, h]
            except Exception as e:
                print("get_work_area error:", e)
        return [0, 0, 1920, 1040]

    def get_window_bounds(self):
        """Retrieve current native window coordinates [x, y, width, height]"""
        if self.hwnd:
            try:
                rect = wintypes.RECT()
                user32.GetWindowRect(self.hwnd, ctypes.byref(rect))
                return [rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top]
            except Exception as e:
                print("get_window_bounds error:", e)
        if self._window:
            try:
                return [self._window.x, self._window.y, self._window.width, self._window.height]
            except Exception:
                pass
        return [0, 0, 1040, 720]

    def get_normal_dimensions(self):
        """Retrieve pre-maximized normal dimensions from WINDOWPLACEMENT"""
        if self.hwnd:
            try:
                wp = WINDOWPLACEMENT()
                wp.length = ctypes.sizeof(WINDOWPLACEMENT)
                user32.GetWindowPlacement(self.hwnd, ctypes.byref(wp))
                w = wp.rcNormalPosition.right - wp.rcNormalPosition.left
                h = wp.rcNormalPosition.bottom - wp.rcNormalPosition.top
                if w > 100 and h > 100:
                    return [w, h]
            except Exception:
                pass
        return [1040, 720]

    def move_window(self, x, y):
        """Atomically move window position with SWP_NOSIZE for ultra-fast, smooth dragging"""
        if self.hwnd:
            try:
                if user32.IsZoomed(self.hwnd):
                    user32.ShowWindow(self.hwnd, SW_RESTORE)
                user32.SetWindowPos(
                    self.hwnd, 0,
                    int(x), int(y), 0, 0,
                    0x0615  # SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING
                )
            except Exception as e:
                print("move_window error:", e)

    def set_window_bounds(self, x, y, w, h):
        """Atomically reposition and resize window with zero lag or flicker"""
        if self.hwnd:
            try:
                if user32.IsZoomed(self.hwnd):
                    if dwmapi:
                        try:
                            dwmapi.DwmSetWindowAttribute(self.hwnd, 3, ctypes.byref(ctypes.c_int(1)), 4)
                        except Exception:
                            pass
                    user32.ShowWindow(self.hwnd, SW_RESTORE)
                    if dwmapi:
                        try:
                            dwmapi.DwmSetWindowAttribute(self.hwnd, 3, ctypes.byref(ctypes.c_int(0)), 4)
                        except Exception:
                            pass
                user32.SetWindowPos(
                    self.hwnd, 0,
                    int(x), int(y), int(w), int(h),
                    0x0614  # SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING
                )
            except Exception as e:
                print("set_window_bounds error:", e)

    def resize_window(self, w, h):
        """Resize window smoothly from border handles"""
        if self._window:
            try:
                self._window.resize(int(w), int(h))
            except Exception:
                pass

    def save_auto_open(self, is_checked):
        """Save auto_open setting immediately upon change"""
        self.auto_open_state = bool(is_checked)
        save_auto_open_to_file(self.auto_open_state)
        return True

    def close_window(self, auto_open=None):
        if auto_open is not None:
            self.auto_open_state = bool(auto_open)
            save_auto_open_to_file(self.auto_open_state)
        elif self.auto_open_state is not None:
            save_auto_open_to_file(self.auto_open_state)
        if self._window:
            self._window.destroy()

    def minimize_window(self):
        """Native Windows minimize with DWM animation"""
        if self.hwnd:
            try:
                user32.ShowWindow(self.hwnd, SW_MINIMIZE)
                return
            except Exception as e:
                print("minimize_window error:", e)
        if self._window:
            self._window.minimize()

    def select_file(self):
        if not self._window:
            return ""
        try:
            files = self._window.create_file_dialog(
                dialog_type=webview.FileDialog.OPEN,
                file_types=('FortiGate Config (*.conf;*.txt)', 'All files (*.*)')
            )
            return files[0] if files else ""
        except Exception as e:
            print("select_file error:", e)
            return ""

    def select_folder(self):
        if not self._window:
            return ""
        try:
            folders = self._window.create_file_dialog(
                dialog_type=webview.FileDialog.FOLDER
            )
            return folders[0] if folders else ""
        except Exception as e:
            print("select_folder error:", e)
            return ""

    def start_conversion(self, conf_file, out_dir, auto_open):
        if self.is_running:
            return
        self.is_running = True

        if not out_dir:
            out_dir = os.path.dirname(conf_file) or "."

        threading.Thread(
            target=self._worker,
            args=(conf_file, out_dir, auto_open),
            daemon=True
        ).start()

    def _worker(self, conf_file, out_dir, auto_open):
        try:
            def log_fn(msg):
                self.log(msg)

            def status_fn(status_text, pct=None):
                self.set_status(status_text, "running")
                if pct is not None:
                    self.set_progress(pct)

            target_dir = self._run_converter_callback(conf_file, out_dir, log_fn, status_fn)
            self.last_target_dir = target_dir

            if self._window:
                escaped_msg = json.dumps("Export finished successfully!")
                self._window.evaluate_js(f"onExportFinished(true, {escaped_msg});")

            if auto_open and target_dir and os.path.isdir(target_dir):
                self.open_result_folder()

        except Exception as ex:
            import traceback
            err_type = type(ex).__name__
            err_msg = str(ex)
            tb_lines = traceback.format_exc()
            self.log(f"\n[ERROR] {err_type}: {err_msg}")
            self.log(f"[DEBUG] Error details & Traceback:\n{tb_lines.strip()}")
            if self._window:
                escaped = json.dumps(f"{err_type}: {err_msg}")
                self._window.evaluate_js(f"onExportFinished(false, {escaped});")
        finally:
            self.is_running = False

    def log(self, msg):
        if self._window:
            js_code = f"appendLog({json.dumps(str(msg))});"
            self._window.evaluate_js(js_code)

    def set_status(self, text, state="running"):
        if self._window:
            js_code = f"setStatus({json.dumps(str(text))}, {json.dumps(state)});"
            self._window.evaluate_js(js_code)

    def set_progress(self, pct):
        if self._window:
            js_code = f"setProgress({int(pct)});"
            self._window.evaluate_js(js_code)

    def open_result_folder(self):
        if self.last_target_dir and os.path.isdir(self.last_target_dir):
            if sys.platform == 'win32':
                os.startfile(self.last_target_dir)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', self.last_target_dir])
            else:
                subprocess.Popen(['xdg-open', self.last_target_dir])

    def open_external_link(self, url):
        webbrowser.open(url)


def launch_gui(converter_callback):
    """
    Launch PyWebView modern Dark Cyber-Glassmorphism GUI.
    """
    if not HAS_WEBVIEW:
        raise RuntimeError("pywebview is not installed. Please run: pip install pywebview")

    # Set explicit AppUserModelID for Windows Taskbar icon grouping and display
    if sys.platform == 'win32':
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                'fortinet.fortigate.policytoexcel.exporter.2.1'
            )
        except Exception:
            pass

    icon_path = _get_resource_path('fortinet.ico')

    saved_cfg = load_app_config()
    saved_theme = saved_cfg.get("theme", "antigravity-dark")
    saved_auto_open = saved_cfg.get("auto_open", True)
    theme_info = VALID_THEMES.get(saved_theme, VALID_THEMES["antigravity-dark"])

    api = WebviewApi(converter_callback, initial_theme=saved_theme, initial_auto_open=saved_auto_open)
    html_content = get_html_template(saved_theme, saved_auto_open)

    # Calculate center coordinates on the active monitor where program is launched
    start_x, start_y = None, None
    if sys.platform == 'win32' and user32:
        try:
            pt = wintypes.POINT()
            user32.GetCursorPos(ctypes.byref(pt))
            hMon = user32.MonitorFromPoint(pt, 2)  # MONITOR_DEFAULTTONEAREST
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            if user32.GetMonitorInfoW(hMon, ctypes.byref(mi)):
                work_w = mi.rcWork.right - mi.rcWork.left
                work_h = mi.rcWork.bottom - mi.rcWork.top
                start_x = mi.rcWork.left + max(0, (work_w - 1040) // 2)
                start_y = mi.rcWork.top + max(0, (work_h - 720) // 2)
        except Exception:
            pass

    window_kwargs = dict(
        title="FORTIGATE POLICY EXPORTER v2.1",
        html=html_content,
        js_api=api,
        width=1040,
        height=720,
        min_size=(920, 650),
        frameless=True,
        transparent=False,
        easy_drag=False,
        shadow=False,
        resizable=True,
        background_color=theme_info["bg_color"],
    )
    if start_x is not None and start_y is not None:
        window_kwargs['x'] = start_x
        window_kwargs['y'] = start_y

    window = webview.create_window(**window_kwargs)
    api.set_window(window)

    def on_window_shown():
        """Triggered immediately when the native window is mapped and displayed"""
        form = getattr(window, 'native', None)
        if not form:
            return
        try:
            hwnd = form.Handle.ToInt32() if hasattr(form.Handle, 'ToInt32') else int(form.Handle)
            api.set_hwnd(hwnd)

            # 1. Enable WS_CAPTION + WS_THICKFRAME for DWM animations + standard window behavior
            try:
                style = user32.GetWindowLongW(hwnd, GWL_STYLE)
                new_style = style | WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU
                user32.SetWindowLongW(hwnd, GWL_STYLE, new_style)
            except Exception as e:
                print("SetWindowLongW error:", e)

            # 2. Hook window procedure using pure x64 machine code thunk to remove OS titlebar via WM_NCCALCSIZE
            try:
                old_proc = user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC)
                if old_proc:
                    pCallWindowProcW = ctypes.cast(user32.CallWindowProcW, ctypes.c_void_p).value
                    thunk_bytes = _build_nccalcsize_thunk(old_proc, pCallWindowProcW)
                    mem = kernel32.VirtualAlloc(None, len(thunk_bytes), 0x1000 | 0x2000, 0x40)
                    if mem:
                        ctypes.memmove(mem, thunk_bytes, len(thunk_bytes))
                        _allocated_thunk_memory.append(mem)
                        user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, mem)
            except Exception as e:
                print("Native thunk setup error:", e)

            # 3. Enable Windows 11 Native Rounded Corners (DWMWA_WINDOW_CORNER_PREFERENCE = 33, DWMWCP_ROUND = 2)
            try:
                DWMWA_WINDOW_CORNER_PREFERENCE = 33
                DWMWCP_ROUND = 2
                dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ctypes.byref(ctypes.c_int(DWMWCP_ROUND)), 4)
                # Apply initial DWM dark mode attribute
                dark_flag = ctypes.c_int(1 if api.is_dark_theme else 0)
                dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(dark_flag), ctypes.sizeof(dark_flag))
            except Exception:
                pass

            # 4. Set Windows Theme Accent Border on focus, subtle dark border on blur
            api.set_window_active(True)
            try:
                form.Activated += lambda sender, e: api.set_window_active(True)
                form.Deactivate += lambda sender, e: api.set_window_active(False)
            except Exception:
                pass

            # 5. Enable DWM transition animations
            try:
                dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_TRANSITIONS_FORCEDISABLED, ctypes.byref(ctypes.c_int(0)), 4)
            except Exception:
                pass

            # 6. Apply frame change and atomically place in exact center of launched monitor
            try:
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                win_w = rect.right - rect.left
                win_h = rect.bottom - rect.top
                pt = wintypes.POINT()
                user32.GetCursorPos(ctypes.byref(pt))
                hMon = user32.MonitorFromPoint(pt, 2)  # MONITOR_DEFAULTTONEAREST
                mi = MONITORINFO()
                mi.cbSize = ctypes.sizeof(MONITORINFO)
                if user32.GetMonitorInfoW(hMon, ctypes.byref(mi)):
                    work_w = mi.rcWork.right - mi.rcWork.left
                    work_h = mi.rcWork.bottom - mi.rcWork.top
                    cx = mi.rcWork.left + max(0, (work_w - win_w) // 2)
                    cy = mi.rcWork.top + max(0, (work_h - win_h) // 2)
                    user32.SetWindowPos(
                        hwnd, 0, cx, cy, 0, 0,
                        SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
                    )
                else:
                    user32.SetWindowPos(
                        hwnd, 0, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
                    )
            except Exception as e:
                print("SetWindowPos centering error:", e)

            # 7. Ensure Taskbar icon is explicitly visible
            form.ShowInTaskbar = True

            # 8. Set WinForms Form.Icon
            if icon_path and os.path.isfile(icon_path):
                try:
                    import System.Drawing
                    form.Icon = System.Drawing.Icon(icon_path)
                except Exception:
                    pass

                # 9. Apply Win32 WM_SETICON for Taskbar & Alt+Tab icon
                try:
                    IMAGE_ICON = 1
                    LR_LOADFROMFILE = 0x00000010
                    LR_DEFAULTSIZE = 0x00000040
                    h_icon_big = user32.LoadImageW(None, icon_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
                    if h_icon_big:
                        WM_SETICON = 0x80
                        ICON_SMALL = 0
                        ICON_BIG = 1
                        user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_icon_big)
                        user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_icon_big)
                except Exception:
                    pass

            # 10. Hook FormClosing on WinForms
            try:
                def on_form_closing(sender, e):
                    if api.auto_open_state is not None:
                        save_auto_open_to_file(api.auto_open_state)
                form.FormClosing += on_form_closing
            except Exception:
                pass

        except Exception as e:
            print("on_window_shown error:", e)

    def on_window_closing(*args, **kwargs):
        """Ensure preferences are persisted when window closes via any method"""
        if api.auto_open_state is not None:
            save_auto_open_to_file(api.auto_open_state)

    window.events.closing += on_window_closing

    if sys.platform == 'win32':
        window.events.shown += on_window_shown

    webview.start(debug=False)
