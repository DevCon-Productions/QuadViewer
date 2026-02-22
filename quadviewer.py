"""
QuadViewer - Quad-Screen TV Streaming Launcher

Opens a GUI to assign TV channels to four screen quadrants,
then launches borderless Chrome windows for each.
"""

import base64
import ctypes
import ctypes.wintypes
from datetime import datetime, timedelta, timezone
import http.client
import json
import os
import socket
import struct
import subprocess
import sys
import threading
import time
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import urllib.request
import urllib.error

from PIL import Image, ImageTk
import ttkbootstrap as tbs
from ttkbootstrap.constants import *

# ---------------------------------------------------------------------------
# Paths & constants
# ---------------------------------------------------------------------------
# When frozen by PyInstaller, data files are in sys._MEIPASS/_internal;
# otherwise use the script's own directory.
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
    _DATA_DIR = sys._MEIPASS
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    _DATA_DIR = SCRIPT_DIR
# User-writable files live next to the exe (or script); read-only bundled
# data (logos, icon, default configs) lives in _DATA_DIR.
CHANNELS_FILE = os.path.join(SCRIPT_DIR, "channels.json")
ASSIGNMENTS_FILE = os.path.join(SCRIPT_DIR, "assignments.json")
PRESETS_FILE = os.path.join(SCRIPT_DIR, "presets.json")
SETTINGS_FILE = os.path.join(SCRIPT_DIR, "settings.json")
PROFILES_DIR = os.path.join(SCRIPT_DIR, "profiles")
LOGOS_DIR = os.path.join(_DATA_DIR, "logos")
ICO_PATH = os.path.join(_DATA_DIR, "quadviewer.ico")
SPLASH_PATH = os.path.join(_DATA_DIR, "DCPLogo.png")

# Logo display sizes (pixels)
LOGO_SMALL = 24   # channel list
LOGO_LARGE = 48   # quadrant frames

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

# Stable directory name per quadrant so Chrome cookies/state persist
QUAD_PROFILE_NAMES = {
    "Upper Left": "quad_ul",
    "Upper Right": "quad_ur",
    "Lower Left": "quad_ll",
    "Lower Right": "quad_lr",
}

# Base port for Chrome remote debugging (each quadrant gets base + offset)
CDP_BASE_PORT = 19220
CDP_PORT_OFFSETS = {
    "Upper Left": 0,
    "Upper Right": 1,
    "Lower Left": 2,
    "Lower Right": 3,
}

# Profile name to auto-select on Fubo / Spectrum
PROFILE_NAME = "Devon"


def _copy_default_data():
    """On first run from a frozen build, copy bundled data files to SCRIPT_DIR."""
    if _DATA_DIR == SCRIPT_DIR:
        return
    for fname in ("channels.json", "presets.json"):
        dest = os.path.join(SCRIPT_DIR, fname)
        if not os.path.isfile(dest):
            src = os.path.join(_DATA_DIR, fname)
            if os.path.isfile(src):
                shutil.copy2(src, dest)

_copy_default_data()


def _set_icon(window):
    """Set the QuadViewer icon on a Toplevel window."""
    if os.path.isfile(ICO_PATH):
        try:
            window.iconbitmap(ICO_PATH)
        except Exception:
            pass


def find_chrome():
    for path in CHROME_PATHS:
        if os.path.isfile(path):
            return path
    return None


def load_channels():
    if not os.path.isfile(CHANNELS_FILE):
        return []
    with open(CHANNELS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_channels(channels):
    with open(CHANNELS_FILE, "w", encoding="utf-8") as f:
        json.dump(channels, f, indent=2, ensure_ascii=False)


def load_assignments():
    if not os.path.isfile(ASSIGNMENTS_FILE):
        return {}
    with open(ASSIGNMENTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_assignments(assignments):
    """Save quadrant->channel mapping. Stores channel index for each quadrant."""
    with open(ASSIGNMENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(assignments, f, indent=2, ensure_ascii=False)


def load_presets():
    if not os.path.isfile(PRESETS_FILE):
        return []
    with open(PRESETS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_presets(presets):
    with open(PRESETS_FILE, "w", encoding="utf-8") as f:
        json.dump(presets, f, indent=2, ensure_ascii=False)


def load_settings():
    if not os.path.isfile(SETTINGS_FILE):
        return {}
    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)


DEFAULT_THEME = "darkly"


# ---------------------------------------------------------------------------
# Quadrant geometry helpers
# ---------------------------------------------------------------------------
QUADRANTS = {
    "Upper Left":  (0, 0),
    "Upper Right": (1, 0),
    "Lower Left":  (0, 1),
    "Lower Right": (1, 1),
}


# Windows 11 invisible border (DWM shadow) adds ~7-8px on each side.
# Overlap windows by this amount so borders disappear into each other.
WIN_BORDER = 8


def _is_taskbar_autohide():
    """Check if the Windows taskbar is set to auto-hide."""
    class APPBARDATA(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.wintypes.DWORD),
            ("hWnd", ctypes.wintypes.HWND),
            ("uCallbackMessage", ctypes.wintypes.UINT),
            ("uEdge", ctypes.wintypes.UINT),
            ("rc", ctypes.wintypes.RECT),
            ("lParam", ctypes.wintypes.LPARAM),
        ]
    ABM_GETSTATE = 0x00000004
    ABS_AUTOHIDE = 0x0000001
    abd = APPBARDATA()
    abd.cbSize = ctypes.sizeof(APPBARDATA)
    state = ctypes.windll.shell32.SHAppBarMessage(ABM_GETSTATE, ctypes.byref(abd))
    return bool(state & ABS_AUTOHIDE)


def get_work_area():
    """Get usable screen rectangle. Uses full screen if taskbar is auto-hidden."""
    if _is_taskbar_autohide():
        # Taskbar auto-hides: use full screen dimensions
        w = ctypes.windll.user32.GetSystemMetrics(0)   # SM_CXSCREEN
        h = ctypes.windll.user32.GetSystemMetrics(1)   # SM_CYSCREEN
        return 0, 0, w, h
    # Normal taskbar: use work area (excludes taskbar)
    rect = ctypes.wintypes.RECT()
    ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top


def get_quadrant_rect(quad_name, work_x, work_y, work_w, work_h):
    col, row = QUADRANTS[quad_name]
    half_w = work_w // 2
    half_h = work_h // 2
    x = work_x + col * half_w - WIN_BORDER
    y = work_y + row * half_h - WIN_BORDER
    w = half_w + 2 * WIN_BORDER
    h = half_h + 2 * WIN_BORDER
    return x, y, w, h


def get_smart_rects(active_quads, work_x, work_y, work_w, work_h):
    """Compute window rects based on how many quadrants are active.

    1 window  -> full screen
    2 windows -> side-by-side or stacked based on quadrant positions
    3-4       -> standard quadrant layout
    """
    names = list(active_quads)
    n = len(names)

    if n == 1:
        # Single window: full screen
        return {names[0]: (
            work_x - WIN_BORDER,
            work_y - WIN_BORDER,
            work_w + 2 * WIN_BORDER,
            work_h + 2 * WIN_BORDER,
        )}

    if n == 2:
        cols = [QUADRANTS[q][0] for q in names]
        rows = [QUADRANTS[q][1] for q in names]

        if rows[0] == rows[1]:
            # Same row -> side by side, full height each
            rects = {}
            for q in names:
                col = QUADRANTS[q][0]
                x = work_x + col * (work_w // 2) - WIN_BORDER
                y = work_y - WIN_BORDER
                w = work_w // 2 + 2 * WIN_BORDER
                h = work_h + 2 * WIN_BORDER
                rects[q] = (x, y, w, h)
            return rects

        if cols[0] == cols[1]:
            # Same column -> stacked, full width each
            rects = {}
            for q in names:
                row = QUADRANTS[q][1]
                x = work_x - WIN_BORDER
                y = work_y + row * (work_h // 2) - WIN_BORDER
                w = work_w + 2 * WIN_BORDER
                h = work_h // 2 + 2 * WIN_BORDER
                rects[q] = (x, y, w, h)
            return rects

        # Diagonal -> side by side, ordered left/right by column
        rects = {}
        sorted_names = sorted(names, key=lambda q: QUADRANTS[q][0])
        for i, q in enumerate(sorted_names):
            x = work_x + i * (work_w // 2) - WIN_BORDER
            y = work_y - WIN_BORDER
            w = work_w // 2 + 2 * WIN_BORDER
            h = work_h + 2 * WIN_BORDER
            rects[q] = (x, y, w, h)
        return rects

    # 3 or 4 windows: standard quadrant layout
    return {q: get_quadrant_rect(q, work_x, work_y, work_w, work_h) for q in names}


# ---------------------------------------------------------------------------
# Chrome DevTools Protocol — minimal WebSocket client (stdlib only)
# ---------------------------------------------------------------------------

def _ws_send_text(sock, text):
    """Send a WebSocket text frame (client-masked)."""
    payload = text.encode("utf-8")
    length = len(payload)
    mask_key = os.urandom(4)
    header = bytearray()
    header.append(0x81)                       # FIN + text opcode
    if length < 126:
        header.append(0x80 | length)          # MASK bit + length
    elif length < 65536:
        header.append(0x80 | 126)
        header.extend(struct.pack(">H", length))
    else:
        header.append(0x80 | 127)
        header.extend(struct.pack(">Q", length))
    header.extend(mask_key)
    masked = bytearray(payload)
    for i in range(len(masked)):
        masked[i] ^= mask_key[i % 4]
    sock.sendall(header + masked)


def _ws_recv(sock, timeout=10):
    """Receive one WebSocket frame. Returns decoded text or None."""
    sock.settimeout(timeout)
    try:
        header = b""
        while len(header) < 2:
            header += sock.recv(2 - len(header))
        length = header[1] & 0x7F
        if length == 126:
            raw = b""
            while len(raw) < 2:
                raw += sock.recv(2 - len(raw))
            length = struct.unpack(">H", raw)[0]
        elif length == 127:
            raw = b""
            while len(raw) < 8:
                raw += sock.recv(8 - len(raw))
            length = struct.unpack(">Q", raw)[0]
        data = b""
        while len(data) < length:
            data += sock.recv(length - len(data))
        return data.decode("utf-8", errors="replace")
    except Exception:
        return None


def cdp_send(port, method, params, retries=15, delay=2):
    """Connect to Chrome's DevTools Protocol and send an arbitrary command."""
    for attempt in range(retries):
        try:
            conn = http.client.HTTPConnection("127.0.0.1", port, timeout=5)
            conn.request("GET", "/json/list")
            resp = conn.getresponse()
            targets = json.loads(resp.read())
            conn.close()

            page = None
            for t in targets:
                if t.get("type") == "page":
                    page = t
                    break
            if page is None:
                time.sleep(delay)
                continue

            ws_url = page.get("webSocketDebuggerUrl", "")
            if not ws_url:
                time.sleep(delay)
                continue

            stripped = ws_url.replace("ws://", "")
            slash = stripped.find("/")
            host_port = stripped[:slash] if slash != -1 else stripped
            path = stripped[slash:] if slash != -1 else "/"
            ws_host, ws_port = host_port.split(":")

            sock = socket.create_connection((ws_host, int(ws_port)), timeout=10)
            key = base64.b64encode(os.urandom(16)).decode()
            handshake = (
                f"GET {path} HTTP/1.1\r\n"
                f"Host: {host_port}\r\n"
                f"Upgrade: websocket\r\n"
                f"Connection: Upgrade\r\n"
                f"Sec-WebSocket-Key: {key}\r\n"
                f"Sec-WebSocket-Version: 13\r\n"
                f"\r\n"
            )
            sock.sendall(handshake.encode())

            buf = b""
            while b"\r\n\r\n" not in buf:
                buf += sock.recv(4096)
            if b"101" not in buf:
                sock.close()
                time.sleep(delay)
                continue

            msg = json.dumps({"id": 1, "method": method, "params": params})
            _ws_send_text(sock, msg)
            _ws_recv(sock, timeout=5)

            sock.close()
            return True

        except Exception:
            time.sleep(delay)

    return False


def cdp_evaluate(port, js_code, retries=15, delay=2, user_gesture=False):
    """Evaluate JavaScript via CDP Runtime.evaluate."""
    params = {"expression": js_code, "awaitPromise": False}
    if user_gesture:
        params["userGesture"] = True
    return cdp_send(port, "Runtime.evaluate", params, retries, delay)


def cdp_press_key(port, key=" ", code="Space", key_code=32):
    """Send a trusted key press via CDP Input.dispatchKeyEvent."""
    params = {
        "type": "keyDown",
        "key": key,
        "code": code,
        "windowsVirtualKeyCode": key_code,
        "nativeVirtualKeyCode": key_code,
    }
    cdp_send(port, "Input.dispatchKeyEvent", params, retries=3, delay=1)
    params["type"] = "keyUp"
    cdp_send(port, "Input.dispatchKeyEvent", params, retries=3, delay=1)


def cdp_mouse_click(port, x, y):
    """Send a trusted mouse click via CDP Input.dispatchMouseEvent."""
    base = {"x": x, "y": y, "button": "left", "clickCount": 1}
    cdp_send(port, "Input.dispatchMouseEvent",
             {**base, "type": "mousePressed"}, retries=3, delay=1)
    cdp_send(port, "Input.dispatchMouseEvent",
             {**base, "type": "mouseReleased"}, retries=3, delay=1)


def cdp_navigate(port, url, retries=5, delay=2):
    """Navigate an existing Chrome tab to a new URL via CDP Page.navigate."""
    return cdp_send(port, "Page.navigate", {"url": url}, retries, delay)


def cdp_set_window_bounds(port, left, top, width, height, retries=5, delay=1):
    """Move/resize the Chrome window via CDP Browser.setWindowBounds."""
    for attempt in range(retries):
        try:
            conn = http.client.HTTPConnection("127.0.0.1", port, timeout=5)
            conn.request("GET", "/json/list")
            resp = conn.getresponse()
            targets = json.loads(resp.read())
            conn.close()

            page = None
            for t in targets:
                if t.get("type") == "page":
                    page = t
                    break
            if page is None:
                time.sleep(delay)
                continue

            target_id = page.get("id", "")
            ws_url = page.get("webSocketDebuggerUrl", "")
            if not ws_url or not target_id:
                time.sleep(delay)
                continue

            stripped = ws_url.replace("ws://", "")
            slash = stripped.find("/")
            host_port = stripped[:slash] if slash != -1 else stripped
            path = stripped[slash:] if slash != -1 else "/"
            ws_host, ws_port = host_port.split(":")

            sock = socket.create_connection((ws_host, int(ws_port)), timeout=10)
            key = base64.b64encode(os.urandom(16)).decode()
            handshake = (
                f"GET {path} HTTP/1.1\r\n"
                f"Host: {host_port}\r\n"
                f"Upgrade: websocket\r\n"
                f"Connection: Upgrade\r\n"
                f"Sec-WebSocket-Key: {key}\r\n"
                f"Sec-WebSocket-Version: 13\r\n"
                f"\r\n"
            )
            sock.sendall(handshake.encode())

            buf = b""
            while b"\r\n\r\n" not in buf:
                buf += sock.recv(4096)
            if b"101" not in buf:
                sock.close()
                time.sleep(delay)
                continue

            # Get windowId for this target
            msg = json.dumps({
                "id": 1, "method": "Browser.getWindowForTarget",
                "params": {"targetId": target_id},
            })
            _ws_send_text(sock, msg)
            resp_text = _ws_recv(sock, timeout=5)
            if not resp_text:
                sock.close()
                time.sleep(delay)
                continue

            resp_data = json.loads(resp_text)
            window_id = resp_data.get("result", {}).get("windowId")
            if window_id is None:
                sock.close()
                time.sleep(delay)
                continue

            # Set new window bounds
            msg = json.dumps({
                "id": 2, "method": "Browser.setWindowBounds",
                "params": {
                    "windowId": window_id,
                    "bounds": {
                        "left": left, "top": top,
                        "width": width, "height": height,
                        "windowState": "normal",
                    },
                },
            })
            _ws_send_text(sock, msg)
            _ws_recv(sock, timeout=5)

            sock.close()
            return True

        except Exception:
            time.sleep(delay)

    return False


def bring_os_window_to_front(pid):
    """Bring a Chrome window to the OS foreground by process ID.

    Uses SetWindowPos with HWND_TOPMOST/NOTOPMOST which is more reliable
    than SetForegroundWindow (which Windows blocks from background threads).
    """
    try:
        user32 = ctypes.windll.user32
        WNDENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
        )

        found_hwnd = ctypes.wintypes.HWND(0)
        target_pid = ctypes.wintypes.DWORD()

        def enum_cb(hwnd, _):
            nonlocal found_hwnd
            # Only consider visible windows
            if not user32.IsWindowVisible(hwnd):
                return True
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(target_pid))
            if target_pid.value == pid:
                found_hwnd = hwnd
                return False  # stop enumeration
            return True

        user32.EnumWindows(WNDENUMPROC(enum_cb), 0)

        if found_hwnd:
            SW_RESTORE = 9
            if user32.IsIconic(found_hwnd):
                user32.ShowWindow(found_hwnd, SW_RESTORE)
            # SetWindowPos TOPMOST then NOTOPMOST — reliably brings to front
            # without making the window permanently always-on-top
            HWND_TOPMOST = ctypes.wintypes.HWND(-1)
            HWND_NOTOPMOST = ctypes.wintypes.HWND(-2)
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            flags = SWP_NOMOVE | SWP_NOSIZE
            user32.SetWindowPos(found_hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
            user32.SetWindowPos(found_hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
            user32.SetForegroundWindow(found_hwnd)
            return True

    except Exception:
        pass
    return False


# The JavaScript to inject into each Chrome page.
INJECTED_JS = r"""
(function() {
  // Guard against double-injection (we re-inject multiple times to survive navigations)
  if (window.__qv_injected) return;
  window.__qv_injected = true;

  const PROFILE = "__PROFILE__";
  const isFubo = location.hostname.includes("fubo.tv");
  const isSpectrum = location.hostname.includes("spectrum.net");
  const isYouTube = location.hostname.includes("youtube.com");
  const isTwitch = location.hostname.includes("twitch.tv");
  let profileDone = false;
  let overlayDone = false;
  let attempts = 0;

  // Visual indicator that injection worked
  const origTitle = document.title;
  document.title = "[QV] " + origTitle;
  setTimeout(function() { document.title = origTitle; }, 5000);

  // --- Helper: send a key event to active element ---
  function sendKey(key, code, keyCode) {
    var t = document.activeElement || document.body;
    t.dispatchEvent(new KeyboardEvent("keydown", {key:key,code:code,keyCode:keyCode,bubbles:true,cancelable:true}));
    t.dispatchEvent(new KeyboardEvent("keyup", {key:key,code:code,keyCode:keyCode,bubbles:true,cancelable:true}));
  }

  // --- Fubo profile: press Escape then Space ---
  function fuboSelectProfile() {
    var body = document.body.innerText || "";
    if (!body.includes(PROFILE)) return false;
    if (!body.includes("Who") && !body.includes("profile") && !body.includes("Profile")
        && !body.includes(PROFILE)) return false;

    sendKey("Escape", "Escape", 27);
    setTimeout(function() { sendKey(" ", "Space", 32); }, 300);
    return true;
  }

  // --- Generic profile click ---
  function clickProfileElement() {
    var all = document.querySelectorAll("*");
    var best = null;
    var bestLen = Infinity;
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      if (el.offsetParent === null && el.tagName !== "BODY") continue;
      var t = el.textContent.trim();
      if (t === PROFILE || t.startsWith(PROFILE + "\n") || t.startsWith(PROFILE + " ")) {
        if (t.length < bestLen) { best = el; bestLen = t.length; }
      }
    }
    if (best) {
      best.dispatchEvent(new MouseEvent("mousedown", {bubbles:true}));
      best.dispatchEvent(new MouseEvent("mouseup", {bubbles:true}));
      best.click();
      return true;
    }
    return false;
  }

  // --- Spectrum: dismiss modal and channel guide ---
  function dismissSpectrumOverlay() {
    var dismissed = false;
    // Close modal popup
    var closeBtn = document.querySelector("#modal-close");
    if (closeBtn) {
      closeBtn.dispatchEvent(new MouseEvent("mousedown", {bubbles:true}));
      closeBtn.dispatchEvent(new MouseEvent("mouseup", {bubbles:true}));
      closeBtn.click();
      dismissed = true;
    }
    // Close channel guide sidebar — try common selectors and Escape key
    var guideSelectors = [
      ".guide-close", "[class*='guide'] [class*='close']",
      "[class*='channel-guide'] [class*='close']",
      "[aria-label='Close']", "[aria-label='Close Guide']",
      ".sidebar-close", "[class*='sidebar'] [class*='close']",
    ];
    for (var i = 0; i < guideSelectors.length; i++) {
      var el = document.querySelector(guideSelectors[i]);
      if (el && el.offsetParent !== null) {
        el.click();
        dismissed = true;
        break;
      }
    }
    // Press Escape once to dismiss any overlay/guide (only early on)
    if (attempts <= 10) {
      sendKey("Escape", "Escape", 27);
    }
    return dismissed;
  }

  function handleProfile() {
    // Only attempt profile selection on streaming TV sites
    if (isFubo) {
      if (fuboSelectProfile()) return true;
      return clickProfileElement();
    }
    if (isSpectrum) return clickProfileElement();
    // Skip profile click on other sites (Amazon, Netflix, Twitch, YouTube, etc.)
    return false;
  }

  // --- YouTube: hide everything except the video player ---
  function youtubeCleanup() {
    if (!isYouTube) return false;
    var style = document.getElementById("__qv_yt_style");
    if (style) return true;
    style = document.createElement("style");
    style.id = "__qv_yt_style";
    style.textContent = [
      "ytd-masthead, #masthead, #masthead-container { display:none !important; }",
      "#secondary, #related, #comments, #chat, ytd-live-chat-frame { display:none !important; }",
      "#meta, #info, #below, ytd-watch-metadata { display:none !important; }",
      "#page-manager { margin-top:0 !important; }",
      "ytd-watch-flexy { --ytd-watch-flexy-panel-max-height:0px !important; }",
      "#columns { max-width:100% !important; }",
      "#primary { max-width:100% !important; width:100% !important; }",
      "#player-container-outer, #player-container-inner, #ytd-player, .html5-video-container, video {",
      "  width:100vw !important; height:100vh !important; max-width:100vw !important; max-height:100vh !important;",
      "}",
      "#movie_player { width:100vw !important; height:100vh !important; position:fixed !important; top:0 !important; left:0 !important; z-index:9999 !important; }",
      ".ytp-chrome-bottom { opacity:0; transition:opacity .3s; }",
      "#movie_player:hover .ytp-chrome-bottom { opacity:1; }",
    ].join("\n");
    document.head.appendChild(style);
    // Click the video to dismiss any overlay / start playback
    var vid = document.querySelector("video");
    if (vid) vid.click();
    return true;
  }

  // --- Twitch: hide everything except the video player ---
  function twitchCleanup() {
    if (!isTwitch) return false;
    var style = document.getElementById("__qv_tw_style");
    if (style) return true;
    style = document.createElement("style");
    style.id = "__qv_tw_style";
    style.textContent = [
      "nav, .top-nav, .channel-header, .stream-chat, .chat-shell, .right-column, .side-nav { display:none !important; }",
      "[class*='side-nav'], [class*='channel-info'], [class*='metadata-layout'] { display:none !important; }",
      "[class*='chat-room'], [class*='right-column'], [class*='community-points'] { display:none !important; }",
      "[data-a-target='right-column-chat-bar'] { display:none !important; }",
      ".persistent-player, .video-player, .video-player__container, video {",
      "  width:100vw !important; height:100vh !important; max-width:100vw !important; max-height:100vh !important;",
      "  position:fixed !important; top:0 !important; left:0 !important; z-index:9999 !important;",
      "}",
      "[class*='video-player__overlay'] { opacity:0; transition:opacity .3s; }",
      ".video-player:hover [class*='video-player__overlay'] { opacity:1; }",
    ].join("\n");
    document.head.appendChild(style);
    return true;
  }

  function handleOverlay() {
    if (isSpectrum) return dismissSpectrumOverlay();
    if (isYouTube) return youtubeCleanup();
    if (isTwitch) return twitchCleanup();
    return false;
  }

  // --- Spectrum: block video.pause() entirely ---
  // Live TV should never pause.  Override the pause method so that
  // resize / visibility-change / player-internal logic cannot stop playback.
  // Buffering still works (that fires 'waiting', not pause()).
  if (isSpectrum) {
    var _origPause = HTMLVideoElement.prototype.pause;
    HTMLVideoElement.prototype.pause = function() {
      // no-op — live TV should never pause
      return undefined;
    };
  }

  // --- Universal autoplay: force-play any paused video elements ---
  // Skip Spectrum — its player interprets play() as user interaction and
  // shows the control panel / chromecast banner.  The guard above handles it.
  function ensurePlayback() {
    if (isSpectrum) return;
    var vids = document.querySelectorAll("video");
    for (var i = 0; i < vids.length; i++) {
      if (vids[i].paused && vids[i].readyState >= 2) {
        try { vids[i].play(); } catch(e) {}
      }
    }
  }

  var timer = setInterval(function() {
    attempts++;
    if (!profileDone && handleProfile()) profileDone = true;
    if (!overlayDone && handleOverlay()) overlayDone = true;
    ensurePlayback();
    // For Spectrum, keep retrying overlay dismissal longer (guide can appear late)
    var done = profileDone && (overlayDone || isSpectrum ? attempts >= 30 : overlayDone);
    if (done || attempts >= 60) clearInterval(timer);
  }, 1000);
})();
""".replace("__PROFILE__", PROFILE_NAME)

MUTE_ALL_JS = "document.querySelectorAll('video, audio').forEach(function(el){el.muted=true;});"
RESUME_VIDEO_JS = "document.querySelectorAll('video').forEach(function(v){if(v.paused)try{v.play()}catch(e){}});"

# YouTube's internal player overrides the HTML5 .muted property, so we must
# use the player API (#movie_player.mute()/.unMute()) AND set the element
# property to keep them in sync.
YT_MUTE_JS = (
    "(function(){"
    "var p=document.getElementById('movie_player');"
    "if(p&&p.mute)p.mute();"
    "document.querySelectorAll('video, audio').forEach(function(el){el.muted=true;});"
    "})();"
)
YT_UNMUTE_JS = (
    "(function(){"
    "var p=document.getElementById('movie_player');"
    "if(p&&p.unMute)p.unMute();"
    "document.querySelectorAll('video, audio').forEach(function(el){el.muted=false;});"
    "})();"
)


def inject_js_thread(port, start_muted=False, url=""):
    """Background thread: inject JS multiple times to survive page navigations."""
    is_yt = "youtube.com" in url or "youtu.be" in url
    mute_js = YT_MUTE_JS if is_yt else MUTE_ALL_JS
    # First injection: as soon as Chrome is reachable
    cdp_evaluate(port, INJECTED_JS)
    if start_muted:
        cdp_evaluate(port, mute_js, retries=2, delay=1)
    # Re-inject after delays to catch pages that navigate/redirect
    for wait in (8, 12, 15, 20):
        time.sleep(wait)
        cdp_evaluate(port, INJECTED_JS, retries=3, delay=1)
        if start_muted:
            cdp_evaluate(port, mute_js, retries=1, delay=1)


def unpause_thread(port, win_w, win_h, url=""):
    """Background thread: click center of video multiple times to dismiss play overlays.
    Skips Spectrum — clicking the video area triggers its channel guide sidebar.
    """
    if "spectrum.net" in url:
        return
    cx, cy = win_w // 2, win_h // 2
    for wait in (15, 10, 10, 10):
        time.sleep(wait)
        cdp_mouse_click(port, cx, cy)


# ---------------------------------------------------------------------------
# TV schedule cache (TVGuide backend API)
# ---------------------------------------------------------------------------
TVGUIDE_API_HOST = "backend.tvguide.com"
TVGUIDE_PROVIDER_ID = "9100001138"  # Eastern - National Listings

_schedule_cache = {
    "timestamp": 0,         # last fetch unix time
    "data": {},             # tvguide_name (e.g. "ESPN") -> list of {title, start, end}
    "fetching": False,
}


def _fmt_time(dt):
    """Format a datetime as '3:00 PM' style local time."""
    if dt is None:
        return ""
    return dt.strftime("%I:%M %p").lstrip("0")


def _fetch_schedule():
    """Fetch current TV schedule from TVGuide backend API."""
    now_ts = int(time.time())
    # Re-fetch at most every 10 minutes
    if _schedule_cache["data"] and now_ts - _schedule_cache["timestamp"] < 600:
        return
    if _schedule_cache["fetching"]:
        return
    _schedule_cache["fetching"] = True
    try:
        url = (
            f"https://{TVGUIDE_API_HOST}/tvschedules/tvguide/"
            f"{TVGUIDE_PROVIDER_ID}/web?start={now_ts}&duration=240"
        )
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json",
        })
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        items = raw.get("data", {}).get("items", [])
        by_name = {}
        for item in items:
            ch = item.get("channel", {})
            ch_name = ch.get("name", "")
            if not ch_name:
                continue
            schedules = item.get("programSchedules", [])
            entries = []
            for s in schedules:
                start_dt = datetime.fromtimestamp(s["startTime"]).astimezone()
                end_dt = datetime.fromtimestamp(s["endTime"]).astimezone()
                entries.append({
                    "title": s.get("title", "Unknown"),
                    "start": start_dt,
                    "end": end_dt,
                })
            entries.sort(key=lambda e: e["start"])
            by_name[ch_name] = entries
        _schedule_cache["data"] = by_name
        _schedule_cache["timestamp"] = now_ts
    except Exception:
        pass
    finally:
        _schedule_cache["fetching"] = False


def get_current_show(tvguide_name):
    """Return info about what's currently airing on a channel, or None."""
    if not tvguide_name:
        return None
    # Auto-refresh if cache is stale
    now_ts = int(time.time())
    if now_ts - _schedule_cache["timestamp"] > 600 and not _schedule_cache["fetching"]:
        threading.Thread(target=_fetch_schedule, daemon=True).start()
    if not _schedule_cache["data"]:
        return None
    episodes = _schedule_cache["data"].get(tvguide_name, [])
    if not episodes:
        return None
    now = datetime.now().astimezone()
    current = None
    next_show = None
    for ep in episodes:
        if ep["start"] <= now < ep["end"]:
            current = ep
        elif ep["start"] > now and next_show is None:
            next_show = ep
    if current:
        lines = [f"Now: {current['title']}"]
        lines.append(f"  {_fmt_time(current['start'])} - {_fmt_time(current['end'])}")
        if next_show:
            lines.append(f"Next: {next_show['title']} ({_fmt_time(next_show['start'])})")
        return "\n".join(lines)
    elif next_show:
        return f"Next: {next_show['title']} at {_fmt_time(next_show['start'])}"
    return None


# ---------------------------------------------------------------------------
# Hover Tooltip
# ---------------------------------------------------------------------------
class ToolTip(tk.Toplevel):
    """A floating tooltip window that follows the cursor near a widget."""

    def __init__(self, parent):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self._label = tk.Label(
            self,
            justify=tk.LEFT,
            background="#2d2d2d",
            foreground="#e0e0e0",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
            padx=6,
            pady=4,
        )
        self._label.pack()
        self.withdraw()

    def show(self, text, x, y):
        self._label.config(text=text)
        self.geometry(f"+{x + 16}+{y + 10}")
        self.deiconify()

    def hide(self):
        self.withdraw()


# ---------------------------------------------------------------------------
# Channel Editor Dialog
# ---------------------------------------------------------------------------
class ChannelDialog(tk.Toplevel):
    """Modal dialog for adding or editing a channel."""

    def __init__(self, parent, title="Channel", channel=None):
        super().__init__(parent)
        _set_icon(self)
        self.title(title)
        self.resizable(False, False)
        self.grab_set()
        self.result = None

        pad = {"padx": 8, "pady": 4}

        ttk.Label(self, text="Channel Name:").grid(row=0, column=0, sticky="w", **pad)
        self.name_var = tk.StringVar(value=channel["name"] if channel else "")
        self._name_entry = ttk.Entry(self, textvariable=self.name_var, width=50)
        self._name_entry.grid(row=0, column=1, **pad)

        ttk.Label(self, text="URL:").grid(row=1, column=0, sticky="w", **pad)
        self.url_var = tk.StringVar(value=channel.get("url", "") if channel else "")
        ttk.Entry(self, textvariable=self.url_var, width=50).grid(row=1, column=1, **pad)

        # Logo picker
        ttk.Label(self, text="Logo:").grid(row=2, column=0, sticky="w", **pad)
        logo_frame = ttk.Frame(self)
        logo_frame.grid(row=2, column=1, sticky="w", **pad)
        self.logo_var = tk.StringVar(value=channel.get("logo", "") if channel else "")
        ttk.Entry(logo_frame, textvariable=self.logo_var, width=36).pack(side=tk.LEFT)
        ttk.Button(logo_frame, text="Browse...", command=self._browse_logo).pack(
            side=tk.LEFT, padx=(4, 0)
        )

        # Logo preview
        self._logo_preview_label = ttk.Label(self, text="")
        self._logo_preview_label.grid(row=3, column=1, sticky="w", **pad)
        self._preview_img = None
        self._update_logo_preview()
        self.logo_var.trace_add("write", lambda *_: self._update_logo_preview())

        # TVGuide channel name (optional, for programming guide)
        ttk.Label(self, text="TVGuide Name:").grid(row=4, column=0, sticky="w", **pad)
        self.tvguide_var = tk.StringVar(value=channel.get("tvguide_name", "") if channel else "")
        ttk.Entry(self, textvariable=self.tvguide_var, width=20).grid(row=4, column=1, sticky="w", **pad)

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="OK", command=self._ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=6)

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())

        # Center on parent
        self.transient(parent)
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        w = self.winfo_width()
        h = self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

        # Auto-focus the name field so user can start typing immediately
        self._name_entry.focus_set()
        self._name_entry.icursor(tk.END)

    def _browse_logo(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select Logo Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.ico *.webp"),
                ("All files", "*.*"),
            ],
        )
        if path:
            # Copy to logos dir with a sanitised filename
            os.makedirs(LOGOS_DIR, exist_ok=True)
            fname = os.path.basename(path)
            dest = os.path.join(LOGOS_DIR, fname)
            if os.path.abspath(path) != os.path.abspath(dest):
                shutil.copy2(path, dest)
            self.logo_var.set(fname)

    def _update_logo_preview(self):
        fname = self.logo_var.get().strip()
        logo_path = os.path.join(LOGOS_DIR, fname) if fname else ""
        if logo_path and os.path.isfile(logo_path):
            try:
                img = Image.open(logo_path)
                img.thumbnail((LOGO_LARGE, LOGO_LARGE), Image.LANCZOS)
                self._preview_img = ImageTk.PhotoImage(img)
                self._logo_preview_label.config(image=self._preview_img, text="")
                return
            except Exception:
                pass
        self._preview_img = None
        self._logo_preview_label.config(image="", text="(no logo)" if not fname else "(not found)")

    def _ok(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Missing Name", "Channel name is required.", parent=self)
            return
        result = {
            "name": name,
            "url": self.url_var.get().strip(),
            "logo": self.logo_var.get().strip(),
        }
        tvg = self.tvguide_var.get().strip()
        if tvg:
            result["tvguide_name"] = tvg
        self.result = result
        self.destroy()


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------
class QuadViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DevCon QuadViewer")
        self.root.resizable(True, True)

        self.channels = load_channels()
        self.assignments = {name: None for name in QUADRANTS}
        self.processes = []
        self.active_ports = {}        # quad_name -> CDP debug port
        self.active_pids = {}         # quad_name -> Chrome process ID
        self.audio_quad = "Upper Left"  # currently unmuted quadrant
        self.block_spectrum_iha = tk.BooleanVar(value=True)  # block IHA by default

        self._ctrl9_showing_app = False  # toggle state for Ctrl+9

        # Window maximize state tracking
        self._quad_rects = {}       # quad_name -> (x, y, w, h) original rect
        self._quad_maximized = {}   # quad_name -> bool

        # Presets
        self.presets = load_presets()

        # Logo image caches (keep references so tkinter doesn't GC them)
        self._logo_small = {}   # logo filename -> PhotoImage (24x24)
        self._logo_large = {}   # logo filename -> PhotoImage (48x48)

        # Restore saved quadrant assignments
        self._restore_assignments()

        # Drag state
        self._drag_channel = None
        self._drag_source_idx = None
        self._drag_indicator = None
        self._drag_is_reorder = False

        self._build_gui()
        self._apply_restored_assignments()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Fetch TV schedule in background for hover tooltips
        threading.Thread(target=_fetch_schedule, daemon=True).start()

        # Global hotkeys (Ctrl+1..4) — work even when QuadViewer isn't focused
        self._hotkey_thread = threading.Thread(
            target=self._hotkey_loop, daemon=True
        )
        self._hotkey_thread.start()

    def _restore_assignments(self):
        """Load saved quadrant assignments from disk."""
        saved = load_assignments()
        for quad_name in QUADRANTS:
            idx = saved.get(quad_name)
            if idx is not None and 0 <= idx < len(self.channels):
                self.assignments[quad_name] = self.channels[idx]

    def _update_quad_display(self, quad_name, channel):
        """Update a quadrant's label and logo image."""
        if channel:
            self.quad_labels[quad_name].config(text=channel["name"])
            logo = self._get_logo(channel.get("logo", ""), LOGO_LARGE)
            if logo:
                self.quad_logos[quad_name].config(image=logo)
            else:
                self.quad_logos[quad_name].config(image="")
        else:
            self.quad_labels[quad_name].config(text="(empty)")
            self.quad_logos[quad_name].config(image="")

    def _apply_restored_assignments(self):
        """Update the GUI labels to reflect restored assignments."""
        for quad_name, ch in self.assignments.items():
            if ch is not None:
                self._update_quad_display(quad_name, ch)

    def _save_assignments(self):
        """Persist current quadrant assignments to disk."""
        data = {}
        for quad_name, ch in self.assignments.items():
            if ch is not None and ch in self.channels:
                data[quad_name] = self.channels.index(ch)
        save_assignments(data)

    # ---- Logo helpers -------------------------------------------------------

    def _get_logo(self, logo_filename, size):
        """Return a cached PhotoImage for the given logo file at the given size."""
        if not logo_filename:
            return None
        cache = self._logo_small if size == LOGO_SMALL else self._logo_large
        if logo_filename in cache:
            return cache[logo_filename]
        path = os.path.join(LOGOS_DIR, logo_filename)
        if not os.path.isfile(path):
            return None
        try:
            img = Image.open(path)
            img.thumbnail((size, size), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            cache[logo_filename] = photo
            return photo
        except Exception:
            return None

    def _invalidate_logo_cache(self, logo_filename):
        """Remove a logo from caches so it gets reloaded next time."""
        self._logo_small.pop(logo_filename, None)
        self._logo_large.pop(logo_filename, None)

    # ---- GUI construction --------------------------------------------------

    def _build_gui(self):
        # Menu bar
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Import Channels & Presets...", command=self._import_data)
        file_menu.add_command(label="Export Channels & Presets...", command=self._export_data)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Help", command=self._show_help)
        help_menu.add_command(label="YouTube Tutorial", command=self._open_youtube_tutorial)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self._show_about)
        options_menu = tk.Menu(menubar, tearoff=0)
        options_menu.add_command(label="Preferences...", command=self._show_preferences)
        menubar.add_cascade(label="File", menu=file_menu)
        menubar.add_cascade(label="Options", menu=options_menu)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # ----- Left panel: channel list + management buttons -----
        left_frame = ttk.LabelFrame(main, text="Channels", padding=4)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 3))

        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.channel_tree = ttk.Treeview(
            tree_frame, columns=("name",), show="tree", selectmode="browse"
        )
        self.channel_tree.heading("name", text="Channel", anchor="w")
        self.channel_tree.column("#0", width=50, minwidth=50, stretch=False)  # logo icon
        self.channel_tree.column("name", width=160, anchor="w")

        # Row height to fit logo icons + bold left-aligned header
        tree_style = self.root.style
        tree_style.configure("Treeview", rowheight=LOGO_SMALL + 10)
        tree_style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), anchor="w")
        scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.channel_tree.yview
        )
        self.channel_tree.configure(yscrollcommand=scrollbar.set)
        self.channel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.channel_tree.tag_configure("drop_target", background="#1a3a5c")
        self._drop_line = None  # Canvas line showing insertion point

        self._populate_tree()

        # Channel management buttons
        mgmt_frame = ttk.Frame(left_frame)
        mgmt_frame.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(mgmt_frame, text="Add", command=self._add_channel).pack(
            side=tk.LEFT, padx=2, expand=True, fill=tk.X
        )
        ttk.Button(mgmt_frame, text="Edit", command=self._edit_channel).pack(
            side=tk.LEFT, padx=2, expand=True, fill=tk.X
        )
        ttk.Button(mgmt_frame, text="Delete", command=self._delete_channel).pack(
            side=tk.LEFT, padx=2, expand=True, fill=tk.X
        )

        # Bind drag-and-drop events on the treeview
        self.channel_tree.bind("<ButtonPress-1>", self._drag_start)
        self.channel_tree.bind("<B1-Motion>", self._drag_motion)
        self.channel_tree.bind("<ButtonRelease-1>", self._drag_drop)

        # Tooltip for showing current programming on hover
        self._tooltip = ToolTip(self.root)
        self._tooltip_item = None
        self.channel_tree.bind("<Motion>", self._on_tree_hover)
        self.channel_tree.bind("<Leave>", self._on_tree_leave)

        # ----- Right panel: quadrant grid + controls -----
        right_frame = ttk.Frame(main, padding=4)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(3, 0))

        grid_frame = ttk.Frame(right_frame)
        grid_frame.pack(fill=tk.BOTH, expand=True)

        self.quad_labels = {}
        self.quad_logos = {}    # quad_name -> ttk.Label for logo image
        self.quad_frames = {}
        self.quad_max_btns = {}  # quad_name -> ttk.Button for Max/Shrink
        positions = [
            ("Upper Left", 0, 0),
            ("Upper Right", 0, 1),
            ("Lower Left", 1, 0),
            ("Lower Right", 1, 1),
        ]
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)
        grid_frame.rowconfigure(0, weight=1)
        grid_frame.rowconfigure(1, weight=1)

        for quad_name, row, col in positions:
            frame = ttk.LabelFrame(grid_frame, text=quad_name, padding=8)
            frame.grid(row=row, column=col, padx=4, pady=4, sticky="nsew")
            self.quad_frames[quad_name] = frame

            btn_frame = ttk.Frame(frame)
            btn_frame.pack(side=tk.BOTTOM, pady=(6, 0))

            label = ttk.Label(frame, text="(empty)", anchor="center", width=18)
            label.pack(side=tk.BOTTOM, pady=(0, 2))

            logo_label = ttk.Label(frame, anchor="center")
            logo_label.pack(side=tk.TOP, pady=(2, 0))
            self.quad_logos[quad_name] = logo_label
            self.quad_labels[quad_name] = label
            ttk.Button(
                btn_frame,
                text="Set",
                command=lambda q=quad_name: self._set_quadrant(q),
            ).pack(side=tk.LEFT, padx=2)
            ttk.Button(
                btn_frame,
                text="Clear",
                command=lambda q=quad_name: self._clear_quadrant(q),
            ).pack(side=tk.LEFT, padx=2)
            ttk.Button(
                btn_frame,
                text="Front",
                command=lambda q=quad_name: self._bring_quad_to_front(q),
            ).pack(side=tk.LEFT, padx=2)
            max_btn = ttk.Button(
                btn_frame,
                text="Max",
                command=lambda q=quad_name: self._toggle_maximize(q),
            )
            max_btn.pack(side=tk.LEFT, padx=2)
            self.quad_max_btns[quad_name] = max_btn
            ttk.Button(
                btn_frame,
                text="Switch",
                command=lambda q=quad_name: self._switch_quadrant(q),
            ).pack(side=tk.LEFT, padx=2)
            ttk.Button(
                btn_frame,
                text="Close",
                command=lambda q=quad_name: self._close_quadrant(q),
            ).pack(side=tk.LEFT, padx=2)
            ttk.Button(
                btn_frame,
                text="Open",
                command=lambda q=quad_name: self._open_quadrant(q),
            ).pack(side=tk.LEFT, padx=2)

        # Audio controls
        audio_frame = ttk.LabelFrame(right_frame, text="Audio (Ctrl+1-4, 0=Mute)  |  Max/Shrink (Ctrl+5-8)  |  App/TV (Ctrl+9)", padding=4)
        audio_frame.pack(fill=tk.X, pady=(6, 0))
        audio_map = [
            ("1: Upper Left", "Upper Left"),
            ("2: Upper Right", "Upper Right"),
            ("3: Lower Left", "Lower Left"),
            ("4: Lower Right", "Lower Right"),
        ]
        for label_text, quad_name in audio_map:
            ttk.Button(
                audio_frame, text=label_text,
                command=lambda q=quad_name: self._set_audio_solo(q),
            ).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)

        # Keyboard shortcuts for audio
        self.root.bind("1", lambda e: self._set_audio_solo("Upper Left"))
        self.root.bind("2", lambda e: self._set_audio_solo("Upper Right"))
        self.root.bind("3", lambda e: self._set_audio_solo("Lower Left"))
        self.root.bind("4", lambda e: self._set_audio_solo("Lower Right"))

        # Spectrum IHA toggle
        ttk.Checkbutton(
            right_frame, text="Block Spectrum auto-login (use saved account)",
            variable=self.block_spectrum_iha,
        ).pack(fill=tk.X, pady=(6, 0))

        # Presets
        preset_frame = ttk.LabelFrame(right_frame, text="Presets", padding=4)
        preset_frame.pack(fill=tk.X, pady=(6, 0))

        self._preset_var = tk.StringVar()
        self._preset_combo = ttk.Combobox(
            preset_frame, textvariable=self._preset_var,
            state="readonly", width=20,
        )
        self._preset_combo.pack(side=tk.LEFT, padx=(0, 4), fill=tk.X, expand=True)
        self._refresh_preset_combo()

        ttk.Button(preset_frame, text="Load", command=self._load_preset).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(preset_frame, text="Save", command=self._save_preset).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(preset_frame, text="Overwrite", command=self._overwrite_preset).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(preset_frame, text="Delete", command=self._delete_preset).pack(
            side=tk.LEFT, padx=2
        )

        # Bottom controls (centered)
        controls = ttk.Frame(right_frame, padding=(0, 8, 0, 0))
        controls.pack(pady=(4, 0))

        ttk.Button(controls, text="Execute", command=self._execute).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(controls, text="Clear All", command=self._clear_all_quadrants).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(controls, text="Bring to Front", command=self._bring_to_front).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(controls, text="Close All", command=self._close_all).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(controls, text="Exit", command=self._on_close).pack(
            side=tk.LEFT, padx=4
        )

    def _populate_tree(self):
        for item in self.channel_tree.get_children():
            self.channel_tree.delete(item)
        for ch in self.channels:
            logo = self._get_logo(ch.get("logo", ""), LOGO_SMALL)
            if logo:
                self.channel_tree.insert("", tk.END, values=(ch["name"],), image=logo)
            else:
                self.channel_tree.insert("", tk.END, values=(ch["name"],))

    # ---- Channel hover (programming guide) -----------------------------------

    def _on_tree_hover(self, event):
        """Show a tooltip with current programming when hovering over a channel."""
        item = self.channel_tree.identify_row(event.y)
        if not item:
            self._tooltip.hide()
            self._tooltip_item = None
            return
        if item == self._tooltip_item:
            # Same item — just update position
            self._tooltip.geometry(
                f"+{event.x_root + 16}+{event.y_root + 10}"
            )
            return
        self._tooltip_item = item
        children = list(self.channel_tree.get_children())
        idx = children.index(item)
        if idx < 0 or idx >= len(self.channels):
            self._tooltip.hide()
            return
        ch = self.channels[idx]
        tvg_name = ch.get("tvguide_name")
        info = get_current_show(tvg_name)
        if info:
            self._tooltip.show(info, event.x_root, event.y_root)
        else:
            self._tooltip.hide()

    def _on_tree_leave(self, event):
        self._tooltip.hide()
        self._tooltip_item = None

    # ---- Drag and drop -----------------------------------------------------

    def _drag_start(self, event):
        item = self.channel_tree.identify_row(event.y)
        if not item:
            self._drag_channel = None
            return
        self.channel_tree.selection_set(item)
        children = self.channel_tree.get_children()
        tree_idx = list(children).index(item)
        if 0 <= tree_idx < len(self.channels):
            self._drag_channel = self.channels[tree_idx]
            self._drag_source_idx = tree_idx
        else:
            self._drag_channel = None
            self._drag_source_idx = None
        self._drag_is_reorder = False

    def _drag_motion(self, event):
        if self._drag_channel is None:
            return

        abs_x = self.root.winfo_pointerx()
        abs_y = self.root.winfo_pointery()

        tx = self.channel_tree.winfo_rootx()
        ty = self.channel_tree.winfo_rooty()
        tw = self.channel_tree.winfo_width()
        th = self.channel_tree.winfo_height()
        over_tree = tx <= abs_x <= tx + tw and ty <= abs_y <= ty + th
        self._drag_is_reorder = over_tree

        if self._drag_indicator is None:
            self._drag_indicator = tk.Toplevel(self.root)
            self._drag_indicator.overrideredirect(True)
            self._drag_indicator.attributes("-topmost", True)
            lbl = tk.Label(
                self._drag_indicator,
                text=self._drag_channel["name"],
                bg="#3a3a5c",
                fg="#e0e0e0",
                relief="solid",
                borderwidth=1,
                padx=6,
                pady=2,
                font=("Segoe UI", 9),
            )
            lbl.pack()
            self._drag_indicator.update_idletasks()

        self._drag_indicator.geometry(f"+{abs_x + 14}+{abs_y + 10}")

        # Highlight drop target row when reordering within the list
        self._clear_drop_highlight()
        if over_tree:
            local_y = abs_y - self.channel_tree.winfo_rooty()
            target_item = self.channel_tree.identify_row(local_y)
            if target_item:
                children = list(self.channel_tree.get_children())
                target_idx = children.index(target_item)
                if target_idx != self._drag_source_idx:
                    self.channel_tree.item(target_item, tags=("drop_target",))
                    # Draw insertion line
                    bbox = self.channel_tree.bbox(target_item)
                    if bbox:
                        lx, ly, lw, lh = bbox
                        if target_idx > self._drag_source_idx:
                            line_y = ly + lh  # insert below
                        else:
                            line_y = ly  # insert above
                        if self._drop_line is None:
                            self._drop_line = tk.Toplevel(self.root)
                            self._drop_line.overrideredirect(True)
                            self._drop_line.attributes("-topmost", True)
                            self._drop_line_canvas = tk.Canvas(
                                self._drop_line, height=3, highlightthickness=0,
                                bg="#57a5ff"
                            )
                            self._drop_line_canvas.pack(fill=tk.X)
                        line_abs_x = self.channel_tree.winfo_rootx() + lx
                        line_abs_y = self.channel_tree.winfo_rooty() + line_y - 1
                        self._drop_line.geometry(f"{lw}x3+{line_abs_x}+{line_abs_y}")
                        self._drop_line.deiconify()

        for quad_name, frame in self.quad_frames.items():
            try:
                if not over_tree:
                    fx = frame.winfo_rootx()
                    fy = frame.winfo_rooty()
                    fw = frame.winfo_width()
                    fh = frame.winfo_height()
                    if fx <= abs_x <= fx + fw and fy <= abs_y <= fy + fh:
                        frame.configure(style="Hover.TLabelframe")
                    else:
                        frame.configure(style="TLabelframe")
                else:
                    frame.configure(style="TLabelframe")
            except tk.TclError:
                pass

    def _drag_drop(self, event):
        if self._drag_channel is None:
            self._cleanup_drag()
            return

        abs_x = self.root.winfo_pointerx()
        abs_y = self.root.winfo_pointery()

        if self._drag_is_reorder:
            local_y = abs_y - self.channel_tree.winfo_rooty()
            target_item = self.channel_tree.identify_row(local_y)
            if target_item and self._drag_source_idx is not None:
                children = list(self.channel_tree.get_children())
                target_idx = children.index(target_item)
                if target_idx != self._drag_source_idx:
                    ch = self.channels.pop(self._drag_source_idx)
                    self.channels.insert(target_idx, ch)
                    save_channels(self.channels)
                    self._populate_tree()
                    new_children = self.channel_tree.get_children()
                    if 0 <= target_idx < len(new_children):
                        self.channel_tree.selection_set(new_children[target_idx])
                        self.channel_tree.see(new_children[target_idx])
        else:
            for quad_name, frame in self.quad_frames.items():
                fx = frame.winfo_rootx()
                fy = frame.winfo_rooty()
                fw = frame.winfo_width()
                fh = frame.winfo_height()
                if fx <= abs_x <= fx + fw and fy <= abs_y <= fy + fh:
                    self.assignments[quad_name] = self._drag_channel
                    self._update_quad_display(quad_name, self._drag_channel)
                    self._save_assignments()
                    break

        self._cleanup_drag()

    def _clear_drop_highlight(self):
        """Remove the row highlight and insertion line from the treeview."""
        for item in self.channel_tree.get_children():
            self.channel_tree.item(item, tags=())
        if self._drop_line is not None:
            self._drop_line.withdraw()

    def _cleanup_drag(self):
        self._drag_channel = None
        self._drag_source_idx = None
        self._drag_is_reorder = False
        self._clear_drop_highlight()
        if self._drop_line is not None:
            self._drop_line.destroy()
            self._drop_line = None
        if self._drag_indicator:
            self._drag_indicator.destroy()
            self._drag_indicator = None
        for frame in self.quad_frames.values():
            try:
                frame.configure(style="TLabelframe")
            except tk.TclError:
                pass

    # ---- Channel management ------------------------------------------------

    def _add_channel(self):
        dlg = ChannelDialog(self.root, title="Add Channel")
        self.root.wait_window(dlg)
        if dlg.result:
            self.channels.append(dlg.result)
            save_channels(self.channels)
            self._populate_tree()
            children = self.channel_tree.get_children()
            if children:
                last = children[-1]
                self.channel_tree.selection_set(last)
                self.channel_tree.see(last)

    def _edit_channel(self):
        sel = self.channel_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a channel to edit.")
            return
        children = list(self.channel_tree.get_children())
        idx = children.index(sel[0])
        if idx < 0 or idx >= len(self.channels):
            return
        dlg = ChannelDialog(self.root, title="Edit Channel", channel=self.channels[idx])
        self.root.wait_window(dlg)
        if dlg.result:
            old_name = self.channels[idx]["name"]
            old_logo = self.channels[idx].get("logo", "")
            self.channels[idx] = dlg.result
            save_channels(self.channels)
            # Invalidate logo cache if logo changed
            new_logo = dlg.result.get("logo", "")
            if old_logo:
                self._invalidate_logo_cache(old_logo)
            if new_logo:
                self._invalidate_logo_cache(new_logo)
            self._populate_tree()
            for q, ch in self.assignments.items():
                if ch and ch["name"] == old_name:
                    self.assignments[q] = dlg.result
                    self._update_quad_display(q, dlg.result)
            self._save_assignments()

    def _delete_channel(self):
        sel = self.channel_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a channel to delete.")
            return
        children = list(self.channel_tree.get_children())
        idx = children.index(sel[0])
        if idx < 0 or idx >= len(self.channels):
            return
        name = self.channels[idx]["name"]
        if not messagebox.askyesno("Confirm Delete", f"Delete '{name}'?"):
            return
        removed = self.channels.pop(idx)
        save_channels(self.channels)
        self._populate_tree()
        for q, ch in self.assignments.items():
            if ch is removed:
                self.assignments[q] = None
                self._update_quad_display(q, None)
        self._save_assignments()

    # ---- Assignment logic ---------------------------------------------------

    def _get_selected_channel(self):
        sel = self.channel_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a channel from the list first.")
            return None
        children = list(self.channel_tree.get_children())
        idx = children.index(sel[0])
        if 0 <= idx < len(self.channels):
            return self.channels[idx]
        return None

    def _set_quadrant(self, quad_name):
        channel = self._get_selected_channel()
        if channel is None:
            return
        self.assignments[quad_name] = channel
        self._update_quad_display(quad_name, channel)
        self._save_assignments()

    def _clear_quadrant(self, quad_name):
        self.assignments[quad_name] = None
        self._update_quad_display(quad_name, None)
        self._save_assignments()

    def _clear_all_quadrants(self):
        for quad_name in QUADRANTS:
            self.assignments[quad_name] = None
            self._update_quad_display(quad_name, None)
        self._save_assignments()

    def _switch_quadrant(self, quad_name):
        """Navigate an already-running quadrant's Chrome window to the assigned channel."""
        port = self.active_ports.get(quad_name)
        if not port:
            messagebox.showinfo(
                "Not Running",
                f"{quad_name} has no active window.\n"
                "Use 'Execute' to launch channels first.",
            )
            return
        channel = self.assignments.get(quad_name)
        if not channel:
            messagebox.showwarning(
                "No Channel",
                f"No channel assigned to {quad_name}.\n"
                "Use 'Set' or drag a channel first.",
            )
            return
        url = channel.get("url", "")
        if not url:
            messagebox.showwarning(
                "Missing URL",
                f"No URL set for {channel['name']}.",
            )
            return

        # Determine mute state: should this quadrant be muted?
        start_muted = quad_name != self.audio_quad

        def do_switch():
            cdp_navigate(port, url)
            # Re-inject automation JS after page loads
            inject_js_thread(port, start_muted=start_muted, url=url)

        threading.Thread(target=do_switch, daemon=True).start()

    # ---- Presets -------------------------------------------------------------

    def _refresh_preset_combo(self):
        """Update the preset combobox values from the current presets list."""
        names = [p["name"] for p in self.presets]
        self._preset_combo["values"] = names
        if names and not self._preset_var.get():
            self._preset_var.set(names[0])
        elif self._preset_var.get() not in names:
            self._preset_var.set(names[0] if names else "")

    def _get_selected_preset_idx(self):
        """Return the index of the currently selected preset, or None."""
        name = self._preset_var.get()
        for i, p in enumerate(self.presets):
            if p["name"] == name:
                return i
        return None

    def _channel_by_name(self, name):
        """Find a channel dict by name, or None."""
        for ch in self.channels:
            if ch["name"] == name:
                return ch
        return None

    def _load_preset(self):
        """Apply the selected preset to all four quadrants."""
        idx = self._get_selected_preset_idx()
        if idx is None:
            messagebox.showwarning("No Preset", "Select a preset to load.")
            return
        preset = self.presets[idx]
        assignments = preset.get("assignments", {})
        for quad_name in QUADRANTS:
            ch_name = assignments.get(quad_name)
            ch = self._channel_by_name(ch_name) if ch_name else None
            self.assignments[quad_name] = ch
            self._update_quad_display(quad_name, ch)
        self._save_assignments()

    def _save_preset(self):
        """Save current quadrant assignments as a new preset."""
        # Build assignment map from current state
        assignment_map = {}
        for quad_name, ch in self.assignments.items():
            if ch is not None:
                assignment_map[quad_name] = ch["name"]

        if not assignment_map:
            messagebox.showwarning(
                "Nothing to Save",
                "Assign at least one channel to a quadrant before saving a preset.",
            )
            return

        # Prompt for name
        name = simpledialog.askstring(
            "Save Preset", "Preset name:", parent=self.root
        )
        if not name or not name.strip():
            return
        name = name.strip()

        # Check for duplicate name
        for p in self.presets:
            if p["name"] == name:
                if not messagebox.askyesno(
                    "Preset Exists",
                    f"A preset named '{name}' already exists. Overwrite it?",
                    parent=self.root,
                ):
                    return
                p["assignments"] = assignment_map
                save_presets(self.presets)
                self._refresh_preset_combo()
                self._preset_var.set(name)
                return

        self.presets.append({"name": name, "assignments": assignment_map})
        save_presets(self.presets)
        self._refresh_preset_combo()
        self._preset_var.set(name)

    def _overwrite_preset(self):
        """Update the selected preset with the current quadrant assignments."""
        idx = self._get_selected_preset_idx()
        if idx is None:
            messagebox.showwarning("No Preset", "Select a preset to overwrite.")
            return

        assignment_map = {}
        for quad_name, ch in self.assignments.items():
            if ch is not None:
                assignment_map[quad_name] = ch["name"]

        name = self.presets[idx]["name"]
        if not messagebox.askyesno(
            "Confirm Overwrite",
            f"Overwrite preset '{name}' with current assignments?",
            parent=self.root,
        ):
            return

        self.presets[idx]["assignments"] = assignment_map
        save_presets(self.presets)

    def _delete_preset(self):
        """Delete the selected preset."""
        idx = self._get_selected_preset_idx()
        if idx is None:
            messagebox.showwarning("No Preset", "Select a preset to delete.")
            return
        name = self.presets[idx]["name"]
        if not messagebox.askyesno(
            "Confirm Delete", f"Delete preset '{name}'?", parent=self.root
        ):
            return
        self.presets.pop(idx)
        save_presets(self.presets)
        self._refresh_preset_combo()

    # ---- Window management --------------------------------------------------

    def _close_quadrant(self, quad_name):
        """Close a single quadrant's Chrome window."""
        pid = self.active_pids.get(quad_name)
        if not pid:
            return
        # Find and remove the matching process from self.processes
        proc_to_close = None
        for proc in self.processes:
            if proc.pid == pid:
                proc_to_close = proc
                break

        def do_close():
            if proc_to_close and proc_to_close.poll() is None:
                try:
                    subprocess.run(
                        ["taskkill", "/PID", str(proc_to_close.pid)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                    proc_to_close.wait(timeout=5)
                except (OSError, subprocess.TimeoutExpired):
                    try:
                        proc_to_close.kill()
                    except OSError:
                        pass

        threading.Thread(target=do_close, daemon=True).start()

        if proc_to_close:
            self.processes.remove(proc_to_close)
        self.active_ports.pop(quad_name, None)
        self.active_pids.pop(quad_name, None)
        self._quad_rects.pop(quad_name, None)
        self._quad_maximized.pop(quad_name, None)
        self.quad_max_btns[quad_name].config(text="Max")

    def _open_quadrant(self, quad_name):
        """Launch a single quadrant's Chrome window (e.g. after closing one)."""
        if self.active_ports.get(quad_name):
            messagebox.showinfo(
                "Already Running",
                f"{quad_name} already has an active window.",
            )
            return
        channel = self.assignments.get(quad_name)
        if not channel:
            messagebox.showwarning(
                "No Channel",
                f"No channel assigned to {quad_name}.\n"
                "Use 'Set' or drag a channel first.",
            )
            return
        url = channel.get("url", "")
        if not url:
            messagebox.showwarning(
                "Missing URL", f"No URL set for {channel['name']}.",
            )
            return
        chrome = find_chrome()
        if chrome is None:
            messagebox.showerror(
                "Chrome Not Found",
                "Google Chrome was not found at the expected location.",
            )
            return

        # Compute geometry based on all currently active quadrants + this one
        work_x, work_y, work_w, work_h = get_work_area()
        x, y, w, h = get_quadrant_rect(quad_name, work_x, work_y, work_w, work_h)
        self._quad_rects[quad_name] = (x, y, w, h)

        profile_dir = self._get_profile_dir(quad_name)
        debug_port = CDP_BASE_PORT + CDP_PORT_OFFSETS[quad_name]
        self.active_ports[quad_name] = debug_port

        cmd = [
            chrome,
            f"--app={url}",
            f"--window-position={x},{y}",
            f"--window-size={w},{h}",
            f"--user-data-dir={profile_dir}",
            f"--remote-debugging-port={debug_port}",
            "--disable-infobars",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-features=MediaRouter",
        ]
        if "spectrum.net" in url and self.block_spectrum_iha.get():
            cmd.append(
                "--host-rules="
                "MAP login.spectrum.net 0.0.0.0,"
                "MAP idp.spectrum.net 0.0.0.0"
            )
            cmd.append("--test-type")

        proc = subprocess.Popen(cmd)
        self.processes.append(proc)
        self.active_pids[quad_name] = proc.pid
        self._quad_maximized[quad_name] = False

        muted = quad_name != self.audio_quad
        threading.Thread(
            target=inject_js_thread, args=(debug_port, muted, url), daemon=True
        ).start()
        threading.Thread(
            target=unpause_thread, args=(debug_port, w, h, url), daemon=True
        ).start()

        # Keep the app in front after Chrome launches
        self.root.after(500, self._raise_app)

    def _raise_app(self):
        """Bring the QuadViewer app window to the front."""
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(100, lambda: self.root.attributes("-topmost", False))
        self.root.focus_force()

    def _bring_to_front(self):
        """Bring all Chrome windows to the OS foreground."""
        if not self.active_pids:
            return
        for pid in self.active_pids.values():
            threading.Thread(
                target=bring_os_window_to_front,
                args=(pid,),
                daemon=True,
            ).start()

    def _toggle_app_or_tv(self):
        """Ctrl+9: toggle between showing the app and showing the TV windows."""
        if self._ctrl9_showing_app:
            # Second press: bring TV windows back to front
            self._bring_to_front()
            self._ctrl9_showing_app = False
        else:
            # First press: bring the app to front
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(100, lambda: self.root.attributes("-topmost", False))
            self.root.focus_force()
            self._ctrl9_showing_app = True

    def _bring_quad_to_front(self, quad_name):
        """Bring a single quadrant's Chrome window to the OS foreground."""
        pid = self.active_pids.get(quad_name)
        if pid:
            threading.Thread(
                target=bring_os_window_to_front,
                args=(pid,),
                daemon=True,
            ).start()

    def _toggle_maximize(self, quad_name):
        """Toggle a quadrant's window between maximized (fullscreen) and original size."""
        port = self.active_ports.get(quad_name)
        if not port:
            return

        ch = self.assignments.get(quad_name)
        url = ch.get("url", "") if ch else ""
        is_spectrum = "spectrum.net" in url

        if self._quad_maximized.get(quad_name, False):
            # Restore to original rect (don't bring to front)
            self._restore_quad(quad_name, is_spectrum)
        else:
            # Restore any other maximized window first
            for other_q in QUADRANTS:
                if other_q != quad_name and self._quad_maximized.get(other_q, False):
                    other_ch = self.assignments.get(other_q)
                    other_spectrum = "spectrum.net" in (other_ch.get("url", "") if other_ch else "")
                    self._restore_quad(other_q, other_spectrum)

            # Maximize to full work area
            work_x, work_y, work_w, work_h = get_work_area()
            x = work_x - WIN_BORDER
            y = work_y - WIN_BORDER
            w = work_w + 2 * WIN_BORDER
            h = work_h + 2 * WIN_BORDER
            pid = self.active_pids.get(quad_name)
            threading.Thread(
                target=self._do_resize_and_front,
                args=(port, pid, x, y, w, h, is_spectrum),
                daemon=True,
            ).start()
            self._quad_maximized[quad_name] = True
            self.quad_max_btns[quad_name].config(text="Shrink")
            # Switch audio to the maximized window
            self._set_audio_solo(quad_name)

    def _restore_quad(self, quad_name, is_spectrum=False):
        """Restore a maximized quadrant to its original size without bringing to front."""
        port = self.active_ports.get(quad_name)
        rect = self._quad_rects.get(quad_name)
        if port and rect:
            x, y, w, h = rect
            threading.Thread(
                target=self._do_resize,
                args=(port, x, y, w, h, is_spectrum),
                daemon=True,
            ).start()
        self._quad_maximized[quad_name] = False
        self.quad_max_btns[quad_name].config(text="Max")

    def _do_resize(self, port, x, y, w, h, is_spectrum=False):
        """Background: resize window, resume Spectrum video if needed."""
        cdp_set_window_bounds(port, x, y, w, h)
        if is_spectrum:
            # Browser may natively pause video on resize; resume with
            # userGesture=True so Chrome's autoplay policy allows play().
            time.sleep(2)
            cdp_evaluate(port, RESUME_VIDEO_JS, retries=2, delay=1, user_gesture=True)
            time.sleep(3)
            cdp_evaluate(port, RESUME_VIDEO_JS, retries=2, delay=1, user_gesture=True)

    def _do_resize_and_front(self, port, pid, x, y, w, h, is_spectrum=False):
        """Background: resize window then bring to front."""
        cdp_set_window_bounds(port, x, y, w, h)
        time.sleep(0.3)
        if pid:
            bring_os_window_to_front(pid)
        if is_spectrum:
            # Browser may natively pause video on resize; resume with
            # userGesture=True so Chrome's autoplay policy allows play().
            time.sleep(2)
            cdp_evaluate(port, RESUME_VIDEO_JS, retries=2, delay=1, user_gesture=True)
            time.sleep(3)
            cdp_evaluate(port, RESUME_VIDEO_JS, retries=2, delay=1, user_gesture=True)

    # ---- Launch / close ----------------------------------------------------

    def _get_profile_dir(self, quad_name):
        profile = os.path.join(PROFILES_DIR, QUAD_PROFILE_NAMES[quad_name])
        os.makedirs(profile, exist_ok=True)
        return profile

    def _execute(self):
        chrome = find_chrome()
        if chrome is None:
            messagebox.showerror(
                "Chrome Not Found",
                "Google Chrome was not found at the expected location.\n"
                "Please install Chrome or update the CHROME_PATHS in the script.",
            )
            return

        active = {q: ch for q, ch in self.assignments.items() if ch is not None}
        if not active:
            messagebox.showwarning(
                "Nothing to Launch", "Assign at least one channel to a quadrant."
            )
            return

        work_x, work_y, work_w, work_h = get_work_area()

        self._close_all()
        self.active_ports.clear()
        self.active_pids.clear()
        self._quad_rects.clear()
        self._quad_maximized.clear()
        for btn in self.quad_max_btns.values():
            btn.config(text="Max")

        rects = get_smart_rects(active, work_x, work_y, work_w, work_h)

        for quad_name, channel in active.items():
            url = channel.get("url", "")
            if not url:
                messagebox.showwarning(
                    "Missing URL",
                    f"No URL set for {channel['name']}. Skipping.\n"
                    "Use Edit to add a URL for this channel.",
                )
                continue

            x, y, w, h = rects[quad_name]
            self._quad_rects[quad_name] = (x, y, w, h)
            profile_dir = self._get_profile_dir(quad_name)
            debug_port = CDP_BASE_PORT + CDP_PORT_OFFSETS[quad_name]
            self.active_ports[quad_name] = debug_port

            cmd = [
                chrome,
                f"--app={url}",
                f"--window-position={x},{y}",
                f"--window-size={w},{h}",
                f"--user-data-dir={profile_dir}",
                f"--remote-debugging-port={debug_port}",
                "--disable-infobars",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-features=MediaRouter",
            ]

            # Block Spectrum in-home authentication (IP-based auto-login)
            # so the saved session cookies for the desired account persist
            if "spectrum.net" in url and self.block_spectrum_iha.get():
                cmd.append(
                    "--host-rules="
                    "MAP login.spectrum.net 0.0.0.0,"
                    "MAP idp.spectrum.net 0.0.0.0"
                )
                cmd.append("--test-type")  # suppress "unsupported flag" banner

            proc = subprocess.Popen(cmd)
            self.processes.append(proc)
            self.active_pids[quad_name] = proc.pid

            # Inject auto-click JS in a background thread after Chrome loads
            muted = quad_name != "Upper Left"
            t = threading.Thread(
                target=inject_js_thread, args=(debug_port, muted, url), daemon=True
            )
            t.start()

            # Click center of video at 20s to unpause
            t2 = threading.Thread(
                target=unpause_thread, args=(debug_port, w, h, url), daemon=True
            )
            t2.start()

        # Set initial audio: mute all except Upper Left after videos load
        if self.active_ports:
            threading.Thread(
                target=self._initial_mute, daemon=True
            ).start()

    # ---- Audio control -------------------------------------------------------

    MUTE_JS = "document.querySelectorAll('video, audio').forEach(function(el){el.muted=true;});"
    UNMUTE_JS = "document.querySelectorAll('video, audio').forEach(function(el){el.muted=false;});"

    def _mute_js_for(self, url):
        """Return the correct mute JS for a given channel URL."""
        if "youtube.com" in url or "youtu.be" in url:
            return YT_MUTE_JS
        return self.MUTE_JS

    def _unmute_js_for(self, url):
        """Return the correct unmute JS for a given channel URL."""
        if "youtube.com" in url or "youtu.be" in url:
            return YT_UNMUTE_JS
        return self.UNMUTE_JS

    _QUAD_ORDER = ["Upper Left", "Upper Right", "Lower Left", "Lower Right"]

    def _hotkey_loop(self):
        """Background thread: listen for global Ctrl+1..4 hotkeys."""
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        # Store thread ID so _on_close can post WM_QUIT to exit GetMessageW
        self._hotkey_thread_id = kernel32.GetCurrentThreadId()

        # Set proper ctypes signatures for reliability
        user32.RegisterHotKey.argtypes = [
            ctypes.wintypes.HWND, ctypes.c_int,
            ctypes.wintypes.UINT, ctypes.wintypes.UINT,
        ]
        user32.RegisterHotKey.restype = ctypes.wintypes.BOOL
        user32.GetMessageW.argtypes = [
            ctypes.POINTER(ctypes.wintypes.MSG), ctypes.wintypes.HWND,
            ctypes.wintypes.UINT, ctypes.wintypes.UINT,
        ]
        user32.GetMessageW.restype = ctypes.wintypes.BOOL

        MOD_CTRL = 0x0002
        VK_1 = 0x31
        WM_HOTKEY = 0x0312

        # Register Ctrl+1..4 (audio), Ctrl+5..8 (maximize), Ctrl+9 (app/tv toggle), Ctrl+0 (mute all)
        VK_0 = 0x30
        for i in range(9):
            user32.RegisterHotKey(None, i + 1, MOD_CTRL, VK_1 + i)
        user32.RegisterHotKey(None, 10, MOD_CTRL, VK_0)

        # Blocking message loop — GetMessageW returns 0 on WM_QUIT
        msg = ctypes.wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == WM_HOTKEY:
                hk_id = msg.wParam
                if 1 <= hk_id <= 4:
                    quad = self._QUAD_ORDER[hk_id - 1]
                    self.root.after(0, self._set_audio_solo, quad)
                elif 5 <= hk_id <= 8:
                    quad = self._QUAD_ORDER[hk_id - 5]
                    self.root.after(0, self._toggle_maximize, quad)
                elif hk_id == 9:
                    self.root.after(0, self._toggle_app_or_tv)
                elif hk_id == 10:
                    self.root.after(0, self._mute_all)

        # Cleanup
        for i in range(10):
            user32.UnregisterHotKey(None, i + 1)

    def _initial_mute(self):
        """Wait for videos to load, then mute all except Upper Left."""
        time.sleep(30)
        for q, port in self.active_ports.items():
            ch = self.assignments.get(q)
            url = ch.get("url", "") if ch else ""
            js = self._unmute_js_for(url) if q == "Upper Left" else self._mute_js_for(url)
            cdp_evaluate(port, js, retries=3, delay=1)
        self.audio_quad = "Upper Left"
        self.root.after(0, self._update_audio_indicator)

    def _set_audio_solo(self, quad_name):
        """Unmute the given quadrant, mute all others."""
        if quad_name not in self.active_ports:
            return
        self.audio_quad = quad_name
        for q, port in self.active_ports.items():
            ch = self.assignments.get(q)
            url = ch.get("url", "") if ch else ""
            js = self._unmute_js_for(url) if q == quad_name else self._mute_js_for(url)
            is_spectrum = "spectrum.net" in url
            threading.Thread(
                target=self._do_audio_switch,
                args=(port, js, is_spectrum),
                daemon=True,
            ).start()
        self._update_audio_indicator()

    def _mute_all(self):
        """Mute all quadrants."""
        if not self.active_ports:
            return
        self.audio_quad = None
        for q, port in self.active_ports.items():
            ch = self.assignments.get(q)
            url = ch.get("url", "") if ch else ""
            js = self._mute_js_for(url)
            threading.Thread(
                target=cdp_evaluate, args=(port, js, 3, 1), daemon=True
            ).start()
        self._update_audio_indicator()

    def _do_audio_switch(self, port, js, is_spectrum):
        """Background: mute/unmute, then resume Spectrum if it paused."""
        cdp_evaluate(port, js, retries=3, delay=1)
        if is_spectrum:
            time.sleep(1)
            cdp_evaluate(port, RESUME_VIDEO_JS, retries=2, delay=1, user_gesture=True)

    def _update_audio_indicator(self):
        """Update quadrant labels to show which has audio (preserves logos)."""
        for quad_name, label in self.quad_labels.items():
            ch = self.assignments.get(quad_name)
            name = ch["name"] if ch else "(empty)"
            if quad_name == self.audio_quad and quad_name in self.active_ports:
                label.config(text=f"{name}  \u266b")  # musical note
            else:
                label.config(text=name)
            # Refresh logo in case it was lost
            if ch:
                logo = self._get_logo(ch.get("logo", ""), LOGO_LARGE)
                if logo:
                    self.quad_logos[quad_name].config(image=logo)

    def _close_all(self):
        """Gracefully close Chrome windows so cookies/sessions are saved."""
        for proc in self.processes:
            if proc.poll() is not None:
                continue
            try:
                subprocess.run(
                    ["taskkill", "/PID", str(proc.pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            except OSError:
                pass
        for proc in self.processes:
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
        self.processes.clear()
        self.active_ports.clear()
        self.active_pids.clear()

    # ---- Preferences --------------------------------------------------------

    def _show_preferences(self):
        """Show the Preferences dialog with theme picker."""
        win = tk.Toplevel(self.root)
        _set_icon(win)
        win.title("Preferences")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        pad = {"padx": 10, "pady": 6}

        ttk.Label(win, text="Visual Theme:", font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w", **pad
        )

        # Use ttkbootstrap's own style object (not bare ttk.Style())
        tbs_style = self.root.style
        all_themes = sorted([
            "cosmo", "flatly", "litera", "minty", "lumen", "sandstone", "yeti",
            "pulse", "united", "morph", "journal", "darkly", "superhero",
            "solar", "cyborg", "vapor", "simplex", "cerculean",
        ])
        current_theme = tbs_style.theme_use()

        theme_var = tk.StringVar(value=current_theme)
        theme_combo = ttk.Combobox(
            win, textvariable=theme_var, values=all_themes,
            state="readonly", width=25,
        )
        theme_combo.grid(row=0, column=1, **pad)

        # Live preview
        preview_label = ttk.Label(win, text="")
        preview_label.grid(row=1, column=0, columnspan=2, **pad)

        def on_theme_change(event=None):
            name = theme_var.get()
            try:
                tbs_style.theme_use(name)
                preview_label.config(text=f"Preview: {name}")
            except Exception:
                preview_label.config(text=f"Could not load: {name}")

        theme_combo.bind("<<ComboboxSelected>>", on_theme_change)

        def apply_and_close():
            chosen = theme_var.get()
            try:
                tbs_style.theme_use(chosen)
            except Exception:
                pass
            settings = load_settings()
            settings["theme"] = chosen
            save_settings(settings)
            # Re-apply hover style for the new theme
            tbs_style.configure("Hover.TLabelframe", background="#1a3a5c")
            tbs_style.configure("Hover.TLabelframe.Label", background="#1a3a5c")
            win.destroy()

        def cancel():
            # Restore original theme
            try:
                tbs_style.theme_use(current_theme)
            except Exception:
                pass
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="OK", command=apply_and_close).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(btn_frame, text="Cancel", command=cancel).pack(
            side=tk.LEFT, padx=6
        )

        # Center on parent
        win.update_idletasks()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        px = self.root.winfo_x()
        py = self.root.winfo_y()
        w = win.winfo_width()
        h = win.winfo_height()
        win.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    # ---- Help / About --------------------------------------------------------

    def _show_help(self):
        """Display the QuadViewer help documentation."""
        help_text = (
            "DEVCON QUADVIEWER HELP\n"
            "======================\n\n"
            "DevCon QuadViewer launches up to four streaming TV channels in borderless\n"
            "Chrome windows arranged in screen quadrants.\n\n"
            "GETTING STARTED\n"
            "---------------\n"
            "1. Select a channel from the list on the left.\n"
            "2. Click 'Set' on one of the four quadrants (Upper Left, Upper Right,\n"
            "   Lower Left, Lower Right), or drag the channel onto a quadrant.\n"
            "3. Repeat for up to four channels.\n"
            "4. Click 'Execute' to launch all assigned channels.\n\n"
            "CHANNEL MANAGEMENT\n"
            "------------------\n"
            "Add      - Create a new channel (name, URL, logo, TVGuide name).\n"
            "Edit     - Modify the selected channel's settings.\n"
            "Delete   - Remove the selected channel.\n"
            "Drag     - Drag channels in the list to reorder them, or drag\n"
            "           onto a quadrant to assign.\n\n"
            "QUADRANT CONTROLS\n"
            "-----------------\n"
            "Set      - Assign the selected channel to this quadrant.\n"
            "Clear    - Remove the channel from this quadrant.\n"
            "Front    - Bring this quadrant's window to the foreground.\n"
            "Max      - Maximize this window to full screen (becomes 'Shrink'\n"
            "           to restore). Maximizing auto-restores any other\n"
            "           maximized window and switches audio to this channel.\n"
            "Switch   - Navigate a running window to its assigned channel.\n"
            "           Use this to swap channels without re-launching:\n"
            "           1. Set a new channel on the quadrant.\n"
            "           2. Click Switch to navigate the existing window.\n"
            "Close    - Close just this quadrant's Chrome window.\n"
            "Open     - Launch this quadrant's window individually\n"
            "           (e.g. after closing it).\n\n"
            "KEYBOARD SHORTCUTS (global - work even when QuadViewer is behind)\n"
            "-------------------------------------------------------------------\n"
            "Ctrl+1   - Audio to Upper Left\n"
            "Ctrl+2   - Audio to Upper Right\n"
            "Ctrl+3   - Audio to Lower Left\n"
            "Ctrl+4   - Audio to Lower Right\n"
            "Ctrl+0   - Mute all channels\n"
            "Ctrl+5   - Maximize / Shrink Upper Left\n"
            "Ctrl+6   - Maximize / Shrink Upper Right\n"
            "Ctrl+7   - Maximize / Shrink Lower Left\n"
            "Ctrl+8   - Maximize / Shrink Lower Right\n"
            "Ctrl+9   - Toggle between QuadViewer app and TV windows\n\n"
            "PRESETS\n"
            "-------\n"
            "Save     - Save current quadrant assignments as a named preset.\n"
            "Load     - Restore a previously saved preset.\n"
            "Overwrite- Update an existing preset with current assignments.\n"
            "Delete   - Remove a saved preset.\n\n"
            "PROGRAMMING GUIDE\n"
            "-----------------\n"
            "Hover over a channel in the list to see what's currently airing\n"
            "and what's on next (requires a TVGuide Name set for the channel).\n\n"
            "OTHER CONTROLS\n"
            "--------------\n"
            "Bring to Front - Bring all TV windows to the foreground.\n"
            "Clear All      - Remove all quadrant assignments.\n"
            "Close All      - Close all launched Chrome windows.\n"
            "Block Spectrum auto-login - Prevents Spectrum from overriding\n"
            "                 the saved login with in-home authentication.\n"
        )
        win = tk.Toplevel(self.root)
        _set_icon(win)
        win.title("DevCon QuadViewer Help")
        win.geometry("600x520")
        win.transient(self.root)
        text = tk.Text(
            win, wrap=tk.WORD, padx=12, pady=12,
            font=("Consolas", 10),
            bg="#2b2b2b", fg="#e0e0e0",
            insertbackground="#e0e0e0",
            selectbackground="#3a5a8c",
        )
        text.insert("1.0", help_text)
        text.config(state=tk.DISABLED)
        scroll = ttk.Scrollbar(win, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(fill=tk.BOTH, expand=True)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=8)

    def _open_youtube_tutorial(self):
        """Open the YouTube tutorial in the default browser."""
        import webbrowser
        webbrowser.open("https://www.youtube.com/watch?v=1tzupqW2n1g")

    def _show_about(self):
        """Show the About dialog with developer photo."""
        win = tk.Toplevel(self.root)
        _set_icon(win)
        win.title("About DevCon QuadViewer")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        frame = ttk.Frame(win, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Developer photo
        self._about_photo = None
        dev_path = os.path.join(_DATA_DIR, "developer.jpg")
        if os.path.isfile(dev_path):
            try:
                img = Image.open(dev_path)
                img.thumbnail((150, 150), Image.LANCZOS)
                self._about_photo = ImageTk.PhotoImage(img)
                ttk.Label(frame, image=self._about_photo).pack(pady=(0, 10))
            except Exception:
                pass

        about_text = (
            "DevCon QuadViewer\n"
            "Version 1.0\n\n"
            "by DevCon Productions\n"
            "Cleveland, Ohio, USA\n\n"
            "Copyright \u00a9 2026 by DevCon Productions\n"
            "MIT License\n\n"
            "This software is provided as-is, without warranty\n"
            "of any kind, express or implied.\n\n"
            "Vibe Coded with Claude Code"
        )
        ttk.Label(
            frame, text=about_text, justify=tk.CENTER,
            font=("Segoe UI", 10),
        ).pack(pady=(0, 10))

        ttk.Button(frame, text="OK", command=win.destroy).pack()

        # Center on parent
        win.update_idletasks()
        pw = self.root.winfo_width()
        ph = self.root.winfo_height()
        px = self.root.winfo_x()
        py = self.root.winfo_y()
        w = win.winfo_width()
        h = win.winfo_height()
        win.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    # ---- Import / Export -----------------------------------------------------

    def _export_data(self):
        """Export channels, presets, and logos to a single JSON file."""
        from tkinter import filedialog
        import base64
        path = filedialog.asksaveasfilename(
            title="Export Channels & Presets",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile="QuadViewer_Export.json",
        )
        if not path:
            return
        channels = load_channels()
        # Collect referenced logo files as base64
        logos = {}
        for ch in channels:
            logo_name = ch.get("logo", "")
            if logo_name and logo_name not in logos:
                logo_path = os.path.join(LOGOS_DIR, logo_name)
                if os.path.isfile(logo_path):
                    with open(logo_path, "rb") as lf:
                        logos[logo_name] = base64.b64encode(lf.read()).decode("ascii")
        data = {
            "channels": channels,
            "presets": load_presets(),
            "logos": logos,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        messagebox.showinfo("Export Complete", f"Channels and presets exported to:\n{path}")

    def _import_data(self):
        """Import channels, presets, and logos from a previously exported JSON file."""
        from tkinter import filedialog
        import base64
        path = filedialog.askopenfilename(
            title="Import Channels & Presets",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            messagebox.showerror("Import Error", f"Could not read file:\n{e}")
            return
        if "channels" not in data and "presets" not in data:
            messagebox.showerror(
                "Import Error",
                "This file does not appear to be a valid QuadViewer export.\n"
                "Expected 'channels' and/or 'presets' keys.",
            )
            return
        confirm = messagebox.askyesno(
            "Confirm Import",
            "This will replace your current channels and presets.\n\nContinue?",
        )
        if not confirm:
            return
        if "channels" in data:
            save_channels(data["channels"])
        if "presets" in data:
            save_presets(data["presets"])
        # Restore logo images
        if "logos" in data:
            os.makedirs(LOGOS_DIR, exist_ok=True)
            for name, b64 in data["logos"].items():
                logo_path = os.path.join(LOGOS_DIR, name)
                with open(logo_path, "wb") as lf:
                    lf.write(base64.b64decode(b64))
        # Reload into the running app
        self.channels = load_channels()
        self._populate_tree()
        self.presets = load_presets()
        self._refresh_preset_combo()
        messagebox.showinfo("Import Complete", "Channels and presets imported successfully.")

    def _on_close(self):
        # Post WM_QUIT to the hotkey thread so GetMessageW returns 0
        if hasattr(self, "_hotkey_thread_id"):
            ctypes.windll.user32.PostThreadMessageW(
                self._hotkey_thread_id, 0x0012, 0, 0  # WM_QUIT
            )
        self._close_all()
        self.root.destroy()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    # Set AppUserModelID so Windows shows our icon in the taskbar
    # (without this, Python's default icon is used)
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DevCon.QuadViewer.1")

    # ---- Splash screen -------------------------------------------------------
    if os.path.isfile(SPLASH_PATH):
        import tkinter as _tk
        splash_root = _tk.Tk()
        splash_root.overrideredirect(True)
        from PIL import Image, ImageTk
        _pil_img = Image.open(SPLASH_PATH)
        _pil_img = _pil_img.resize(
            (int(_pil_img.width * 0.6), int(_pil_img.height * 0.6)),
            Image.LANCZOS,
        )
        splash_img = ImageTk.PhotoImage(_pil_img)
        img_w, img_h = _pil_img.width, _pil_img.height
        sx = (splash_root.winfo_screenwidth() - img_w) // 2
        sy = (splash_root.winfo_screenheight() - img_h) // 2
        splash_root.geometry(f"{img_w}x{img_h}+{sx}+{sy}")
        _tk.Label(splash_root, image=splash_img, bd=0).pack()
        splash_root.update()
        splash_root.after(3000, splash_root.destroy)
        splash_root.mainloop()

    # ---- Main window ----------------------------------------------------------
    settings = load_settings()
    theme = settings.get("theme", DEFAULT_THEME)

    # Validate theme is a real ttkbootstrap theme
    TTKB_THEMES = [
        "cosmo", "flatly", "litera", "minty", "lumen", "sandstone", "yeti",
        "pulse", "united", "morph", "journal", "darkly", "superhero",
        "solar", "cyborg", "vapor", "simplex", "cerculean",
    ]
    if theme not in TTKB_THEMES:
        theme = DEFAULT_THEME

    root = tbs.Window(
        title="DevCon QuadViewer",
        themename=theme,
        size=(1180, 700),
        minsize=(1180, 700),
        hdpi=False,
    )

    # Window / taskbar icon
    if os.path.isfile(ICO_PATH):
        root.iconbitmap(ICO_PATH)

    style = root.style
    style.configure("Hover.TLabelframe", background="#1a3a5c")
    style.configure("Hover.TLabelframe.Label", background="#1a3a5c")

    app = QuadViewerApp(root)

    root.mainloop()


if __name__ == "__main__":
    main()
