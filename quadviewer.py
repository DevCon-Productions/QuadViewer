"""
QuadViewer - Quad-Screen TV Streaming Launcher

Opens a GUI to assign TV channels to four screen quadrants,
then launches borderless Chrome windows for each.
"""

import base64
import ctypes
import ctypes.wintypes
import http.client
import json
import os
import socket
import struct
import subprocess
import threading
import time
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from PIL import Image, ImageTk

# ---------------------------------------------------------------------------
# Paths & constants
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CHANNELS_FILE = os.path.join(SCRIPT_DIR, "channels.json")
ASSIGNMENTS_FILE = os.path.join(SCRIPT_DIR, "assignments.json")
PROFILES_DIR = os.path.join(SCRIPT_DIR, "profiles")
LOGOS_DIR = os.path.join(SCRIPT_DIR, "logos")

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


def cdp_evaluate(port, js_code, retries=15, delay=2):
    """Evaluate JavaScript via CDP Runtime.evaluate."""
    return cdp_send(port, "Runtime.evaluate",
                    {"expression": js_code, "awaitPromise": False},
                    retries, delay)


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
    // Press Escape to dismiss any overlay/guide
    sendKey("Escape", "Escape", 27);
    // Click the video player area to ensure focus and dismiss guide
    var video = document.querySelector("video");
    if (video && video.offsetParent !== null) {
      video.click();
      dismissed = true;
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

  // --- Universal autoplay: force-play any paused video elements ---
  function ensurePlayback() {
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

def inject_js_thread(port):
    """Background thread: inject JS multiple times to survive page navigations."""
    # First injection: as soon as Chrome is reachable
    cdp_evaluate(port, INJECTED_JS)
    # Re-inject after delays to catch pages that navigate/redirect
    for wait in (8, 12, 15, 20):
        time.sleep(wait)
        cdp_evaluate(port, INJECTED_JS, retries=3, delay=1)


def unpause_thread(port, win_w, win_h):
    """Background thread: click center of video multiple times to dismiss play overlays."""
    cx, cy = win_w // 2, win_h // 2
    for wait in (15, 10, 10, 10):
        time.sleep(wait)
        cdp_mouse_click(port, cx, cy)


# ---------------------------------------------------------------------------
# Channel Editor Dialog
# ---------------------------------------------------------------------------
class ChannelDialog(tk.Toplevel):
    """Modal dialog for adding or editing a channel."""

    def __init__(self, parent, title="Channel", channel=None):
        super().__init__(parent)
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

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=10)
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
        self.result = {
            "name": name,
            "url": self.url_var.get().strip(),
            "logo": self.logo_var.get().strip(),
        }
        self.destroy()


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------
class QuadViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QuadViewer")
        self.root.resizable(True, True)

        self.channels = load_channels()
        self.assignments = {name: None for name in QUADRANTS}
        self.processes = []
        self.active_ports = {}        # quad_name -> CDP debug port
        self.audio_quad = "Upper Left"  # currently unmuted quadrant
        self.block_spectrum_iha = tk.BooleanVar(value=True)  # block IHA by default

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
            tree_frame, columns=("name",), show="tree headings", selectmode="browse"
        )
        self.channel_tree.heading("name", text="Channel", anchor="w")
        self.channel_tree.column("#0", width=50, minwidth=50, stretch=False)  # logo icon
        self.channel_tree.column("name", width=160, anchor="w")

        # Row height to fit logo icons + bold left-aligned header
        tree_style = ttk.Style()
        tree_style.configure("Treeview", rowheight=max(28, LOGO_SMALL + 4))
        tree_style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), anchor="w")
        scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.channel_tree.yview
        )
        self.channel_tree.configure(yscrollcommand=scrollbar.set)
        self.channel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.channel_tree.tag_configure("drop_target", background="#b3d9ff")
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

        # ----- Right panel: quadrant grid + controls -----
        right_frame = ttk.Frame(main, padding=4)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(3, 0))

        grid_frame = ttk.Frame(right_frame)
        grid_frame.pack(fill=tk.BOTH, expand=True)

        self.quad_labels = {}
        self.quad_logos = {}    # quad_name -> ttk.Label for logo image
        self.quad_frames = {}
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

        # Bottom controls
        controls = ttk.Frame(right_frame, padding=(0, 8, 0, 0))
        controls.pack(fill=tk.X)

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

        # Audio controls
        audio_frame = ttk.LabelFrame(right_frame, text="Audio (Ctrl+1-4)", padding=4)
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

    def _populate_tree(self):
        for item in self.channel_tree.get_children():
            self.channel_tree.delete(item)
        for ch in self.channels:
            logo = self._get_logo(ch.get("logo", ""), LOGO_SMALL)
            if logo:
                self.channel_tree.insert("", tk.END, values=(ch["name"],), image=logo)
            else:
                self.channel_tree.insert("", tk.END, values=(ch["name"],))

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
                bg="#ffffcc",
                fg="#333333",
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
                                bg="#0066cc"
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

    # ---- Window management --------------------------------------------------

    def _bring_to_front(self):
        """Bring all Chrome windows to the foreground via CDP."""
        if not self.active_ports:
            return
        for port in self.active_ports.values():
            threading.Thread(
                target=cdp_send,
                args=(port, "Page.bringToFront", {}),
                kwargs={"retries": 2, "delay": 1},
                daemon=True,
            ).start()

    def _bring_quad_to_front(self, quad_name):
        """Bring a single quadrant's Chrome window to the foreground."""
        port = self.active_ports.get(quad_name)
        if port:
            threading.Thread(
                target=cdp_send,
                args=(port, "Page.bringToFront", {}),
                kwargs={"retries": 2, "delay": 1},
                daemon=True,
            ).start()

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

            # Inject auto-click JS in a background thread after Chrome loads
            t = threading.Thread(
                target=inject_js_thread, args=(debug_port,), daemon=True
            )
            t.start()

            # Click center of video at 20s to unpause
            t2 = threading.Thread(
                target=unpause_thread, args=(debug_port, w, h), daemon=True
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

        # Register Ctrl+1 through Ctrl+4
        for i in range(4):
            user32.RegisterHotKey(None, i + 1, MOD_CTRL, VK_1 + i)

        # Blocking message loop — GetMessageW returns 0 on WM_QUIT
        msg = ctypes.wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == WM_HOTKEY:
                hk_id = msg.wParam
                if 1 <= hk_id <= 4:
                    quad = self._QUAD_ORDER[hk_id - 1]
                    self.root.after(0, self._set_audio_solo, quad)

        # Cleanup
        for i in range(4):
            user32.UnregisterHotKey(None, i + 1)

    def _initial_mute(self):
        """Wait for videos to load, then mute all except Upper Left."""
        time.sleep(30)
        for q, port in self.active_ports.items():
            js = self.UNMUTE_JS if q == "Upper Left" else self.MUTE_JS
            cdp_evaluate(port, js, retries=3, delay=1)
        self.audio_quad = "Upper Left"
        self.root.after(0, self._update_audio_indicator)

    def _set_audio_solo(self, quad_name):
        """Unmute the given quadrant, mute all others."""
        if quad_name not in self.active_ports:
            return
        self.audio_quad = quad_name
        for q, port in self.active_ports.items():
            js = self.UNMUTE_JS if q == quad_name else self.MUTE_JS
            threading.Thread(
                target=cdp_evaluate, args=(port, js, 3, 1), daemon=True
            ).start()
        self._update_audio_indicator()

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

    root = tk.Tk()
    root.geometry("700x540")

    # Window / taskbar icon
    ico_path = os.path.join(SCRIPT_DIR, "quadviewer.ico")
    if os.path.isfile(ico_path):
        root.iconbitmap(ico_path)

    style = ttk.Style()
    style.configure("Hover.TLabelframe", background="#d0e8ff")
    style.configure("Hover.TLabelframe.Label", background="#d0e8ff")

    app = QuadViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
