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
import tkinter as tk
from tkinter import ttk, messagebox

# ---------------------------------------------------------------------------
# Paths & constants
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CHANNELS_FILE = os.path.join(SCRIPT_DIR, "channels.json")
ASSIGNMENTS_FILE = os.path.join(SCRIPT_DIR, "assignments.json")
PROFILES_DIR = os.path.join(SCRIPT_DIR, "profiles")

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


def get_work_area():
    """Get usable screen rectangle excluding the taskbar (Windows)."""
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

  // --- Spectrum overlay: click #modal-close ---
  function dismissSpectrumOverlay() {
    var closeBtn = document.querySelector("#modal-close");
    if (closeBtn) {
      closeBtn.dispatchEvent(new MouseEvent("mousedown", {bubbles:true}));
      closeBtn.dispatchEvent(new MouseEvent("mouseup", {bubbles:true}));
      closeBtn.click();
      return true;
    }
    return false;
  }

  function handleProfile() {
    if (isFubo) {
      if (fuboSelectProfile()) return true;
      return clickProfileElement();
    }
    if (isSpectrum) return clickProfileElement();
    return clickProfileElement();
  }

  function handleOverlay() {
    if (isSpectrum) return dismissSpectrumOverlay();
    return false;
  }

  var timer = setInterval(function() {
    attempts++;
    if (!profileDone && handleProfile()) profileDone = true;
    if (!overlayDone && handleOverlay()) overlayDone = true;
    if ((profileDone && overlayDone) || attempts >= 60) clearInterval(timer);
  }, 1000);
})();
""".replace("__PROFILE__", PROFILE_NAME)

def inject_js_thread(port):
    """Background thread: inject JS multiple times to survive page navigations."""
    # First injection: as soon as Chrome is reachable
    cdp_evaluate(port, INJECTED_JS)
    # Re-inject after delays to catch pages that navigate/redirect
    for delay in (10, 15):
        time.sleep(delay)
        cdp_evaluate(port, INJECTED_JS, retries=3, delay=1)


def unpause_thread(port, win_w, win_h):
    """Background thread: click center of video at 20s to unpause."""
    time.sleep(20)
    # Click center of the viewport to toggle play
    cdp_mouse_click(port, win_w // 2, win_h // 2)


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
        ttk.Entry(self, textvariable=self.name_var, width=50).grid(row=0, column=1, **pad)

        ttk.Label(self, text="URL:").grid(row=1, column=0, sticky="w", **pad)
        self.url_var = tk.StringVar(value=channel.get("url", "") if channel else "")
        ttk.Entry(self, textvariable=self.url_var, width=50).grid(row=1, column=1, **pad)

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
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

    def _ok(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Missing Name", "Channel name is required.", parent=self)
            return
        self.result = {
            "name": name,
            "url": self.url_var.get().strip(),
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

    def _apply_restored_assignments(self):
        """Update the GUI labels to reflect restored assignments."""
        for quad_name, ch in self.assignments.items():
            if ch is not None:
                self.quad_labels[quad_name].config(text=ch["name"])

    def _save_assignments(self):
        """Persist current quadrant assignments to disk."""
        data = {}
        for quad_name, ch in self.assignments.items():
            if ch is not None and ch in self.channels:
                data[quad_name] = self.channels.index(ch)
        save_assignments(data)

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
            tree_frame, columns=("name",), show="headings", selectmode="browse"
        )
        self.channel_tree.heading("name", text="Channel")
        self.channel_tree.column("name", width=180)
        scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.channel_tree.yview
        )
        self.channel_tree.configure(yscrollcommand=scrollbar.set)
        self.channel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

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

            label = ttk.Label(frame, text="(empty)", anchor="center", width=18)
            label.pack(pady=(0, 6))
            self.quad_labels[quad_name] = label

            btn_frame = ttk.Frame(frame)
            btn_frame.pack()
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

        # Bottom controls
        controls = ttk.Frame(right_frame, padding=(0, 8, 0, 0))
        controls.pack(fill=tk.X)

        ttk.Button(controls, text="Execute", command=self._execute).pack(
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

    def _populate_tree(self):
        for item in self.channel_tree.get_children():
            self.channel_tree.delete(item)
        for ch in self.channels:
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
                    self.quad_labels[quad_name].config(
                        text=self._drag_channel["name"]
                    )
                    self._save_assignments()
                    break

        self._cleanup_drag()

    def _cleanup_drag(self):
        self._drag_channel = None
        self._drag_source_idx = None
        self._drag_is_reorder = False
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
            self.channels[idx] = dlg.result
            save_channels(self.channels)
            self._populate_tree()
            for q, ch in self.assignments.items():
                if ch and ch["name"] == old_name:
                    self.assignments[q] = dlg.result
                    self.quad_labels[q].config(text=dlg.result["name"])
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
                self.quad_labels[q].config(text="(empty)")
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
        self.quad_labels[quad_name].config(text=channel["name"])
        self._save_assignments()

    def _clear_quadrant(self, quad_name):
        self.assignments[quad_name] = None
        self.quad_labels[quad_name].config(text="(empty)")
        self._save_assignments()

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

        for quad_name, channel in active.items():
            url = channel.get("url", "")
            if not url:
                messagebox.showwarning(
                    "Missing URL",
                    f"No URL set for {channel['name']}. Skipping.\n"
                    "Use Edit to add a URL for this channel.",
                )
                continue

            x, y, w, h = get_quadrant_rect(quad_name, work_x, work_y, work_w, work_h)
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
        """Update quadrant labels to show which has audio."""
        for quad_name, label in self.quad_labels.items():
            ch = self.assignments.get(quad_name)
            name = ch["name"] if ch else "(empty)"
            if quad_name == self.audio_quad and quad_name in self.active_ports:
                label.config(text=f"{name}  \u266b")  # musical note
            else:
                label.config(text=name)

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
    root = tk.Tk()
    root.geometry("700x460")

    style = ttk.Style()
    style.configure("Hover.TLabelframe", background="#d0e8ff")
    style.configure("Hover.TLabelframe.Label", background="#d0e8ff")

    app = QuadViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
