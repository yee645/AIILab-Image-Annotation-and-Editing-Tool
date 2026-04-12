"""Azure Kinect 影片標記工具

快捷鍵：
  A / D       — 後退 / 前進 1 幀
  W / S       — 後退 / 前進 30 幀
  Shift+A/D   — 跳到影片開頭 / 結尾
  Left/Right  — 等同 A/D（後退 / 前進 1 幀）
  Shift+Left/Right — 跳到影片開頭 / 結尾
  I           — 標記 start_us
  O           — 標記 end_us
  Enter       — 儲存區間到暫存清單
  P           — 擷取目前畫面
  Ctrl+S      — 立即將暫存標記寫入 CSV（不換片）
  Shift+S     — 顯示目前所有已暫存的標記區間
  Ctrl+Z      — 復原上一筆標記區間
  Backspace   — 回到上一部影片
  N           — 載入下一部影片（暫存標記保留）
  Ctrl+E      — 將暫存區間匯出為剪輯影片（ffmpeg）
  Ctrl+Q      — 退出

影片預覽縮放：
  + / =       — 放大
  -           — 縮小
  滑鼠滾輪    — 放大 / 縮小（以游標位置為中心）
  滑鼠左鍵拖曳 — 平移檢視（縮放後可檢視特定位置）
  0           — 還原縮放
"""

import csv
import glob
import json
import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from collections import OrderedDict
from tkinter import filedialog, messagebox

import cv2
from PIL import Image, ImageTk

# Windows high-DPI 適配
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

# Windows IME 控制（避免中文輸入法攔截按鍵事件造成快捷鍵卡頓）
_IMM32 = None
_USER32 = None
if sys.platform == "win32":
    try:
        import ctypes
        from ctypes import wintypes
        _IMM32 = ctypes.WinDLL("imm32", use_last_error=True)
        _IMM32.ImmAssociateContextEx.argtypes = [
            wintypes.HWND, wintypes.HANDLE, wintypes.DWORD]
        _IMM32.ImmAssociateContextEx.restype = wintypes.BOOL
        _USER32 = ctypes.WinDLL("user32", use_last_error=True)
        _USER32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
        _USER32.GetAncestor.restype = wintypes.HWND
    except Exception:
        _IMM32 = None
        _USER32 = None

_IACE_DEFAULT = 0x0010  # 恢復系統預設 IME context
_GA_ROOT = 2


def _ime_set(widget, enabled):
    """對指定 Tk widget 對應之 Win32 HWND 啟用或停用 IME。"""
    if _IMM32 is None or _USER32 is None:
        return
    try:
        widget.update_idletasks()
        hwnd = widget.winfo_id()
        if not hwnd:
            return
        flags = _IACE_DEFAULT if enabled else 0
        _IMM32.ImmAssociateContextEx(hwnd, None, flags)
        top = _USER32.GetAncestor(hwnd, _GA_ROOT)
        if top and top != hwnd:
            _IMM32.ImmAssociateContextEx(top, None, flags)
    except Exception:
        pass

FRAME_CACHE_MAX = 32
DEFAULT_CSV_NAME = "annotation.csv"
DEFAULT_IMAGE_FMT = "{filename}_{frame}_{us}.png"
DEFAULT_EXPORT_FMT = "{filename}_seg{segment}.mkv"

# ── 主題配色（科技風）─────────────────────────────────
FONT_BTN = ("Consolas", 10, "bold")
FONT_UI = ("Consolas", 10)
FONT_SM = ("Consolas", 9)

THEMES = {
    "light": {
        "nav_start": "#4facfe", "nav_end": "#00f2fe",
        "btn_bg": "#e8f4ff", "btn_fg": "#0a3d6b",
        "btn_hover": "#cce8ff", "btn_active": "#a0d4f8",
        "btn_outline": "#4facfe",
        "info_start": "#dceefb", "info_end": "#c4e0f5",
        "info_text": "#0e4a7a",
        "footer_start": "#8ec8f0", "footer_end": "#5aade0",
        "footer_text": "#0a3050",
        "video_bg": "#0c1018",
        "nav_text": "#ffffff",
        "slider_trough": "#a0d0f0",
        "slider_bg": "#c4e0f5",
        "slider_active": "#00b4f0",
        "toggle_label": "// NIGHT",
        "dlg_bg": "#dceefb", "dlg_fg": "#0e4a7a",
        "dlg_entry_bg": "#f0f8ff", "dlg_entry_fg": "#0a3d6b",
    },
    "dark": {
        "nav_start": "#0a0e27", "nav_end": "#1a0533",
        "btn_bg": "#101428", "btn_fg": "#00d4ff",
        "btn_hover": "#182040", "btn_active": "#1e2850",
        "btn_outline": "#00d4ff",
        "info_start": "#080c18", "info_end": "#0c1225",
        "info_text": "#5a9cc0",
        "footer_start": "#070a14", "footer_end": "#0a0e20",
        "footer_text": "#406888",
        "video_bg": "#04060c",
        "nav_text": "#00d4ff",
        "slider_trough": "#101830",
        "slider_bg": "#080c18",
        "slider_active": "#00d4ff",
        "toggle_label": "// DAY",
        "dlg_bg": "#0c1020", "dlg_fg": "#5a9cc0",
        "dlg_entry_bg": "#101428", "dlg_entry_fg": "#00d4ff",
    },
}


# ── 工具函式 ─────────────────────────────────────────

def _hex_rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _gradient(canvas, c1, c2, direction="horizontal"):
    """在 Canvas 上繪製漸層，以 'grad' tag 標記。"""
    canvas.delete("grad")
    w, h = canvas.winfo_width(), canvas.winfo_height()
    if w <= 1 or h <= 1:
        return
    r1, g1, b1 = _hex_rgb(c1)
    r2, g2, b2 = _hex_rgb(c2)
    if direction == "horizontal":
        n = min(w, 256)
        s = w / n
        for i in range(n):
            t = i / max(n - 1, 1)
            c = (f"#{int(r1+(r2-r1)*t):02x}"
                 f"{int(g1+(g2-g1)*t):02x}"
                 f"{int(b1+(b2-b1)*t):02x}")
            canvas.create_rectangle(
                int(i * s), 0, int((i + 1) * s) + 1, h,
                fill=c, outline="", tags="grad")
    else:
        n = min(h, 256)
        s = h / n
        for i in range(n):
            t = i / max(n - 1, 1)
            c = (f"#{int(r1+(r2-r1)*t):02x}"
                 f"{int(g1+(g2-g1)*t):02x}"
                 f"{int(b1+(b2-b1)*t):02x}")
            canvas.create_rectangle(
                0, int(i * s), w, int((i + 1) * s) + 1,
                fill=c, outline="", tags="grad")
    canvas.tag_lower("grad")


# ── 圓角按鈕 ─────────────────────────────────────────

class RoundedButton(tk.Canvas):
    """以 Canvas 繪製的圓角按鈕（含發光邊框）。"""

    def __init__(self, parent, text="", command=None, radius=14,
                 bg="#fff", fg="#333", hover_bg="#eee", active_bg="#ddd",
                 outline_color="", outer_bg=None,
                 font=FONT_BTN, padx=20, pady=7, **kwargs):
        self._bg, self._fg = bg, fg
        self._hover, self._active = hover_bg, active_bg
        self._outline = outline_color
        self._radius = radius
        self._command = command
        self._text = text
        self._font = font

        _tmp = tk.Label(parent, text=text, font=font)
        _tmp.update_idletasks()
        tw, th = _tmp.winfo_reqwidth(), _tmp.winfo_reqheight()
        _tmp.destroy()

        self._bw = tw + padx * 2
        self._bh = th + pady * 2

        super().__init__(parent, width=self._bw, height=self._bh,
                         highlightthickness=0, bd=0,
                         bg=outer_bg or bg, cursor="hand2", **kwargs)
        self._paint(self._bg)
        self.bind("<Enter>", lambda e: self._paint(self._hover))
        self.bind("<Leave>", lambda e: self._paint(self._bg))
        self.bind("<ButtonPress-1>", lambda e: self._paint(self._active))
        self.bind("<ButtonRelease-1>", self._click)

    def _rrect(self, x1, y1, x2, y2, r, **kw):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
               x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
               x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return self.create_polygon(pts, smooth=True, **kw)

    def _paint(self, fill):
        self.delete("all")
        if self._outline:
            self._rrect(1, 1, self._bw - 1, self._bh - 1,
                        self._radius, fill="", outline=self._outline,
                        width=2)
        self._rrect(3, 3, self._bw - 3, self._bh - 3,
                    max(self._radius - 2, 4), fill=fill, outline="")
        self.create_text(self._bw / 2, self._bh / 2,
                         text=self._text, font=self._font, fill=self._fg)

    def _click(self, event):
        self._paint(self._hover)
        if self._command:
            self._command()

    def recolor(self, bg, fg, hover, active, outer=None, outline=""):
        self._bg, self._fg = bg, fg
        self._hover, self._active = hover, active
        if outline:
            self._outline = outline
        if outer:
            self.configure(bg=outer)
        self._paint(self._bg)

    def set_text(self, text):
        self._text = text
        self._paint(self._bg)


# ── 輔助 ────────────────────────────────────────────

def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ── 主程式 ───────────────────────────────────────────

class Annotator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Azure Kinect Annotator")
        self.root.geometry("1080x760")
        self.root.minsize(800, 600)

        self.input_folder = ""
        self.output_folder = ""
        self.csv_name = DEFAULT_CSV_NAME
        self.image_format = DEFAULT_IMAGE_FMT
        self.export_format = DEFAULT_EXPORT_FMT
        self.files = []
        self.file_idx = 0
        self.cap = None
        self.total_frames = 0
        self.frame_no = 0
        self.start_us = None
        self.end_us = None
        self.segments = []
        self.photo = None
        self._slider_lock = False
        self._cached_us = 0
        self._cached_frame = None
        self._last_read_pos = -1
        self._frame_cache = OrderedDict()
        self._step_after_id = None
        self._slider_after_id = None
        self._resize_after_id = None
        self.theme_name = "light"

        # 縮放與平移狀態
        self.zoom = 1.0
        self.zoom_min = 1.0
        self.zoom_max = 16.0
        self.pan_cx = 0.5
        self.pan_cy = 0.5
        self._drag_start = None
        self._drag_start_pan = None
        self._last_fit_scale = 1.0

        self._load_config()
        self._build_ui()
        self._bind_keys()
        self._update_footer()
        # 鎖定主視窗 IME，避免中文輸入法攔截按鍵造成快捷鍵卡頓
        _ime_set(self.root, enabled=False)
        self._auto_load_input()
        self.root.mainloop()

    # ── Config ──────────────────────────────────────

    def _config_path(self):
        return os.path.join(get_app_dir(), "config.json")

    def _load_config(self):
        path = self._config_path()
        if not os.path.exists(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.input_folder = cfg.get("input_folder", "")
            self.output_folder = cfg.get("output_folder", "")
            self.csv_name = cfg.get("csv_name", DEFAULT_CSV_NAME)
            self.image_format = cfg.get("image_format", DEFAULT_IMAGE_FMT)
            self.export_format = cfg.get("export_format", DEFAULT_EXPORT_FMT)
            _theme = cfg.get("theme", "light")
            self.theme_name = _theme if _theme in THEMES else "light"
        except Exception:
            pass

    def _save_config(self):
        cfg = {
            "input_folder": self.input_folder,
            "output_folder": self.output_folder,
            "csv_name": self.csv_name,
            "image_format": self.image_format,
            "export_format": self.export_format,
            "theme": self.theme_name,
        }
        try:
            with open(self._config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _auto_load_input(self):
        if not self.input_folder or not os.path.isdir(self.input_folder):
            return
        self.files = sorted(glob.glob(os.path.join(self.input_folder, "*.mkv")))
        if not self.files:
            return
        self.file_idx = 0
        self.root.after(100, self._load_current_video)

    @property
    def _t(self):
        return THEMES[self.theme_name]

    # ── UI layout ───────────────────────────────────

    def _build_ui(self):
        t = self._t

        # -- Footer（最底部，最先 pack）--
        self.footer_cv = tk.Canvas(
            self.root, height=28, highlightthickness=0, bd=0)
        self.footer_cv.pack(fill=tk.X, side=tk.BOTTOM)
        self._ftr_txt = self.footer_cv.create_text(
            12, 14, anchor="w", font=FONT_SM,
            fill=t["footer_text"], text="")
        self.footer_cv.bind("<Configure>", lambda e: self._grad_footer())

        # -- Info panel --
        self.info_cv = tk.Canvas(
            self.root, height=62, highlightthickness=0, bd=0)
        self.info_cv.pack(fill=tk.X, side=tk.BOTTOM)
        self._info_txt = self.info_cv.create_text(
            12, 6, anchor="nw", font=FONT_UI,
            fill=t["info_text"],
            text="請先選擇 Input Folder 載入 .mkv 檔案")
        self.info_cv.bind("<Configure>", lambda e: self._grad_info())

        # -- Timeline slider --
        self.slider = tk.Scale(
            self.root, from_=0, to=0, orient=tk.HORIZONTAL,
            showvalue=False, command=self._on_slider,
            troughcolor=t["slider_trough"], bg=t["slider_bg"],
            activebackground=t["slider_active"],
            highlightthickness=0, bd=0, sliderrelief=tk.FLAT,
            sliderlength=20, width=12, cursor="hand2",
        )
        self.slider.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=(4, 2))

        # -- Nav bar --
        self.nav_cv = tk.Canvas(
            self.root, height=54, highlightthickness=0, bd=0)
        self.nav_cv.pack(fill=tk.X, side=tk.TOP)

        nc1, nc2 = t["nav_start"], t["nav_end"]
        ol = t["btn_outline"]

        self.btn_input = RoundedButton(
            self.nav_cv, "Input Folder", self._select_input,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=ol, outer_bg=nc1)
        self.btn_output = RoundedButton(
            self.nav_cv, "Output Folder", self._select_output,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=ol, outer_bg=nc1)
        self.btn_settings = RoundedButton(
            self.nav_cv, "Settings", self._open_settings,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=ol, outer_bg=nc1)
        self.btn_export = RoundedButton(
            self.nav_cv, "Export", self._export_segments,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=ol, outer_bg=nc1)
        self.btn_export_all = RoundedButton(
            self.nav_cv, "Export All", self._export_all,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=ol, outer_bg=nc1)
        self.btn_theme = RoundedButton(
            self.nav_cv, t["toggle_label"], self._toggle_theme,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=ol, outer_bg=nc2)

        self._nw = []
        for btn in (self.btn_input, self.btn_output, self.btn_settings,
                     self.btn_export, self.btn_export_all):
            self._nw.append(
                self.nav_cv.create_window(0, 10, window=btn, anchor="nw"))
        self._nw_toggle = self.nav_cv.create_window(
            0, 10, window=self.btn_theme, anchor="ne")
        self._nav_ftxt = self.nav_cv.create_text(
            0, 27, anchor="w", font=FONT_UI,
            fill=t["nav_text"], text="")

        self.nav_cv.bind("<Configure>", lambda e: self._on_nav_cfg())

        # -- Video display --
        self.vframe = tk.Frame(self.root, bg=t["video_bg"])
        self.vframe.pack(fill=tk.BOTH, expand=True)
        self.vframe.pack_propagate(False)
        self.canvas = tk.Label(self.vframe, bg=t["video_bg"])
        self.canvas.pack(expand=True)
        self.vframe.bind("<Configure>", self._on_vframe_resize)

        # 初始繪製漸層
        self.root.update_idletasks()
        self._grad_nav()
        self._grad_info()
        self._grad_footer()

    # ── Gradient 繪製 ───────────────────────────────

    def _grad_nav(self):
        _gradient(self.nav_cv, self._t["nav_start"], self._t["nav_end"])

    def _grad_info(self):
        _gradient(self.info_cv, self._t["info_start"], self._t["info_end"])
        w = self.info_cv.winfo_width()
        if w > 24:
            self.info_cv.itemconfigure(self._info_txt, width=w - 24)

    def _grad_footer(self):
        _gradient(self.footer_cv,
                  self._t["footer_start"], self._t["footer_end"])

    def _on_nav_cfg(self):
        self._repos_nav()
        self._grad_nav()

    def _repos_nav(self):
        """重新排列導覽列按鈕位置（RWD）。"""
        x = 12
        btns = (self.btn_input, self.btn_output, self.btn_settings,
                self.btn_export, self.btn_export_all)
        for i, btn in enumerate(btns):
            btn.update_idletasks()
            self.nav_cv.coords(self._nw[i], x, 10)
            x += btn.winfo_reqwidth() + 8
        self.nav_cv.coords(self._nav_ftxt, x + 8, 27)
        nw = self.nav_cv.winfo_width()
        self.nav_cv.coords(self._nw_toggle, nw - 12, 10)

    # ── Theme 切換 ──────────────────────────────────

    def _toggle_theme(self):
        self.theme_name = "dark" if self.theme_name == "light" else "light"
        self._apply_theme()
        self._save_config()

    def _apply_theme(self):
        t = self._t
        nc1, nc2 = t["nav_start"], t["nav_end"]

        ol = t["btn_outline"]
        for btn, outer in [(self.btn_input, nc1),
                           (self.btn_output, nc1),
                           (self.btn_settings, nc1),
                           (self.btn_export, nc1),
                           (self.btn_export_all, nc1),
                           (self.btn_theme, nc2)]:
            btn.recolor(t["btn_bg"], t["btn_fg"],
                        t["btn_hover"], t["btn_active"], outer, ol)
        self.btn_theme.set_text(t["toggle_label"])

        self.nav_cv.itemconfigure(self._nav_ftxt, fill=t["nav_text"])
        self.info_cv.itemconfigure(self._info_txt, fill=t["info_text"])
        self.footer_cv.itemconfigure(self._ftr_txt, fill=t["footer_text"])

        self.slider.configure(
            troughcolor=t["slider_trough"], bg=t["slider_bg"],
            activebackground=t["slider_active"])

        self.vframe.configure(bg=t["video_bg"])
        self.canvas.configure(bg=t["video_bg"])

        self._grad_nav()
        self._grad_info()
        self._grad_footer()

    # ── Key bindings ────────────────────────────────

    def _bind_keys(self):
        for key, cb in [
            ("<a>", lambda e: self._step(-1)),
            ("<d>", lambda e: self._step(1)),
            ("<w>", lambda e: self._step(-30)),
            ("<s>", lambda e: self._step(30)),
            ("<A>", lambda e: self._jump_start()),
            ("<D>", lambda e: self._jump_end()),
            ("<Left>", lambda e: self._step(-1)),
            ("<Right>", lambda e: self._step(1)),
            ("<Shift-Left>", lambda e: self._jump_start()),
            ("<Shift-Right>", lambda e: self._jump_end()),
            ("<i>", lambda e: self._mark_start()),
            ("<o>", lambda e: self._mark_end()),
            ("<Return>", lambda e: self._save_segment()),
            ("<p>", lambda e: self._capture_frame()),
            ("<n>", lambda e: self._next_video()),
            ("<BackSpace>", lambda e: self._prev_video()),
            ("<Control-s>", lambda e: self._ctrl_save()),
            ("<S>", lambda e: self._show_segments()),
            ("<Control-z>", lambda e: self._undo_segment()),
            ("<Control-e>", lambda e: self._export_segments()),
            ("<Control-q>", lambda e: self._quit()),
            ("<Escape>", lambda e: self._quit()),
            ("<plus>", lambda e: self._zoom_step(1.2)),
            ("<KP_Add>", lambda e: self._zoom_step(1.2)),
            ("<equal>", lambda e: self._zoom_step(1.2)),
            ("<minus>", lambda e: self._zoom_step(1 / 1.2)),
            ("<KP_Subtract>", lambda e: self._zoom_step(1 / 1.2)),
            ("<Key-0>", lambda e: self._zoom_reset()),
            ("<KP_0>", lambda e: self._zoom_reset()),
        ]:
            self.root.bind(key, cb)

        # 影片預覽滑鼠事件（縮放 / 拖曳）
        for widget in (self.vframe, self.canvas):
            widget.bind("<MouseWheel>", self._on_mousewheel)
            widget.bind("<Button-4>", self._on_mousewheel)  # Linux 上滾
            widget.bind("<Button-5>", self._on_mousewheel)  # Linux 下滾
            widget.bind("<ButtonPress-1>", self._on_drag_start)
            widget.bind("<B1-Motion>", self._on_drag_move)
            widget.bind("<ButtonRelease-1>", self._on_drag_end)

    # ── Settings dialog ─────────────────────────────

    def _open_settings(self):
        t = self._t
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.configure(bg=t["dlg_bg"])
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        # 在 Settings 視窗啟用 IME（允許中文輸入），關閉時自動恢復主視窗鎖定
        _ime_set(win, enabled=True)
        win.bind(
            "<Destroy>",
            lambda e, w=win: (
                _ime_set(self.root, enabled=False)
                if e.widget is w else None
            ),
        )

        lbl_kw = {"bg": t["dlg_bg"], "fg": t["dlg_fg"], "font": FONT_UI}
        ent_kw = {"bg": t["dlg_entry_bg"], "fg": t["dlg_entry_fg"],
                  "insertbackground": t["dlg_fg"],
                  "font": FONT_UI, "relief": tk.FLAT, "bd": 4}
        pad = {"padx": 8, "pady": 5}

        # Output folder
        tk.Label(win, text="Output Folder:", anchor=tk.W,
                 **lbl_kw).grid(row=0, column=0, sticky=tk.W, **pad)
        out_var = tk.StringVar(value=self.output_folder)
        tk.Entry(win, textvariable=out_var, width=48,
                 **ent_kw).grid(row=0, column=1, **pad)

        def browse_out():
            d = filedialog.askdirectory(
                parent=win, initialdir=self.output_folder or None)
            if d:
                out_var.set(d)

        RoundedButton(
            win, "...", browse_out, radius=10, padx=8, pady=4,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=t["btn_outline"],
            outer_bg=t["dlg_bg"],
        ).grid(row=0, column=2, **pad)

        # CSV name
        tk.Label(win, text="CSV Filename:", anchor=tk.W,
                 **lbl_kw).grid(row=1, column=0, sticky=tk.W, **pad)
        csv_var = tk.StringVar(value=self.csv_name)
        tk.Entry(win, textvariable=csv_var, width=48,
                 **ent_kw).grid(row=1, column=1, **pad)

        # Image format
        tk.Label(win, text="Image Filename:", anchor=tk.W,
                 **lbl_kw).grid(row=2, column=0, sticky=tk.W, **pad)
        img_var = tk.StringVar(value=self.image_format)
        tk.Entry(win, textvariable=img_var, width=48,
                 **ent_kw).grid(row=2, column=1, **pad)

        # Export format
        tk.Label(win, text="Export Filename:", anchor=tk.W,
                 **lbl_kw).grid(row=3, column=0, sticky=tk.W, **pad)
        exp_var = tk.StringVar(value=self.export_format)
        tk.Entry(win, textvariable=exp_var, width=48,
                 **ent_kw).grid(row=3, column=1, **pad)

        # Help
        tk.Label(
            win, text="Variables:  {filename}  {frame}  {us}  {segment}",
            font=FONT_SM, fg=t["dlg_fg"], bg=t["dlg_bg"],
        ).grid(row=4, column=0, columnspan=3, sticky=tk.W, **pad)

        # Preview
        preview_var = tk.StringVar()

        def update_preview(*_):
            lines = []
            try:
                csv_name = csv_var.get().format(filename="example")
            except (KeyError, ValueError):
                csv_name = "(format error)"
            lines.append(f"CSV:  {csv_name}")
            try:
                img_name = img_var.get().format(
                    filename="example", frame=120, us=4000000)
            except (KeyError, ValueError):
                img_name = "(format error)"
            lines.append(f"Image:  {img_name}")
            try:
                exp_name = exp_var.get().format(
                    filename="example", segment=1)
            except (KeyError, ValueError):
                exp_name = "(format error)"
            lines.append(f"Export:  {exp_name}")
            preview_var.set("\n".join(lines))

        csv_var.trace_add("write", update_preview)
        img_var.trace_add("write", update_preview)
        exp_var.trace_add("write", update_preview)
        update_preview()

        tk.Label(
            win, textvariable=preview_var, justify=tk.LEFT,
            font=FONT_SM, fg=t["dlg_fg"], bg=t["dlg_bg"],
        ).grid(row=5, column=0, columnspan=3, sticky=tk.W, **pad)

        # Buttons
        btn_row = tk.Frame(win, bg=t["dlg_bg"])
        btn_row.grid(row=6, column=0, columnspan=3, pady=12)

        def save():
            self.output_folder = out_var.get()
            self.csv_name = csv_var.get() or DEFAULT_CSV_NAME
            self.image_format = img_var.get() or DEFAULT_IMAGE_FMT
            self.export_format = exp_var.get() or DEFAULT_EXPORT_FMT
            self._save_config()
            self._update_footer()
            win.destroy()

        RoundedButton(
            btn_row, "Save", save,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=t["btn_outline"], outer_bg=t["dlg_bg"],
        ).pack(side=tk.LEFT, padx=8)
        RoundedButton(
            btn_row, "Cancel", win.destroy,
            bg=t["btn_bg"], fg=t["btn_fg"],
            hover_bg=t["btn_hover"], active_bg=t["btn_active"],
            outline_color=t["btn_outline"], outer_bg=t["dlg_bg"],
        ).pack(side=tk.LEFT, padx=8)

    # ── Folder selection ────────────────────────────

    def _select_input(self):
        folder = filedialog.askdirectory(
            title="選擇 .mkv 檔案所在資料夾",
            initialdir=self.input_folder or None)
        if not folder:
            return
        self.input_folder = folder
        self.files = sorted(glob.glob(os.path.join(folder, "*.mkv")))
        if not self.files:
            messagebox.showwarning("提示", "該資料夾找不到 .mkv 檔案")
            return
        self.file_idx = 0
        if not self.output_folder:
            self.output_folder = folder
        self._save_config()
        self._update_footer()
        self._load_current_video()

    def _select_output(self):
        folder = filedialog.askdirectory(
            title="選擇輸出資料夾",
            initialdir=self.output_folder or None)
        if not folder:
            return
        self.output_folder = folder
        self._save_config()
        self._update_footer()

    def _update_footer(self):
        inp = self.input_folder or "(未設定)"
        out = self.output_folder or "(未設定)"
        self.footer_cv.itemconfigure(
            self._ftr_txt,
            text=f"Input: {inp}  |  Output: {out}")

    # ── Video loading ───────────────────────────────

    def _load_current_video(self):
        if self.cap:
            self.cap.release()
            self.cap = None
        self._cached_frame = None
        self._frame_cache.clear()
        self._last_read_pos = -1

        if self.file_idx >= len(self.files):
            messagebox.showinfo("完成", "所有影片皆已標記完畢")
            self.info_cv.itemconfigure(
                self._info_txt, text="所有影片皆已標記完畢")
            self.nav_cv.itemconfigure(self._nav_ftxt, text="")
            self.canvas.config(image="")
            return

        path = self.files[self.file_idx]
        self.cap = cv2.VideoCapture(path, cv2.CAP_FFMPEG)
        if not self.cap.isOpened():
            self.cap = cv2.VideoCapture(path)
        if not self.cap.isOpened():
            messagebox.showerror("錯誤", f"無法開啟: {path}")
            return

        self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
        self.frame_no = 0
        self.start_us = None
        self.end_us = None
        self.zoom = 1.0
        self.pan_cx = 0.5
        self.pan_cy = 0.5

        self.slider.configure(to=max(self.total_frames - 1, 0))
        self.nav_cv.itemconfigure(
            self._nav_ftxt,
            text=(f"{os.path.basename(path)}"
                  f"  ({self.file_idx + 1}/{len(self.files)})"))
        self._render_current()

    # ── Frame cache ─────────────────────────────────

    def _cache_put(self, frame_no, frame, us):
        self._frame_cache[frame_no] = (frame, us)
        if len(self._frame_cache) > FRAME_CACHE_MAX:
            self._frame_cache.popitem(last=False)

    def _cache_get(self, frame_no):
        if frame_no in self._frame_cache:
            self._frame_cache.move_to_end(frame_no)
            return self._frame_cache[frame_no]
        return None

    # ── Frame read ──────────────────────────────────

    def _read_frame_at(self, frame_no):
        cached = self._cache_get(frame_no)
        if cached:
            return cached

        if not self.cap:
            return None

        if self._last_read_pos == frame_no - 1:
            ret, frame = self.cap.read()
            if ret:
                us = int(self.cap.get(cv2.CAP_PROP_POS_MSEC) * 1000)
                self._last_read_pos = frame_no
                self._cache_put(frame_no, frame, us)
                return (frame, us)

        self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame_no)
        ret, frame = self.cap.read()
        if ret:
            us = int(self.cap.get(cv2.CAP_PROP_POS_MSEC) * 1000)
            self._last_read_pos = frame_no
            self._cache_put(frame_no, frame, us)
            return (frame, us)
        return None

    def _render_current(self):
        result = self._read_frame_at(self.frame_no)
        if result:
            self._cached_frame, self._cached_us = result
            self._display()

    # ── Display ─────────────────────────────────────

    def _display(self):
        if self._cached_frame is None:
            return

        self._slider_lock = True
        self.slider.set(self.frame_no)
        self._slider_lock = False

        s = self.start_us if self.start_us is not None else "---"
        e = self.end_us if self.end_us is not None else "---"
        zoom_txt = f"{self.zoom:.2f}x" if self.zoom > 1.0 + 1e-6 else "1.00x"
        self.info_cv.itemconfigure(self._info_txt, text=(
            f"Frame: {self.frame_no}/{self.total_frames - 1}  |  "
            f"Time: {self._cached_us} us  |  "
            f"Start: {s}  |  End: {e}  |  "
            f"Segments: {len(self.segments)}  |  Zoom: {zoom_txt}\n"
            f"[A/D/Arrow]+-1  [W/S]+-30  [Shift+A/D]Head/Tail  "
            f"[I]Start  [O]End  [Enter]Save  [P]Capture  "
            f"[+/-/Wheel]Zoom  [Drag]Pan  [0]Reset  "
            f"[Ctrl+Z]Undo  [Ctrl+S]Write  [Shift+S]View  "
            f"[BS]Prev  [N]Next  [Ctrl+Q]Quit"))

        cw = self.vframe.winfo_width()
        ch = self.vframe.winfo_height()
        if cw <= 1 or ch <= 1:
            return

        frame = self._cached_frame
        ih, iw = frame.shape[:2]

        fit_scale = min(cw / iw, ch / ih)
        self._last_fit_scale = fit_scale
        disp_w = max(1, int(iw * fit_scale))
        disp_h = max(1, int(ih * fit_scale))

        if self.zoom <= 1.0 + 1e-6:
            self.zoom = 1.0
            self.pan_cx = 0.5
            self.pan_cy = 0.5
            src = frame
        else:
            roi_w = iw / self.zoom
            roi_h = ih / self.zoom
            cx = self.pan_cx * iw
            cy = self.pan_cy * ih
            x1 = cx - roi_w / 2
            y1 = cy - roi_h / 2
            if x1 < 0:
                x1 = 0
            if y1 < 0:
                y1 = 0
            if x1 + roi_w > iw:
                x1 = iw - roi_w
            if y1 + roi_h > ih:
                y1 = ih - roi_h
            self.pan_cx = (x1 + roi_w / 2) / iw
            self.pan_cy = (y1 + roi_h / 2) / ih

            ix1 = max(0, int(round(x1)))
            iy1 = max(0, int(round(y1)))
            ix2 = min(iw, int(round(x1 + roi_w)))
            iy2 = min(ih, int(round(y1 + roi_h)))
            if ix2 - ix1 < 1 or iy2 - iy1 < 1:
                return
            src = frame[iy1:iy2, ix1:ix2]

        disp_frame = cv2.resize(
            src, (disp_w, disp_h), interpolation=cv2.INTER_LINEAR)

        rgb = cv2.cvtColor(disp_frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(rgb)

        if (self.photo is not None
                and self.photo.width() == disp_w
                and self.photo.height() == disp_h):
            self.photo.paste(img)
        else:
            self.photo = ImageTk.PhotoImage(img)
            self.canvas.configure(image=self.photo)

    # ── 縮放 / 平移 ────────────────────────────────

    def _clamp_zoom(self, value):
        if value < self.zoom_min:
            return self.zoom_min
        if value > self.zoom_max:
            return self.zoom_max
        return value

    def _viewport_to_image(self, vx, vy):
        """將 viewport 座標換算為來源影像座標（None 表示不在影像範圍內）。"""
        if self._cached_frame is None:
            return None
        ih, iw = self._cached_frame.shape[:2]
        cw = self.vframe.winfo_width()
        ch = self.vframe.winfo_height()
        if cw <= 1 or ch <= 1:
            return None
        fit_scale = min(cw / iw, ch / ih)
        disp_w = iw * fit_scale
        disp_h = ih * fit_scale
        off_x = (cw - disp_w) / 2
        off_y = (ch - disp_h) / 2
        lx = vx - off_x
        ly = vy - off_y
        if lx < 0 or ly < 0 or lx > disp_w or ly > disp_h:
            return None
        # lx/ly 為 fit 後影像座標（0..disp_w / 0..disp_h）
        # 將其換算為縮放後 ROI 的影像座標
        roi_w = iw / self.zoom
        roi_h = ih / self.zoom
        cx = self.pan_cx * iw
        cy = self.pan_cy * ih
        rx1 = cx - roi_w / 2
        ry1 = cy - roi_h / 2
        if self.zoom > 1.0:
            if rx1 < 0:
                rx1 = 0
            if ry1 < 0:
                ry1 = 0
            if rx1 + roi_w > iw:
                rx1 = iw - roi_w
            if ry1 + roi_h > ih:
                ry1 = ih - roi_h
        img_x = rx1 + (lx / disp_w) * roi_w
        img_y = ry1 + (ly / disp_h) * roi_h
        return img_x, img_y

    def _zoom_to(self, new_zoom, anchor_vx=None, anchor_vy=None):
        new_zoom = self._clamp_zoom(new_zoom)
        if abs(new_zoom - self.zoom) < 1e-6:
            return

        if self._cached_frame is None:
            self.zoom = new_zoom
            return

        ih, iw = self._cached_frame.shape[:2]
        anchor_img = None
        if anchor_vx is not None and anchor_vy is not None:
            anchor_img = self._viewport_to_image(anchor_vx, anchor_vy)

        if anchor_img is None or new_zoom <= 1.0:
            self.zoom = new_zoom
        else:
            ax, ay = anchor_img
            # 以 anchor 為固定點調整 pan_cx/cy
            cw = self.vframe.winfo_width()
            ch = self.vframe.winfo_height()
            fit_scale = min(cw / iw, ch / ih)
            disp_w = iw * fit_scale
            disp_h = ih * fit_scale
            off_x = (cw - disp_w) / 2
            off_y = (ch - disp_h) / 2
            # anchor 在 viewport 中的正規化位置 (0..1) 相對顯示框
            nx = (anchor_vx - off_x) / disp_w
            ny = (anchor_vy - off_y) / disp_h
            nx = min(max(nx, 0.0), 1.0)
            ny = min(max(ny, 0.0), 1.0)
            new_roi_w = iw / new_zoom
            new_roi_h = ih / new_zoom
            new_x1 = ax - nx * new_roi_w
            new_y1 = ay - ny * new_roi_h
            self.pan_cx = (new_x1 + new_roi_w / 2) / iw
            self.pan_cy = (new_y1 + new_roi_h / 2) / ih
            self.zoom = new_zoom

        self._display()

    def _zoom_step(self, factor):
        self._zoom_to(self.zoom * factor)

    def _zoom_reset(self):
        self.zoom = 1.0
        self.pan_cx = 0.5
        self.pan_cy = 0.5
        self._display()

    def _on_mousewheel(self, event):
        if self._cached_frame is None:
            return
        delta = 0
        if getattr(event, "num", None) == 4:
            delta = 1
        elif getattr(event, "num", None) == 5:
            delta = -1
        elif getattr(event, "delta", 0) > 0:
            delta = 1
        elif getattr(event, "delta", 0) < 0:
            delta = -1
        if delta == 0:
            return
        factor = 1.2 if delta > 0 else (1 / 1.2)
        # 事件座標相對觸發 widget，轉為 vframe 座標
        vx = event.x_root - self.vframe.winfo_rootx()
        vy = event.y_root - self.vframe.winfo_rooty()
        self._zoom_to(self.zoom * factor, vx, vy)

    def _on_drag_start(self, event):
        if self._cached_frame is None or self.zoom <= 1.0 + 1e-6:
            self._drag_start = None
            return
        vx = event.x_root - self.vframe.winfo_rootx()
        vy = event.y_root - self.vframe.winfo_rooty()
        self._drag_start = (vx, vy)
        self._drag_start_pan = (self.pan_cx, self.pan_cy)
        try:
            self.canvas.configure(cursor="fleur")
        except tk.TclError:
            pass

    def _on_drag_move(self, event):
        if self._drag_start is None or self._cached_frame is None:
            return
        if self.zoom <= 1.0 + 1e-6:
            return
        ih, iw = self._cached_frame.shape[:2]
        cw = self.vframe.winfo_width()
        ch = self.vframe.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        fit_scale = min(cw / iw, ch / ih)
        effective = fit_scale * self.zoom
        if effective <= 0:
            return
        vx = event.x_root - self.vframe.winfo_rootx()
        vy = event.y_root - self.vframe.winfo_rooty()
        dx = vx - self._drag_start[0]
        dy = vy - self._drag_start[1]
        img_dx = dx / effective
        img_dy = dy / effective
        start_cx, start_cy = self._drag_start_pan
        self.pan_cx = start_cx - img_dx / iw
        self.pan_cy = start_cy - img_dy / ih
        self._display()

    def _on_drag_end(self, event):
        self._drag_start = None
        self._drag_start_pan = None
        try:
            self.canvas.configure(cursor="")
        except tk.TclError:
            pass

    # ── Navigation ──────────────────────────────────

    def _step(self, delta):
        if not self.cap:
            return
        new_no = max(0, min(self.frame_no + delta, self.total_frames - 1))
        if new_no == self.frame_no:
            return
        self.frame_no = new_no

        if self._step_after_id is not None:
            self.root.after_cancel(self._step_after_id)
        self._step_after_id = self.root.after(16, self._on_step_done)

    def _on_step_done(self):
        self._step_after_id = None
        self._render_current()

    def _jump_start(self):
        if not self.cap:
            return
        self.frame_no = 0
        self._render_current()

    def _jump_end(self):
        if not self.cap:
            return
        self.frame_no = max(self.total_frames - 2, 0)
        self._render_current()

    def _on_slider(self, val):
        if self._slider_lock or not self.cap:
            return
        self.frame_no = int(val)
        if self._slider_after_id is not None:
            self.root.after_cancel(self._slider_after_id)
        self._slider_after_id = self.root.after(30, self._render_current)

    def _on_vframe_resize(self, event):
        if self._cached_frame is None:
            return
        if self._resize_after_id is not None:
            self.root.after_cancel(self._resize_after_id)
        self._resize_after_id = self.root.after(50, self._display)

    # ── Marking ─────────────────────────────────────

    def _mark_start(self):
        if not self.cap:
            return
        self.start_us = self._cached_us
        self._display()

    def _mark_end(self):
        if not self.cap:
            return
        self.end_us = self._cached_us
        self._display()

    def _save_segment(self):
        if self.start_us is None or self.end_us is None:
            return
        fname = (os.path.basename(self.files[self.file_idx])
                 if self.files else "")
        self.segments.append((fname, self.start_us, self.end_us))
        self.start_us = None
        self.end_us = None
        self._display()

    # ── Capture frame ───────────────────────────────

    def _capture_frame(self):
        if self._cached_frame is None or not self.output_folder:
            return
        fname = os.path.splitext(
            os.path.basename(self.files[self.file_idx])
        )[0] if self.files else "frame"
        try:
            img_name = self.image_format.format(
                filename=fname, frame=self.frame_no, us=self._cached_us)
        except (KeyError, ValueError):
            img_name = f"{fname}_{self.frame_no}_{self._cached_us}.png"
        img_path = os.path.join(self.output_folder, img_name)
        cv2.imwrite(img_path, self._cached_frame)
        self.info_cv.itemconfigure(
            self._info_txt, text=f"Captured: {img_name}")

    # ── CSV write / view / undo ─────────────────────

    def _resolve_csv_path(self):
        fname = os.path.splitext(
            os.path.basename(self.files[self.file_idx])
        )[0] if self.files else ""
        try:
            csv_filename = self.csv_name.format(filename=fname)
        except (KeyError, ValueError):
            csv_filename = DEFAULT_CSV_NAME
        return os.path.join(self.output_folder, csv_filename)

    def _write_csv(self):
        if not self.output_folder or not self.segments or not self.files:
            return 0
        csv_path = self._resolve_csv_path()
        write_header = not os.path.exists(csv_path)

        with open(csv_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if write_header:
                writer.writerow(["filename", "segment", "start_us", "end_us"])
            seg_counter = {}
            for fn, s, e in self.segments:
                seg_counter[fn] = seg_counter.get(fn, 0) + 1
                writer.writerow([fn, seg_counter[fn], s, e])
        return len(self.segments)

    def _ctrl_save(self):
        if not self.segments:
            messagebox.showinfo("提示", "目前沒有暫存的標記區間")
            return
        count = self._write_csv()
        if count:
            messagebox.showinfo(
                "已儲存",
                f"已寫入 {count} 筆區間至\n{self._resolve_csv_path()}",
            )

    def _show_segments(self):
        if not self.segments:
            messagebox.showinfo("暫存標記", "目前沒有暫存的標記區間")
            return
        lines = []
        seg_counter = {}
        for fn, s, e in self.segments:
            seg_counter[fn] = seg_counter.get(fn, 0) + 1
            lines.append(f"  {fn} seg:{seg_counter[fn]}  |  {s}  ~  {e}  us")
        msg = f"共 {len(self.segments)} 段:\n\n" + "\n".join(lines)
        messagebox.showinfo("暫存標記一覽", msg)

    def _undo_segment(self):
        if not self.segments:
            return
        removed = self.segments.pop()
        self.start_us = removed[1]
        self.end_us = removed[2]
        self._display()

    # ── Export segments (ffmpeg) ────────────────────

    def _read_csv_segments(self):
        """從 CSV 檔讀取當前影片的標記區間，回傳 list of (start_us, end_us)。"""
        csv_path = self._resolve_csv_path()
        if not os.path.exists(csv_path):
            return []
        fname = os.path.basename(self.files[self.file_idx]) if self.files else ""
        segs = []
        with open(csv_path, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("filename") == fname:
                    segs.append((int(row["start_us"]), int(row["end_us"])))
        return segs

    def _find_mkvmerge(self):
        """尋找 mkvmerge 執行檔，回傳完整路徑或 None。"""
        found = shutil.which("mkvmerge")
        if found:
            return found
        # MKVToolNix 常見安裝位置
        for prog_dir in [os.environ.get("PROGRAMFILES", ""),
                         os.environ.get("PROGRAMFILES(X86)", "")]:
            candidate = os.path.join(prog_dir, "MKVToolNix", "mkvmerge.exe")
            if os.path.isfile(candidate):
                return candidate
        return None

    def _export_segments(self):
        mkvmerge_path = self._find_mkvmerge()
        if not mkvmerge_path:
            messagebox.showerror("錯誤",
                                 "找不到 mkvmerge，請安裝 MKVToolNix\n"
                                 "https://mkvtoolnix.download/")
            return
        if not self.files:
            return
        if not self.output_folder:
            messagebox.showinfo("提示", "請先設定 Output Folder")
            return

        # 優先使用 CSV 檔案的區間，若無 CSV 則使用暫存標記
        csv_segs = self._read_csv_segments()
        if csv_segs:
            cur = os.path.basename(self.files[self.file_idx])
            export_segs = [(cur, s, e) for s, e in csv_segs]
        elif self.segments:
            export_segs = list(self.segments)  # (filename, start_us, end_us)
        else:
            messagebox.showinfo("提示", "找不到 CSV 檔案，暫存標記也是空的，無法匯出")
            return

        # 建立檔名到完整路徑的對應
        file_map = {os.path.basename(f): f for f in self.files}

        # 建立進度視窗
        win = tk.Toplevel(self.root)
        win.title("匯出剪輯")
        win.configure(bg=self._t["dlg_bg"])
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.geometry("420x120")

        lbl = tk.Label(win, text="準備中...",
                       bg=self._t["dlg_bg"], fg=self._t["dlg_fg"],
                       font=FONT_UI)
        lbl.pack(pady=(16, 8))

        progress_lbl = tk.Label(win, text="",
                                bg=self._t["dlg_bg"], fg=self._t["dlg_fg"],
                                font=FONT_SM)
        progress_lbl.pack()

        segments = export_segs
        export_fmt = self.export_format
        results = {"ok": 0, "fail": 0, "errors": []}

        def _us_to_ts(us):
            """微秒轉為 HH:MM:SS.nnnnnnnnn 時間格式。"""
            total_s = us / 1_000_000
            h = int(total_s // 3600)
            m = int((total_s % 3600) // 60)
            s = total_s % 60
            return f"{h:02d}:{m:02d}:{s:012.9f}"

        def _run():
            seg_counter = {}
            for idx, (fn, s_us, e_us) in enumerate(segments, 1):
                seg_counter[fn] = seg_counter.get(fn, 0) + 1
                seg_no = seg_counter[fn]
                bn = os.path.splitext(fn)[0]
                src = file_map.get(fn, "")

                try:
                    out_name = export_fmt.format(
                        filename=bn, segment=seg_no)
                except (KeyError, ValueError):
                    out_name = f"{bn}_seg{seg_no}.mkv"

                def _update_label(_idx=idx, _name=out_name):
                    try:
                        lbl.configure(
                            text=f"正在匯出第 {_idx}/{len(segments)} 段...")
                        progress_lbl.configure(text=_name)
                    except tk.TclError:
                        pass
                try:
                    win.after(0, _update_label)
                except tk.TclError:
                    pass

                if not src or not os.path.isfile(src):
                    results["fail"] += 1
                    results["errors"].append(
                        f"{fn} seg{seg_no}: 找不到來源影片")
                    continue
                out_path = os.path.join(self.output_folder, out_name)

                ss = _us_to_ts(s_us)
                to = _us_to_ts(e_us)

                # 使用 mkvmerge --split parts 裁剪，完整保留所有軌道與 codec tag
                cmd = [
                    mkvmerge_path,
                    "-o", out_path,
                    "--split", f"parts:{ss}-{to}",
                    src,
                ]
                try:
                    subprocess.run(
                        cmd, check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        creationflags=(subprocess.CREATE_NO_WINDOW
                                       if sys.platform == "win32" else 0),
                    )
                    results["ok"] += 1
                except subprocess.CalledProcessError as exc:
                    results["fail"] += 1
                    output = (exc.stdout or exc.stderr or b"").decode(
                        errors="replace")[-500:]
                    results["errors"].append(
                        f"{fn} seg{seg_no}: {output}")

            def _done():
                try:
                    win.destroy()
                except tk.TclError:
                    pass
                msg = f"成功匯出 {results['ok']}/{len(segments)} 段"
                if results["fail"]:
                    msg += f"\n失敗 {results['fail']} 段"
                    for e in results["errors"]:
                        msg += f"\n  {e}"
                msg += f"\n\n輸出目錄: {self.output_folder}"
                messagebox.showinfo("匯出完成", msg)

            try:
                self.root.after(0, _done)
            except tk.TclError:
                pass

        threading.Thread(target=_run, daemon=True).start()

    # ── Export All (from CSV) ───────────────────────

    def _read_all_csv_segments(self):
        """讀取 output_folder 中所有 CSV 檔的標記區間，
        回傳 list of (filename, start_us, end_us)。"""
        if not self.output_folder:
            return []
        all_segs = []
        for csv_file in glob.glob(os.path.join(self.output_folder, "*.csv")):
            with open(csv_file, "r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    fn = row.get("filename", "")
                    s = row.get("start_us", "")
                    e = row.get("end_us", "")
                    if fn and s and e:
                        all_segs.append((fn, int(s), int(e)))
        return all_segs

    def _export_all(self):
        mkvmerge_path = self._find_mkvmerge()
        if not mkvmerge_path:
            messagebox.showerror("錯誤",
                                 "找不到 mkvmerge，請安裝 MKVToolNix\n"
                                 "https://mkvtoolnix.download/")
            return
        if not self.output_folder:
            messagebox.showinfo("提示", "請先設定 Output Folder")
            return

        all_segs = self._read_all_csv_segments()
        if not all_segs:
            messagebox.showinfo("提示", "CSV 檔案中沒有任何標記區間")
            return

        # 建立檔名到完整路徑的對應
        file_map = {os.path.basename(f): f for f in self.files}

        # 建立進度視窗
        win = tk.Toplevel(self.root)
        win.title("匯出全部剪輯")
        win.configure(bg=self._t["dlg_bg"])
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.geometry("420x120")

        lbl = tk.Label(win, text="準備中...",
                       bg=self._t["dlg_bg"], fg=self._t["dlg_fg"],
                       font=FONT_UI)
        lbl.pack(pady=(16, 8))

        progress_lbl = tk.Label(win, text="",
                                bg=self._t["dlg_bg"], fg=self._t["dlg_fg"],
                                font=FONT_SM)
        progress_lbl.pack()

        segments = all_segs
        export_fmt = self.export_format
        results = {"ok": 0, "fail": 0, "errors": []}

        def _us_to_ts(us):
            total_s = us / 1_000_000
            h = int(total_s // 3600)
            m = int((total_s % 3600) // 60)
            s = total_s % 60
            return f"{h:02d}:{m:02d}:{s:012.9f}"

        def _run():
            seg_counter = {}
            for idx, (fn, s_us, e_us) in enumerate(segments, 1):
                seg_counter[fn] = seg_counter.get(fn, 0) + 1
                seg_no = seg_counter[fn]
                bn = os.path.splitext(fn)[0]
                src = file_map.get(fn, "")

                try:
                    out_name = export_fmt.format(
                        filename=bn, segment=seg_no)
                except (KeyError, ValueError):
                    out_name = f"{bn}_seg{seg_no}.mkv"

                def _update_label(_idx=idx, _name=out_name):
                    try:
                        lbl.configure(
                            text=f"正在匯出第 {_idx}/{len(segments)} 段...")
                        progress_lbl.configure(text=_name)
                    except tk.TclError:
                        pass
                try:
                    win.after(0, _update_label)
                except tk.TclError:
                    pass

                if not src or not os.path.isfile(src):
                    results["fail"] += 1
                    results["errors"].append(
                        f"{fn} seg{seg_no}: 找不到來源影片")
                    continue
                out_path = os.path.join(self.output_folder, out_name)

                ss = _us_to_ts(s_us)
                to = _us_to_ts(e_us)

                cmd = [
                    mkvmerge_path,
                    "-o", out_path,
                    "--split", f"parts:{ss}-{to}",
                    src,
                ]
                try:
                    subprocess.run(
                        cmd, check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        creationflags=(subprocess.CREATE_NO_WINDOW
                                       if sys.platform == "win32" else 0),
                    )
                    results["ok"] += 1
                except subprocess.CalledProcessError as exc:
                    results["fail"] += 1
                    output = (exc.stdout or exc.stderr or b"").decode(
                        errors="replace")[-500:]
                    results["errors"].append(
                        f"{fn} seg{seg_no}: {output}")

            def _done():
                try:
                    win.destroy()
                except tk.TclError:
                    pass
                msg = f"成功匯出 {results['ok']}/{len(segments)} 段"
                if results["fail"]:
                    msg += f"\n失敗 {results['fail']} 段"
                    for e in results["errors"]:
                        msg += f"\n  {e}"
                msg += f"\n\n輸出目錄: {self.output_folder}"
                messagebox.showinfo("匯出全部完成", msg)

            try:
                self.root.after(0, _done)
            except tk.TclError:
                pass

        threading.Thread(target=_run, daemon=True).start()

    # ── Next / Prev video ──────────────────────────

    def _next_video(self):
        if not self.cap or not self.output_folder:
            return
        if self.file_idx >= len(self.files) - 1:
            messagebox.showinfo("提示", "這是最後一部影片")
            return
        self.file_idx += 1
        self._load_current_video()

    def _prev_video(self):
        if self.file_idx <= 0 or not self.files:
            return
        self.file_idx -= 1
        self._load_current_video()

    def _quit(self):
        if self.cap:
            self.cap.release()
        self.root.destroy()


if __name__ == "__main__":
    Annotator()
