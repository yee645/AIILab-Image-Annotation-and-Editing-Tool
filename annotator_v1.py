"""Azure Kinect 影片標記工具

快捷鍵：
  A / D       — 後退 / 前進 1 幀
  W / S       — 後退 / 前進 30 幀
  Shift+A/D   — 跳到影片開頭 / 結尾
  Left/Right  — 等同 A/D（後退 / 前進 1 幀）
  Shift+Left/Right — 跳到影片開頭 / 結尾
  J           — 標記 start_us
  K           — 標記 end_us
  Enter       — 儲存區間到暫存清單
  P           — 擷取目前畫面
  Ctrl+S      — 立即將暫存標記寫入 CSV（不換片）
  Shift+S     — 顯示目前所有已暫存的標記區間
  Ctrl+Z      — 復原上一筆標記區間
  Backspace   — 回到上一部影片
  N           — 寫入 CSV 並載入下一部影片
  Ctrl+Q      — 退出
"""

import csv
import glob
import json
import os
import sys
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

FRAME_CACHE_MAX = 32
DEFAULT_CSV_NAME = "annotation.csv"
DEFAULT_IMAGE_FMT = "{filename}_{frame}_{us}.png"


def get_app_dir():
    """取得應用程式所在目錄（相容 PyInstaller 打包）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class Annotator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Azure Kinect Annotator")
        self.root.geometry("1080x760")
        self.root.minsize(800, 600)

        # State
        self.input_folder = ""
        self.output_folder = ""
        self.csv_name = DEFAULT_CSV_NAME
        self.image_format = DEFAULT_IMAGE_FMT
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

        self._load_config()
        self._build_ui()
        self._bind_keys()
        self._update_footer()
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
        except Exception:
            pass

    def _save_config(self):
        cfg = {
            "input_folder": self.input_folder,
            "output_folder": self.output_folder,
            "csv_name": self.csv_name,
            "image_format": self.image_format,
        }
        try:
            with open(self._config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _auto_load_input(self):
        """啟動時若 config 中有有效的 input_folder，自動載入影片。"""
        if not self.input_folder or not os.path.isdir(self.input_folder):
            return
        self.files = sorted(glob.glob(os.path.join(self.input_folder, "*.mkv")))
        if not self.files:
            return
        self.file_idx = 0
        self.root.after(100, self._load_current_video)

    # ── UI layout ───────────────────────────────────

    def _build_ui(self):
        # Footer（最底部，最先 pack）
        self.footer_var = tk.StringVar()
        tk.Label(
            self.root, textvariable=self.footer_var,
            anchor=tk.W, font=("Consolas", 9),
            bg="#e0e0e0", relief=tk.SUNKEN, padx=8, pady=3,
        ).pack(fill=tk.X, side=tk.BOTTOM)

        # Info panel
        self.info_var = tk.StringVar(value="請先選擇 Input Folder 載入 .mkv 檔案")
        tk.Label(
            self.root, textvariable=self.info_var,
            anchor=tk.W, justify=tk.LEFT, font=("Consolas", 10),
            padx=8, pady=4,
        ).pack(fill=tk.X, side=tk.BOTTOM)

        # Timeline slider
        self.slider = tk.Scale(
            self.root, from_=0, to=0, orient=tk.HORIZONTAL,
            showvalue=False, command=self._on_slider,
        )
        self.slider.pack(fill=tk.X, side=tk.BOTTOM, padx=8)

        # Top navigation bar
        nav = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        nav.pack(fill=tk.X, side=tk.TOP)
        tk.Button(nav, text="Input Folder", command=self._select_input).pack(
            side=tk.LEFT, padx=4, pady=4)
        tk.Button(nav, text="Output Folder", command=self._select_output).pack(
            side=tk.LEFT, padx=4, pady=4)
        tk.Button(nav, text="Settings", command=self._open_settings).pack(
            side=tk.LEFT, padx=4, pady=4)
        self.file_label = tk.Label(nav, text="", font=("Consolas", 10))
        self.file_label.pack(side=tk.LEFT, padx=12)

        # Video display
        self.vframe = tk.Frame(self.root, bg="black")
        self.vframe.pack(fill=tk.BOTH, expand=True)
        self.vframe.pack_propagate(False)
        self.canvas = tk.Label(self.vframe, bg="black")
        self.canvas.pack(expand=True)
        self.vframe.bind("<Configure>", self._on_vframe_resize)

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
            ("<j>", lambda e: self._mark_start()),
            ("<k>", lambda e: self._mark_end()),
            ("<Return>", lambda e: self._save_segment()),
            ("<p>", lambda e: self._capture_frame()),
            ("<n>", lambda e: self._next_video()),
            ("<BackSpace>", lambda e: self._prev_video()),
            ("<Control-s>", lambda e: self._ctrl_save()),
            ("<S>", lambda e: self._show_segments()),
            ("<Control-z>", lambda e: self._undo_segment()),
            ("<Control-q>", lambda e: self._quit()),
            ("<Escape>", lambda e: self._quit()),
        ]:
            self.root.bind(key, cb)

    # ── Settings dialog ─────────────────────────────

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        pad = {"padx": 8, "pady": 5}

        # Output folder
        tk.Label(win, text="Output Folder:", anchor=tk.W).grid(
            row=0, column=0, sticky=tk.W, **pad)
        out_var = tk.StringVar(value=self.output_folder)
        tk.Entry(win, textvariable=out_var, width=48).grid(
            row=0, column=1, **pad)

        def browse_out():
            d = filedialog.askdirectory(
                parent=win, initialdir=self.output_folder or None)
            if d:
                out_var.set(d)

        tk.Button(win, text="...", command=browse_out, width=3).grid(
            row=0, column=2, **pad)

        # CSV name
        tk.Label(win, text="CSV Filename:", anchor=tk.W).grid(
            row=1, column=0, sticky=tk.W, **pad)
        csv_var = tk.StringVar(value=self.csv_name)
        tk.Entry(win, textvariable=csv_var, width=48).grid(
            row=1, column=1, **pad)

        # Image format
        tk.Label(win, text="Image Filename:", anchor=tk.W).grid(
            row=2, column=0, sticky=tk.W, **pad)
        img_var = tk.StringVar(value=self.image_format)
        tk.Entry(win, textvariable=img_var, width=48).grid(
            row=2, column=1, **pad)

        # Help
        tk.Label(
            win,
            text="Variables:  {filename}  {frame}  {us}",
            font=("Consolas", 9), fg="gray",
        ).grid(row=3, column=0, columnspan=3, sticky=tk.W, **pad)

        # Preview
        preview_var = tk.StringVar()

        def update_preview(*_):
            try:
                name = img_var.get().format(
                    filename="example", frame=120, us=4000000)
            except (KeyError, ValueError):
                name = "(format error)"
            preview_var.set(f"Preview:  {name}")

        img_var.trace_add("write", update_preview)
        update_preview()

        tk.Label(
            win, textvariable=preview_var,
            font=("Consolas", 9), fg="#555",
        ).grid(row=4, column=0, columnspan=3, sticky=tk.W, **pad)

        # Buttons
        btn = tk.Frame(win)
        btn.grid(row=5, column=0, columnspan=3, pady=12)

        def save():
            self.output_folder = out_var.get()
            self.csv_name = csv_var.get() or DEFAULT_CSV_NAME
            self.image_format = img_var.get() or DEFAULT_IMAGE_FMT
            self._save_config()
            self._update_footer()
            win.destroy()

        tk.Button(btn, text="Save", command=save, width=10).pack(
            side=tk.LEFT, padx=8)
        tk.Button(btn, text="Cancel", command=win.destroy, width=10).pack(
            side=tk.LEFT, padx=8)

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
        self.footer_var.set(f"Input: {inp}  |  Output: {out}")

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
            self.info_var.set("所有影片皆已標記完畢")
            self.file_label.config(text="")
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
        self.segments = []

        self.slider.configure(to=max(self.total_frames - 1, 0))
        self.file_label.config(
            text=f"{os.path.basename(path)}  ({self.file_idx + 1}/{len(self.files)})"
        )
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
        self.info_var.set(
            f"Frame: {self.frame_no}/{self.total_frames - 1}  |  "
            f"Time: {self._cached_us} us  |  "
            f"Start: {s}  |  End: {e}  |  "
            f"Segments: {len(self.segments)}\n"
            f"[A/D/Arrow]+-1  [W/S]+-30  [Shift+A/D]Head/Tail  "
            f"[J]Start  [K]End  [Enter]Save  [P]Capture  "
            f"[Ctrl+Z]Undo  [Ctrl+S]Write  [Shift+S]View  "
            f"[BS]Prev  [N]Next  [Ctrl+Q]Quit"
        )

        cw = self.vframe.winfo_width()
        ch = self.vframe.winfo_height()
        if cw <= 1 or ch <= 1:
            return

        frame = self._cached_frame
        ih, iw = frame.shape[:2]
        scale = min(cw / iw, ch / ih)
        new_w, new_h = int(iw * scale), int(ih * scale)
        if new_w <= 0 or new_h <= 0:
            return

        if (new_w, new_h) != (iw, ih):
            frame = cv2.resize(frame, (new_w, new_h), interpolation=cv2.INTER_LINEAR)

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(rgb)

        if (self.photo is not None
                and self.photo.width() == new_w
                and self.photo.height() == new_h):
            self.photo.paste(img)
        else:
            self.photo = ImageTk.PhotoImage(img)
            self.canvas.configure(image=self.photo)

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
        self.segments.append((self.start_us, self.end_us))
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
        self.info_var.set(f"Captured: {img_name}")

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
        fname = os.path.basename(self.files[self.file_idx])
        csv_path = self._resolve_csv_path()
        write_header = not os.path.exists(csv_path)

        with open(csv_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if write_header:
                writer.writerow(["filename", "segment", "start_us", "end_us"])
            for i, (s, e) in enumerate(self.segments, 1):
                writer.writerow([fname, i, s, e])
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
        for i, (s, e) in enumerate(self.segments, 1):
            lines.append(f"  [{i}]  {s}  ~  {e}  us")
        fname = os.path.basename(self.files[self.file_idx]) if self.files else ""
        msg = f"檔案: {fname}\n共 {len(self.segments)} 段:\n\n" + "\n".join(lines)
        messagebox.showinfo("暫存標記一覽", msg)

    def _undo_segment(self):
        if not self.segments:
            return
        removed = self.segments.pop()
        self.start_us = removed[0]
        self.end_us = removed[1]
        self._display()

    # ── Next / Prev video ──────────────────────────

    def _next_video(self):
        if not self.cap or not self.output_folder:
            return
        self._write_csv()
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
