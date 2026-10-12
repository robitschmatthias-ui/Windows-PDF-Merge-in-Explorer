"""tkinter UI: Advanced dialog, progress window, tooltips."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import ttk


class Tooltip:
    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, _event=None):
        if self.tip:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, justify="left", background="#ffffe0",
                 relief="solid", borderwidth=1, padx=6, pady=4).pack()

    def _hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


def center(win: tk.Toplevel | tk.Tk) -> None:
    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth() - w) // 2
    y = (win.winfo_screenheight() - h) // 3
    win.geometry(f"+{x}+{y}")


# ── Advanced dialog ─────────────────────────────────────────────────────────

@dataclass
class AdvancedOptions:
    files: list[Path]
    size: str        # "original" | "a4"
    orient: str      # "original" | "portrait" | "landscape"
    open_after: bool


class AdvancedDialog:
    def __init__(self, root: tk.Tk, files: list[Path]):
        self.files = list(files)
        self.result: AdvancedOptions | None = None
        self._drag_index: int | None = None

        self.win = win = tk.Toplevel(root)
        win.title("PDF Advanced Merge")
        win.minsize(560, 480)
        win.protocol("WM_DELETE_WINDOW", self._cancel)

        main = ttk.Frame(win, padding=12)
        main.pack(fill="both", expand=True)

        # File list
        frame_files = ttk.LabelFrame(main, text=" File order ", padding=8)
        frame_files.pack(fill="both", expand=True)

        list_row = ttk.Frame(frame_files)
        list_row.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(list_row, activestyle="none", height=12,
                                  selectmode="browse", exportselection=False)
        scroll = ttk.Scrollbar(list_row, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        self.listbox.bind("<Button-1>", self._drag_start)
        self.listbox.bind("<B1-Motion>", self._drag_motion)
        Tooltip(self.listbox, "Drag & drop files into the desired order")

        buttons = ttk.Frame(list_row, padding=(8, 0, 0, 0))
        buttons.pack(side="left", fill="y")
        ttk.Button(buttons, text="▲ Up", command=lambda: self._move(-1)).pack(fill="x")
        ttk.Button(buttons, text="▼ Down", command=lambda: self._move(1)).pack(fill="x", pady=(4, 0))
        self._refresh_list()

        # Options
        options = ttk.Frame(main)
        options.pack(fill="x", pady=(12, 0))

        self.size = tk.StringVar(value="original")
        frame_size = ttk.LabelFrame(options, text=" Page size ", padding=10)
        frame_size.pack(side="left", fill="both", expand=True)
        rb = ttk.Radiobutton(frame_size, text="Keep original", value="original", variable=self.size)
        rb.pack(anchor="w")
        Tooltip(rb, "Pages keep their original size.\n"
                    "Mixed page formats are preserved.")
        rb = ttk.Radiobutton(frame_size, text="Fit to A4", value="a4", variable=self.size)
        rb.pack(anchor="w", pady=(6, 0))
        Tooltip(rb, "• PDF and Office pages → scaled to A4 and centered\n"
                    "• Images with at least 200 DPI on A4 → fill the page\n"
                    "• Smaller images → original size, centered\n"
                    "  (only shrunk if larger than A4)")

        self.orient = tk.StringVar(value="original")
        frame_orient = ttk.LabelFrame(options, text=" Orientation ", padding=10)
        frame_orient.pack(side="left", fill="both", expand=True, padx=(12, 0))
        for text, value, tip in (
            ("Keep original", "original",
             "Every page keeps its current orientation."),
            ("Portrait", "portrait",
             "Landscape pages are rotated by 90° to portrait.\n"
             "Portrait pages stay unchanged."),
            ("Landscape", "landscape",
             "Portrait pages are rotated by 90° to landscape.\n"
             "Landscape pages stay unchanged."),
        ):
            rb = ttk.Radiobutton(frame_orient, text=text, value=value, variable=self.orient)
            rb.pack(anchor="w", pady=(0, 6))
            Tooltip(rb, tip)

        self.open_after = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(main, text="Open the PDF in the default program after merging",
                             variable=self.open_after)
        cb.pack(anchor="w", pady=(12, 0))
        Tooltip(cb, "Opens the finished PDF in the default PDF program,\n"
                    "e.g. to rotate or move individual pages.")

        bottom = ttk.Frame(main)
        bottom.pack(fill="x", pady=(12, 0))
        ttk.Button(bottom, text="Next", command=self._ok).pack(side="right")
        ttk.Button(bottom, text="Cancel", command=self._cancel).pack(side="right", padx=(0, 8))
        win.bind("<Escape>", lambda _e: self._cancel())

        center(win)
        win.lift()
        win.focus_force()

    def _refresh_list(self, select: int | None = None) -> None:
        self.listbox.delete(0, "end")
        for i, f in enumerate(self.files, start=1):
            self.listbox.insert("end", f"{i:>3}.  {f.name}")
        if select is not None:
            self.listbox.selection_set(select)
            self.listbox.see(select)

    def _drag_start(self, event) -> None:
        self._drag_index = self.listbox.nearest(event.y)

    def _drag_motion(self, event) -> None:
        if self._drag_index is None:
            return
        target = self.listbox.nearest(event.y)
        if target != self._drag_index:
            item = self.files.pop(self._drag_index)
            self.files.insert(target, item)
            self._drag_index = target
            self._refresh_list(select=target)

    def _move(self, delta: int) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if 0 <= j < len(self.files):
            self.files[i], self.files[j] = self.files[j], self.files[i]
            self._refresh_list(select=j)

    def _ok(self) -> None:
        self.result = AdvancedOptions(self.files, self.size.get(), self.orient.get(),
                                      self.open_after.get())
        self.win.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.win.destroy()

    def show(self) -> AdvancedOptions | None:
        self.win.wait_window()
        return self.result


# ── Progress ────────────────────────────────────────────────────────────────

class _ProgressWindow:
    def __init__(self, root: tk.Tk, title: str, text: str):
        self.win = win = tk.Toplevel(root)
        win.title(title)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", lambda: None)  # no cancel, as on Linux
        frame = ttk.Frame(win, padding=16)
        frame.pack(fill="both", expand=True)
        self.label = ttk.Label(frame, text=text, width=50)
        self.label.pack(anchor="w")
        self.bar = ttk.Progressbar(frame, length=360, maximum=100)
        self.bar.pack(fill="x", pady=(10, 0))
        center(win)
        win.lift()

    def update(self, percent: int, text: str | None) -> None:
        self.bar["value"] = percent
        if text:
            self.label["text"] = text


def run_with_progress(root: tk.Tk, title: str, text: str, job, after_done=None):
    """Runs job(report) in a background thread while showing a progress
    window. report(percent, text) updates the display.

    after_done() runs in the same thread AFTER the window is closed
    (e.g. quitting Office without making the user wait for it).
    Returns job's return value or re-raises its exception.
    """
    messages: queue.Queue = queue.Queue()
    outcome: dict = {}

    def report(percent: int, msg: str | None = None) -> None:
        messages.put(("progress", percent, msg))

    def worker() -> None:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            try:
                outcome["value"] = job(report)
            except BaseException as exc:  # handed over to the UI thread
                outcome["error"] = exc
            messages.put(("done",))
            if after_done:
                after_done()
        finally:
            pythoncom.CoUninitialize()

    window = _ProgressWindow(root, title, text)
    thread = threading.Thread(target=worker, name="pdf-merge-worker")
    thread.start()

    def poll() -> None:
        try:
            while True:
                item = messages.get_nowait()
                if item[0] == "done":
                    window.win.destroy()
                    root.quit()
                    return
                window.update(item[1], item[2])
        except queue.Empty:
            pass
        root.after(80, poll)

    root.after(80, poll)
    root.mainloop()

    if "error" in outcome:
        raise outcome["error"]
    return outcome.get("value")
