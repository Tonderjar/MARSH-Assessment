"""
Marsh — FP&A Report Generator
Run: python src/main.py
"""
import os
import sys
import glob
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Crisp rendering on Windows high-DPI displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# When frozen by PyInstaller, sys._MEIPASS holds the extraction dir.
# In dev, fall back to the directory of this file.
_base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _base)
from engine import generate_report

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY    = "#1C3F6E"
BLUE    = "#2E75B6"
LIGHT   = "#BDD7EE"
BG      = "#EFF3F8"
CARD    = "#FFFFFF"
TEXT    = "#1F2D3D"
MUTED   = "#6B7C93"
DIVIDER = "#D8E2EE"
SUCCESS = "#1A7A3F"
S_HOVER = "#145C2F"
ERROR   = "#B71C1C"
E_HOVER = "#8B0000"
B_HOVER = "#245D96"

F_TITLE  = ("Calibri", 20, "bold")
F_SUB    = ("Calibri", 11)
F_LABEL  = ("Calibri", 10, "bold")
F_INPUT  = ("Calibri", 10)
F_MONO   = ("Consolas", 9)
F_BTN    = ("Calibri", 13, "bold")
F_SMALL  = ("Calibri", 9)
F_HINT   = ("Calibri", 9, "italic")


def _hover(widget, normal, hot):
    widget.bind("<Enter>", lambda _: widget.configure(bg=hot))
    widget.bind("<Leave>", lambda _: widget.configure(bg=normal))


# ── Application ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Marsh  ·  FP&A Report Generator")
        self.configure(bg=BG)
        self.resizable(False, False)

        self._data_var   = tk.StringVar()
        self._fx_var     = tk.StringVar()
        self._output_var = tk.StringVar()
        self._status_var = tk.StringVar(value="Ready to generate.")
        self._queue      = queue.Queue()
        self._running    = False
        self._result     = None

        # Auto-fill derived fields whenever the data folder changes
        self._data_var.trace_add("write", self._on_data_change)

        self._build_ui()
        self._center()
        self._poll()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Banner ────────────────────────────────────────────────────────────
        banner = tk.Frame(self, bg=NAVY, height=72)
        banner.pack(fill="x")
        banner.pack_propagate(False)
        tk.Label(banner, text="MARSH", font=F_TITLE,
                 bg=NAVY, fg="white").place(x=28, y=10)
        tk.Label(banner, text="FP&A Report Generator", font=F_SUB,
                 bg=NAVY, fg=LIGHT).place(x=30, y=46)

        # Blue accent line under banner
        tk.Frame(self, bg=BLUE, height=4).pack(fill="x")

        # ── Main card ─────────────────────────────────────────────────────────
        card = tk.Frame(self, bg=CARD, padx=34, pady=26)
        card.pack(fill="both", expand=True, padx=18, pady=14)

        # Input rows
        self._row(card, "1.  Data Folder",
                  "Root folder that contains all DataTrans_* files",
                  self._data_var, mode="dir")

        self._row(card, "2.  FX Rates File",
                  "Auto-detected inside the data folder if left blank",
                  self._fx_var, mode="open",
                  filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])

        self._row(card, "3.  Save Report As",
                  "Destination path for the generated .xlsx report",
                  self._output_var, mode="save",
                  filetypes=[("Excel workbook", "*.xlsx")])

        # Divider
        tk.Frame(card, bg=DIVIDER, height=1).pack(fill="x", pady=(22, 18))

        # ── Generate button ───────────────────────────────────────────────────
        self._btn = tk.Button(
            card, text="⚡  Generate Report",
            font=F_BTN, bg=BLUE, fg="white",
            activebackground=B_HOVER, activeforeground="white",
            relief="flat", cursor="hand2", pady=12, padx=40,
            command=self._run,
        )
        self._btn.pack()
        _hover(self._btn, BLUE, B_HOVER)

        # Progress bar
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Marsh.Horizontal.TProgressbar",
                        troughcolor="#D8E2EE", background=BLUE,
                        thickness=6)
        self._pb = ttk.Progressbar(card, mode="indeterminate", length=520,
                                   style="Marsh.Horizontal.TProgressbar")
        self._pb.pack(pady=(14, 0))

        # ── Log terminal ──────────────────────────────────────────────────────
        term = tk.Frame(card, bg="#0E1D2D", bd=0)
        term.pack(fill="both", expand=True, pady=(14, 0))

        self._log = tk.Text(
            term, bg="#0E1D2D", fg="#7EB3D4",
            font=F_MONO, height=9, relief="flat",
            state="disabled", padx=14, pady=10, wrap="word",
            selectbackground=BLUE, insertbackground=BLUE,
        )
        sb = ttk.Scrollbar(term, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        self._log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._log.tag_config("head",  foreground="#FFFFFF",
                             font=("Consolas", 9, "bold"))
        self._log.tag_config("meta",  foreground="#4A7A9B")
        self._log.tag_config("ok",    foreground="#66BB6A")
        self._log.tag_config("err",   foreground="#EF5350")

        # ── Status bar ────────────────────────────────────────────────────────
        bar = tk.Frame(self, bg=NAVY, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        tk.Label(bar, textvariable=self._status_var, font=F_SMALL,
                 bg=NAVY, fg=LIGHT, anchor="w", padx=20).pack(fill="x", pady=4)

    def _row(self, parent, label, hint, var, mode, filetypes=None):
        """One labelled file-picker row."""
        frame = tk.Frame(parent, bg=CARD)
        frame.pack(fill="x", pady=6)

        # Label + hint on same line
        top = tk.Frame(frame, bg=CARD)
        top.pack(fill="x")
        tk.Label(top, text=label, font=F_LABEL, bg=CARD, fg=TEXT).pack(side="left")
        tk.Label(top, text=f"  —  {hint}", font=F_HINT, bg=CARD, fg=MUTED
                 ).pack(side="left")

        # Entry + browse button
        row = tk.Frame(frame, bg=CARD)
        row.pack(fill="x", pady=(4, 0))

        entry = tk.Entry(row, textvariable=var, font=F_INPUT,
                         bg="#F4F8FC", fg=TEXT, relief="solid", bd=1,
                         insertbackground=BLUE,
                         highlightthickness=1,
                         highlightcolor=BLUE,
                         highlightbackground=DIVIDER)
        entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        def _browse():
            if mode == "dir":
                p = filedialog.askdirectory(title=label, parent=self)
            elif mode == "save":
                p = filedialog.asksaveasfilename(
                    title=label, parent=self,
                    defaultextension=".xlsx", filetypes=filetypes or [],
                )
            else:
                p = filedialog.askopenfilename(
                    title=label, parent=self, filetypes=filetypes or [],
                )
            if p:
                var.set(p)

        btn = tk.Button(row, text="Browse …", command=_browse,
                        font=F_SMALL, bg=LIGHT, fg=NAVY,
                        activebackground=BLUE, activeforeground="white",
                        relief="flat", cursor="hand2", padx=14, pady=6)
        btn.pack(side="right")
        _hover(btn, LIGHT, BLUE)

    # ── Smart defaults ────────────────────────────────────────────────────────

    def _on_data_change(self, *_):
        root = self._data_var.get().strip()
        if not root or not os.path.isdir(root):
            return

        # Auto-fill output path (sibling of the data folder)
        if not self._output_var.get().strip():
            parent = os.path.dirname(os.path.abspath(root))
            self._output_var.set(
                os.path.join(parent, "Deliverable_1_FPA_Analysis.xlsx")
            )

        # Auto-detect FX rates file inside the data folder
        if not self._fx_var.get().strip():
            hits = glob.glob(
                os.path.join(root, "**", "*FX*Rates*.xlsx"), recursive=True
            ) or glob.glob(
                os.path.join(root, "**", "*FX*.xlsx"), recursive=True
            )
            if hits:
                self._fx_var.set(hits[0])

    # ── Generate pipeline ─────────────────────────────────────────────────────

    def _run(self):
        if self._running:
            return

        data_root = self._data_var.get().strip()
        fx_path   = self._fx_var.get().strip() or None
        out_path  = self._output_var.get().strip()

        # Validate
        if not data_root:
            messagebox.showerror("Missing input",
                                 "Please select a Data Folder.", parent=self)
            return
        if not os.path.isdir(data_root):
            messagebox.showerror("Invalid path",
                                 f"Folder not found:\n{data_root}", parent=self)
            return
        if fx_path and not os.path.isfile(fx_path):
            messagebox.showerror("Invalid path",
                                 f"FX rates file not found:\n{fx_path}", parent=self)
            return
        if not out_path:
            messagebox.showerror("Missing output",
                                 "Please specify where to save the report.", parent=self)
            return

        # Prepare UI
        self._running = True
        self._result  = None
        self._btn.configure(state="disabled", text="Running…",
                            bg="#5A8FC0", cursor="wait")
        self._pb.start(10)
        self._status_var.set("Processing — please wait…")

        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

        self._write("FP&A Report Generator", "head")
        self._write(f"  Data root : {data_root}", "meta")
        self._write(f"  FX file   : {fx_path or '(auto-detect)'}", "meta")
        self._write(f"  Output    : {out_path}", "meta")
        self._write("", None)

        def _worker():
            try:
                path = generate_report(
                    data_root=data_root,
                    fx_path=fx_path,
                    output_path=out_path,
                    progress_cb=lambda m: self._queue.put((m, None)),
                )
                self._result = path
                self._queue.put(("", None))
                self._queue.put(("✓  Report saved successfully.", "ok"))
                self.after(0, self._on_done)
            except Exception as exc:
                self._queue.put(("", None))
                self._queue.put((f"✗  {exc}", "err"))
                self.after(0, self._on_error)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self):
        self._running = False
        self._pb.stop()
        self._btn.configure(
            state="normal", text="📂  Open Report",
            cursor="hand2", bg=SUCCESS, command=self._open,
        )
        _hover(self._btn, SUCCESS, S_HOVER)
        self._status_var.set(f"Done  →  {self._result}")

    def _on_error(self):
        self._running = False
        self._pb.stop()
        self._btn.configure(
            state="normal", text="Retry",
            cursor="hand2", bg=ERROR, command=self._reset,
        )
        _hover(self._btn, ERROR, E_HOVER)
        self._status_var.set("Error — see log above.")

    def _open(self):
        if self._result and os.path.isfile(self._result):
            os.startfile(self._result)
        self._reset()

    def _reset(self):
        self._btn.configure(text="⚡  Generate Report", bg=BLUE, command=self._run)
        _hover(self._btn, BLUE, B_HOVER)
        self._status_var.set("Ready to generate.")

    # ── Utilities ─────────────────────────────────────────────────────────────

    def _write(self, msg, tag):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n", tag or "")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _poll(self):
        try:
            while True:
                msg, tag = self._queue.get_nowait()
                self._write(msg, tag)
                if msg:
                    self._status_var.set(msg)
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def _center(self):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
