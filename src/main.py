"""
main.py — Tkinter GUI front-end for the Marsh FP&A Report Generator.

Provides a desktop interface that:
  - Lets the user pick a data folder, FX rates file, and output path.
  - Offers checkboxes to select which reports to generate (D1 and/or D2).
  - Provides a month-range filter; year is inferred from the data folder name.
  - Runs the engine pipeline on a background thread to keep the UI responsive.
  - Streams progress messages into an on-screen terminal log via a thread-safe queue.
  - Writes a timestamped log file to logs/ at project root on every run.
  - After generation, offers to open the report and/or send it via Outlook.

Run standalone:  python src/main.py
"""
import os
import sys
import glob
import queue
import logging
import pathlib
import calendar
import threading
import urllib.parse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime

# Crisp rendering on Windows high-DPI displays
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# When frozen by PyInstaller, sys._MEIPASS holds the extraction dir.
_base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _base)

# Add project root (parent of src/) so we can import build_deck
_project_root = os.path.dirname(_base)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from engine import generate_report, generate_latam_report


# ── Logging setup ─────────────────────────────────────────────────────────────

def _setup_logging() -> pathlib.Path:
    """
    Configure file-based logging for the entire application (and the engine module).

    Log location:
      - Frozen (PyInstaller .exe): <exe folder>/logs/
      - Development:               <project root>/logs/

    Each run creates a new timestamped log file so sessions are never overwritten.
    Both marsh.app and marsh.engine loggers write to this file automatically because
    basicConfig installs a root-level handler and both loggers propagate by default.

    Returns the Path of the log file (logged at startup so it is easy to locate).
    """
    if getattr(sys, "frozen", False):
        # sys.executable is <project>/dist/<name>.exe — go up one level so logs/
        # lands next to dist/ and output/, not inside dist/.
        log_dir = pathlib.Path(os.path.dirname(sys.executable)).parent / "logs"
    else:
        log_dir = pathlib.Path(_project_root) / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"marsh_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8")],
    )
    return log_file


_log_file = _setup_logging()
_app_logger = logging.getLogger("marsh.app")
_app_logger.info("Application started — log file: %s", _log_file)


# ── Palette ───────────────────────────────────────────────────────────────────
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

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _hover(widget, normal, hot):
    """Bind Enter/Leave events to swap the widget's background colour."""
    widget.bind("<Enter>", lambda _: widget.configure(bg=hot))
    widget.bind("<Leave>", lambda _: widget.configure(bg=normal))


# ── Application ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    """
    Main application window for the Marsh FP&A Report Generator.

    Architecture overview:
      - All widget creation happens in _build_ui(), called once from __init__.
      - Report generation runs on a daemon thread (_worker inside _run) so the
        Tk event loop never blocks and the UI stays responsive.
      - Thread-to-GUI communication uses a Queue: the worker puts (message, tag)
        tuples and _poll() drains the queue every 80 ms on the main thread.
      - _on_done / _on_error are scheduled via self.after(0, ...) from the worker
        thread, ensuring all Tk widget mutations happen on the main thread.
    """

    def __init__(self):
        """
        Initialise the application: set up state variables, build the UI,
        centre the window, and start the 80 ms polling loop.
        """
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
        self._results    = []   # absolute paths of successfully generated reports

        # Report type toggles
        self._d1_var = tk.BooleanVar(value=True)
        self._d2_var = tk.BooleanVar(value=False)

        # Month filter
        self._mfrom_var = tk.StringVar(value="Jan")
        self._mto_var   = tk.StringVar(value="Dec")

        # Auto-fill output and FX paths whenever the data folder changes
        self._data_var.trace_add("write", self._on_data_change)

        self._build_ui()
        self._center()
        self._poll()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        """Build and lay out all widgets — called once from __init__."""
        # Banner
        banner = tk.Frame(self, bg=NAVY, height=72)
        banner.pack(fill="x")
        banner.pack_propagate(False)
        tk.Label(banner, text="MARSH", font=F_TITLE,
                 bg=NAVY, fg="white").place(x=28, y=10)
        tk.Label(banner, text="FP&A Report Generator", font=F_SUB,
                 bg=NAVY, fg=LIGHT).place(x=30, y=46)

        tk.Frame(self, bg=BLUE, height=4).pack(fill="x")

        # Main card
        card = tk.Frame(self, bg=CARD, padx=34, pady=22)
        card.pack(fill="both", expand=True, padx=18, pady=14)

        # File inputs
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

        # Report type section
        tk.Frame(card, bg=DIVIDER, height=1).pack(fill="x", pady=(14, 10))
        self._section_label(card, "Report Type")

        chk_frame = tk.Frame(card, bg=CARD)
        chk_frame.pack(fill="x", pady=(2, 0))
        self._chk(chk_frame, "D1 — Full Analysis Report (.xlsx)", self._d1_var, 0)
        self._chk(chk_frame, "D2 — LATAM Report (.xlsx)",         self._d2_var, 1)

        # Date filter section
        tk.Frame(card, bg=DIVIDER, height=1).pack(fill="x", pady=(14, 10))
        self._section_label(card, "Month Filter")

        filt = tk.Frame(card, bg=CARD)
        filt.pack(fill="x", pady=(2, 0))

        tk.Label(filt, text="From:", font=F_LABEL, bg=CARD, fg=TEXT).pack(side="left")
        ttk.Combobox(filt, textvariable=self._mfrom_var,
                     values=MONTHS, state="readonly", width=6, font=F_INPUT
                     ).pack(side="left", padx=(4, 14))

        tk.Label(filt, text="To:", font=F_LABEL, bg=CARD, fg=TEXT).pack(side="left")
        ttk.Combobox(filt, textvariable=self._mto_var,
                     values=MONTHS, state="readonly", width=6, font=F_INPUT
                     ).pack(side="left", padx=(4, 0))

        # Generate button
        tk.Frame(card, bg=DIVIDER, height=1).pack(fill="x", pady=(18, 14))
        self._btn = tk.Button(
            card, text="Generate Report",
            font=F_BTN, bg=BLUE, fg="white",
            activebackground=B_HOVER, activeforeground="white",
            relief="flat", cursor="hand2", pady=12, padx=40,
            command=self._run,
        )
        self._btn.pack()
        _hover(self._btn, BLUE, B_HOVER)

        # Outlook button — always visible; opens a file picker directed at the output folder
        self._email_btn = tk.Button(
            card, text="Send via Outlook",
            font=F_BTN, bg=NAVY, fg="white",
            activebackground=B_HOVER, activeforeground="white",
            relief="flat", cursor="hand2", pady=12, padx=28,
            command=self._send_outlook,
        )
        _hover(self._email_btn, NAVY, B_HOVER)
        self._email_btn.pack(pady=(8, 0))

        # Progress bar
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Marsh.Horizontal.TProgressbar",
                        troughcolor="#D8E2EE", background=BLUE, thickness=6)
        self._pb = ttk.Progressbar(card, mode="indeterminate", length=520,
                                   style="Marsh.Horizontal.TProgressbar")
        self._pb.pack(pady=(14, 0))

        # Log terminal
        term = tk.Frame(card, bg="#0E1D2D", bd=0)
        term.pack(fill="both", expand=True, pady=(14, 0))
        self._log_widget = tk.Text(
            term, bg="#0E1D2D", fg="#7EB3D4",
            font=F_MONO, height=9, relief="flat",
            state="disabled", padx=14, pady=10, wrap="word",
            selectbackground=BLUE, insertbackground=BLUE,
        )
        sb = ttk.Scrollbar(term, command=self._log_widget.yview)
        self._log_widget.configure(yscrollcommand=sb.set)
        self._log_widget.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._log_widget.tag_config("head",  foreground="#FFFFFF", font=("Consolas", 9, "bold"))
        self._log_widget.tag_config("meta",  foreground="#4A7A9B")
        self._log_widget.tag_config("ok",    foreground="#66BB6A")
        self._log_widget.tag_config("err",   foreground="#EF5350")

        # Status bar
        bar = tk.Frame(self, bg=NAVY, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        tk.Label(bar, textvariable=self._status_var, font=F_SMALL,
                 bg=NAVY, fg=LIGHT, anchor="w", padx=20).pack(fill="x", pady=4)

    def _section_label(self, parent, text):
        tk.Label(parent, text=text, font=F_LABEL, bg=CARD, fg=NAVY).pack(anchor="w")

    def _chk(self, parent, text, var, col):
        cb = tk.Checkbutton(parent, text=text, variable=var,
                            font=F_INPUT, bg=CARD, fg=TEXT,
                            activebackground=CARD, selectcolor=LIGHT,
                            cursor="hand2")
        cb.grid(row=0, column=col, padx=(0, 24), sticky="w")

    def _row(self, parent, label, hint, var, mode, filetypes=None):
        """
        Create a labelled input row consisting of a text entry and a Browse button.

        mode: 'dir'  → directory picker dialog
              'open' → file open dialog
              'save' → file save-as dialog
        filetypes: list of (description, pattern) tuples passed to the dialog.
        """
        frame = tk.Frame(parent, bg=CARD)
        frame.pack(fill="x", pady=6)
        top = tk.Frame(frame, bg=CARD)
        top.pack(fill="x")
        tk.Label(top, text=label, font=F_LABEL, bg=CARD, fg=TEXT).pack(side="left")
        tk.Label(top, text=f"  —  {hint}", font=F_HINT, bg=CARD, fg=MUTED).pack(side="left")
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
                    defaultextension=".xlsx", filetypes=filetypes or [])
            else:
                p = filedialog.askopenfilename(title=label, parent=self, filetypes=filetypes or [])
            if p:
                var.set(p)

        btn = tk.Button(row, text="Browse ...", command=_browse,
                        font=F_SMALL, bg=LIGHT, fg=NAVY,
                        activebackground=BLUE, activeforeground="white",
                        relief="flat", cursor="hand2", padx=14, pady=6)
        btn.pack(side="right")
        _hover(btn, LIGHT, BLUE)

    # ── Month filter helpers ──────────────────────────────────────────────────

    def _parse_dates(self):
        """
        Convert the From/To month dropdowns to datetime boundaries for the engine filter.

        Year derivation: the data folder is expected to be named after the reporting
        year (e.g. '2024').  The base name of the selected path is parsed as an int;
        if it is not a plain year number the current calendar year is used as a fallback.

        Returns (None, None) when the full Jan–Dec range is selected, signalling to
        the engine that no date filtering should be applied.
        """
        mf = MONTHS.index(self._mfrom_var.get()) + 1
        mt = MONTHS.index(self._mto_var.get())   + 1
        if mf > mt:
            mf, mt = mt, mf
        if mf == 1 and mt == 12:
            return None, None  # full year — no filter needed
        # Derive year from the selected data folder name (expected to be a year folder)
        folder = os.path.basename(self._data_var.get().strip())
        try:
            year = int(folder)
        except ValueError:
            year = datetime.now().year
        last_day = calendar.monthrange(year, mt)[1]
        return datetime(year, mf, 1), datetime(year, mt, last_day)

    # ── Smart defaults ────────────────────────────────────────────────────────

    def _on_data_change(self, *_):
        """
        Auto-fill the Output path and FX Rates path when the user selects a data folder.

        Output path walk-up logic:
          The selected folder may be a nested year sub-folder (e.g. data/2024).
          Walking straight to its parent would place the output inside data/, not
          at the project level.  Instead we walk up the directory tree until we find
          a sibling 'output/' folder that already exists — that is always the correct
          project-level output directory regardless of how deep the user drills down.
          If no existing output/ sibling is found we fall back to creating one next
          to the immediate parent of the selected folder.
        """
        root = self._data_var.get().strip()
        if not root or not os.path.isdir(root):
            return
        if not self._output_var.get().strip():
            p = os.path.abspath(root)
            out_dir = None
            while True:
                parent = os.path.dirname(p)
                if os.path.isdir(os.path.join(parent, "output")):
                    out_dir = os.path.join(parent, "output")
                    break
                if parent == p:  # reached filesystem root without finding output/
                    break
                p = parent
            if out_dir is None:
                out_dir = os.path.join(os.path.dirname(os.path.abspath(root)), "output")
            self._output_var.set(os.path.join(out_dir, "Deliverable_1_FPA_Analysis.xlsx"))
        if not self._fx_var.get().strip():
            hits = glob.glob(os.path.join(root, "**", "*FX*Rates*.xlsx"), recursive=True) \
                or glob.glob(os.path.join(root, "**", "*FX*.xlsx"), recursive=True)
            if hits:
                self._fx_var.set(hits[0])

    # ── Generate pipeline ─────────────────────────────────────────────────────

    def _run(self):
        """
        Validate inputs, then spawn a daemon thread to run the engine pipeline.

        Thread-to-GUI communication pattern:
          The worker thread puts (message, tag) tuples into self._queue for
          progress text, and a sentinel ("", None) to mark a stage boundary.
          The main thread drains the queue every 80 ms via _poll().
          _on_done and _on_error are invoked via self.after(0, ...) from the
          worker thread — this ensures all Tk widget mutations occur on the
          main thread, which is required by Tkinter.
        """
        if self._running:
            return

        data_root = self._data_var.get().strip()
        fx_path   = self._fx_var.get().strip() or None
        out_path  = self._output_var.get().strip()

        if not data_root:
            messagebox.showerror("Missing input", "Please select a Data Folder.", parent=self)
            return
        if not os.path.isdir(data_root):
            messagebox.showerror("Invalid path", f"Folder not found:\n{data_root}", parent=self)
            return
        if fx_path and not os.path.isfile(fx_path):
            messagebox.showerror("Invalid path", f"FX rates file not found:\n{fx_path}", parent=self)
            return
        if not out_path:
            messagebox.showerror("Missing output", "Please specify where to save the report.", parent=self)
            return

        selected = []
        if self._d1_var.get(): selected.append("d1")
        if self._d2_var.get(): selected.append("d2")
        if not selected:
            messagebox.showerror("No Report Selected",
                                 "Please check at least one report type.", parent=self)
            return

        try:
            date_from, date_to = self._parse_dates()
        except Exception as exc:
            messagebox.showerror("Date Error", f"Invalid date range:\n{exc}", parent=self)
            return

        self._running = True
        self._results = []
        self._btn.configure(state="disabled", text="Running...",
                            bg="#5A8FC0", cursor="wait")
        self._pb.start(10)
        self._status_var.set("Processing — please wait...")

        self._log_widget.configure(state="normal")
        self._log_widget.delete("1.0", "end")
        self._log_widget.configure(state="disabled")

        self._write("Marsh FP&A Report Generator", "head")
        self._write(f"  Data root  : {data_root}", "meta")
        self._write(f"  FX file    : {fx_path or '(auto-detect)'}", "meta")
        self._write(f"  Output     : {out_path}", "meta")
        reports_str = " | ".join(s.upper() for s in selected)
        self._write(f"  Reports    : {reports_str}", "meta")
        if date_from:
            self._write(f"  Date range : {date_from:%Y-%m-%d} to {date_to:%Y-%m-%d}", "meta")
        self._write("", None)
        _app_logger.info("Run started — data=%s fx=%s out=%s reports=%s dates=[%s,%s]",
                         data_root, fx_path, out_path, selected, date_from, date_to)

        out_dir = os.path.dirname(os.path.abspath(out_path))

        def _worker():
            try:
                if "d1" in selected:
                    path = generate_report(
                        data_root=data_root, fx_path=fx_path, output_path=out_path,
                        progress_cb=lambda m: self._queue.put((m, None)),
                        date_from=date_from, date_to=date_to,
                    )
                    self._results.append(path)

                if "d2" in selected:
                    d2_path = os.path.join(out_dir, "Deliverable_2_LatinAmerica.xlsx")
                    path = generate_latam_report(
                        data_root=data_root, fx_path=fx_path, output_path=d2_path,
                        progress_cb=lambda m: self._queue.put((m, None)),
                        date_from=date_from, date_to=date_to,
                    )
                    self._results.append(path)

                self._queue.put(("", None))
                n = len(self._results)
                self._queue.put((f"Done — {n} report{'s' if n != 1 else ''} generated.", "ok"))
                _app_logger.info("Run complete — %d reports: %s", n, self._results)
                self.after(0, self._on_done)
            except Exception as exc:
                _app_logger.exception("Report generation failed")
                self._queue.put(("", None))
                self._queue.put((f"Error: {exc}", "err"))
                self.after(0, self._on_error)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self):
        """Switch the main button to 'Open Report' and update the status bar."""
        self._running = False
        self._pb.stop()
        self._btn.configure(
            state="normal", text="Open Report",
            cursor="hand2", bg=SUCCESS, command=self._open,
        )
        _hover(self._btn, SUCCESS, S_HOVER)
        if self._results:
            self._status_var.set(f"Done  |  {self._results[0]}")

    def _on_error(self):
        """Switch the main button to 'Retry' in red."""
        self._running = False
        self._pb.stop()
        self._btn.configure(state="normal", text="Retry",
                            cursor="hand2", bg=ERROR, command=self._reset)
        _hover(self._btn, ERROR, E_HOVER)
        self._status_var.set("Error — see log above.")

    def _open(self):
        """
        Open the first generated report in the default application (Excel).

        Does not call _reset() so the Outlook button stays visible — the user
        may still want to email after reviewing the file.  The main button
        transitions to 'New Report' instead, allowing a fresh run.
        """
        if self._results and os.path.isfile(self._results[0]):
            os.startfile(self._results[0])
        self._btn.configure(text="New Report", bg=BLUE, command=self._reset)
        _hover(self._btn, BLUE, B_HOVER)

    def _reset(self):
        """Return the UI to its initial state, ready for a new report run."""
        self._btn.configure(text="Generate Report", bg=BLUE, command=self._run)
        _hover(self._btn, BLUE, B_HOVER)
        self._status_var.set("Ready to generate.")

    # ── Outlook email ─────────────────────────────────────────────────────────

    def _send_outlook(self):
        """
        Open a file picker directed at the output folder, copy the selected files
        to the Windows clipboard, then open a pre-filled compose window via the
        mailto: URI scheme.

        Using mailto: (rather than win32com COM automation) opens whatever app is
        set as the Windows default mail handler — including the new Outlook, which
        does not expose the classic COM interface that win32com relies on.

        The selected files are copied to the clipboard via PowerShell Set-Clipboard
        so the user can paste them (Ctrl+V) directly into the attachment area.
        """
        # Open the file picker in the output folder so D1/D2 are immediately visible
        out_path = self._output_var.get().strip()
        initial_dir = os.path.dirname(os.path.abspath(out_path)) if out_path else os.path.expanduser("~")

        paths = filedialog.askopenfilenames(
            title="Select files to attach",
            initialdir=initial_dir,
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
            parent=self,
        )
        if not paths:
            return

        recipient = simpledialog.askstring(
            "Send via Outlook",
            "Recipient email address:",
            parent=self,
        )
        if not recipient:
            return

        valid_paths = [os.path.abspath(p) for p in paths if os.path.isfile(p)]

        # Copy files to the Windows clipboard in CF_HDROP format — the same format
        # Explorer uses when you Ctrl+C a file, which Outlook recognises as attachments.
        # PowerShell Set-Clipboard -Path does not produce CF_HDROP so files don't paste
        # as attachments; writing the raw clipboard data directly is the reliable fix.
        if valid_paths:
            try:
                import struct
                import win32clipboard

                # DROPFILES header (20 bytes): pFiles offset, drop point x/y, fNC, fWide=1
                file_list = "\0".join(valid_paths) + "\0\0"
                data = struct.pack("<IIIII", 20, 0, 0, 0, 1) + file_list.encode("utf-16-le")

                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, data)
                win32clipboard.CloseClipboard()
                _app_logger.info("Copied %d file(s) to clipboard (CF_HDROP)", len(valid_paths))
            except Exception as exc:
                _app_logger.warning("Clipboard copy failed: %s", exc)

        # mailto: URI opens the system default mail app — new Outlook, classic Outlook,
        # or any other client the user has set as default in Windows settings
        subject = urllib.parse.quote(f"Marsh FP&A Report — {datetime.now().strftime('%Y-%m-%d')}")
        body = urllib.parse.quote(
            "Please find attached the Marsh FP&A analysis report(s).\n\n"
            "This report was generated automatically by the Marsh FP&A Report Generator.\n\n"
            "Regards,\nMarsh FP&A Team"
        )
        mailto_uri = f"mailto:{urllib.parse.quote(recipient)}?subject={subject}&body={body}"

        try:
            _app_logger.info("Opening compose window for: %s", recipient)
            os.startfile(mailto_uri)
        except Exception as exc:
            _app_logger.exception("Failed to open mail client")
            messagebox.showerror("Mail Error", str(exc), parent=self)
            return

        file_names = "\n  •  ".join(os.path.basename(p) for p in valid_paths)
        messagebox.showinfo(
            "Compose Window Opened",
            f"Outlook has been opened with the recipient and subject pre-filled.\n\n"
            f"The following file(s) are ready in your clipboard:\n  •  {file_names}\n\n"
            f"Press Ctrl+V in the attachment area to attach them.",
            parent=self,
        )

    # ── Utilities ─────────────────────────────────────────────────────────────

    def _write(self, msg, tag):
        """Append a line to the on-screen terminal log, optionally coloured by tag."""
        self._log_widget.configure(state="normal")
        self._log_widget.insert("end", msg + "\n", tag or "")
        self._log_widget.see("end")
        self._log_widget.configure(state="disabled")

    def _poll(self):
        """
        Drain the progress queue and update the terminal log.

        Scheduled to run every 80 ms via Tk's after() scheduler.  All queue
        consumption and widget updates happen on the main thread, so there are
        no Tkinter thread-safety concerns.  The loop exits when the queue is
        empty (queue.Empty exception), then reschedules itself.
        """
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
        """Centre the window on the screen after all widgets have been laid out."""
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
