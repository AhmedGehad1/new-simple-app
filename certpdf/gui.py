"""The main window."""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
import tkinter as tk
import tkinter.font as tkfont
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import APP_NAME, __version__
from .builder import BuildResult, Cancelled, OutputError, build_pdfs
from .config import (
    CERTIFICATES,
    ENGINE_AUTO,
    ENGINE_EXCEL,
    ENGINE_LIBREOFFICE,
    PASS_FAIL,
    SHEET_SPECS,
    Settings,
    base_name,
    default_sheet_names,
)
from .engines import LIBREOFFICE_DOWNLOAD_URL, EngineError, available_engines, open_engine
from .excel_output import build_workbooks
from .icon import CHECKBOX_PNG_BASE64, HEADER_ICON_PNG_BASE64, ICON_PNG_BASE64
from .workbook_info import WorkbookReadError, is_excel_file, match_sheets, read_sheet_names

# Colours
BG = "#EEF2F7"
CARD = "#FFFFFF"
BORDER = "#D5DDE8"
HEADER = "#174A84"
HEADER_TEXT = "#D6E4F5"
PRIMARY = "#1F5FA8"
PRIMARY_DARK = "#174A84"
PRIMARY_DISABLED = "#A9BEDA"
TEXT = "#1E2A38"
MUTED = "#5F6F82"
OK = "#1E7B34"
WARN = "#A15C00"
ERROR = "#B3261E"
ROW_ALT = "#F6F8FB"
SELECTED = "#DCE8F7"
BUTTON = "#F3F6FA"
LOG_BG = "#FAFBFD"

ENGINE_CHOICES = (
    (ENGINE_AUTO, "Automatic (recommended)"),
    (ENGINE_EXCEL, "Microsoft Excel"),
    (ENGINE_LIBREOFFICE, "LibreOffice"),
)
ENGINE_NAMES = {ENGINE_EXCEL: "Microsoft Excel", ENGINE_LIBREOFFICE: "LibreOffice"}
INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

FOUND, MISSING, UNKNOWN = "✔", "✖", "?"


@dataclass
class FileEntry:
    path: Path
    sheet_names: list[str] | None = None     # None when the tabs could not be read
    read_error: str = ""
    found: dict[str, str | None] = field(default_factory=dict)
    state: str = "ready"                     # ready, warning, error, working, done
    status: str = ""


def natural_key(text: str):
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def open_path(path: Path, reveal: bool = False) -> None:
    """Open a file or folder with the default program (or show it in the file manager)."""
    if sys.platform == "win32":
        if reveal and path.is_file():
            subprocess.Popen(["explorer", "/select,", str(path)])
        else:
            os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-R", str(path)] if reveal else ["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path.parent if reveal and path.is_file() else path)])


class Tooltip:
    def __init__(self, widget: tk.Widget, text: str, font) -> None:
        self.widget, self.text, self.font = widget, text, font
        self.window: tk.Toplevel | None = None
        self.job = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None) -> None:
        self._hide()
        self.job = self.widget.after(600, self._show)

    def _show(self) -> None:
        x = self.widget.winfo_rootx() + 8
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(f"+{x}+{y}")
        tk.Label(self.window, text=self.text, bg="#2B3645", fg="white", font=self.font,
                 padx=8, pady=4, justify="left").pack()

    def _hide(self, _event=None) -> None:
        if self.job:
            self.widget.after_cancel(self.job)
            self.job = None
        if self.window:
            self.window.destroy()
            self.window = None


class QueueReporter:
    """Passes progress from the worker thread to the window."""

    def __init__(self, events: queue.Queue) -> None:
        self.events = events

    def log(self, level: str, text: str) -> None:
        self.events.put(("log", level, text))

    def progress(self, done: int, total: int, text: str) -> None:
        self.events.put(("progress", done, total, text))

    def file_status(self, index: int, state: str, text: str) -> None:
        self.events.put(("file", index, state, text))


@dataclass
class RunResult:
    excel: BuildResult | None = None
    pdf: BuildResult | None = None
    pdf_error: str = ""      # why no PDFs could be made at all


class StageReporter:
    """Spreads the progress of the Excel step and the PDF step over one progress bar."""

    def __init__(self, reporter: QueueReporter, stage: int, stages: int) -> None:
        self.reporter, self.stage, self.stages = reporter, stage, max(stages, 1)

    def log(self, level: str, text: str) -> None:
        self.reporter.log(level, text)

    def progress(self, done: int, total: int, text: str) -> None:
        fraction = (self.stage + done / max(total, 1)) / self.stages
        self.reporter.progress(int(fraction * 1000), 1000, text)

    def file_status(self, index: int, state: str, text: str) -> None:
        self.reporter.file_status(index, state, text)


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.withdraw()
        self.title(APP_NAME)
        self.settings = Settings.load()
        self.entries: list[FileEntry] = []
        self.events: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.closing = False
        self.last_outputs: dict[str, Path] = {}
        self._rematch_job = None
        self._lockable: list[tk.Widget] = []
        self.scale = max(1.0, self.winfo_fpixels("1i") / 96.0)

        self._init_fonts()
        self._init_style()
        self._icon = tk.PhotoImage(data=ICON_PNG_BASE64)
        self._header_icon = tk.PhotoImage(data=HEADER_ICON_PNG_BASE64)
        self.iconphoto(True, self._icon)

        self._build()
        self.engines = available_engines()
        self._update_engine_note()
        self._refresh_table()
        self._place_window()
        self.deiconify()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind_all("<Control-o>", lambda _e: self._add_files())
        self.after(100, self._drain_events)

    # ------------------------------------------------------------------ setup

    def px(self, value: float) -> int:
        return int(round(value * self.scale))

    def _init_fonts(self) -> None:
        default = tkfont.nametofont("TkDefaultFont")
        if sys.platform == "win32":
            default.configure(family="Segoe UI")
        default.configure(size=10)
        for name in ("TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            tkfont.nametofont(name).configure(family=default.actual("family"), size=10)
        family = default.actual("family")
        self.font_title = tkfont.Font(family=family, size=16, weight="bold")
        self.font_subtitle = tkfont.Font(family=family, size=10)
        self.font_section = tkfont.Font(family=family, size=11, weight="bold")
        self.font_bold = tkfont.Font(family=family, size=10, weight="bold")
        self.font_small = tkfont.Font(family=family, size=9)
        self.font_big_button = tkfont.Font(family=family, size=11, weight="bold")
        self.font_mark = tkfont.Font(family=family, size=11)
        mono = "Consolas" if sys.platform == "win32" else tkfont.nametofont("TkFixedFont").actual("family")
        self.font_log = tkfont.Font(family=mono, size=9)

    def _init_style(self) -> None:
        px = self.px
        self.configure(background=BG)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background=BG, foreground=TEXT, bordercolor=BORDER, focuscolor=PRIMARY)
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Card.TLabel", background=CARD)
        style.configure("Field.Card.TLabel", background=CARD, font=self.font_bold)
        style.configure("Muted.Card.TLabel", background=CARD, foreground=MUTED, font=self.font_small)
        style.configure("Section.Card.TLabel", background=CARD, font=self.font_section)
        style.configure("Status.Card.TLabel", background=CARD, foreground=MUTED)

        style.configure("TButton", padding=(px(12), px(5)), background=BUTTON, bordercolor=BORDER,
                        lightcolor=BUTTON, darkcolor=BUTTON, relief="flat")
        style.map("TButton",
                  background=[("disabled", BUTTON), ("pressed", "#D8E2EE"), ("active", "#E6EDF6")],
                  lightcolor=[("pressed", "#D8E2EE"), ("active", "#E6EDF6")],
                  darkcolor=[("pressed", "#D8E2EE"), ("active", "#E6EDF6")],
                  foreground=[("disabled", "#A3ADB9")])

        style.configure("Primary.TButton", background=PRIMARY, foreground="white", font=self.font_big_button,
                        padding=(px(28), px(9)), bordercolor=PRIMARY, lightcolor=PRIMARY, darkcolor=PRIMARY)
        style.map("Primary.TButton",
                  background=[("disabled", PRIMARY_DISABLED), ("pressed", PRIMARY_DARK), ("active", PRIMARY_DARK)],
                  bordercolor=[("disabled", PRIMARY_DISABLED), ("pressed", PRIMARY_DARK), ("active", PRIMARY_DARK)],
                  lightcolor=[("disabled", PRIMARY_DISABLED), ("pressed", PRIMARY_DARK), ("active", PRIMARY_DARK)],
                  darkcolor=[("disabled", PRIMARY_DISABLED), ("pressed", PRIMARY_DARK), ("active", PRIMARY_DARK)],
                  foreground=[("disabled", "#EEF3FA")])

        style.configure("Link.TButton", background=CARD, foreground=PRIMARY, bordercolor=CARD,
                        lightcolor=CARD, darkcolor=CARD, padding=(px(6), px(2)), font=self.font_bold)
        style.map("Link.TButton", background=[("active", "#EEF4FB")], lightcolor=[("active", "#EEF4FB")],
                  darkcolor=[("active", "#EEF4FB")], bordercolor=[("active", "#EEF4FB")],
                  foreground=[("disabled", "#A3ADB9")])

        # Checkbox with a modern tick instead of the theme's cross.
        size = 24 if self.scale >= 1.25 else 16
        self._check_images = {state: tk.PhotoImage(data=CHECKBOX_PNG_BASE64[(size, state)])
                              for state in ("off", "on", "off_disabled", "on_disabled")}
        images = self._check_images
        style.element_create("Tick.indicator", "image", images["off"],
                             ("selected", "disabled", images["on_disabled"]), ("disabled", images["off_disabled"]),
                             ("selected", images["on"]), width=size + px(7), sticky="w")
        style.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
            ("Tick.indicator", {"side": "left", "sticky": ""}),
            ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
                ("Checkbutton.label", {"sticky": "nswe"})]})]})])
        style.configure("TCheckbutton", background=CARD, foreground=TEXT, padding=(0, px(2)))
        style.map("TCheckbutton", background=[("active", CARD)], foreground=[("disabled", "#A3ADB9")])
        style.configure("TEntry", padding=px(5), fieldbackground="white", bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER)
        style.map("TEntry", bordercolor=[("focus", PRIMARY)], lightcolor=[("focus", PRIMARY)],
                  fieldbackground=[("disabled", BUTTON)])
        style.configure("Select.TMenubutton", padding=(px(8), px(5)), background="white", foreground=TEXT,
                        bordercolor=BORDER, lightcolor="white", darkcolor="white", arrowcolor=TEXT,
                        relief="raised")
        style.map("Select.TMenubutton", background=[("disabled", BUTTON), ("active", "#F3F6FA")],
                  lightcolor=[("disabled", BUTTON), ("active", "#F3F6FA")],
                  darkcolor=[("disabled", BUTTON), ("active", "#F3F6FA")],
                  bordercolor=[("active", PRIMARY)], foreground=[("disabled", "#A3ADB9")])

        style.configure("Files.Treeview", rowheight=px(28), background="white", fieldbackground="white",
                        foreground=TEXT, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        style.map("Files.Treeview", background=[("selected", SELECTED)], foreground=[("selected", TEXT)])
        style.configure("Files.Treeview.Heading", background="#E8EEF6", foreground=TEXT, font=self.font_bold,
                        relief="flat", padding=(px(6), px(5)), bordercolor=BORDER,
                        lightcolor="#E8EEF6", darkcolor=BORDER)
        style.map("Files.Treeview.Heading", background=[("active", "#DCE5F0")])
        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#E3E9F1", background=PRIMARY,
                        bordercolor="#E3E9F1", lightcolor=PRIMARY, darkcolor=PRIMARY, thickness=px(10))
        style.configure("Vertical.TScrollbar", background=BUTTON, troughcolor=CARD, bordercolor=CARD,
                        lightcolor=BUTTON, darkcolor=BUTTON, arrowcolor=MUTED)

    def _place_window(self) -> None:
        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        width = min(self.px(1080), screen_w - 40)
        height = min(self.px(820), screen_h - 80)
        self.geometry(f"{width}x{height}+{max(0, (screen_w - width) // 2)}+{max(0, (screen_h - height) // 3)}")
        self.minsize(min(self.px(860), width), min(self.px(600), height))

    # ------------------------------------------------------------------ layout

    def _build(self) -> None:
        px = self.px
        self._build_header()
        content = ttk.Frame(self, padding=(px(16), px(12), px(16), px(14)))
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=11, uniform="bottom")
        content.columnconfigure(1, weight=9, uniform="bottom")
        content.rowconfigure(0, weight=1)

        self._build_files_card(content).grid(row=0, column=0, columnspan=2, sticky="nsew")
        self._build_output_card(content).grid(row=1, column=0, sticky="nsew", pady=(px(12), 0), padx=(0, px(6)))
        self._build_run_card(content).grid(row=1, column=1, sticky="nsew", pady=(px(12), 0), padx=(px(6), 0))

    def _build_header(self) -> None:
        px = self.px
        header = tk.Frame(self, bg=HEADER)
        header.pack(fill="x")
        inner = tk.Frame(header, bg=HEADER)
        inner.pack(fill="x", padx=px(18), pady=px(10))
        tk.Label(inner, image=self._header_icon, bg=HEADER).pack(side="left", padx=(0, px(12)))
        text = tk.Frame(inner, bg=HEADER)
        text.pack(side="left", fill="x", expand=True)
        tk.Label(text, text=APP_NAME, font=self.font_title, bg=HEADER, fg="white", anchor="w").pack(fill="x")
        tk.Label(text, bg=HEADER, fg=HEADER_TEXT, font=self.font_subtitle, anchor="w", justify="left",
                 text="Combine the ECG & NIBP certificates and pass/fail test sheets of many devices "
                      "into Excel and PDF files.").pack(fill="x")
        tk.Label(inner, text=f"v{__version__}", bg=HEADER, fg=HEADER_TEXT, font=self.font_small).pack(
            side="right", anchor="n")

    def _card(self, parent, number: int, title: str, hint: str = ""):
        px = self.px
        outer = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightcolor=BORDER,
                         highlightthickness=1, bd=0)
        head = ttk.Frame(outer, style="Card.TFrame")
        head.pack(fill="x", padx=px(16), pady=(px(10), px(8)))
        size = px(24)
        badge = tk.Canvas(head, width=size, height=size, bg=CARD, highlightthickness=0)
        badge.create_oval(1, 1, size - 1, size - 1, fill=PRIMARY, outline=PRIMARY)
        badge.create_text(size / 2, size / 2, text=str(number), fill="white", font=self.font_bold)
        badge.pack(side="left")
        ttk.Label(head, text=title, style="Section.Card.TLabel").pack(side="left", padx=(px(8), 0))
        if hint:
            ttk.Label(head, text=hint, style="Muted.Card.TLabel").pack(side="left", padx=(px(12), 0), pady=(px(2), 0))
        body = ttk.Frame(outer, style="Card.TFrame")
        body.pack(fill="both", expand=True, padx=px(16), pady=(0, px(12)))
        return outer, body, head

    def _button(self, parent, text: str, command, tooltip: str = "", style: str = "TButton", lock=True, **kwargs):
        button = ttk.Button(parent, text=text, command=command, style=style, takefocus=False, **kwargs)
        if tooltip:
            Tooltip(button, tooltip, self.font_small)
        if lock:
            self._lockable.append(button)
        return button

    def _build_files_card(self, parent) -> tk.Frame:
        px = self.px
        card, body, head = self._card(parent, 1, "Choose the Excel files",
                                      "one workbook per device – the output follows this order")
        self.count_label = ttk.Label(head, style="Muted.Card.TLabel")
        self.count_label.pack(side="right")

        bar = ttk.Frame(body, style="Card.TFrame")
        bar.pack(fill="x", pady=(0, px(8)))
        self._button(bar, "＋  Add files…", self._add_files, "Choose one or more Excel files (Ctrl+O)").pack(side="left")
        self._button(bar, "Add folder…", self._add_folder, "Add every Excel file in a folder").pack(
            side="left", padx=(px(6), 0))
        self._button(bar, "Remove", self._remove_selected, "Remove the selected files from the list (Delete)").pack(
            side="left", padx=(px(18), 0))
        self._button(bar, "Clear", self._clear, "Remove all files from the list").pack(side="left", padx=(px(6), 0))
        self._button(bar, "Sort A–Z", self._sort, "Sort the list by file name").pack(side="right")
        self._button(bar, "▼", lambda: self._move(1), "Move the selected files down", width=3).pack(
            side="right", padx=(0, px(6)))
        self._button(bar, "▲", lambda: self._move(-1), "Move the selected files up", width=3).pack(
            side="right", padx=(0, px(6)))

        table = ttk.Frame(body, style="Card.TFrame")
        table.pack(fill="both", expand=True)
        columns = ("num", "file") + tuple(spec.key for spec in SHEET_SPECS) + ("status",)
        self.tree = ttk.Treeview(table, columns=columns, show="headings", style="Files.Treeview",
                                 selectmode="extended", height=4)
        self.tree.heading("num", text="#")
        self.tree.column("num", width=px(40), minwidth=px(34), anchor="center", stretch=False)
        self.tree.heading("file", text="Excel file", anchor="w")
        self.tree.column("file", width=px(220), minwidth=px(140), anchor="w")
        for spec in SHEET_SPECS:
            self.tree.heading(spec.key, text=spec.label)
            self.tree.column(spec.key, width=px(138), minwidth=px(100), anchor="center", stretch=False)
        self.tree.heading("status", text="Status", anchor="w")
        self.tree.column("status", width=px(200), minwidth=px(120), anchor="w")
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.tag_configure("alt", background=ROW_ALT)
        self.tree.tag_configure("ready", foreground=TEXT)
        self.tree.tag_configure("warning", foreground=WARN)
        self.tree.tag_configure("error", foreground=ERROR)
        self.tree.tag_configure("working", foreground=PRIMARY)
        self.tree.tag_configure("done", foreground=OK)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._update_details())
        self.tree.bind("<Delete>", lambda _e: self._remove_selected())
        self.tree.bind("<Double-1>", self._open_clicked_file)

        self.empty_label = tk.Label(
            table, bg="white", fg=MUTED, font=self.font_subtitle, cursor="hand2", justify="center",
            text="No Excel files yet.\nClick “＋ Add files…” (or here) to choose the device workbooks.")
        self.empty_label.bind("<Button-1>", lambda _e: self._add_files())

        self.details = ttk.Label(body, style="Muted.Card.TLabel", anchor="w", justify="left")
        self.details.pack(fill="x", pady=(px(6), 0))
        body.bind("<Configure>", lambda e: self.details.configure(wraplength=max(200, e.width - px(10))))
        return card

    def _build_output_card(self, parent) -> tk.Frame:
        px = self.px
        card, body, _ = self._card(parent, 2, "Where to save")
        body.columnconfigure(1, weight=1)

        self.output_var = tk.StringVar(value=self.settings.output_dir)
        self.cert_name_var = tk.StringVar(value=self.settings.certificates_name)
        self.pf_name_var = tk.StringVar(value=self.settings.passfail_name)
        self.make_excel_var = tk.BooleanVar(value=self.settings.make_excel)
        self.make_pdf_var = tk.BooleanVar(value=self.settings.make_pdf)
        self.open_folder_var = tk.BooleanVar(value=self.settings.open_folder_when_done)
        self.engine_var = tk.StringVar(value=self.settings.engine)
        self.engine_button: ttk.Menubutton | None = None  # lives in the Options dialog

        ttk.Label(body, text="Folder", style="Field.Card.TLabel").grid(row=0, column=0, sticky="w")
        folder_row = ttk.Frame(body, style="Card.TFrame")
        folder_row.grid(row=0, column=1, sticky="ew", padx=(px(12), 0))
        folder_row.columnconfigure(0, weight=1)
        folder = ttk.Entry(folder_row, textvariable=self.output_var)
        folder.grid(row=0, column=0, sticky="ew", padx=(0, px(6)))
        self._lockable.append(folder)
        self._button(folder_row, "Browse…", self._choose_output_dir, "Choose the folder for the files").grid(
            row=0, column=1)

        rows = (
            (self.cert_name_var, "Certificates",
             "File name for the ECG certificate, then the NIBP certificate, of every device\n"
             "(.xlsx and .pdf are added automatically)"),
            (self.pf_name_var, "Pass/Fail sheets",
             "File name for the ECG pass/fail sheet, then the NIBP pass/fail sheet, of every device\n"
             "(.xlsx and .pdf are added automatically)"),
        )
        for row, (var, label, hint) in enumerate(rows, start=1):
            name_label = ttk.Label(body, text=label, style="Field.Card.TLabel")
            name_label.grid(row=row, column=0, sticky="w", pady=(px(8), 0))
            entry = ttk.Entry(body, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", padx=(px(12), 0), pady=(px(8), 0))
            self._lockable.append(entry)
            Tooltip(name_label, hint, self.font_small)
            Tooltip(entry, hint, self.font_small)

        ttk.Label(body, text="Save as", style="Field.Card.TLabel").grid(row=3, column=0, sticky="w",
                                                                     pady=(px(8), 0))
        kinds = ttk.Frame(body, style="Card.TFrame")
        kinds.grid(row=3, column=1, sticky="ew", padx=(px(12), 0), pady=(px(8), 0))
        excel = ttk.Checkbutton(kinds, text="Excel files", variable=self.make_excel_var, takefocus=False)
        excel.pack(side="left")
        Tooltip(excel, "Certificates.xlsx: sheets “ECG certificates” and “NIBP certificates”.\n"
                       "Pass-Fail file: sheets “ECG reports” and “NIBP reports”.\n"
                       "Each sheet has every device one after another, a page break between them,\n"
                       "and prints exactly like the original tab. The app itself never prints.", self.font_small)
        pdf = ttk.Checkbutton(kinds, text="PDF files", variable=self.make_pdf_var, takefocus=False,
                              command=self._update_engine_note)
        pdf.pack(side="left", padx=(px(14), 0))
        self._lockable += [excel, pdf]
        self.engine_note = tk.Label(kinds, bg=CARD, fg=MUTED, font=self.font_small, anchor="w")
        self.engine_note.pack(side="left", padx=(px(6), 0))

        options = ttk.Frame(body, style="Card.TFrame")
        options.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(px(10), 0))
        check = ttk.Checkbutton(options, text="Open the folder when finished", variable=self.open_folder_var,
                                takefocus=False)
        check.pack(side="left")
        self._lockable.append(check)
        self._button(options, "Options…", self._edit_options,
                     "Tab names to look for, and which program makes the PDFs",
                     style="Link.TButton").pack(side="right")

        self.sheet_vars: dict[str, tk.StringVar] = {}
        for spec in SHEET_SPECS:
            var = tk.StringVar(value=self.settings.sheet_names.get(spec.key, spec.default_name))
            var.trace_add("write", lambda *_: self._schedule_rematch())
            self.sheet_vars[spec.key] = var
        return card

    def _build_run_card(self, parent) -> tk.Frame:
        px = self.px
        card, body, _ = self._card(parent, 3, "Create the files")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(4, weight=1)

        buttons = ttk.Frame(body, style="Card.TFrame")
        buttons.grid(row=0, column=0, sticky="ew")
        self.create_button = self._button(buttons, "Create files", self._start, style="Primary.TButton")
        self.create_button.pack(side="left")
        self.cancel_button = self._button(buttons, "Cancel", self._cancel, lock=False)
        self.cancel_button.pack(side="left", padx=(px(8), 0), fill="y")
        self.cancel_button.state(["disabled"])

        self.progress = ttk.Progressbar(body, style="Accent.Horizontal.TProgressbar", mode="determinate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(px(12), 0))
        self.status_label = ttk.Label(body, text="Add your Excel files, then click “Create files”.",
                                      style="Status.Card.TLabel", justify="left")
        self.status_label.grid(row=2, column=0, sticky="w", pady=(px(4), 0))
        body.bind("<Configure>", lambda e: self.status_label.configure(wraplength=max(150, e.width - px(4))))

        self.results_bar = ttk.Frame(body, style="Card.TFrame")
        ttk.Label(self.results_bar, text="Open:", style="Card.TLabel").pack(side="left")
        self.open_cert_button = self._button(self.results_bar, "Certificates",
                                             lambda: self._open_output(CERTIFICATES), style="Link.TButton")
        self.open_pf_button = self._button(self.results_bar, "Pass/Fail",
                                           lambda: self._open_output(PASS_FAIL), style="Link.TButton")
        self.open_folder_button = self._button(self.results_bar, "Folder", self._open_output_folder,
                                               style="Link.TButton")
        for button in (self.open_cert_button, self.open_pf_button, self.open_folder_button):
            button.pack(side="left", padx=(px(4), 0))

        log_frame = tk.Frame(body, bg=LOG_BG, highlightbackground=BORDER, highlightthickness=1)
        log_frame.grid(row=4, column=0, sticky="nsew", pady=(px(8), 0))
        self.log_text = tk.Text(log_frame, height=3, width=20, wrap="word", bg=LOG_BG, fg=TEXT, relief="flat",
                                bd=0, font=self.font_log, padx=px(8), pady=px(5), state="disabled",
                                cursor="arrow", highlightthickness=0)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self.log_text.tag_configure("time", foreground="#8A97A8")
        self.log_text.tag_configure("info", foreground=TEXT)
        self.log_text.tag_configure("warning", foreground=WARN)
        self.log_text.tag_configure("error", foreground=ERROR)
        self.log_text.tag_configure("success", foreground=OK)
        return card

    def _edit_options(self) -> None:
        px = self.px
        dialog = tk.Toplevel(self)
        dialog.withdraw()
        dialog.title("Options")
        dialog.configure(bg=CARD)
        dialog.transient(self)
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, style="Card.TFrame", padding=px(18))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Tabs to look for in every Excel file", style="Section.Card.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(frame, style="Muted.Card.TLabel", justify="left",
                  text="Capital letters, spaces, dots and dashes are ignored, so\n"
                       "“ECG Certificate” also finds a tab called “ECG.certificate”.\n"
                       "The list of files updates as you type.").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(px(2), px(10)))
        first = None
        for row, spec in enumerate(SHEET_SPECS, start=2):
            ttk.Label(frame, text=spec.label, style="Field.Card.TLabel").grid(row=row, column=0, sticky="w",
                                                                              pady=px(4))
            entry = ttk.Entry(frame, textvariable=self.sheet_vars[spec.key], width=30)
            entry.grid(row=row, column=1, sticky="ew", padx=(px(12), 0), pady=px(4))
            first = first or entry
        row = len(SHEET_SPECS) + 2
        ttk.Button(frame, text="Reset tab names", command=self._reset_tab_names, takefocus=False).grid(
            row=row, column=1, sticky="e", pady=(px(4), 0))

        ttk.Label(frame, text="PDFs are made with", style="Section.Card.TLabel").grid(
            row=row + 1, column=0, columnspan=2, sticky="w", pady=(px(16), 0))
        ttk.Label(frame, style="Muted.Card.TLabel", justify="left",
                  text="Automatic uses Excel when it can save PDFs (Excel 2010 or newer),\n"
                       "otherwise the free LibreOffice. Nothing is ever sent to a printer.").grid(
            row=row + 2, column=0, columnspan=2, sticky="w", pady=(px(2), px(8)))
        self.engine_button = ttk.Menubutton(frame, style="Select.TMenubutton", width=24, takefocus=False)
        menu = tk.Menu(self.engine_button, tearoff=False)
        for key, label in ENGINE_CHOICES:
            menu.add_radiobutton(label=label, value=key, variable=self.engine_var, command=self._update_engine_note)
        self.engine_button.configure(menu=menu)
        self.engine_button.grid(row=row + 3, column=0, columnspan=2, sticky="w")
        self._update_engine_note()

        close = ttk.Button(frame, text="Done", command=dialog.destroy, style="Primary.TButton", takefocus=False)
        close.grid(row=row + 4, column=0, columnspan=2, sticky="e", pady=(px(18), 0))

        def closed(_event=None):
            if _event is None or _event.widget is dialog:
                self.engine_button = None
        dialog.bind("<Destroy>", closed)
        dialog.bind("<Return>", lambda _e: dialog.destroy())
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_reqwidth()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_reqheight()) // 3
        dialog.geometry(f"+{max(0, x)}+{max(0, y)}")
        dialog.deiconify()
        dialog.grab_set()
        first.focus_set()

    # ------------------------------------------------------------------ file list

    def _add_files(self) -> None:
        if self.worker:
            return
        chosen = filedialog.askopenfilenames(
            parent=self, title="Choose the Excel files",
            initialdir=self.settings.last_browse_dir or None,
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if chosen:
            self.settings.last_browse_dir = str(Path(chosen[0]).parent)
            self.add_paths([Path(p) for p in chosen])

    def _add_folder(self) -> None:
        if self.worker:
            return
        folder = filedialog.askdirectory(parent=self, title="Choose a folder with Excel files",
                                         initialdir=self.settings.last_browse_dir or None)
        if not folder:
            return
        self.settings.last_browse_dir = folder
        paths = sorted((p for p in Path(folder).iterdir() if p.is_file() and is_excel_file(p)),
                       key=lambda p: natural_key(p.name))
        if not paths:
            messagebox.showinfo(APP_NAME, "There are no Excel files in that folder.", parent=self)
            return
        self.add_paths(paths)

    def add_paths(self, paths: list[Path]) -> None:
        known = {os.path.normcase(str(e.path)) for e in self.entries}
        skipped = []
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            for path in paths:
                path = path.resolve()
                if not is_excel_file(path):
                    skipped.append(path.name)
                    continue
                if os.path.normcase(str(path)) in known:
                    continue
                known.add(os.path.normcase(str(path)))
                entry = FileEntry(path)
                try:
                    entry.sheet_names = read_sheet_names(path)
                except WorkbookReadError as exc:
                    entry.read_error = str(exc)
                self.entries.append(entry)
        finally:
            self.config(cursor="")
        if not self.output_var.get().strip() and self.entries:
            self.output_var.set(str(self.entries[0].path.parent))
        self._rematch()
        if skipped:
            messagebox.showwarning(APP_NAME, "These files are not Excel workbooks and were skipped:\n\n"
                                   + "\n".join(skipped[:15]), parent=self)

    def _selected_indices(self) -> list[int]:
        return sorted(self.tree.index(item) for item in self.tree.selection())

    def _remove_selected(self) -> None:
        if self.worker:
            return
        for index in reversed(self._selected_indices()):
            del self.entries[index]
        self._refresh_table()

    def _clear(self) -> None:
        if self.worker or not self.entries:
            return
        if messagebox.askyesno(APP_NAME, f"Remove all {len(self.entries)} files from the list?", parent=self):
            self.entries.clear()
            self._refresh_table()

    def _move(self, step: int) -> None:
        if self.worker:
            return
        selected = self._selected_indices()
        if not selected:
            return
        order = selected if step < 0 else list(reversed(selected))
        chosen = set(selected)
        for index in order:
            target = index + step
            if 0 <= target < len(self.entries) and target not in chosen:
                self.entries[index], self.entries[target] = self.entries[target], self.entries[index]
                chosen.discard(index)
                chosen.add(target)
        self._refresh_table(select=sorted(chosen))

    def _sort(self) -> None:
        if self.worker:
            return
        self.entries.sort(key=lambda e: natural_key(e.path.name))
        self._refresh_table()

    def _open_clicked_file(self, event) -> None:
        item = self.tree.identify_row(event.y)
        if item:
            self._safe_open(self.entries[self.tree.index(item)].path)

    def _wanted_names(self) -> dict[str, str]:
        names = {}
        for spec in SHEET_SPECS:
            value = self.sheet_vars[spec.key].get().strip()
            names[spec.key] = value or spec.default_name
        return names

    def _schedule_rematch(self) -> None:
        if self._rematch_job:
            self.after_cancel(self._rematch_job)
        self._rematch_job = self.after(300, self._rematch)

    def _rematch(self) -> None:
        """Work out which tabs each file has and refresh the table."""
        self._rematch_job = None
        wanted = self._wanted_names()
        for entry in self.entries:
            if entry.sheet_names is None:
                entry.found = {}
                entry.state, entry.status = "warning", f"Tabs could not be read ({entry.read_error})"
                continue
            entry.found = match_sheets(wanted, entry.sheet_names)
            missing = [spec.label for spec in SHEET_SPECS if entry.found.get(spec.key) is None]
            if missing:
                entry.state, entry.status = "warning", "Missing: " + ", ".join(missing)
            else:
                entry.state, entry.status = "ready", "Ready"
        self._refresh_table()

    def _refresh_table(self, select: list[int] | None = None) -> None:
        if select is None:
            select = [i for i in self._selected_indices() if i < len(self.entries)] if self.entries else []
        self.tree.delete(*self.tree.get_children())
        for index, entry in enumerate(self.entries):
            self.tree.insert("", "end", values=self._row_values(index, entry), tags=self._row_tags(index, entry))
        children = self.tree.get_children()
        if select and children:
            self.tree.selection_set([children[i] for i in select if i < len(children)])

        if self.entries:
            self.empty_label.place_forget()
        else:
            self.empty_label.place(relx=0.5, rely=0.55, anchor="center")
        found = sum(1 for e in self.entries for v in e.found.values() if v)
        files = len(self.entries)
        self.count_label.configure(
            text=f"{files} file{'s' if files != 1 else ''} · {found} of {files * len(SHEET_SPECS)} tabs found"
            if files else "")
        self._update_details()

    def _row_values(self, index: int, entry: FileEntry) -> tuple:
        marks = []
        for spec in SHEET_SPECS:
            if entry.sheet_names is None:
                marks.append(UNKNOWN)
            else:
                marks.append(FOUND if entry.found.get(spec.key) else MISSING)
        return (index + 1, entry.path.name, *marks, entry.status)

    @staticmethod
    def _row_tags(index: int, entry: FileEntry) -> tuple:
        return (entry.state, "alt") if index % 2 else (entry.state,)

    def _update_row(self, index: int) -> None:
        children = self.tree.get_children()
        if index < len(children):
            entry = self.entries[index]
            self.tree.item(children[index], values=self._row_values(index, entry), tags=self._row_tags(index, entry))

    def _update_details(self) -> None:
        indices = self._selected_indices()
        if len(indices) != 1:
            text = ("Tip: double-click a file to open it in Excel. ✔ = tab found, ✖ = tab missing."
                    if self.entries else "")
        else:
            entry = self.entries[indices[0]]
            if entry.sheet_names is None:
                text = f"{entry.path}\nThe tabs could not be read: {entry.read_error}"
            else:
                text = f"{entry.path}\nTabs in this file: {', '.join(entry.sheet_names) or '(none)'}"
        self.details.configure(text=text)

    # ------------------------------------------------------------------ options

    def _choose_output_dir(self) -> None:
        folder = filedialog.askdirectory(parent=self, title="Save the files in…",
                                         initialdir=self.output_var.get() or self.settings.last_browse_dir or None)
        if folder:
            self.output_var.set(str(Path(folder)))

    def _reset_tab_names(self) -> None:
        for key, name in default_sheet_names().items():
            self.sheet_vars[key].set(name)

    def _engine_choice(self) -> str:
        choice = self.engine_var.get()
        return choice if choice in dict(ENGINE_CHOICES) else ENGINE_AUTO

    def _update_engine_note(self) -> None:
        """Say next to "PDF files" which program will make the PDFs."""
        excel, libre = self.engines.get(ENGINE_EXCEL), self.engines.get(ENGINE_LIBREOFFICE)
        choice = self._engine_choice()
        if self.engine_button is not None:
            self.engine_button.configure(text=dict(ENGINE_CHOICES)[choice])
        self.engine_note.unbind("<Button-1>")
        if not self.make_pdf_var.get():
            self.engine_note.configure(text="", cursor="")
            return
        if choice == ENGINE_AUTO:
            ok = excel or libre
            text = ("with Excel or LibreOffice" if excel and libre else "with LibreOffice" if libre
                    else "with Excel (2010 or newer)" if excel else "needs LibreOffice – click here")
        else:
            ok = self.engines.get(choice)
            name = ENGINE_NAMES[choice]
            text = f"with {name}" if ok else f"{name} not found – click here"
        self.engine_note.configure(text=text, fg=MUTED if ok else ERROR, cursor="" if ok else "hand2")
        if not ok:
            self.engine_note.bind("<Button-1>", lambda _e: self._explain_missing_engine())

    def _explain_missing_engine(self) -> None:
        if messagebox.askyesno(
                "PDFs need LibreOffice",
                "To make PDFs, this app needs Excel 2010 or newer, or the free LibreOffice.\n\n"
                "Excel 2007 cannot save PDFs. The Excel files work without it.\n\n"
                "Open the LibreOffice download page now?", parent=self):
            webbrowser.open(LIBREOFFICE_DOWNLOAD_URL)

    def _collect_settings(self) -> None:
        s = self.settings
        s.output_dir = self.output_var.get().strip()
        s.certificates_name = base_name(self.cert_name_var.get()) or Settings.certificates_name
        s.passfail_name = base_name(self.pf_name_var.get()) or Settings.passfail_name
        s.make_excel = bool(self.make_excel_var.get())
        s.make_pdf = bool(self.make_pdf_var.get())
        s.sheet_names = self._wanted_names()
        s.engine = self._engine_choice()
        s.open_folder_when_done = bool(self.open_folder_var.get())

    # ------------------------------------------------------------------ running

    def _file_base(self, var: tk.StringVar, default: str) -> str | None:
        name = base_name(var.get()) or default
        if INVALID_FILENAME.search(name) or not name.strip(". "):
            return None
        var.set(name)
        return name

    def _start(self) -> None:
        if self.worker:
            return
        self._collect_settings()
        if not self.entries:
            messagebox.showinfo(APP_NAME, "Add at least one Excel file first.", parent=self)
            return
        make_excel, make_pdf = self.settings.make_excel, self.settings.make_pdf
        if not make_excel and not make_pdf:
            messagebox.showinfo(APP_NAME, "Tick “Excel files”, “PDF files” or both.", parent=self)
            return
        gone = [e.path.name for e in self.entries if not e.path.exists()]
        if gone:
            messagebox.showerror(APP_NAME, "These files no longer exist – remove them from the list:\n\n"
                                 + "\n".join(gone[:15]), parent=self)
            return

        folder = self.output_var.get().strip() or str(self.entries[0].path.parent)
        self.output_var.set(folder)
        cert_name = self._file_base(self.cert_name_var, Settings.certificates_name)
        pf_name = self._file_base(self.pf_name_var, Settings.passfail_name)
        if not cert_name or not pf_name:
            messagebox.showerror(APP_NAME, 'A file name is not valid. Names cannot contain  < > : " / \\ | ? *',
                                 parent=self)
            return
        if cert_name.lower() == pf_name.lower():
            messagebox.showerror(APP_NAME, "The certificates and the pass/fail sheets need different file names.",
                                 parent=self)
            return

        self.engines = available_engines()
        self._update_engine_note()
        choice = self._engine_choice()
        if make_pdf and not make_excel and not (self.engines.get(choice) if choice != ENGINE_AUTO
                                                else any(self.engines.values())):
            self._explain_missing_engine()
            return

        try:
            Path(folder).mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"The folder could not be created:\n{folder}\n\n{exc}", parent=self)
            return
        names = {CERTIFICATES: cert_name, PASS_FAIL: pf_name}
        excel_outputs = {g: Path(folder) / f"{n}.xlsx" for g, n in names.items()} if make_excel else {}
        pdf_outputs = {g: Path(folder) / f"{n}.pdf" for g, n in names.items()} if make_pdf else {}
        existing = [p.name for p in (*excel_outputs.values(), *pdf_outputs.values()) if p.exists()]
        if existing and not messagebox.askyesno(
                APP_NAME, "These files already exist in the folder:\n\n" + "\n".join(existing)
                + "\n\nReplace them?", icon="warning", parent=self):
            return

        missing = sum(1 for e in self.entries for v in e.found.values() if v is None)
        files_missing = sum(1 for e in self.entries if any(v is None for v in e.found.values()))
        if missing and not messagebox.askyesno(
                APP_NAME, f"{missing} tab(s) are missing in {files_missing} file(s) and will be left out.\n\n"
                          "Continue anyway?", parent=self):
            return

        self.settings.save()
        for index, entry in enumerate(self.entries):
            entry.state, entry.status = "ready", "Waiting…"
            self._update_row(index)
        self._clear_log()
        self.results_bar.grid_forget()
        self.progress.configure(value=0, maximum=1000)
        self.status_label.configure(text="Starting…", foreground=MUTED)
        self._set_running(True)

        self.cancel_event.clear()
        files = [e.path for e in self.entries]
        self.worker = threading.Thread(target=self._work, name="builder", daemon=True,
                                       args=(files, self._wanted_names(), excel_outputs, pdf_outputs, choice))
        self.worker.start()

    def _work(self, files, sheet_names, excel_outputs, pdf_outputs, engine_choice) -> None:
        reporter = QueueReporter(self.events)
        stages = int(bool(excel_outputs)) + int(bool(pdf_outputs))
        run = RunResult()
        try:
            if excel_outputs:
                run.excel = build_workbooks(files, sheet_names, excel_outputs,
                                            StageReporter(reporter, 0, stages), self.cancel_event)
            if pdf_outputs:
                try:
                    with open_engine(engine_choice, reporter.log) as engine:
                        run.pdf = build_pdfs(files, sheet_names, pdf_outputs, engine,
                                             StageReporter(reporter, stages - 1, stages), self.cancel_event)
                except EngineError as exc:  # no program on this PC can make PDFs
                    run.pdf_error = str(exc)
                    reporter.log("warning", f"PDF files were not made: {exc}")
            self.events.put(("finished", run))
        except Cancelled:
            self.events.put(("cancelled",))
        except OutputError as exc:
            self.events.put(("failed", str(exc)))
        except Exception as exc:  # keep the window alive whatever happens
            self.events.put(("log", "error", traceback.format_exc().strip()))
            self.events.put(("failed", f"Unexpected error: {exc}"))

    def _cancel(self) -> None:
        if self.worker:
            self.cancel_event.set()
            self.cancel_button.state(["disabled"])
            self.status_label.configure(text="Stopping after the current file…")

    def _set_running(self, running: bool) -> None:
        for widget in self._lockable:
            widget.state(["disabled"] if running else ["!disabled"])
        self.cancel_button.state(["!disabled"] if running else ["disabled"])
        self.config(cursor="watch" if running else "")

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "log":
                    self._log(event[1], event[2])
                elif kind == "progress":
                    _, done, total, text = event
                    self.progress.configure(maximum=total, value=done)
                    self.status_label.configure(text=text, foreground=MUTED)
                elif kind == "file":
                    _, index, state, text = event
                    if index < len(self.entries):
                        self.entries[index].state, self.entries[index].status = state, text
                        self._update_row(index)
                elif kind == "finished":
                    self._finished(event[1])
                elif kind == "cancelled":
                    self._stopped("Cancelled.", WARN)
                elif kind == "failed":
                    self._stopped("Failed – see the messages below.", ERROR)
                    self._log("error", event[1])
                    if not self.closing:
                        messagebox.showerror("The files could not be created", event[1], parent=self)
        except queue.Empty:
            pass
        if self.closing and self.worker is None:
            self.destroy()
            return
        self.after(80, self._drain_events)

    def _stopped(self, text: str, colour: str) -> None:
        self.worker = None
        self._set_running(False)
        for index, entry in enumerate(self.entries):
            if entry.state == "ready":  # never reached
                entry.status = "Not processed"
                self._update_row(index)
        self.status_label.configure(text=text, foreground=colour)
        self._log("warning" if colour == WARN else "error", text)

    def _finished(self, run: RunResult) -> None:
        self.worker = None
        self._set_running(False)
        if self.closing:
            return
        self.progress.configure(value=self.progress["maximum"])
        results = [r for r in (run.excel, run.pdf) if r]
        self.last_outputs = {}
        for result in results:  # PDFs come last, so they win for the "Open" links
            self.last_outputs.update({group: out.path for group, out in result.outputs.items() if out})

        # One status per file, from both steps
        failed_files, missing = [], 0
        for index, entry in enumerate(self.entries):
            per_step = [r.workbooks[index] for r in results if index < len(r.workbooks)]
            problems = [p for w in per_step for p in w.problems]
            tabs_missing = max((len(w.missing) for w in per_step), default=0)
            missing += tabs_missing
            if problems:
                failed_files.append(entry.path.name)
                entry.state, entry.status = "error", "Failed – see the messages below"
            elif tabs_missing:
                entry.state, entry.status = "warning", f"Done – {tabs_missing} tab(s) missing"
            else:
                entry.state, entry.status = "done", "Done"
            self._update_row(index)

        saved = [out for r in results for out in r.outputs.values() if out]
        lines = [f"• {out.path.name} – {out.tabs} tabs, {out.pages} pages" if out.pages
                 else f"• {out.path.name} – {out.tabs} tabs, ECG and NIBP on their own sheet"
                 for out in saved]
        notes = []
        if run.pdf_error:
            notes.append(f"PDF files were not made: {run.pdf_error}.")
        if missing:
            notes.append(f"{missing} tab(s) were missing and left out.")
        if failed_files:
            notes.append(f"{len(failed_files)} file(s) had problems: {', '.join(failed_files)}. "
                         "See the messages in the window.")
        warn = bool(notes)

        if saved:
            summary = "Finished with warnings" if warn else "Finished"
            self.status_label.configure(
                text=f"{'⚠' if warn else '✔'}  {summary} – {len(saved)} file{'s' if len(saved) != 1 else ''} saved.",
                foreground=WARN if warn else OK)
            self.open_cert_button.state(["!disabled"] if CERTIFICATES in self.last_outputs else ["disabled"])
            self.open_pf_button.state(["!disabled"] if PASS_FAIL in self.last_outputs else ["disabled"])
            self.results_bar.grid(row=3, column=0, sticky="w", pady=(self.px(4), 0))
            message = f"Saved in {saved[0].path.parent}:\n\n" + "\n".join(lines)
            if notes:
                message += "\n\n" + "\n\n".join(notes)
            if run.pdf_error and not self.engines.get(ENGINE_LIBREOFFICE):
                if messagebox.askyesno("Files created", message + "\n\nTo get PDFs as well, install LibreOffice "
                                       "(free). Open the download page now?", icon="warning", parent=self):
                    webbrowser.open(LIBREOFFICE_DOWNLOAD_URL)
            else:
                (messagebox.showwarning if warn else messagebox.showinfo)("Files created", message, parent=self)
            if self.open_folder_var.get():
                self._open_output_folder()
        else:
            self.status_label.configure(text="Nothing was saved – see the messages below.", foreground=ERROR)
            hint = ("Check the tab names under “Options…”." if not failed_files and not run.pdf_error
                    else "See the messages in the window for details.")
            messagebox.showerror(APP_NAME, "\n\n".join(["No file was created.", *notes, hint]), parent=self)

    def _open_output(self, group: str) -> None:
        path = self.last_outputs.get(group)
        if path:
            self._safe_open(path)

    def _open_output_folder(self) -> None:
        paths = list(self.last_outputs.values())
        if paths:
            self._safe_open(paths[0], reveal=True)
        elif self.output_var.get().strip():
            self._safe_open(Path(self.output_var.get().strip()))

    def _safe_open(self, path: Path, reveal: bool = False) -> None:
        try:
            open_path(path, reveal)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Could not open {path}\n\n{exc}", parent=self)

    # ------------------------------------------------------------------ log

    def _log(self, level: str, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", time.strftime("%H:%M:%S  "), "time")
        self.log_text.insert("end", text + "\n", level if level in ("warning", "error", "success") else "info")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ------------------------------------------------------------------ closing

    def _on_close(self) -> None:
        self._collect_settings()
        self.settings.save()
        if self.worker:
            if not messagebox.askyesno(APP_NAME, "The files are still being created. Stop and close?", parent=self):
                return
            self.closing = True
            self._cancel()
            return  # _drain_events closes the window once the worker has stopped
        self.destroy()
