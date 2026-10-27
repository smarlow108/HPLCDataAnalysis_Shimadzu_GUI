"""Shimadzu calibration GUI with ten calibration-file slots.

Select up to ten spectrum files in one operation, or load one LabSolutions
ASCII batch export containing Water plus Cal 1 ... Cal 10. The Water record is
kept separately as a blank/reference and the calibration records are assigned
to Cal 1 ... Cal 10. Choose a calibration point to display its spectrum,
select a peak region by dragging, and store the integrated area.

The reader accepts two-column text/CSV/ASCII data and identifies Shimadzu
LabSolutions compound-file (OLE/CFB) data files, including files exported or
renamed with a .txt extension. Binary LabSolutions signal decoding is isolated
in read_labsolutions_file() so the GUI can report a clear diagnostic when the
required decoder is unavailable.
"""

from __future__ import annotations

import csv
import os
import re
import struct
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib.widgets import SpanSelector
from matplotlib.ticker import MultipleLocator


ANALYTES = [
    "Glycerol", "Glyceric acid", "Glycolic acid", "Oxalic acid",
    "Formic acid", "Lactic acid", "Tartronic acid",
]
CALIBRATIONS = [f"Cal {i}" for i in range(1, 11)]
DETECTORS = ["PDA 210 nm", "ELSD"]

CONCENTRATIONS = {
    "Glycerol":       [0.0250, 0.0800, 0.1350, 0.1900, 0.2500, 0.0250, 0.0800, 0.1350, 0.1900, 0.2500],
    "Glyceric acid":  [0.0500, 0.0350, 0.0250, 0.0150, 0.0050, None, None, None, None, None],
    "Glycolic acid":  [0.0500, 0.0350, 0.0250, 0.0150, 0.0050, None, None, None, None, None],
    "Oxalic acid":    [0.0050, 0.0150, 0.0250, 0.0350, 0.0500, None, None, None, None, None],
    "Formic acid":    [0.0050, 0.0150, 0.0250, 0.0350, 0.0500, None, None, None, None, None],
    "Lactic acid":    [0.0500, 0.0350, 0.0250, 0.0150, 0.0050, 0.0500, 0.0350, 0.0250, 0.0150, 0.0050],
    "Tartronic acid": [None, None, None, None, None, 0.0500, 0.0350, 0.0250, 0.0150, 0.0050],
}

RETENTION_TIMES = {
    "Oxalic acid": ("PDA 210 nm", 6.6),
    "Tartronic acid": ("PDA 210 nm", 8.1),
    "Glyceric acid": ("PDA 210 nm", 10.8),
    "Glycolic acid": ("PDA 210 nm", 12.3),
    "Lactic acid": ("PDA 210 nm", 12.7),
    "Formic acid": ("PDA 210 nm", 13.8),
    "Glycerol": ("ELSD", 13.2),
}


def compound_label(compound: str) -> str:
    expected = RETENTION_TIMES.get(compound)
    if expected is None:
        return compound
    return f"{compound} ({expected[1]:g} min)"


def compound_from_label(label: str) -> str:
    return str(label).split(" (", 1)[0]


def read_labsolutions_file(filename: str) -> tuple[np.ndarray, np.ndarray]:
    """Read a LabSolutions OLE/CFB file when a compatible decoder is available.

    These .txt files are actually binary LabSolutions .lcd compound files. The
    file structure identifies ELSD and PDA streams, but the numerical stream
    packing is vendor-specific and must not be guessed. Install the optional
    decoder package used by your lab, or export the chromatogram as ASCII/CSV.
    """
    with open(filename, "rb") as fh:
        header = fh.read(8)
    if header != bytes.fromhex("D0CF11E0A1B11AE1"):
        raise ValueError("Not a LabSolutions compound file.")
    raise ValueError(
        "LabSolutions binary .lcd detected. This file contains ELSD and PDA "
        "streams, but the numerical stream decoder is not included yet. "
        "Export ELSD/PDA 210 nm as ASCII/CSV, or provide the lab's validated "
        "LabSolutions decoder/API."
    )


def _ascii_section(text: str, section: str) -> tuple[np.ndarray, np.ndarray]:
    """Extract the time/intensity table from a LabSolutions ASCII section."""
    start = text.find(section)
    if start < 0:
        raise ValueError(f"ASCII section not found: {section}")
    end = text.find("\n[", start + len(section))
    block = text[start:] if end < 0 else text[start:end]
    rows = []
    data_started = False
    for line in block.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("R.Time (min)"):
            data_started = True
            continue
        if not data_started:
            continue
        fields = re.split(r"[;,\t ]+", line)
        if len(fields) < 2:
            continue
        try:
            rows.append((float(fields[0]), float(fields[1])))
        except ValueError:
            continue
    if len(rows) < 2:
        raise ValueError(f"No numerical data found in {section}")
    arr = np.asarray(rows, dtype=float)
    return arr[:, 0], arr[:, 1]


def split_ascii_batch(text: str) -> list[str]:
    """Split a LabSolutions ASCII batch export into individual sample records."""
    starts = [m.start() for m in re.finditer(r"(?m)^\[Header\]\s*$", text)]
    if not starts:
        return [text]
    return [text[a:b] for a, b in zip(starts, starts[1:] + [len(text)])]


def ascii_record_label(text: str, fallback: str) -> str:
    """Return the calibration slot from the record's Sample Name/metadata.

    The Sample Name is authoritative.  This prevents a file selected in an
    arbitrary order from being assigned to the wrong calibration slot.
    """
    if re.search(
        r"^\s*Sample Name\s*[:=]?\s*WATER\b|"
        r"^\s*Sample Type\s*[:=]?\s*.*blank",
        text, re.MULTILINE | re.IGNORECASE,
    ):
        return "Water"

    patterns = (
        # Preferred source: the Sample Name field.
        r"^\s*Sample Name\s*[:=]?\s*.*?\bCAL\s*[-_ ]?(\d+)\b",
        # Other LabSolutions metadata fields.
        r"^\s*Data File Name\s*[:=]?\s*.*?[_-]?CAL\s*[-_ ]?(\d+)(?:[_\. ]|$)",
        r"^\s*Sample ID\s*[:=]?\s*.*?[_-]?CAL\s*[-_ ]?(\d+)\b",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match:
            number = int(match.group(1))
            if 1 <= number <= 10:
                return f"Cal {number}"
    return fallback


def read_spectrum(source, detector: str = "PDA 210 nm") -> tuple[np.ndarray, np.ndarray]:
    """Read one ASCII record or file, selecting PDA 210 nm or ELSD."""
    if isinstance(source, tuple):
        text = source[0]
        raw = None
    else:
        with open(source, "rb") as fh:
            raw = fh.read()
        if raw.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
            return read_labsolutions_file(source)
        text = raw.decode("utf-8-sig", errors="replace")

    if "[LC Chromatogram(ELSD)]" in text or "[PDA Multi Chromatogram" in text:
        section = "[LC Chromatogram(ELSD)]" if detector == "ELSD" else "[PDA Multi Chromatogram(Ch1)]"
        return _ascii_section(text, section)

    rows = []
    for line in text.splitlines():
        fields = re.split(r"[;,\\t ]+", line.strip())
        nums = []
        for field in fields:
            try:
                nums.append(float(field.replace(",", ".")))
            except ValueError:
                continue
        if len(nums) >= 2:
            rows.append((nums[0], nums[1]))
    if len(rows) < 2:
        raise ValueError("No two-column numeric spectrum or LabSolutions ASCII channel was found.")
    arr = np.asarray(rows, dtype=float)
    arr = arr[np.argsort(arr[:, 0])]
    unique = np.r_[True, np.diff(arr[:, 0]) != 0]
    return arr[unique, 0], arr[unique, 1]


def integrate_peak(
    x: np.ndarray,
    y: np.ndarray,
    left: float,
    right: float,
    baseline_left: float | None = None,
    baseline_right: float | None = None,
) -> float:
    if right <= left:
        left, right = right, left
        baseline_left, baseline_right = baseline_right, baseline_left
    mask = (x >= left) & (x <= right)
    if mask.sum() < 2:
        raise ValueError("The selected region contains fewer than two data points.")
    xx, yy = x[mask], y[mask]
    if baseline_left is None:
        baseline_left = float(np.interp(left, x, y))
    if baseline_right is None:
        baseline_right = float(np.interp(right, x, y))
    baseline = np.interp(xx, [left, right], [baseline_left, baseline_right])
    return float(np.trapz(yy - baseline, xx))



class CalibrationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Shimadzu HPLC Calibration")
        self.geometry("1450x900")
        self.minsize(1100, 720)

        self.files: dict[str, str | tuple[str, str]] = {}
        self.water_record = None
        self.x = self.y = None
        self.selector = None
        self.overlay_selector = None
        self.baseline_selector = None
        self.overlay_spectra = {}
        self.peak_range = None
        self.baseline_range = None
        self.overlay_peak_artist = None
        self.overlay_baseline_artist = None
        self.areas = {(d, a, c): "" for d in DETECTORS for a in ANALYTES for c in CALIBRATIONS}
        self.concentrations = {a: list(vals) for a, vals in CONCENTRATIONS.items()}
        self.concentration_vars = {}
        self.last_data_folder = None
        self.overlay_mode = "peak"
        self.individual_compound_vars = {}
        self.individual_rt_artists = {}
        self.overlay_rt_artists = []
        self._drag_rt_compound = None
        self._drag_rt_detector = None

        self._build_ui()
        self.update_rt_label()
        self.refresh_concentration_table()
        self.refresh_overlay()

    def _build_ui(self):
        self._build_file_bar()
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=5)
        self.page1 = ttk.Frame(self.tabs)
        self.page2 = ttk.Frame(self.tabs)
        self.page3 = ttk.Frame(self.tabs)
        self.tabs.add(self.page1, text="1. Individual curves")
        self.tabs.add(self.page2, text="2. Overlay and integration")
        self.tabs.add(self.page3, text="3. Calibration curves")
        self._build_page1()
        self._build_page2()
        self._build_page3()

    def _build_file_bar(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(fill="x")
        ttk.Button(bar, text="Open calibration files…", command=self.open_files).pack(side="left")
        ttk.Button(bar, text="Open batch ASCII…", command=self.open_batch_ascii).pack(side="left", padx=5)
        self.file_label = ttk.Label(bar, text="No files loaded", width=38)
        self.file_label.pack(side="left", padx=8)
        ttk.Label(bar, text="Detector:").pack(side="left")
        self.detector = ttk.Combobox(bar, values=DETECTORS, state="readonly", width=14)
        self.detector.current(0)
        self.detector.pack(side="left", padx=5)
        self.detector.bind("<<ComboboxSelected>>", lambda e: self._detector_changed())
        ttk.Button(bar, text="Export Excel…", command=self.export_excel).pack(side="right")
        self.status = ttk.Label(self, text="Load calibration files to begin.", padding=(8, 0))
        self.status.pack(fill="x")

    def _build_page1(self):
        controls = ttk.Frame(self.page1, padding=(4, 4, 4, 0))
        controls.pack(fill="x")
        ttk.Label(controls, text="Calibration file:").pack(side="left")
        self.cal = ttk.Combobox(controls, values=CALIBRATIONS, state="readonly", width=8)
        self.cal.current(0)
        self.cal.pack(side="left", padx=5)
        self.cal.bind("<<ComboboxSelected>>", lambda e: self.load_selected_file())
        ttk.Label(controls, text="Analyte:").pack(side="left", padx=(15, 0))
        self.analyte = ttk.Combobox(controls, values=ANALYTES, state="readonly", width=18)
        self.analyte.current(0)
        self.analyte.pack(side="left", padx=5)
        self.analyte.bind("<<ComboboxSelected>>", lambda e: self.update_rt_label())
        self.rt_label = ttk.Label(controls, text="")
        self.rt_label.pack(side="left", padx=8)
        ttk.Button(controls, text="Clear selected area", command=self.clear_selected).pack(side="right")

        rt_frame = ttk.LabelFrame(self.page1, text="Retention-time lines (drag a line on the plot to update its value)", padding=4)
        rt_frame.pack(fill="x", padx=4, pady=(4, 0))
        for compound in ANALYTES:
            var = tk.BooleanVar(value=True)
            self.individual_compound_vars[compound] = var
            ttk.Checkbutton(
                rt_frame, text=compound, variable=var,
                command=self._retention_lines_changed
            ).pack(side="left", padx=(0, 8))

        self.figure = Figure(figsize=(9, 4.4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.grid(alpha=.25)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.page1)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.canvas, self.page1).update()
        self.canvas.mpl_connect("button_press_event", self._rt_press)
        self.canvas.mpl_connect("motion_notify_event", self._rt_motion)
        self.canvas.mpl_connect("button_release_event", self._rt_release)
        self.selector = SpanSelector(self.ax, self.on_individual_select, "horizontal",
                                     useblit=True, props=dict(alpha=.25, facecolor="tab:orange"),
                                     interactive=True)

        frame = ttk.LabelFrame(self.page1, text="Editable concentrations and individual integrated areas", padding=4)
        frame.pack(fill="x", padx=4, pady=4)
        self.conc_tree = ttk.Treeview(frame, columns=["sample"] + CALIBRATIONS, show="headings", height=8)
        self.conc_tree.heading("sample", text="Sample Name")
        self.conc_tree.column("sample", width=145, anchor="w")
        for c in CALIBRATIONS:
            self.conc_tree.heading(c, text=c)
            self.conc_tree.column(c, width=95, anchor="center")
        for a in ANALYTES:
            self.conc_tree.insert("", "end", iid=a, values=[a] + [""] * 10)
        self.conc_tree.pack(fill="x")
        self.conc_tree.bind("<Double-1>", self.edit_concentration_cell)
        ttk.Label(frame, text="Double-click a concentration cell to edit it. Values are shared by all pages.",
                  foreground="#555").pack(anchor="w")

    def _build_page2(self):
        controls = ttk.Frame(self.page2, padding=4)
        controls.pack(fill="x")
        ttk.Label(controls, text="Assign selected peak to:").pack(side="left")
        self.overlay_analyte = ttk.Combobox(controls, values=[compound_label(a) for a in ANALYTES], state="readonly", width=26)
        self.overlay_analyte.current(0)
        self.overlay_analyte.pack(side="left", padx=5)
        ttk.Button(controls, text="1 Select peak range", command=lambda: self.set_overlay_mode("peak")).pack(side="left", padx=3)
        ttk.Button(controls, text="2 Select baseline range", command=lambda: self.set_overlay_mode("baseline")).pack(side="left", padx=3)
        ttk.Button(controls, text="3 Integrate all curves", command=self.integrate_overlay_peak).pack(side="left", padx=3)
        ttk.Button(controls, text="Reset selection", command=self.reset_overlay_selection).pack(side="left", padx=3)
        ttk.Button(controls, text="Refresh overlay", command=self.refresh_overlay).pack(side="left", padx=3)
        ttk.Button(controls, text="Clear assigned area", command=self.clear_overlay_area).pack(side="left", padx=3)
        self.overlay_status = ttk.Label(controls, text="Select a peak range.")
        self.overlay_status.pack(side="left", padx=10)

        self.overlay_figure = Figure(figsize=(9, 4.3), dpi=100)
        self.overlay_ax = self.overlay_figure.add_subplot(111)
        self.overlay_ax.grid(alpha=.25)
        self.overlay_canvas = FigureCanvasTkAgg(self.overlay_figure, master=self.page2)
        self.overlay_canvas.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.overlay_canvas, self.page2).update()
        self.overlay_selector = SpanSelector(self.overlay_ax, self.on_overlay_select, "horizontal",
                                             useblit=True, props=dict(alpha=.25, facecolor="tab:orange"),
                                             interactive=True)

        frame = ttk.LabelFrame(self.page2, text="Integrated peak areas", padding=4)
        frame.pack(fill="x", padx=4, pady=4)
        cols = ["sample"] + CALIBRATIONS
        self.area_tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
        self.area_tree.heading("sample", text="Compound")
        self.area_tree.column("sample", width=145, anchor="w")
        for c in CALIBRATIONS:
            self.area_tree.heading(c, text=c)
            self.area_tree.column(c, width=95, anchor="center")
        for a in ANALYTES:
            self.area_tree.insert("", "end", iid=a, values=[a] + [""] * 10)
        self.area_tree.pack(fill="x")

    def _build_page3(self):
        top = ttk.Frame(self.page3, padding=4)
        top.pack(fill="x")
        ttk.Label(top, text="Compound:").pack(side="left")
        self.curve_analyte = ttk.Combobox(top, values=ANALYTES, state="readonly", width=20)
        self.curve_analyte.current(0)
        self.curve_analyte.pack(side="left", padx=5)
        self.curve_analyte.bind("<<ComboboxSelected>>", lambda e: self.refresh_curve_page())
        self.curve_info = ttk.Label(top, text="")
        self.curve_info.pack(side="left", padx=10)
        self.curve_figure = Figure(figsize=(9, 4.5), dpi=100)
        self.curve_ax = self.curve_figure.add_subplot(111)
        self.curve_canvas = FigureCanvasTkAgg(self.curve_figure, master=self.page3)
        self.curve_canvas.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.curve_canvas, self.page3).update()
        frame = ttk.LabelFrame(self.page3, text="Concentration versus integrated area", padding=4)
        frame.pack(fill="x", padx=4, pady=4)
        self.curve_tree = ttk.Treeview(frame, columns=["cal", "concentration", "area"], show="headings", height=10)
        for col, heading, width in [("cal", "Calibration", 120), ("concentration", "Concentration", 150), ("area", "Integrated area", 180)]:
            self.curve_tree.heading(col, text=heading)
            self.curve_tree.column(col, width=width, anchor="center")
        self.curve_tree.pack(fill="x")

    def _detector_changed(self):
        self.load_selected_file()
        self.refresh_overlay()
        self.refresh_concentration_table()
        self.refresh_curve_page()

    def _source_name(self, source):
        return Path(source[1] if isinstance(source, tuple) else source).name

    def _source_folder(self, source):
        return str(Path(source[1] if isinstance(source, tuple) else source).parent)

    def open_files(self):
        selected = filedialog.askopenfilenames(
            title="Select calibration files", filetypes=[("Spectra", "*.lcd *.csv *.txt *.asc *.dat"), ("All files", "*.*")]
        )
        if not selected:
            return
        self.files, self.water_record = {}, None
        used = set()
        for filename in selected:
            try:
                with open(filename, "r", encoding="utf-8-sig", errors="replace") as fh:
                    record = fh.read()

                # Use Sample Name (for example, "Sample Name CAL9") before
                # considering the filename.  Selection order is irrelevant.
                label = ascii_record_label(record, "")
                if not label:
                    match = re.search(
                        r"\bCAL\s*[-_ ]?(\d+)\b",
                        Path(filename).stem,
                        re.IGNORECASE,
                    )
                    if match and 1 <= int(match.group(1)) <= 10:
                        label = f"Cal {int(match.group(1))}"

                if label == "Water":
                    self.water_record = (record, filename)
                elif label in CALIBRATIONS and label not in used:
                    self.files[label] = (record, filename)
                    used.add(label)
                else:
                    raise ValueError
            except (OSError, ValueError):
                # Files without identifying metadata retain the old
                # sequential fallback, but labelled files always win.
                unassigned = [c for c in CALIBRATIONS if c not in self.files]
                if unassigned:
                    self.files[unassigned[0]] = filename
        if self.files:
            self.last_data_folder = self._source_folder(next(iter(self.files.values())))
        self._after_files_loaded()

    def open_batch_ascii(self):
        filename = filedialog.askopenfilename(
            title="Select LabSolutions multi-sample ASCII export",
            filetypes=[("ASCII/text exports", "*.txt *.asc *.dat *.csv"), ("All files", "*.*")]
        )
        if not filename:
            return
        try:
            with open(filename, "r", encoding="utf-8-sig", errors="replace") as fh:
                text = fh.read()
            self.files, self.water_record = {}, None
            next_cal = 1
            for record in split_ascii_batch(text)[:11]:
                label = ascii_record_label(record, f"Cal {next_cal}")
                if label == "Water":
                    self.water_record = (record, filename)
                    continue
                while label in self.files and next_cal <= 10:
                    next_cal += 1
                if label not in CALIBRATIONS:
                    label = f"Cal {next_cal}"
                self.files[label] = (record, filename)
                next_cal = max(next_cal, int(label.split()[-1]) + 1)
            self.last_data_folder = str(Path(filename).parent)
            self._after_files_loaded()
        except Exception as exc:
            messagebox.showerror("Could not read batch ASCII file", str(exc))

    def _after_files_loaded(self):
        loaded = [c for c in CALIBRATIONS if c in self.files]
        self.cal.set(loaded[0] if loaded else CALIBRATIONS[0])
        self.load_selected_file()
        self.refresh_overlay()
        self.status.config(text=f"Loaded {len(loaded)} calibration file(s).")
        self.file_label.config(text=", ".join(loaded) if loaded else "No readable files")

    def load_selected_file(self):
        cal = self.cal.get()
        source = self.files.get(cal)
        if not source:
            self.x = self.y = None
            self.file_label.config(text=f"{cal}: no file")
            self.ax.clear()
            self.ax.grid(alpha=.25)
            self.canvas.draw_idle()
            return
        try:
            self.x, self.y = read_spectrum(source, self.detector.get())
        except Exception as exc:
            self.x = self.y = None
            messagebox.showerror(f"Could not read {cal}", str(exc))
            return
        self.file_label.config(text=f"{cal}: {self._source_name(source)}")
        self.ax.clear()
        self.ax.plot(self.x, self.y, color="tab:blue", lw=1)
        self.ax.set_title(f"{cal} — {self.detector.get()}")
        self.ax.set_xlabel("Retention time / X")
        self.ax.set_ylabel("Signal")
        self.ax.xaxis.set_major_locator(MultipleLocator(1))
        self.ax.grid(alpha=.25)
        self._draw_individual_retention_lines()
        self.canvas.draw_idle()
        self.update_rt_label()

    def _refresh_compound_labels(self):
        if hasattr(self, "overlay_analyte"):
            current = compound_from_label(self.overlay_analyte.get()) if self.overlay_analyte.get() else ANALYTES[0]
            self.overlay_analyte.configure(values=[compound_label(a) for a in ANALYTES])
            self.overlay_analyte.set(compound_label(current))

    def _retention_lines_changed(self):
        self._draw_individual_retention_lines()
        self._draw_overlay_retention_lines()
        self.canvas.draw_idle()
        self.overlay_canvas.draw_idle()

    def _draw_individual_retention_lines(self):
        if not hasattr(self, "ax"):
            return
        for artists in self.individual_rt_artists.values():
            for artist in (artists if isinstance(artists, tuple) else (artists,)):
                try:
                    artist.remove()
                except Exception:
                    pass
        self.individual_rt_artists = {}
        detector = self.detector.get()
        for compound, var in self.individual_compound_vars.items():
            expected = RETENTION_TIMES.get(compound)
            if var.get() and expected and expected[0] == detector:
                rt = expected[1]
                line = self.ax.axvline(
                    rt, color="tab:red", linestyle="--", linewidth=1.2, alpha=.85, zorder=4
                )
                label = self.ax.text(rt, 0.98, f"{compound} ({rt:g} min)",
                                     rotation=90, color="tab:red",
                                     transform=self.ax.get_xaxis_transform(),
                                     ha="right", va="top", fontsize=8, clip_on=True, zorder=5)
                self.individual_rt_artists[compound] = (line, label)
        self.ax.xaxis.set_major_locator(MultipleLocator(1))

    def _draw_overlay_retention_lines(self):
        if not hasattr(self, "overlay_ax"):
            return
        for artist in self.overlay_rt_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self.overlay_rt_artists = []
        detector = self.detector.get()
        for compound, var in self.individual_compound_vars.items():
            expected = RETENTION_TIMES.get(compound)
            if var.get() and expected and expected[0] == detector:
                rt = expected[1]
                self.overlay_rt_artists.append(self.overlay_ax.axvline(
                    rt, color="tab:red", linestyle="--", linewidth=1.1, alpha=.85, zorder=4
                ))
                self.overlay_rt_artists.append(self.overlay_ax.text(
                    rt, 0.98, f"{compound} ({rt:g} min)",
                    rotation=90, color="tab:red",
                    transform=self.overlay_ax.get_xaxis_transform(),
                    ha="right", va="top", fontsize=8, clip_on=True, zorder=5
                ))
        self.overlay_ax.xaxis.set_major_locator(MultipleLocator(1))

    def _rt_press(self, event):
        if event.inaxes is not self.ax or event.xdata is None or event.button != 1:
            return
        detector = self.detector.get()
        candidates = []
        for compound, var in self.individual_compound_vars.items():
            expected = RETENTION_TIMES.get(compound)
            if var.get() and expected and expected[0] == detector:
                candidates.append((abs(expected[1] - event.xdata), compound))
        if candidates and min(candidates)[0] <= 0.35:
            self._drag_rt_compound = min(candidates)[1]
            self._drag_rt_detector = detector

    def _rt_motion(self, event):
        if self._drag_rt_compound is None or event.inaxes is not self.ax or event.xdata is None:
            return
        compound = self._drag_rt_compound
        detector = self._drag_rt_detector
        RETENTION_TIMES[compound] = (detector, float(event.xdata))
        self._refresh_compound_labels()
        self._draw_individual_retention_lines()
        self._draw_overlay_retention_lines()
        self.update_rt_label()
        self.canvas.draw_idle()
        self.overlay_canvas.draw_idle()
        self.status.config(text=f"{compound} retention time updated to {event.xdata:.3f} min.")

    def _rt_release(self, event):
        self._drag_rt_compound = None
        self._drag_rt_detector = None

    def update_rt_label(self):
        if not hasattr(self, "rt_label"):
            return
        expected = RETENTION_TIMES.get(self.analyte.get())
        self.rt_label.config(text=f"Expected RT ~{expected[1]:.1f} min" if expected and expected[0] == self.detector.get() else "Expected RT —")

    def edit_concentration_cell(self, event):
        item = self.conc_tree.identify_row(event.y)
        column = self.conc_tree.identify_column(event.x)
        if not item or column == "#1":
            return
        idx = int(column[1:]) - 1
        bbox = self.conc_tree.bbox(item, column)
        if not bbox:
            return
        old = self.concentrations[item][idx]
        entry = ttk.Entry(self.conc_tree)
        entry.insert(0, "" if old is None else f"{old:.8g}")
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.focus_set()
        def finish(_=None):
            raw = entry.get().strip()
            try:
                value = None if raw == "" else float(raw)
                self.concentrations[item][idx] = value
            except ValueError:
                messagebox.showwarning("Invalid concentration", "Enter a number or leave the cell blank.")
            entry.destroy()
            self.refresh_concentration_table()
            self.refresh_curve_page()
        entry.bind("<Return>", finish)
        entry.bind("<FocusOut>", finish)

    def refresh_concentration_table(self):
        if not hasattr(self, "conc_tree"):
            return
        detector = self.detector.get()
        for a in ANALYTES:
            vals = [a]
            for i, c in enumerate(CALIBRATIONS):
                conc = self.concentrations[a][i]
                area = self.areas[(detector, a, c)]
                if conc is None:
                    vals.append("")
                elif area:
                    vals.append(f"{conc:.4f} | {area}")
                else:
                    vals.append(f"{conc:.4f}")
            self.conc_tree.item(a, values=vals)
            self.area_tree.item(a, values=[a] + [self.areas[(detector, a, c)] for c in CALIBRATIONS])

    def on_individual_select(self, xmin, xmax):
        if self.x is None:
            return
        try:
            area = integrate_peak(self.x, self.y, xmin, xmax)
            a, c, d = self.analyte.get(), self.cal.get(), self.detector.get()
            if self.concentrations[a][CALIBRATIONS.index(c)] is None:
                messagebox.showwarning("No concentration", f"{a} has no concentration for {c}.")
                return
            self.areas[(d, a, c)] = f"{area:.8g}"
            self.refresh_concentration_table()
            self.refresh_curve_page()
            self.status.config(text=f"{d} / {a} / {c}: area = {area:.8g}")
        except Exception as exc:
            self.status.config(text=str(exc))

    def set_overlay_mode(self, mode):
        self.overlay_mode = mode
        self.overlay_status.config(text=f"Select a {mode} range by dragging on the overlay.")
        self.overlay_selector.set_active(True)

    def on_overlay_select(self, xmin, xmax):
        if xmax < xmin:
            xmin, xmax = xmax, xmin
        if self.overlay_mode == "peak":
            self.peak_range = (float(xmin), float(xmax))
            color = "tab:orange"
            if self.overlay_peak_artist:
                self.overlay_peak_artist.remove()
            self.overlay_peak_artist = self.overlay_ax.axvspan(xmin, xmax, color=color, alpha=.25)

            # Automatically assign the closest visible retention-time line
            # contained in the selected range.
            candidates = []
            detector = self.detector.get()
            for compound, var in self.individual_compound_vars.items():
                expected = RETENTION_TIMES.get(compound)
                if var.get() and expected and expected[0] == detector:
                    rt = expected[1]
                    if min(xmin, xmax) <= rt <= max(xmin, xmax):
                        candidates.append((abs(rt - (xmin + xmax) / 2), compound))
            if candidates:
                self.overlay_analyte.set(compound_label(min(candidates)[1]))
                self.overlay_status.config(
                    text=f"Automatically assigned to {compound_label(min(candidates)[1])}."
                )
        else:
            self.baseline_range = (float(xmin), float(xmax))
            if self.overlay_baseline_artist:
                self.overlay_baseline_artist.remove()
            self.overlay_baseline_artist = self.overlay_ax.axvspan(xmin, xmax, color="tab:green", alpha=.22)
        self.overlay_canvas.draw_idle()

    def integrate_overlay_peak(self):
        """Integrate the selected peak in every loaded calibration curve.

        A single peak assignment therefore populates the complete row in the
        area table.  Calibration slots where the assigned compound has no
        concentration are stored as zero, as are unavailable calibration
        spectra.
        """
        if self.peak_range is None:
            self.overlay_status.config(text="Select a peak range first.")
            return
        if not self.overlay_spectra:
            self.overlay_status.config(text="Load calibration curves first.")
            return

        try:
            left, right = self.peak_range
            a, d = compound_from_label(self.overlay_analyte.get()), self.detector.get()
            results = []

            for i, cal in enumerate(CALIBRATIONS):
                # A blank concentration means this compound is absent from
                # that calibration solution.
                if self.concentrations[a][i] is None:
                    self.areas[(d, a, cal)] = "0"
                    continue

                spectrum = self.overlay_spectra.get(cal)
                if spectrum is None:
                    self.areas[(d, a, cal)] = "0"
                    continue

                x, y = spectrum
                if self.baseline_range:
                    bl, br = self.baseline_range
                    mask = (x >= min(bl, br)) & (x <= max(bl, br))
                    if mask.sum() < 2:
                        raise ValueError(
                            f"Baseline range contains fewer than two points in {cal}."
                        )
                    # The baseline is defined by the signal at the selected
                    # baseline-range endpoints for this individual curve.
                    baseline_left = float(np.interp(bl, x, y))
                    baseline_right = float(np.interp(br, x, y))
                    area = integrate_peak(
                        x, y, left, right, baseline_left, baseline_right
                    )
                else:
                    area = integrate_peak(x, y, left, right)

                self.areas[(d, a, cal)] = f"{area:.8g}"
                results.append((cal, area))

            self.refresh_concentration_table()
            self.refresh_curve_page()
            self.overlay_status.config(
                text=f"Saved {a}: integrated {len(results)} calibration curve(s); "
                     f"absent/missing samples set to 0."
            )
        except Exception as exc:
            self.overlay_status.config(text=str(exc))

    def clear_overlay_area(self):
        compound = compound_from_label(self.overlay_analyte.get())
        self.areas[(self.detector.get(), compound, self.cal.get())] = ""
        self.refresh_concentration_table()
        self.refresh_curve_page()

    def reset_overlay_selection(self):
        self.peak_range = self.baseline_range = None
        for artist in (self.overlay_peak_artist, self.overlay_baseline_artist):
            if artist:
                artist.remove()
        self.overlay_peak_artist = self.overlay_baseline_artist = None
        self.overlay_canvas.draw_idle()

    def refresh_overlay(self):
        if not hasattr(self, "overlay_ax"):
            return
        self.overlay_ax.clear()
        self.overlay_spectra = {}
        detector = self.detector.get()
        plotted = 0
        for cal in CALIBRATIONS:
            source = self.files.get(cal)
            if not source:
                continue
            try:
                x, y = read_spectrum(source, detector)
                self.overlay_spectra[cal] = (x, y)
                self.overlay_ax.plot(x, y, lw=1, label=cal)
                plotted += 1
            except Exception:
                pass
        self.overlay_ax.set_title(f"Overlay — {detector}")
        self.overlay_ax.set_xlabel("Retention time / X")
        self.overlay_ax.set_ylabel("Signal")
        self.overlay_ax.xaxis.set_major_locator(MultipleLocator(1))
        self.overlay_ax.grid(alpha=.25)
        self._draw_overlay_retention_lines()
        if plotted:
            self.overlay_ax.legend(fontsize=8, ncol=2)
        self.overlay_canvas.draw_idle()
        self.reset_overlay_selection()
        if hasattr(self, "overlay_status"):
            self.overlay_status.config(text=f"{plotted} calibration file(s) plotted.")

    def refresh_curve_page(self):
        if not hasattr(self, "curve_ax"):
            return
        a, d = self.curve_analyte.get(), self.detector.get()
        points = []
        for i, c in enumerate(CALIBRATIONS):
            conc = self.concentrations[a][i]
            raw = self.areas[(d, a, c)]
            if conc is not None and raw != "":
                try:
                    points.append((float(conc), float(raw), c))
                except ValueError:
                    pass
        self.curve_ax.clear()
        self.curve_ax.set_xlabel("Concentration")
        self.curve_ax.set_ylabel("Integrated area")
        self.curve_ax.xaxis.set_major_locator(MultipleLocator(1))
        self.curve_ax.grid(alpha=.25)
        self.curve_tree.delete(*self.curve_tree.get_children())
        for c in CALIBRATIONS:
            i = CALIBRATIONS.index(c)
            conc = self.concentrations[a][i]
            raw = self.areas[(d, a, c)]
            self.curve_tree.insert("", "end", values=[c, "" if conc is None else f"{conc:.8g}", raw])
        if points:
            xx = np.array([p[0] for p in points])
            yy = np.array([p[1] for p in points])
            slope = float(np.sum(xx * yy) / np.sum(xx * xx)) if np.sum(xx * xx) else 0.0
            yfit = slope * xx
            sse = float(np.sum((yy - yfit) ** 2))
            sst = float(np.sum((yy - np.mean(yy)) ** 2))
            r2 = 1 - sse / sst if sst else float("nan")
            self.curve_ax.scatter(xx, yy, label="Measured points")
            line_x = np.linspace(0, max(xx) * 1.05 if max(xx) else 1, 100)
            self.curve_ax.plot(line_x, slope * line_x, color="tab:red", label="Fit through zero")
            self.curve_ax.legend()
            self.curve_info.config(text=f"Area = {slope:.8g} × concentration    R² = {r2:.6f}")
        else:
            self.curve_info.config(text="No integrated areas available.")
        self.curve_ax.set_title(f"{a} — {d}")
        self.curve_canvas.draw_idle()

    def export_excel(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            messagebox.showerror("Missing package", "Install openpyxl with: pip install openpyxl")
            return
        folder = self.last_data_folder
        if not folder:
            messagebox.showwarning("No data folder", "Load calibration data first.")
            return
        path = os.path.join(folder, "HPLC_calibration_results.xlsx")
        if os.path.exists(path):
            stem, ext = os.path.splitext(path)
            n = 2
            while os.path.exists(f"{stem}_v{n}{ext}"):
                n += 1
            path = f"{stem}_v{n}{ext}"
        wb = Workbook()
        ws = wb.active
        ws.title = "Concentrations"
        ws.append(["Compound"] + CALIBRATIONS)
        for a in ANALYTES:
            ws.append([a] + self.concentrations[a])
        for d in DETECTORS:
            ws = wb.create_sheet(f"Areas {d[:20]}")
            ws.append(["Compound"] + CALIBRATIONS)
            for a in ANALYTES:
                ws.append([a] + [self.areas[(d, a, c)] for c in CALIBRATIONS])
        ws = wb.create_sheet("Calibration curves")
        ws.append(["Detector", "Compound", "Calibration", "Concentration", "Area", "Slope through zero", "R2"])
        for d in DETECTORS:
            for a in ANALYTES:
                vals = []
                for i, c in enumerate(CALIBRATIONS):
                    conc, raw = self.concentrations[a][i], self.areas[(d, a, c)]
                    if conc is not None and raw != "":
                        try: vals.append((float(conc), float(raw), c))
                        except ValueError: pass
                slope = np.sum([x*y for x,y,_ in vals]) / np.sum([x*x for x,y,_ in vals]) if vals else ""
                r2 = ""
                if vals:
                    yy = np.array([y for _,y,_ in vals])
                    fit = np.array([slope*x for x,_,_ in vals])
                    sse, sst = np.sum((yy-fit)**2), np.sum((yy-np.mean(yy))**2)
                    r2 = 1-sse/sst if sst else ""
                for conc, area, c in vals:
                    ws.append([d, a, c, conc, area, slope, r2])
        ws = wb.create_sheet("Retention times")
        ws.append(["Compound", "Detector", "Retention time (min)"])
        for a in ANALYTES:
            detector_name, rt = RETENTION_TIMES.get(a, ("", ""))
            ws.append([a, detector_name, rt])

        ws = wb.create_sheet("Files")
        ws.append(["Calibration", "File"])
        for c in CALIBRATIONS:
            source = self.files.get(c)
            ws.append([c, self._source_name(source) if source else ""])
        wb.save(path)
        messagebox.showinfo("Export complete", f"Saved {os.path.basename(path)}")

    def clear_selected(self):
        self.areas[(self.detector.get(), self.analyte.get(), self.cal.get())] = ""
        self.refresh_concentration_table()
        self.refresh_curve_page()


if __name__ == "__main__":
    CalibrationApp().mainloop()
