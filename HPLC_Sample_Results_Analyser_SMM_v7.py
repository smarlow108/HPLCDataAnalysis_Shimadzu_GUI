"""HPLC sample-results analyser using an exported calibration workbook.

Workflow:
1. Load HPLC calibration results workbook and any number of sample data files.
2. Page 1: inspect each sample curve individually.
3. Page 2: select a peak and baseline range, assign the peak to a compound, and
   integrate that range across every loaded sample curve.
4. Page 3: view calculated concentrations and export them to Excel.

The calibration workbook is the workbook produced by the calibration program:
Concentrations, Areas PDA 210 nm, Areas ELSD, Calibration curves, and Files.
"""

from __future__ import annotations

import os
import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib.widgets import SpanSelector
from matplotlib.ticker import MultipleLocator


COMPOUNDS = [
    "Oxalic Acid", "Tartronic Acid", "Glyceric Acid", "Glycolic Acid",
    "Lactic Acid", "Formic Acid", "Glycerol",
]
DETECTOR_FOR = {name: ("ELSD" if name == "Glycerol" else "PDA 210 nm") for name in COMPOUNDS}
DISPLAY_ORDER = COMPOUNDS

# Approximate retention times used only for dropdown labels.
# Update these values if your method has different retention times.
RETENTION_TIMES = {
    "Oxalic Acid": 6.6,
    "Tartronic Acid": 8.1,
    "Glyceric Acid": 10.8,
    "Glycolic Acid": 12.3,
    "Lactic Acid": 12.7,
    "Formic Acid": 13.8,
    "Glycerol": 13.2,
}


def compound_label(compound: str) -> str:
    rt = RETENTION_TIMES.get(compound)
    return f"{compound} ({rt:g} min)" if rt is not None else compound


def compound_from_label(label: str) -> str:
    return str(label).split(" (", 1)[0]


def _canonical(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def _compound_from_calibration(value: str) -> str | None:
    key = _canonical(value)
    for name in COMPOUNDS:
        if _canonical(name) == key:
            return name
    return None


def read_labsolutions_file(filename: str):
    with open(filename, "rb") as fh:
        header = fh.read(8)
    if header == bytes.fromhex("D0CF11E0A1B11AE1"):
        raise ValueError(
            "Binary LabSolutions .lcd files are not decoded by this script. "
            "Export the PDA/ELSD chromatogram as ASCII or CSV first."
        )
    raise ValueError("Not a readable ASCII file.")


def _ascii_section(text: str, section: str):
    start = text.find(section)
    if start < 0:
        raise ValueError(f"ASCII section not found: {section}")
    end = text.find("\n[", start + len(section))
    block = text[start:] if end < 0 else text[start:end]
    rows, started = [], False
    for line in block.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("R.Time (min)"):
            started = True
            continue
        if not started:
            continue
        fields = re.split(r"[;,\t ]+", line)
        if len(fields) < 2:
            continue
        try:
            rows.append((float(fields[0]), float(fields[1])))
        except ValueError:
            pass
    if len(rows) < 2:
        raise ValueError(f"No numerical data found in {section}")
    arr = np.asarray(rows, dtype=float)
    return arr[:, 0], arr[:, 1]


def read_spectrum(filename: str, detector: str):
    with open(filename, "rb") as fh:
        raw = fh.read()
    if raw.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        return read_labsolutions_file(filename)
    text = raw.decode("utf-8-sig", errors="replace")
    if "[LC Chromatogram(ELSD)]" in text or "[PDA Multi Chromatogram" in text:
        section = "[LC Chromatogram(ELSD)]" if detector == "ELSD" else "[PDA Multi Chromatogram(Ch1)]"
        return _ascii_section(text, section)

    rows = []
    for line in text.splitlines():
        fields = re.split(r"[;,\t ]+", line.strip())
        nums = []
        for field in fields:
            try:
                nums.append(float(field.replace(",", ".")))
            except ValueError:
                pass
        if len(nums) >= 2:
            rows.append((nums[0], nums[1]))
    if len(rows) < 2:
        raise ValueError("No two-column numeric spectrum was found.")
    arr = np.asarray(rows, dtype=float)
    arr = arr[np.argsort(arr[:, 0])]
    unique = np.r_[True, np.diff(arr[:, 0]) != 0]
    return arr[unique, 0], arr[unique, 1]


def sample_name_from_file(filename: str) -> str:
    try:
        text = Path(filename).read_text(encoding="utf-8-sig", errors="replace")
        for pattern in (
            r"(?im)^\s*Sample Name\s*[:;\t ]+\s*(.+?)\s*$",
            r"(?im)^\s*SampleName\s*[:;\t ]+\s*(.+?)\s*$",
        ):
            match = re.search(pattern, text)
            if match:
                value = match.group(1).strip().strip(";,\t")
                if value:
                    return value
    except OSError:
        pass
    return Path(filename).stem


def integrate_peak(x, y, left, right, baseline_left=None, baseline_right=None):
    if right < left:
        left, right = right, left
        baseline_left, baseline_right = baseline_right, baseline_left
    mask = (x >= left) & (x <= right)
    if mask.sum() < 2:
        raise ValueError("The selected peak range contains fewer than two data points.")
    xx, yy = x[mask], y[mask]
    if baseline_left is None:
        baseline_left = float(np.interp(left, x, y))
    if baseline_right is None:
        baseline_right = float(np.interp(right, x, y))
    baseline = np.interp(xx, [left, right], [baseline_left, baseline_right])
    return float(np.trapz(yy - baseline, xx))


def clean_integration_area(value):
    """Return a usable area; negative and sub-threshold values are treated as zero."""
    try:
        area = float(value)
    except (TypeError, ValueError):
        return 0.0
    if not np.isfinite(area) or area < 20.0:
        return 0.0
    return area


def load_calibration_workbook(filename):
    try:
        import openpyxl
    except ImportError as exc:
        raise RuntimeError("Install openpyxl with: pip install openpyxl") from exc

    wb = openpyxl.load_workbook(filename, data_only=True)
    concentrations = {compound: {} for compound in COMPOUNDS}
    slopes = {compound: None for compound in COMPOUNDS}
    r2_values = {compound: None for compound in COMPOUNDS}

    if "Concentrations" not in wb.sheetnames:
        raise ValueError("The workbook does not contain a 'Concentrations' sheet.")
    ws = wb["Concentrations"]
    headers = [str(ws.cell(1, c).value or "") for c in range(2, ws.max_column + 1)]
    for row in range(2, ws.max_row + 1):
        compound = _compound_from_calibration(ws.cell(row, 1).value or "")
        if not compound:
            continue
        for col, cal in enumerate(headers, start=2):
            value = ws.cell(row, col).value
            if value not in (None, ""):
                try:
                    concentrations[compound][cal] = float(value)
                except (TypeError, ValueError):
                    pass

    # Use the calibration-curves sheet when available; calculate a fallback
    # directly from the Areas sheets when it is not.
    if "Calibration curves" in wb.sheetnames:
        ws = wb["Calibration curves"]
        for row in range(2, ws.max_row + 1):
            detector = str(ws.cell(row, 1).value or "")
            compound = _compound_from_calibration(ws.cell(row, 2).value or "")
            if not compound:
                continue
            try:
                slope = float(ws.cell(row, 6).value)
                r2 = float(ws.cell(row, 7).value)
            except (TypeError, ValueError):
                continue
            if DETECTOR_FOR[compound] == detector and slopes[compound] is None:
                slopes[compound], r2_values[compound] = slope, r2

    areas = {detector: {compound: {} for compound in COMPOUNDS}
             for detector in ("PDA 210 nm", "ELSD")}
    for detector, sheet in (("PDA 210 nm", "Areas PDA 210 nm"), ("ELSD", "Areas ELSD")):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        headers = [str(ws.cell(1, c).value or "") for c in range(2, ws.max_column + 1)]
        for row in range(2, ws.max_row + 1):
            compound = _compound_from_calibration(ws.cell(row, 1).value or "")
            if not compound:
                continue
            for col, cal in enumerate(headers, start=2):
                value = ws.cell(row, col).value
                if value not in (None, ""):
                    try:
                        areas[detector][compound][cal] = float(value)
                    except (TypeError, ValueError):
                        pass

    for compound in COMPOUNDS:
        if slopes[compound] is not None:
            continue
        points = []
        detector = DETECTOR_FOR[compound]
        for cal, conc in concentrations[compound].items():
            area = areas[detector][compound].get(cal)
            if conc is not None and area is not None:
                points.append((conc, area))
        if points:
            xx = np.array([p[0] for p in points], dtype=float)
            yy = np.array([p[1] for p in points], dtype=float)
            denom = float(np.sum(xx * xx))
            if denom:
                slope = float(np.sum(xx * yy) / denom)
                fit = slope * xx
                sse = float(np.sum((yy - fit) ** 2))
                sst = float(np.sum((yy - np.mean(yy)) ** 2))
                slopes[compound] = slope
                r2_values[compound] = 1 - sse / sst if sst else None

    return concentrations, slopes, r2_values


class ResultsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HPLC Sample Results Analyser")
        self.geometry("1500x920")
        self.minsize(1150, 720)

        self.calibration_file = None
        self.calibration_concentrations = {}
        self.slopes = {}
        self.r2_values = {}
        self.sample_files = {}             # display name -> filename
        self.sample_spectra = {}           # detector -> sample -> (x, y)
        self.sample_names = []
        self.current_x = self.current_y = None
        self.peak_range = None
        self.baseline_range = None
        self.overlay_mode = "peak"
        self.peak_artist = self.baseline_artist = None
        self.preview_artists = []
        self.overlay_rt_artists = []
        self.integration_areas = {compound: {} for compound in COMPOUNDS}
        self.individual_compound_vars = {}

        self._build_ui()
        self._refresh_everything()

    def _build_ui(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(fill="x")
        ttk.Button(bar, text="Open calibration workbook…", command=self.open_calibration).pack(side="left")
        ttk.Button(bar, text="Open sample data files…", command=self.open_samples).pack(side="left", padx=5)
        ttk.Label(bar, text="Detector:").pack(side="left", padx=(18, 4))
        self.detector = ttk.Combobox(bar, values=["PDA 210 nm", "ELSD"], state="readonly", width=14)
        self.detector.current(0)
        self.detector.pack(side="left")
        self.detector.bind("<<ComboboxSelected>>", lambda e: self._detector_changed())
        ttk.Button(bar, text="Export results to Excel…", command=self.export_excel).pack(side="right")
        self.status = ttk.Label(self, text="Load a calibration workbook and sample files.", padding=(8, 0))
        self.status.pack(fill="x")

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=5)
        self.page1, self.page2, self.page3 = (ttk.Frame(self.tabs) for _ in range(3))
        self.tabs.add(self.page1, text="1. Individual curves")
        self.tabs.add(self.page2, text="2. Overlay and integration")
        self.tabs.add(self.page3, text="3. Results")

        self._build_page1()
        self._build_page2()
        self._build_page3()

    def _build_page1(self):
        top = ttk.Frame(self.page1, padding=4)
        top.pack(fill="x")
        ttk.Label(top, text="Sample:").pack(side="left")
        self.sample_combo = ttk.Combobox(top, state="readonly", width=30)
        self.sample_combo.pack(side="left", padx=5)
        self.sample_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_individual())
        ttk.Label(top, text="Retention-time lines:").pack(side="left", padx=(15, 4))
        line_frame = ttk.Frame(top)
        line_frame.pack(side="left", fill="x", expand=True)
        for compound in COMPOUNDS:
            var = tk.BooleanVar(value=False)
            self.individual_compound_vars[compound] = var
            ttk.Checkbutton(
                line_frame, text=compound, variable=var,
                command=self._retention_lines_changed
            ).pack(side="left", padx=(0, 8))

        self.fig1 = Figure(figsize=(9, 4.6), dpi=100)
        self.ax1 = self.fig1.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master=self.page1)
        self.canvas1.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.canvas1, self.page1).update()

        frame = ttk.LabelFrame(self.page1, text="Loaded sample files", padding=4)
        frame.pack(fill="x", padx=4, pady=4)
        self.file_tree = ttk.Treeview(frame, columns=("sample", "file"), show="headings", height=7)
        self.file_tree.heading("sample", text="Sample Name")
        self.file_tree.heading("file", text="File")
        self.file_tree.column("sample", width=260)
        self.file_tree.column("file", width=700)
        self.file_tree.pack(fill="x")

    def _build_page2(self):
        top = ttk.Frame(self.page2, padding=4)
        top.pack(fill="x")
        ttk.Label(top, text="Assign selected peak to:").pack(side="left")
        self.assign_combo = ttk.Combobox(
            top, values=[compound_label(c) for c in COMPOUNDS],
            state="readonly", width=28
        )
        self.assign_combo.set(compound_label(COMPOUNDS[0]))
        self.assign_combo.pack(side="left", padx=5)
        ttk.Button(top, text="1 Select peak range", command=lambda: self.set_overlay_mode("peak")).pack(side="left", padx=3)
        ttk.Button(top, text="2 Select baseline range", command=lambda: self.set_overlay_mode("baseline")).pack(side="left", padx=3)
        ttk.Button(top, text="3 Integrate all sample curves", command=self.integrate_all).pack(side="left", padx=3)
        ttk.Button(top, text="Reset selection", command=self.reset_selection).pack(side="left", padx=3)
        ttk.Button(top, text="Refresh overlay", command=self.refresh_overlay).pack(side="left", padx=3)
        self.overlay_status = ttk.Label(top, text="Load sample files, then select a peak range.")
        self.overlay_status.pack(side="left", padx=10)

        self.fig2 = Figure(figsize=(9, 4.4), dpi=100)
        self.ax2 = self.fig2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.fig2, master=self.page2)
        self.canvas2.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.canvas2, self.page2).update()
        self.selector = SpanSelector(
            self.ax2, self.on_select, "horizontal", useblit=True,
            props=dict(alpha=.25, facecolor="tab:orange"), interactive=True,
        )

        frame = ttk.LabelFrame(self.page2, text="Integrated areas (the assigned compound is integrated for every sample)", padding=4)
        frame.pack(fill="x", padx=4, pady=4)
        self.area_tree = ttk.Treeview(frame, show="headings", height=7)
        self.area_tree.pack(fill="x")

    def _build_page3(self):
        frame = ttk.LabelFrame(self.page3, text="Calculated concentrations", padding=4)
        frame.pack(fill="both", expand=True, padx=4, pady=4)
        self.result_tree = ttk.Treeview(frame, show="headings")
        self.result_tree.pack(fill="both", expand=True, side="left")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.result_tree.yview)
        scroll.pack(side="right", fill="y")
        self.result_tree.configure(yscrollcommand=scroll.set)
        ttk.Label(self.page3, text="Concentration = integrated area / zero-intercept calibration slope. "
                  "Glycerol uses ELSD; all other compounds use PDA 210 nm.",
                  foreground="#555").pack(anchor="w", padx=8, pady=4)

    def _detector_changed(self):
        self.refresh_individual()
        self.refresh_overlay()
        self.refresh_area_table()

    def open_calibration(self):
        filename = filedialog.askopenfilename(
            title="Select calibration results workbook",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if not filename:
            return
        try:
            self.calibration_concentrations, self.slopes, self.r2_values = load_calibration_workbook(filename)
            self.calibration_file = filename
            self.status.config(text=f"Loaded calibration workbook: {Path(filename).name}")
            self.refresh_results()
        except Exception as exc:
            messagebox.showerror("Could not load calibration workbook", str(exc))

    def open_samples(self):
        filenames = filedialog.askopenfilenames(
            title="Select sample data files",
            filetypes=[("HPLC data", "*.txt *.csv *.asc *.dat *.lcd"), ("All files", "*.*")]
        )
        if not filenames:
            return
        added = 0
        for filename in filenames:
            name = sample_name_from_file(filename)
            base, n = name, 2
            while name in self.sample_files:
                name = f"{base} ({n})"
                n += 1
            self.sample_files[name] = filename
            added += 1
        self.sample_names = list(self.sample_files)
        self.sample_combo["values"] = self.sample_names
        if self.sample_names and not self.sample_combo.get():
            self.sample_combo.set(self.sample_names[0])
        for name, filename in self.sample_files.items():
            self.file_tree.insert("", "end", iid=name, values=(name, Path(filename).name))
        self.status.config(text=f"Loaded {added} sample file(s).")
        self._refresh_everything()

    def _source_for(self, sample):
        return self.sample_files.get(sample)

    def _retention_lines_changed(self):
        """Keep the retention-line selection synchronized across Pages 1 and 2."""
        self.refresh_individual()
        if hasattr(self, "ax2"):
            self._draw_overlay_retention_lines()
            self.canvas2.draw_idle()

    def refresh_individual(self):
        self.ax1.clear()
        sample = self.sample_combo.get()
        if sample and sample in self.sample_files:
            try:
                x, y = read_spectrum(self.sample_files[sample], self.detector.get())
                self.current_x, self.current_y = x, y
                self.ax1.plot(x, y, lw=1, color="tab:blue")
                self.ax1.set_title(f"{sample} — {self.detector.get()}")
                self.ax1.set_xlabel("Retention time / X")
                self.ax1.set_ylabel("Signal")
            except Exception as exc:
                self.current_x = self.current_y = None
                self.ax1.set_title(str(exc))
        else:
            self.current_x = self.current_y = None
            self.ax1.set_title("No sample selected")
        for compound, var in self.individual_compound_vars.items():
            if var.get() and RETENTION_TIMES.get(compound) is not None:
                rt = RETENTION_TIMES[compound]
                self.ax1.axvline(
                    rt, color="tab:red", linestyle="--", linewidth=1.1, alpha=.85
                )
                self.ax1.text(
                    rt, 0.98, compound, rotation=90, color="tab:red",
                    transform=self.ax1.get_xaxis_transform(), ha="right", va="top",
                    fontsize=8, clip_on=True
                )
        self.ax1.xaxis.set_major_locator(MultipleLocator(1))
        self.ax1.grid(alpha=.25)
        self.canvas1.draw_idle()

    def _draw_overlay_retention_lines(self):
        """Draw the retention-time lines selected on Page 1 and keep them on Page 2."""
        self.overlay_rt_artists = []
        for compound, var in self.individual_compound_vars.items():
            if var.get() and RETENTION_TIMES.get(compound) is not None:
                rt = RETENTION_TIMES[compound]
                self.overlay_rt_artists.append(
                    self.ax2.axvline(
                        rt, color="tab:red", linestyle="--",
                        linewidth=1.1, alpha=.85, zorder=4
                    )
                )
                self.overlay_rt_artists.append(
                    self.ax2.text(
                        rt, 0.98, compound,
                        rotation=90, color="tab:red",
                        transform=self.ax2.get_xaxis_transform(),
                        ha="right", va="top", fontsize=8,
                        clip_on=True, zorder=5
                    )
                )

    def refresh_overlay(self):
        self.ax2.clear()
        self.overlay_rt_artists = []
        self.sample_spectra[self.detector.get()] = {}
        detector = self.detector.get()
        plotted = 0
        for sample, filename in self.sample_files.items():
            try:
                x, y = read_spectrum(filename, detector)
                self.sample_spectra[detector][sample] = (x, y)
                self.ax2.plot(x, y, lw=1, label=sample)
                plotted += 1
            except Exception:
                pass
        self.ax2.set_title(f"Overlay — {detector}")
        self.ax2.set_xlabel("Retention time / X")
        self.ax2.set_ylabel("Signal")
        self.ax2.xaxis.set_major_locator(MultipleLocator(1))
        self.ax2.grid(alpha=.25)

        # Re-draw the Page 1 retention-time markers after the overlay axes
        # are cleared, so they remain visible on Page 2.
        self._draw_overlay_retention_lines()

        if plotted:
            self.ax2.legend(fontsize=8, ncol=2)
        self.canvas2.draw_idle()
        self.reset_selection()
        self.overlay_status.config(text=f"{plotted} sample curve(s) plotted.")

    def refresh_area_table(self):
        columns = ["compound"] + self.sample_names
        self.area_tree.configure(columns=columns)
        for col in columns:
            self.area_tree.heading(col, text="Compound" if col == "compound" else col)
            self.area_tree.column(col, width=150 if col == "compound" else 120, anchor="center")
        self.area_tree.delete(*self.area_tree.get_children())
        for compound in COMPOUNDS:
            values = [compound]
            for sample in self.sample_names:
                value = clean_integration_area(
                    self.integration_areas[compound].get(sample, 0.0)
                )
                values.append(value)
            self.area_tree.insert("", "end", iid=compound, values=values)

    def refresh_results(self):
        columns = ["compound"] + self.sample_names
        self.result_tree.configure(columns=columns)
        for col in columns:
            self.result_tree.heading(col, text="Compound" if col == "compound" else col)
            self.result_tree.column(col, width=150 if col == "compound" else 120, anchor="center")
        self.result_tree.delete(*self.result_tree.get_children())
        for compound in DISPLAY_ORDER:
            row = [compound]
            slope = self.slopes.get(compound)
            for sample in self.sample_names:
                area = clean_integration_area(
                    self.integration_areas[compound].get(sample, 0.0)
                )
                if slope in (None, "", 0):
                    row.append("")
                else:
                    row.append(f"{area / float(slope):.8g}")
            self.result_tree.insert("", "end", iid=compound, values=row)

    def _refresh_everything(self):
        if hasattr(self, "file_tree"):
            self.refresh_individual()
            self.refresh_overlay()
            self.refresh_area_table()
            self.refresh_results()

    def set_overlay_mode(self, mode):
        self.overlay_mode = mode
        self.overlay_status.config(text=f"Drag on the plot to select a {mode} range.")
        self.selector.set_active(True)

    def _clear_preview_artists(self):
        for artist in self.preview_artists:
            try:
                artist.remove()
            except (ValueError, AttributeError):
                pass
        self.preview_artists = []

    def _draw_integration_preview(self):
        """Show the selected baseline and red integrated area on every curve."""
        self._clear_preview_artists()
        detector = self.detector.get()
        spectra = self.sample_spectra.get(detector, {})

        if not self.baseline_range:
            return

        bl, br = self.baseline_range
        if br < bl:
            bl, br = br, bl

        # Show the baseline over the selected baseline range even before a
        # peak range has been selected.
        if self.peak_range:
            left, right = self.peak_range
            if right < left:
                left, right = right, left
        else:
            left, right = bl, br

        for x, y in spectra.values():
            mask = (x >= left) & (x <= right)
            if mask.sum() < 2:
                continue
            xx = x[mask]
            yy = y[mask]
            baseline = np.interp(xx, [bl, br], [
                float(np.interp(bl, x, y)),
                float(np.interp(br, x, y)),
            ])

            # Baseline line across the selected peak (or baseline range).
            self.preview_artists.extend(
                self.ax2.plot(xx, baseline, color="black", lw=1.2, alpha=.85)
            )

            # Shade only the positive area between signal and baseline once
            # both the peak and baseline ranges have been selected.
            if self.peak_range:
                self.preview_artists.append(
                    self.ax2.fill_between(
                        xx, baseline, yy, where=(yy >= baseline),
                        color="red", alpha=.28, linewidth=0
                    )
                )

    def _auto_assign_peak_compound(self, left, right):
        """Assign a selected peak to a visible retention-time marker when possible."""
        candidates = [
            (compound, RETENTION_TIMES[compound])
            for compound, var in self.individual_compound_vars.items()
            if var.get() and RETENTION_TIMES.get(compound) is not None
            and left <= RETENTION_TIMES[compound] <= right
        ]
        if not candidates:
            return None

        # If more than one marker is inside the range, use the marker nearest
        # the centre of the selected range.
        centre = (left + right) / 2.0
        compound = min(candidates, key=lambda item: abs(item[1] - centre))[0]
        self.assign_combo.set(compound_label(compound))
        return compound

    def on_select(self, xmin, xmax):
        if xmax < xmin:
            xmin, xmax = xmax, xmin
        if self.overlay_mode == "peak":
            self.peak_range = (float(xmin), float(xmax))
            assigned = self._auto_assign_peak_compound(float(xmin), float(xmax))
            if assigned:
                self.overlay_status.config(
                    text=f"Automatically assigned peak to {compound_label(assigned)}."
                )
            if self.peak_artist:
                self.peak_artist.remove()
            self.peak_artist = self.ax2.axvspan(
                xmin, xmax, color="tab:orange", alpha=.25, zorder=0
            )
        else:
            self.baseline_range = (float(xmin), float(xmax))
            if self.baseline_artist:
                self.baseline_artist.remove()
            self.baseline_artist = self.ax2.axvspan(
                xmin, xmax, color="tab:green", alpha=.22, zorder=0
            )

        self._draw_integration_preview()
        self.canvas2.draw_idle()

    def reset_selection(self):
        self.peak_range = self.baseline_range = None
        for artist in (self.peak_artist, self.baseline_artist):
            if artist is not None:
                try:
                    artist.remove()
                except ValueError:
                    pass
        self.peak_artist = self.baseline_artist = None
        self._clear_preview_artists()
        if hasattr(self, "canvas2"):
            self.canvas2.draw_idle()

    def integrate_all(self):
        compound = compound_from_label(self.assign_combo.get())
        detector = DETECTOR_FOR[compound]
        if not self.peak_range:
            self.overlay_status.config(text="Select a peak range first.")
            return
        if not self.sample_names:
            self.overlay_status.config(text="Load sample files first.")
            return
        # The assigned compound determines the detector, so an assignment to
        # glycerol automatically uses ELSD and all other compounds use PDA.
        if detector != self.detector.get():
            selected_peak = self.peak_range
            selected_baseline = self.baseline_range
            self.detector.set(detector)
            self.refresh_overlay()
            self.peak_range = selected_peak
            self.baseline_range = selected_baseline
            self._draw_integration_preview()

        left, right = self.peak_range
        blbr = self.baseline_range
        spectra = self.sample_spectra.get(detector, {})
        count = 0
        for sample in self.sample_names:
            if sample not in spectra:
                self.integration_areas[compound][sample] = 0.0
                continue
            x, y = spectra[sample]
            try:
                if blbr:
                    bl, br = blbr
                    mask = (x >= bl) & (x <= br)
                    if mask.sum() < 2:
                        raise ValueError("Baseline range has fewer than two points.")
                    area = integrate_peak(
                        x, y, left, right,
                        float(np.interp(bl, x, y)),
                        float(np.interp(br, x, y)),
                    )
                else:
                    area = integrate_peak(x, y, left, right)
                area = clean_integration_area(area)
                self.integration_areas[compound][sample] = area
                count += 1
            except Exception:
                self.integration_areas[compound][sample] = 0.0
        self.refresh_area_table()
        self.refresh_results()
        self._draw_integration_preview()
        self.canvas2.draw_idle()
        self.overlay_status.config(text=f"Integrated {compound} across {count} sample curve(s).")

    def export_excel(self):
        if not self.calibration_file:
            messagebox.showwarning("No calibration workbook", "Load the calibration workbook first.")
            return
        if not self.sample_names:
            messagebox.showwarning("No samples", "Load at least one sample data file first.")
            return
        try:
            from openpyxl import Workbook
        except ImportError:
            messagebox.showerror("Missing package", "Install openpyxl with: pip install openpyxl")
            return
        folder = str(Path(self.calibration_file).parent)
        output = os.path.join(folder, "HPLC_sample_results.xlsx")
        if os.path.exists(output):
            stem, ext = os.path.splitext(output)
            n = 2
            while os.path.exists(f"{stem}_v{n}{ext}"):
                n += 1
            output = f"{stem}_v{n}{ext}"

        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(["Compound"] + self.sample_names)
        for compound in DISPLAY_ORDER:
            row = [compound]
            slope = self.slopes.get(compound)
            for sample in self.sample_names:
                area = clean_integration_area(
                    self.integration_areas[compound].get(sample, 0.0)
                )
                row.append("" if not slope else area / float(slope))
            ws.append(row)

        ws = wb.create_sheet("Integrated areas")
        ws.append(["Compound"] + self.sample_names)
        for compound in DISPLAY_ORDER:
            ws.append([compound] + [
                clean_integration_area(
                    self.integration_areas[compound].get(sample, 0.0)
                ) for sample in self.sample_names
            ])

        ws = wb.create_sheet("Calibration coefficients")
        ws.append(["Compound", "Detector", "Slope through zero", "R2"])
        for compound in DISPLAY_ORDER:
            ws.append([compound, DETECTOR_FOR[compound], self.slopes.get(compound, ""), self.r2_values.get(compound, "")])

        ws = wb.create_sheet("Sample files")
        ws.append(["Sample Name", "File"])
        for sample, filename in self.sample_files.items():
            ws.append([sample, filename])

        wb.save(output)
        messagebox.showinfo("Export complete", f"Saved {Path(output).name}")

    def destroy(self):
        super().destroy()


if __name__ == "__main__":
    ResultsApp().mainloop()
