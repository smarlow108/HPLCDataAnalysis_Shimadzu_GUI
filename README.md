# HPLC Calibration and Sample Analysis Tools

Python/Tkinter applications for processing Shimadzu HPLC chromatograms, creating calibration curves, integrating peaks, and calculating compound concentrations in unknown samples.

## Scripts

### `HPLC_Calibration_Shimadzu_SMM_3.3.py`

Creates calibration curves from up to ten calibration solutions. It supports the following compounds:

- Oxalic acid
- Tartronic acid
- Glyceric acid
- Glycolic acid
- Lactic acid
- Formic acid
- Glycerol

The current detector assignment is:

- **PDA 210 nm:** oxalic, tartronic, glyceric, glycolic, lactic, and formic acid
- **ELSD:** glycerol

### `HPLC_Sample_Results_Analyser_SMM_v7.py`

Uses the calibration workbook produced by the calibration script to analyse any number of sample datasets and calculate concentrations for the seven compounds.

## Features

### Calibration application

The calibration application provides three pages:

1. **Individual curves**
   - View each calibration curve individually.
   - Select calibration files using their `Sample Name` or calibration identifier.
   - Display retention-time markers for the compounds.
   - Move retention-time markers to align with peaks.
   - Edit calibration concentrations in the on-screen table.
   - Use one-minute major x-axis spacing.

2. **Overlay and integration**
   - Display calibration curves together for the selected detector.
   - Select a peak range and baseline range by dragging on the plot.
   - Display the selected baseline and shaded integration region.
   - Assign peaks to compounds using labels that include retention time, for example `Glycerol (13.2 min)`.
   - Automatically assign a peak when a selected retention-time marker falls within the peak range.
   - Integrate the selected peak across all calibration curves.
   - Record zero for calibration solutions where the compound is absent.
   - Refresh the overlay after adding or changing datasets.

3. **Calibration curves**
   - Display concentration versus integrated area for each compound.
   - Fit a linear calibration through zero.
   - Display the fitted equation and R² value.
   - Show the calibration points and calculated fit.

The calibration results can be exported to an Excel workbook containing calibration areas, concentrations, fitted coefficients, R² values, retention times, and file mappings.

### Sample-results application

The sample analyser also provides three pages:

1. **Individual curves**
   - Load and inspect sample curves individually.
   - Display sample names read from the data files.
   - Show selected retention-time markers.
   - Use one-minute major x-axis spacing.

2. **Overlay and integration**
   - Overlay all loaded sample curves for the selected detector.
   - Select peak and baseline ranges.
   - Display the baseline and shaded integrated region.
   - Assign peaks manually or automatically using retention-time markers.
   - Integrate an assigned peak across every loaded sample.
   - Refresh the overlay after adding datasets.

3. **Results**
   - Calculate concentrations from the calibration slopes.
   - Combine PDA and ELSD results into one table.
   - Use this compound order:

     1. Oxalic Acid
     2. Tartronic Acid
     3. Glyceric Acid
     4. Glycolic Acid
     5. Lactic Acid
     6. Formic Acid
     7. Glycerol

   - Export results, integrated areas, calibration coefficients, and sample mappings to Excel.

Integration values below 20, including negative values, are reported as zero. Compounds with no assigned or selected peak are also reported as zero.

## Requirements

- Python 3.10 or newer recommended
- NumPy
- Matplotlib
- OpenPyXL
- Tkinter, normally included with standard Python installations

Install the Python packages with:

```bash
pip install numpy matplotlib openpyxl
```

For Linux, Tkinter may need to be installed separately, for example:

```bash
sudo apt-get install python3-tk
```

The scripts use `numpy.trapz` for compatibility with older NumPy versions.

## Running the applications

From a terminal in the repository folder:

```bash
python HPLC_Calibration_Shimadzu_SMM_3.3.py
```

After creating and saving the calibration workbook, run:

```bash
python HPLC_Sample_Results_Analyser_SMM_v7.py
```

In the sample analyser, first select the calibration workbook, then load one or more sample data files.

## Input data

The readers accept:

- Shimadzu LabSolutions ASCII exports
- CSV or tab-delimited two-column files containing time and intensity
- Text files containing LabSolutions ASCII sections such as:
  - `[LC Chromatogram(ELSD)]`
  - `[PDA Multi Chromatogram(Ch1)]`

The calibration program can load individual files or a LabSolutions ASCII batch export containing records such as `Water` and `Cal 1` to `Cal 10`. Calibration assignment uses the `Sample Name` where available, with filename information used as a fallback.

Sample names are read from the `Sample Name` field in the exported file when available. If no sample name is found, the filename is used.

## Important limitation: binary `.lcd` files

Binary LabSolutions `.lcd` compound files are detected but are not decoded by these scripts. Export the PDA and ELSD chromatograms from LabSolutions as ASCII or CSV before loading them.

## Typical workflow

1. Export calibration chromatograms from LabSolutions as ASCII or CSV.
2. Run `HPLC_Calibration_Shimadzu_SMM_3.3.py`.
3. Load the calibration files or batch export.
4. Check or edit the concentrations.
5. Adjust retention-time markers on Page 1.
6. On Page 2, select peak and baseline ranges and integrate assigned compounds across all calibration curves.
7. Review the zero-intercept calibration plots on Page 3.
8. Export the calibration workbook.
9. Run `HPLC_Sample_Results_Analyser_SMM_v7.py`.
10. Load the exported calibration workbook and sample chromatograms.
11. Integrate the relevant peaks on Page 2.
12. Review and export the concentration results on Page 3.

## Excel output

The calibration workbook includes sheets for calibration data, integrated areas, calibration curves, and file mappings. The sample-results workbook includes sheets for integrated areas, calculated results, calibration coefficients, and sample-file mappings.

Always review peak assignments, baselines, calibration fits, and exported values before using results for reporting or release decisions.

## Customisation

Retention times and compound names are defined near the top of both scripts. Update the `RETENTION_TIMES` dictionaries if the chromatographic method changes. Detector assignments can be changed in the calibration script and in the `DETECTOR_FOR` mapping in the sample analyser.

## License and use

Add the appropriate licence for your project before publishing this repository. Validate the workflow with known standards and laboratory quality procedures before using it for regulated or reportable measurements.
