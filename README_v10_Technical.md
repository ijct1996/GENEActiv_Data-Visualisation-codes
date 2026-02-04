
---

## Technical README (v10.5)

```markdown
Actograph Analysis and Summary GUI for GENEActiv Participants (v10.5) [Technical README]
-----------------------------------------------------------------------------------------------
Authored by Isaiah J. Ting, Chronobiology Postdoctoral Scientist
Date: 4 February 2026

## Introduction
This MATLAB GUI is a single-function pipeline for ingesting GENEActiv-style actigraphy exports from Excel, regularising onto a fixed epoch grid, and producing a standardised set of 0 to 48 h visual outputs and daily summary metrics (with optional PowerPoint compilation).

It is designed to be robust to common real-world issues: timestamp formatting differences, irregular sampling, missing epochs, partial start or end days, and optional temperature channels.

## Prerequisites
- MATLAB R2019a or later (requires App Designer UI components, i.e. `uifigure`)  
- Report Generator toolbox (only if PowerPoint export is selected)  
- Excel input workbook with a **RawData** sheet (preferred) or any first sheet matching the required columns

## Preparing Input Data
1. Use **GENEActiv Data Template.xlsx** where possible.  
2. Paste raw data into **RawData** with the expected columns.  
3. Save the workbook and use it as the GUI input.

Sheet selection rule:
- If a sheet named **RawData** exists, it is used. Otherwise, the first sheet is used.

## Input Format
### Required channels
Columns are detected using loose matching against candidate names:

- Time: `Time stamp` (also accepts `Timestamp`, `Time Stamp`, `Time`)  
- Activity: `Sum of vector (SVMg)` (also accepts `SVMg`, `SVM`, `Activity`, `Activity (SVMg)`)  
- Light: `Light level (LUX)` (also accepts `Light level (Lux)`, `Light (LUX)`, `Lux`, `Light`)  

### Optional channel
- Temperature: `Temperature` (also accepts `Temp`, `Temperature (C)`, `Temperature (°C)`)  
If temperature is absent, temperature outputs are skipped and combined plots fall back to activity + light only.

### Timestamp parsing
`Time stamp` can be:
- MATLAB `datetime`  
- Excel serial timestamps (numeric, converted via `datetime(...,'ConvertFrom','excel')`)  
- Text timestamps, parsed against common formats (e.g. `yyyy-MM-dd HH:mm:ss.SSS`, `dd/MM/yyyy HH:mm`, `MM/dd/yyyy HH:mm:ss`, etc.)

Timezone behaviour:
- The **Timezone (labels only)** field sets `timestampsLabel.TimeZone` for display labels only.  
- Binning, regularisation, and metric computation remain naive (timezone-agnostic), so DST transitions are not modelled.

## GUI Features
- Input file and output folder selection  
- `Light threshold (lux)` used for:
  - light shading in 0 to 48 h plots  
  - daily hours-in-light calculation (`Light > threshold`, within valid lux epochs)  
- Axis style selector for y-labelled 0 to 48 h outputs:
  - Days only: `Day 1..N`  
  - Dated only: `dd/MM`  
  - Both: exports both label variants
- Output selection toggles:
  - DailyLightTracker (per day, 3-panel plot)  
  - LightDistribution (mean and mean ± SD, block-level)  
  - Activity profile 0 to 48 h (bar + light shading)  
  - Temperature profile 0 to 48 h (bar + light shading, right-axis ticks)  
  - Combined profile 0 to 48 h (activity bars + optional temperature line, forced PDF export)  
  - Activity heatmap 0 to 48 h (5th to 95th percentile scaling, NaN-transparent)  
  - Daily activity totals  
  - Low-activity call-outs (complete-day thresholding + complete-day-only plot)  
  - Excel outputs  
  - PowerPoint compilation
- Optional close-GUI-on-finish

## Analysis Workflow (Implementation Notes)
When you click **Run**, the pipeline proceeds as follows:

1. **Load workbook and detect columns**  
   - Uses `readtable(...,'VariableNamingRule','preserve')`  
   - Column detection is by normalised string matching against candidate names.

2. **Timestamp parsing and sorting**  
   - Invalid timestamps are dropped (`~isnat`) and remaining data are sorted by time.

3. **Epoch inference and snapping**  
   - Epoch is inferred from the median of positive `diff(timestamps)` (seconds).  
   - Epoch is snapped to a day-divisor candidate set (60, 90, 120, 180, 300, 360, 600, 720, 900, 1200, 1800, 3600).  
   - If snapping produces a non day-divisor, the app falls back to 60 s.

4. **Regularisation to a fixed epoch grid**  
   - Data are converted to a timetable and retimed to regular spacing using `retime(...,'regular','mean','TimeStep',seconds(epochSec))`.  
   - A full day-aligned grid is created from `day0 = dateshift(firstTime,'start','day')` through the last aligned day end.  
   - Retime to the full grid; missing epochs remain NaN (no interpolation).

5. **Daily binning**  
   - The full grid is reshaped into `binsPerDay x totalDays` and transposed to `totalDays x binsPerDay`.  
   - The last partial day (if any) is dropped by flooring `nTotal/binsPerDay`.

6. **Daily metrics**
   - `TotalActivity`: `sum(binnedActivity,2,'omitnan')`  
   - `HoursInLight`: `sum((Light > threshold) & validLux,2) * (epochSec/3600)`  
   - `MinTemperature` and `MaxTemperature`: per-day min/max omitting NaN  
   - **L5** and **M10**:
     - Sliding mean windows are computed with NaN-aware convolution (`slidingMeanNan`)  
     - Minimum valid fraction per window is 0.90  
     - Output is start time and window mean

7. **IS and IV (hourly)**
   - Daily data are collapsed to 24 hourly means per day based on bin hour indices.  
   - IS and IV are computed using the hourly formulation (Vectorised approach across all days).

8. **0 to 48 h matrices**
   - Double-plot matrices are constructed as `Day i` (0 to 24 h) followed by `Day i+1` (24 to 48 h).  
   - Final row second half is blank.

9. **Export rules**
   - Default export is JPG at 600 dpi.  
   - If `totalDays > 30`, large 0 to 48 h outputs may export as vector PDF to avoid oversized raster files.  
   - Combined profile is always exported as vector PDF and is not added to the PowerPoint queue.  
   - Temperature profile uses global min and max across the dataset, outward rounded to 0.1 °C for stable tick endpoints, and adds a right-side blanket axis label.

10. **Low-activity call-outs (v10.5 behaviour)**
   - Complete day definition is coverage-based on activity epochs: `>= 0.95*binsPerDay` non-NaN epochs.  
   - The low-activity threshold is `mu - sd` computed from complete days (or all days if fewer than three complete days exist).  
   - In the full plot, incomplete days are coloured grey and never labelled “Low”.  
   - A second plot is exported with complete days only.

11. **PowerPoint compilation**
   - If Report Generator is available, the app compiles queued images into `AllFigures_Report.pptx`.  
   - Slide order is deterministic based on the export queue. If the queue is empty, it falls back to a recursive JPG search.

## Output Files
The output folder may contain:

- `Participant_Results.xlsx`  
  - **Summary**: daily metrics and L5/M10  
  - **Metrics**: epoch details, thresholds, axis style, timezone label setting, export mode, IS/IV, sheet used, days analysed, temperature availability  
  - **Definitions**: terms and interpretation notes

- Figures (JPG 600 dpi and or vector PDF depending on export rules)  
  - `DailyLightTracker/DailyLightTracker_yyyy-mm-dd.jpg`  
  - `LightDistribution_Mean_0-24h_yyyy-mm-dd_to_yyyy-mm-dd.jpg`  
  - `LightDistribution_MeanSD_0-24h_… .jpg`  
  - `Actogram_Activity_0-48h_<Days|Dates>_… .jpg or .pdf`  
  - `Heatmap_Activity_0-48h_<Days|Dates>_… .jpg or .pdf`  
  - `Actogram_Temperature_0-48h_<Days|Dates>_… .jpg or .pdf`  
  - `Actogram_Combined_0-48h_<Days|Dates>_… .pdf`  
  - `DailyActivity_… .jpg`  
  - `LowActivity_… .jpg`  
  - `LowActivity_CompleteDaysOnly_… .jpg`

- `AllFigures_Report.pptx` (optional)

## Usage
1. Prepare the input workbook and confirm the required columns are present.  
2. In MATLAB, run:  
   ```matlab
   >> Actograph_v10_forparticipants_GUI
