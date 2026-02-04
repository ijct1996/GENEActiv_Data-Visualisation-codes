Actograph Analysis and Summary GUI for GENEActiv Participants (v10.5)
------------------------------------------------------------------------
Authored by Isaiah J. Ting, Chronobiology Postdoctoral Scientist
Date: 4 February 2026

## Introduction
This MATLAB app provides a single, self-contained graphical interface to process wrist-worn GENEActiv actigraphy data (and similarly formatted actigraphy datasets). Load your Excel file, set thresholds and options, then click **Run** to generate figures, daily metrics, and report-ready exports.

## Prerequisites
- MATLAB R2019a or later (App Designer UI components required: `uifigure`)  
- Report Generator toolbox (only if PowerPoint export is selected)  
- An Excel workbook containing a **RawData** sheet (the provided template is recommended)

## Preparing Input Data
1. Open **GENEActiv Data Template.xlsx** (recommended).  
2. Copy your raw data into the sheet named **RawData**.  
3. Save and close the template. This becomes your input file.

If your workbook does not contain a sheet called **RawData**, the app will use the first sheet in the workbook.

## Input Format
Your Excel file should contain these columns (names are matched loosely, but keep them close to the template):

- **Time stamp** (supported as text, Excel serial time, or MATLAB datetime)  
- **Sum of vector (SVMg)** (minute-by-minute activity counts)  
- **Light level (LUX)** (ambient light intensity)  
- **Temperature** (°C) optional

## GUI Features
- **Input selectors** for the Excel file and an output folder  
- **Timezone (labels only)** field (IANA format, e.g. `Europe/London`)  
- **Numeric field** to set the light threshold (lux) used for light shading and hours-in-light  
- **Dropdown** to choose axis style for y-labelled outputs: anonymised day number, calendar date (`dd/MM`), or both  
- **Output panel** to select which figures and exports to generate:  
  • Activity profile (0 to 48 h)  
  • Activity heatmap (0 to 48 h)  
  • Temperature profile (0 to 48 h, if temperature available)  
  • Combined profile (0 to 48 h, exported as PDF)  
  • Daily activity totals  
  • Low-activity call-outs (with complete-day handling)  
  • Daily light tracker (per day, 0 to 24 h)  
  • Light distribution (block-level, 0 to 24 h)  
  • Excel workbook (Summary, Metrics, Definitions)  
  • PowerPoint deck compiling exported images  
- **Close GUI when finished** option

## Analysis Workflow
When you click **Run**, the app will:

1. Validate inputs and detect required columns  
2. Parse and sort timestamps  
3. Infer the sampling epoch and regularise the recording onto a fixed grid (missing epochs remain blank)  
4. Bin the recording into complete 24-hour day rows for plotting  
5. Compute daily metrics:  
   - Total activity, hours in light, min and max temperature  
   - **L5** (lowest-5-hour) and **M10** (highest-10-hour) metrics  
6. Compute rhythm stability and fragmentation:  
   - **Interdaily Stability (IS)**  
   - **Intradaily Variability (IV)**  
7. Generate selected figures and exports  
8. Export outputs (high-resolution images, Excel, optional PowerPoint)

Notes on missing and partial days:
- Missing epochs are not interpolated. They remain NaN and appear blank in 0 to 48 h outputs.  
- Low-activity call-outs compute the threshold using complete days and do not label incomplete days as “Low”. A second plot is exported showing only complete days.

## Output Files
Your output folder will contain, depending on your selections:

- High-resolution figure exports (JPG 600 dpi, with some large 0 to 48 h outputs exported as vector PDF)  
- `Participant_Results.xlsx`:  
  • **Summary**: daily metrics, L5/M10, min/max temperature, hours in light  
  • **Metrics**: sampling, thresholds, IS, IV, export mode, sheet used  
  • **Definitions**: glossary and interpretation notes  
- `AllFigures_Report.pptx`: one slide per exported image (PDF-only figures are not added)

Typical figure locations and names:
- `DailyLightTracker/` folder: `DailyLightTracker_yyyy-mm-dd.jpg`  
- `Actogram_Activity_0-48h_…` (Days and or Dates variants if you selected “Both”)  
- `Heatmap_Activity_0-48h_…`  
- `Actogram_Temperature_0-48h_…` (if temperature available)  
- `Actogram_Combined_0-48h_…` (PDF)

## Usage
1. Prepare your Excel file (ideally using **GENEActiv Data Template.xlsx**) with raw data in **RawData**.  
2. In MATLAB, run:  
   ```matlab
   >> Actograph_v10_forparticipants_GUI
