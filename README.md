# Actograph Analysis and Summary for Parkinson’s Participants

## Introduction
This MATLAB script processes raw actigraphy data to produce comprehensive visualisations and a quantitative summary of key rest–activity and light‐exposure metrics. It is tailored for Parkinson’s research but can be adapted for any actigraphy dataset.

## Features
- **Data‐range selection**: choose to include only complete 24-hour days or all days of recording  
- **Binned visualisations** at 1-minute resolution:
  - Weekly rest–activity profiles (bar plots, with “Day #” and “dd/mm” labels)  
  - Weekly heatmaps of minute-by-minute activity  
  - Weekly bar charts of daily total activity and hours in light  
  - Low-activity call-out plots highlighting unusually low days  
  - A full-period area plot of average light by hour  
- **Rhythm metrics**:
  - **L5** (lowest‐5-hour activity window) start time & mean  
  - **M10** (highest-10-hour activity window) start time & mean  
  - **Interdaily Stability (IS)**  
  - **Intradaily Variability (IV)**  
- **Output bundling**:
  - High-resolution JPEGs of every figure  
  - An Excel workbook with two sheets:
    1. **Summary** of daily metrics, L5/M10, IS/IV  
    2. **Definitions** of every term with interpretation guidance  
  - A PowerPoint deck collecting all JPEGs as individual slides  

## Requirements
- MATLAB R2018a or later  
- Report Generator toolbox (for PowerPoint export)  
- Source data in `.xlsx` format  

## Input Format
The input Excel file must contain the following columns (exactly as named):
- **Time stamp** (format: `yyyy-MM-dd HH:mm:ss:SSS`)  
- **Sum of vector (SVMg)** (minute-by-minute activity counts)  
- **Light level (LUX)** (ambient light intensity)  
- **Button (1/0)** (event markers, if any)  

## Usage
1. Open and run the script in the MATLAB editor or command window.  
2. When prompted, select your actigraphy `.xlsx` file.  
3. Choose whether to **include only complete 24-hour days** or **all days**.  
4. Select (or create) an output folder where JPEGs, the Excel workbook and PowerPoint will be saved.  
5. The script will:
   - Bin the data into daily matrices  
   - Produce weekly figures (both “Day #” and dated versions)  
   - Compute L5/M10, IS, IV and other daily summary metrics  
   - Export everything to high-resolution JPEG, an Excel workbook, and a PowerPoint  

## Output
- **JPEG images** (`01_…` to `06_…`), with both `…_dated` and non-dated versions  
- **Excel workbook** `07_Participant_results.xlsx` containing:
  - **Summary** sheet: metrics table  
  - **Definitions** sheet: term glossary with interpretation  
- **PowerPoint** `AllFigures_Report.pptx` collecting every JPEG as its own slide  

## Metric Definitions (see “Definitions” sheet)
- **L5_StartTime**: clock time when the lowest 5-hour activity window begins  
- **L5_Mean**: average activity level during that trough  
- **M10_StartTime**: clock time when the highest 10-hour activity window begins  
- **M10_Mean**: average activity level during that peak  
- **Interdaily Stability (IS)**: consistency of daily rhythm; higher = more stable  
- **Intradaily Variability (IV)**: fragmentation of activity; higher = more fragmented  

## Notes
- Choosing only complete days removes any partial records and may improve the reliability of rhythm metrics.  
- All visuals and summaries automatically reflect your choice of data range.  

---

*Authored by Isaiah J. Ting, Chronobiology Postdoctoral Scientist*  
*Date: June 2025*  
