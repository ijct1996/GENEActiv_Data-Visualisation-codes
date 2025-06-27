 Actograph Analysis  Summary GUI for Parkinson's Participants (v7)
 ------------------------------------------------------------------------
 Authored by Isaiah J. Ting, Chronobiology Postdoctoral Scientist
 Date: 27 June 2025

 ## Introduction
 This MATLAB app provides a single, self-contained graphical interface to
 process wrist-worn actigraphy data for Parkinson's research (or any
 actigraphy dataset). Load your data, set thresholds and options, then
 click **Run** to generate visualisations, metrics and reports.

 ## Prerequisites
 - MATLAB R2019a or later  
 - Report Generator toolbox (for PowerPoint export)  
 - The file **GENEActiv Data Template.xlsx** (included with this code)  

 ## Preparing Input Data
 1. Open **GENEActiv Data Template.xlsx**.  
 2. Copy your raw CSV data into the sheet named **RawData** (columns must match exactly).  
 3. Save and close the template—it becomes your input file.  

 ## Input Format
 Your Excel file (template) must contain these columns in the **RawData** sheet:  
 - **Time stamp** (`yyyy-MM-dd HH:mm:ss:SSS`)  
 - **Sum of vector (SVMg)** (minute-by-minute activity counts)  
 - **Light level (LUX)** (ambient light intensity)  
 - **Temperature** (°C)  

 ## GUI Features
 - **Input selectors** for raw Excel file and output folder  
 - **Numeric field** to set a global light threshold (lux)  
 - **Checkbox** to include only complete 24-hour days or all days  
 - **Dropdown** to choose axis style: anonymised day number, calendar date (`dd/mm`), or both  
 - **Output panel** to select which figures and exports to generate:  
   • Activity profiles (bar plots with light shading)  
   • Daily activity bar charts  
   • Low-activity call-outs  
   • Activity heatmaps  
   • Temperature profiles (15–50 °C with shading)  
   • Daily light tracker (full-range and 0–10 LUX)  
   • Weekly light distribution area plots  
   • Combined profiles (activity, light, temperature)  
   • Excel workbook (Summary, Metrics and Definitions sheets)  
   • PowerPoint deck of all figures  

 ## Analysis Workflow
 When you click **Run**, the app will:  
 1. Validate inputs  
 2. Auto-detect sampling interval and bin to 1-minute resolution per day  
 3. Optionally filter to complete days  
 4. Compute daily metrics:  
    - Total activity, hours in light, min/max temperature  
    - **L5** (lowest-5-hour) and **M10** (highest-10-hour) metrics  
 5. Compute rhythm stability:  
    - **Interdaily Stability (IS)**  
    - **Intradaily Variability (IV)**  
 6. Generate weekly figures as selected  
 7. Export:  
    - High-resolution JPEGs (white background)  
    - **Excel workbook** (`Participant_Results.xlsx`) with **Summary**, **Metrics**, **Definitions**  
    - **PowerPoint** (`AllFigures_Report.pptx`) compiling every JPEG  

 ## Output Files
 - `01_…` to `07_…` JPEGs (dated and/or anonymised)  
 - `Participant_Results.xlsx`:  
    • **Summary**: daily metrics, L5/M10, min/max temp, hours in light  
    • **Metrics**: IS, IV and normal ranges  
    • **Definitions**: glossary and interpretation guidance  
 - `AllFigures_Report.pptx`: one slide per figure  

 ## Usage
 Follow these steps to run the GUI and process your data:

 1. Ensure you have **GENEActiv Data Template.xlsx** prepared with  
    your raw CSV data in the **RawData** sheet.  
 2. In MATLAB's Command Window or Editor pane, type and enter:  
    ```matlab
    >> Actograph_v7_forparticipants_GUI
    ```  
 3. In the GUI that appears:  
    a. Click **Browse…** next to **Input file:** and select your prepared **GENEActiv Data Template.xlsx**.  
    b. Click **Browse…** next to **Output folder:** and choose (or create) a folder for results.  
    c. Enter your desired **Light threshold (lux)**.  
    d. Tick **Include only complete 24-hour days** if required.  
    e. Select an **Axis style**: "Days only", "Dated only" or "Both".  
    f. In **Select Outputs**, check the figures and exports you need.  
 4. Click the **Run** button.  
 5. Monitor the **Status** label—once it reads **Done.**, your outputs are ready in the chosen folder.  

 *Contact:* Isaiah J. Ting, Chronobiology Postdoctoral Scientist - i.ting@kent.ac.uk

 ------------------------------------------------------------------------  
