# Neighborhood Property Value Tracker

This **Excel + Python** project visualizes property values in the fictional neighborhood **Syrup Hollow** and tracks changes over time.  
Clicking parcels on the map highlights the corresponding row in the data table, and selecting a row highlights its parcel on the map.

---

## Features
- Dynamic row highlighting in Excel based on map clicks or table selection  
- Visual linking between map parcels and table data  
- Tracks last sale date, most recent sale, market appraisal, and potential ROI  
- Built-in macros (VBA) for interactivity  
- Python script for scraping & preparing property data  

## Project Background

This project was developed as a **learning exercise** to practice:  
- Excel VBA for dynamic interactivity  
- Python for data scraping and preprocessing  
- Structuring a project for maintainability and clarity  

The focus was on understanding workflows, debugging, and connecting Excel tables with map visuals.  
Some parts may be experimental or imperfect, but the goal was to **practice problem-solving and improve coding skills**.

---

## Project Files
- `data/` → Sample or anonymized property data  
- `gifs/` → Demo of interactive workflow  
- `images/` → Screenshots of the tracker and map  
- `src/python/` → Data scraping and preprocessing scripts
- `src/vba/` → Macros and worksheet event code  

---

## Screenshots & Demo

### Excel Table Fresh/Unselected
![Excel Table Fresh](images/unselected.png)

### Interactive Map Selection
![Map Selection](images/map_select.png)

### Interactive Row Selection
![Row Selection](images/row_select.png)

### Demo (GIF)
![Workflow Demo](gifs/demo.gif)

---

## Usage

To recreate the interactive workbook locally:  

1. Open a new Excel workbook.  
2. Open the VBA editor (`Alt + F11`).  
3. Import the `.bas` and `.cls` files from `src/vba/` into the project.  
4. Enable macros in Excel.  
5. Use the sample data in `data/property_data.xlsx` to populate the table.  
6. Click on parcels or rows to see dynamic highlighting.

---

## Notes
- All personal and sensitive information has been anonymized  
- Parcel numbers are randomized between **1–21**  
- Owner names, street names, and sale dates are **fake for privacy**  
- A real map is included for visualization, but no resident-identifying data is exposed  

---

## Acknowledgments
This project used **OpenAI's ChatGPT (GPT-5 Mini)** for:  
- Web scraping assistance  
- Summarizing data  
- Formatting code  
- Randomizing private information  

All code was actively developed and refined with human oversight — issues were identified and corrected during the process.  

