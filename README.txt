PCS APPLE-TO-APPLE DASHBOARD — complete package
================================================

TO EDIT DATA:
  1. Open PCS_Master_All.xlsx
  2. Edit the cream cells on any tab (see the HOW TO USE tab inside)
     - MATRIX tab = the 92-parameter comparison + scoring
     - EFFICIENCY / PQ / TEMP_DERATE / VOLTAGE_DERATE / LVRT_HVRT = detailed curves
     - CHART_POINTS = the summary charts (efficiency bars, DC window, overload)
  3. Save the file (keep it .xlsx)

TO GENERATE CSVs:
  4. Run:  python xlsx_to_csv_all.py PCS_Master_All.xlsx
     (needs Python + openpyxl:  pip install openpyxl)
  5. It creates all 7 CSVs.

TO PUBLISH:
  6. Upload the 7 CSVs + index.html to your GitHub repo (all in the same folder)
  7. Hard-refresh the dashboard: Ctrl+Shift+R (Windows) / Cmd+Shift+R (Mac)

FILES:
  index.html            the dashboard (never edit)
  PCS_Master_All.xlsx   <-- YOU EDIT THIS
  xlsx_to_csv_all.py    the converter
  data.csv              generated: 92-parameter matrix
  chart_points.csv      generated: summary charts
  curve_*.csv (x5)      generated: detailed curves

Vendor names must match exactly across all tabs.
Leave a cell blank = the dashboard shows "Awaiting data" for that point.
