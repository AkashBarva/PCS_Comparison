#!/usr/bin/env python3
"""
Turn the master Excel (PCS_Master_All.xlsx) into ALL the CSV files the
dashboard reads. One command, all 7 files.

Usage:
    python xlsx_to_csv_all.py                       # uses PCS_Master_All.xlsx here
    python xlsx_to_csv_all.py MyFile.xlsx           # or point at a specific file

Tab  ->  CSV mapping:
    MATRIX          -> data.csv
    EFFICIENCY      -> curve_efficiency.csv
    PQ              -> curve_pq.csv
    TEMP_DERATE     -> curve_temp_derate.csv
    VOLTAGE_DERATE  -> curve_voltage_derate.csv
    LVRT_HVRT       -> curve_lvrt_hvrt.csv

The header row is row 2 on every tab (row 1 is the yellow instruction note).
Upload all generated CSVs to your GitHub repo next to index.html, then
hard-refresh the dashboard.
"""
import sys, csv, os

TABS = {
    "MATRIX":         "data.csv",
    "EFFICIENCY":     "curve_efficiency.csv",
    "PQ":             "curve_pq.csv",
    "TEMP_DERATE":    "curve_temp_derate.csv",
    "VOLTAGE_DERATE": "curve_voltage_derate.csv",
    "LVRT_HVRT":      "curve_lvrt_hvrt.csv",
    "AUX_CONSUMPTION": "curve_aux.csv",
    "CERTIFICATES":   "curve_certs.csv",
    "CHART_POINTS":   "chart_points.csv",
    "IGBT_INFO":      "igbt_info.csv",
    "IGBT_TEMP":      "curve_igbt_temp.csv",
}

def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else "PCS_Master_All.xlsx"
    if not os.path.exists(xlsx):
        print(f"ERROR: can't find {xlsx}.")
        print("Put this script next to the Excel file, or pass the name:")
        print("   python xlsx_to_csv_all.py PCS_Master_All.xlsx")
        sys.exit(1)

    try:
        from openpyxl import load_workbook
    except ImportError:
        print("ERROR: openpyxl not installed. Run:  pip install openpyxl")
        sys.exit(1)

    wb = load_workbook(xlsx, data_only=True)
    made = 0
    for tab, out in TABS.items():
        if tab not in wb.sheetnames:
            print(f"  skip {tab:15}  (tab not found)")
            continue
        ws = wb[tab]
        rows = list(ws.values)
        # header is row 2 (index 1); row 1 is the instruction note
        table = rows[1:]
        cleaned = []
        for r in table:
            r = ["" if v is None else str(v).strip() for v in r]
            if any(cell != "" for cell in r):
                cleaned.append(r)
        # trim trailing all-empty columns that Excel sometimes adds
        if cleaned:
            width = len(cleaned[0])
            cleaned = [r[:width] for r in cleaned]
        with open(out, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(cleaned)
        n = max(len(cleaned) - 1, 0)
        print(f"  OK  {tab:15}  ->  {out:26} ({n} rows)")
        made += 1

    print(f"\nGenerated {made} CSV files.")
    print("Upload them to your GitHub repo (next to index.html), then hard-refresh the dashboard.")

if __name__ == "__main__":
    main()


# ============================================================
# AUTO-GENERATE chart_points.csv from the matrix + curve tabs
# (so it never drifts out of sync with vendors again)
# ============================================================
def autogen_chart_points(xlsx="PCS_Master_All.xlsx"):
    import re
    from openpyxl import load_workbook
    wb = load_workbook(xlsx, data_only=True)
    ws = wb["MATRIX"]
    V = [str(ws.cell(2,c).value).split("\n")[0].strip()
         for c in range(5, ws.max_column+1) if ws.cell(2,c).value]

    def matrixrow(name):
        for r in range(3, ws.max_row+1):
            if str(ws.cell(r,3).value or "").strip().lower() == name.lower():
                return {V[i]: ws.cell(r,5+i).value for i in range(len(V))}
        return {}

    def num(x):
        if x is None: return ""
        s = str(x).strip()
        if s.upper() in ("","NA","N/A","-","TBC","NOT STATED"): return ""
        m = re.search(r"\d+\.?\d*", s.replace(",",""))
        if not m: return ""
        val = float(m.group(0))
        # normalise efficiency written as fraction (0.988 -> 98.8)
        if 0 < val <= 1: val = val*100
        return str(round(val,2)) if val%1 else str(int(val))

    peak = matrixrow("Peak efficiency")
    euro = matrixrow("Euro / CEC efficiency")
    dcw  = matrixrow("DC voltage window")
    ovl  = matrixrow("Overload capability")

    out = [["chart","vendor","x","y"]]

    # derate: pull real kW/°C points from TEMP_DERATE tab if present
    if "TEMP_DERATE" in wb.sheetnames:
        w = wb["TEMP_DERATE"]
        for r in range(3, w.max_row+1):
            v = w.cell(r,1).value
            if v is None or str(v).strip() not in V: continue
            x = w.cell(r,2).value; y = w.cell(r,3).value
            if x not in (None,"") and y not in (None,""):
                out.append(["derate", str(v).strip(), x, y])

    # dc window
    for v in V:
        s = str(dcw.get(v,"") or "")
        nums = re.findall(r"\d{3,4}", s.replace(",",""))
        if len(nums) >= 2:
            out.append(["dcwin", v, min(map(int,nums)), 1500])

    # overload (parse simple points)
    for v in V:
        s = str(ovl.get(v,"") or "").lower()
        if "150%" in s or "150 %" in s: out.append(["overload", v, 1, 150])
        if "120" in s: out.append(["overload", v, 10, 120])
        if "110" in s: out.append(["overload", v, "cont", 110])

    # efficiency (all vendors, normalised)
    for v in V:
        out.append(["eff_peak", v, "", num(peak.get(v,""))])
        out.append(["eff_euro", v, "", num(euro.get(v,""))])

    with open("chart_points.csv","w",newline="",encoding="utf-8") as f:
        csv.writer(f).writerows(out)
    print(f"  AUTO  chart_points.csv regenerated from matrix ({len(V)} vendors, all in sync)")

# run autogen after the normal export
if __name__ == "__main__":
    try:
        _xlsx = sys.argv[1] if len(sys.argv) > 1 else "PCS_Master_All.xlsx"
        autogen_chart_points(_xlsx)
    except Exception as e:
        print("  (chart_points autogen skipped:", e, ")")
