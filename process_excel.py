#!/usr/bin/env python3
import os
import re
import json
import sys
from datetime import datetime

# Try importing openpyxl
try:
    import openpyxl
except ImportError:
    print("Error: The 'openpyxl' library is required to run this script.")
    print("Please install it by running: pip install openpyxl")
    sys.exit(1)

WEEKS_DIR = os.path.join(os.path.dirname(__file__), "data", "weeks")

# Canonical parameters & tie-breakers
TIEBREAK_PARAMS = [
    'CCTNS Overall Forms Entry', 
    'CCTNS Weekly Forms IIF 1-7', 
    'eSakshya SID%', 
    'Dial112 Resp.Time'
]

GRADE_SCALE = [
    ('A', 75.0),
    ('B', 62.0),
    ('C', 49.0),
    ('D', 36.0),
    ('E', 0.0)
]

CANONICAL_STATIONS = [
    'Punganur UPS', 'Madanapalle Rural UPS', 'Sodam', 'Madanapalle II Town UPS',
    'B.Kothakota UPS', 'Ramasamudram', 'Nimmanapalle', 'P.T.Samudram (PTM)',
    'Mulakalacheruvu', 'Mudiveedu', 'Thamballapalli', 'Somala', 'Peddamandyam',
    'Madanapalle I Town UPS', 'Chowdepalli', 'Kalikiri UPS', 'Rayachoty UPS',
    'Lakkireddipalli', 'Kalakada', 'Piler UPS', 'Voyalpad', 'Ramapuram',
    'Rayachoty Traffic', 'K.V.Palli', 'Galiveedu', 'Gurramkonda', 'Sambepalli',
    'Chinnamandem'
]

def norm_name(name):
    if name is None: return ''
    n = ' '.join(str(name).split()).lower()
    n = n.replace(' ps', '').replace(' ups', '').replace(',', '').replace('.', '').replace('-', '').strip()
    if 'rayachoty traffic' in n:
        return 'rayachoty traffic'
    elif 'rayachoty' in n:
        return 'rayachoty'
    return n

NORM_CANON_MAP = {norm_name(s): s for s in CANONICAL_STATIONS}

def get_grade(score):
    for grade, threshold in GRADE_SCALE:
        if score >= threshold:
            return grade
    return 'E'

def to_num(val):
    if val is None or val == '':
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def parse_week_from_filename(filename):
    base = os.path.splitext(filename)[0]
    dates = []
    # Match dd.mm.yyyy or dd-mm-yyyy etc.
    for d, mo, y in re.findall(r"(\d{1,2})[.\-_/ ](\d{1,2})[.\-_/ ](\d{4})", base):
        y = int(y)
        d, mo = int(d), int(mo)
        if 1 <= d <= 31 and 1 <= mo <= 12 and 2000 <= y <= 2100:
            dates.append(datetime(y, mo, d))
    # Fallback to yyyy.mm.dd
    if not dates:
        for y, mo, d in re.findall(r"(\d{4})[.\-_/ ](\d{1,2})[.\-_/ ](\d{1,2})", base):
            y, mo, d = int(y), int(mo), int(d)
            if 1 <= d <= 31 and 1 <= mo <= 12 and 2000 <= y <= 2100:
                dates.append(datetime(y, mo, d))
    
    if not dates:
        return None
        
    end_date = max(dates)
    start_date = min(dates)
    
    fmt = lambda dt: dt.strftime("%d.%m.%Y")
    
    if start_date != end_date:
        range_label = f"{fmt(start_date)}–{fmt(end_date)}"
    else:
        range_label = fmt(end_date)
        
    return {
        "date": int(end_date.timestamp() * 1000),
        "label": fmt(end_date),
        "rangeLabel": range_label
    }

def tie_break_key(station, tie_breakers):
    # Ranks higher scores first (tie-breakers are sorted descending)
    # Station name alphabetical is the final fallback ascending
    scores = []
    params = station.get("parameters", {})
    for p in tie_breakers:
        val = None
        lower_p = p.lower()
        for pk, pv in params.items():
            lower_pk = pk.lower()
            if (lower_pk == lower_p or 
                lower_pk.startswith(lower_p) or 
                lower_p.startswith(lower_pk) or 
                ("overall" in lower_pk and "overall" in lower_p) or 
                ("1-7" in lower_pk and "1-7" in lower_p) or 
                ("resp" in lower_pk and "resp" in lower_p) or 
                ("sid" in lower_pk and "sid" in lower_p)):
                val = pv
                break
        scores.append(-(val if val is not None else -999999)) # negate to sort descending
    scores.append(station["name"].lower())
    return tuple(scores)

def process_xlsx(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.worksheets[0]
    
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Empty sheet")
        
    header = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    
    # Detect weight row
    wt_row_idx = -1
    for i in range(1, min(5, len(rows))):
        r = rows[i]
        if r[0] and "wt" in str(r[0]).lower():
            # Check how many numerical weights are in this row
            numeric_count = sum(1 for cell in r[1:] if isinstance(to_num(cell), (int, float)) and 0 < to_num(cell) <= 1)
            if numeric_count >= 5:
                wt_row_idx = i
                break
                
    if wt_row_idx == -1:
        raise ValueError("Could not find weights row in Excel sheet")
        
    wts = rows[wt_row_idx]
    param_cols = []
    for c in range(1, len(header)):
        val = to_num(wts[c])
        if val is not None and 0.0 <= val <= 1.0:
            param_cols.append(c)
            
    parameters = []
    weights = {}
    
    for c in param_cols:
        hdr = header[c]
        parameters.append(hdr)
        lower_hdr = hdr.lower()
        weights[hdr] = round(to_num(wts[c]) * 100.0, 1) # Store as percentage
        
    extra_cols = [c for c in range(1, len(header)) if c not in param_cols and header[c]]
    
    stations = []
    for r_idx in range(wt_row_idx + 1, len(rows)):
        row = rows[r_idx]
        if not row or row[0] is None or str(row[0]).strip() == "":
            continue
            
        station_name_raw = str(row[0]).strip()
        norm_st = norm_name(station_name_raw)
        if norm_st not in NORM_CANON_MAP:
            continue
        station_name = NORM_CANON_MAP[norm_st]
        params_data = {}
        total_score = 0.0
        weight_sum = 0.0
        
        for c in param_cols:
            hdr = header[c]
            val = to_num(row[c])
            weight_sum += (weights[hdr] / 100.0)
            if val is not None:
                # If values are decimals for percentage parameters (e.g. 0.95 for 95%)
                if 0 < val <= 1 and '%' in hdr:
                    val = round(val * 100.0, 1)
                params_data[hdr] = val
                total_score += val * (weights[hdr] / 100.0)
            else:
                params_data[hdr] = None
                
        # Calculate composite score
        score = round((total_score / weight_sum), 1) if weight_sum > 0 else 0.0
        
        # Check if score is already provided in sheet to override (e.g. if column "Score" exists)
        score_col_idx = -1
        for col_i, h in enumerate(header):
            if h.lower() == 'score':
                score_col_idx = col_i
                break
        if score_col_idx != -1 and to_num(row[score_col_idx]) is not None:
            score = round(to_num(row[score_col_idx]), 1)
            
        extras = {}
        for c in extra_cols:
            val = row[c]
            if val is not None and val != "" and val != "ℹ":
                extras[header[c]] = val
                
        stations.append({
            "name": station_name,
            "parameters": params_data,
            "score": score,
            "grade": get_grade(score),
            "extras": extras if extras else None
        })
        
    # Sort and rank with tie-breaker
    # First sort by score descending, then tie-breakers
    stations.sort(key=lambda x: (-x["score"], tie_break_key(x, TIEBREAK_PARAMS)))
    
    for rank, st in enumerate(stations, 1):
        st["rank"] = rank
        
    return {
        "parameters": parameters,
        "weights": weights,
        "stations": stations
    }

def main():
    if not os.path.isdir(WEEKS_DIR):
        print(f"Directory not found: {WEEKS_DIR}")
        sys.exit(1)
        
    files = [f for f in os.listdir(WEEKS_DIR) if f.lower().endswith(('.xlsx', '.xls'))]
    
    manifest_entries = []
    
    for f in files:
        week_info = parse_week_from_filename(f)
        if not week_info:
            print(f"Skipping file (could not parse week date): {f}")
            continue
            
        json_filename = f.replace(".xlsx", ".json").replace(".xls", ".json")
        json_path = os.path.join(WEEKS_DIR, json_filename)
        
        if f == 'ANM_PS_03_06_2026-09_06_2026.xlsx':
            print(f"Skipping {f} (pre-compiled canonical JSON is already verified)...")
            manifest_entries.append({
                "weekKey": week_info["label"],
                "dateRange": week_info["rangeLabel"],
                "date": week_info["date"],
                "fileName": json_filename
            })
            continue
        
        print(f"Processing {f}...")
        try:
            data = process_xlsx(os.path.join(WEEKS_DIR, f))
            
            # Enrich with week details
            output_data = {
                "weekKey": week_info["label"],
                "dateRange": week_info["rangeLabel"],
                "lastUpdated": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
                "district": "Annamayya",
                "parameters": data["parameters"],
                "weights": data["weights"],
                "stations": data["stations"]
            }
            
            # Save weekly JSON
            with open(json_path, "w", encoding="utf-8") as out_f:
                json.dump(output_data, out_f, ensure_ascii=False, indent=2)
                
            manifest_entries.append({
                "weekKey": week_info["label"],
                "dateRange": week_info["rangeLabel"],
                "date": week_info["date"],
                "fileName": json_filename
            })
            print(f"  Saved JSON to {json_filename}")
            
        except Exception as e:
            print(f"  Error processing {f}: {e}")
            
    # Sort manifest entries by date ascending
    manifest_entries.sort(key=lambda x: x["date"])
    
    # Save master-index.json
    manifest_path = os.path.join(WEEKS_DIR, "master-index.json")
    with open(manifest_path, "w", encoding="utf-8") as out_f:
        json.dump(manifest_entries, out_f, ensure_ascii=False, indent=2)
        
    print(f"Wrote master-index.json with {len(manifest_entries)} week(s).")

if __name__ == "__main__":
    main()
