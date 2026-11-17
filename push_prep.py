"""
Prepare Tapuach_Payments_Tracker.xlsx for a direct push into Michal's Google
Sheet (spreadsheetId 1_2HBIlRutX2BM09afFyJW8m-57RhmyGO7_1Q779EozA), replacing
the old flow of sending her the .xlsx to import by hand.

Run AFTER build.py and the recalc validation step (STEP 6), so
/tmp/recalc/Tapuach_Payments_Tracker.xlsx already exists with every formula
cell resolved to its computed value (data_only=True reads that correctly).

For each worksheet, writes one or more small JSON chunk files to
/home/claude/tapuach_report/sheets_push/, each named
  "<Sheet Name>__<start cell>.json"
e.g. "Transactions__A1.json", "Transactions__A151.json" -- the "<start cell>"
is exactly the top-left cell to pass as the `range` for that chunk's
update_values call (combine with the sheet name: "Transactions!A151").
Chunks are split by estimated JSON byte size (not a fixed row count), capped
well under the ~43KB point where the Read tool's line-numbered text view has
been seen to silently truncate a single huge line -- this matters because a
sheet with unusually long free-text cells (e.g. a verbose Notes column) can
blow past a fixed row-count chunk's size even when neighboring chunks are
fine, so a byte cap is what actually keeps every chunk safely readable.

Also writes /home/claude/tapuach_report/sheets_push/manifest.json:
  {"<Sheet Name>": {"clear_range": "<Sheet Name>", "chunks": ["<Sheet Name>__A1.json", ...]}, ...}
ordered the same order the chunks must be written in (top to bottom).

Usage for the daily job: run this script, then for each sheet in
manifest.json, call Google Sheets batch_clear_values on clear_range first
(wipes stale content before writing), then for each chunk file in its
"chunks" list (in order), read the file's raw JSON (via Bash/python -- NOT
the Read tool, which can silently truncate a single huge line) and call
update_values with range "<Sheet Name>!<start cell from filename>" and
values = the loaded 2D array.
"""
import json
import os
import re
import datetime
import openpyxl

SRC = "/tmp/recalc/Tapuach_Payments_Tracker.xlsx"
OUT_DIR = "/home/claude/tapuach_report/sheets_push"
MAX_CHUNK_BYTES = 30000  # stay well under the ~43KB Read-tool truncation point

os.makedirs(OUT_DIR, exist_ok=True)
for f in os.listdir(OUT_DIR):
    os.remove(os.path.join(OUT_DIR, f))

wb = openpyxl.load_workbook(SRC, data_only=True)

def convert(v):
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M:%S")
    return v

def col_letter(n):  # 0-indexed column -> "A", "B", ... "AA"
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

manifest = {}
for ws in wb.worksheets:
    rows = []
    for row in ws.iter_rows():
        rows.append([convert(c.value) for c in row])
    while rows and all(v is None for v in rows[-1]):
        rows.pop()
    max_col = 0
    for r in rows:
        for i in range(len(r) - 1, -1, -1):
            if r[i] is not None:
                max_col = max(max_col, i + 1)
                break
    rows = [r[:max_col] for r in rows]

    safe_name = re.sub(r"[^A-Za-z0-9 ]", "_", ws.title)
    chunks = []
    cur, cur_bytes, start_row = [], 2, 0  # 2 for "[]"
    for idx, r in enumerate(rows):
        r_bytes = len(json.dumps(r))
        if cur and cur_bytes + r_bytes > MAX_CHUNK_BYTES:
            fname = f"{safe_name}__A{start_row + 1}.json"
            with open(os.path.join(OUT_DIR, fname), "w") as f:
                json.dump(cur, f)
            chunks.append(fname)
            cur, cur_bytes, start_row = [], 2, idx
        cur.append(r)
        cur_bytes += r_bytes + 1
    if cur:
        fname = f"{safe_name}__A{start_row + 1}.json"
        with open(os.path.join(OUT_DIR, fname), "w") as f:
            json.dump(cur, f)
        chunks.append(fname)

    manifest[ws.title] = {"clear_range": ws.title, "chunks": chunks, "rows": len(rows), "cols": max_col}

with open(os.path.join(OUT_DIR, "manifest.json"), "w") as f:
    json.dump(manifest, f, indent=2)

print(json.dumps({k: {"rows": v["rows"], "cols": v["cols"], "chunk_count": len(v["chunks"])} for k, v in manifest.items()}, indent=2))
