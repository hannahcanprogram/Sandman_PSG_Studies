# get_next.py
import sys, json
import win32com.client as win32

"""
how to use this utility:
python get_next.py "D:\\PSG_test\\test_paths.xlsx" "Sheet1"
stdout: {"path":"...", "name":"...", "row":12}
nothing to return: {}
path invalid: {"path":"na"}
"""

XL_UP = -4162  # xlUp

def to_str(v):
    if v is None:
        return ""
    return str(v).strip()

def find_col(ws, header):
    target = header.lower()
    for col in range(1, 256):
        v = to_str(ws.Cells(1, col).Value).lower()
        if v == target:
            return col
    raise RuntimeError(f"Column '{header}' not found in row 1.")

def is_na(s: str) -> bool:
    s = to_str(s).lower()
    return s in ("", "na", "n/a", "none", "nan")


def is_one(v) -> bool:
    if v is None:
        return False
    if isinstance(v, (int, float)):
        try:
            return float(v) == 1.0
        except Exception:
            return False
    s = to_str(v).lower()
    if s == "":
        return False
    try:
        return float(s) == 1.0
    except Exception:
        return s in {"done", "complete"}

def should_skip(v) -> bool:
    s = to_str(v).lower()
    return is_one(v) or s == "running"

def main(xlsx, sheet="Sheet1"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(xlsx)
    try:
        ws = wb.Worksheets(sheet)
        c_path = find_col(ws, "Path")
        c_name = find_col(ws, "Name")
        c_stat = find_col(ws, "Status")

        last_row = ws.Cells(ws.Rows.Count, c_path).End(XL_UP).Row

        target_row = None
        for r in range(2, last_row + 1):
            status_raw = ws.Cells(r, c_stat).Value
            if not should_skip(status_raw):
                target_row = r
                break

        if not target_row:
            print("{}"); sys.stdout.flush()
            return

        p = ws.Cells(target_row, c_path).Value
        n = ws.Cells(target_row, c_name).Value

        ws.Cells(target_row, c_stat).Value = "running"
        wb.Save()

        if is_na(p):
            print(json.dumps({"path": "na"})); sys.stdout.flush()
            return

        out = {"path": to_str(p), "name": to_str(n), "row": int(target_row)}
        print(json.dumps(out)); sys.stdout.flush()

    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()

if __name__ == "__main__":
    try:
        xlsx = sys.argv[1]
        sheet = sys.argv[2] if len(sys.argv) > 2 else "Sheet1"
        main(xlsx, sheet)
    except Exception as e:
        sys.stderr.write("ERROR: " + str(e) + "\n")
        sys.stderr.flush()
        sys.exit(1)
