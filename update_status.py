# update_status.py
import sys, csv, time
import win32com.client as win32

LOG_CSV = r"D:\PSG_test\sandman_batch_log.csv"

def find_col(ws, header):
    for col in range(1, 200):
        v = (ws.Cells(1, col).Value or "").strip() if ws.Cells(1, col).Value else ""
        if v.lower() == header.lower():
            return col
    raise RuntimeError(f"Column '{header}' not found.")

def main(xlsx, sheet, row, status, extra):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(xlsx)
    try:
        ws = wb.Worksheets(sheet)
        col_stat = find_col(ws, "Status")
        col_name = find_col(ws, "Name")
        col_path = find_col(ws, "Path")

        if status == "success":
            ws.Cells(row, col_stat).Value = "1"
        else:
            ws.Cells(row, col_stat).Value = f"error:{extra}"[:255]

        wb.Save()

        with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if f.tell() == 0:
                w.writerow(["timestamp","row","name","path","status","info"])
            w.writerow([time.strftime("%Y-%m-%d %H:%M:%S"),
                        row,
                        ws.Cells(row, col_name).Value,
                        ws.Cells(row, col_path).Value,
                        status,
                        extra])
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()

if __name__ == "__main__":
    xlsx  = sys.argv[1]
    sheet = sys.argv[2]
    row   = int(sys.argv[3])
    status= sys.argv[4]            # success / error
    extra = sys.argv[5] if len(sys.argv)>5 else ""
    main(xlsx, sheet, row, status, extra)
