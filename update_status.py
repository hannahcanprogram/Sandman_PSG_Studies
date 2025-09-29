# update_status_csv_log.py
import sys, os, csv, time

# 日志文件路径，可以改成你自己的
LOG_CSV = r"G:\Sandman\LOGS\sandman_batch_log_group1B.csv"

def to_str(v): return "" if v is None else str(v).strip()

def ci_lookup(headers, name):
    tgt = name.lower()
    for h in headers:
        if to_str(h).lower() == tgt:
            return h
    raise ValueError(f"Header '{name}' not found (got: {headers})")

def read_all_rows(csv_path, encoding_list=("utf-8-sig","utf-8","latin-1")):
    last_err = None
    for enc in encoding_list:
        try:
            with open(csv_path, "r", newline="", encoding=enc) as f:
                r = csv.DictReader(f)
                rows = list(r)
                return r.fieldnames, rows, enc
        except Exception as e:
            last_err = e
    raise last_err

def write_all_rows_atomic(csv_path, headers, rows, encoding):
    tmp = csv_path + ".tmp"
    with open(tmp, "w", newline="", encoding=encoding) as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, csv_path)

def append_log(row, name, path, status, extra):
    try:
        os.makedirs(os.path.dirname(LOG_CSV), exist_ok=True)
        newfile = (not os.path.exists(LOG_CSV)) or os.path.getsize(LOG_CSV) == 0
        with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if newfile:
                w.writerow(["timestamp", "row", "name", "path", "status", "info"])
            w.writerow([
                time.strftime("%Y-%m-%d %H:%M:%S"),
                row, name, path, status, extra
            ])
    except Exception:
        pass

def main(csv_path, row_num, status, extra=""):
    headers, rows, enc = read_all_rows(csv_path)
    if row_num < 2 or row_num > len(rows) + 1:
        raise ValueError(f"Row {row_num} out of range (max={len(rows)+1}).")

    h_stat = ci_lookup(headers, "Status")
    h_name = ci_lookup(headers, "Name")
    h_path = ci_lookup(headers, "Path")
    h_extra = None
    for h in headers:
        if to_str(h).lower() == "extra":
            h_extra = h
            break

    idx = row_num - 2  # 数据行索引
    if status.lower() == "success":
        rows[idx][h_stat] = "1"
    else:
        rows[idx][h_stat] = ("error:" + to_str(extra))[:255]
    if h_extra:
        rows[idx][h_extra] = to_str(extra)

    name = to_str(rows[idx].get(h_name, ""))
    path = to_str(rows[idx].get(h_path, ""))

    write_all_rows_atomic(csv_path, headers, rows, enc)
    append_log(row_num, name, path, status, to_str(extra))

if __name__ == "__main__":
    try:
        csv_path = sys.argv[1]
        row_num  = int(sys.argv[2])
        status   = sys.argv[3]               # success / error / skip_lock ...
        extra    = sys.argv[4] if len(sys.argv)>4 else ""
        main(csv_path, row_num, status, extra)
    except Exception as e:
        sys.stderr.write("ERROR: " + repr(e) + "\n")
        sys.stderr.flush()
        sys.exit(1)