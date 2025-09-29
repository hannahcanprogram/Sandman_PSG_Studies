# get_next_csv.py
import sys, os, json, csv, time

def to_str(v): return "" if v is None else str(v).strip()

def is_na(s):
    s = to_str(s).lower()
    return s in ("", "na", "n/a", "none", "nan")

def is_one(v):
    s = to_str(v).lower()
    if s == "": return False
    try:
        return float(s) == 1.0
    except Exception:
        return s in {"done", "complete"}

def should_skip(v):
    s = to_str(v).lower()
    return is_one(s) or s == "running"

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

def main(csv_path):
    headers, rows, enc = read_all_rows(csv_path)
    if not rows:
        print("{}"); return

    h_path = ci_lookup(headers, "Path")
    h_name = ci_lookup(headers, "Name")
    h_stat = ci_lookup(headers, "Status")

    # CSV 的第 1 行是表头，因此数据行的“人类行号”= 索引 + 2
    target_idx = None
    for i, row in enumerate(rows):
        if not should_skip(row.get(h_stat, "")):
            target_idx = i
            break

    if target_idx is None:
        print("{}"); return

    row = rows[target_idx]
    p = row.get(h_path, "")
    n = row.get(h_name, "")

    # 标记 running 并保存
    rows[target_idx][h_stat] = "running"
    write_all_rows_atomic(csv_path, headers, rows, enc)

    if is_na(p):
        print(json.dumps({"path": "na"})); return

    out = {"path": to_str(p), "name": to_str(n), "row": int(target_idx + 2)}
    print(json.dumps(out, ensure_ascii=False))

if __name__ == "__main__":
    try:
        csv_path = sys.argv[1]
        main(csv_path)
    except Exception as e:
        sys.stderr.write("ERROR: " + repr(e) + "\n")
        sys.stderr.flush()
        sys.exit(1)