import argparse
import glob
import os
import re
import sys
from collections import Counter, defaultdict

import pandas as pd

KEYS = ["Site", "Project", "Board"]
QTY = ["Input Qty", "Output Qty", "Failure Qty"]
PCT = ["FPY (%)", "VPY (%)", "NDF"]
KEEP = KEYS + QTY + PCT
TOP_COLS = ["Top Issue 1", "Top Issue 2", "Top Issue 3"]
TOP_N = len(TOP_COLS)
PREFIX = "monthly"  # 輸出檔名：monthly_summary.csv / .xlsx

# "TestCaseName (12)" -> ("TestCaseName", 12)；名稱本身含括號也沒關係，只取最後一組 (數字)
TOP_RE = re.compile(r"^(.*\S)\s*\((\d+)\)\s*$")


def pct_to_num(v):
    """'98.50%' -> 98.5；'-'、空白 -> NaN"""
    if pd.isna(v):
        return float("nan")
    try:
        return float(str(v).strip().rstrip("%"))
    except ValueError:
        return float("nan")


def parse_top(v):
    """'name (12)' -> ('name', 12)；'-'、空白、格式不符 -> None"""
    if pd.isna(v):
        return None
    m = TOP_RE.match(str(v).strip())
    return (m.group(1).strip(), int(m.group(2))) if m else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True)   # YYYY-MM-DD
    ap.add_argument("--end", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--reports", default="reports")
    a = ap.parse_args()

    s = a.start.replace("-", "")
    e = a.end.replace("-", "")

    frames, days = [], []
    for path in sorted(glob.glob(f"{a.reports}/*/summary.csv")):
        tag = os.path.basename(os.path.dirname(path))
        if not re.fullmatch(r"\d{8}", tag) or not (s <= tag <= e):
            continue
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
        miss = [c for c in KEEP if c not in df.columns]
        if miss:
            sys.exit(f"{path} 缺少欄位：{miss}")
        for c in TOP_COLS:          # 舊的每日檔可能還沒有 Top Issue 欄位
            if c not in df.columns:
                df[c] = pd.NA
        frames.append(df[KEEP + TOP_COLS].copy())
        days.append(tag)

    if not frames:
        sys.exit(f"{a.start} ~ {a.end} 之間找不到任何 reports/*/summary.csv")

    print(f"找到 {len(days)} 天：{days[0]} ~ {days[-1]}")
    all_days = pd.date_range(a.start, a.end).strftime("%Y%m%d")
    missing = sorted(set(all_days) - set(days))
    if missing:
        print(f"::warning::缺少 {len(missing)} 天的資料：{missing}")

    df = pd.concat(frames, ignore_index=True)
    for k in KEYS:
        df[k] = df[k].fillna("")
    for c in QTY:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Top Issue：把每天括號裡的次數，依 (Site, Project, Board, testCase) 加總，再取前三名
    issue_cnt = defaultdict(Counter)
    for _, r in df.iterrows():
        key = (r["Site"], r["Project"], r["Board"])
        for c in TOP_COLS:
            t = parse_top(r[c])
            if t:
                issue_cnt[key][t[0]] += t[1]

    def top_cells(key):
        # 次數多的在前；次數相同時依名稱排序，結果才穩定
        ranked = sorted(issue_cnt.get(key, {}).items(), key=lambda kv: (-kv[1], kv[0]))[:TOP_N]
        cells = [f"{n} ({c})" for n, c in ranked]
        return cells + ["-"] * (TOP_N - len(cells))

    # 百分比欄位：以 Input Qty 加權平均（沒資料的天不列入分母）
    for p in PCT:
        df[p] = df[p].map(pct_to_num)
        df[f"_w_{p}"] = (df[p] * df["Input Qty"]).fillna(0)
        df[f"_m_{p}"] = df["Input Qty"].where(df[p].notna(), 0)

    agg = {c: "sum" for c in QTY}
    agg.update({f"_w_{p}": "sum" for p in PCT})
    agg.update({f"_m_{p}": "sum" for p in PCT})
    g = df.groupby(KEYS, as_index=False, sort=True).agg(agg)

    for p in PCT:
        g[p] = (g[f"_w_{p}"] / g[f"_m_{p}"].where(g[f"_m_{p}"] > 0)).map(
            lambda v: "" if pd.isna(v) else f"{v:.2f}%")

    tops = [top_cells((r["Site"], r["Project"], r["Board"])) for _, r in g.iterrows()]
    for i, c in enumerate(TOP_COLS):
        g[c] = [t[i] for t in tops]

    out = g[KEEP + TOP_COLS].copy()
    for c in QTY:
        out[c] = out[c].astype(int)

    os.makedirs(a.out, exist_ok=True)
    out.to_csv(f"{a.out}/{PREFIX}_summary.csv", index=False, encoding="utf-8-sig")
    xlsx = f"{a.out}/{PREFIX}_summary.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="summary")
        ws = w.sheets["summary"]
        for i, col in enumerate(out.columns, start=1):
            if col in TOP_COLS:
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 45
    print(f"已輸出 {len(out)} 列到 {a.out}")


if __name__ == "__main__":
    main()
