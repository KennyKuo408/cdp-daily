import argparse
import glob
import os
import re
import sys

import pandas as pd

KEYS = ["Site", "Project", "Board"]
QTY = ["Input Qty", "Output Qty", "Failure Qty"]
PCT = ["FPY (%)", "VPY (%)", "NDF"]
KEEP = KEYS + QTY + PCT


def pct_to_num(v):
    """'98.50%' -> 98.5；'-'、空白 -> NaN"""
    if pd.isna(v):
        return float("nan")
    try:
        return float(str(v).strip().rstrip("%"))
    except ValueError:
        return float("nan")


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
        frames.append(df[KEEP].copy())
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

    out = g[KEEP]
    for c in QTY:
        out[c] = out[c].astype(int)

    os.makedirs(a.out, exist_ok=True)
    out.to_csv(f"{a.out}/monthly_summary.csv", index=False, encoding="utf-8-sig")
    out.to_excel(f"{a.out}/monthly_summary.xlsx", index=False)
    print(f"已輸出 {len(out)} 列到 {a.out}")


if __name__ == "__main__":
    main()
