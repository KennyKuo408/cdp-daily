#!/usr/bin/env python3
"""
每日資料擷取 + FPY 分析 一站式腳本

流程:
  1. 連線 Databricks SQL Warehouse,分別查詢 WYMY / WYMTN 兩張表
  2. 對每張表的查詢結果直接做 FPY / VPY / Cycle Time 分析(邏輯照搬自 FPY.py)
  3. 分析結果直接輸出成多分頁 .xlsx,檔名規則跟原本抓資料時一樣:
       WYMY_<日期>_<model>_<teststage>.xlsx
       WYMTN_<日期>_<model>_<teststage>.xlsx
     (不會另外產生一份原始單分頁檔案,分析結果直接用這個檔名蓋掉)

環境變數(跟 fetch_daily_data.py 一樣):
    DATABRICKS_HOST      例如 adb-xxxxxxxxxxxx.x.azuredatabricks.net
    DATABRICKS_HTTP_PATH 例如 /sql/1.0/warehouses/xxxxxxxxxxxx
    DATABRICKS_TOKEN     你的 Personal Access Token(DBeaver Password 欄位那組,不是 Username)

用法:
    python3 fetch_and_fpy.py
    python3 fetch_and_fpy.py --start 2026-09-13 --end 2026-09-14
    python3 fetch_and_fpy.py --start 2026-09-13 --end 2026-09-14 --model YV4 --teststage MB

注意:
    FPY 分析邏輯假設查詢結果裡有這些欄位:
        usn, tresult, startTimeLocal, endTimeLocal, startTime,
        plant, customer, childusn, stage
    這些欄位名稱跟大小寫要跟 Databricks 表裡的一致(SELECT * 出來的欄位名稱
    就是表裡原本的欄位名,理論上會跟你原本 CSV 版本一致,但第一次跑完後
    建議打開「原始資料」分頁確認一下欄位都對得上)。
"""

import os
import re
import sys
import argparse
from datetime import datetime, timedelta


# ---------- SSL: 公司網路有 SSL 攔截,在公司 Mac 上需要用合併過公司根證書的信任清單 ----------
# 只有在 Mac 上找得到這份檔案時才套用。GitHub Actions 的機器不經過公司代理,不需要也沒有這個檔案。
_CA_BUNDLE = os.path.expanduser("~/Library/CloudStorage/OneDrive-WiwynnCorporation/CDP/combined-ca-bundle.pem")
if os.path.exists(_CA_BUNDLE):
    os.environ.setdefault("SSL_CERT_FILE", _CA_BUNDLE)
    os.environ.setdefault("REQUESTS_CA_BUNDLE", _CA_BUNDLE)
elif not os.environ.get("GITHUB_ACTIONS"):
    print(f"[提醒] 找不到 {_CA_BUNDLE},如果之後出現 SSL CERTIFICATE_VERIFY_FAILED 錯誤,"
          f"代表需要先重新建立這份合併憑證檔案。")

try:
    from databricks import sql
except ImportError:
    print("缺少套件,請先執行: pip3 install databricks-sql-connector pandas openpyxl")
    sys.exit(1)

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


TABLES = {
    "WYMY": "dw_manufacturing.viw_tpspcbainfo_wymy",
    "WYMTN": "dw_manufacturing.viw_tpspcbainfo_wymtn",
    "WMX": "dw_manufacturing.viw_tpspcbainfo_wmx"
}

# 要查的 (site, model, teststage) 組合清單。要加/減組合,只要改這裡,不用動其他程式碼。
# site 要對應 TABLES 裡的 key("WYMY" / "WYMTN")。不是每個 model/teststage 都兩邊都有,
# 用三元素就能只列出某個 site 真的有在跑的組合,不會多查到另一邊去。
# 跑一次 main() 會把下面每一組都查一遍,全部結果一起進同一份彙總表。
JOBS = [
    #LoB1
    ("WYMTN", "YV4", "MB"),("WYMTN", "NL", "MB"),("WYMTN", "GC", "ServerBoard"),
    ("WYMY", "YV4", "MB"),("WYMY", "NL", "MB"),("WYMY", "GC", "FDPB"),("WYMY", "GC", "RDPB"),
    
    #LoB2
    ("WYMTN", "C41A8", "MB"),("WYMTN", "C219B", "MB"),
    ("WYMY", "C41A8", "MB"),("WYMY", "M1120", "SCM"),("WYMY", "C219B", "MB"),
    ("WMX", "M1120", "SCM"), ("WMX", "C219B", "MB"),
    
    
    
    
    
]

BLANK_COLS = ["plant", "customer", "usn", "childusn", "stage"]


# ---------- YV4 / MB 特例分類(其他 model/teststage 組合完全不受影響) ----------

GEN_TYPE_LOOKUP = {
    # 之後補: upn -> (gen, type)
    # 例如: "B91.10H1": ("GEN4", "T1C"),
    # WMMTN
    "B91.10H10.0019":("GEN5","T1M"),
    "B91.10H10.001C":("GEN5","T1C"), "B91.10H10.0051":("GEN5","T1C"),
    "B91.10H10.0017":("GEN4","T1M"), "B91.10H10.0018":("GEN4","T1M"), "B91.10H10.003K":("GEN4","T1M"), "B91.10H10.003T":("GEN4","T1M"),
    "B91.10H10.0055":("GEN4","T1C"), "B91.10H10.0058":("GEN4","T1C"), "B91.10H10.005A":("GEN4","T1C"), "B91.10H10.0059":("GEN4","T1C"), "B91.10H10.001B":("GEN4","T1C"), "B91.10H10.001A":("GEN4","T1C"),
    #WYMY
    "B91.10H10.0037":("GEN5","T1M"), "B91.10H10.001Z":("GEN5","T1M"), "B91.10H10.004V":("GEN5","T1M"), "B91.10H10.003W":("GEN5","T1M"), "B91.10H10.003E":("GEN5","T1M"),
    "B91.10H10.0054":("GEN5","T1C"), "B91.10H10.003G":("GEN5","T1C"), "B91.10H10.0023":("GEN5","T1C"), "B91.10H10.0038":("GEN5","T1C"), "B91.10H10.003Y":("GEN5","T1C"),
    "B91.10H10.004P":("GEN4","T1M"), "B91.10H10.001X":("GEN4","T1M"), "B91.10H10.001Y":("GEN4","T1M"),
    "B91.10H10.004X":("GEN4","T1C"), "B91.10H10.0022":("GEN4","T1C"), "B91.10H10.0021":("GEN4","T1C")

}

STD_LOOKUP = {
    # 之後補: (site, model, board) -> std 數值(分鐘,直接對應彙總表 Std(sec) 那欄,不再用 P90-P10 現場算)
    # 例如: ("WYMY", "YV4", "GEN4_T1C_AUTO"): 12.34,
    #WYMTN
    ("WYMTN", "YV4", "GEN5_T1M"): 1980, ("WYMTN", "YV4", "GEN5_T1C"): 1800, ("WYMTN", "YV4", "GEN4_T1M"): 1980, ("WYMTN", "YV4", "GEN4_T1C"): 1800,
    ("WYMTN", "YV4", "SPD"): 1410, ("WYMTN", "YV4", "WF"): 750, ("WYMTN", "YV4", "FAN"): 300,("WYMTN", "YV4", "MEDUSA"):250 ,
    ("WYMTN", "YV4", "MGM"): 750,("WYMTN", "YV4", "FIO"): 40, ("WYMTN", "YV4", "BSM"): 200,
    ("WYMTN", "NL", "MB"): 5300,("WYMTN", "OLPC 2", "MB"):1520 ,("WYMTN", "OLPC 2", "EXPA"):1250,("WYMTN", "OLPC 2", "EXPB"):1380 ,
    ("WYMTN", "OLPC 2", "ITPS"):60 ,("WYMTN", "OLPC 2", "BB"):1265 ,("WYMTN", "OLPC 2", "BSM"):240 ,("WYMTN", "OLPC 2", "FIO"):60 ,
    ("WYMTN", "OLPC 2", "MEDUSA"):180 ,("WYMTN", "OLPC 2", "TPM"):340 ,
    #WYMY
    ("WYMY", "YV4", "GEN5_T1M_AUTO"):2449 ,("WYMY", "YV4", "GEN5_T1C_AUTO"):1713 ,("WYMY", "YV4", "GEN4_T1M_AUTO"):2142, ("WYMY", "YV4", "GEN4_T1C_AUTO"):1602,
    ("WYMY", "YV4", "GEN5_T1M_MANUAL"):2449 ,("WYMY", "YV4", "GEN5_T1C_MANUAL"):1713 ,("WYMY", "YV4", "GEN4_T1M_MANUAL"):2142, ("WYMY", "YV4", "GEN4_T1C_MANUAL"):1602,
    ("WYMY", "YV4", "FAN"):152, ("WYMY", "YV4", "FIO"):40,("WYMY", "YV4", "BSM"):158,("WYMY", "YV4", "SPD"):1503,("WYMY", "YV4", "MGM"):756,
    ("WYMY", "YV4", "WF"):1239,("WYMY", "GC", "FDPB"):1139,("WYMY", "GC", "RDPB"):1218,("WYMY", "NL", "MB"):5333,
    
}

FIXTURE_NUM_RE = re.compile(r"YV4_MB_(\d{3})_MP", re.IGNORECASE)
AUTO_THRESHOLD = 156


def lookup_gen_type(upn):
    """upn -> (gen, type) 對照。查不到回傳 None,呼叫端要處理查不到的狀況。"""
    return GEN_TYPE_LOOKUP.get(upn)


def extract_fixture_number(fixture_id):
    """從 fixtureId 抓出 YV4_MB_XXX_MP 裡的 XXX(固定三位數字)。抓不到回傳 None。"""
    if not isinstance(fixture_id, str):
        return None
    m = FIXTURE_NUM_RE.search(fixture_id.upper())
    return int(m.group(1)) if m else None


def classify_board(site, model, teststage, upn, fixture_id):
    """
    YV4/MB 特例: 依 upn 分 GEN4/GEN5 x T1C/T1M;WYMY 再依 fixtureId 數字門檻加 _AUTO/_MANUAL。
    其他 model/teststage 組合完全照舊,直接回傳 teststage,行為跟改之前一樣。
    """
    if (model, teststage) != ("YV4", "MB"):
        return teststage

    gen_type = lookup_gen_type(upn)
    if gen_type is None:
        return "UNKNOWN"  # GEN_TYPE_LOOKUP 還沒填,或這個 upn 真的對不到
    gen, ttype = gen_type
    board = f"{gen}_{ttype}"

    if site == "WYMY":
        n = extract_fixture_number(fixture_id)
        board += "_UNKNOWN" if n is None else ("_AUTO" if n >= AUTO_THRESHOLD else "_MANUAL")

    return board


# ---------- 輸出路徑(OneDrive 同步資料夾,不是本機 Desktop) ----------
 
# 有設定環境變數 OUTPUT_DIR(GitHub Actions 上會設)就存到那裡,
# 沒設定(在 Mac 上直接跑)就照舊存到 OneDrive 同步資料夾。
ONEDRIVE_CDP_DIR = os.path.expanduser("~/Library/CloudStorage/OneDrive-WiwynnCorporation/CDP")
ONEDRIVE_DAILY_DIR = os.environ.get("OUTPUT_DIR") or os.path.join(ONEDRIVE_CDP_DIR, "daily")
 

# ---------- 連線 / 抓資料 ----------

CREDENTIALS_FILE = os.path.expanduser("~/Library/CloudStorage/OneDrive-WiwynnCorporation/CDP/.databricks_credentials")


def load_credentials_file():
    """
    退回讀取本機憑證檔案 ~/.databricks_credentials(格式: KEY=VALUE,一行一個)。
    這個檔案跟你目前用哪個 Terminal 視窗、哪個 shell(bash/zsh)完全無關,
    之後排程執行(cron/launchd)也一定讀得到,不會再有「換個地方跑就找不到
    環境變數」的問題。沒有這個檔案就回傳空字典,不報錯(改交給呼叫端判斷缺什麼)。
    """
    if not os.path.exists(CREDENTIALS_FILE):
        return {}
    values = {}
    with open(CREDENTIALS_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            values[key.strip()] = val.strip().strip('"').strip("'")
    return values


def get_connection_params():
    """
    連線資訊優先順序: 環境變數 > ~/.databricks_credentials 檔案。
    這樣不管你是用 export 設好的 Terminal、還是換一台/換個路徑執行,
    只要那份憑證檔案存在,就一定能連上。
    """
    file_values = load_credentials_file()

    host = os.environ.get("DATABRICKS_HOST") or file_values.get("DATABRICKS_HOST")
    http_path = os.environ.get("DATABRICKS_HTTP_PATH") or file_values.get("DATABRICKS_HTTP_PATH")
    token = os.environ.get("DATABRICKS_TOKEN") or file_values.get("DATABRICKS_TOKEN")

    missing = [name for name, val in [
        ("DATABRICKS_HOST", host),
        ("DATABRICKS_HTTP_PATH", http_path),
        ("DATABRICKS_TOKEN", token),
    ] if not val]

    if missing:
        print(f"缺少連線資訊: {', '.join(missing)}")
        print(f"請設定環境變數,或建立檔案 {CREDENTIALS_FILE},內容例如:")
        print("  DATABRICKS_HOST=adb-xxxx.azuredatabricks.net")
        print("  DATABRICKS_HTTP_PATH=/sql/1.0/warehouses/xxxx")
        print("  DATABRICKS_TOKEN=你的token (DBeaver Password 欄位那組,不是 Username)")
        sys.exit(1)

    return host, http_path, token


def compute_default_window():
    now = datetime.now()
    today_8am = now.replace(hour=7, minute=0, second=0, microsecond=0)
    end = today_8am if now >= today_8am else today_8am - timedelta(days=1)
    start = end - timedelta(days=1)
    return start, end


def parse_args():
    parser = argparse.ArgumentParser(description="抓取每日資料並跑 FPY 分析 (WYMY + WYMTN)")
    parser.add_argument("--start", type=str, help="窗口起始日期 YYYY-MM-DD (預設抓昨天08:00)")
    parser.add_argument("--end", type=str, help="窗口結束日期 YYYY-MM-DD (預設抓今天08:00)")
    parser.add_argument("--model", type=str, default=None,
                         help="只查單一 model(搭配 --teststage 一起用)。不指定的話,改成跑 JOBS 清單裡全部組合")
    parser.add_argument("--teststage", type=str, default=None,
                         help="只查單一 testStage(搭配 --model 一起用)。不指定的話,改成跑 JOBS 清單裡全部組合")
    parser.add_argument("--site", type=str, default=None, choices=list(TABLES.keys()),
                         help="搭配 --model/--teststage 時,只查這一個 site;不給的話兩個 site 都查")
    return parser.parse_args()


def fetch_table(connection, table_name, model, teststage, start_str, end_str):
    query = f"""
        SELECT *
        FROM {table_name}
        WHERE model = %(model)s
            AND testStage = %(teststage)s
            AND stage IN ('TD')
            AND UPPER(fixtureId) NOT LIKE '%DOA%'
            AND startTime >= %(start)s
            AND startTime < %(end)s
        ORDER BY endTimeLocal ASC
    """
    cursor = connection.cursor()
    cursor.execute(query, {
        "model": model,
        "teststage": teststage,
        "start": start_str,
        "end": end_str,
    })
    columns = [desc[0] for desc in cursor.description]
    rows = cursor.fetchall()
    return pd.DataFrame(rows, columns=columns)


def stringify_like_csv(df):
    """把資料轉成跟原本 pd.read_csv(dtype=str) 一樣的型態:
    非空值轉字串,缺值維持 NaN(不會變成字串 'nan' / 'None'),
    這樣後面沿用 FPY.py 原本的字串比較邏輯才會一致。"""
    return df.apply(lambda col: col.map(lambda v: v if pd.isna(v) else str(v)))


# ---------- FPY 分析(邏輯照搬 FPY.py,改成吃記憶體中的 DataFrame、寫到指定路徑) ----------

def safe_counta(sheet_name, col_letter, n_rows):
    """安全建立COUNTA公式字串。n_rows=0時直接回傳"=0",
    避免產生 2:1 這種頭尾顛倒的範圍(Excel會自動normalize成1:2,
    導致誤把標題列也算進去,多算1筆)。"""
    if n_rows == 0:
        return "=0"
    return f"=COUNTA({sheet_name}!{col_letter}2:{col_letter}{1+n_rows})"


def blank_repeat_rows(data: pd.DataFrame, cols_to_blank):
    out = data.copy()
    is_first = out.groupby("usn").cumcount() == 0
    for col in cols_to_blank:
        if pd.api.types.is_numeric_dtype(out[col]):
            out.loc[~is_first, col] = float("nan")
        else:
            out.loc[~is_first, col] = ""
    return out, is_first


def run_fpy_analysis(df: pd.DataFrame, out_path: str):
    """對一份原始查詢結果做 FPY/VPY/Cycle Time 分析,輸出多分頁 xlsx 到 out_path。"""
    df = stringify_like_csv(df)
    df["tresult"] = df["tresult"].astype(str).str.strip()

    # 依usn排序(stable保留同SN內原本時間先後順序),方便閱讀、也讓後面的分組blank邏輯視覺連續
    df = df.sort_values(by="usn", kind="stable").reset_index(drop=True)

    # ---- cycleTime改成「每個SN只有一個值」:該SN所有列裡,最晚endTimeLocal - 最早startTimeLocal ----
    _start_dt = pd.to_datetime(df["startTime"], errors="coerce")
    _end_dt = pd.to_datetime(df["endTime"], errors="coerce")
    _row_duration = (_end_dt - _start_dt).dt.total_seconds()

    _cycle_by_sn = (
        pd.DataFrame({"usn": df["usn"], "_dur": _row_duration})
        .groupby("usn")["_dur"]
        .sum()
        .round(3)
    )
    cycle_map = _cycle_by_sn
    df["cycleTime"] = df["usn"].map(cycle_map)

    sn_has_zero = set(df.loc[df["tresult"] == "0", "usn"])

    sheet2_df = df[df["usn"].isin(sn_has_zero)].copy()
    sheet2_df = sheet2_df.sort_values(["usn", "startTime"], kind="stable").reset_index(drop=True)

    remaining_df = df[~df["usn"].isin(sn_has_zero)].copy()
    remaining_df = remaining_df.sort_values(["usn", "startTime"], kind="stable").reset_index(drop=True)

    sheet3_df = remaining_df.copy()

    dup_mask = remaining_df.duplicated(subset="usn", keep="first")
    sheet4_df = remaining_df[dup_mask].copy()

    only_zero_usn_series = sheet2_df.groupby("usn")["tresult"].apply(lambda x: (x == "0").all())
    only_zero_usn = only_zero_usn_series[only_zero_usn_series].index.tolist()
    only_zero_usn_count = len(only_zero_usn)

    # ---- 額外用 Python 直接算出跟「統計」分頁公式對應的數值 ----
    # (Excel 公式寫進 xlsx 後,openpyxl 讀不到算好的結果,只有真的用 Excel/Numbers
    #  打開存檔過才有快取值。要做跨表彙總表,直接在這裡用 Python 算一次最單純可靠。)
    fail_sn_count = sheet2_df["usn"].nunique()
    pass_sn_count = sheet3_df["usn"].nunique()
    base_input_qty = fail_sn_count + pass_sn_count
    base_output_qty = base_input_qty - only_zero_usn_count

    if fail_sn_count > 0:
        _grp = sheet2_df.groupby("usn")["tresult"]
        _has0 = _grp.apply(lambda s: (s == "0").any())
        _has1 = _grp.apply(lambda s: (s == "1").any())
        ndf_count = int((_has0 & _has1).sum())
    else:
        ndf_count = 0

    fpy = (base_input_qty - fail_sn_count) / base_input_qty if base_input_qty else None
    vpy = (base_input_qty - fail_sn_count + ndf_count) / base_input_qty if base_input_qty else None
    ndf_rate = (vpy - fpy) if (fpy is not None and vpy is not None) else None

    _valid_ct = cycle_map.drop(index=[u for u in only_zero_usn if u in cycle_map.index], errors="ignore").dropna()
    if len(_valid_ct) > 0:
        p10_sec = _valid_ct.quantile(0.1)
        p50_sec = _valid_ct.quantile(0.5)
        p90_sec = _valid_ct.quantile(0.9)
        spread_sec = p90_sec - p10_sec
    else:
        p10_sec = p50_sec = p90_sec = spread_sec = None

    metrics = {
        "base_input_qty": base_input_qty,
        "base_output_qty": base_output_qty,
        "failure_qty": fail_sn_count,
        "pass_qty": pass_sn_count,
        "ndf_count": ndf_count,
        "fpy": fpy,
        "vpy": vpy,
        "ndf_rate": ndf_rate,
        "p10_sec": p10_sec,
        "p50_sec": p50_sec,
        "p90_sec": p90_sec,
        "spread_sec": spread_sec,
    }

    df_display, _ = blank_repeat_rows(df, ["cycleTime"])
    df_display.loc[df_display["usn"].isin(only_zero_usn), "cycleTime"] = float("nan")

    sheet2_display, sheet2_is_first = blank_repeat_rows(sheet2_df, BLANK_COLS + ["cycleTime"])
    sheet3_display, _ = blank_repeat_rows(sheet3_df, BLANK_COLS + ["cycleTime"])

    sheet2_display["_helper_usn"] = sheet2_df["usn"].values
    sheet2_display["_is_first"] = sheet2_is_first.astype(int).values

    onlyfail_detail_df = sheet2_df[sheet2_df["usn"].isin(only_zero_usn)].copy()
    onlyfail_detail_df = onlyfail_detail_df.sort_values(["usn", "startTime"], kind="stable").reset_index(drop=True)
    onlyfail_display, _ = blank_repeat_rows(onlyfail_detail_df, BLANK_COLS + ["cycleTime"])

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_display.to_excel(writer, sheet_name="原始資料", index=False)
        sheet2_display.to_excel(writer, sheet_name="Fail資料", index=False)
        sheet3_display.to_excel(writer, sheet_name="Pass資料", index=False)
        sheet4_df.to_excel(writer, sheet_name="Pass重複資料", index=False)
        onlyfail_display.to_excel(writer, sheet_name="只有Fail的SN", index=False)

    n_data_cols = len(df.columns)
    usn_col_idx = list(df.columns).index("usn") + 1
    usn_col_letter = get_column_letter(usn_col_idx)
    tresult_col_idx = list(df.columns).index("tresult") + 1
    tresult_col_letter = get_column_letter(tresult_col_idx)
    cycletime_col_idx = list(df.columns).index("cycleTime") + 1
    cycletime_col_letter = get_column_letter(cycletime_col_idx)

    helper_usn_col_idx = n_data_cols + 1
    helper_isfirst_col_idx = n_data_cols + 2
    helper_count0_col_idx = n_data_cols + 3
    helper_count1_col_idx = n_data_cols + 4
    helper_match_col_idx = n_data_cols + 5

    helper_usn_letter = get_column_letter(helper_usn_col_idx)
    helper_isfirst_letter = get_column_letter(helper_isfirst_col_idx)
    helper_count0_letter = get_column_letter(helper_count0_col_idx)
    helper_count1_letter = get_column_letter(helper_count1_col_idx)
    helper_match_letter = get_column_letter(helper_match_col_idx)

    n_orig = len(df)
    n_sheet2 = len(sheet2_display)
    n_sheet3 = len(sheet3_display)
    n_sheet4 = len(sheet4_df)

    wb = openpyxl.load_workbook(out_path)
    ws2 = wb["Fail資料"]
    ws2.cell(row=1, column=helper_count0_col_idx, value="_count0")
    ws2.cell(row=1, column=helper_count1_col_idx, value="_count1")
    ws2.cell(row=1, column=helper_match_col_idx, value="_match_1and1")

    if n_sheet2 > 0:
        for i in range(2, 2 + n_sheet2):
            ws2.cell(row=i, column=helper_count0_col_idx, value=(
                f"=COUNTIFS({helper_usn_letter}2:{helper_usn_letter}{1+n_sheet2},{helper_usn_letter}{i},"
                f"{tresult_col_letter}2:{tresult_col_letter}{1+n_sheet2},\"0\")"
            ))
            ws2.cell(row=i, column=helper_count1_col_idx, value=(
                f"=COUNTIFS({helper_usn_letter}2:{helper_usn_letter}{1+n_sheet2},{helper_usn_letter}{i},"
                f"{tresult_col_letter}2:{tresult_col_letter}{1+n_sheet2},\"1\")"
            ))
            ws2.cell(row=i, column=helper_match_col_idx, value=(
                f"=IF(AND({helper_count0_letter}{i}>=1,{helper_count1_letter}{i}>=1),1,0)"
            ))

    for col_idx in [helper_usn_col_idx, helper_isfirst_col_idx, helper_count0_col_idx,
                    helper_count1_col_idx, helper_match_col_idx]:
        ws2.column_dimensions[get_column_letter(col_idx)].hidden = True

    ws5 = wb.create_sheet("統計")
    wb.save(out_path)

    ws5["A1"] = "項目"
    ws5["B1"] = "數量"
    r = 2
    ws5.cell(row=r, column=1, value="原始資料總筆數")
    ws5.cell(row=r, column=2, value=safe_counta("原始資料", tresult_col_letter, n_orig))
    r += 1

    ws5.cell(row=r, column=1, value="Fail總筆數")
    ws5.cell(row=r, column=2, value=safe_counta("Fail資料", tresult_col_letter, n_sheet2))
    r += 1

    ws5.cell(row=r, column=1, value="Pass總筆數")
    ws5.cell(row=r, column=2, value=safe_counta("Pass資料", tresult_col_letter, n_sheet3))
    r += 1

    ws5.cell(row=r, column=1, value="Pass重複筆數")
    ws5.cell(row=r, column=2, value=safe_counta("Pass重複資料", tresult_col_letter, n_sheet4))
    r += 1

    ws5.cell(row=r, column=1, value="只有Fail筆數")
    ws5.cell(row=r, column=2, value=only_zero_usn_count)
    r += 2

    ws5.cell(row=r, column=1, value="Base Input Quanity")
    n_total_distinct_row = r
    formula_fail_usn = safe_counta("Fail資料", usn_col_letter, n_sheet2)
    formula_pass_usn = safe_counta("Pass資料", usn_col_letter, n_sheet3)
    ws5.cell(row=r, column=2, value=f"={formula_fail_usn.lstrip('=')}+{formula_pass_usn.lstrip('=')}")
    r += 1

    ws5.cell(row=r, column=1, value="Base Output Quanity")
    ws5.cell(row=r, column=2, value=f"={formula_fail_usn.lstrip('=')}+{formula_pass_usn.lstrip('=')}-{only_zero_usn_count}")
    r += 1

    ws5.cell(row=r, column=1, value="Failure Quanity")
    n_zero_distinct_row = r
    ws5.cell(row=r, column=2, value=formula_fail_usn)
    r += 1

    ws5.cell(row=r, column=1, value="Pass Quanity")
    ws5.cell(row=r, column=2, value=formula_pass_usn)
    r += 1

    ws5.cell(row=r, column=1, value="複測確認無defect")
    n_ndf_row = r
    if n_sheet2 > 0:
        ws5.cell(row=r, column=2, value=(
            f"=SUMPRODUCT((Fail資料!{helper_isfirst_letter}2:{helper_isfirst_letter}{1+n_sheet2}=1)*"
            f"(Fail資料!{helper_match_letter}2:{helper_match_letter}{1+n_sheet2}=1))"
        ))
    else:
        ws5.cell(row=r, column=2, value=0)
    r += 2

    ws5.cell(row=r, column=1, value="FPY")
    row_fpy = r
    ws5.cell(row=r, column=2, value=f"=(B{n_total_distinct_row}-B{n_zero_distinct_row})/B{n_total_distinct_row}")
    ws5.cell(row=r, column=2).number_format = "0.00%"
    r += 1

    ws5.cell(row=r, column=1, value="VPY")
    row_vpy = r
    ws5.cell(row=r, column=2, value=f"=(B{n_total_distinct_row}-B{n_zero_distinct_row}+B{n_ndf_row})/B{n_total_distinct_row}")
    ws5.cell(row=r, column=2).number_format = "0.00%"
    r += 1

    ws5.cell(row=r, column=1, value="NDF Rate")
    ws5.cell(row=r, column=2, value=f"=B{row_vpy}-B{row_fpy}")
    ws5.cell(row=r, column=2).number_format = "0.00%"
    r += 2

    ct_range = f"原始資料!{cycletime_col_letter}2:{cycletime_col_letter}{1+n_orig}"

    ws5.cell(row=r, column=1, value="Cycle Time有效SN數(應等於Base Output Quantity)")
    ws5.cell(row=r, column=2, value=f"=COUNT({ct_range})")
    r += 1

    ws5.cell(row=r, column=1, value="Pure Test Average CT")
    ws5.cell(row=r, column=2, value=f"=AVERAGE({ct_range})")
    ws5.cell(row=r, column=2).number_format = "0.00"
    r += 1

    ws5.cell(row=r, column=1, value="Pure Test P50 CT")
    ws5.cell(row=r, column=2, value=f"=MEDIAN({ct_range})")
    ws5.cell(row=r, column=2).number_format = "0.00"
    r += 1

    ws5.cell(row=r, column=1, value="Pure Test P10 CT")
    row_p10 = r
    ws5.cell(row=r, column=2, value=f"=PERCENTILE({ct_range},0.1)")
    ws5.cell(row=r, column=2).number_format = "0.00"
    r += 1

    ws5.cell(row=r, column=1, value="Pure Test P90 CT")
    row_p90 = r
    ws5.cell(row=r, column=2, value=f"=PERCENTILE({ct_range},0.9)")
    ws5.cell(row=r, column=2).number_format = "0.00"
    r += 1

    ws5.cell(row=r, column=1, value="Pure Test CT Spread")
    ws5.cell(row=r, column=2, value=f"=B{row_p90}-B{row_p10}")
    ws5.cell(row=r, column=2).number_format = "0.00"
    r += 1

    wb.save(out_path)

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    body_font = Font(name="Arial")

    for ws in wb.worksheets:
        if ws.max_row < 1:
            continue
        max_col = ws.max_column
        max_row = ws.max_row
        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        for col in range(1, max_col + 1):
            letter = get_column_letter(col)
            max_len = len(str(ws.cell(row=1, column=col).value or ""))
            for row in range(2, max_row + 1):
                v = ws.cell(row=row, column=col).value
                ws.cell(row=row, column=col).font = body_font
                if v is not None:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[letter].width = min(max_len + 4, 30)
        ws.freeze_panes = "A2"

    wb.save(out_path)

    ws7 = wb.create_sheet("cycletime小於０的SN")
    ws7.cell(row=1, column=1, value="USN")
    ws7.cell(row=1, column=2, value="cycleTime")

    df_filtered = df[df['cycleTime'] <0].sort_values('cycleTime', ascending=False)

    row_num = 2
    for idx, row in df_filtered.iterrows():
        if (row['cycleTime'] < 0):
            print(f"行號: {idx}, USN: {row['usn']}, cycleTime: {row['cycleTime']}")
            ws7.cell(row=row_num, column=1, value=row['usn'])
            ws7.cell(row=row_num, column=2, value=row['cycleTime'])
            row_num += 1
    wb.save(out_path)


    return metrics


# ---------- 彙總表(跨 WYMY / WYMTN 一張表,格式參考你要的樣式) ----------

SUMMARY_HEADERS = [
    "Site", "Project", "Board", "Input Qty", "Output Qty", "Failure Qty",
    "FPY (%)", "VPY (%)", "NDF", "P10 (sec)", "P50 (sec)", "P90 (sec)", "Std (sec)",
]


def build_summary(results: dict, out_path: str):
    """
    results: { (site, model, board): metrics_dict }
    一個 (site, model, board) 一列。

    注意:
    - P10/P50/P90 仍是從資料現場算的百分位數,單位秒,跟個別 board 的「統計」分頁一致,不做任何換算。
    - Std 改成查 STD_LOOKUP[(site, model, board)],不再用 P90-P10 現場算。
      查不到就印提醒、格子留空,不會偷偷放錯值。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "彙總"

    for col, header in enumerate(SUMMARY_HEADERS, start=1):
        ws.cell(row=1, column=col, value=header)

    row = 2
    for (site, model, board), m in results.items():
        if m is None:
            continue
        ws.cell(row=row, column=1, value=site)
        ws.cell(row=row, column=2, value=model)
        ws.cell(row=row, column=3, value=board)
        ws.cell(row=row, column=4, value=m["base_input_qty"])
        ws.cell(row=row, column=5, value=m["base_output_qty"])
        ws.cell(row=row, column=6, value=m["failure_qty"])

        c = ws.cell(row=row, column=7, value=m["fpy"])
        c.number_format = "0.00%"
        c = ws.cell(row=row, column=8, value=m["vpy"])
        c.number_format = "0.00%"
        c = ws.cell(row=row, column=9, value=m["ndf_rate"])
        c.number_format = "0.00%"

        for col, key in [(10, "p10_sec"), (11, "p50_sec"), (12, "p90_sec")]:
            val = m[key]
            ws.cell(row=row, column=col, value=val).number_format = "0.00"

        std_val = STD_LOOKUP.get((site, model, board))
        if std_val is None:
            print(f"[提醒] STD_LOOKUP 沒有 {(site, model, board)} 這組的值,Std 欄位留空")
        ws.cell(row=row, column=13, value=std_val).number_format = "0.00"

        row += 1

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    body_font = Font(name="Arial")

    max_col = len(SUMMARY_HEADERS)
    max_row = ws.max_row
    for col in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        max_len = len(str(ws.cell(row=1, column=col).value or ""))
        for r in range(2, max_row + 1):
            v = ws.cell(row=r, column=col).value
            ws.cell(row=r, column=col).font = body_font
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="center")
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = min(max_len + 4, 20)
    ws.freeze_panes = "A2"

    wb.save(out_path)


def main():
    args = parse_args()

    os.makedirs(ONEDRIVE_DAILY_DIR, exist_ok=True)

    if args.start and args.end:
        start_dt = datetime.strptime(args.start, "%Y-%m-%d").replace(hour=8, minute=0, second=0)
        end_dt = datetime.strptime(args.end, "%Y-%m-%d").replace(hour=8, minute=0, second=0)
    else:
        start_dt, end_dt = compute_default_window()

    start_str = start_dt.strftime("%Y-%m-%d %H:%M:%S")
    end_str = end_dt.strftime("%Y-%m-%d %H:%M:%S")
    date_tag = start_dt.strftime("%Y%m%d")

    print(f"查詢窗口: {start_str}  ~  {end_str}")

    # --model/--teststage 兩個都給的話,只查這一組(方便單獨測試),--site 可選擇只查一邊;
    # 沒給 --model/--teststage 的話,改成把 JOBS 清單裡全部 (site, model, teststage) 組合都查一遍。
    if args.model and args.teststage:
        sites = [args.site] if args.site else list(TABLES.keys())
        jobs = [(site, args.model, args.teststage) for site in sites]
    else:
        jobs = JOBS
    print(f"本次要查的 (site, model, teststage) 組合: {jobs}")

    host, http_path, token = get_connection_params()

    print("連線 Databricks SQL Warehouse ...")
    connection = sql.connect(
        server_hostname=host,
        http_path=http_path,
        access_token=token,
    )

    results = {}

    try:
        for site, model, teststage in jobs:
            table_name = TABLES[site]
            print(f"=== site={site}  model={model}  testStage={teststage} ===")
            print(f"查詢 {table_name} ...")
            df = fetch_table(connection, table_name, model, teststage, start_str, end_str)
            print(f"  取得 {len(df)} 筆資料")

            if df.empty:
                print(f"  沒有資料,跳過")
                continue

            df["board"] = df.apply(
                lambda r: classify_board(site, model, teststage, r.get("upn"), r.get("fixtureId")),
                axis=1,
            )

            for board, group_df in df.groupby("board"):
                out_path = f"{ONEDRIVE_DAILY_DIR}/{site}_{date_tag}_{model}_{board}.xlsx"
                print(f"  [{board}] 跑 FPY 分析 -> {out_path}")
                metrics = run_fpy_analysis(group_df.copy(), out_path)
                results[(site, model, board)] = metrics
                print(f"  已存成: {out_path}")
    finally:
        connection.close()

    summary_path = f"{ONEDRIVE_DAILY_DIR}/{date_tag}.xlsx"
    print(f"產生彙總表 -> {summary_path}")
    build_summary(results, summary_path)
    print(f"已存成: {summary_path}")


if __name__ == "__main__":
    main()
