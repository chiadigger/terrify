
import os
import time
import requests
import pandas as pd

from pathlib import Path
from datetime import date, timedelta
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment

# =========================================================
# 1. 基本設定
# =========================================================

TOKEN = os.getenv("FINMIND_TOKEN", "").strip()

if not TOKEN:
    raise RuntimeError("找不到 FINMIND_TOKEN，請設定 GitHub Secrets")

API_URL = "https://api.finmindtrade.com/api/v4/data"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

SUMMARY_FILE = DATA_DIR / "market_margin_history.csv"
DETAIL_FILE = DATA_DIR / "stock_margin_history.csv"
EXCEL_FILE = OUTPUT_DIR / "TW_Margin_Monitor.xlsx"

# 初次執行的歷史起始日期
START_DATE = os.getenv("START_DATE", "2026-01-01")

# 每次重新檢查最近幾天
LOOKBACK_DAYS = 10

session = requests.Session()
session.headers.update({
    "Authorization": f"Bearer {TOKEN}"
})


# =========================================================
# 2. 下載指定日期全市場資料
# =========================================================

def fetch_market_data(target_date):

    params = {
        "dataset": "TaiwanStockMarginMaintenance",
        "start_date": target_date
    }

    for attempt in range(3):

        try:
            response = session.get(
                API_URL,
                params=params,
                timeout=90
            )

            response.raise_for_status()
            result = response.json()

            if result.get("status") != 200:
                raise RuntimeError(
                    result.get("msg", str(result))
                )

            df = pd.DataFrame(result.get("data", []))

            if not df.empty:
                df["date"] = df["date"].astype(str)
                df = df[df["date"] == target_date]

            return df

        except Exception as e:

            print(
                f"[ERROR] {target_date} "
                f"第 {attempt + 1} 次下載失敗：{e}"
            )

            time.sleep(5 * (attempt + 1))

    return pd.DataFrame()


# =========================================================
# 3. 計算全市場融資維持率
# =========================================================

def calculate_market_metrics(df, target_date):

    required_columns = [
        "date",
        "stock_id",
        "margin_balance",
        "margin_cost",
        "margin_ratio",
        "margin_maintenance"
    ]

    if df.empty:
        return None

    if not set(required_columns).issubset(df.columns):
        raise ValueError(
            f"API 欄位不完整：{df.columns.tolist()}"
        )

    df = df.copy()

    for col in required_columns[2:]:
        df[col] = pd.to_numeric(
            df[col], errors="coerce"
        )

    df = df.dropna(
        subset=required_columns[2:]
    )

    df = df[
        (df["margin_balance"] > 0) &
        (df["margin_cost"] > 0) &
        (df["margin_ratio"] > 0) &
        (df["margin_maintenance"] > 0)
    ].copy()

    if df.empty:
        return None

    # 融資餘額單位：張
    # 一張 = 1,000 股
    # 融資成數例如 0.6、0.5
    #
    # 推估融資金額 = 融資張數 × 1000
    #                × 融資成本 × 融資成數

    df["estimated_loan"] = (
        df["margin_balance"]
        * 1000
        * df["margin_cost"]
        * df["margin_ratio"]
    )

    total_loan = df["estimated_loan"].sum()

    if total_loan <= 0:
        return None

    # 以融資金額加權的市場估算維持率
    market_ratio = (
        df["estimated_loan"]
        * df["margin_maintenance"]
    ).sum() / total_loan

    # 低維持率股票
    under_130 = df[
        df["margin_maintenance"] < 130
    ]

    under_150 = df[
        df["margin_maintenance"] < 150
    ]

    under_166 = df[
        df["margin_maintenance"] < 166
    ]

    # 低維持率融資金額占比
    risk_130 = (
        under_130["estimated_loan"].sum()
        / total_loan * 100
    )

    risk_150 = (
        under_150["estimated_loan"].sum()
        / total_loan * 100
    )

    result = {
        "date": target_date,

        "market_maintenance": round(
            market_ratio, 2
        ),

        "estimated_loan_100m": round(
            total_loan / 100000000, 2
        ),

        "margin_balance_lots": int(
            df["margin_balance"].sum()
        ),

        "stock_count": int(len(df)),

        "under_130_count": int(len(under_130)),

        "under_150_count": int(len(under_150)),

        "under_166_count": int(len(under_166)),

        "under_130_loan_pct": round(
            risk_130, 2
        ),

        "under_150_loan_pct": round(
            risk_150, 2
        )
    }

    return result, df


# =========================================================
# 4. 讀取歷史資料
# =========================================================

def load_history(file_path):

    if file_path.exists():

        return pd.read_csv(
            file_path,
            dtype={"stock_id": str}
        )

    return pd.DataFrame()


summary_history = load_history(SUMMARY_FILE)
detail_history = load_history(DETAIL_FILE)

today = date.today()

if summary_history.empty:

    fetch_start = date.fromisoformat(START_DATE)

else:

    latest_date = pd.to_datetime(
        summary_history["date"]
    ).max().date()

    fetch_start = max(
        date.fromisoformat(START_DATE),
        latest_date - timedelta(days=LOOKBACK_DAYS)
    )


# =========================================================
# 5. 逐日更新
# =========================================================

all_dates = pd.date_range(
    start=fetch_start,
    end=today,
    freq="B"
)

new_summaries = []
new_details = []

for current_date in all_dates:

    target_date = current_date.strftime(
        "%Y-%m-%d"
    )

    print(f"\n[DOWNLOAD] {target_date}")

    df = fetch_market_data(target_date)

    if df.empty:
        print("[INFO] 無資料，可能是假日或尚未更新")
        continue

    calculated = calculate_market_metrics(
        df,
        target_date
    )

    if calculated is None:
        continue

    summary, detail = calculated

    new_summaries.append(summary)

    detail["date"] = target_date
    new_details.append(detail)

    print(
        f"[OK] 維持率："
        f"{summary['market_maintenance']}%"
    )

    time.sleep(0.3)


# =========================================================
# 6. 合併歷史資料，避免重複
# =========================================================

if new_summaries:

    summary_history = pd.concat(
        [
            summary_history,
            pd.DataFrame(new_summaries)
        ],
        ignore_index=True
    )

    summary_history["date"] = pd.to_datetime(
        summary_history["date"]
    ).dt.strftime("%Y-%m-%d")

    summary_history = (
        summary_history
        .drop_duplicates(
            subset=["date"],
            keep="last"
        )
        .sort_values("date")
        .reset_index(drop=True)
    )


if new_details:

    detail_history = pd.concat(
        [detail_history] + new_details,
        ignore_index=True
    )

    detail_history["date"] = pd.to_datetime(
        detail_history["date"]
    ).dt.strftime("%Y-%m-%d")

    detail_history = (
        detail_history
        .drop_duplicates(
            subset=["date", "stock_id"],
            keep="last"
        )
        .sort_values(
            ["date", "stock_id"]
        )
        .reset_index(drop=True)
    )


# =========================================================
# 7. 儲存 CSV
# =========================================================

if summary_history.empty:
    raise RuntimeError("目前沒有取得任何市場資料")

summary_history.to_csv(
    SUMMARY_FILE,
    index=False,
    encoding="utf-8-sig"
)

detail_history.to_csv(
    DETAIL_FILE,
    index=False,
    encoding="utf-8-sig"
)


# =========================================================
# 8. 產生 Excel
# =========================================================

with pd.ExcelWriter(
    EXCEL_FILE,
    engine="openpyxl"
) as writer:

    summary_history.to_excel(
        writer,
        sheet_name="市場維持率",
        index=False
    )

    detail_history.to_excel(
        writer,
        sheet_name="個股歷史",
        index=False
    )


# =========================================================
# 9. Excel 格式與圖表
# =========================================================

wb = load_workbook(EXCEL_FILE)

ws = wb["市場維持率"]

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

# 標題格式
for cell in ws[1]:

    cell.font = Font(
        bold=True,
        color="FFFFFF"
    )

    cell.fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    cell.alignment = Alignment(
        horizontal="center"
    )

# 欄寬
for column in ws.columns:

    max_length = max(
        len(str(cell.value or ""))
        for cell in column
    )

    ws.column_dimensions[
        column[0].column_letter
    ].width = min(max_length + 3, 30)


# 維持率歷史圖
chart = LineChart()

chart.title = "台股全市場融資維持率"
chart.style = 13
chart.y_axis.title = "維持率 (%)"
chart.x_axis.title = "日期"
chart.height = 10
chart.width = 22

ratio_col = 2

data = Reference(
    ws,
    min_col=ratio_col,
    min_row=1,
    max_row=ws.max_row
)

dates = Reference(
    ws,
    min_col=1,
    min_row=2,
    max_row=ws.max_row
)

chart.add_data(
    data,
    titles_from_data=True
)

chart.set_categories(dates)

ws.add_chart(
    chart,
    "L2"
)


# 個股資料表格式
detail_ws = wb["個股歷史"]

detail_ws.freeze_panes = "A2"
detail_ws.auto_filter.ref = detail_ws.dimensions

for cell in detail_ws[1]:

    cell.font = Font(
        bold=True,
        color="FFFFFF"
    )

    cell.fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

for column in detail_ws.columns:

    max_length = max(
        len(str(cell.value or ""))
        for cell in column
    )

    detail_ws.column_dimensions[
        column[0].column_letter
    ].width = min(max_length + 3, 30)


wb.save(EXCEL_FILE)


# =========================================================
# 10. 執行結果
# =========================================================

latest = summary_history.iloc[-1]

print("\n" + "=" * 50)
print("台股融資維持率更新完成")
print("=" * 50)

print("資料日期：", latest["date"])

print(
    "市場估算維持率：",
    latest["market_maintenance"], "%"
)

print(
    "融資金額估算：",
    latest["estimated_loan_100m"],
    "億元"
)

print(
    "低於130%股票數：",
    latest["under_130_count"]
)

print(
    "低於150%股票數：",
    latest["under_150_count"]
)

print(
    "低於130%融資權重：",
    latest["under_130_loan_pct"],
    "%"
)

print("\nExcel：", EXCEL_FILE)
