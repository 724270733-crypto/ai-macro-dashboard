"""Build the GitHub Pages data snapshot from the research workbooks.

Usage (Windows):
  set AI_FRAMEWORK_DIR=<your local research data folder>
  set TUSHARE_TOKEN=...
  python scripts/refresh_dashboard.py

The script never writes to the original workbooks.  It exports a compact JSON
snapshot used by the static dashboard.  iFinD export files can be placed in
data/imports/ and will be added in a later pass; the existing iFinD Excel
formulas are deliberately not treated as values when Excel has not refreshed
them.
"""
from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "dashboard-data.json"
SOURCE_VALUE = os.environ.get("AI_FRAMEWORK_DIR")
if not SOURCE_VALUE:
    raise RuntimeError("Set AI_FRAMEWORK_DIR to the folder containing the source workbooks.")
SOURCE = Path(SOURCE_VALUE)


def clean(v):
    if pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d")
    if hasattr(v, "item"):
        return v.item()
    return v


def records(df: pd.DataFrame):
    return [{str(k): clean(v) for k, v in row.items()} for row in df.to_dict("records")]


def find_workbook(name: str) -> Path:
    path = SOURCE / name
    if not path.exists():
        raise FileNotFoundError(f"Missing source workbook: {path}")
    return path


def workbook_data():
    db = find_workbook("AI宏观数据库_指标与获取路径.xlsx")
    token = pd.read_excel(db, sheet_name="top50模型token数量2")
    token.iloc[:, 0] = pd.to_datetime(token.iloc[:, 0])
    providers = [x for x in ["Claude", "DeepSeek", "Gemini", "Grok", "OpenAI", "Qwen"] if x in token.columns]
    token["total_tokens"] = token[providers].fillna(0).sum(axis=1)
    token = token.rename(columns={token.columns[0]: "date"})
    token["date"] = token["date"].dt.strftime("%Y-%m-%d")
    keep = ["date", "total_tokens"] + providers
    token = token[keep]

    apps = pd.read_excel(db, sheet_name="按代币使用率排名前列应用")
    apps = apps.iloc[:, :5].dropna(subset=[apps.columns[0]])
    apps.columns = ["rank", "app_id", "app_name", "total_tokens", "total_requests"]
    apps["tokens_per_request"] = apps["total_tokens"] / apps["total_requests"]
    apps = apps.head(12)

    changes = pd.read_excel(db, sheet_name="模型价格变动跟踪")
    changes = changes.iloc[:, :12]
    changes.columns = ["id", "date", "model_id", "model", "provider", "field", "field_label", "change_type", "old", "new", "delta", "pct"]
    changes["date"] = pd.to_datetime(changes["date"]).dt.strftime("%Y-%m-%d")
    events = changes.groupby("date").size().reset_index(name="events").sort_values("date")
    provider_events = changes.groupby("provider").size().reset_index(name="events").sort_values("events", ascending=False).head(8)

    state = find_workbook("state_of_ai.xlsx")
    funding = pd.read_excel(state, sheet_name="Equity Funding", header=None)
    quarter = funding.iloc[17:32, 1:5].copy()
    quarter.columns = ["year", "quarter", "deal_count", "funding_usd_m"]
    quarter = quarter.dropna(subset=["year", "quarter"])
    quarter["period"] = quarter.apply(lambda r: f"{int(r.year)} Q{int(r.quarter)}", axis=1)
    return {
        "token": records(token),
        "apps": records(apps),
        "price_events": records(events),
        "price_events_by_provider": records(provider_events),
        "funding": records(quarter[["period", "deal_count", "funding_usd_m"]]),
    }


def tushare_financials():
    """Optional live Tushare refresh for A-share AI hardware financial evidence."""
    token = os.environ.get("TUSHARE_TOKEN")
    if not token:
        return [], "Tushare token not configured; retained last static snapshot only."
    try:
        import tushare as ts
        pro = ts.pro_api(token)
        pro._DataApi__token = token
        pro._DataApi__http_url = "http://jiaoch.site"
        companies = {
            "688256.SH": "寒武纪",
            "688041.SH": "海光信息",
            "300308.SZ": "中际旭创",
        }
        output = []
        for code, name in companies.items():
            inc = pro.income(ts_code=code, start_date="20240101", end_date="20271231", fields="ts_code,ann_date,end_date,report_type,revenue,operate_profit,n_income")
            cf = pro.cashflow(ts_code=code, start_date="20240101", end_date="20271231", fields="ts_code,ann_date,end_date,report_type,n_cashflow_act,c_pay_acq_const_fiolta")
            if inc.empty:
                continue
            inc = inc.drop_duplicates(subset=["end_date"]).copy()
            cf = cf.drop_duplicates(subset=["end_date"]).copy() if not cf.empty else cf
            merged = inc.merge(cf[["end_date", "n_cashflow_act", "c_pay_acq_const_fiolta"]], on="end_date", how="left")
            merged = merged.sort_values("end_date")
            for _, row in merged.iterrows():
                output.append({
                    "code": code, "company": name, "period": str(row["end_date"]),
                    "revenue_cny_bn": round(float(row["revenue"]) / 1e9, 3),
                    "operating_profit_cny_bn": round(float(row["operate_profit"]) / 1e9, 3),
                    "net_income_cny_bn": round(float(row["n_income"]) / 1e9, 3),
                    "operating_cashflow_cny_bn": round(float(row["n_cashflow_act"]) / 1e9, 3) if pd.notna(row.get("n_cashflow_act")) else None,
                    "capex_cash_cny_bn": round(float(row["c_pay_acq_const_fiolta"]) / 1e9, 3) if pd.notna(row.get("c_pay_acq_const_fiolta")) else None,
                })
        return output, "Tushare Pro"
    except Exception as exc:
        return [], f"Tushare refresh failed: {type(exc).__name__}"


def manual_evidence():
    # Event-led / company-disclosed items are intentionally tagged rather than
    # converted into artificial high-frequency time series.
    return {
        "cloud_capex": [
            {"period": "2023", "usd_bn": 149.1}, {"period": "2024", "usd_bn": 224.2},
            {"period": "2025", "usd_bn": 387.1}, {"period": "2026 Q1", "usd_bn": 126.8},
        ],
        "hardware": [
            {"company": "NVIDIA", "metric": "Data Center revenue", "period": "FY2027 Q1", "value": 75.2, "unit": "USD bn", "yoy": 92, "source_type": "company disclosure"},
            {"company": "AMD", "metric": "Data Center revenue", "period": "2026 Q1", "value": 5.8, "unit": "USD bn", "yoy": 57, "source_type": "company disclosure"},
            {"company": "Broadcom", "metric": "AI semiconductor revenue", "period": "FY2026 Q2", "value": 10.8, "unit": "USD bn", "yoy": 143, "source_type": "company disclosure"},
            {"company": "Marvell", "metric": "Data Center revenue", "period": "FY2027 Q1", "value": 1.83, "unit": "USD bn", "yoy": 27, "source_type": "company disclosure"},
        ],
        "commercialization_events": [
            {"date": "2025-12", "company": "Anthropic", "metric": "ARR", "value": 9.0, "unit": "USD bn", "confidence": "medium", "source_type": "event / media"},
            {"date": "2026-05", "company": "Anthropic", "metric": "ARR", "value": 44.0, "unit": "USD bn", "confidence": "medium", "source_type": "event / third party"},
            {"date": "2026-02", "company": "Meta", "metric": "monthly token estimate", "value": 70.0, "unit": "trillion tokens", "confidence": "medium", "source_type": "third-party estimate"},
        ],
    }


def main():
    data = workbook_data()
    financials, financial_source = tushare_financials()
    data.update(manual_evidence())
    data["a_share_financials"] = financials
    data["metadata"] = {
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source_snapshot": "AI宏观数据库_指标与获取路径.xlsx; state_of_ai.xlsx",
        "financial_refresh": financial_source,
        "notes": [
            "OpenRouter series is a public sample proxy, not the whole AI market.",
            "Private-company ARR is maintained as an event ledger; do not calculate strict month-on-month growth.",
            "The iFinD capex sheet stores formulas. Only refreshed numeric exports should be imported as time series.",
        ],
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()

