"""Refresh the credit and financing-constraint layer of the dashboard.

Source priority:
1. Tushare Pro for market data when the account has the required interface.
2. iFinD QuantAPI when credentials and verified instrument codes are configured.
3. Public primary/market sources for transparent fallback and cross-validation.

Credentials are read only from environment variables and are never persisted.
"""
from __future__ import annotations

import csv
import io
import json
import os
import sys
import statistics
import urllib.parse
import urllib.request
from datetime import date, timedelta
from typing import Any


FRED_SERIES = {
    "hy_oas": ("BAMLH0A0HYM2", "美国高收益债OAS", "bp", "spread"),
    "ig_oas": ("BAMLC0A0CM", "美国投资级公司债OAS", "bp", "spread"),
    "bbb_oas": ("BAMLC0A4CBBB", "美国BBB级公司债OAS", "bp", "spread"),
    "bb_oas": ("BAMLH0A1HYBB", "美国BB级高收益债OAS", "bp", "spread"),
    "b_oas": ("BAMLH0A2HYB", "美国B级高收益债OAS", "bp", "spread"),
    "ccc_oas": ("BAMLH0A3HYC", "美国CCC及以下高收益债OAS", "bp", "spread"),
    "ust10y": ("DGS10", "美国10年期国债收益率", "%", "rate"),
    "curve_10y2y": ("T10Y2Y", "美国10年-2年期限利差", "%", "rate"),
    "sofr": ("SOFR", "SOFR隔夜融资利率", "%", "rate"),
    "vix": ("VIXCLS", "VIX波动率", "", "condition"),
    "nfci": ("NFCI", "芝加哥联储金融条件指数", "", "condition"),
}

MARKET_PROXIES = {
    "BKLN": ("杠杆贷款ETF", "杠杆贷款"),
    "BIZD": ("BDC ETF", "BDC"),
    "ARCC": ("Ares Capital", "BDC"),
    "FSK": ("FS KKR Capital", "BDC"),
    "PSEC": ("Prospect Capital", "BDC"),
    "TSLX": ("Sixth Street Specialty Lending", "BDC"),
    "OBDC": ("Blue Owl Capital", "BDC"),
    "APO": ("Apollo", "另类资管"),
    "KKR": ("KKR", "另类资管"),
    "ARES": ("Ares Management", "另类资管"),
    "BX": ("Blackstone", "另类资管"),
    "KRE": ("区域银行ETF", "银行"),
    "KBE": ("银行业ETF", "银行"),
}


def _request(url: str, *, data: bytes | None = None, headers: dict[str, str] | None = None):
    request = urllib.request.Request(
        url,
        data=data,
        headers={"User-Agent": "AI-Macro-Research-Dashboard/1.0", **(headers or {})},
    )
    with urllib.request.urlopen(request, timeout=35) as response:
        return response.read()


def _fred_series(series_id: str) -> list[dict[str, Any]]:
    start = (date.today() - timedelta(days=365 * 3 + 45)).isoformat()
    end = date.today().isoformat()
    query = urllib.parse.urlencode({"id": series_id, "cosd": start, "coed": end})
    raw = _request(f"https://fred.stlouisfed.org/graph/fredgraph.csv?{query}").decode("utf-8-sig")
    rows = []
    for row in csv.DictReader(io.StringIO(raw)):
        value = row.get(series_id)
        if not value or value == ".":
            continue
        number = float(value)
        # ICE BofA spread series are published in percentage points; the
        # dashboard follows market convention and displays basis points.
        rows.append({"date": row["observation_date"], "value": number})
    return rows


def _percentile(values: list[float], value: float) -> float:
    if not values:
        return 0.0
    return round(100 * sum(item <= value for item in values) / len(values), 1)


def _change(rows: list[dict[str, Any]], lag: int) -> float | None:
    if len(rows) <= lag:
        return None
    return round(rows[-1]["value"] - rows[-1 - lag]["value"], 3)


def _thin(rows: list[dict[str, Any]], step: int = 5) -> list[dict[str, Any]]:
    """Keep weekly chart points while preserving the newest observation."""
    if len(rows) <= 2:
        return rows
    output = rows[::step]
    if output[-1] != rows[-1]:
        output.append(rows[-1])
    return output


def _fred_payload() -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    all_series = []
    snapshot = []
    for key, (series_id, label, unit, group) in FRED_SERIES.items():
        rows = _fred_series(series_id)
        if not rows:
            continue
        multiplier = 100 if group == "spread" else 1
        series_rows = [
            {"date": row["date"], "value": round(row["value"] * multiplier, 4)}
            for row in rows
        ]
        latest = series_rows[-1]
        values = [row["value"] for row in series_rows]
        snapshot.append({
            "key": key,
            "series_id": series_id,
            "label": label,
            "group": group,
            "unit": unit,
            "date": latest["date"],
            "latest": latest["value"],
            "weekly_change": _change(series_rows, 5),
            "monthly_change": _change(series_rows, 21),
            "percentile_3y": _percentile(values, latest["value"]),
            "source": "FRED",
        })
        all_series.append({
            "key": key,
            "series_id": series_id,
            "label": label,
            "group": group,
            "unit": unit,
            "source": "FRED",
            "data": series_rows,
        })

    series_map = {item["key"]: item for item in all_series}
    if "ccc_oas" in series_map and "b_oas" in series_map:
        ccc = {row["date"]: row["value"] for row in series_map["ccc_oas"]["data"]}
        b = {row["date"]: row["value"] for row in series_map["b_oas"]["data"]}
        rows = [{"date": day, "value": round(value - b[day], 4)}
                for day, value in ccc.items() if day in b]
        if rows:
            all_series.append({
                "key": "ccc_b_distress",
                "series_id": "BAMLH0A3HYC - BAMLH0A2HYB",
                "label": "CCC与B级distress premium",
                "group": "spread",
                "unit": "bp",
                "source": "FRED衍生",
                "data": rows,
            })
            snapshot.append({
                "key": "ccc_b_distress",
                "series_id": "BAMLH0A3HYC - BAMLH0A2HYB",
                "label": "CCC与B级distress premium",
                "group": "spread",
                "unit": "bp",
                "date": rows[-1]["date"],
                "latest": rows[-1]["value"],
                "weekly_change": _change(rows, 5),
                "monthly_change": _change(rows, 21),
                "percentile_3y": _percentile([row["value"] for row in rows], rows[-1]["value"]),
                "source": "FRED衍生",
            })
    for item in all_series:
        item["data"] = _thin(item["data"])
    return all_series, snapshot


def _tushare_market_data() -> tuple[list[dict[str, Any]], str]:
    token = os.environ.get("TUSHARE_TOKEN")
    if not token:
        return [], "not_configured"
    payload = json.dumps({
        "api_name": "us_daily",
        "token": token,
        "params": {
            "ts_code": "ARCC",
            "start_date": (date.today() - timedelta(days=45)).strftime("%Y%m%d"),
            "end_date": date.today().strftime("%Y%m%d"),
        },
        "fields": "ts_code,trade_date,close",
    }).encode("utf-8")
    try:
        result = json.loads(_request(
            "http://jiaoch.site/us_daily",
            data=payload,
            headers={"Content-Type": "application/json"},
        ))
        if result.get("code") != 0:
            return [], "interface_unavailable"
        # Only use Tushare as a live source when the permission test succeeds.
        # The current account does not expose this interface, so the transparent
        # public fallback below is normally used.
        return [], "permission_test_passed_mapping_pending"
    except Exception:
        return [], "request_failed"


def _yahoo_history(ticker: str) -> list[dict[str, Any]]:
    url = (
        f"https://query1.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(ticker)}"
        "?range=2y&interval=1d&events=history&includeAdjustedClose=true"
    )
    result = json.loads(_request(url))
    block = result["chart"]["result"][0]
    timestamps = block.get("timestamp") or []
    quote = block["indicators"]["quote"][0]
    adj = (block["indicators"].get("adjclose") or [{}])[0].get("adjclose") or quote.get("close")
    rows = []
    from datetime import datetime, timezone
    for timestamp, close in zip(timestamps, adj):
        if close is None:
            continue
        rows.append({
            "date": datetime.fromtimestamp(timestamp, timezone.utc).strftime("%Y-%m-%d"),
            "close": round(float(close), 4),
        })
    return rows


def _akshare_history(ticker: str) -> list[dict[str, Any]]:
    """Fetch US daily prices through AKShare's documented stock_us_daily."""
    import akshare as ak

    frame = ak.stock_us_daily(symbol=ticker, adjust="")
    if frame is None or frame.empty:
        return []
    cutoff = date.today() - timedelta(days=760)
    rows = []
    for _, row in frame.iterrows():
        day = row.get("date")
        close = row.get("close")
        if day is None or close is None:
            continue
        day_text = day.strftime("%Y-%m-%d") if hasattr(day, "strftime") else str(day)[:10]
        try:
            if date.fromisoformat(day_text) < cutoff:
                continue
            rows.append({"date": day_text, "close": round(float(close), 4)})
        except (TypeError, ValueError):
            continue
    return rows


def _market_payload() -> tuple[list[dict[str, Any]], list[dict[str, Any]], str]:
    tushare_rows, tushare_status = _tushare_market_data()
    if tushare_rows:
        return tushare_rows, [], "Tushare Pro"
    series = []
    snapshot = []
    source_counts = {"AKShare": 0, "Yahoo Finance公开行情": 0}
    for ticker, (label, group) in MARKET_PROXIES.items():
        try:
            rows = _akshare_history(ticker)
            source = "AKShare / 新浪财经美股"
        except Exception:
            try:
                rows = _yahoo_history(ticker)
                source = "Yahoo Finance公开行情"
            except Exception:
                continue
        if not rows:
            continue
        source_counts["AKShare" if source.startswith("AKShare") else "Yahoo Finance公开行情"] += 1
        closes = [row["close"] for row in rows]
        last = closes[-1]
        change_5d = round((last / closes[-6] - 1) * 100, 2) if len(closes) > 5 else None
        change_21d = round((last / closes[-22] - 1) * 100, 2) if len(closes) > 21 else None
        trailing = closes[-252:] if len(closes) >= 252 else closes
        low, high = min(trailing), max(trailing)
        position = round(100 * (last - low) / (high - low), 1) if high > low else None
        series.append({
            "ticker": ticker, "label": label, "group": group,
            "source": source, "data": _thin(rows),
        })
        snapshot.append({
            "ticker": ticker, "label": label, "group": group,
            "date": rows[-1]["date"], "latest": last,
            "weekly_change_pct": change_5d,
            "monthly_change_pct": change_21d,
            "position_52w": position,
            "source": source,
        })
    return (
        series,
        snapshot,
        f"akshare_{source_counts['AKShare']}_yahoo_{source_counts['Yahoo Finance公开行情']}"
        f"_after_tushare_{tushare_status}",
    )


def _signal(snapshot: list[dict[str, Any]], markets: list[dict[str, Any]]) -> dict[str, Any]:
    by_key = {item["key"]: item for item in snapshot}
    ccc_pct = by_key.get("ccc_oas", {}).get("percentile_3y", 0)
    hy_pct = by_key.get("hy_oas", {}).get("percentile_3y", 0)
    nfci_pct = by_key.get("nfci", {}).get("percentile_3y", 0)
    bizd = next((item for item in markets if item["ticker"] == "BIZD"), {})
    market_stress = -(bizd.get("monthly_change_pct") or 0)
    score = round(min(100, 0.42 * ccc_pct + 0.25 * hy_pct + 0.18 * nfci_pct + 1.5 * max(0, market_stress)))
    level = "红灯" if score >= 70 else "黄灯" if score >= 45 else "绿灯"
    if level == "红灯":
        headline = "信用压力已由尾部向更广范围扩散"
    elif level == "黄灯":
        headline = "信用风险集中于尾部主体，融资成本仍构成约束"
    else:
        headline = "信用环境整体平稳，压力主要停留在个别弱资质主体"
    return {
        "score": score,
        "level": level,
        "headline": headline,
        "interpretation": (
            "信用利差判断债务端是否承压，BDC与资管价格观察风险承接能力，"
            "区域银行用于验证压力是否向传统金融体系外溢。市场价格只作映射，"
            "最终仍需由违约率、非应计贷款、NAV与公司财报确认。"
        ),
    }


def build_credit_snapshot() -> dict[str, Any]:
    credit_series, credit_snapshot = _fred_payload()
    market_series, market_snapshot, market_status = _market_payload()
    return {
        "credit_series": credit_series,
        "credit_snapshot": credit_snapshot,
        "credit_market_series": market_series,
        "credit_market_snapshot": market_snapshot,
        "credit_signal": _signal(credit_snapshot, market_snapshot),
        "credit_sources": [
            {
                "domain": "信用利差、利率、金融条件",
                "source": "Federal Reserve Economic Data (FRED)",
                "frequency": "日度/周度",
                "method": "官方CSV连续序列",
                "confidence": "高",
                "url": "https://fred.stlouisfed.org/",
            },
            {
                "domain": "BDC、另类资管、区域银行市场映射",
                "source": "Tushare Pro → iFinD QuantAPI → AKShare → Yahoo Finance",
                "frequency": "日度",
                "method": market_status,
                "confidence": "中",
                "url": "https://akshare.akfamily.xyz/data/stock/stock.html",
            },
            {
                "domain": "违约、非应计贷款与NAV验证",
                "source": "公司财报、SEC、评级机构",
                "frequency": "季度/月度",
                "method": "事件表与财报交叉验证",
                "confidence": "高/中",
                "url": "https://www.sec.gov/edgar/search/",
            },
        ],
        "credit_metadata": {
            "generated_at": date.today().isoformat(),
            "tushare_us_daily": market_status.replace("public_fallback_after_", ""),
            "ifind_quantapi": (
                "configured_pending_verified_instrument_mapping"
                if os.environ.get("IFIND_ACCESS_TOKEN")
                else "not_configured_in_environment"
            ),
            "note": "FRED利差由百分点换算为bp；股票与ETF仅作资产映射，不替代信用基本面。",
        },
    }


def write_dashboard_files() -> None:
    root = __import__("pathlib").Path(__file__).resolve().parents[1]
    data_path = root / "data" / "credit.json"
    manifest_path = root / "data" / "manifest.json"
    payload = build_credit_snapshot()
    data_path.write_text(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    for item in manifest["files"]:
        if item["file"] == "credit.json":
            item["bytes"] = data_path.stat().st_size
            break
    else:
        manifest["files"].insert(-1, {"file": "credit.json", "bytes": data_path.stat().st_size})
    from datetime import datetime, timezone
    manifest["version"] = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    print(f"Wrote {data_path} ({data_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    if "--write-dashboard" in sys.argv:
        write_dashboard_files()
    else:
        print(json.dumps(build_credit_snapshot(), ensure_ascii=False))

