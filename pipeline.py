"""04 API → Excel 数据管道:拉取 OKX 公开行情快照 → 清洗 → 带样式的 Excel + CSV。

只调用公开、免登录的只读接口 /api/v5/market/tickers,不需要 API Key。
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DEFAULT_BASE_URL = "https://www.okx.com"
TICKERS_PATH = "/api/v5/market/tickers"
CN_TZ = timezone(timedelta(hours=8))
NUMERIC_COLS = ["last", "open24h", "high24h", "low24h", "bidPx", "askPx", "vol24h", "volCcy24h"]

log = logging.getLogger("pipeline")


class ApiError(RuntimeError):
    """接口返回错误或重试耗尽。"""


# ---------------------------------------------------------------- 拉取
def fetch_json(url: str, params: dict | None = None, *, timeout: float = 10.0, retries: int = 3,
               backoff: float = 1.0, session: requests.Session | None = None, sleep=time.sleep) -> dict:
    """GET 请求,带超时;对超时/连接错误/429/5xx 做指数退避重试,4xx 直接报错。"""
    sess = session or requests.Session()
    last_err: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            resp = sess.get(url, params=params, timeout=timeout)
            if resp.status_code == 429 or resp.status_code >= 500:
                raise requests.HTTPError(f"HTTP {resp.status_code}", response=resp)
            if resp.status_code >= 400:
                raise ApiError(f"HTTP {resp.status_code}: {resp.text[:200]}")
            return resp.json()
        except (requests.Timeout, requests.ConnectionError, requests.HTTPError) as e:
            last_err = e
            log.warning("第 %d/%d 次请求失败:%s", attempt, retries, e)
            if attempt < retries:
                sleep(backoff * 2 ** (attempt - 1))
    raise ApiError(f"重试 {retries} 次后仍失败:{last_err}")


def fetch_tickers(inst_type: str = "SPOT", base_url: str = DEFAULT_BASE_URL, **kw) -> list[dict]:
    payload = fetch_json(base_url.rstrip("/") + TICKERS_PATH, {"instType": inst_type}, **kw)
    if str(payload.get("code")) != "0":
        raise ApiError(f"接口错误 code={payload.get('code')} msg={payload.get('msg')}")
    return payload.get("data", [])


# ---------------------------------------------------------------- 清洗
def clean_tickers(rows: list[dict], quote: str | None = "USDT", min_volume: float = 0.0,
                  top: int | None = 30) -> pd.DataFrame:
    """转数值、去重、过滤计价币和成交额、计算涨跌幅/价差/振幅,按成交额降序取前 N。"""
    if not rows:
        return pd.DataFrame(columns=["instId", "base", "quote", "last", "change_24h_pct", "amplitude_pct",
                                     "spread_pct", "volume_quote_24h", "time_cn"])
    df = pd.DataFrame(rows)
    df = df.drop_duplicates(subset="instId", keep="last")
    for c in NUMERIC_COLS:
        df[c] = pd.to_numeric(df.get(c), errors="coerce")
    df = df.dropna(subset=["last", "open24h"])
    df = df[(df["last"] > 0) & (df["open24h"] > 0)]
    parts = df["instId"].str.split("-", n=1, expand=True)
    df["base"], df["quote"] = parts[0], parts[1]
    if quote:
        df = df[df["quote"].str.upper() == quote.upper()]
    # SPOT 的 volCcy24h 是计价币成交额
    df["volume_quote_24h"] = df["volCcy24h"].fillna(0.0)
    df = df[df["volume_quote_24h"] >= min_volume]
    df["change_24h_pct"] = (df["last"] / df["open24h"] - 1) * 100
    df["amplitude_pct"] = (df["high24h"] - df["low24h"]) / df["open24h"] * 100
    mid = (df["bidPx"] + df["askPx"]) / 2
    df["spread_pct"] = ((df["askPx"] - df["bidPx"]) / mid * 100).where(mid > 0)
    ts = pd.to_numeric(df.get("ts"), errors="coerce")
    df["time_cn"] = pd.to_datetime(ts, unit="ms", utc=True).dt.tz_convert(CN_TZ).dt.strftime("%Y-%m-%d %H:%M:%S")
    df = df.sort_values("volume_quote_24h", ascending=False)
    if top:
        df = df.head(top)
    out = df[["instId", "base", "quote", "last", "change_24h_pct", "amplitude_pct", "spread_pct",
              "volume_quote_24h", "time_cn"]].reset_index(drop=True)
    for c in ["change_24h_pct", "amplitude_pct"]:
        out[c] = out[c].round(2)
    out["spread_pct"] = out["spread_pct"].round(4)
    out["volume_quote_24h"] = out["volume_quote_24h"].round(0)
    return out


def summarize(df: pd.DataFrame) -> dict:
    if df.empty:
        return {"count": 0}
    best = df.loc[df["change_24h_pct"].idxmax()]
    worst = df.loc[df["change_24h_pct"].idxmin()]
    return {
        "count": int(len(df)),
        "up_count": int((df["change_24h_pct"] > 0).sum()),
        "down_count": int((df["change_24h_pct"] < 0).sum()),
        "median_change_pct": round(float(df["change_24h_pct"].median()), 2),
        "total_volume_quote": float(df["volume_quote_24h"].sum()),
        "top_gainer": f"{best['instId']} ({best['change_24h_pct']:+.2f}%)",
        "top_loser": f"{worst['instId']} ({worst['change_24h_pct']:+.2f}%)",
    }


# ---------------------------------------------------------------- 输出
HEADERS_CN = {
    "instId": "交易对", "base": "币种", "quote": "计价", "last": "最新价", "change_24h_pct": "24h涨跌%",
    "amplitude_pct": "24h振幅%", "spread_pct": "买卖价差%", "volume_quote_24h": "24h成交额(计价币)",
    "time_cn": "数据时间(北京时间)",
}


def write_excel(df: pd.DataFrame, summary: dict, path: Path, title: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "行情快照"
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ncol = len(df.columns)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    head_fill = PatternFill("solid", fgColor="1F4E78")
    for j, col in enumerate(df.columns, 1):
        c = ws.cell(row=2, column=j, value=HEADERS_CN.get(col, col))
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    fmt = {"last": "#,##0.########", "change_24h_pct": "0.00", "amplitude_pct": "0.00",
           "spread_pct": "0.0000", "volume_quote_24h": "#,##0"}
    for i, row in enumerate(df.itertuples(index=False), 3):
        for j, (col, val) in enumerate(zip(df.columns, row), 1):
            c = ws.cell(row=i, column=j, value=None if pd.isna(val) else val)
            c.border = border
            if col in fmt:
                c.number_format = fmt[col]
    last_row = len(df) + 2
    if len(df):
        idx = list(df.columns).index("change_24h_pct") + 1
        rng = f"{get_column_letter(idx)}3:{get_column_letter(idx)}{last_row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"],
                                                      font=Font(color="008000")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"],
                                                      font=Font(color="C00000")))
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ncol)}{last_row}"
    for j, col in enumerate(df.columns, 1):
        width = max([len(str(HEADERS_CN.get(col, col))) * 2] + [len(str(v)) for v in df[col].tolist()] + [8])
        ws.column_dimensions[get_column_letter(j)].width = min(width + 2, 28)

    ws2 = wb.create_sheet("汇总")
    labels = {"count": "交易对数量", "up_count": "上涨数量", "down_count": "下跌数量",
              "median_change_pct": "涨跌幅中位数%", "total_volume_quote": "成交额合计",
              "top_gainer": "涨幅最大", "top_loser": "跌幅最大"}
    ws2.append(["指标", "数值"])
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
    for k, v in summary.items():
        ws2.append([labels.get(k, k), v])
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 30
    wb.save(path)
    return path


# ---------------------------------------------------------------- 通知(占位,不真实发送)
def build_notification(summary: dict, title: str) -> str:
    if not summary.get("count"):
        return f"{title}:本次没有符合条件的数据。"
    return (f"{title}\n共 {summary['count']} 个交易对,上涨 {summary['up_count']} / 下跌 {summary['down_count']},"
            f"涨跌幅中位数 {summary['median_change_pct']}%\n"
            f"涨幅最大:{summary['top_gainer']};跌幅最大:{summary['top_loser']}")


def notify(message: str, config_path: Path | None, out_dir: Path) -> str:
    """可配置的通知占位。

    本作品**不会真实发送任何消息**:无论配置成哪个渠道,都只把将要发送的内容写到
    output/notify_preview.txt。接入真实渠道(企业微信/飞书/钉钉 webhook、邮件)时,
    在 `send_real()` 里实现,并由使用者自行提供并保管 webhook 地址或密钥。
    """
    cfg = {"enabled": False, "channel": "none"}
    if config_path and config_path.exists():
        cfg.update(json.loads(config_path.read_text(encoding="utf-8")))
    out_dir.mkdir(parents=True, exist_ok=True)
    preview = out_dir / "notify_preview.txt"
    preview.write_text(f"[channel={cfg.get('channel')} enabled={cfg.get('enabled')} 仅预览,未发送]\n{message}\n",
                       encoding="utf-8")
    return "preview_only"


# ---------------------------------------------------------------- 命令行
def parse_args(argv=None):
    p = argparse.ArgumentParser(description="拉取 OKX 公开行情快照 → 清洗 → 带样式 Excel/CSV")
    p.add_argument("--inst-type", default="SPOT", help="SPOT / SWAP / FUTURES,默认 SPOT")
    p.add_argument("--quote", default="USDT", help="只保留该计价币,传空字符串表示不过滤")
    p.add_argument("--top", type=int, default=30, help="按成交额取前 N 个,0 表示全部")
    p.add_argument("--min-volume", type=float, default=0.0, help="24h 成交额下限(计价币)")
    p.add_argument("--out", default="output", help="输出目录")
    p.add_argument("--timeout", type=float, default=10.0)
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--base-url", default=DEFAULT_BASE_URL)
    p.add_argument("--offline", help="读取本地 JSON(接口原始返回格式),不联网")
    p.add_argument("--notify-config", help="通知配置 JSON(仅生成预览,不真实发送)")
    return p.parse_args(argv)


def run(argv=None) -> dict:
    args = parse_args(argv)
    out_dir = Path(args.out)
    if args.offline:
        payload = json.loads(Path(args.offline).read_text(encoding="utf-8"))
        rows = payload.get("data", [])
    else:
        rows = fetch_tickers(args.inst_type, args.base_url, timeout=args.timeout, retries=args.retries)
    df = clean_tickers(rows, quote=args.quote or None, min_volume=args.min_volume, top=args.top or None)
    summary = summarize(df)
    stamp = datetime.now(CN_TZ).strftime("%Y%m%d")
    name = f"okx_{args.inst_type.lower()}_{(args.quote or 'all').lower()}_{stamp}"
    title = f"OKX {args.inst_type} {args.quote or '全部'} 行情快照 Top{args.top or len(df)}(生成于 {datetime.now(CN_TZ):%Y-%m-%d %H:%M} 北京时间)"
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / f"{name}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    xlsx_path = write_excel(df, summary, out_dir / f"{name}.xlsx", title)
    message = build_notification(summary, title)
    notify(message, Path(args.notify_config) if args.notify_config else None, out_dir)
    print(message)
    print(f"已输出:{csv_path}、{xlsx_path}、{out_dir / 'notify_preview.txt'}")
    return {"csv": csv_path, "xlsx": xlsx_path, "summary": summary}


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    try:
        run(argv)
    except ApiError as e:
        log.error("%s", e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
