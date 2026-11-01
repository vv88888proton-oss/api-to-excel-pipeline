from pathlib import Path
from unittest import mock

import pytest
import requests
from openpyxl import load_workbook

import pipeline as pl


def ticker(inst, last, open_, vol=1000.0, bid=None, ask=None, ts="1793500000000"):
    bid = bid if bid is not None else last * 0.999
    ask = ask if ask is not None else last * 1.001
    return {"instId": inst, "last": str(last), "open24h": str(open_), "high24h": str(max(last, open_) * 1.01),
            "low24h": str(min(last, open_) * 0.99), "bidPx": str(bid), "askPx": str(ask),
            "vol24h": "1", "volCcy24h": str(vol), "ts": ts}


ROWS = [
    ticker("BTC-USDT", 110, 100, vol=5000),
    ticker("ETH-USDT", 90, 100, vol=3000),
    ticker("ETH-USDT", 90, 100, vol=3000),      # 重复
    ticker("SOL-USDC", 50, 50, vol=9000),        # 计价币不符
    {**ticker("BAD-USDT", 1, 1), "last": ""},    # 缺价格
    ticker("TINY-USDT", 2, 1, vol=10),
]


def resp(status=200, payload=None):
    r = mock.Mock()
    r.status_code = status
    r.json.return_value = payload if payload is not None else {"code": "0", "data": ROWS}
    r.text = "err"
    return r


def test_clean_filters_dedups_and_computes():
    df = pl.clean_tickers(ROWS, quote="USDT", top=None)
    assert df["instId"].tolist() == ["BTC-USDT", "ETH-USDT", "TINY-USDT"]  # 按成交额降序
    assert df.loc[0, "change_24h_pct"] == pytest.approx(10.0)
    assert df.loc[1, "change_24h_pct"] == pytest.approx(-10.0)
    assert df.loc[0, "spread_pct"] == pytest.approx(0.2, abs=1e-3)
    assert df.loc[0, "time_cn"].startswith("2026-")


def test_clean_min_volume_and_top():
    df = pl.clean_tickers(ROWS, quote="USDT", min_volume=100, top=1)
    assert df["instId"].tolist() == ["BTC-USDT"]


def test_clean_empty():
    assert pl.clean_tickers([]).empty


def test_fetch_retries_then_succeeds():
    sess = mock.Mock()
    sess.get.side_effect = [requests.Timeout("t"), resp(503), resp()]
    sleeps = []
    rows = pl.fetch_tickers(session=sess, retries=3, backoff=0.5, sleep=sleeps.append)
    assert len(rows) == len(ROWS)
    assert sess.get.call_count == 3
    assert sleeps == [0.5, 1.0]  # 指数退避
    assert sess.get.call_args.kwargs["timeout"] == 10.0


def test_fetch_gives_up_after_retries():
    sess = mock.Mock()
    sess.get.side_effect = requests.ConnectionError("down")
    with pytest.raises(pl.ApiError):
        pl.fetch_tickers(session=sess, retries=2, sleep=lambda s: None)
    assert sess.get.call_count == 2


def test_fetch_4xx_not_retried_and_api_error_code():
    sess = mock.Mock()
    sess.get.return_value = resp(400)
    with pytest.raises(pl.ApiError):
        pl.fetch_tickers(session=sess, sleep=lambda s: None)
    assert sess.get.call_count == 1
    sess.get.return_value = resp(200, {"code": "51001", "msg": "bad instType"})
    with pytest.raises(pl.ApiError, match="51001"):
        pl.fetch_tickers(session=sess, sleep=lambda s: None)


def test_run_end_to_end_with_mock(tmp_path):
    cfg = tmp_path / "notify.json"
    cfg.write_text('{"enabled": true, "channel": "webhook", "webhook_url": "https://example.invalid"}', encoding="utf-8")
    with mock.patch.object(pl, "fetch_tickers", return_value=ROWS) as f, \
         mock.patch.object(pl.requests, "post") as post:
        res = pl.run(["--out", str(tmp_path), "--top", "5", "--notify-config", str(cfg)])
    f.assert_called_once()
    post.assert_not_called()  # 通知只做预览,绝不真实发送
    assert res["summary"]["count"] == 3
    wb = load_workbook(res["xlsx"])
    ws = wb["行情快照"]
    assert ws["A2"].value == "交易对" and ws["A3"].value == "BTC-USDT"
    assert ws.freeze_panes == "A3"
    assert "汇总" in wb.sheetnames
    preview = (tmp_path / "notify_preview.txt").read_text(encoding="utf-8")
    assert "未发送" in preview and "TINY-USDT" in preview


def test_main_returns_1_on_api_error(tmp_path):
    with mock.patch.object(pl, "fetch_tickers", side_effect=pl.ApiError("boom")):
        assert pl.main(["--out", str(tmp_path)]) == 1
