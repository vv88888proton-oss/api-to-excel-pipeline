"""生成离线演示用的模拟 tickers 数据(随机数,不是真实行情)。"""
import json
import random
from pathlib import Path

random.seed(42)
coins = ["AAA", "BBB", "CCC", "DDD", "EEE", "FFF", "GGG", "HHH", "III", "JJJ", "KKK", "LLL"]
ts = 1793500000000
rows = []
for i, c in enumerate(coins):
    for q in ["USDT", "USDC"]:
        open_ = round(random.uniform(0.5, 500), 4)
        last = round(open_ * random.uniform(0.9, 1.1), 4)
        hi, lo = max(open_, last) * 1.02, min(open_, last) * 0.98
        vol = random.uniform(1e3, 1e6)
        rows.append({"instType": "SPOT", "instId": f"{c}-{q}", "last": str(last), "open24h": str(open_),
                     "high24h": f"{hi:.4f}", "low24h": f"{lo:.4f}", "bidPx": f"{last*0.9995:.4f}",
                     "askPx": f"{last*1.0005:.4f}", "vol24h": f"{vol:.2f}", "volCcy24h": f"{vol*last:.2f}",
                     "ts": str(ts + i)})
rows.append(dict(rows[0]))                       # 重复行
rows.append({**rows[1], "instId": "BAD-USDT", "last": ""})  # 缺失价格
out = Path(__file__).parent / "sample_data" / "tickers_sample.json"
out.parent.mkdir(exist_ok=True)
out.write_text(json.dumps({"code": "0", "msg": "", "data": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"已生成 {out}({len(rows)} 行,含 1 条重复、1 条缺价)")
