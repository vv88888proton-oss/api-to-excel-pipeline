# 04 API → Excel 数据管道(拉取 → 清洗 → 带样式报表 → 通知预览)

## 解决什么问题
常见需求:“每天定时从某个接口拉一次数据,整理成一张好看的 Excel 发给我/发到群里。”
本项目以 OKX **公开、免登录**的行情快照接口(`/api/v5/market/tickers`)为例,一条命令完成:

1. 拉取:带超时、指数退避重试(超时 / 连接错误 / 429 / 5xx 重试,4xx 和业务错误码直接报错)
2. 清洗:去重、字符串转数值、剔除缺价行、按计价币和成交额过滤、计算 24h 涨跌幅 / 振幅 / 买卖价差,时间转北京时间
3. 输出:UTF-8-BOM CSV(Excel 直接打开不乱码)+ 带样式的 Excel(中文表头、冻结首行、自动筛选、千分位、涨绿跌红条件格式、“汇总”工作表)
4. 通知:**只做可配置的占位**,把将要发送的内容写到 `output/notify_preview.txt`,**不会真实发送任何消息**

**只读公开数据,不需要 API Key,不涉及任何下单或账户操作。** 换成其他 JSON 接口时,只需改 `fetch_tickers()` 和 `clean_tickers()` 两个函数。

## 使用方法
```bash
pip install -r requirements.txt
python pipeline.py --top 30 --notify-config notify_config.example.json      # 实时拉取
python pipeline.py --inst-type SWAP --quote USDT --min-volume 1000000 --out output/swap
python make_sample_data.py && python pipeline.py --offline sample_data/tickers_sample.json --top 10 --out output/offline_demo
python -m pytest -q   # 运行测试(8 个,网络请求全部 mock,不联网)
```

| 参数 | 说明 | 默认 |
|---|---|---|
| `--inst-type` | SPOT / SWAP / FUTURES | SPOT |
| `--quote` | 只保留该计价币,传 `""` 不过滤 | USDT |
| `--top` | 按成交额取前 N,0 为全部 | 30 |
| `--min-volume` | 24h 成交额下限 | 0 |
| `--timeout` / `--retries` | 单次请求超时秒数 / 最大尝试次数 | 10 / 3 |
| `--base-url` | 接口域名,可替换 | https://www.okx.com |
| `--offline` | 读取本地 JSON(接口原始格式),不联网 | — |
| `--notify-config` | 通知配置 JSON,仅生成预览 | — |
| `--out` | 输出目录 | output |

定时运行可交给系统自带的计划任务(Linux `cron` / Windows 任务计划程序),例如每天 9 点:
`0 9 * * * cd /path/to/04_api_to_excel_pipeline && python pipeline.py`

## 关于通知
`notify_config.example.json` 是占位配置(`webhook_url` 为无效示例地址)。`notify()` 只写预览文件,
测试里也断言了不会发出任何 HTTP POST。如需接入企业微信 / 飞书 / 钉钉机器人或邮件,由使用者自行提供并保管 webhook 地址或密钥,再在代码中实现发送。

## 测试覆盖
- 清洗:去重、计价币过滤、缺价剔除、成交额排序、涨跌幅/价差计算、空数据
- 请求:超时 + 503 后重试成功并校验退避间隔;重试耗尽报错;4xx 不重试;业务错误码报错
- 端到端:mock 接口后跑完整流程,校验 Excel 表头、冻结窗格、汇总表、通知未真实发送
- 命令行:接口失败时返回退出码 1

## 示例输出(真实运行结果,2026-11-01 13:42 北京时间,SPOT USDT 成交额 Top30)
```
共 30 个交易对,上涨 22 / 下跌 8,涨跌幅中位数 1.01%
instId,last,change_24h_pct,amplitude_pct,spread_pct,volume_quote_24h
BTC-USDT,77764.4,0.7,0.87,0.0001,220966327
ETH-USDT,2265.07,1.08,1.67,0.0004,129842837
```
输出文件:`output/okx_spot_usdt_20261101.xlsx`、`output/okx_spot_usdt_20261101.csv`、`output/notify_preview.txt`;
离线演示在 `output/offline_demo/`(`sample_data/` 是随机生成的模拟数据,不是真实行情)。

## 技术栈
Python 3.10+、requests、pandas、openpyxl(样式、条件格式)、argparse、pytest(unittest.mock)

## 免责声明
仅为数据处理演示,不构成任何投资建议。
