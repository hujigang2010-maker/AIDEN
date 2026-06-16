# 杨浦试点数据采集与产出管线

一套可真实接入 **高德 / 企查查 / 房源平台** 的采集脚本，并内置「演示数据模式」把杨浦占位数据端到端跑通，自动产出图表、数据填充版报告与 PPT 路演稿。

## 运行模式
- **真实模式**：在 `scripts/.env`（复制自 `.env.example`）填入 `AMAP_KEY`、`QCC_KEY` 后自动启用。
- **演示模式（默认）**：无 Key 时降级为合成数据，结构/口径与真实一致，便于快速验证全链路。

## 一键跑通
```bash
pip install -r scripts/requirements.txt
cd scripts
python3 run_yangpu_pipeline.py     # 采集→清洗→聚合→落盘 data/yangpu/*.csv + headline.json
python3 make_charts.py             # 生成 charts/*.png（中文字体）
python3 build_filled_report.py     # 生成「杨浦区数据填充版」Markdown
python3 make_ppt.py                # 生成 PPT 路演稿（下载版本/*.pptx）
```

## 脚本说明
| 文件 | 作用 | 对应数据库表 |
|---|---|---|
| `config.py` | 配置与运行模式判定、中文字体路径 | — |
| `amap_collector.py` | 高德 POI/AOI 楼宇·园区基础数据 | `dim_property_asset` |
| `qcc_collector.py` | 企查查/天眼查 企业画像 + 工商变更（迁徙） | `dim_company` / `fact_company_migration` |
| `listing_collector.py` | 房源平台挂牌行情（报价/成交/空置） | `fact_rental_transaction` / `fact_property_vacancy` |
| `run_yangpu_pipeline.py` | 编排采集与聚合，输出 CSV + headline.json | DWS 聚合层 |
| `make_charts.py` | 6 张分析图表（PNG，Noto CJK） | — |
| `build_filled_report.py` | 数据填充版 Markdown 报告 | ADS 应用层 |
| `make_ppt.py` | PPT 路演稿（资本/产品视角） | — |

## 产出物
- `data/yangpu/*.csv`、`headline.json`：结构化数据集与关键指标。
- `charts/*.png`：分街道供给、等级结构、租金/空置分化、产业画像、In/Out、迁徙来源。
- `下载版本/…杨浦区数据填充版.docx/.pdf`、`…PPT路演稿.pptx/.pdf`。

> 合规：仅采公开数据，企业信息脱敏；遵循各平台服务条款与《上海市数据条例》。
