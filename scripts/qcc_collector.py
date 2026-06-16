"""企查查/天眼查 企业画像与工商变更（迁徙）采集。

真实模式：调用企查查开放平台 API（企业检索 / 地址检索 / 变更记录）。
演示模式：合成杨浦区入驻企业与迁徙样本，覆盖主导产业。

字段对齐 dim_company / fact_company_occupancy / fact_company_migration。
"""
import random
import requests

import config

QCC_SEARCH_URL = "https://api.qichacha.com/ECIV4/Search"  # 示例端点，以官方文档为准

INDUSTRIES = [
    ("AI/大模型", "信息技术"),
    ("在线新经济", "信息技术"),
    ("科技金融", "金融业"),
    ("设计创意", "文化创意"),
    ("检验检测/科技服务", "科学研究"),
    ("专精特新制造服务", "制造业"),
]

PLATES = ["大创智", "杨浦滨江", "新江湾", "五角场商圈", "国定路高校带", "内环商务"]
SIGNALS = ["近6月融资+招聘扩张", "工商稳定+岗位激增", "新设研发中心", "—"]


def fetch_real_company(keyword: str):
    params = {"key": config.QCC_KEY, "keyword": keyword}
    resp = requests.get(QCC_SEARCH_URL, params=params, timeout=15)
    return resp.json()


def fetch_demo_companies(n=320):
    random.seed(7)
    rows = []
    for i in range(n):
        ind, ind1 = random.choice(INDUSTRIES)
        plate = random.choices(PLATES, weights=[5, 4, 3, 3, 2, 2])[0]
        size = random.choices(["微型", "小型", "中型", "大型"], weights=[3, 4, 2, 1])[0]
        area = {"微型": (50, 200), "小型": (200, 600),
                "中型": (600, 2000), "大型": (2000, 8000)}[size]
        rows.append({
            "company_id": f"YP{i:04d}",
            "name": f"杨浦示例企业{i:04d}",
            "industry_tag": ind,
            "industry_lvl1": ind1,
            "scale": size,
            "area_sqm": random.randint(*area),
            "plate": plate,
            "reg_capital_wan": random.choice([100, 300, 500, 1000, 3000]),
            "status": "存续",
            "signal": random.choices(SIGNALS, weights=[2, 2, 1, 5])[0],
        })
    return rows


def fetch_demo_migration(n=90):
    random.seed(11)
    areas = ["黄浦区", "静安区", "徐汇区", "浦东新区", "杨浦-大创智", "杨浦-滨江",
             "杨浦-新江湾", "张江", "临港"]
    rows = []
    for i in range(n):
        frm = random.choice(areas)
        to = random.choice(["杨浦-大创智", "杨浦-滨江", "杨浦-新江湾", "张江", "临港", "徐汇区"])
        if frm == to:
            continue
        ind, _ = random.choice(INDUSTRIES)
        rows.append({
            "company_id": f"MIG{i:04d}",
            "from_area": frm,
            "to_area": to,
            "industry_tag": ind,
            "area_change": random.choice([-500, -200, 0, 300, 800, 1500]),
            "reason": random.choice(["成本驱动", "产业链协同", "政策牵引", "扩张"]),
        })
    return rows


def collect():
    if config.DEMO_MODE:
        return fetch_demo_companies(), fetch_demo_migration(), "DEMO"
    # 真实模式：按楼宇地址批量检索企业（此处留出对接位）
    companies, migrations = [], []
    return companies, migrations, "REAL"


if __name__ == "__main__":
    c, m, mode = collect()
    print(f"[qcc] mode={mode} companies={len(c)} migrations={len(m)}")
