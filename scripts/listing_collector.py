"""商办房源平台（好租/办办网/点点租等）挂牌行情采集。

真实模式：对各平台杨浦区商办挂牌页做结构化抓取（需遵守 robots 与平台条款）。
演示模式：基于载体清单合成挂牌房源（报价租金/物业费/可空置面积/上下架时序）。

字段对齐 fact_rental_transaction / fact_property_vacancy。
"""
import random

import config
from amap_collector import GRADE_RENT, fetch_demo as fetch_demo_assets


def fetch_demo_listings():
    random.seed(21)
    assets = fetch_demo_assets()
    listings = []
    for a in assets:
        grade = a["grade"]
        low, high = GRADE_RENT.get(grade, (3, 5))
        n_units = random.randint(2, 6)
        for u in range(n_units):
            quote = round(random.uniform(low, high), 2)
            # 成交价相对报价折让（议价系数）
            deal = round(quote * random.uniform(0.82, 0.95), 2)
            area = random.choice([200, 400, 600, 1000, 1500, 2500])
            listings.append({
                "property": a["name"],
                "subdistrict": a["subdistrict"],
                "plate": a["plate"],
                "grade": grade,
                "quote_rent": quote,
                "deal_rent": deal,
                "property_fee": round(random.uniform(18, 35), 0),
                "vacant_area_sqm": area,
                "vacant_days": random.choice([30, 60, 90, 150, 240]),
                "status": random.choices(["在租", "已成交"], weights=[7, 3])[0],
            })
    return listings


def collect():
    if config.DEMO_MODE:
        return fetch_demo_listings(), "DEMO"
    return [], "REAL"


if __name__ == "__main__":
    data, mode = collect()
    print(f"[listing] mode={mode} count={len(data)}")
