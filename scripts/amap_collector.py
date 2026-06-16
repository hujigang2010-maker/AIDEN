"""高德 POI/AOI 楼宇·园区基础数据采集。

真实模式：调用高德 Web 服务 API（关键字 POI 搜索）。
演示模式：返回合成的杨浦区典型载体清单。

字段对齐 dim_property_asset：名称/类型/等级/坐标/街道/建面/可租面积/运营方/配套。
"""
import time
import random
import requests

import config

AMAP_POI_URL = "https://restapi.amap.com/v3/place/text"

# 关键字 → 物业类型映射
KEYWORDS = {
    "写字楼": "office",
    "产业园区": "park",
    "科技园": "park",
    "孵化器": "park",
}


def fetch_real(keyword: str, page_size: int = 25, max_pages: int = 4):
    """调用高德关键字 POI 搜索（真实模式）。"""
    rows = []
    for page in range(1, max_pages + 1):
        params = {
            "key": config.AMAP_KEY,
            "keywords": keyword,
            "city": config.TARGET_CITY,
            "types": "120000|170000",  # 商务住宅/科教文化（含写字楼、产业园）
            "offset": page_size,
            "page": page,
            "extensions": "all",
        }
        resp = requests.get(AMAP_POI_URL, params=params, timeout=15)
        data = resp.json()
        pois = data.get("pois", [])
        if not pois:
            break
        for p in pois:
            adname = p.get("adname", "")
            if config.TARGET_DISTRICT not in adname:
                continue
            loc = (p.get("location") or ",").split(",")
            rows.append({
                "name": p.get("name"),
                "asset_type": KEYWORDS.get(keyword, "office"),
                "district": adname,
                "subdistrict": p.get("business_area") or "",
                "address": p.get("address"),
                "longitude": loc[0] if loc[0] else None,
                "latitude": loc[1] if len(loc) > 1 else None,
            })
        time.sleep(0.3)
    return rows


# ---- 演示模式：杨浦区典型载体（真实板块/园区名，面积等为合成示例值）----
DEMO_ASSETS = [
    ("创智天地企业中心", "office", "五角场街道", "甲级", "大创智"),
    ("INNO创智", "park", "五角场街道", "科创园", "大创智"),
    ("国正中心", "office", "五角场镇", "甲级", "大创智"),
    ("合生汇写字楼", "office", "五角场镇", "甲级", "五角场商圈"),
    ("万达广场写字楼", "office", "五角场镇", "乙级", "五角场商圈"),
    ("湾谷科技园", "park", "新江湾城街道", "科创园", "新江湾"),
    ("尚浦中心", "office", "新江湾城街道", "甲级", "新江湾"),
    ("互联宝地", "park", "平凉路街道", "科创园", "杨浦滨江"),
    ("financial street 滨江", "office", "平凉路街道", "甲级", "杨浦滨江"),
    ("长阳创谷", "park", "大桥街道", "科创园", "杨浦滨江"),
    ("城市概念园区", "park", "大桥街道", "孵化器", "杨浦滨江"),
    ("控江路商务楼", "office", "控江路街道", "乙级", "内环商务"),
    ("四平科技园", "park", "四平路街道", "科创园", "同济周边"),
    ("江浦路商务中心", "office", "江浦路街道", "乙级", "内环商务"),
    ("复旦软件园", "park", "五角场街道", "科创园", "国定路高校带"),
    ("国定东路孵化基地", "park", "五角场街道", "孵化器", "国定路高校带"),
    ("殷行科创空间", "park", "殷行街道", "孵化器", "外环"),
    ("延吉商务楼", "office", "延吉新村街道", "乙级", "内环商务"),
    ("长白科创园", "park", "长白新村街道", "科创园", "内环商务"),
    ("定海科创社区", "park", "定海路街道", "孵化器", "滨江东延伸"),
]

GRADE_RENT = {"甲级": (5.5, 7.5), "乙级": (3.2, 4.8), "科创园": (3.0, 5.0), "孵化器": (2.5, 4.0)}


def fetch_demo():
    random.seed(42)
    rows = []
    for name, atype, sub, grade, plate in DEMO_ASSETS:
        gfa = round(random.uniform(3, 18), 1)
        rentable = round(gfa * random.uniform(0.55, 0.75), 1)
        rows.append({
            "name": name,
            "asset_type": atype,
            "district": "杨浦区",
            "subdistrict": sub,
            "plate": plate,
            "grade": grade,
            "address": f"上海市杨浦区{sub}（示例）",
            "total_gfa_wan": gfa,
            "rentable_area_wan": rentable,
            "metro_distance_m": random.choice([200, 350, 500, 700, 900]),
            "operator": random.choice(["国资平台", "品牌开发商", "专业运营商", "外资基金"]),
        })
    return rows


def collect():
    if config.DEMO_MODE:
        return fetch_demo(), "DEMO"
    rows = []
    for kw in KEYWORDS:
        rows.extend(fetch_real(kw))
    return rows, "REAL"


if __name__ == "__main__":
    data, mode = collect()
    print(f"[amap] mode={mode} count={len(data)}")
