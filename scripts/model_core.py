# -*- coding: utf-8 -*-
"""港交所18C手册演示案例的唯一计算源。

数字只复现正文里的虚构算例，用来核对公式。
不是行业均值，也不是任何真实发行人的估值。
"""

from __future__ import annotations

BASE_REV = 2.14  # 2026年收入，亿元人民币
GROWTH = (0.60, 0.45, 0.30, 0.20, 0.12)
GM = (0.50, 0.56, 0.61, 0.65, 0.68)
RND = (0.42, 0.35, 0.28, 0.23, 0.20)
SGA = (0.32, 0.26, 0.23, 0.20, 0.18)
CAPEX = (0.08, 0.08, 0.07, 0.06, 0.05)
DA_RATIO = 0.03
NWC_RATIO = 0.04
TAX = 0.25
WACC = 0.16
G = 0.03
YEARS = (2027, 2028, 2029, 2030, 2031)

ARR0 = 1.20
NRR = 1.15
NEW_ARR = 0.80
TOKEN_REV = 0.45

NTM_MULTIPLE = 3.0
QUALITY_ADJ = -0.20

RNPV_SUCCESS_PV = 12.0
RNPV_POS = 0.15
RNPV_DEV_PV = 1.5

BRIDGE_OLD = 8.00
BRIDGE_OPS = -1.10
BRIDGE_MARGIN = -0.45
BRIDGE_MULTIPLE = -0.80
BRIDGE_CASH = 0.25


def arr_bridge():
    end_arr = ARR0 * NRR + NEW_ARR
    avg_arr = (ARR0 + end_arr) / 2
    total = avg_arr + TOKEN_REV
    return {
        "end_arr": end_arr,
        "avg_arr": avg_arr,
        "total": total,
    }


def forecast(base_rev: float = BASE_REV):
    rows = []
    prev = base_rev
    for i, year in enumerate(YEARS):
        rev = prev * (1 + GROWTH[i])
        gm = GM[i]
        rnd = RND[i]
        sga = SGA[i]
        ebitda = rev * (gm - rnd - sga)
        da = rev * DA_RATIO
        ebit = ebitda - da
        tax = max(0.0, ebit) * TAX
        nopat = ebit - tax
        capex = rev * CAPEX[i]
        delta_nwc = (rev - prev) * NWC_RATIO
        fcff = nopat + da - capex - delta_nwc
        rows.append(
            {
                "year": year,
                "growth": GROWTH[i],
                "rev": rev,
                "gm": gm,
                "rnd": rnd,
                "sga": sga,
                "ebitda": ebitda,
                "da": da,
                "ebit": ebit,
                "tax": tax,
                "nopat": nopat,
                "capex_ratio": CAPEX[i],
                "capex": capex,
                "delta_nwc": delta_nwc,
                "fcff": fcff,
            }
        )
        prev = rev
    return rows


def dcf(rows=None, wacc: float = WACC, g: float = G):
    if rows is None:
        rows = forecast()
    if wacc <= g:
        return None
    pv = 0.0
    pvs = []
    for t, row in enumerate(rows, start=1):
        factor = 1 / (1 + wacc) ** t
        value = row["fcff"] * factor
        pvs.append(value)
        pv += value
    last = rows[-1]["fcff"]
    tv = last * (1 + g) / (wacc - g)
    tv_pv = tv / (1 + wacc) ** len(rows)
    ev = pv + tv_pv
    return {
        "pv_fcff": pv,
        "pvs": pvs,
        "tv": tv,
        "tv_pv": tv_pv,
        "ev": ev,
        "tv_share": tv_pv / ev if ev else None,
    }


def sensitivity(waccs=(0.12, 0.14, 0.16, 0.18, 0.20), gs=(0.01, 0.02, 0.03, 0.04, 0.05)):
    rows = forecast()
    grid = []
    for w in waccs:
        line = []
        for g in gs:
            result = dcf(rows, w, g)
            line.append(None if result is None else result["ev"])
        grid.append(line)
    return list(waccs), list(gs), grid


def market_method(rows=None):
    if rows is None:
        rows = forecast()
    ntm = rows[0]["rev"]
    raw = ntm * NTM_MULTIPLE
    adjusted = raw * (1 + QUALITY_ADJ)
    ev = dcf(rows)["ev"]
    return {
        "ntm": ntm,
        "raw": raw,
        "adjusted": adjusted,
        "dcf": ev,
        "gap": adjusted - ev,
    }


def equity_bridge():
    new = BRIDGE_OLD + BRIDGE_OPS + BRIDGE_MARGIN + BRIDGE_MULTIPLE + BRIDGE_CASH
    return {
        "old": BRIDGE_OLD,
        "ops": BRIDGE_OPS,
        "margin": BRIDGE_MARGIN,
        "multiple": BRIDGE_MULTIPLE,
        "cash": BRIDGE_CASH,
        "new": new,
        "residual": new - (BRIDGE_OLD + BRIDGE_OPS + BRIDGE_MARGIN + BRIDGE_MULTIPLE + BRIDGE_CASH),
    }


def dilution(pre: float, cash: float):
    post = pre + cash
    return {
        "pre": pre,
        "cash": cash,
        "post": post,
        "new_share": cash / post,
        "old_share": pre / post,
    }


def rnpv():
    return RNPV_SUCCESS_PV * RNPV_POS - RNPV_DEV_PV


PUBLISHED_SENSITIVITY = {
    (0.14, 0.02): 6.58,
    (0.14, 0.03): 7.27,
    (0.14, 0.04): 8.10,
    (0.16, 0.02): 5.05,
    (0.16, 0.03): 5.52,
    (0.16, 0.04): 6.07,
    (0.18, 0.02): 3.94,
    (0.18, 0.03): 4.27,
    (0.18, 0.04): 4.65,
}


def verify():
    """对照手册正文的关键数字。失败即抛出 AssertionError。"""
    bridge = arr_bridge()
    assert abs(bridge["end_arr"] - 2.18) < 1e-12
    assert abs(bridge["avg_arr"] - 1.69) < 1e-12
    assert abs(bridge["total"] - BASE_REV) < 1e-12

    rows = forecast()
    assert abs(rows[0]["rev"] - 3.424) < 1e-12
    assert abs(rows[0]["ebitda"] - (-0.82176)) < 1e-9
    assert abs(rows[0]["fcff"] - (-1.14704)) < 1e-9
    assert abs(rows[1]["fcff"] - (-0.707056)) < 1e-8
    assert abs(rows[4]["fcff"] - 1.5459195648) < 1e-8

    result = dcf(rows)
    assert abs(result["tv"] - 12.2484) < 5e-4
    assert abs(result["pv_fcff"] - (-0.3120)) < 5e-4
    assert abs(result["tv_pv"] - 5.8316) < 5e-4
    assert abs(result["ev"] - 5.5196) < 5e-4
    assert result["tv_pv"] > result["ev"]

    _, _, grid = sensitivity()
    waccs = (0.12, 0.14, 0.16, 0.18, 0.20)
    gs = (0.01, 0.02, 0.03, 0.04, 0.05)
    for (w, g), published in PUBLISHED_SENSITIVITY.items():
        ev = grid[waccs.index(w)][gs.index(g)]
        assert abs(ev - published) < 0.015, (w, g, ev, published)

    mkt = market_method(rows)
    assert abs(mkt["raw"] - 10.272) < 1e-9
    assert abs(mkt["adjusted"] - 8.2176) < 1e-9
    assert abs(mkt["gap"] - 2.6980) < 5e-4

    eq = equity_bridge()
    assert abs(eq["new"] - 5.90) < 1e-12
    assert abs(eq["residual"]) < 1e-12

    base = dilution(8, 2)
    down = dilution(6, 2)
    assert abs(base["new_share"] - 0.20) < 1e-12
    assert abs(down["new_share"] - 0.25) < 1e-12
    assert abs(down["new_share"] - base["new_share"] - 0.05) < 1e-12
    assert abs(rnpv() - 0.30) < 1e-12
    return {
        "rows": rows,
        "dcf": result,
        "market": mkt,
        "bridge": eq,
        "sensitivity": grid,
    }


if __name__ == "__main__":
    out = verify()
    d = out["dcf"]
    print(f"EV {d['ev']:.4f}  TV {d['tv']:.4f}  TV_PV {d['tv_pv']:.4f}  FCFF_PV {d['pv_fcff']:.4f}")
    print("核对通过")
