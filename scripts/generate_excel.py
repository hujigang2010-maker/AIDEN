#!/usr/bin/env python3
"""Generate Excel workbook from Sam's Club business plan tables."""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/exports/山姆选品商机与商务路径方案.xlsx"

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14)
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def write_table(ws, start_row, title, headers, rows, col_widths=None):
    ws.cell(row=start_row, column=1, value=title).font = TITLE_FONT
    hr = start_row + 1
    for i, h in enumerate(headers, 1):
        ws.cell(row=hr, column=i, value=h)
    style_header(ws, hr, len(headers))
    for ri, row_data in enumerate(rows, hr + 1):
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
    if col_widths:
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    return hr + len(rows) + 2


def main():
    wb = Workbook()
    wb.remove(wb.active)

    # Sheet 1: 选品机会
    ws1 = wb.create_sheet("选品机会")
    r = 1
    r = write_table(ws1, r, "表1：已验证成功品类", 
        ["品类大类", "已跑通方向", "商机切入点", "难度", "备注"],
        [
            ["零食", "各类休闲零食", "健康化、低糖、儿童向", "中", "复购高"],
            ["酒水", "已成熟品类", "精酿、低度果酒、米酒", "中", "关注政策税务"],
            ["调料", "李锦记等", "复合调味料、减盐系列", "中低", "差异化替代"],
            ["有机农产品", "已成熟", "溯源化、礼盒化", "中", "产地直采"],
            ["生鲜", "大米、大闸蟹", "青岛虾米、特色海产", "高", "冷链产能要求高"],
            ["功能性洗护", "防过敏儿童", "母婴、敏感肌", "中", "非食品已跑通"],
            ["药食同源饮品", "姜柠檬饮、石斛饮", "短保健康饮品", "中高", "当前风口"],
        ], [14, 18, 28, 8, 18])

    r = write_table(ws1, r, "表2：差异化替代机会",
        ["现有坑位", "替代产品建议", "逻辑"],
        [
            ["火锅牛肉丸", "黑猪肉马蹄丸、墨鱼虾滑", "同坑位替换"],
            ["普通甜饮", "药食同源养生饮、低糖植物饮", "健康升级"],
            ["普通大米", "地标稻米、胚芽米", "产地溢价"],
            ["常规调味酱", "地域复合酱、低钠版", "健康化升级"],
        ], [18, 28, 20])

    write_table(ws1, r, "表3：空白机会",
        ["品类", "现状", "机会判断"],
        [
            ["牛肉", "检疫严、十亿级", "中长期，需大资方"],
            ["蔬菜", "≥3吨/日、5品起签", "产地农业集团打包"],
            ["国际选品", "结算复杂未做", "跨境/保税空白"],
        ], [14, 22, 28])

    # Sheet 2: 商务路径
    ws2 = wb.create_sheet("商务路径")
    r = 1
    r = write_table(ws2, r, "表5：合作模式对比",
        ["维度", "自有品牌", "山姆贴牌"],
        [
            ["合同周期", "3个月→6/12个月", "24个月"],
            ["铺货范围", "区域（华东/上海）", "全国60店"],
            ["盲测/验厂", "标准", "更高"],
            ["额外费用", "无", "+15万贴牌费"],
            ["退货", "按常规", "买断制不退货"],
            ["适合", "中小品牌试点", "有产能打爆款"],
        ], [16, 28, 28])

    r = write_table(ws2, r, "表6：流程费用",
        ["阶段", "内容", "费用/时间"],
        [
            ["报名", "提交资料", "报名费先付"],
            ["盲测", "第三方消费者调研", "不可造假"],
            ["多轮筛选", "评估", "—"],
            ["验厂", "第三阶段", "退出收20%商务费"],
            ["上架", "全流程通过", "约6个月"],
            ["总费用", "单品", "约15万（贴牌+15万）"],
        ], [12, 28, 22])

    r = write_table(ws2, r, "表7：分润结构",
        ["分配方", "比例", "说明"],
        [
            ["厂家", "77%", "生产+供货"],
            ["山姆", "23%", "渠道"],
            ["中介总抽点", "1-5%", "综合服务"],
            ["采购方", "3%", "三三三一"],
            ["中介", "3%", ""],
            ["对接方", "3%", ""],
            ["预留", "1%", ""],
        ], [14, 10, 30])

    # Sheet 3: 10产品清单
    ws3 = wb.create_sheet("10产品清单")
    write_table(ws3, 1, "10个试点产品",
        ["优先级", "产品", "替代对象", "品牌候选", "评分", "建议模式"],
        [
            ["★★★", "姜黄/生姜柠檬冷压饮", "普通甜饮", "好望水、一甸园", "88", "自有品牌试点"],
            ["★★★", "黑猪肉马蹄丸", "牛肉丸", "海欣、潮庭", "86", "自有品牌试点"],
            ["★★★", "青岛干虾米", "普通海产干货", "青岛本地加工厂", "85", "自有→贴牌"],
            ["★★☆", "石斛即饮植物饮", "功能饮料", "仙草集、同仁堂", "82", "自有品牌试点"],
            ["★★☆", "低钠复合调味酱", "常规调味酱", "李锦记、欣和", "80", "自有品牌"],
            ["★★☆", "有机胚芽米5kg", "普通大米", "北大荒、十月稻田", "79", "自有品牌"],
            ["★★☆", "墨鱼虾滑", "虾滑", "海欣、安井", "78", "自有品牌"],
            ["★★☆", "儿童防敏洗护", "常规母婴", "启初、薇诺娜宝贝", "77", "自有品牌"],
            ["★☆☆", "精酿米酒礼盒", "普通甜酒", "苏州桥、石库门", "74", "区域试点"],
            ["★☆☆", "低糖冻干水果脆", "普通零食", "三只松鼠、良品铺子", "73", "自有品牌"],
        ], [8, 22, 16, 22, 8, 16])

    # Sheet 4: 行动计划
    ws4 = wb.create_sheet("行动计划")
    r = 1
    r = write_table(ws4, r, "项目推进表",
        ["阶段", "动作", "输出"],
        [
            ["第1周", "收集品牌", "50-100名单"],
            ["第2周", "初筛", "10个候选"],
            ["第3周", "深度评估", "3个产品"],
            ["第4周", "对接山姆/中介", "确定1-2试点"],
            ["1-6月", "跑流程", "上架"],
        ], [12, 24, 24])

    write_table(ws4, r, "待办清单",
        ["事项", "时间", "状态"],
        [
            ["提交初步品牌名单+方案", "本周五前", "待办"],
            ["三方见面洽谈入仓", "下周", "待办"],
            ["联系青岛虾米供应商", "持续", "待办"],
            ["1-2试点全流程验证", "同步推进", "待办"],
        ], [30, 14, 12])

    # Sheet 5: 选品模型
    ws5 = wb.create_sheet("选品评估模型")
    write_table(ws5, 1, "五维评估模型",
        ["维度", "权重", "判断标准"],
        [
            ["复购率", "30%", "是否日常消费"],
            ["健康属性", "20%", "是否有功能卖点"],
            ["差异化", "20%", "是否替代现有SKU"],
            ["供应能力", "20%", "是否稳定供货"],
            ["成本结构", "10%", "是否有利润空间"],
        ], [16, 10, 40])

    wb.save(OUTPUT)
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
