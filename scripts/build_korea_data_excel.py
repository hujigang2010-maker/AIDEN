#!/usr/bin/env python3
"""生成《以韩为鉴》数据底稿 Excel：output/以韩为鉴_数据底稿.xlsx。

工作表：
1. 核心数据（指标、数值、时间、所属章节、来源、链接）
2. 中韩对照
3. 周期坐标
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output" / "以韩为鉴_数据底稿.xlsx"

KR_DATA = [
    ("宏观", "实际GDP增速", "1.0%（0.97%）", "2025年", "第三章", "韩国银行2025年四季度及全年GDP速报",
     "https://www.bok.or.kr/portal/bbs/B0000170/view.do?nttId=10096012&menuNo=201297"),
    ("宏观", "名义GDP同比增速", "17.1%（1995年三季度以来最高）", "2026年一季度", "第三章", "首尔经济日报（BIS数据报道）",
     "https://en.sedaily.com/finance/2026/06/16/korea-household-debt-ratio-hits-lowest-level-in-over-six"),
    ("宏观", "实际GDP环比", "一季度+1.8%，二季度+0.6%", "2026年", "第三章", "韩国银行（经Global Economic）",
     "https://www.g-enews.com/article/Finance/2026/07/202607230826467005bb91c46fcd_1"),
    ("宏观", "全年出口", "7097亿美元（+3.8%，首破7000亿）", "2025年", "第三章", "韩国产业通商资源部",
     "https://english.motir.go.kr/eng/article/EATCLdfa319ada/2470/view"),
    ("宏观", "半导体出口", "1734亿美元（+22.2%）", "2025年", "第三、十三章", "韩国产业通商资源部",
     "https://english.motir.go.kr/eng/article/EATCLdfa319ada/2470/view"),
    ("宏观", "6月出口", "1022.5亿美元（+70.9%），半导体448.2亿美元", "2026年6月", "第三章", "证券时报",
     "https://stcn.com/article/detail/4037104.html"),
    ("资产", "KOSPI历史高点", "9063.84点", "2026-06-18", "第十九章", "Businesskorea",
     "https://www.businesskorea.co.kr/news/articleView.html?idxno=271588"),
    ("资产", "KOSPI与汇率", "7003.74点；1350.6韩元/美元", "2026-10-02", "第十九章", "Businesskorea",
     "https://www.businesskorea.co.kr/news/articleView.html?idxno=278173"),
    ("宏观", "基准利率", "2.75%（2023年1月以来首次加息）", "2026年7月", "第三章", "Trading Economics/韩国银行",
     "https://tradingeconomics.com/south-korea/interest-rate"),
    ("人口", "总和生育率", "0.72（2023）/0.75（2024）/0.80（2025）", "2025年", "第七章", "韩国政策简报（国家数据处）",
     "https://www.korea.kr/news/policyNewsView.do?newsId=148970652"),
    ("人口", "出生人口", "25.4341万（+6.7%）", "2025年", "第七章", "同上",
     "https://www.korea.kr/news/policyNewsView.do?newsId=148970652"),
    ("人口", "首尔/全南生育率", "0.63 / 1.09", "2025年", "第六、七章", "同上",
     "https://www.korea.kr/news/policyNewsView.do?newsId=148970652"),
    ("人口", "上半年出生人口", "14.5804万（+15.4%），二季度TFR 0.88", "2026年上半年", "第七章", "韩国日报",
     "https://www.hankookilbo.com/news/article/A2026082611200004875"),
    ("人口", "母亲平均生育年龄", "33.8岁；35岁以上占37.3%", "2025年", "第七章", "韩国政策简报",
     "https://www.korea.kr/briefing/policyBriefingView.do?newsId=156745912"),
    ("人口", "低生育对策投入", "约280万亿韩元（2006—2021年）", "2006—2021", "第七章", "韩民族日报/文化日报",
     "https://www.hani.co.kr/arti/economy/economy_general/1111560.html"),
    ("就业", "15—29岁就业率", "44.1%（连续28个月下降）；就业人数连续46个月减少", "2026年8月", "第五章", "Herald Business",
     "https://biz.heraldcorp.com/article/10867335"),
    ("就业", "15—29岁失业率", "5.4%", "2026年8月", "第五章", "首尔经济日报",
     "https://en.sedaily.com/finance/2026/09/09/korea-adds-184000-jobs-in-august-as-youth-employment-slides"),
    ("就业", "20—39岁“休息”人数", "71.7万（历史最高）", "2025年", "第五章", "MoneyToday",
     "https://www.mt.co.kr/en/economy/2026/08/28/2026082814495072507"),
    ("就业", "“休息”青年占非经济活动人口比例", "14.6%（2019）→22.3%（2025）", "2025年", "第五章", "韩国银行BOK Issue Note（经首尔经济日报）",
     "https://en.sedaily.com/finance/2026/01/20/prolonged-unemployment-not-high-expectations-drives-youth"),
    ("就业", "主力岗位平均退休年龄", "52.9岁", "2025年", "第五、十六章", "中央日报",
     "https://www.joongang.co.kr/article/25385192"),
    ("就业", "65岁以上就业率", "37.3%（OECD 13.6%，日本25.3%）", "2023年", "第五、十五章", "韩国日报（国会预算政策处）",
     "https://www.hankookilbo.com/news/article/A2025052710310005247"),
    ("就业", "IMF危机失业率", "1997.10 2.1%→1998.07 7.6%→1999.02 8.7%", "1997—1999", "第四章", "KDI School w99-01",
     "https://archives.kdischool.ac.kr/bitstream/11125/29132/1/w99-01.pdf"),
    ("就业", "1998年失业人数", "146.3万（失业率6.8%）", "1998年", "第四章", "KDI",
     "https://www.kdi.re.kr/research/reportView?pub_no=83"),
    ("就业", "收集黄金运动", "约351万人参与，约227吨，出口约22亿美元", "1998年", "第四章", "韩国民族文化大百科",
     "https://encykorea.aks.ac.kr/Article/E0080908"),
    ("收入", "大企业/中小企业月均收入", "613万/307万韩元", "2024年", "第十三章", "KDI经济信息中心（国家数据处）",
     "https://eiec.kdi.re.kr/policy/materialView.do?num=277139"),
    ("收入", "工资岗位平均/中位收入", "375万/288万韩元", "2024年", "第十三章", "同上",
     "https://eiec.kdi.re.kr/policy/materialView.do?num=277139"),
    ("职业", "9级公务员竞争率", "93.3（2011）/53.8（2016）/21.8（2024）/24.3（2025）", "2025年", "第十章", "韩国人事革新处；韩联社",
     "https://www.yna.co.kr/view/AKR20250208027900001"),
    ("职业", "9级起薪", "269万→300万韩元（2027年目标）", "2025年", "第十章", "世界日报",
     "https://www.segye.com/newsView/20250209501324"),
    ("职业", "公立小学教师录用", "7386（2014）→3157（2024）→3113（2026学年度，-27.1%）", "2026年", "第十一章", "东亚日报；Newspim",
     "https://www.donga.com/news/Society/article/all/20250910/132356290/1"),
    ("职业", "小学教师录用合格率", "68.9%（2018）→43.6%（2024）", "2024年", "第十一章", "Newspim",
     "https://www.newspim.com/news/view/20240411000615"),
    ("职业", "医师国考合格者", "3045（2024）→269（2025）", "2025年", "第十二章", "韩联社",
     "https://www.yna.co.kr/view/AKR20250807103900530"),
    ("职业", "国立大学医院医生补充率", "73.1%→54.9%→73.5%", "2024.6—2026.5", "第十二章", "财经日报",
     "https://news.jkn.co.kr/post/1007088"),
    ("职业", "理工科硕博3年内考虑出国", "42.9%（20—30多岁约70%）", "2025年", "第十三章", "韩国银行BOK Issue Note 2025-31",
     "https://www.bok.or.kr/portal/bbs/P0002353/view.do?nttId=10094375&menuNo=200433"),
    ("职业", "首尔大学工学院退学", "523人（2020—2025），约占30%", "2025年", "第十三章", "韩国大学新闻",
     "https://news.unn.net/news/articleView.html?idxno=584835"),
    ("自营业", "停业经营者", "100.8万（2024）/97.6万（2025）", "2025年", "第十四章", "韩联社；Daum",
     "https://www.yna.co.kr/view/AKR20250705035000002"),
    ("自营业", "自营业者贷款", "1069.6万亿韩元；低收入逾期率2.07%", "2025年二季度", "第十四章", "韩联社",
     "https://www.yna.co.kr/view/AKR20251011040900002"),
    ("自营业", "60岁以上自营业者", "269.7万（41.2%）", "2025年", "第十四章", "E-Daily（韩国银行金融稳定报告）",
     "https://www.edaily.co.kr/news/read?mediaCodeNo=257&newsId=03703126645485000"),
    ("自营业", "自营业者比重（OECD口径）", "23.2%，OECD第7", "2023年", "第十四章", "韩联社（韩国银行）",
     "https://www.yna.co.kr/view/AKR20250515064700002"),
    ("教育", "课外补习总额", "约27.5万亿韩元（-5.7%）；2024年约29.2万亿", "2025年", "第八章", "韩国国家数据处",
     "https://mods.go.kr/board.es?act=view&bid=245&list_no=443953&mid=a10301070100"),
    ("教育", "参与学生人均月补习费", "60.4万韩元（高中79.3万）", "2025年", "第八章", "同上",
     "https://mods.go.kr/board.es?act=view&bid=245&list_no=443953&mid=a10301070100"),
    ("住房", "首尔公寓平均成交价", "13.6214亿韩元", "2026年7月", "第二章", "首尔新闻（韩国不动产院）",
     "https://www.seoul.co.kr/news/economy/estate/2026/08/19/20260819030003"),
    ("住房", "首尔PIR（分位）", "1分位29.36/3分位10.49/5分位4.44", "2026年3月", "第二章", "首尔新闻通信（KB不动产）",
     "http://www.snakorea.com/news/articleView.html?idxno=1026390"),
    ("住房", "传贳诈骗受害者", "累计40278人，40岁以下75.9%", "2026年8月", "第二、九章", "首尔新闻（国土交通部）",
     "https://www.seoul.co.kr/news/economy/2026/08/07/20260807500013"),
    ("住房", "首尔传贳价格累计涨幅", "5.91%（15年来最高）", "2026年1—8月", "第二章", "首尔经济日报",
     "https://en.sedaily.com/finance/2026/09/29/seoul-jeonse-prices-climb-at-fastest-pace-in-15-years"),
    ("债务", "家庭债务/GDP（BIS）", "88.6%（峰值2021Q3 99.1%）", "2025年末", "第二章", "KBS World",
     "https://world.kbs.co.kr/service/news_view.htm?lang=e&Seq_Code=202213"),
    ("债务", "银行传贳贷款余额", "约165.7万亿韩元", "2026年一季度末", "第二章", "朝鲜日报经济版",
     "https://biz.chosun.com/en/en-finance/2026/08/26/H4D6SCMLHFD4DDWVPW3WR72MGU/"),
    ("区域", "首都圈人口占比", "50.8%（2630万）", "2024年", "第六章", "仁川Today（统计厅）",
     "https://www.incheontoday.com/news/articleView.html?idxno=305327"),
    ("区域", "消亡风险市郡区", "62个（新分类，其中7个严重）", "2025年6月", "第六章", "Economic Post（韩国雇佣信息院）",
     "https://economicpost.co.kr/146386"),
    ("养老", "66岁以上老年贫困率", "39.7%（OECD 14.8%）；76岁以上54.0%", "OECD 2025", "第十五章", "OECD Pensions at a Glance 2025",
     "https://www.oecd.org/en/publications/pensions-at-a-glance-2025-country-notes_8a53ef12-en/korea-republic-of_5cd52913-en.html"),
    ("养老", "国民年金改革", "缴费率9%→13%（2033年）；替代率43%；耗尽推迟至约2065年", "2025年4月", "第十六章", "OECD国别报告",
     "https://www.oecd.org/en/publications/pensions-at-a-glance-2025-country-notes_8a53ef12-en/korea-republic-of_5cd52913-en.html"),
    ("养老", "继续雇佣建议", "2033年继续雇佣义务至65岁", "2025年5月", "第十六章", "韩联社（经社劳委）",
     "https://www.yna.co.kr/view/AKR20250508105751530"),
    ("健康", "自杀死亡率", "29.1/10万（+6.6%）；OECD标准化26.2", "2024年", "第十七章", "韩国政策简报（统计厅）",
     "https://www.korea.kr/briefing/policyBriefingView.do?newsId=156721880"),
    ("健康", "孤独死人数", "3924人（男性81.7%）", "2024年", "第十七章", "韩国保健福祉部",
     "https://www.mohw.go.kr/board.es?act=view&bid=0027&list_no=1488039&mid=a10503000000"),
    ("出海", "化妆品出口", "114亿美元（+12.3%）", "2025年", "第十八章", "KDI经济信息中心（食药处）",
     "https://eiec.kdi.re.kr/policy/materialView.do?num=275786"),
    ("出海", "K-food+出口", "136.2亿美元（+5.1%）", "2025年", "第十八章", "韩国农林畜产食品部",
     "https://www.mafra.go.kr/bbs/home/792/593514/download.do"),
    ("出海", "净买入美股", "326亿美元（创纪录）", "2025年", "第十九章", "东亚日报（预托结算院）",
     "https://www.donga.com/news/Economy/article/all/20251231/133070056/1"),
    ("出海", "外币证券保管额", "2202.6亿美元", "2025年三季度末", "第十九章", "Metro Seoul（预托结算院）",
     "http://pdf.metroseoul.co.kr/article/20251027500093"),
]

CN_COMPARE = [
    ("实际GDP增速", "1.0%（2025）", "5.0%（2025）", "国家统计局2025年统计公报",
     "https://www.stats.gov.cn/zwfwck/sjfb/202602/t20260228_1962662.html"),
    ("出生人口", "25.43万（2025）", "792万（2025）", "国家统计局",
     "https://www.cncaprc.gov.cn/xxllsy/770049.jhtml"),
    ("总人口变化", "自然减少", "-339万（2025）", "国家统计局", "https://www.cncaprc.gov.cn/xxllsy/770049.jhtml"),
    ("65岁以上占比", "20.3%（2025）", "15.9%（2025年末）", "国家统计局", "https://www.cncaprc.gov.cn/xxllsy/770049.jhtml"),
    ("城镇化/首都圈", "首都圈50.8%（2024）", "城镇化率67.89%（2025年末）", "国家统计局",
     "https://www.cncaprc.gov.cn/xxllsy/770049.jhtml"),
    ("青年就业/失业", "15—29岁就业率44.1%、失业率5.4%（2026.8）", "16—24岁（不含在校生）失业率18.9%、25—29岁7.5%（2026.8）",
     "财新网（国家统计局）", "https://economy.caixin.com/2026-09-18/102486206.html"),
    ("高校毕业生", "—", "2026届约1270万", "央视网（教育部）",
     "https://news.cctv.cn/2025/11/20/ARTI0xYbzeyS5Y6Zky3R3VZg251120.shtml"),
    ("考公竞争比", "9级24.3:1（2025）", "国考98:1（2026年度，过审371.8万）", "京报网（国家公务员局）",
     "https://news.bjd.com.cn/2025/11/30/11434817.shtml"),
    ("居民杠杆率", "88.6%（2025年末，BIS）", "59.4%（2025年末，NIFD）", "第一财经（国家金融与发展实验室）",
     "https://www.yicai.com/news/103026298.html"),
    ("居民收入中位数增速", "—", "4.5%（2025，历史低位）", "第一财经", "https://www.yicai.com/news/103026298.html"),
]

CYCLE = [
    ("第四次康波·繁荣", "1948—1966", "汽车、石化、电气化", "1962年起出口导向起飞", "1949年建国，工业化起步"),
    ("第四次康波·衰退/萧条", "1966—1982", "石油危机、滞胀", "重化工业化、财阀成型", "1978年改革开放"),
    ("第五次康波·回升", "1982—1991", "个人电脑、半导体", "三低景气、民主化、奥运", "价格闯关、乡镇企业"),
    ("第五次康波·繁荣", "1991—2007", "互联网、移动通信", "加入OECD、IMF危机、IT崛起", "入世、城镇化、房地产上行"),
    ("第五次康波·衰退", "2008—2015", "全球金融危机", "出口复苏、潜在增速下降", "四万亿、2015年股灾与去库存"),
    ("第五次康波·萧条", "2015—2025", "增长放缓、债务高企", "生育率跌破1、房价2021见顶、2025增长1.0%", "房地产拐点、人口负增长"),
    ("第六次康波·回升（推测）", "2025—2035", "人工智能、新能源、生物技术", "存储芯片超级周期、KOSPI 9000点", "新能源汽车、AI、先进制造"),
]

HEAD_FILL = PatternFill("solid", fgColor="1F3A5F")
HEAD_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10.5)
BODY_FONT = Font(name="微软雅黑", size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_sheet(ws, headers, rows, widths):
    ws.append(headers)
    for row in rows:
        ws.append(list(row))
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for r_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            if r_idx == 1:
                cell.fill = HEAD_FILL
                cell.font = HEAD_FONT
            else:
                cell.font = BODY_FONT
                if r_idx % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="EEF3F8")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "韩国核心数据"
    write_sheet(ws, ["领域", "指标", "数值", "时间", "所属章节", "来源", "链接"], KR_DATA, [8, 26, 40, 14, 12, 30, 60])

    ws2 = wb.create_sheet("中韩对照")
    write_sheet(ws2, ["指标", "韩国", "中国", "中国数据来源", "链接"], CN_COMPARE, [18, 34, 40, 28, 60])

    ws3 = wb.create_sheet("康波周期坐标")
    write_sheet(ws3, ["康波阶段", "时间（参照周金涛框架）", "全球主导技术/事件", "韩国", "中国"], CYCLE, [24, 20, 26, 36, 32])

    ws4 = wb.create_sheet("说明")
    notes = [
        ("说明",),
        ("1. 数值优先取官方一手来源，部分经主流媒体转引，链接为检索时可访问的页面。",),
        ("2. 中韩统计口径不同（如青年失业率），对照用于看方向与量级。",),
        ("3. 康波分期是分析框架，不是确定性规律；第六次康波为推测。",),
        ("4. 数据截至2026年10月。",),
    ]
    for row in notes:
        ws4.append(row)
    ws4.column_dimensions["A"].width = 80
    ws4["A1"].font = Font(name="微软雅黑", bold=True, size=12)

    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"已生成 {OUT.name}：韩国数据 {len(KR_DATA)} 条，中韩对照 {len(CN_COMPARE)} 条")


if __name__ == "__main__":
    main()
