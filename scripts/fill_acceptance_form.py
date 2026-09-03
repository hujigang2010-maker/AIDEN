"""填写《绿城·潮鸣外滩-营销验收单(含现场照片)-2》。

以空白「基础营销验收单」为版式，嵌入 2026-05-22 峰会现场绿城/潮鸣露出照片，
首页改为 3×3 图+图注，并增加「权益核验清单」。不与开盘花束验收混用。
参观邀约、朋友圈九宫格等无单独成片的项目如实标注，不拔高。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties
from PIL import Image as PILImage

TEMPLATE = Path("/home/ubuntu/.cursor/projects/workspace/uploads/____b1b1.xlsx")
OUT = Path("/workspace/deliverables/绿城·潮鸣外滩-营销验收单(含现场照片)-2.xlsx")
OUT_ALIAS = Path("/workspace/deliverables/绿城·潮鸣外滩-营销验收单(含现场照片).xlsx")
PHOTO_DIR = Path("/workspace/deliverables/验收照片")
TMP = Path("/tmp/acceptance-photos")
THUMB = Path("/tmp/acceptance-thumbs-v2")
ZIP_OUT = Path("/workspace/deliverables/下载包/绿城潮鸣外滩-营销验收单(含现场照片)-2.zip")

ALBUM = "https://live.pailixiang.com/album/a12523836160"
VIDEO = "https://pan.baidu.com/s/1S8cNdqnCR_anuVPD6vlhjg"
VIDEO_CODE = "8888"

FONT = "等线"
TITLE_FONT = "黑体"
THIN = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")
LEFT_TOP = Alignment(wrap_text=True, vertical="top", horizontal="left")
GREEN = PatternFill("solid", fgColor="1F6B4A")
GREEN_LIGHT = PatternFill("solid", fgColor="E8F5E9")
GREEN_MID = PatternFill("solid", fgColor="C8E6C9")
AMBER = PatternFill("solid", fgColor="FFF8E1")
WHITE = PatternFill("solid", fgColor="FFFFFF")
WHITE_FONT = Font(name=TITLE_FONT, size=16, bold=True, color="FFFFFF")

# 首页 9 张：优先「绿城中国 / 潮鸣 / 外滩潮鸣」字样可辨的画面
COVER = [
    ("060_941A7785.JPG", "①议程看板【绿城·潮鸣】晚宴冠名", "议程看板：晚宴场次【绿城·潮鸣】星耀北外滩——AI领袖定制晚宴"),
    ("065_941A7794.JPG", "②1号位展台：绿城礼盒+潮鸣立牌", "1号位展台：绿城礼盒、「潮鸣」立牌、VR 与项目物料"),
    ("001_941A7663.JPG", "③主背景板：晚宴冠名·绿城中国", "主背景板全幅：晚宴冠名战略合作伙伴·绿城中国"),
    ("074_941A7832.JPG", "④双屏：战略合作「绿城·外滩潮鸣」", "主会场双屏底部：战略合作伙伴「绿城·外滩潮鸣」"),
    ("071_941A7821.JPG", "⑤主持：晚宴冠名及战略合作伙伴", "主持环节大屏：「晚宴冠名及战略合作伙伴：绿城中国」"),
    ("055_941A7777.JPG", "⑥主背景板合影核验冠名位", "主背景板合影：晚宴冠名位「绿城中国」可核验"),
    ("388_941A8841.JPG", "⑦晚宴桌卡「潮鳴」", "晚宴桌卡露出「潮鳴」，江景厅圆桌"),
    ("395_941A8859.JPG", "⑧祝酒大屏「绿城·外滩潮鸣」", "晚宴祝酒：弧形大屏露出「绿城·外滩潮鸣」"),
    ("00_banner.jpg", "⑨主视觉KV：战略合作含绿城中国", "主视觉 KV：战略合作伙伴名单含绿城中国"),
]

# 露出页续图（含场地与晚宴氛围；不把无字样的画面写成潮鸣看板）
EXTRA = [
    ("075_941A7835.JPG", "讲台/大屏顶栏合作 logo 带（主办致辞）"),
    ("073_941A7829.JPG", "座席手拎袋物料（嘉宾席）"),
    ("077_941A7839.JPG", "主办致辞画面：大屏顶栏合作 logo"),
    ("056_941A7779.JPG", "主背景板合影（晚宴冠名位特写）"),
    ("p066_941A7798.JPG", "主背景板合影核验晚宴冠名位"),
    ("v04_941A8762.JPG", "晚宴现场举杯"),
    ("v01_941A8878.JPG", "晚宴桌次互动"),
    ("399_941A8881.JPG", "晚宴桌次物料"),
    ("p144_941A8048.JPG", "主会场全景（北外滩·一滴水）"),
    ("p080_941A7864.JPG", "主论坛致辞（讲台合作伙伴 logo 带）"),
    ("394_941A8857.JPG", "晚宴合影（江景厅）"),
    ("068_941A7802.JPG", "场地江景露台（北外滩一滴水，正对陆家嘴）"),
]

ALIASES = {
    "395_941A8859.JPG": ["v02_941A8859.JPG", "395_941A8859.JPG"],
    "001_941A7663.JPG": ["p000_941A7663.JPG", "001_941A7663.JPG"],
    "065_941A7794.JPG": ["p064_941A7794.JPG", "065_941A7794.JPG"],
}


def resolve_src(name: str) -> Path:
    names = ALIASES.get(name, [name])
    dirs = [PHOTO_DIR, TMP / "greentown", TMP]
    for n in names:
        for d in dirs:
            p = d / n
            if p.exists() and p.stat().st_size > 2000:
                return p
    raise FileNotFoundError(name)


def thumb(src: Path, dest: Path, max_w=520, max_h=390, quality=84) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    im = PILImage.open(src)
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    im.thumbnail((max_w, max_h), PILImage.Resampling.LANCZOS)
    dest = dest.with_suffix(".jpg")
    im.save(dest, "JPEG", quality=quality, optimize=True)
    return dest


def place_image(ws, path: Path, cell: str, width: int, height: int) -> None:
    img = XLImage(str(path))
    img.width = width
    img.height = height
    img.anchor = cell
    ws.add_image(img)


def fill_cell(ws, coord, text, size=10, bold=False, align=WRAP, fill=None, font_name=FONT):
    cell = ws[coord]
    cell.value = text
    cell.font = Font(name=font_name, size=size, bold=bold)
    cell.alignment = align
    cell.border = THIN
    if fill is not None:
        cell.fill = fill


def paint_range(ws, start_row, end_row, start_col, end_col, fill=WHITE):
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(r, c)
            cell.border = THIN
            if fill is not None:
                cell.fill = fill


def unmerge_if(ws, coord: str) -> None:
    hits = [str(rng) for rng in list(ws.merged_cells.ranges) if coord in rng]
    for rng in hits:
        ws.unmerge_cells(rng)


def setup_print(ws, area: str, landscape: bool = False, fit_height: int = 1) -> None:
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = area
    ws.page_setup.horizontalCentered = True
    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.6, bottom=0.6, header=0.2, footer=0.2
    )


def fill_cover(ws) -> None:
    ws.title = "营销验收单"
    fill_cell(ws, "B3", "晚宴冠名战略合作伙伴\n（绿城中国 | 绿城·潮鸣外滩）", 11, True)
    fill_cell(ws, "D3", "人民币 100,000 元整\n（含税 · 四模块整体打包）", 11, True)
    fill_cell(
        ws,
        "B4",
        "主办：上海市杨浦区科技企业联合会\n赞助：上海绿城泓盛建设发展有限公司\n项目品牌：绿城中国 · 绿城·潮鸣外滩",
        10,
    )
    fill_cell(ws, "D4", "2026年5月22日\n13:00 – 20:30\n上海·北外滩·一滴水", 11, True)
    fill_cell(
        ws,
        "B5",
        "【活动】第四届 / 2026人工智能商业化落地与硬核投资破局峰会。\n"
        "【身份】晚宴冠名战略合作伙伴（物料亦作「绿城·外滩潮鸣」「潮鳴」）。\n"
        "【晚宴】【绿城·潮鸣】星耀北外滩——AI领袖定制晚宴。\n"
        "【四模块】①晚宴冠名和专属权益 ¥55,000；②主会场权益 ¥30,000；"
        "③专场项目参观邀约 ¥5,000；④宣发配合 ¥10,000。含税合计 ¥100,000，整体打包验收。\n"
        "【现场已核验】主背景板「晚宴冠名战略合作伙伴·绿城中国」、议程看板晚宴冠名、"
        "1号位展台、主会场双屏「绿城·外滩潮鸣」、晚宴大屏/桌卡「潮鳴」。\n"
        "【证书】事项载明颁发“2026年度智慧人居新质资产领军企业 暨卓越战略合作伙伴”；"
        "本相册未见证书特写，不以照片推定已颁。",
        8,
        align=LEFT_TOP,
    )
    ws.row_dimensions[5].height = 118

    fill_cell(
        ws,
        "B13",
        "【露出】主背景板、议程看板、1号位展台、双屏、晚宴大屏/桌卡已核验，"
        "绿城中国 / 潮鸣 / 外滩潮鸣字样可辨（见本页图①–⑨及「绿城·潮鸣外滩露出」页）。\n"
        "【营销评估】非常好（☐）　良好（☑）　一般（☐）　较差（☐）\n"
        "【打包验收】四模块整体打包；参观邀约无单独成片、朋友圈九宫格/回顾视频未附本表，"
        "见「权益核验清单」黄底项。效果良好，符合约定要求（以已核验露出为准）。",
        9,
        align=LEFT_TOP,
    )
    ws.row_dimensions[13].height = 48
    ws.row_dimensions[14].height = 36
    ws.row_dimensions[15].height = 28

    fill_cell(
        ws,
        "B16",
        "①本页 9 张均为绿城/潮鸣可辨露出，含图注；"
        "②「权益核验清单」按报价单 4 模块对照照片；"
        "③「绿城·潮鸣外滩露出」页共 21 张；"
        f"④拍立享直播 {ALBUM}（相册约 425 张）；"
        f"⑤全程录像 {VIDEO} 提取码 {VIDEO_CODE}；"
        "⑥费用清单（经办人签字）。本表不与 2026年3月开盘花束验收混用。",
        8,
        align=LEFT,
    )
    ws.row_dimensions[16].height = 72
    fill_cell(ws, "C17", "胡继刚\n13262607888\n（主办执行）\n签字：____________", 10, align=LEFT)
    fill_cell(
        ws,
        "C18",
        "绿城对接：笪浩 18678408669\n营销负责人签字：____________\n日期：____________",
        10,
        align=LEFT,
    )

    unmerge_if(ws, "B6")
    paint_range(ws, 6, 12, 2, 4, WHITE)
    # 3×3：6/8/10 放图，7/9/11 图注，12 总注
    ws.merge_cells("B12:D12")
    fill_cell(
        ws,
        "B12",
        "图①–⑨对应「权益核验清单」；更多露出见「绿城·潮鸣外滩露出」页。江景露台等无品牌字样画面不作为潮鸣看板证据。",
        8,
        align=LEFT,
    )

    photo_rows = (6, 8, 10)
    cap_rows = (7, 9, 11)
    cols = ("B", "C", "D")
    for i, (name, short, _full) in enumerate(COVER):
        r, c = divmod(i, 3)
        fill_cell(ws, f"{cols[c]}{cap_rows[r]}", short, 7, align=LEFT)
        t = thumb(resolve_src(name), THUMB / f"cover_{name}", max_w=420, max_h=300)
        place_image(ws, t, f"{cols[c]}{photo_rows[r]}", 168, 108)

    for r, h in [(6, 86), (7, 32), (8, 86), (9, 32), (10, 86), (11, 32), (12, 22)]:
        ws.row_dimensions[r].height = h

    ws["A20"] = "照片直播（拍立享）"
    ws["A20"].font = Font(name=FONT, size=9, bold=True)
    ws["B20"] = ALBUM
    ws["B20"].hyperlink = ALBUM
    ws["B20"].font = Font(name=FONT, size=9, color="0563C1", underline="single")
    ws.merge_cells("B20:D20")
    ws["A21"] = "全程录像（百度网盘）"
    ws["A21"].font = Font(name=FONT, size=9, bold=True)
    ws["B21"] = f"{VIDEO}  提取码：{VIDEO_CODE}"
    ws["B21"].hyperlink = VIDEO
    ws["B21"].font = Font(name=FONT, size=9, color="0563C1", underline="single")
    ws.merge_cells("B21:D21")

    setup_print(ws, "A2:D18", landscape=False, fit_height=1)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_properties.tabColor = "1F6B4A"


def add_rights_sheet(wb) -> None:
    ws = wb.create_sheet("权益核验清单", 1)
    widths = [6, 18, 36, 12, 42, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:F1")
    fill_cell(ws, "A1", "绿城·潮鸣外滩 · 晚宴冠名四模块权益核验清单", 16, True, WRAP, GREEN, TITLE_FONT)
    paint_range(ws, 1, 1, 1, 6, GREEN)
    ws["A1"].font = WHITE_FONT
    ws["A1"].fill = GREEN
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:F2")
    fill_cell(
        ws,
        "A2",
        "活动：2026-05-22 上海·北外滩·一滴水　身份：晚宴冠名战略合作伙伴　"
        "金额：含税 ¥100,000（报价单 4 模块整体打包，不拆子项验收）　"
        "依据：绿城中国-潮鸣外滩-10万元晚宴冠名服务报价单",
        9,
        False,
        LEFT,
        GREEN_MID,
    )
    paint_range(ws, 2, 2, 1, 6, GREEN_MID)
    ws.row_dimensions[2].height = 32

    headers = ["序号", "服务模块", "约定内容", "金额（元）", "现场照片核验", "结论"]
    for col, title in enumerate(headers, 1):
        cell = ws.cell(3, col, title)
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = GREEN
        cell.alignment = WRAP
        cell.border = THIN
    ws.row_dimensions[3].height = 22

    rows = [
        (
            "1",
            "晚宴冠名和专属权益",
            "晚宴主题冠名、专场PPT/视频、主持人口播、席卡等物料植入",
            "55,000",
            "图①议程看板晚宴冠名【绿城·潮鸣】；图③⑤⑥主背景板「晚宴冠名战略合作伙伴·绿城中国」；"
            "图⑦桌卡「潮鳴」；图⑧祝酒大屏「绿城·外滩潮鸣」。专场PPT宣讲无单独成片，以录像为准。",
            "通过（物料/冠名）",
        ),
        (
            "2",
            "主会场权益",
            "1号位展台、手拎袋项目物料、视频轮播/奖项颁发画面",
            "30,000",
            "图②1号位展台（绿城礼盒+潮鸣立牌+VR）；图④双屏战略合作伙伴「绿城·外滩潮鸣」；"
            "图⑨主视觉KV含绿城中国；露出页手拎袋/讲台顶栏。奖项颁发证书未见特写。",
            "通过（展台/双屏）",
        ),
        (
            "3",
            "专场项目参观邀约",
            "论坛/颁奖结束后主持人口播 + 动线引导，邀约前往案场",
            "5,000",
            "图⑤证明主会场主持人口播位已具备。相册抽样未见「参观邀约/案场接驳」专项成片，"
            "不以本表照片推定口播已执行，请以全程录像及现场执行回执为准。",
            "待录像核验",
        ),
        (
            "4",
            "宣发配合",
            "现场摄影、朋友圈九宫格、回顾视频项目鸣谢、执行回执",
            "10,000",
            "现场摄影：本表 21 张露出 + 拍立享约 425 张 + 网盘录像。本表即执行回执（图片+合影）。"
            "朋友圈九宫格、回顾视频片头鸣谢未附，不以照片推定已发。",
            "摄影通过；其余待补",
        ),
        (
            "—",
            "合计（含税）",
            "4 模块整体打包，一次性验收",
            "100,000",
            "已核验项以绿城/潮鸣字样可辨露出为准；黄底项需录像或另行回执，不在本表拔高为已完成。",
            "部分待补",
        ),
        (
            "—",
            "证书",
            "2026年度智慧人居新质资产领军企业 暨卓越战略合作伙伴",
            "—",
            "事项描述已载明颁发。本相册未见证书/颁奖特写，不作为已交付证据。",
            "待补照片",
        ),
    ]
    fills = {
        "通过（物料/冠名）": GREEN_LIGHT,
        "通过（展台/双屏）": GREEN_LIGHT,
        "待录像核验": AMBER,
        "摄影通过；其余待补": AMBER,
        "部分待补": AMBER,
        "待补照片": AMBER,
    }
    for i, vals in enumerate(rows, 4):
        result = vals[-1]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(i, col, val)
            cell.font = Font(name=FONT, size=9, bold=(col in (2, 4, 6)))
            cell.alignment = LEFT if col in (3, 5) else WRAP
            cell.border = THIN
            cell.fill = fills.get(result, WHITE) if col == 6 else WHITE
        ws.row_dimensions[i].height = 72

    note = 10
    ws.merge_cells(f"A{note}:F{note}")
    fill_cell(
        ws,
        f"A{note}",
        "验收口径：报价单约定「不逐项贴照片、四模块整体打包」。本清单把照片映射到模块，便于绿城复核，"
        "不等于把未成片的子项改成已交付。营销评估预勾「良好」，请绿城营销负责人复核签字。"
        "本清单仅对应 2026-05-22 峰会，不与 3 月开盘花束验收混用。",
        9,
        False,
        LEFT,
        GREEN_LIGHT,
    )
    paint_range(ws, note, note, 1, 6, GREEN_LIGHT)
    ws.row_dimensions[note].height = 48

    setup_print(ws, f"A1:F{note}", landscape=True, fit_height=1)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FFF8E1"


def add_photo_sheet(wb) -> None:
    ws = wb.create_sheet("绿城·潮鸣外滩露出")
    ws.merge_cells("A1:D1")
    fill_cell(
        ws,
        "A1",
        "绿城中国｜绿城·潮鸣外滩（外滩潮鸣 / 潮鳴）现场露出验收照片",
        14,
        True,
        WRAP,
        GREEN,
        TITLE_FONT,
    )
    paint_range(ws, 1, 1, 1, 4, GREEN)
    ws["A1"].font = WHITE_FONT
    ws["A1"].fill = GREEN
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:D2")
    fill_cell(
        ws,
        "A2",
        f"活动时间：2026-05-22　地点：上海·北外滩·一滴水　"
        f"本页 {len(COVER) + len(EXTRA)} 张（首页 9 张 + 续图）　"
        f"完整相册约 425 张：{ALBUM}　录像：{VIDEO}（提取码 {VIDEO_CODE}）",
        9,
        False,
        LEFT,
        GREEN_MID,
    )
    paint_range(ws, 2, 2, 1, 4, GREEN_MID)
    ws["A2"].hyperlink = ALBUM
    ws.row_dimensions[2].height = 36
    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 42

    all_items = [(n, full) for n, _s, full in COVER] + EXTRA
    row = 4
    col_pair = [("A", "B"), ("C", "D")]
    i = 0
    while i < len(all_items):
        ws.row_dimensions[row].height = 168
        ws.row_dimensions[row + 1].height = 36
        for j, (img_col, cap_col) in enumerate(col_pair):
            if i + j >= len(all_items):
                break
            name, cap = all_items[i + j]
            t = thumb(resolve_src(name), THUMB / f"sheet_{name}", max_w=640, max_h=480)
            place_image(ws, t, f"{img_col}{row}", 290, 205)
            cell = ws[f"{img_col}{row + 1}"]
            ws.merge_cells(f"{img_col}{row + 1}:{cap_col}{row + 1}")
            cell.value = f"{i + j + 1}. {cap}"
            cell.font = Font(name=FONT, size=9)
            cell.alignment = LEFT
            cell.border = THIN
        i += 2
        row += 2

    note_row = row + 1
    ws.merge_cells(f"A{note_row}:D{note_row}")
    fill_cell(
        ws,
        f"A{note_row}",
        "说明：优先收录绿城中国、绿城·潮鸣外滩、绿城·外滩潮鸣、潮鳴可辨露出。"
        "江景露台等无品牌字样画面仅证明场地，不作为潮鸣看板证据。"
        "完整 425 张以拍立享为准；MP4 以网盘「5.22」为准。不与开盘花束 4 张混用。",
        9,
        False,
        LEFT,
        WHITE,
    )
    paint_range(ws, note_row, note_row, 1, 4, WHITE)
    ws.row_dimensions[note_row].height = 42
    setup_print(ws, f"A1:D{note_row}", landscape=True, fit_height=0)
    ws.print_title_rows = "1:2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "1F6B4A"


def make_zip() -> None:
    ZIP_OUT.parent.mkdir(parents=True, exist_ok=True)
    used = [n for n, *_ in COVER] + [n for n, _ in EXTRA]
    with zipfile.ZipFile(ZIP_OUT, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(OUT, OUT.name)
        for name in used:
            src = resolve_src(name)
            zf.write(src, f"验收照片-露出精选/{name}")


def main() -> None:
    if not TEMPLATE.exists():
        raise SystemExit(f"缺少模板 {TEMPLATE}")
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    THUMB.mkdir(parents=True, exist_ok=True)

    used = [n for n, *_ in COVER] + [n for n, _ in EXTRA]
    for name in used:
        src = resolve_src(name)
        dest = PHOTO_DIR / name
        if not dest.exists():
            dest.write_bytes(src.read_bytes())

    wb = load_workbook(str(TEMPLATE))
    fill_cover(wb.active)
    add_rights_sheet(wb)
    add_photo_sheet(wb)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(OUT))
    wb.save(str(OUT_ALIAS))
    make_zip()
    print(f"saved {OUT} size={OUT.stat().st_size}")
    print(f"alias {OUT_ALIAS} size={OUT_ALIAS.stat().st_size}")
    print(f"zip {ZIP_OUT} size={ZIP_OUT.stat().st_size}")


if __name__ == "__main__":
    main()
