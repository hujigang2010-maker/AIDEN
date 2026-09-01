"""填写绿城·潮鸣外滩营销验收单，并嵌入 2026-05-22 现场照片。"""
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from PIL import Image as PILImage

SRC = Path("/home/ubuntu/.cursor/projects/workspace/uploads/____b1b1.xlsx")
OUT = Path("/workspace/deliverables/绿城·潮鸣外滩-营销验收单(含现场照片).xlsx")
PHOTO_DIR = Path("/workspace/deliverables/验收照片")
TMP = Path("/tmp/acceptance-photos")
THUMB = Path("/tmp/acceptance-thumbs")

ALBUM = "https://live.pailixiang.com/album/a12523836160"
VIDEO = "https://pan.baidu.com/s/1S8cNdqnCR_anuVPD6vlhjg"
VIDEO_CODE = "8888"

FONT = "等线"
THIN = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")


def thumb(src: Path, dest: Path, max_w=480, max_h=360, quality=82) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    im = PILImage.open(src)
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    im.thumbnail((max_w, max_h), PILImage.Resampling.LANCZOS)
    im.save(dest, "JPEG", quality=quality, optimize=True)
    return dest


def place_image(ws, path: Path, cell: str, width: int, height: int):
    img = XLImage(str(path))
    img.width = width
    img.height = height
    img.anchor = cell
    ws.add_image(img)


def fill_cell(ws, coord, text, size=10, bold=False, align=WRAP):
    cell = ws[coord]
    cell.value = text
    cell.font = Font(name=FONT, size=size, bold=bold)
    cell.alignment = align
    cell.border = THIN


def main():
    if not SRC.exists():
        raise SystemExit(f"missing template: {SRC}")
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    THUMB.mkdir(parents=True, exist_ok=True)

    # 主表 6 张 + 附图页
    main_photos = [
        ("p000_941A7663.JPG", "主背景板：晚宴冠名战略合作伙伴·绿城中国露出"),
        ("p064_941A7794.JPG", "绿城·潮鸣外滩现场展位与物料"),
        ("p080_941A7864.JPG", "主论坛：主办方代表致辞"),
        ("p144_941A8048.JPG", "主会场全景（北外滩·一滴水）"),
        ("v02_941A8859.JPG", "晚宴祝酒：大屏露出绿城中国·潮鸣外滩"),
        ("v04_941A8762.JPG", "晚宴现场：宾客举杯与驻场演出"),
    ]
    extra_photos = [
        ("00_banner.jpg", "活动主视觉 KV（战略合作伙伴含绿城中国）"),
        ("p066_941A7798.JPG", "主背景板合影核验晚宴冠名位"),
        ("p093_941A7905.JPG", "主论坛听众席与双屏演讲"),
        ("p128_941A7978.JPG", "主旨演讲现场"),
        ("v01_941A8878.JPG", "晚宴桌次互动与举杯"),
        ("v05_941A8600.JPG", "主题演讲（晚宴前衔接场次）"),
        ("p048_941A7767.JPG", "现场展区交流"),
        ("p013_941A7690.JPG", "现场物料陈列"),
    ]

    copied = []
    for name, cap in main_photos + extra_photos:
        src = TMP / name
        if not src.exists():
            raise SystemExit(f"missing photo: {src}")
        dest = PHOTO_DIR / name
        dest.write_bytes(src.read_bytes())
        copied.append((dest, cap))

    wb = load_workbook(str(SRC))
    ws = wb.active
    ws.title = "营销验收单"

    fill_cell(ws, "B3", "晚宴冠名战略合作伙伴\n（绿城中国 | 绿城·潮鸣外滩）", 11, True)
    fill_cell(ws, "D3", "人民币 100,000 元整\n（含税）", 12, True)
    fill_cell(
        ws,
        "B4",
        "主办：上海市杨浦区科技企业联合会\n赞助：上海绿城泓盛建设发展有限公司",
        10,
    )
    fill_cell(ws, "D4", "2026年5月22日\n13:00 – 20:30", 11, True)
    fill_cell(
        ws,
        "B5",
        "「2026人工智能商业化落地与硬核投资破局峰会」于上海·北外滩·一滴水举办。"
        "绿城·潮鸣外滩作为晚宴冠名战略合作伙伴，现场交付：晚宴冠名及专属权益、主会场露出、"
        "专场项目参观邀约、宣发配合；共享现场展位；颁发“2026年度智慧人居新质资产领军企业 暨卓越战略合作伙伴”证书。",
        9,
        align=LEFT,
    )
    fill_cell(
        ws,
        "B13",
        "现场已完成主背景板/讲台晚宴冠名露出、绿城·潮鸣外滩展位接待、主论坛及晚宴执行。"
        "拍立享相册归档 425 张；全程录像见百度网盘「5.22」。效果良好，符合约定要求。",
        10,
        align=LEFT,
    )
    fill_cell(
        ws,
        "B16",
        "①本表嵌入现场照片，详见「现场照片」页；"
        "②照片直播：live.pailixiang.com/album/a12523836160；"
        "③全程录像（百度网盘 5.22）提取码 8888；"
        "④费用清单（经办人签字）。",
        9,
        align=LEFT,
    )
    fill_cell(ws, "C17", "胡继刚\n13262607888\n（主办执行）", 10)
    fill_cell(ws, "C18", "（绿城营销负责人签字）", 10)

    # 验收照片区加高，便于嵌图
    ws.row_dimensions[6].height = 118
    for r in range(7, 10):
        ws.row_dimensions[r].height = 22
    ws.row_dimensions[10].height = 118
    ws.row_dimensions[11].height = 28
    ws.row_dimensions[12].height = 28
    ws.row_dimensions[16].height = 72

    # 主表 2×3 照片
    anchors = ["B6", "C6", "D6", "B10", "C10", "D10"]
    sizes = [(210, 148)] * 6
    for (src, cap), cell, (w, h) in zip(main_photos, anchors, sizes):
        t = thumb(TMP / src, THUMB / f"main_{src}", max_w=420, max_h=300)
        place_image(ws, t, cell, w, h)

    # 超链接说明行（不改表头结构，写在打印区外的注释行）
    ws["A20"] = "照片直播（拍立享）"
    ws["B20"] = ALBUM
    ws["B20"].hyperlink = ALBUM
    ws["B20"].font = Font(name=FONT, size=9, color="0563C1", underline="single")
    ws["A21"] = "全程录像（百度网盘）"
    ws["B21"] = f"{VIDEO}  提取码：{VIDEO_CODE}"
    ws["B21"].hyperlink = VIDEO
    ws["B21"].font = Font(name=FONT, size=9, color="0563C1", underline="single")
    ws["A20"].font = Font(name=FONT, size=9, bold=True)
    ws["A21"].font = Font(name=FONT, size=9, bold=True)

    # 第二页：现场照片
    ws2 = wb.create_sheet("现场照片")
    ws2["A1"] = "绿城·潮鸣外滩 × 2026人工智能商业化落地峰会｜现场验收照片"
    ws2["A1"].font = Font(name="黑体", size=14, bold=True)
    ws2.merge_cells("A1:D1")
    ws2.row_dimensions[1].height = 28
    ws2["A2"] = (
        f"活动时间：2026-05-22　地点：上海·北外滩·一滴水　"
        f"相册共 425 张　直播：{ALBUM}　录像：{VIDEO}（提取码 {VIDEO_CODE}）"
    )
    ws2["A2"].font = Font(name=FONT, size=9)
    ws2.merge_cells("A2:D2")
    ws2.row_dimensions[2].height = 36
    ws2["A2"].alignment = LEFT
    ws2["A2"].hyperlink = ALBUM

    for col, w in enumerate([42, 42, 42, 42], 1):
        ws2.column_dimensions[get_column_letter(col)].width = w

    # 每行 2 张：图 + 说明 在下一行
    all_items = main_photos + extra_photos
    row = 4
    col_pair = [("A", "B"), ("C", "D")]
    i = 0
    while i < len(all_items):
        ws2.row_dimensions[row].height = 160
        ws2.row_dimensions[row + 1].height = 32
        for j, (img_col, cap_col) in enumerate(col_pair):
            if i + j >= len(all_items):
                break
            name, cap = all_items[i + j]
            t = thumb(TMP / name, THUMB / f"sheet_{name}", max_w=520, max_h=390)
            place_image(ws2, t, f"{img_col}{row}", 280, 198)
            cell = ws2[f"{img_col}{row + 1}"]
            ws2.merge_cells(f"{img_col}{row + 1}:{cap_col}{row + 1}")
            cell.value = f"{i + j + 1}. {cap}"
            cell.font = Font(name=FONT, size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        i += 2
        row += 2

    note_row = row + 1
    ws2[f"A{note_row}"] = (
        "说明：以上为代表性验收照片（主视觉、晚宴冠名背景板、潮鸣外滩展位、主论坛、晚宴大屏露出及举杯现场）。"
        "完整 425 张以拍立享相册为准；MP4 全程录像以百度网盘「5.22」为准。"
    )
    ws2[f"A{note_row}"].alignment = LEFT
    ws2[f"A{note_row}"].font = Font(name=FONT, size=9)
    ws2.merge_cells(f"A{note_row}:D{note_row}")
    ws2.row_dimensions[note_row].height = 40
    ws2.page_setup.orientation = "landscape"
    ws2.page_setup.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0
    ws2.print_title_rows = "1:2"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(OUT))
    print(f"saved {OUT} size={OUT.stat().st_size}")
    print(f"photos copied to {PHOTO_DIR} n={len(list(PHOTO_DIR.glob('*')))}")


if __name__ == "__main__":
    main()
