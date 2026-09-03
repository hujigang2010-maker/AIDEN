"""把《绿城中国策划活动类效果评估表-开盘花束》文字与现场照片填入营销验收单。"""
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

SRC = Path("/home/ubuntu/.cursor/projects/workspace/uploads/____feef.xlsx")
if not SRC.exists():
    SRC = Path("/home/ubuntu/.cursor/projects/workspace/uploads/____b1b1.xlsx")
IMG_SRC = Path("/tmp/flower-eval")
OUT = Path("/workspace/deliverables/绿城·潮鸣外滩-开盘花束营销验收单.xlsx")
PHOTO_DIR = Path("/workspace/deliverables/开盘花束验收照片")
THUMB = Path("/tmp/flower-thumbs")

FONT = "等线"
THIN = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")

PHOTOS = [
    ("image1.jpeg", "案场大厅：绿城中国标识 +「潮鸣」背景，花束装箱陈列"),
    ("image2.jpeg", "开盘现场氛围：礼仪接待、花束阵列与拍摄设备"),
    ("image3.jpeg", "花束送达：厢式货车卸货（纸箱内红粉花束）"),
    ("image4.jpeg", "花束清点转运：鲜切花轻拿轻放，沪A·ZR373 送达现场"),
]


def thumb(src: Path, dest: Path, max_w=640, max_h=480, quality=85) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    im = PILImage.open(src)
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    im.thumbnail((max_w, max_h), PILImage.Resampling.LANCZOS)
    dest = dest.with_suffix(".jpg")
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
        raise SystemExit(f"missing template {SRC}")
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    THUMB.mkdir(parents=True, exist_ok=True)

    for name, _ in PHOTOS:
        src = IMG_SRC / name
        if not src.exists():
            raise SystemExit(f"missing image {src}")
        (PHOTO_DIR / name).write_bytes(src.read_bytes())

    wb = load_workbook(str(SRC))
    ws = wb.active
    ws.title = "营销验收单"

    fill_cell(ws, "B3", "绿城潮鳴开盘花束定制\n（潮鸣外滩开盘 / 潮鳴东方示范区开放）", 11, True)
    fill_cell(ws, "D3", "45 束鲜花\n（金额按约定）", 12, True)
    fill_cell(
        ws,
        "B4",
        "项目：绿城中国 · 绿城·潮鸣外滩\n（潮鳴东方项目示范区）",
        10,
    )
    fill_cell(ws, "D4", "2026年3月\n示范区开放：3月26日", 11, True)
    fill_cell(
        ws,
        "B5",
        "依据《绿城中国策划活动类效果评估表》（冠名、活动赞助、活动执行等）："
        "本次活动为潮鳴东方项目示范区开放（3月26日），为潮鸣外滩开盘活动提供 45 束鲜花。"
        "按照约定完成本次活动的整体策划与执行，活动取得圆满成功。",
        9,
        align=LEFT,
    )
    fill_cell(
        ws,
        "B13",
        "总体活动评价：现场氛围良好，花束按时送达并完成陈列，绿城中国 /「潮鸣」品牌露出清晰。"
        "策划负责人评估为活动取得圆满成功；效果良好，符合要求。"
        "营销负责人勾选栏见评估表原件（非常好 / 良好 / 一般 / 较差）。",
        10,
        align=LEFT,
    )
    fill_cell(
        ws,
        "B16",
        "①本表嵌入评估表现场照片 4 张，详见「开盘花束现场」页；"
        "②附件：绿城中国策划活动类效果评估表-开盘花束；"
        "③费用清单（经办人签字）。",
        9,
        align=LEFT,
    )
    fill_cell(ws, "C17", "策划负责人签名：\n（见评估表原件）", 10)
    fill_cell(ws, "C18", "营销负责人签名：\n（见评估表原件）", 10)

    # 2×2 照片
    ws.row_dimensions[6].height = 110
    ws.row_dimensions[7].height = 10
    ws.row_dimensions[8].height = 18
    ws.row_dimensions[9].height = 10
    ws.row_dimensions[10].height = 110
    ws.row_dimensions[11].height = 18
    ws.row_dimensions[12].height = 18
    ws.row_dimensions[16].height = 68

    anchors = ["B6", "C6", "B10", "C10"]
    for (name, _cap), cell in zip(PHOTOS, anchors):
        t = thumb(IMG_SRC / name, THUMB / f"main_{name}")
        place_image(ws, t, cell, 220, 148)

    ws["A20"] = "对应评估表"
    ws["A20"].font = Font(name=FONT, size=9, bold=True)
    ws["B20"] = "绿城中国策划活动类效果评估表（冠名、活动赞助、活动执行等）· 开盘花束"
    ws["B20"].font = Font(name=FONT, size=9)

    # 第二页：大图 + 原文
    ws2 = wb.create_sheet("开盘花束现场")
    ws2["A1"] = "绿城潮鳴开盘花束定制｜活动现场氛围（评估表附图）"
    ws2["A1"].font = Font(name="黑体", size=14, bold=True)
    ws2.merge_cells("A1:D1")
    ws2.row_dimensions[1].height = 28
    ws2["A2"] = (
        "活动主题：绿城潮鳴开盘花束定制　　举办时间：2026年3月（示范区开放 3月26日）　　"
        "交付：为潮鸣外滩开盘活动提供 45 束鲜花。按约定完成整体策划与执行，活动取得圆满成功。"
    )
    ws2["A2"].font = Font(name=FONT, size=10)
    ws2["A2"].alignment = LEFT
    ws2.merge_cells("A2:D2")
    ws2.row_dimensions[2].height = 42
    for col, w in enumerate([42, 42, 42, 42], 1):
        ws2.column_dimensions[get_column_letter(col)].width = w

    row = 4
    for i, (name, cap) in enumerate(PHOTOS):
        col = "A" if i % 2 == 0 else "C"
        if i % 2 == 0:
            ws2.row_dimensions[row].height = 170
            ws2.row_dimensions[row + 1].height = 36
        t = thumb(IMG_SRC / name, THUMB / f"page_{name}", max_w=720, max_h=540)
        place_image(ws2, t, f"{col}{row}", 300, 210)
        cap_cell = ws2[f"{col}{row + 1}"]
        merge_end = "B" if col == "A" else "D"
        ws2.merge_cells(f"{col}{row + 1}:{merge_end}{row + 1}")
        cap_cell.value = f"{i + 1}. {cap}"
        cap_cell.font = Font(name=FONT, size=9)
        cap_cell.alignment = LEFT
        if i % 2 == 1:
            row += 2

    note = row + 1 if i % 2 == 1 else row + 2
    # after 4 photos (0,1 then 2,3) row ends at 8
    ws2["A8"] = (
        "策划负责人填写｜总体活动评价：本次活动为潮鳴东方项目示范区开放（3月26日）。"
        "为潮鸣外滩开盘活动提供了 45 束鲜花。按照约定进行本次活动的整体策划与执行，活动取得圆满成功。"
    )
    ws2["A8"].alignment = LEFT
    ws2["A8"].font = Font(name=FONT, size=10)
    ws2.merge_cells("A8:D8")
    ws2.row_dimensions[8].height = 48
    ws2["A9"] = "营销负责人填写｜总体效果评估：非常好（  ）  良好（  ）  一般（  ）  较差（  ）　　营销负责人签名：__________"
    ws2["A9"].font = Font(name=FONT, size=10)
    ws2.merge_cells("A9:D9")
    ws2.row_dimensions[9].height = 28
    ws2.page_setup.orientation = "landscape"
    ws2.page_setup.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(OUT))
    print(f"saved {OUT} size={OUT.stat().st_size}")


if __name__ == "__main__":
    main()
