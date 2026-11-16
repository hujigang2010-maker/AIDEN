"""把“修改对照”和“事实核验清单”从 JSON 生成 Markdown 与 XLSX。

JSON 结构：{"title": ..., "notes": [...], "sheets": [{"name", "intro", "columns", "widths", "rows"}]}
清单内容涉及个人申报事实，不放在公开仓库里；本脚本只负责排版。
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def to_markdown(data: dict) -> str:
    out = [f"# {data['title']}", ""]
    out += [f"- {note}" for note in data.get("notes", [])]
    for sheet in data["sheets"]:
        out += ["", f"## {sheet['name']}", ""]
        if sheet.get("intro"):
            out += [sheet["intro"], ""]
        out.append("| " + " | ".join(sheet["columns"]) + " |")
        out.append("|" + "---|" * len(sheet["columns"]))
        for row in sheet["rows"]:
            cells = [str(cell).replace("|", "\\|").replace("\n", "<br>") for cell in row]
            out.append("| " + " | ".join(cells) + " |")
    return "\n".join(out) + "\n"


def to_xlsx(data: dict, out_path: Path) -> None:
    book = Workbook()
    book.remove(book.active)
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    cover = book.create_sheet("说明")
    cover.column_dimensions["A"].width = 120
    cover["A1"] = data["title"]
    cover["A1"].font = Font(name="宋体", size=14, bold=True)
    for index, note in enumerate(data.get("notes", []), start=3):
        cell = cover.cell(row=index, column=1, value=note)
        cell.font = Font(name="宋体", size=10.5)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for sheet in data["sheets"]:
        ws = book.create_sheet(sheet["name"][:31])
        columns = sheet["columns"]
        for index, title in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=index, value=title)
            cell.font = Font(name="宋体", size=10.5, bold=True)
            cell.fill = PatternFill("solid", fgColor="E8E8E8")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            ws.column_dimensions[get_column_letter(index)].width = sheet["widths"][index - 1]
        for r, row in enumerate(sheet["rows"], start=2):
            if len(row) != len(columns):
                raise ValueError(f"{sheet['name']} 第 {r - 1} 行有 {len(row)} 列，应为 {len(columns)} 列")
            for c, value in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=value)
                cell.font = Font(name="宋体", size=10)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    book.save(out_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", required=True, type=Path)
    parser.add_argument("--out-md", required=True, type=Path)
    parser.add_argument("--out-xlsx", required=True, type=Path)
    args = parser.parse_args()
    data = json.loads(args.data.read_text(encoding="utf-8"))
    args.out_md.write_text(to_markdown(data), encoding="utf-8")
    to_xlsx(data, args.out_xlsx)
    print({sheet["name"]: len(sheet["rows"]) for sheet in data["sheets"]})


if __name__ == "__main__":
    main()
