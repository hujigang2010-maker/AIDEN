# 赛普客户：房企“十五五”竞争的破局关键——客户服务体系化建设

以赛普咨询（赛普研究院）视角撰写的长篇行业研究报告，完全参照赛普咨询公开报告的版式规范与品牌视觉（官方 logo、赛普红 `#e94d4f`、深藏青 `#1f2a44`），报告时间纹路锚定 **2025年10月**（党的二十届四中全会审议通过“十五五”规划《建议》、赛普咨询公众号同名原创文发布之时），面向“十五五”（2026—2030）作前瞻规划。

## 交付物

`report/` 目录下：

- `赛普客户-房企十五五竞争的破局关键-客户服务体系化建设.docx`（Word 正稿，33 页）
- `赛普客户-房企十五五竞争的破局关键-客户服务体系化建设.pdf`（PDF 转换版）

报告结构：封面 / 版权与免责声明 / 摘要 / 目录（页码已回填）/ 引言 / 第一章 时代之变 / 第二章 认知之困 / 第三章 破局之道（“14554”客户服务体系蓝图）/ 第四章 价值之证 / 第五章 落地之径 / 结语 / 附录A 政策脉络 / 附录B 蓝图数字索引 / 附录C 赛普方法论演进 / 关于赛普。

## 脚本复现

```bash
pip install python-docx matplotlib cairosvg Pillow pypdf
# 需系统安装 LibreOffice（soffice）与 Noto CJK 字体（fonts-noto-cjk）

python3 scripts/generate_charts.py                 # 1. 生成 assets/ 下全部配图
python3 scripts/generate_report.py                 # 2. 第一遍生成 docx（目录无页码）
soffice --headless --convert-to pdf --outdir report/ report/*.docx
python3 scripts/extract_toc.py                     # 3. 从 PDF 提取章节页码 → scripts/toc.json
python3 scripts/generate_report.py --tocmap scripts/toc.json   # 4. 第二遍回填目录页码
soffice --headless --convert-to pdf --outdir report/ report/*.docx
```

## 文件说明

- `scripts/report_content.py`：报告全部文字内容（数据层）
- `scripts/generate_report.py`：python-docx 排版渲染器（封面、页眉页脚、目录、表格样式）
- `scripts/generate_charts.py`：matplotlib 绘制 10 张品牌风格图表
- `scripts/extract_toc.py`：目录页码两遍回填工具
- `assets/`：赛普官方 logo（提取自 chinasap.cn 官网内联 SVG 渲染）与报告配图

## 时间纹路说明

报告撰写视角为 2025年10月下旬：政策脉络引用至 2025年10月23日四中全会及10月28日《建议》全文发布为止；“当前”“今年”均指 2025 年；2026—2030 年内容均为规划前瞻表述。
