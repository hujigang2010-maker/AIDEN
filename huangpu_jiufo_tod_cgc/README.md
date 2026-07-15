# 广州黄埔区九佛TOD「全球自贸365街区」策划方案 + CGC国际会议厅未来方向

编制方:中国广告协会文创与IP专委会（CGC）· 秘书长 徐超 · 2026/01

## 交付文件(`deliverables/`)

| 文件 | 格式 | 内容 |
| --- | --- | --- |
| 广州黄埔九佛TOD全球自贸365街区_策划方案.pptx | PPT(35页) | 优化商务风:封面、目录、9大板块(概述定位/战略价值/功能布局/招商/出口体系/效益/运营/支持保障/总结)、CGC国际会议厅与全球贸易港未来方向、2张附件数据表、封底 |
| CGC合作意向与授权提资证明申请函.docx | Word | 致对方的正式函件:表达承接CGC国际会议厅与贸易港建设意向 + 申请出具《合作意向/授权(提资)证明》,附证明书建议格式 |
| 广州黄埔九佛TOD全球自贸365街区_附件数据表.xlsx | Excel(3表) | 出口TOP20 / 消费类出口TOP20 / 楼层功能布局 |

## 说明

- PPT 已按商务风格重构(深藏青+金色主题、卡片式版式、统一页眉与分节页),并新增结合两份材料的「CGC国际会议厅与全球贸易港建设」未来方向章节。
- Word 函件用于向对方表达未来业务意向,并申请可用于后续业务的授权/提资证明(内附可直接盖章使用的证明书建议格式)。
- 文案统一维护在 `scripts/content_cgc.py`。

## 重新生成

```bash
pip install python-pptx python-docx openpyxl
cd scripts
python3 generate_ppt_cgc.py
python3 generate_word_cgc.py
python3 generate_excel_cgc.py
```
