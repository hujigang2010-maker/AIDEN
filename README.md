# 冠松静安智能驾驶研发中心 · 招商方案（GS · iDrive Hub）

> 项目代号：**GS · iDrive Hub**
> 战略定位：**让中心城区跑通智能驾驶最后一公里**
> 文档库：本仓库 `docs/` 目录，按 5 个 Phase / 8 个任务交付
> 当前版本：v1.0（招商策划阶段；签约前由各业务部门按实际数据复核）

---

## 商务汇报 PPT

- [docs/deck/GS-iDrive-Hub-招商方案.pptx](docs/deck/GS-iDrive-Hub-招商方案.pptx) — **47 页 16:9 商务汇报版**（v1.2.1 新增"服务平台佣金构成"与"后市场协同分成构成"2 页可视化解读）（封面 / 议程 / 摘要 / 区位 / SWOT / 5 阶段 8 任务 + 5 张链主一页纸 + 服务包 + 报价单 + 风险矩阵 + **谈判策略 4 页** + **冠松资源协同** + **22 人组织架构** + **薪酬带宽** + **KPI 仪表盘** + **FAQ** + **文档索引** / 投决建议 / 致谢）
- 生成脚本：[scripts/build_pptx.py](scripts/build_pptx.py)（依赖 `python-pptx`，模块化函数 `anchor_one_pager()` / `section_cover()` 便于二次扩页）
- 中文字体：WenQuanYi Micro Hei（系统默认）；如需替换为 PingFang/Source Han，编辑脚本顶部 `CN_FONT` 即可
- 重新生成：`pip install python-pptx && python3 scripts/build_pptx.py`

## 业态转型顾问材料（招什么）

- [docs/advisory/00-业态转型建议.md](docs/advisory/00-业态转型建议.md) — 3 类业态 + 8 大案例
- [docs/advisory/deck/冠松01楼-业态转型顾问材料.pptx](docs/advisory/deck/冠松01楼-业态转型顾问材料.pptx) — 18 页（给董事长看「招什么」）
- 他到底要什么（8 月 5 日得到大脑口径）：[docs/advisory/10-他想要什么.md](docs/advisory/10-他想要什么.md)
- 招商怎么落地、怎么付钱：见下一节 **招商委托**

## 9 月 2 日之后（当前主文件）

> 2026-09-02 沟通已转向：**政府公关为主、招商为辅**。董事长要公寓/酒店路径，不信纯写字楼招商。9 月 30 日必须进场施工。
> 收费以见面为准：启动金 **15 万覆盖头 90 天** + 之后月度 **8 万** + 散户 **8%** / 整层 **12%**，启动金抵佣金。孵化器品牌授权另计。
> **不要再把下面「招商委托会上包」当主菜**——那是转向前的工具箱。

- [docs/advisory/12-9月2日沟通-需求与新方向.md](docs/advisory/12-9月2日沟通-需求与新方向.md) — 对方要什么、谈成的新方向（内部）
- [docs/advisory/13-9月2日后-优化方案.md](docs/advisory/13-9月2日后-优化方案.md) — **交给甲方的优化方案**
- [docs/advisory/14-静安研发用地业态调整-案例与政策窗口.md](docs/advisory/14-静安研发用地业态调整-案例与政策窗口.md) — 静安本地窗口（诚实距离）
- [docs/advisory/deck/冠松01楼-政策路径-董事会版.pptx](docs/advisory/deck/冠松01楼-政策路径-董事会版.pptx) — **8 页董事会 PPT**
- [docs/legal/06-业态政策路径-董事会简版.docx](docs/legal/06-业态政策路径-董事会简版.docx) — 纸质简版（确认栏）
- 生成脚本：[`scripts/build_policy_pptx.py`](scripts/build_policy_pptx.py) · [`scripts/build_policy_brief.py`](scripts/build_policy_brief.py)

## 招商委托（转向前的工具箱 · 收费已被 9 月 2 日见面覆盖）

> 商务简版 05b + 会上版 6 页仍可当招商附件。收费请以见面为准（整层 12%、15 万覆盖头 90 天），不要按 v1.2「一律 8%、第 2 个月起付月度」对外讲。

- [docs/legal/05b-招商合作协议-商务简版.docx](docs/legal/05b-招商合作协议-商务简版.docx) — **会上主文件**（90 天动作 + 收费确认栏）
- [docs/advisory/deck/冠松01楼-招商合作-会上版.pptx](docs/advisory/deck/冠松01楼-招商合作-会上版.pptx) — **6 页会上版**
- [docs/advisory/deck/冠松01楼-招商合作-工作版.pptx](docs/advisory/deck/冠松01楼-招商合作-工作版.pptx) — **12 页工作版（给你深化）**
- [docs/advisory/11-90天行动与收费对齐.md](docs/advisory/11-90天行动与收费对齐.md) — 90 天排期 + 收费原口径柔化
- [docs/legal/05-合作协议-招商服务委托协议.docx](docs/legal/05-合作协议-招商服务委托协议.docx) — 法务全稿（口径已与 v1.2 对齐）
- 生成脚本：[`scripts/build_mandate_brief.py`](scripts/build_mandate_brief.py) · [`scripts/build_meeting_pptx.py`](scripts/build_meeting_pptx.py) · [`scripts/build_working_pptx.py`](scripts/build_working_pptx.py) · [`scripts/build_mandate_agreement.py`](scripts/build_mandate_agreement.py)
- 给董事长的一页纸：[docs/advisory/09-招商委托协议要点.md](docs/advisory/09-招商委托协议要点.md)

## 法务文档（Word · 草案口径）

> 01–04 由 [`scripts/build_legal_docs.py`](scripts/build_legal_docs.py) 生成（园区对外）；05 见上一节。签约前由法务最终定稿。

- [docs/legal/01-合作协议-链主总部租赁合同.docx](docs/legal/01-合作协议-链主总部租赁合同.docx) — 链主总部租赁合同（甲档样张，17 章 + 5 附件）
- [docs/legal/02-合作协议-中介居间服务协议.docx](docs/legal/02-合作协议-中介居间服务协议.docx) — 5 家中介通用居间协议
- [docs/legal/03-合作协议-联合实验室共建协议.docx](docs/legal/03-合作协议-联合实验室共建协议.docx) — 3F 联合研发实验室共建（链主 + 我方）
- [docs/legal/04-合作协议-政府专班合作备忘录.docx](docs/legal/04-合作协议-政府专班合作备忘录.docx) — 静安区政府 × 冠松集团 战略 MOU
- [docs/legal/05b-招商合作协议-商务简版.docx](docs/legal/05b-招商合作协议-商务简版.docx) — **会上确认收费**（两种签法 + 确认栏）
- [docs/legal/05-合作协议-招商服务委托协议.docx](docs/legal/05-合作协议-招商服务委托协议.docx) — **冠松委托操盘方招商**（法务全稿）

## 财务测算 Excel

> 由 [`scripts/build_finance_xlsx.py`](scripts/build_finance_xlsx.py) 一键生成；含 8 个 Sheet。

- [docs/finance/财务测算与商务模型.xlsx](docs/finance/财务测算与商务模型.xlsx)
  - Sheet 1：摘要 Dashboard
  - Sheet 2：假设与参数（黄色单元格可调）
  - Sheet 3：三年损益（含柱状图）
  - Sheet 4：36 个月滚动现金流（含折线图）
  - Sheet 5：入驻爬坡进度（链主 + 生态）
  - Sheet 6：单/双变量敏感性矩阵
  - Sheet 7：22 人薪酬带宽
  - Sheet 8：链主谈判让步阶梯计算器

## 目录索引

### Phase 1 · 策略与定位（奠基）

| 任务 | 文档 | 主要交付 |
| --- | --- | --- |
| 任务 1 · 产业定位研究与竞品对标 | [docs/phase1-strategy/01-industry-positioning-report.md](docs/phase1-strategy/01-industry-positioning-report.md) | 1 页摘要 + 产业链图谱 + 静安/跨区竞品对标 + SWOT |
| 任务 2 · 空间功能规划 | [docs/phase1-strategy/02-space-planning.md](docs/phase1-strategy/02-space-planning.md) | A~E 栋功能 + 户外测试区方案 + 平面示意 + 面积表 |

### Phase 2 · 招商执行（核心）

| 任务 | 文档 | 主要交付 |
| --- | --- | --- |
| 任务 3 · 链主企业攻坚 | [docs/phase2-execution/03-anchor-tenant-tracker.md](docs/phase2-execution/03-anchor-tenant-tracker.md) | TOP5 进度表 + 每家一页纸定制提案（华为 / 百度 / 小鹏 / 地平线 / Momenta） |
| 任务 3 附录 · 谈判策略 | [docs/phase2-execution/03b-negotiation-playbook.md](docs/phase2-execution/03b-negotiation-playbook.md) | 链主/政府/中介/生态 四类谈判 Playbook + 让步阶梯 + 24 条话术 + 红线 |
| 任务 4 · 生态企业招商漏斗 | [docs/phase2-execution/04-ecosystem-funnel.md](docs/phase2-execution/04-ecosystem-funnel.md) | 300 家库结构 + 漏斗 + 中介策略 |
|  | [docs/phase2-execution/04-ecosystem-target-db.csv](docs/phase2-execution/04-ecosystem-target-db.csv) | 300 家目标企业数据库样表 |
|  | [docs/phase2-execution/04-broker-agreement-template.md](docs/phase2-execution/04-broker-agreement-template.md) | 中介合作协议模板（要点版） |
| 任务 5 · 政府关系对接 | [docs/phase2-execution/05-government-relations.md](docs/phase2-execution/05-government-relations.md) | 对接路径 + 纪要模板 + 政策适配清单（牌照/补贴/人才/财税/数据/用地） |

### Phase 3 · 品牌与活动

| 任务 | 文档 | 主要交付 |
| --- | --- | --- |
| 任务 6 · 品牌活动策划与执行 | [docs/phase3-brand/06-launch-and-events.md](docs/phase3-brand/06-launch-and-events.md) | 9 月发布会 + 年度 10 场活动日历 + 媒体策略 + 80 家媒体清单 |

### Phase 4 · 商业条款

| 任务 | 文档 | 主要交付 |
| --- | --- | --- |
| 任务 7 · 商业模式与合同 | [docs/phase4-commercial/07-pricing-and-contract.md](docs/phase4-commercial/07-pricing-and-contract.md) | 四档报价 + 合同框架 + 三年财务测算 + 链主报价单样张 |
|  | [docs/phase4-commercial/07-financial-model.csv](docs/phase4-commercial/07-financial-model.csv) | 三年财务测算明细 (CSV) |
| 任务 7 补充 · 非租金收入解读 | [docs/phase4-commercial/07c-non-rental-revenue-explained.md](docs/phase4-commercial/07c-non-rental-revenue-explained.md) | **服务平台佣金 + 后市场协同分成** 产业逻辑、6+6 子项构成、合规边界、投决会备答（Markdown 版） |
|  | [docs/phase4-commercial/07c-非租金收入解读-投决会备答版.docx](docs/phase4-commercial/07c-非租金收入解读-投决会备答版.docx) | 同上内容 · Word 投决会备答版（含表格 + 应答口径） |
|  | PPT 第 30–31 页 | 服务平台佣金 600 万 + 后市场协同 300 万 · 可视化构成图（含三年爬坡 + 4 道护城河） |

### Phase 5 · 落地推进

| 任务 | 文档 | 主要交付 |
| --- | --- | --- |
| 任务 8 · 12 个月执行计划 | [docs/phase5-rollout/08-execution-plan.md](docs/phase5-rollout/08-execution-plan.md) | 里程碑甘特图 + 4 人核心团队 + 22 人扩编 + RACI 分工 + 运营 SOP |
| 任务 8 附录 · 22 人 JD | [docs/phase5-rollout/08b-team-and-jd.md](docs/phase5-rollout/08b-team-and-jd.md) | 22 个岗位详细 JD + 薪酬带宽 + KPI + 招聘节奏 + 期权激励机制 |

---

## 顶层逻辑（One-Pager · 给决策人）

- **项目载体**：**01# 新建研发楼**（永和社区 075b-07 地块）· 地上 9F **15,152.75 ㎡** + 地下 6,992.87 ㎡ · 高 44.95 m · **C6 教育科研用地** · 装配式 100% / 绿建二星
- **客群垂直分层**：链主总部（8–9F）→ 核心研发（3–4F · 5.4–5.7 m 高层高）→ 算法软件（6–7F）
- **差异化壁垒**：① 中心城区罕见 5.4–5.7 m 高层高硬科技研发空间；② 绿建二星 + 装配式 + 540 ㎡ 光伏 + 海绵城市 全套绿色低碳成绩单；③ 静安"一企一策"政策包；④ "三段式"测试解决方案（园区静态 + 区内 1.5 km 路测延伸 + 嘉定/临港会员通道，**轻资产**）；⑤ 冠松汽车后市场反哺
- **TOP5 链主**：华为车 BU / 百度 Apollo / 小鹏 / 地平线 / Momenta（落位 8–9F）
- **300 家生态库**：5 家中介 + 9 类来源；线索→签约 17% 转化模型
- **9 月发布会**：200 人 · 5+ 家签约 · "iDrive · 静安 10 条"政策包发布
- **四档商业**：甲(链主总部+政策返还) / 乙(核心研发高层高) / 丙(算法软件) / 丁(冠名+对赌)
- **三年目标**：入驻率 92% · 链主 ≥ 1 · 入驻企业 12–18 家 · Y3 EBITDA 转正约 1,400 万元

---

## 使用建议

- 决策汇报：建议从 README → Phase1 任务1 摘要 → Phase4 报价单 → Phase5 甘特图，约 30 分钟可形成完整认知
- 销售工具：Phase2 链主一页纸 + 报价单 + 客户报备表
- 政府汇报：Phase2 任务5 政策适配清单 + Phase4 一企一策模板
- 内部管理：Phase5 RACI + KPI 看板 + 漏斗例会节奏

## 审阅与签字（建议）

| 角色 | 审阅范围 |
| --- | --- |
| 集团董事长 | 全部，重点 Phase4 + Phase5 |
| 项目总监 | Phase2 + Phase3 + Phase5 |
| GR 总监 | Phase2 任务5 + Phase4 政策返还 |
| 招商总监 | Phase2 任务3、4 |
| 法务总监 | Phase4 合同 / 对赌 / 数据合规 |
| CFO | Phase4 财务模型 |
