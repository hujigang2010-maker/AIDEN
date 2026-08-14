# 住房即服务：医疗、养老与生产效率提升的 2030 展望

本分支交付河南大学住房政策研究中心白皮书，从住房与社区空间出发，结合 2025—2026 年政策和最新科技进展，展望 2030 年的医疗服务、养老服务与生产效率提升服务。

## 生成

```bash
python3 -m pip install python-docx matplotlib pillow
python3 scripts/build_charts.py
python3 scripts/build_whitepaper.py
python3 scripts/verify_whitepaper.py
```

输出文件：

- `dist/河南大学住房政策研究中心_住房即服务_医疗养老与提效服务2030展望白皮书.docx`
- `whitepaper/assets/charts/` 配图

资料核对见 `source/资料摘编.md`。
