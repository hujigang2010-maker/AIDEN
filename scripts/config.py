"""全局配置与运行模式判定。

- 若存在有效 API Key（通过 .env 或环境变量注入），脚本走"真实采集模式"。
- 否则自动降级为"演示数据模式"（DEMO），用合成数据把管线端到端跑通。
"""
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" / "yangpu"
CHART_DIR = ROOT / "charts"
DATA_DIR.mkdir(parents=True, exist_ok=True)
CHART_DIR.mkdir(parents=True, exist_ok=True)


def _load_dotenv():
    env_path = Path(__file__).resolve().parent / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


_load_dotenv()

AMAP_KEY = os.getenv("AMAP_KEY", "").strip()
QCC_KEY = os.getenv("QCC_KEY", "").strip()
TARGET_DISTRICT = os.getenv("TARGET_DISTRICT", "杨浦区").strip()
TARGET_CITY = os.getenv("TARGET_CITY", "上海市").strip()

# 任一关键 Key 缺失即进入演示模式
DEMO_MODE = not (AMAP_KEY and QCC_KEY)

# 杨浦区下辖街道/镇
YANGPU_SUBDISTRICTS = [
    "五角场街道", "五角场镇", "新江湾城街道", "长海路街道", "控江路街道",
    "四平路街道", "江浦路街道", "定海路街道", "平凉路街道", "大桥街道",
    "延吉新村街道", "长白新村街道", "殷行街道",
]

# 中文字体（matplotlib）
CJK_FONT = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
