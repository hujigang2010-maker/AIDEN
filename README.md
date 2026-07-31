# AIDEN

AIDEN 是一个多任务工作区。`main` 分支存放共享配置；具体应用和文档交付物位于各自的 `cursor/*` 功能分支。

## 显示语言

本仓库默认使用**简体中文**。首次打开项目时，请安装推荐的中文语言包，并在命令面板中选择 **Configure Display Language → 中文(简体)**，然后重启 Cursor。

更多说明见 [AGENTS.md](./AGENTS.md)。

## 本分支：九三学社申请人社登记表（胡继刚）

依据《九三学社申请人社登记表》（2020年版）栏目，结合申请人简历整理填表稿。人名统一为**胡继刚**。

### 生成

```bash
python3 scripts/generate_jiusan_application.py
```

### 产出

- `output/九三学社申请人社登记表-胡继刚.docx`
