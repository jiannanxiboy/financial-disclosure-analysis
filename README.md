# Finance Skill — 财务报表分析与提取

从 A股（巨潮资讯网）和港股（港交所披露易）自动下载年报 PDF，提取财务指标，生成 Excel 数据底稿和 HTML 分析报告。

## 安装

```bash
git clone https://github.com/<your-username>/finance-skill.git
cd finance-skill
pip install -r requirements.txt
playwright install chromium
```

## 目录结构

```
finance-skill/
├── SKILL.md              # AI Agent 使用说明书
├── README.md             # 人类读的本文档
├── requirements.txt      # Python 依赖
├── scripts/
│   ├── a_share.py        # A股年报搜索下载
│   ├── hk_share.py       # 港股年报搜索下载
│   ├── pdf_to_text.py    # PDF 批量转 TXT
│   ├── generate_excel.py # TSV → Excel 数据底稿
│   ├── generate_charts.py # 图表数据JSON
│   └── _common.py        # 共用工具模块
```

## 手动使用

```bash
# 1. 下载 A 股年报 PDF
python scripts/a_share.py search-annual --codes 000002 600048 --year 2024 --download-dir ./data/pdfs

# 2. 下载港股年报 PDF
python scripts/hk_share.py search-annual --codes "00688,01109" --year 2024 --download-dir ./data/pdfs

# 3. PDF 转 TXT
python scripts/pdf_to_text.py --input-dir ./data/pdfs --output-dir ./data/txt --skip-existing

# 4. 生成 Excel 数据底稿
python scripts/generate_excel.py --config ./data/excel_config.json
```

## AI Agent 使用

将本仓库克隆到 Agent 可访问的路径，将 `SKILL.md` 作为 system prompt 注入即可。

**Claude Code 用户**：放到 `.claude/skills/finance/` 目录，自动发现。

## 输出文件

| 文件 | 说明 |
|------|------|
| `数据底稿.xlsx` | Sheet 1 汇总透视表 + 各公司明细 Sheet，纯数字可 SUM |
| `分析报告.html` | 结构化 HTML 报告，浏览器直接打开 |

## 支持的市场

- **A股**：通过巨潮资讯网 (cninfo.com.cn) 搜索下载
- **港股**：通过港交所披露易 (hkexnews.hk) 搜索下载

## License

MIT
