# Finance Skill

从 A股（巨潮资讯网）和港股（港交所披露易）自动下载年报 PDF，提取财务指标，生成 Excel 数据底稿和 HTML 分析报告。

## 快速开始

```bash
pip install -r requirements.txt
playwright install chromium

# 下载万科 + 保利 2024 年报
python scripts/a_share.py search-annual --codes 000002 600048 --year 2024 --download-dir ./data/pdfs
python scripts/pdf_to_text.py --input-dir ./data/pdfs --output-dir ./data/txt --skip-existing
# 后续流程由 AI Agent 驱动，详见 SKILL.md
```

## 如何使用

- **作为 Skill**：放到 `.claude/skills/finance/`，Claude Code 自动发现并加载
- **手动调用**：上面脚本可独立使用，完整流程参考 [SKILL.md](SKILL.md)

## License

MIT
