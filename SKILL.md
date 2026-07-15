---
name: financial-disclosure-analysis
description: Download, extract, verify, and analyze official financial disclosures from A-share and Hong Kong-listed companies. Use when the user needs source-traceable financial data from annual reports, interim reports, results announcements, or other filings, as well as peer comparisons, period analysis, Excel workpapers, or editable PPTX reports.
---

# Financial Disclosure Analysis

将官方财务披露转化为可追溯数据底稿，并将最终报告交由 PPT Master 生成原生可编辑 `.pptx`。

## 依赖

安装本项目依赖：

```bash
pip install -r requirements.txt
playwright install chromium
```

另行安装 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) 及其依赖。设置 `PPT_MASTER_HOME` 指向仓库根目录或 `skills/ppt-master`，或在调用桥接脚本时传 `--ppt-master-dir`。

开始任务前执行：

```bash
python scripts/ppt_master_bridge.py check
```

如果检查失败，停止报告阶段并给出安装指引；数据下载、提取和 Excel 生成仍可继续。

## 核心原则

- 只从官方披露渠道下载并保留原始 PDF。
- 不修改原始 PDF/TXT；所有派生文件写入新的输出目录。
- 使每个数据点可在 30 秒内定位回原文。
- 不用估算值补齐缺失数据，不编造图表、KPI 或分析结论中的数字。
- 将 `数据底稿.xlsx` 作为数值事实源，将年报原文作为披露事实源。
- 将 PPT Master 的 `svg_output/` 作为页面设计源，将 `exports/*.pptx` 作为正式报告。
- 不再生成 Chart.js HTML 报告；除非用户明确要求网页附件，否则 HTML 不属于交付物。

## 工作流

### 1. 确认范围

确认目标公司、证券代码、市场、期间、报表类型、输出目录和分析深度。默认：

- 输出目录：`<当前目录>/finance_analysis/data`
- 分析深度：标准
- 报告格式：16:9 可编辑 PPTX
- 标准报告：10–15 页；简要报告：5–8 页；深度报告：15 页以上

一次性请用户确认范围，然后自动执行到 PPT Master 的强制设计确认门。

### 2. 下载与转文本

令 `SD` 指向本 Skill 的 `scripts/`。

A 股：

```bash
python {SD}/a_share.py search-annual --codes {代码...} --years {年份...} -d "{output_dir}/pdfs" --quiet
```

港股：

```bash
python {SD}/hk_share.py search-annual --codes "{代码,...}" --years "{年份,...}" -d "{output_dir}/pdfs" --quiet
```

批量转文本：

```bash
python {SD}/pdf_to_text.py --input-dir "{output_dir}/pdfs" --output-dir "{output_dir}/txt" --skip-existing --quiet
```

需要排错时才启用详细日志。不要把下载明细、解析警告或 TXT 正文刷入主会话。

### 3. 提取、复核与 Excel

读取 [references/data-extraction.md](references/data-extraction.md) 并严格执行。完成后必须得到：

- `{output_dir}/tsv/*.tsv`
- `{output_dir}/数据底稿.xlsx`
- 缺失项、口径差异、交叉验证异常的摘要

在对话中展示简洁的“指标 × 公司-期间”Markdown 汇总表。数据单元格只写数字，单位单列，缺失填 `-`。

### 4. 形成报告素材

根据 Excel 和已核验原文编写 `{output_dir}/报告素材.md`。读取 [references/ppt-master-report.md](references/ppt-master-report.md) 获取内容契约、页面建议和质量门。

`报告素材.md` 至少包含：

1. 分析范围、期间、单位与口径
2. 执行摘要和关键结论
3. 报告大纲与逐页信息目标
4. 用于每页的已核验数据表
5. 趋势、同行差异、驱动因素和风险分析
6. 数据局限、缺失项与来源定位

先核验每个结论都能回指 Excel 或原文，再进入 PPT Master。

### 5. 交接给 PPT Master

创建 PPT Master 项目并导入 `报告素材.md` 和 `数据底稿.xlsx`：

```bash
python {SD}/ppt_master_bridge.py --ppt-master-dir "{ppt_master_dir}" prepare \
  --project-name "{项目名}" \
  --projects-dir "{output_dir}/ppt-master-projects" \
  --source "{output_dir}/报告素材.md" \
  --source "{output_dir}/数据底稿.xlsx" \
  --format ppt169
```

桥接脚本会复制后再按 PPT Master 契约使用 `import-sources --move`，不会移动 Financial Disclosure Analysis 原件。
它同时在项目的 `analysis/financial-disclosure-analysis-handoff.json` 中记录输入文件 SHA-256 和可用的 PPT Master Git commit。

随后读取解析出的 PPT Master `SKILL.md`、`workflows/routing.md` 和其要求的引用文件，按上游当前版本完整执行。不得把 PPT Master 当作普通 Python 库绕过其 Strategist、逐页 SVG 创作、质量检查、`finalize_svg.py` 和 `svg_to_pptx.py` 流程。

在 PPT Master Step 4 的 ⛔ BLOCKING 设计确认处停止，向用户展示建议并等待明确确认。确认后连续完成非阻塞阶段。

### 6. 验收与交付

至少验收：

- PPT Master 项目校验通过。
- `svg_output/` 页数与确认后的大纲一致。
- 每页通过 SVG 质量检查，无溢出、截断、重叠或不可读字号。
- 图表和 KPI 与 `数据底稿.xlsx` 抽样核对；关键结论逐条核对。
- `analysis/financial-disclosure-analysis-handoff.json` 中的输入哈希与交付时文件一致。
- 正式文件位于 PPT Master 项目的 `exports/`，为可在 PowerPoint 中逐元素编辑的 `.pptx`。
- `svg_final/` 仅作预览和排错，不冒充正式 PPTX 源。
- Excel 底稿与 PPTX 一并交付；需要 PDF 时从最终 PPTX 另行导出并做分页检查。

最终消息仅列出交付路径、报告页数、数据缺失/局限和验证结果，不粘贴原始披露正文。

## 脚本

| 脚本 | 用途 |
|---|---|
| `a_share.py` | 搜索、下载 A 股年报 |
| `hk_share.py` | 搜索、下载港股年报 |
| `pdf_to_text.py` | 批量 PDF 转 TXT |
| `generate_excel.py` | 从 TSV 生成 Excel 底稿 |
| `ppt_master_bridge.py` | 定位 PPT Master、初始化项目并安全导入素材 |

## 故障边界

- 官方渠道下载失败：列出公告、期间、失败原因和重试情况。
- PDF 无法解析：保留 PDF，标记需人工/OCR处理，不猜测数值。
- PPT Master 未安装：完成数据和 Excel，停止 PPT 报告阶段并给出明确安装命令。
- PPTX 导出失败：保留 PPT Master 项目及 `svg_output/`，按其 failure-recovery 流程处理，不回退为 HTML 正式报告。
- 涉及未公开资料、个人信息或内部敏感数据：提醒用户确认脱敏和共享范围。
