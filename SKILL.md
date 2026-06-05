---
description: 下载、提取并分析公司财务报表。支持A股（巨潮资讯网）和港股（港交所披露易），自动下载年报PDF、转TXT、提取指标、生成Excel数据底稿和HTML分析报告。
---

# Finance Skill — 财务报表分析与提取

## 适用平台

- **Claude Code**（原生）：克隆到 `.claude/skills/finance/` 自动发现
- **其他 Agent**：将本文件作为 system prompt 注入，`scripts/` 加入 PATH 或 `cd` 到本目录执行

## 前置依赖

```bash
pip install -r requirements.txt
playwright install chromium
```

## 脚本清单

所有脚本位于 `<本skill目录>/scripts/` 下，Python 3.10+ 可直接运行：

| 脚本 | 用途 |
|------|------|
| `a_share.py` | A股年报搜索与PDF下载（巨潮资讯网） |
| `hk_share.py` | 港股年报搜索与PDF下载（港交所披露易） |
| `pdf_to_text.py` | 批量PDF转TXT（多进程） |
| `generate_excel.py` | 从TSV生成Excel数据底稿 |

---

## 阶段1 — 确认范围

**1a — 确认分析目标**：当用户未明确指定时，逐一询问确认：目标公司候选、时间范围、报表类型、输出目录。默认输出目录为当前工作目录下的 `finance_analysis/data`。

**1b — 确认**：汇总目标公司、时间范围、报表类型、输出目录，一次性请用户确认。这是唯一一轮确认，后续阶段自动推进。

---

## 阶段2 — 全量下载与提取

下载和PDF转txt是确定性脚本调用，由主线程直接执行。

### 阶段2a — 批量搜索并下载PDF（按市场分组）

用变量 `SD` 指向本 skill 的 `scripts/` 目录。

- **A股**：`python {SD}/a_share.py search-annual --codes {代码1} {代码2} ... --year {年份} --download-dir "{output_dir}/pdfs"`
- **港股**：`python {SD}/hk_share.py search-annual --codes "{代码1},{代码2},..." --year {年份} --download-dir "{output_dir}/pdfs"`

脚本内部复用浏览器、自动处理限流。输出每家公司搜索和下载结果。

### 阶段2b — 批量转PDF为TXT（一条命令）

```bash
python {SD}/pdf_to_text.py --input-dir "{output_dir}/pdfs" --output-dir "{output_dir}/txt" --skip-existing
```

这条命令使用多进程并行处理所有PDF，无需逐文件调用。完成后列出每个txt的行数和首页内容片段。

### 阶段2c — 制定指标清单

用 Read 工具只读每个txt的前50-100行（目录/摘要部分），了解实际报表结构。制定统一指标清单，分"报表科目"和"经营披露指标"两组，确定数据单位。无需等用户确认，直接进入阶段3。

---

## 阶段3 — 口径统一与产出

### 阶段3设计原则

本阶段产出的数据将被后续分析反复使用，需遵循两个原则：

1. **可溯源**：每个数据点的备注中应包含足够信息，让人能在30秒内翻到对应PDF位置找到原文。
2. **可分析**：Excel输出的主表是一张展平的数据表，导入pandas或透视表后无需任何布局调整即可开始分析。不要把数据拆成多个Sheet做"排版"——那是Word的活，不是Excel的。

### 阶段3a-3b — 子Agent提取（并行，批量分组）

将公司按 2-3 家一组分批，每批启动一个子Agent。通常启动 3-4 个子Agent即可，复杂场景可适当增减。

使用 `Agent` 工具并行启动，subagent_type 用 `general-purpose`。

**子Agent Prompt 模板**：

```
你是财务数据提取助手。从多家公司的报表txt中按指定指标清单提取数值，分别为每家公司输出TSV文件。

=== 任务清单 ===
{逐项列出：公司名、股票代码、年份、txt路径、输出TSV路径}

=== 指标清单 ===
{从阶段2c确定的完整指标清单，包括指标名、查找说明、优先级}

=== TSV格式规范 ===
header: 名称 \t 数据 \t 备注

字段说明：
- 名称: 统一指标名
- 数据: 统一单位后的值，数值+单位写在一起（如"2,435亿""18.5%""3.21元/股"）；找不到填"-"
- 备注: 由你自行决定详细程度，可包含原始科目名、数据来源路径、提取依据、口径说明等

=== 提取要求 ===
1. 对每家公司逐一读取txt文件，找到指标清单中每个指标对应的数值
2. 同一指标在不同公司的单位须统一（亿元、%、元/股等），单位不统一的自行换算后再填入
3. 找不到的指标金额填"-"，备注简述原因
4. 每家公司+期间组合写入一个独立的TSV文件
5. 优先定位到财务报表章节（合并资产负债表、合并利润表、合并现金流量表），不要从头到尾通读

完成后只报告以下内容（总计<500字）：
- 每家公司：实际提取到的指标数 / 指标总数
- 缺失项清单
- 发现的口径差异（科目名称差异、单位差异等）

严禁返回TSV文件内容。
```

**并行执行**：在单个消息中同时启动所有批次的子Agent调用。

### 阶段3c — 处理口径差异

汇总各子Agent报告的口径差异，自行判断并统一：
- 同一指标在不同公司的科目名称映射
- 发现口径差异较大无法自行判断时再向用户确认，否则直接更新TSV并继续。

### 阶段3d — 复核

读取各公司TSV文件，逐公司复核数据一致性：
- 交叉验证（如资产=负债+权益的勾稽关系）
- 如涉及多年份，检查同比变化是否合理
- 发现异常自行纠正，汇总缺失项和口径差异。

### 阶段3e — 产出Excel

编写 JSON 配置文件，然后调用预置脚本生成 **一个** Excel 文件（`数据底稿.xlsx`），第一个 Sheet 为汇总透视表，后续为各公司明细。

**Step 1** — 创建 `{output_dir}/excel_config.json`：

```json
{
  "companies": [
    {
      "name": "万科",
      "periods": { "2024": "{output_dir}/tsv/万科_2024.tsv", "2025": "{output_dir}/tsv/万科_2025.tsv" }
    },
    {
      "name": "金茂",
      "periods": { "2024": "{output_dir}/tsv/金茂_2024.tsv", "2025": "{output_dir}/tsv/金茂_2025.tsv" }
    }
  ],
  "indicators": ["营业收入", "营业成本", "毛利", "毛利率", ...],
  "output": "{output_dir}/数据底稿.xlsx"
}
```

- `companies`：按用户指定的公司顺序排列。`name` 为公司简称，`periods` 为期间标签→TSV路径的映射（key 可以是年度"2024"、季度"202403"、半年度"2024H1"等任意字符串）
- `indicators`：指标显示顺序，先绝对值后比率；未在清单中但 TSV 里存在的指标会自动追加到末尾

**Step 2** — 调用脚本：

```bash
python {SD}/generate_excel.py --config "{output_dir}/excel_config.json"
```

脚本自动生成 `数据底稿.xlsx`，结构：Sheet 1 = 汇总透视表（指标 × 公司-期间），后续 Sheet = 各公司各期间明细（名称 / 数据 / 备注）。

---

## 阶段4 — 分析报告

数据提取和 Excel 产出完成后进入本阶段，依次产出：汇总表 → 报告大纲 → HTML 报告。

### 阶段4a — 输出 Markdown 汇总表

在对话中输出一张 **指标 × 公司-期间** 交叉透视表，严格按以下模板：

```
|               | 单位 | 万科       | 万科       | 金茂       | 金茂       |
|               |      | 2024年     | 2025年     | 2024年     | 2025年     |
|---------------|------|------------|------------|------------|------------|
| 投资性房地产   | 亿   | 1,200      | 1,350      | 880        | 920        |
| 总负债         | 亿   | 8,500      | 9,200      | 6,100      | 6,800      |
| 毛利率         | %    | 16.0       | 14.2       | 18.5       | 15.5       |
```

规则：
1. **列结构**：第1列指标名，第2列单位，其余按"公司A YYYY年 | …"排列。第1行表头为公司简称，第2行为年份
2. **列顺序**：按用户指定的公司顺序，每家公司内期间升序
3. **行顺序**：先绝对值/规模类，再比率/计算类
4. **值格式**：数据单元格只写纯数字（不加单位），保留适当小数位；NA 填 `-`
5. **禁止**：不输出数据来源列、备注列、TSV 原文或任何调试信息

### 阶段4b — 制作分析报告大纲

基于汇总数据，列出分析报告大纲（各级标题即可），内容结构由你根据本次数据和公司特点自行发挥，列出后一次性向用户确认。

### 阶段4c — 生成 HTML 分析报告

分三步完成。

**Step 1 — 生成图表数据**

```bash
python {SD}/generate_charts.py --excel "{output_dir}/数据底稿.xlsx" --output "{output_dir}/charts_data.json"
```

脚本自动从 Excel 透视表提取数据，生成可在 Chart.js 中直接使用的图表配置 JSON。会根据数据自动选择图表类型（柱状图/折线图），覆盖收入利润、毛利率、负债水平、资产结构、现金流等维度。

**Step 2 — 组装 HTML 报告**

使用以下模板骨架，将 `{CHART_JSON}` 和你的分析文字填入对应位置：

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>财务分析报告</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:"Microsoft YaHei","PingFang SC",sans-serif;color:#333;line-height:1.8;max-width:1100px;margin:0 auto;padding:40px 20px;background:#f8f9fa}
  h1{text-align:center;font-size:28px;color:#2F5496;margin-bottom:8px}
  h2{font-size:22px;color:#2F5496;border-left:4px solid #2F5496;padding-left:14px;margin:48px 0 20px}
  h3{font-size:17px;color:#444;margin:24px 0 12px}
  .subtitle{text-align:center;color:#888;font-size:14px;margin-bottom:40px}
  .kpi-row{display:flex;gap:16px;flex-wrap:wrap;margin:20px 0}
  .kpi-card{flex:1;min-width:160px;background:#fff;border-radius:8px;padding:18px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.06)}
  .kpi-card .label{font-size:13px;color:#999;margin-bottom:6px}
  .kpi-card .value{font-size:24px;font-weight:700;color:#2F5496}
  .kpi-card .change{font-size:13px;color:#888;margin-top:4px}
  .chart-box{background:#fff;border-radius:8px;padding:24px;margin:24px 0;box-shadow:0 2px 8px rgba(0,0,0,0.06)}
  .chart-box canvas{max-height:380px}
  p{font-size:15px;margin:10px 0;text-indent:2em}
  table{width:100%;border-collapse:collapse;margin:20px 0;font-size:14px}
  th{background:#2F5496;color:#fff;padding:10px 12px;text-align:center}
  td{padding:8px 12px;text-align:center;border-bottom:1px solid #e8e8e8}
  tr:nth-child(even) td{background:#f5f7fa}
  .risk{background:#fff3f3;border-left:4px solid #dc4e4e;padding:14px 18px;margin:16px 0;border-radius:0 8px 8px 0;font-size:14px}
  @media print{body{background:#fff;padding:20px}.chart-box{break-inside:avoid}}
</style>
</head>
<body>

<h1>财务分析报告</h1>
<p class="subtitle">覆盖公司：{公司列表}　|　期间：{期间列表}　|　生成日期：{日期}</p>

<!-- ════════ 以下由你按大纲撰写 ════════ -->

<h2>一、公司概况与业务结构</h2>
<div class="kpi-row">
  <div class="kpi-card"><div class="label">营业收入(亿)</div><div class="value">...</div></div>
  <div class="kpi-card"><div class="label">归母净利润(亿)</div><div class="value">...</div></div>
  <div class="kpi-card"><div class="label">总资产(亿)</div><div class="value">...</div></div>
</div>
<p>{分析段落}</p>

<h2>二、盈利能力分析</h2>
<div class="chart-box"><canvas id="chart_revenue"></canvas></div>
<p>{分析段落}</p>

<div class="chart-box"><canvas id="chart_margin"></canvas></div>
<p>{分析段落}</p>

<h2>三、偿债能力与资本结构</h2>
<div class="chart-box"><canvas id="chart_leverage"></canvas></div>
<p>{分析段落}</p>

<h2>四、资产质量</h2>
<div class="chart-box"><canvas id="chart_assetStructure"></canvas></div>
<p>{分析段落}</p>

<h2>五、现金流状况</h2>
<div class="chart-box"><canvas id="chart_cashflow"></canvas></div>
<p>{分析段落}</p>

<h2>六、同业对比与综合评价</h2>
<table>
  <thead><tr><th>指标</th>{各公司期间列}</tr></thead>
  <tbody>{数据行}</tbody>
</table>
<p>{综合评价段落}</p>

<div class="risk"><strong>⚠ 风险提示：</strong>{从数据中识别的风险点}</div>

<!-- ════════ 图表初始化脚本（不要修改） ════════ -->
<script>
const CHART_DATA = {将 charts_data.json 的内容粘贴在此};

const CHART_COLORS = ['#2F5496','#ED7D31','#449E73','#DC4E4E','#8C64B4','#00AAB4'];

function makeChart(id, cfg) {
  const canvas = document.getElementById(id);
  if (!canvas || !cfg) return;
  if (cfg.datasets) {
    cfg.datasets.forEach((ds, i) => {
      if (!ds.backgroundColor) ds.backgroundColor = CHART_COLORS[i % CHART_COLORS.length];
    });
  }
  new Chart(canvas, cfg);
}

document.querySelectorAll('canvas[id^="chart_"]').forEach(canvas => {
  const key = canvas.id.replace('chart_', '');
  const cfg = CHART_DATA.charts[key];
  if (cfg) makeChart(canvas.id, {type: cfg.type, data: {labels: cfg.labels, datasets: cfg.datasets}, options: cfg.options});
});
</script>

</body>
</html>
```

**Step 3 — 写入文件**

将组装好的 HTML 写入 `{output_dir}/分析报告.html`。

关键约束：
- `CHART_DATA` 直接粘贴 `charts_data.json` 的内容，**不修改任何数字**
- KPI 卡片和表格中的数值也必须来自 Excel，禁止编造
- 如果 `charts_data.json` 中某个图表缺失（如"现金流"数据不全），对应 `<canvas>` 自动跳过，安全
- 分析段落是你唯一的创作空间

---

## 规则

- 只从官方披露渠道下载，原始PDF必须保留
- 每个数据点可追溯到报告具体位置
- 子Agent绝不返回原始文件内容（txt正文、TSV数据），只返回路径和摘要
- 子Agent遇到无法解决的问题时标记清楚，不编造数据
