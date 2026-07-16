# Financial Disclosure Analysis

**下载年报、提取财务数据，并生成原生可编辑 PowerPoint 报告。**

你只需要告诉它"分析万科、金茂、华润置地 2023 到 2025 年的财报"，剩下的事情全部自动完成。整个过程分四步：

---

### 第一步：自动下载年报

系统从中国证监会和港交所的官方披露网站，自动搜索并下载你指定公司、指定年份的年报文件（PDF 格式）。不管是 A 股（万科、保利、招商蛇口……）还是港股（华润置地、中海地产、龙湖集团……），都能自动处理。网络偶尔不稳定时会自动重试，不需要人工盯着。

### 第二步：自动读懂报表

下载完 PDF 之后，系统先把这几百页的文件转成可读的文字，然后从里面精准提取关键财务数据。包括：

- **赚钱能力**：营业收入、毛利、毛利率、净利润、每股收益等
- **债务水平**：总负债、资产负债率、净负债率、有息负债等
- **资产结构**：总资产、存货、投资性房地产、货币资金等
- **现金状况**：经营现金流、投资现金流、筹资现金流等
- **业务构成**：房地产开发收入、商业物业收入、物业管理收入等

一共覆盖几十个核心财务指标，基本涵盖了专业研究员做对标分析时需要的全部数据。

不同公司的报表写法不一样——A 股公司用简体中文，港股公司用繁体中文甚至中英双语，科目名称常常不同。比如同一个"营业收入"，有的报告里叫"營業額"，有的叫"收益"。系统会自动识别这些差异，统一成一致的口径，让不同公司的数据可以直接横向比较。

### 第三步：自动整理成数据表

提取出来的数据会生成一份 Excel 表格。表格分两部分：

- **汇总表**：一张大表，行是各项指标，列是各公司各年份，数值全部是纯数字，可以直接拿来求和、算平均值、做透视分析，不用再手动调整格式
- **明细表**：每家公司的每个年份各一张表，三列——指标名称、数据、备注（记录了这个数是从报告哪一页找到的，方便需要时翻回去核对）

### 第四步：用 PPT Master 生成可编辑 PowerPoint 报告

最后，系统把核验后的 Excel 底稿和报告素材交给 [PPT Master](https://github.com/hugohe3/ppt-master)，生成原生可编辑 `.pptx`。报告内容包括：

- **关键数据卡片**：公司概况、核心指标一目了然
- **图表分析**：柱状图对比各公司收入利润、折线图展示毛利率变化趋势，直观看出谁在增长谁在下滑
- **文字分析**：盈利能力、偿债能力、资产质量、现金流状况逐项解读
- **风险提示**：高负债、存货减值、现金流压力等潜在风险自动标注
- **汇总数据表**：完整的数据底表附在报告末尾，方便查阅具体数字
- **可编辑元素**：文本、形状、图表和表格可在 PowerPoint 中继续调整

报告不是网页截图，也不是每页一张扁平图片。PPT Master 以 SVG 为页面设计源，并导出 DrawingML/原生对象组成的 PowerPoint；图表数量、页面结构和分析结论根据实际数据生成。

---

### 总结：原来半天的手工活，现在 1-2 小时自动完成

以前做一份多公司对标分析，流程是这样的：打开好几个网站 → 逐个下载 PDF → 翻几百页找报表那一页 → 手动抄数字到 Excel → 一个个核对 → 再做图表 → 再写分析。一个下午就过去了，还容易抄错。

现在你只需要一句话告诉它分析哪些公司、哪些年份。过程中需要两次明确确认：开头确认分析范围，以及 PPT Master 的报告策略与设计方案确认。完成后会得到可追溯 Excel 底稿和可继续编辑的 PowerPoint 报告。

---

### 如何使用

Financial Disclosure Analysis 是一个 AI Agent Skill，可在具备文件读写和命令执行能力的 Agent 环境中运行。最终报告依赖 PPT Master。

**安装：**

```bash
git clone https://github.com/jiannanxiboy/financial-disclosure-analysis.git ~/.claude/skills/financial-disclosure-analysis/
cd ~/.claude/skills/financial-disclosure-analysis
pip install -r requirements.txt
playwright install chromium

git clone https://github.com/hugohe3/ppt-master.git ~/ppt-master
pip install -r ~/ppt-master/requirements.txt
export PPT_MASTER_HOME=~/ppt-master
```

Windows PowerShell 设置环境变量：

```powershell
$env:PPT_MASTER_HOME = "$HOME\ppt-master"
python scripts/ppt_master_bridge.py check
```

**通过 CC Switch 安装：**

在「Skills → 仓库管理 → 添加仓库」中填写：

```text
Repository: jiannanxiboy/financial-disclosure-analysis
Branch: master
Subdirectory: 留空
```

仓库根目录就是 Skill 目录；CC Switch 会将其识别为 `financial-disclosure-analysis`。

**使用：**

安装后，在 Claude Code 中输入 `/financial-disclosure-analysis` 即可启动，按提示告诉它要分析的公司和年份即可。

**脚本也可独立使用（面向开发者）：**

```bash
# 下载 A 股年报
python scripts/a_share.py search-annual --codes 000002 600048 --year 2024 --download-dir ./data/pdfs

# 下载港股年报
python scripts/hk_share.py search-annual --codes "00688,01109" --year 2024 --download-dir ./data/pdfs

# PDF 转文字
python scripts/pdf_to_text.py --input-dir ./data/pdfs --output-dir ./data/txt --skip-existing

# 检查 PPT Master 集成
python scripts/ppt_master_bridge.py check

# 运行单元与回归测试
python -m unittest discover -s tests -v
```

完整流程参考 [SKILL.md](SKILL.md)。

---

### 支持的市场

- **A 股**：通过巨潮资讯网搜索下载
- **港股**：通过港交所披露易搜索下载

### License

MIT
