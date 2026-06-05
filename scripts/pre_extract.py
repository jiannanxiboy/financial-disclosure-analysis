#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
预提取年报TXT中的关键财务报表章节，输出精简版TXT供子Agent读取。

用法:
  python pre_extract.py --input 年报.txt --output 年报_精简.txt

策略（按优先级）:
  1. 五年财务摘要       — 港股首页就有，包含最全的核心数据
  2. 综合损益表/收益表    — 营收、成本、利润
  3. 综合财务状况表       — 资产、负债、权益
  4. 综合现金流量表       — 经营/投资/筹资现金流
  5. 分部报告/MD&A        — 分业务收入毛利
  6. 附注-存货跌价        — 减值数据

对每个章节:
  - 找到标题行 → 向后取500行（覆盖完整报表）
  - 如果文件大于3000行，只取1-4节（核心报表）
"""

import argparse, os, sys, re

# ── 章节标题模式（支持简繁体、中英文） ──
SECTION_PATTERNS = [
    # (优先级, 标题正则, 提取行数, 标签)
    (1, r'五年財務摘要|五年财务摘要|Five\s*Year.*Summary', 300, '五年财务摘要'),
    (2, r'綜合損益|综合损益|綜合收益|综合收益|Consolidated\s+(Statement\s+of\s+)?(Profit\s+or\s+Loss|Income)', 500, '综合损益表'),
    (3, r'綜合財務狀況|综合财务状况|Consolidated\s+(Statement\s+of\s+)?Financial\s+Position', 500, '综合财务状况表'),
    (4, r'綜合現金流量|综合现金流量|Consolidated\s+(Statement\s+of\s+)?Cash\s+Flows', 400, '综合现金流量表'),
    (5, r'分部(收入|報告|报告|信息|資料)|Segment\s+(Revenue|Information|Report)', 400, '分部报告'),
    (6, r'存貨(跌價|减值|減值)|物業減值|Inventory\s+Impairment|存貨\s*(—|–|-|—)', 300, '存货减值附注'),
    (7, r'利息資本化|利息资本化|Borrowing\s+[Cc]osts?\s*[Cc]apitali[sz]ed', 100, '利息资本化'),
]

# A股特有章节
A_SHARE_PATTERNS = [
    (2, r'合并利润表(?!.*母公司)|合併利潤表', 500, '合并利润表'),
    (3, r'合并资产负债表(?!.*母公司)|合併資產負債表', 500, '合并资产负债表'),
    (4, r'合并现金流量表(?!.*母公司)|合併現金流量表', 400, '合并现金流量表'),
    (5, r'主营业务分行业|主營業務分行業', 300, '主营业务分行业'),
    (6, r'存货跌价准备|存貨跌價準備', 300, '存货跌价准备'),
    (8, r'主要会计数据|主要會計數據', 300, '主要会计数据'),
]


def find_sections(lines: list[str], patterns: list[tuple]) -> dict:
    """在行列表中查找各章节的起止位置。返回 {标签: (start, end)}。

    去重规则: 同一标签只保留第一次命中（优先匹配前面的行）。
    排除规则: 跳过标题行含"附注索引"或"目录"的（TOC条目）。
    """
    found = {}

    for priority, pattern, lines_to_take, label in patterns:
        if label in found:
            continue
        regex = re.compile(pattern)
        for i, line in enumerate(lines):
            stripped = line.strip()
            if not regex.search(stripped):
                continue
            # 排除目录行/附注索引行
            if any(kw in stripped for kw in ['附注索引', '目錄', '目录', '→', '…']):
                continue
            # 排除只有标题的行后紧跟的是分隔符/空行(可能是TOC中的条目)
            end = min(i + lines_to_take, len(lines))
            found[label] = (i, end)
            break

    return found


def extract(lines: list[str], patterns: list[tuple], max_output_lines: int = 8000) -> str:
    """从行列表中提取关键章节，拼接为精简文本。"""
    sections = find_sections(lines, patterns)

    if not sections:
        # 什么都没找到，返回文件头2000行作为fallback
        return "".join(lines[:2000]) + "\n\n[预提取未定位到任何报表章节，以下为文件前2000行]\n"

    # 合并重叠区间，按行号排序
    intervals = sorted(sections.values())
    merged = []
    for start, end in intervals:
        if merged and start <= merged[-1][1] + 5:
            # 与上一个区间重叠或相近（5行内），合并
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))

    # 拼接输出
    out_parts = []
    total_lines = 0
    for start, end in merged:
        chunk = lines[start:end]
        chunk_lines = len(chunk)
        if total_lines + chunk_lines > max_output_lines:
            # 截断
            remaining = max_output_lines - total_lines
            chunk = chunk[:remaining]
            out_parts.append(f"\n{'='*60}\n[截取行 {start+1}-{start+len(chunk)}，原始区域 {start+1}-{end}]\n{'='*60}\n")
            out_parts.extend(chunk)
            break
        out_parts.append(f"\n{'='*60}\n[行 {start+1}-{end}]\n{'='*60}\n")
        out_parts.extend(chunk)
        total_lines += chunk_lines

    # 如果提取的内容太少，追加文件前500行作为补充
    if total_lines < 500:
        out_parts.insert(0, "".join(lines[:500]) + "\n\n[以上为文件前500行，预提取内容不足500行]\n")

    return "".join(out_parts)


def main():
    parser = argparse.ArgumentParser(description="预提取年报TXT中的关键财报章节")
    parser.add_argument("--input", required=True, help="原始TXT路径")
    parser.add_argument("--output", required=True, help="精简输出TXT路径")
    parser.add_argument("--market", default="auto", choices=["auto", "a", "hk"],
                        help="市场: auto=自动检测, a=A股, hk=港股")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"错误: 文件不存在: {args.input}", file=sys.stderr)
        sys.exit(1)

    with open(args.input, "r", encoding="utf-8") as f:
        lines = f.readlines()

    total_lines = len(lines)
    print(f"原始文件: {total_lines} 行")

    # 自动检测市场
    if args.market == "auto":
        # 检测前200行是否含港股特征
        head = "".join(lines[:200])
        hk_markers = ["披露易", "HKEX", "港交所", "綜合財務", "綜合損益", "綜合收益"]
        a_markers = ["巨潮", "cninfo", "上海证券交易所", "深圳证券交易所", "合并资产负债表", "合并利润表"]
        hk_score = sum(1 for m in hk_markers if m in head)
        a_score = sum(1 for m in a_markers if m in head)
        market = "hk" if hk_score > a_score else "a"
    else:
        market = args.market

    print(f"检测市场: {'港股' if market == 'hk' else 'A股'}")

    # 选择匹配模式
    patterns = SECTION_PATTERNS.copy()
    if market == "a":
        patterns.extend(A_SHARE_PATTERNS)
    patterns.sort(key=lambda x: x[0])  # 按优先级排序

    # 提取
    max_lines = 6000 if total_lines > 15000 else 3000
    result = extract(lines, patterns, max_output_lines=max_lines)

    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(result)

    out_lines = result.count("\n")
    reduction = (1 - out_lines / total_lines) * 100 if total_lines > 0 else 0
    print(f"精简输出: {out_lines} 行 (缩减 {reduction:.0f}%)")

    # 报告找到了哪些章节
    found_sections = find_sections(lines, patterns)
    if found_sections:
        print(f"定位章节: {', '.join(found_sections.keys())}")
    else:
        print("⚠ 未能定位任何报表章节，输出了文件前段作为fallback")


if __name__ == "__main__":
    main()
