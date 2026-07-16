# 报告事实包与结论校验

在完成 `数据底稿.xlsx` 后运行：

```bash
python {SD}/report_facts.py \
  --workbook "{output_dir}/数据底稿.xlsx" \
  --claims "{output_dir}/report_claims.json" \
  --output "{output_dir}/report_facts.json"
```

`report_facts.json` 保存每个数字对应的公司、期间、指标、单位和 Excel 单元格地址。声明了 claims 时，任一结论校验失败都会返回非零退出码。

`report_claims.json` 示例：

```json
{
  "claims": [
    {
      "id": "C01",
      "text": "六家公司2025年销售额均低于2023年",
      "checks": [
        {
          "indicator": "合约销售额",
          "operator": "all_lt",
          "left_period": "2025",
          "right_period": "2023"
        }
      ]
    },
    {
      "id": "C02",
      "text": "华润置地2025年毛利率样本最高",
      "checks": [
        {
          "indicator": "毛利率",
          "period": "2025",
          "operator": "max_company",
          "expected": "华润置地"
        }
      ]
    }
  ]
}
```

支持的操作：

- 单点：`eq`、`lt`、`lte`、`gt`、`gte`，需提供公司、期间和 expected。
- 全样本期间比较：`all_lt`、`all_gt`，需提供 left_period 和 right_period。
- 同期极值公司：`max_company`、`min_company`，需提供期间和 expected 公司。

PPT 标题和执行摘要中的事实性结论必须有 claim ID；推测性判断仍需在 `报告素材.md` 中写明依据和不确定性。
