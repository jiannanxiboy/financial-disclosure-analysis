import csv
import sys
import tempfile
import unittest
from collections import OrderedDict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import generate_excel


class ReadTsvTests(unittest.TestCase):
    def test_reads_standard_header_bom_and_quoted_tab(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "测试.tsv"
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle, delimiter="\t")
                writer.writerow(["指标名称", "数据", "备注"])
                writer.writerow(["营业收入", "1,200亿", "年报\t第10页"])
                writer.writerow(["毛利率", "18.5%", "财务摘要"])
            data = generate_excel.read_tsv(str(path))
        self.assertEqual(list(data), ["营业收入", "毛利率"])
        self.assertEqual(data["营业收入"], ("1,200亿", "年报\t第10页"))

    def test_accepts_legacy_header(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "legacy.tsv"
            path.write_text("名称\t金额/比例\t备注\n营业收入\t10亿\t第1页\n", encoding="utf-8")
            data = generate_excel.read_tsv(str(path))
        self.assertEqual(list(data), ["营业收入"])

    def test_rejects_duplicate_indicator(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "duplicate.tsv"
            path.write_text(
                "指标名称\t数据\t备注\n营业收入\t10亿\tA\n营业收入\t11亿\tB\n",
                encoding="utf-8",
            )
            with self.assertRaisesRegex(ValueError, "指标重复"):
                generate_excel.read_tsv(str(path))

    def test_rejects_unknown_header(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.tsv"
            path.write_text("项目\t数字\n营业收入\t10亿\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "缺少必要字段"):
                generate_excel.read_tsv(str(path))


class UnitNormalizationTests(unittest.TestCase):
    def setUp(self):
        self.definitions = generate_excel.load_indicator_definitions({})

    def test_parse_value_supports_parentheses_and_missing(self):
        self.assertEqual(generate_excel.parse_value("(1,234)万元"), (-1234.0, "万元"))
        self.assertEqual(generate_excel.parse_value("—"), (None, ""))

    def test_average_price_converts_to_ten_thousand_yuan(self):
        self.assertEqual(
            generate_excel.normalize_value("平均售价", "23032元/平方米", self.definitions),
            (2.3032, "万元/平方米"),
        )
        self.assertEqual(
            generate_excel.normalize_value("平均售价", "2.05万元/平方米", self.definitions),
            (2.05, "万元/平方米"),
        )

    def test_money_converts_to_hundred_million_yuan(self):
        self.assertEqual(
            generate_excel.normalize_value("营业收入", "125000万元", self.definitions),
            (12.5, "亿"),
        )

    def test_unknown_unit_fails_closed(self):
        with self.assertRaisesRegex(ValueError, "无法换算"):
            generate_excel.normalize_value("平均售价", "2.3美元/平方米", self.definitions)

    def test_pivot_contains_normalized_numeric_values(self):
        all_data = OrderedDict(
            {
                ("公司A", "2025"): OrderedDict({"平均售价": ("23032元/平方米", "A")}),
                ("公司B", "2025"): OrderedDict({"平均售价": ("2.05万元/平方米", "B")}),
            }
        )
        columns, rows = generate_excel.build_pivot(all_data, ["平均售价"], self.definitions)
        self.assertEqual(columns, ["指标", "单位", "公司A 2025", "公司B 2025"])
        self.assertEqual(rows, [["平均售价", "万元/平方米", 2.3032, 2.05]])


if __name__ == "__main__":
    unittest.main()
