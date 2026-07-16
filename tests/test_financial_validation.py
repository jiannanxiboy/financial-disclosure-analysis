import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from financial_validation import MetricRecord, validate_records


def record(indicator, value, unit="亿", company="公司A", period="2025", status="normalized"):
    return MetricRecord(
        company=company,
        period=period,
        indicator=indicator,
        raw_value=f"{value}{unit}" if value is not None else "-",
        raw_unit=unit if value is not None else "",
        normalized_value=value,
        normalized_unit=unit,
        source_file="source.tsv",
        source_location="第1页",
        status="missing" if value is None else status,
    )


class FinancialValidationTests(unittest.TestCase):
    def test_recalculation_warning_and_missing_are_structured(self):
        records = [
            record("营业收入", 100),
            record("毛利", 20),
            record("毛利率", 15, "%"),
            record("归母净利润", None),
        ]
        result = validate_records(records)
        self.assertEqual(result["status"], "passed_with_warnings")
        self.assertEqual(result["stats"]["warnings"], 1)
        self.assertEqual(result["warnings"][0]["category"], "recalculation")
        self.assertEqual(result["stats"]["missing"], 1)

    def test_absurd_ratio_is_blocking_error(self):
        result = validate_records([record("毛利率", 1200, "%")])
        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["errors"][0]["category"], "ratio_range")


if __name__ == "__main__":
    unittest.main()
