import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import report_facts


class ReportFactsTests(unittest.TestCase):
    def test_extracts_cells_and_verifies_cross_period_claim(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "数据底稿.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "汇总透视表"
            ws.append(["指标", "单位", "公司A 2023", "公司A 2025", "公司B 2023", "公司B 2025"])
            ws.append(["合约销售额", "亿", 100, 80, 120, 90])
            ws.append(["毛利率", "%", 20, 15, 18, 16])
            wb.save(path)
            package = report_facts.extract_facts(path)
            claims = [{
                "id": "C01",
                "text": "两家公司销售额均下降",
                "checks": [{
                    "indicator": "合约销售额",
                    "operator": "all_lt",
                    "left_period": "2025",
                    "right_period": "2023",
                }],
            }]
            result = report_facts.verify_claims(package, claims)
        self.assertEqual(len(package["facts"]), 8)
        self.assertEqual(package["facts"][0]["cell"], "汇总透视表!C2")
        self.assertEqual(result[0]["status"], "verified")
        self.assertEqual(len(result[0]["checks"][0]["evidence"]), 4)


if __name__ == "__main__":
    unittest.main()
