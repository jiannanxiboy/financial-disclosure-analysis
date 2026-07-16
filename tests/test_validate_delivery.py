import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import validate_delivery


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


class DeliveryValidationTests(unittest.TestCase):
    def test_complete_delivery_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data = root / "data"
            project = root / "project"
            (project / "analysis").mkdir(parents=True)
            (project / "sources").mkdir()
            (project / "svg_output").mkdir()
            (project / "exports").mkdir()
            data.mkdir()
            original = data / "数据底稿.xlsx"
            original.write_bytes(b"workbook")
            imported = project / "sources" / original.name
            imported.write_bytes(original.read_bytes())
            handoff = {
                "sources": [{
                    "path": str(original),
                    "filename": original.name,
                    "sha256": sha(original),
                    "imported_path": f"sources/{original.name}",
                    "imported_sha256": sha(imported),
                }]
            }
            (project / "analysis" / "financial-disclosure-analysis-handoff.json").write_text(
                json.dumps(handoff, ensure_ascii=False), encoding="utf-8"
            )
            (data / "validation_summary.json").write_text(
                json.dumps({"stats": {"records": 1}, "errors": [], "warnings": []}), encoding="utf-8"
            )
            (data / "normalized_records.json").write_text(
                json.dumps({"records": [{}]}), encoding="utf-8"
            )
            (data / "report_facts.json").write_text(
                json.dumps({"facts": [{}], "claims": [{}], "status": "verified"}), encoding="utf-8"
            )
            (project / "svg_output" / "01_cover.svg").write_text("<svg/>", encoding="utf-8")
            pptx = project / "exports" / "report_editable_shapes.pptx"
            pptx.write_bytes(b"pptx")
            fake_skill = root / "ppt-master"
            (fake_skill / "scripts").mkdir(parents=True)
            (fake_skill / "scripts" / "svg_quality_checker.py").write_text("", encoding="utf-8")
            completed = SimpleNamespace(returncode=0, stdout="", stderr="")
            pptx_report = {"ok": True, "slides": 1, "notes": 1, "errors": []}
            with patch.object(validate_delivery, "resolve_ppt_master", return_value=fake_skill), \
                    patch.object(validate_delivery.subprocess, "run", return_value=completed), \
                    patch.object(validate_delivery, "validate_pptx", return_value=pptx_report):
                report = validate_delivery.validate_delivery(
                    project, data, expected_slides=1, renderer="none"
                )
        self.assertEqual(report["status"], "passed_with_warnings")
        self.assertFalse(report["errors"])
        self.assertEqual(report["checks"]["normalized_records"], 1)


if __name__ == "__main__":
    unittest.main()
