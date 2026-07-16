import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import validate_pptx


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


class PptxValidationTests(unittest.TestCase):
    def test_inspects_slide_notes_chart_table_and_dimensions(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "报告.pptx"
            with zipfile.ZipFile(pptx, "w") as archive:
                archive.writestr("[Content_Types].xml", "<Types/>")
                archive.writestr(
                    "ppt/presentation.xml",
                    f'<p:presentation xmlns:p="{P_NS}"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
                )
                archive.writestr(
                    "ppt/slides/slide1.xml",
                    f'<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld><a:tbl/></p:cSld></p:sld>',
                )
                archive.writestr("ppt/notesSlides/notesSlide1.xml", f'<p:notes xmlns:p="{P_NS}"/>')
                archive.writestr("ppt/charts/chart1.xml", "<chart/>")
            report = validate_pptx.validate(
                pptx, expected_slides=1, require_notes=True, renderer="none"
            )
            self.assertTrue(report["ok"])
            self.assertEqual(report["slides"], 1)
            self.assertEqual(report["notes"], 1)
            self.assertEqual(report["charts"], 1)
            self.assertEqual(report["tables"], 1)
            self.assertEqual(report["dimensions"], {"cx": 12192000, "cy": 6858000})


if __name__ == "__main__":
    unittest.main()
