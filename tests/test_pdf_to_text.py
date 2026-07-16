import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import pdf_to_text


class PdfRoutingTests(unittest.TestCase):
    def test_detect_mode_writes_page_index_and_marks_ocr(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "扫描年报.pdf"
            output = root / "扫描年报.txt"
            source.write_bytes(b"pdf")
            extracted = {
                "total_pages": 2,
                "tool": "pdfplumber",
                "pages": [
                    {"page": 1, "char_start": 100, "char_end": 120, "text_chars": 0, "table_count": 0, "status": "empty"},
                    {"page": 2, "char_start": 120, "char_end": 300, "text_chars": 200, "table_count": 1, "status": "text"},
                ],
            }
            with patch.object(pdf_to_text, "_verify_pdf_readable", return_value=2), \
                    patch.object(pdf_to_text, "_extract_pdf_once", return_value=extracted), \
                    patch.object(pdf_to_text, "_sha256", return_value="a" * 64), \
                    patch.object(pdf_to_text, "_ocr_command", return_value=None):
                pages = pdf_to_text.extract_pdf_to_text(
                    str(source), str(output), ocr_mode="detect", scanned_ratio_threshold=0.3
                )
            self.assertEqual(pages, 2)
            metadata = json.loads((root / "扫描年报.pages.json").read_text(encoding="utf-8"))
            self.assertTrue(metadata["needs_ocr"])
            self.assertEqual(metadata["scanned_pages"], [1])
            self.assertFalse(metadata["ocr_applied"])

    def test_required_mode_fails_when_ocr_is_unavailable(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "scan.pdf"
            output = root / "scan.txt"
            source.write_bytes(b"pdf")
            extracted = {
                "total_pages": 1,
                "tool": "pdfplumber",
                "pages": [{"page": 1, "char_start": 0, "char_end": 1, "text_chars": 0, "table_count": 0, "status": "empty"}],
            }
            with patch.object(pdf_to_text, "_verify_pdf_readable", return_value=1), \
                    patch.object(pdf_to_text, "_extract_pdf_once", return_value=extracted), \
                    patch.object(pdf_to_text, "_sha256", return_value="b" * 64), \
                    patch.object(pdf_to_text, "_ocr_command", return_value=None):
                with self.assertRaisesRegex(RuntimeError, "ocrmypdf"):
                    pdf_to_text.extract_pdf_to_text(str(source), str(output), ocr_mode="required")
            metadata = json.loads((root / "scan.pages.json").read_text(encoding="utf-8"))
            self.assertTrue(metadata["needs_ocr"])
            self.assertIn("ocr_error", metadata)


if __name__ == "__main__":
    unittest.main()
