import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import _common


class FakeResponse:
    status_code = 200
    content = b"%PDF-1.7\n" + b"x" * 12000

    def iter_content(self, chunk_size):
        yield self.content

    def close(self):
        pass


class DownloadCacheTests(unittest.TestCase):
    def test_second_download_restores_content_cache_without_network(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            cache = root / "cache"
            first = root / "first.pdf"
            second = root / "second.pdf"
            env = {"FINANCIAL_DISCLOSURE_CACHE_DIR": str(cache)}
            with patch.dict(os.environ, env, clear=False), \
                    patch.object(_common, "verify_pdf", return_value=True), \
                    patch.object(_common.requests, "get", return_value=FakeResponse()) as get:
                self.assertTrue(_common.download_pdf("https://example.test/report.pdf", first))
                self.assertEqual(get.call_count, 1)
            with patch.dict(os.environ, env, clear=False), \
                    patch.object(_common, "verify_pdf", return_value=True), \
                    patch.object(_common.requests, "get", side_effect=AssertionError("network should not run")):
                self.assertTrue(_common.download_pdf("https://example.test/report.pdf", second))
            self.assertEqual(first.read_bytes(), second.read_bytes())
            self.assertEqual(len(list((cache / "objects").glob("*.pdf"))), 1)
            self.assertEqual(len(list((cache / "urls").glob("*.json"))), 1)


if __name__ == "__main__":
    unittest.main()
