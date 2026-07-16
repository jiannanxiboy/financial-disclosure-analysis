import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import export_pptx_variants


class ExportVariantsTests(unittest.TestCase):
    def test_exports_both_named_variants_and_marks_shapes_preferred(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            project = root / "project"
            (project / "svg_output").mkdir(parents=True)
            (project / "svg_output" / "01_cover.svg").write_text("<svg/>", encoding="utf-8")
            skill = root / "ppt-master"
            (skill / "scripts").mkdir(parents=True)
            (skill / "scripts" / "svg_to_pptx.py").write_text("", encoding="utf-8")

            commands = []

            def fake_run(command, _cwd):
                commands.append(command)
                output = Path(command[command.index("-o") + 1])
                output.write_bytes(b"pptx")

            validation = {"ok": True, "slides": 1, "errors": []}
            capabilities = {"compatible": True, "missing": []}
            with patch.object(export_pptx_variants, "resolve_ppt_master", return_value=skill), \
                    patch.object(export_pptx_variants, "inspect_capabilities", return_value=capabilities), \
                    patch.object(export_pptx_variants, "_run", side_effect=fake_run), \
                    patch.object(export_pptx_variants, "validate", return_value=validation):
                report = export_pptx_variants.export_variants(project, base_name="六家/房企")

            preferred = Path(report["preferred_delivery"])
            self.assertEqual(preferred.name, "六家_房企_editable_shapes.pptx")
            self.assertTrue(preferred.exists())
            self.assertTrue(Path(report["variants"][1]["path"]).exists())
            self.assertNotIn("--native-charts-and-tables", commands[0])
            self.assertIn("--native-charts-and-tables", commands[1])
            self.assertTrue((project / "analysis" / "export_variants.json").exists())


if __name__ == "__main__":
    unittest.main()
