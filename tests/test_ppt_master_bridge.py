import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import ppt_master_bridge


class ManifestTests(unittest.TestCase):
    def test_utf8_manifest_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "分析" / "交接清单.json"
            manifest = {
                "sources": [
                    {
                        "path": r"C:\测试目录\报告素材.md",
                        "filename": "报告素材.md",
                    }
                ]
            }
            ppt_master_bridge._write_manifest(path, manifest)
            self.assertEqual(json.loads(path.read_text(encoding="utf-8")), manifest)
            self.assertIn("报告素材.md", path.read_text(encoding="utf-8"))

    def test_prepare_project_preserves_chinese_paths_and_hashes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            skill_dir = root / "ppt-master" / "skills" / "ppt-master"
            skill_dir.mkdir(parents=True)
            projects_dir = root / "项目目录"
            source_dir = root / "财务资料"
            source_dir.mkdir()
            sources = [source_dir / "报告素材.md", source_dir / "数据底稿.xlsx"]
            sources[0].write_text("六家房企分析", encoding="utf-8")
            sources[1].write_bytes(b"xlsx-test-content")
            project_dir = projects_dir / "六家房企_ppt169_20260716"

            def fake_run(command, cwd):
                action = command[2]
                if action == "init":
                    (project_dir / "sources").mkdir(parents=True)
                    return f"[OK] Project initialized: {project_dir}"
                if action == "import-sources":
                    imported_project = Path(command[3])
                    for item in command[4:-1]:
                        staged = Path(item)
                        shutil.move(str(staged), imported_project / "sources" / staged.name)
                    return "[OK] imported"
                raise AssertionError(command)

            with patch.object(ppt_master_bridge, "_run", side_effect=fake_run):
                result = ppt_master_bridge.prepare_project(
                    skill_dir=skill_dir,
                    project_name="六家房企",
                    projects_dir=projects_dir,
                    sources=sources,
                    canvas_format="ppt169",
                )

            self.assertEqual(result, project_dir.resolve())
            manifest_path = project_dir / "analysis" / "financial-disclosure-analysis-handoff.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["schema_version"], 1)
            self.assertEqual([item["filename"] for item in manifest["sources"]], ["报告素材.md", "数据底稿.xlsx"])
            for item, source in zip(manifest["sources"], sources):
                self.assertEqual(item["path"], str(source.resolve()))
                self.assertEqual(item["sha256"], item["imported_sha256"])
                self.assertEqual(item["imported_path"], f"sources/{source.name}")


if __name__ == "__main__":
    unittest.main()
