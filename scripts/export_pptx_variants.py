#!/usr/bin/env python3
"""Export fidelity-first and native chart/table PPTX variants from one PPT Master project."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

from _common import atomic_write_json
from ppt_master_bridge import inspect_capabilities, ppt_master_env, resolve_ppt_master
from validate_pptx import validate


def _safe_name(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*]+', "_", value).strip(" .")


def _run(command: list[str], cwd: Path) -> None:
    result = subprocess.run(
        command,
        cwd=cwd,
        check=False,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=ppt_master_env(),
    )
    if result.returncode:
        raise RuntimeError(f"PPTX export failed ({result.returncode})")


def export_variants(
    project_path: str | Path,
    ppt_master_dir: str | None = None,
    base_name: str | None = None,
    force: bool = False,
) -> dict:
    project = Path(project_path).resolve()
    skill_dir = resolve_ppt_master(ppt_master_dir)
    capabilities = inspect_capabilities(skill_dir)
    if not capabilities["compatible"]:
        raise RuntimeError("PPT Master缺少能力: " + ", ".join(capabilities["missing"]))
    svg_files = sorted((project / "svg_output").glob("*.svg"))
    if not svg_files:
        raise FileNotFoundError(f"svg_output中没有SVG: {project}")
    exports = project / "exports"
    exports.mkdir(parents=True, exist_ok=True)
    default_base = project.name.split("_ppt", 1)[0]
    name = _safe_name(base_name or default_base)
    shapes = exports / f"{name}_editable_shapes.pptx"
    native = exports / f"{name}_native_charts_tables.pptx"
    for output in (shapes, native):
        if output.exists() and not force:
            raise FileExistsError(f"输出已存在，使用 --force 覆盖: {output}")
    exporter = skill_dir / "scripts" / "svg_to_pptx.py"
    _run([sys.executable, str(exporter), str(project), "-o", str(shapes)], skill_dir)
    _run([
        sys.executable, str(exporter), str(project), "-o", str(native),
        "--native-charts-and-tables",
    ], skill_dir)
    expected = len(svg_files)
    report = {
        "schema_version": 1,
        "project": str(project),
        "preferred_delivery": str(shapes),
        "variants": [
            {
                "role": "editable_shapes",
                "path": str(shapes),
                "validation": validate(shapes, expected_slides=expected, require_notes=True, renderer="none"),
            },
            {
                "role": "native_charts_tables",
                "path": str(native),
                "validation": validate(native, expected_slides=expected, require_notes=True, renderer="none"),
            },
        ],
    }
    report["ok"] = all(item["validation"]["ok"] for item in report["variants"])
    atomic_write_json(project / "analysis" / "export_variants.json", report)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("project")
    parser.add_argument("--ppt-master-dir")
    parser.add_argument("--base-name")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    try:
        report = export_variants(
            args.project, args.ppt_master_dir, args.base_name, args.force,
        )
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["ok"] else 1
    except (FileNotFoundError, FileExistsError, RuntimeError) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
