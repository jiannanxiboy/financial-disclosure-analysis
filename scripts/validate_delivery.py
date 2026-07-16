#!/usr/bin/env python3
"""Run one-command acceptance checks for financial disclosure analysis delivery artifacts."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from _common import atomic_write_json
from ppt_master_bridge import resolve_ppt_master
from validate_pptx import validate as validate_pptx


STANDARD_SVG_NAME = re.compile(r"^\d{2}_[\w-]+\.svg$", re.UNICODE)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_delivery(
    project_path: str | Path,
    data_dir: str | Path | None = None,
    ppt_master_dir: str | None = None,
    expected_slides: int | None = None,
    renderer: str = "auto",
    allow_legacy_svg_names: bool = False,
    allow_missing_facts: bool = False,
) -> dict:
    project = Path(project_path).resolve()
    data = Path(data_dir).resolve() if data_dir else None
    errors: list[str] = []
    warnings: list[str] = []
    checks: dict[str, object] = {}

    handoff_path = project / "analysis" / "financial-disclosure-analysis-handoff.json"
    if not handoff_path.is_file():
        errors.append("missing handoff manifest")
    else:
        handoff = json.loads(handoff_path.read_text(encoding="utf-8"))
        source_checks = []
        for item in handoff.get("sources", []):
            original = Path(item["path"])
            imported = project / item.get("imported_path", "")
            original_ok = original.is_file() and _sha256(original) == item.get("sha256")
            imported_ok = imported.is_file() and _sha256(imported) == item.get("imported_sha256", item.get("sha256"))
            source_checks.append({
                "filename": item.get("filename"),
                "original_ok": original_ok,
                "imported_ok": imported_ok,
            })
            if not original_ok or not imported_ok:
                errors.append(f"handoff source hash/path check failed: {item.get('filename')}")
        checks["handoff"] = source_checks

    validation_path = data / "validation_summary.json" if data else None
    if validation_path and validation_path.is_file():
        validation_summary = json.loads(validation_path.read_text(encoding="utf-8"))
        checks["financial_validation"] = validation_summary.get("stats", {})
        if validation_summary.get("errors"):
            errors.append(f"financial validation has {len(validation_summary['errors'])} blocking errors")
        if validation_summary.get("warnings"):
            warnings.append(f"financial validation has {len(validation_summary['warnings'])} warnings")
    elif data:
        errors.append("missing validation_summary.json")

    records_path = data / "normalized_records.json" if data else None
    if records_path and records_path.is_file():
        records = json.loads(records_path.read_text(encoding="utf-8"))
        checks["normalized_records"] = len(records.get("records", []))
    elif data:
        errors.append("missing normalized_records.json")

    facts_path = data / "report_facts.json" if data else None
    if facts_path and facts_path.is_file():
        facts = json.loads(facts_path.read_text(encoding="utf-8"))
        checks["report_facts"] = {
            "facts": len(facts.get("facts", [])),
            "claims": len(facts.get("claims", [])),
            "status": facts.get("status"),
        }
        if facts.get("status") == "failed":
            errors.append("report_facts contains unverified claims")
    elif data and not allow_missing_facts:
        errors.append("missing report_facts.json")

    svg_files = sorted((project / "svg_output").glob("*.svg"))
    checks["svg_count"] = len(svg_files)
    if not svg_files:
        errors.append("svg_output contains no pages")
    if expected_slides is not None and len(svg_files) != expected_slides:
        errors.append(f"SVG page count mismatch: expected {expected_slides}, got {len(svg_files)}")
    legacy_names = [item.name for item in svg_files if not STANDARD_SVG_NAME.match(item.name)]
    checks["nonstandard_svg_names"] = legacy_names
    if legacy_names:
        message = f"nonstandard SVG names: {', '.join(legacy_names[:5])}"
        (warnings if allow_legacy_svg_names else errors).append(message)

    try:
        skill_dir = resolve_ppt_master(ppt_master_dir)
        checker = skill_dir / "scripts" / "svg_quality_checker.py"
        result = subprocess.run(
            [sys.executable, str(checker), str(project)],
            cwd=skill_dir,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
        )
        checks["svg_quality"] = {"ok": result.returncode == 0}
        if result.returncode:
            errors.append("PPT Master SVG quality checker failed")
    except FileNotFoundError as exc:
        errors.append(str(exc))

    variant_manifest = project / "analysis" / "export_variants.json"
    pptx_paths: list[Path] = []
    if variant_manifest.is_file():
        variants = json.loads(variant_manifest.read_text(encoding="utf-8"))
        pptx_paths = [Path(item["path"]) for item in variants.get("variants", [])]
    if not pptx_paths:
        pptx_paths = sorted((project / "exports").glob("*.pptx"))
        if pptx_paths:
            warnings.append("export_variants.json missing; validating discovered PPTX files")
    if not pptx_paths:
        errors.append("no PPTX exports found")
    pptx_reports = []
    for pptx in pptx_paths:
        report = validate_pptx(
            pptx,
            expected_slides=expected_slides or len(svg_files),
            require_notes=True,
            renderer=renderer,
        )
        pptx_reports.append(report)
        if not report["ok"]:
            errors.append(f"PPTX validation failed: {pptx.name}")
    checks["pptx"] = pptx_reports

    report = {
        "schema_version": 1,
        "project": str(project),
        "data_dir": str(data) if data else None,
        "status": "failed" if errors else "passed_with_warnings" if warnings else "passed",
        "errors": errors,
        "warnings": warnings,
        "checks": checks,
    }
    atomic_write_json(project / "analysis" / "delivery_validation.json", report)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", required=True)
    parser.add_argument("--data-dir")
    parser.add_argument("--ppt-master-dir")
    parser.add_argument("--expected-slides", type=int)
    parser.add_argument("--renderer", choices=("auto", "libreoffice", "none"), default="auto")
    parser.add_argument("--allow-legacy-svg-names", action="store_true")
    parser.add_argument("--allow-missing-facts", action="store_true")
    args = parser.parse_args()
    report = validate_delivery(
        args.project, args.data_dir, args.ppt_master_dir, args.expected_slides,
        args.renderer, args.allow_legacy_svg_names, args.allow_missing_facts,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] != "failed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
