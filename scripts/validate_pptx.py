#!/usr/bin/env python3
"""Validate PPTX package structure and optionally round-trip render with LibreOffice."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree


def _numbered_parts(names: list[str], prefix: str, stem: str) -> list[str]:
    pattern = re.compile(rf"^{re.escape(prefix)}/{re.escape(stem)}(\d+)\.xml$")
    return sorted(
        (name for name in names if pattern.match(name)),
        key=lambda name: int(pattern.match(name).group(1)),
    )


def inspect_pptx(path: str | Path) -> dict:
    pptx = Path(path).resolve()
    errors: list[str] = []
    if not pptx.is_file():
        return {"path": str(pptx), "ok": False, "errors": ["PPTX file not found"]}
    try:
        with zipfile.ZipFile(pptx) as archive:
            names = archive.namelist()
            required = {"[Content_Types].xml", "ppt/presentation.xml"}
            errors.extend(f"missing package part: {name}" for name in sorted(required - set(names)))
            slides = _numbered_parts(names, "ppt/slides", "slide")
            notes = _numbered_parts(names, "ppt/notesSlides", "notesSlide")
            charts = _numbered_parts(names, "ppt/charts", "chart")
            table_count = 0
            for name in slides:
                payload = archive.read(name)
                try:
                    slide_root = ElementTree.fromstring(payload)
                    table_count += len(slide_root.findall(
                        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}tbl"
                    ))
                except ElementTree.ParseError as exc:
                    errors.append(f"invalid XML {name}: {exc}")
            dimensions = {}
            if "ppt/presentation.xml" in names:
                root = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
                size = root.find("{http://schemas.openxmlformats.org/presentationml/2006/main}sldSz")
                if size is not None:
                    dimensions = {"cx": int(size.get("cx", 0)), "cy": int(size.get("cy", 0))}
            if not slides:
                errors.append("presentation contains no slides")
    except (OSError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        return {"path": str(pptx), "ok": False, "errors": [str(exc)]}
    return {
        "path": str(pptx),
        "size": pptx.stat().st_size,
        "slides": len(slides),
        "notes": len(notes),
        "charts": len(charts),
        "tables": table_count,
        "dimensions": dimensions,
        "ok": not errors,
        "errors": errors,
    }


def find_libreoffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    return next((item for item in candidates if item and Path(item).is_file()), None)


def render_with_libreoffice(path: str | Path) -> dict:
    executable = find_libreoffice()
    if not executable:
        return {"renderer": "libreoffice", "status": "unavailable", "ok": False}
    pptx = Path(path).resolve()
    with tempfile.TemporaryDirectory(prefix="financial-disclosure-pptx-render-") as tmp:
        result = subprocess.run(
            [executable, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(pptx)],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )
        pdf = Path(tmp) / f"{pptx.stem}.pdf"
        if result.returncode or not pdf.is_file():
            return {
                "renderer": "libreoffice",
                "status": "failed",
                "ok": False,
                "error": (result.stderr or result.stdout or "PDF was not produced").strip(),
            }
        try:
            import pdfplumber
            with pdfplumber.open(pdf) as document:
                pages = len(document.pages)
        except Exception as exc:
            return {"renderer": "libreoffice", "status": "failed", "ok": False, "error": str(exc)}
        return {"renderer": "libreoffice", "status": "rendered", "ok": True, "pages": pages}


def validate(
    path: str | Path,
    expected_slides: int | None = None,
    require_notes: bool = False,
    renderer: str = "auto",
    require_render: bool = False,
) -> dict:
    report = inspect_pptx(path)
    errors = report.setdefault("errors", [])
    if expected_slides is not None and report.get("slides") != expected_slides:
        errors.append(f"slide count mismatch: expected {expected_slides}, got {report.get('slides')}")
    if require_notes and report.get("notes") != report.get("slides"):
        errors.append(f"speaker notes mismatch: {report.get('notes')} notes for {report.get('slides')} slides")

    render_report = {"renderer": renderer, "status": "skipped", "ok": True}
    if renderer in {"auto", "libreoffice"}:
        render_report = render_with_libreoffice(path)
        if render_report["status"] == "unavailable" and not require_render and renderer == "auto":
            render_report["ok"] = True
            render_report["status"] = "skipped-unavailable"
        elif not render_report.get("ok"):
            errors.append(f"{render_report.get('renderer')} rendering failed or is unavailable")
        if render_report.get("status") == "rendered" and render_report.get("pages") != report.get("slides"):
            errors.append(
                f"rendered page count mismatch: {render_report.get('pages')} for {report.get('slides')} slides"
            )
        if require_render and not render_report.get("ok"):
            errors.append("required cross-application rendering did not succeed")
    report["render"] = render_report
    report["ok"] = not errors
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", help="PPTX file to validate")
    parser.add_argument("--expected-slides", type=int)
    parser.add_argument("--require-notes", action="store_true")
    parser.add_argument("--renderer", choices=("auto", "libreoffice", "none"), default="auto")
    parser.add_argument("--require-render", action="store_true")
    parser.add_argument("--output", help="Optional JSON report path")
    args = parser.parse_args()
    report = validate(
        args.pptx,
        expected_slides=args.expected_slides,
        require_notes=args.require_notes,
        renderer=args.renderer,
        require_render=args.require_render,
    )
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    print(payload)
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(payload + "\n", encoding="utf-8")
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
