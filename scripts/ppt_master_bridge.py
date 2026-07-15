#!/usr/bin/env python3
"""Locate PPT Master and prepare a project from financial-disclosure-analysis outputs."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


REQUIRED_FILES = (
    "SKILL.md",
    "scripts/project_manager.py",
    "scripts/finalize_svg.py",
    "scripts/svg_to_pptx.py",
)


def _normalize_candidate(path: Path) -> Path | None:
    path = path.expanduser().resolve()
    candidates = (path, path / "skills" / "ppt-master")
    for candidate in candidates:
        if all((candidate / item).is_file() for item in REQUIRED_FILES):
            return candidate
    return None


def resolve_ppt_master(explicit: str | None = None) -> Path:
    """Resolve either a ppt-master repository root or its skill directory."""
    raw_candidates: list[Path] = []
    if explicit:
        raw_candidates.append(Path(explicit))
    if os.environ.get("PPT_MASTER_HOME"):
        raw_candidates.append(Path(os.environ["PPT_MASTER_HOME"]))

    here = Path(__file__).resolve().parent.parent
    raw_candidates.extend(
        [
            here.parent / "ppt-master",
            Path.home() / ".agents" / "skills" / "ppt-master",
            Path.home() / ".claude" / "skills" / "ppt-master",
            Path.home() / ".codex" / "skills" / "ppt-master",
        ]
    )

    checked: list[str] = []
    for raw in raw_candidates:
        checked.append(str(raw.expanduser()))
        resolved = _normalize_candidate(raw)
        if resolved:
            return resolved

    locations = "\n  - ".join(checked) or "(none)"
    raise FileNotFoundError(
        "PPT Master was not found. Clone https://github.com/hugohe3/ppt-master "
        "and pass --ppt-master-dir, or set PPT_MASTER_HOME.\n"
        f"Checked:\n  - {locations}"
    )


def _run(command: list[str], cwd: Path) -> str:
    result = subprocess.run(
        command,
        cwd=cwd,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
    )
    output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part.strip())
    if result.returncode:
        raise RuntimeError(output or f"Command failed with exit code {result.returncode}")
    return output


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _git_commit(path: Path) -> str | None:
    root_result = subprocess.run(
        ["git", "-C", str(path), "rev-parse", "--show-toplevel"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if root_result.returncode != 0 or not root_result.stdout.strip():
        return None
    repo_root = Path(root_result.stdout.strip()).resolve()
    expected_locations = {repo_root, (repo_root / "skills" / "ppt-master").resolve()}
    if path.resolve() not in expected_locations:
        return None
    result = subprocess.run(
        ["git", "-C", str(repo_root), "rev-parse", "HEAD"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    commit = result.stdout.strip()
    return commit if result.returncode == 0 and commit else None


def prepare_project(
    skill_dir: Path,
    project_name: str,
    projects_dir: Path,
    sources: list[Path],
    canvas_format: str,
) -> Path:
    projects_dir = projects_dir.expanduser().resolve()
    projects_dir.mkdir(parents=True, exist_ok=True)
    manager = skill_dir / "scripts" / "project_manager.py"
    source_records: list[dict[str, str]] = []
    for source in sources:
        resolved = source.expanduser().resolve()
        if not resolved.is_file():
            raise FileNotFoundError(f"Source file not found: {resolved}")
        source_records.append(
            {
                "path": str(resolved),
                "filename": resolved.name,
                "sha256": _sha256(resolved),
            }
        )

    init_output = _run(
        [
            sys.executable,
            str(manager),
            "init",
            project_name,
            "--format",
            canvas_format,
            "--dir",
            str(projects_dir),
        ],
        cwd=skill_dir,
    )
    match = re.search(r"\[OK\] Project initialized:\s*(.+)", init_output)
    if not match:
        raise RuntimeError(f"Could not determine PPT Master project path:\n{init_output}")
    project_dir = Path(match.group(1).strip()).resolve()

    # PPT Master's main workflow requires --move. Stage copies so the source
    # evidence and Excel workpaper remain untouched.
    with tempfile.TemporaryDirectory(
        prefix="financial-disclosure-analysis-ppt-handoff-", dir=projects_dir
    ) as tmp:
        staging = Path(tmp)
        staged: list[Path] = []
        for source in sources:
            source = source.expanduser().resolve()
            target = staging / source.name
            counter = 2
            while target.exists():
                target = staging / f"{source.stem}_{counter}{source.suffix}"
                counter += 1
            shutil.copy2(source, target)
            staged.append(target)

        if staged:
            _run(
                [
                    sys.executable,
                    str(manager),
                    "import-sources",
                    str(project_dir),
                    *(str(item) for item in staged),
                    "--move",
                ],
                cwd=skill_dir,
            )

    manifest = {
        "integration": "financial-disclosure-analysis -> ppt-master",
        "ppt_master_skill_dir": str(skill_dir),
        "ppt_master_commit": _git_commit(skill_dir),
        "sources": source_records,
    }
    analysis_dir = project_dir / "analysis"
    analysis_dir.mkdir(parents=True, exist_ok=True)
    (analysis_dir / "financial-disclosure-analysis-handoff.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return project_dir


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--ppt-master-dir",
        help="PPT Master repo root or skills/ppt-master directory; otherwise use PPT_MASTER_HOME/autodiscovery",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("check", help="Validate and print the resolved PPT Master skill directory")

    prepare = subparsers.add_parser(
        "prepare",
        help="Create a PPT Master project and import financial-disclosure-analysis outputs",
    )
    prepare.add_argument("--project-name", required=True)
    prepare.add_argument("--projects-dir", required=True)
    prepare.add_argument("--source", action="append", default=[], help="Source file to import; repeat as needed")
    prepare.add_argument("--format", default="ppt169", help="PPT Master canvas format (default: ppt169)")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        skill_dir = resolve_ppt_master(args.ppt_master_dir)
        if args.command == "check":
            print(skill_dir)
            return 0
        sources = [Path(item) for item in args.source]
        if not sources:
            raise ValueError("prepare requires at least one --source")
        project_dir = prepare_project(
            skill_dir=skill_dir,
            project_name=args.project_name,
            projects_dir=Path(args.projects_dir),
            sources=sources,
            canvas_format=args.format,
        )
        print(project_dir)
        return 0
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
