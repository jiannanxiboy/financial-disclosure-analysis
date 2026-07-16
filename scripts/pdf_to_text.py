#!/usr/bin/env python3
"""
PDF to plain-text converter for financial reports.
Extracts ALL content: paragraphs, headers, and tables rendered as aligned text.
Outputs a single .txt file with page markers for source tracing.

Supports single-file mode (--input/--output) and batch mode (--input-dir/--file-list).
Batch mode uses multiprocessing for parallelism.
"""

import argparse
import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from functools import partial
from pathlib import Path


def _configure_third_party_logging(verbose: bool = False) -> None:
    """Suppress noisy parser warnings unless explicitly requested."""
    level = logging.WARNING if verbose else logging.ERROR
    for logger_name in ("pdfminer", "pdfplumber"):
        logging.getLogger(logger_name).setLevel(level)


def _display_width(text: str) -> int:
    """计算字符串的显示宽度：CJK 字符占 2 列，其余占 1 列。"""
    import unicodedata
    w = 0
    for ch in text:
        w += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return w


def render_table_as_text(table: list[list[str | None]], col_widths: list[int] | None = None) -> str:
    """Render a 2D table list into an aligned plain-text format."""
    if not table:
        return ""

    # Calculate column widths from display width (CJK-aware)
    if col_widths is None:
        col_widths = []
        for row in table:
            for i, cell in enumerate(row):
                cell_text = str(cell).strip() if cell else ""
                dw = _display_width(cell_text)
                if i >= len(col_widths):
                    col_widths.append(dw)
                else:
                    col_widths[i] = max(col_widths[i], dw)

    lines = []
    for row in table:
        cells = []
        for i, cell in enumerate(row):
            cell_text = str(cell).strip() if cell else ""
            target_w = col_widths[i] if i < len(col_widths) else 12
            dw = _display_width(cell_text)
            padding = max(0, target_w - dw)
            cells.append(cell_text + " " * padding)
        lines.append(" | ".join(cells))

    return "\n".join(lines)


def _verify_pdf_readable(pdf_path: str) -> int:
    """验证PDF可读并返回页数。失败则抛出明确异常。"""
    import pdfplumber
    from pdfplumber.utils.exceptions import PdfminerException

    p = Path(pdf_path)
    if not p.exists():
        raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
    size_mb = p.stat().st_size / (1024 * 1024)
    if size_mb < 0.01:
        raise ValueError(f"PDF文件过小 ({size_mb:.2f}MB)，可能是HTML或下载失败: {pdf_path}")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            n = len(pdf.pages)
            if n == 0:
                raise ValueError(f"PDF没有页面: {pdf_path}")
            return n
    except PdfminerException as e:
        raise ValueError(f"PDF解析失败（损坏或非PDF格式）: {pdf_path}\n  {e}")


def _sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _sidecar_path(output_path: str | Path) -> Path:
    return Path(output_path).with_suffix(".pages.json")


def _extract_pdf_once(pdf_path: str, output_path: str, include_tables: bool, tool: str) -> dict:
    """Extract once and return page-level source-location and scan-density metadata."""
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        with open(output_path, "w", encoding="utf-8") as out:
            header = (
                f"# 源文件: {Path(pdf_path).name}\n"
                f"# 总页数: {total_pages}\n"
                f"# 提取工具: {tool}\n"
                f"{'=' * 60}\n\n"
            )
            out.write(header)
            char_cursor = len(header)
            page_records = []

            for page_num, page in enumerate(pdf.pages, 1):
                page_text = page.extract_text() or ""
                chunks = [f"========== 第 {page_num} 页 / 共 {total_pages} 页 ==========\n\n"]
                if page_text:
                    chunks.extend([page_text, "\n"])
                table_count = 0
                if include_tables:
                    tables = page.extract_tables()
                    for t_idx, table in enumerate(tables):
                        if table and len(table) > 0:
                            table_count += 1
                            chunks.append(f"\n--- 第{page_num}页 表格{t_idx + 1} ---\n")
                            rendered = render_table_as_text(table)
                            chunks.extend([rendered, "\n"])
                chunks.append("\n")
                page_chunk = "".join(chunks)
                start = char_cursor
                out.write(page_chunk)
                char_cursor += len(page_chunk)
                text_chars = len("".join(page_text.split()))
                page_records.append({
                    "page": page_num,
                    "char_start": start,
                    "char_end": char_cursor,
                    "text_chars": text_chars,
                    "table_count": table_count,
                    "status": "text" if text_chars else "empty",
                })
    return {"total_pages": total_pages, "pages": page_records, "tool": tool}


def _write_page_sidecar(output_path: str, report: dict) -> None:
    path = _sidecar_path(output_path)
    temp = path.with_name(f".{path.name}.tmp")
    try:
        temp.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def _ocr_command() -> str | None:
    return shutil.which("ocrmypdf")


def _apply_ocr(pdf_path: str, output_pdf: str) -> None:
    command = _ocr_command()
    if not command:
        raise RuntimeError("未找到 ocrmypdf；请安装 OCRmyPDF 后重试，或使用 --ocr-mode detect 标记人工处理")
    result = subprocess.run(
        [command, "--skip-text", "--deskew", "--output-type", "pdf", pdf_path, output_pdf],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout or "OCRmyPDF failed").strip())


def extract_pdf_to_text(
    pdf_path: str,
    output_path: str,
    include_tables: bool = True,
    ocr_mode: str = "detect",
    min_text_chars: int = 30,
    scanned_ratio_threshold: float = 0.3,
):
    """Extract PDF text, write page index, and optionally route scanned pages through OCRmyPDF."""
    _verify_pdf_readable(pdf_path)
    report = _extract_pdf_once(pdf_path, output_path, include_tables, "pdfplumber")
    scanned = [page for page in report["pages"] if page["text_chars"] < min_text_chars]
    ratio = len(scanned) / report["total_pages"] if report["total_pages"] else 0
    needs_ocr = bool(scanned) and ratio >= scanned_ratio_threshold
    report.update({
        "source_file": str(Path(pdf_path).resolve()),
        "source_sha256": _sha256(pdf_path),
        "min_text_chars": min_text_chars,
        "scanned_pages": [page["page"] for page in scanned],
        "scanned_ratio": ratio,
        "needs_ocr": needs_ocr,
        "ocr_mode": ocr_mode,
        "ocr_available": bool(_ocr_command()),
        "ocr_applied": False,
    })

    if needs_ocr and ocr_mode in {"auto", "required"}:
        try:
            with tempfile.TemporaryDirectory(prefix="financial-disclosure-ocr-") as tmp:
                ocr_pdf = str(Path(tmp) / "ocr.pdf")
                _apply_ocr(pdf_path, ocr_pdf)
                report = _extract_pdf_once(ocr_pdf, output_path, include_tables, "ocrmypdf+pdfplumber")
                remaining = [page for page in report["pages"] if page["text_chars"] < min_text_chars]
                report.update({
                    "source_file": str(Path(pdf_path).resolve()),
                    "source_sha256": _sha256(pdf_path),
                    "min_text_chars": min_text_chars,
                    "scanned_pages": [page["page"] for page in remaining],
                    "scanned_ratio": len(remaining) / report["total_pages"],
                    "needs_ocr": bool(remaining),
                    "ocr_mode": ocr_mode,
                    "ocr_available": True,
                    "ocr_applied": True,
                })
        except Exception as exc:
            report["ocr_error"] = str(exc)
            if ocr_mode == "required":
                _write_page_sidecar(output_path, report)
                raise

    _write_page_sidecar(output_path, report)
    return report["total_pages"]


# ── 路径工具 ──

def _to_native_path(raw: str) -> Path:
    """将 MSYS2/Unix 风格路径转为 Windows 原生路径（如 /c/foo → C:\\foo）。"""
    p = Path(raw)
    if p.is_absolute():
        return p
    # MSYS2 风格: /c/Users/... → C:\Users\...
    import platform
    if platform.system() == "Windows" and raw.startswith("/"):
        parts = raw.lstrip("/").split("/", 1)
        if len(parts) >= 1 and len(parts[0]) == 1 and parts[0].isalpha():
            drive_letter = parts[0].upper()
            rest = parts[1] if len(parts) > 1 else ""
            candidate = Path(f"{drive_letter}:/{rest}")
            if candidate.exists():
                return candidate
    return p


# ── 批量处理 worker ──

def _extract_one(pdf_path: str, output_path: str, include_tables: bool = True,
                 skip_existing: bool = False, verbose: bool = False,
                 ocr_mode: str = "detect", min_text_chars: int = 30,
                 scanned_ratio_threshold: float = 0.3) -> dict:
    """单个PDF提取的独立入口，供进程池调用。返回结果字典。"""
    _configure_third_party_logging(verbose)
    t0 = time.perf_counter()
    try:
        sidecar = _sidecar_path(output_path)
        if (
            skip_existing
            and Path(output_path).exists()
            and Path(output_path).stat().st_size > 0
            and sidecar.exists()
        ):
            elapsed = time.perf_counter() - t0
            metadata = json.loads(sidecar.read_text(encoding="utf-8"))
            return {"pdf": pdf_path, "output": output_path, "pages": metadata.get("total_pages", 0), "ok": True,
                    "skipped": True, "error": "", "elapsed": elapsed,
                    "needs_ocr": metadata.get("needs_ocr", False),
                    "ocr_applied": metadata.get("ocr_applied", False)}

        pages = extract_pdf_to_text(
            pdf_path, output_path, include_tables, ocr_mode,
            min_text_chars, scanned_ratio_threshold,
        )
        metadata = json.loads(_sidecar_path(output_path).read_text(encoding="utf-8"))
        elapsed = time.perf_counter() - t0
        return {"pdf": pdf_path, "output": output_path, "pages": pages, "ok": True,
                "skipped": False, "error": "", "elapsed": elapsed,
                "needs_ocr": metadata.get("needs_ocr", False),
                "ocr_applied": metadata.get("ocr_applied", False)}
    except Exception as e:
        elapsed = time.perf_counter() - t0
        return {"pdf": pdf_path, "output": output_path, "pages": 0, "ok": False,
                "skipped": False, "error": str(e), "elapsed": elapsed,
                "needs_ocr": False, "ocr_applied": False}


def batch_extract(pdf_list: list[tuple[str, str]], include_tables: bool = True,
                  workers: int | None = None, skip_existing: bool = False,
                  quiet: bool = False, verbose: bool = False,
                  ocr_mode: str = "detect", min_text_chars: int = 30,
                  scanned_ratio_threshold: float = 0.3) -> dict:
    """并行批量提取PDF，返回汇总统计。"""
    if workers is None:
        cpu_count = os.cpu_count() or 4
        # 保守默认：留1核给系统，上限4进程（pdfplumber吃内存，再多收益递减）
        workers = min(4, max(1, cpu_count - 1))

    n = len(pdf_list)
    worker_fn = partial(
        _extract_one,
        include_tables=include_tables,
        skip_existing=skip_existing,
        verbose=verbose,
        ocr_mode=ocr_mode,
        min_text_chars=min_text_chars,
        scanned_ratio_threshold=scanned_ratio_threshold,
    )

    if not quiet:
        print(f"批量提取: {n} 个文件, {workers} 个工作进程\n")

    results = []
    t0 = time.perf_counter()
    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(worker_fn, src, dst): src for src, dst in pdf_list}
        completed = 0
        for future in as_completed(futures):
            r = future.result()
            results.append(r)
            completed += 1
            status = "跳过" if r["skipped"] else ("OK" if r["ok"] else "FAIL")
            if not quiet:
                src_name = Path(r["pdf"]).name
                elapsed = r["elapsed"]
                print(f"[{completed}/{n}] {status}  {src_name}  ({elapsed:.1f}s)"
                      + (f"  {r['error']}" if not r["ok"] else ""))

    total_elapsed = time.perf_counter() - t0
    ok_list = [r for r in results if r["ok"]]
    fail_list = [r for r in results if not r["ok"]]
    skipped_list = [r for r in results if r["skipped"]]
    total_pages = sum(r["pages"] for r in ok_list)
    needs_ocr = [r for r in ok_list if r.get("needs_ocr")]
    ocr_applied = [r for r in ok_list if r.get("ocr_applied")]

    # 汇总
    print(f"\n{'=' * 60}")
    print(f"总文件数: {n}  成功: {len(ok_list)}  失败: {len(fail_list)}  跳过: {len(skipped_list)}")
    print(f"总页数: {total_pages:,}  总耗时: {total_elapsed:.1f}s"
          + (f"  平均: {total_elapsed / n:.1f}s/文件" if n else ""))
    if needs_ocr or ocr_applied:
        print(f"OCR已应用: {len(ocr_applied)}  仍需OCR/人工复核: {len(needs_ocr)}")

    if fail_list:
        print(f"\n失败列表:")
        for r in fail_list:
            print(f"  {Path(r['pdf']).name} — {r['error']}")

    return {"total": n, "ok": len(ok_list), "fail": len(fail_list),
            "skipped": len(skipped_list), "total_pages": total_pages,
            "ocr_applied": len(ocr_applied), "needs_ocr": len(needs_ocr),
            "elapsed": total_elapsed}


def main():
    parser = argparse.ArgumentParser(
        description="将财务报告PDF完整提取为纯文本（包含文字和表格）\n"
                    "单文件模式: --input xxx.pdf --output xxx.txt\n"
                    "批量模式:   --input-dir ./pdfs/ --output-dir ./txts/"
    )
    # 单文件模式
    parser.add_argument("--input", "-i", help="PDF文件路径")
    parser.add_argument("--output", "-o", help="输出文本文件路径 (.txt)")
    # 批量模式
    parser.add_argument("--input-dir", help="包含PDF文件的目录（扫描所有 *.pdf）")
    parser.add_argument("--file-list", help="每行一个PDF路径的文本文件")
    parser.add_argument("--output-dir", help="批量模式的输出目录")
    parser.add_argument("--workers", "-w", type=int, default=None,
                        help=f"并行进程数（默认: CPU核心数-1且≤4，当前默认 {min(4, max(1, (os.cpu_count() or 4) - 1))}）")
    parser.add_argument("--skip-existing", action="store_true", help="跳过已存在的输出文件")
    parser.add_argument("--quiet", "-q", action="store_true", help="静默模式，仅输出汇总")
    parser.add_argument("--verbose", action="store_true", help="显示PDF解析库的详细警告日志")
    parser.add_argument("--ocr-mode", choices=("off", "detect", "auto", "required"), default="detect",
                        help="扫描件处理：off忽略、detect标记、auto有工具时OCR、required无OCR则失败")
    parser.add_argument("--min-text-chars", type=int, default=30,
                        help="每页低于该非空白字符数时视为疑似扫描页（默认30）")
    parser.add_argument("--scanned-ratio-threshold", type=float, default=0.3,
                        help="疑似扫描页比例达到该值时触发OCR路由（默认0.3）")
    # 通用
    parser.add_argument("--no-tables", action="store_true", help="跳过表格提取（仅提取文字）")

    args = parser.parse_args()
    _configure_third_party_logging(args.verbose)

    # ── 判断模式 ──
    batch_mode = bool(args.input_dir or args.file_list)

    if batch_mode and (args.input or args.output):
        parser.error("批量模式（--input-dir/--file-list）与单文件模式（--input/--output）互斥")

    if not batch_mode and not (args.input and args.output):
        parser.error("请指定 --input/--output（单文件）或 --input-dir/--file-list（批量）")

    if batch_mode and not args.output_dir:
        parser.error("批量模式需要指定 --output-dir")

    # ── 单文件模式（保持原行为） ──
    if not batch_mode:
        input_path = _to_native_path(args.input)
        if not input_path.exists():
            print(f"错误: PDF文件不存在: {args.input}", file=sys.stderr)
            sys.exit(1)

        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        pages = extract_pdf_to_text(
            pdf_path=str(input_path),
            output_path=str(output_path),
            include_tables=not args.no_tables,
            ocr_mode=args.ocr_mode,
            min_text_chars=args.min_text_chars,
            scanned_ratio_threshold=args.scanned_ratio_threshold,
        )
        print(f"共处理 {pages} 页")
        metadata = json.loads(_sidecar_path(output_path).read_text(encoding="utf-8"))
        if metadata.get("needs_ocr"):
            print(f"警告: {len(metadata['scanned_pages'])} 页疑似扫描页，详见 {_sidecar_path(output_path)}")
        return

    # ── 批量模式 ──
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # 收集PDF列表
    pdf_list: list[tuple[str, str]] = []

    if args.input_dir:
        input_dir = _to_native_path(args.input_dir)
        if not input_dir.is_dir():
            print(f"错误: 目录不存在: {args.input_dir}", file=sys.stderr)
            sys.exit(1)
        for pdf_path in sorted(input_dir.glob("*.pdf")):
            out_path = output_dir / f"{pdf_path.stem}.txt"
            pdf_list.append((str(pdf_path), str(out_path)))

    if args.file_list:
        flist_path = _to_native_path(args.file_list)
        if not flist_path.exists():
            print(f"错误: 文件列表不存在: {args.file_list}", file=sys.stderr)
            sys.exit(1)
        with open(flist_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                src = _to_native_path(line)
                if not src.is_absolute():
                    src = flist_path.parent / src
                if not src.exists():
                    print(f"警告: 文件不存在，跳过: {src}", file=sys.stderr)
                    continue
                dst = output_dir / f"{src.stem}.txt"
                pdf_list.append((str(src), str(dst)))

    if not pdf_list:
        print("没有找到PDF文件", file=sys.stderr)
        sys.exit(1)

    summary = batch_extract(
        pdf_list=pdf_list,
        include_tables=not args.no_tables,
        workers=args.workers,
        skip_existing=args.skip_existing,
        quiet=args.quiet,
        verbose=args.verbose,
        ocr_mode=args.ocr_mode,
        min_text_chars=args.min_text_chars,
        scanned_ratio_threshold=args.scanned_ratio_threshold,
    )

    if summary["fail"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
