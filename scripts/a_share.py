#!/usr/bin/env python3
"""A股公告搜索与下载 — 通过巨潮资讯网。"""

import sys
import json
import os
import time
from pathlib import Path

# 确保能导入同目录下的 _common 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests

from _common import (
    DEFAULT_UA, CNINFO_API_TIMEOUT,
    download_pdf as _download_pdf, poll_for_button, launch_browser,
)

CNINFO_API = "http://www.cninfo.com.cn/new/hisAnnouncement/query"

# CNINFO 公告分类 — 比关键词匹配更可靠的分类体系
CATEGORY_ANNUAL   = "category_ndbg_szsh"      # 年度报告
CATEGORY_SEMI     = "category_bndbg_szsh"      # 半年度报告
CATEGORY_Q1       = "category_rptfirst_szsh"   # 一季报
CATEGORY_Q3       = "category_rptthird_szsh"   # 三季报
_CACHE_DIR = Path(__file__).resolve().parent
_CACHE_FILE = _CACHE_DIR / ".orgid_cache.json"
_CACHE_TTL = 86400 * 90


# ── 缓存 ──

def _load_cache() -> dict:
    if _CACHE_FILE.exists():
        try:
            return json.loads(_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_cache(cache: dict):
    _CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    _CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def _cache_get(code: str) -> dict | None:
    entry = _load_cache().get(code)
    if not entry:
        return None
    if time.time() - entry.get("ts", 0) > _CACHE_TTL:
        return None
    return entry.get("info")


def _cache_set(code: str, info: dict):
    cache = _load_cache()
    cache[code] = {"ts": time.time(), "info": info}
    _save_cache(cache)


# ── API ──

def _api_headers():
    return {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "User-Agent": DEFAULT_UA,
        "Referer": "http://www.cninfo.com.cn/",
        "Origin": "http://www.cninfo.com.cn",
    }


def _query(stock_code: str, org_id: str, column: str, plate: str,
           category: str = "", se_date: str = "", timeout: int = 30) -> list[dict]:
    data = {
        "stock": f"{stock_code},{org_id}",
        "tabName": "fulltext", "pageSize": "30", "pageNum": "1",
        "column": column, "plate": plate,
        "category": category, "seDate": se_date, "isHLtitle": "true",
    }
    resp = requests.post(CNINFO_API, headers=_api_headers(), data=data, timeout=timeout, verify=False)
    if resp.status_code != 200:
        return []
    return (resp.json().get("announcements") or [])


# ── orgId ──

def _discover_stock_info(stock_code: str, timeout: int = 30) -> dict | None:
    cached = _cache_get(stock_code)
    if cached:
        return cached
    column, plate = ("sse", "sh") if stock_code.startswith(("60", "68")) else ("szse", "sz")
    data = {
        "stock": "", "tabName": "fulltext", "pageSize": "30", "pageNum": "1",
        "column": column, "plate": plate,
        "category": "category_ndbg_szsh", "searchkey": stock_code, "isHLtitle": "true",
    }
    try:
        resp = requests.post(CNINFO_API, headers=_api_headers(), data=data, timeout=timeout, verify=False)
        if resp.status_code != 200:
            return None
        for a in (resp.json().get("announcements") or []):
            if a.get("secCode") == stock_code:
                info = {"orgId": a["orgId"], "column": column, "plate": plate, "name": a.get("secName", "")}
                _cache_set(stock_code, info)
                return info
    except Exception:
        pass
    return None


# ── 搜索 ──

def search(stock_code: str, keywords: list[str], se_date: str = "",
           timeout: int = 30) -> list[dict]:
    """搜索公告。keywords为标题必须同时包含的词，se_date格式"YYYY-MM-DD~YYYY-MM-DD"。"""
    info = _discover_stock_info(stock_code, timeout=timeout)
    if not info:
        return []
    announcements = _query(stock_code, info["orgId"], info["column"], info["plate"],
                           se_date=se_date, timeout=timeout)
    results = []
    for a in announcements:
        title = a.get("announcementTitle", "")
        if not all(kw in title for kw in keywords):
            continue
        ann_time = a.get("announcementTime", "")
        org = a.get("orgId", info["orgId"])
        results.append({
            "orgId": org,
            "announcementId": a.get("announcementId", ""),
            "announcementTime": ann_time,
            "stockCode": stock_code,
            "companyName": info.get("name", ""),
            "title": title.strip(),
            "url": f"http://www.cninfo.com.cn/new/disclosure/detail?orgId={org}&announcementId={a.get('announcementId', '')}&announcementTime={ann_time}",
        })
    return results


def search_by_category(stock_code: str, category: str = CATEGORY_ANNUAL,
                       year: int | None = None, timeout: int = CNINFO_API_TIMEOUT) -> list[dict]:
    """按CNINFO公告分类搜索，可选ASCII年份过滤。

    不依赖中文关键词匹配，从根本上解决bash/CLI传参编码问题。
    分类常量: CATEGORY_ANNUAL(年报), CATEGORY_SEMI(中报),
             CATEGORY_Q1(一季报), CATEGORY_Q3(三季报)。
    """
    info = _discover_stock_info(stock_code, timeout=timeout)
    if not info:
        return []
    announcements = _query(stock_code, info["orgId"], info["column"], info["plate"],
                           category=category, timeout=timeout)
    results = []
    for a in announcements:
        title = a.get("announcementTitle", "")
        # 仅用ASCII年份过滤（永不因编码损坏）
        if year is not None and str(year) not in title:
            continue
        # 排除摘要版（含"摘要"二字）
        if "摘要" in title:
            continue
        ann_time = a.get("announcementTime", "")
        org = a.get("orgId", info["orgId"])
        results.append({
            "orgId": org,
            "announcementId": a.get("announcementId", ""),
            "announcementTime": ann_time,
            "stockCode": stock_code,
            "companyName": info.get("name", ""),
            "title": title.strip(),
            "url": f"http://www.cninfo.com.cn/new/disclosure/detail?"
                   f"orgId={org}&announcementId={a.get('announcementId', '')}"
                   f"&announcementTime={ann_time}",
        })
    return results


# ── 下载 ──

def _extract_pdf_url(page) -> str | None:
    link = page.query_selector('a[href*=".PDF"]')
    return link.get_attribute("href") if link else None


def download(announcement_url: str, output_path: str | Path) -> bool:
    """从公告详情页提取PDF直链并下载，带重试。"""
    from playwright.sync_api import TimeoutError as PwTimeout

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    for attempt in range(3):
        try:
            pw, browser, context, page = launch_browser()
            page.goto(announcement_url, wait_until="domcontentloaded", timeout=30000)

            if not poll_for_button(page, ["公告下载", "下载"]):
                context.close(); browser.close(); pw.stop()
                if attempt < 2: time.sleep([1, 3, 9][attempt])
                continue

            pdf_url = _extract_pdf_url(page)
            context.close(); browser.close(); pw.stop()

            if pdf_url and _download_pdf(pdf_url, path):
                return True
            if attempt < 2:
                time.sleep([1, 3, 9][attempt])
        except (PwTimeout, Exception):
            if attempt < 2:
                time.sleep([1, 3, 9][attempt])
    return False


def download_batch(tasks: list[tuple[str, str]]) -> dict[str, bool]:
    """批量下载，复用浏览器。"""
    from playwright.sync_api import TimeoutError as PwTimeout

    results = {}
    try:
        pw, browser, _, _ = launch_browser()
        for url, output in tasks:
            path = Path(output)
            path.parent.mkdir(parents=True, exist_ok=True)

            for attempt in range(3):
                try:
                    context = browser.new_context(
                        user_agent=DEFAULT_UA,
                        viewport={"width": 1920, "height": 1080}, locale="zh-CN")
                    page = context.new_page()
                    page.goto(url, wait_until="domcontentloaded", timeout=30000)

                    if not poll_for_button(page, ["公告下载", "下载"]):
                        page.close(); context.close(); break

                    pdf_url = _extract_pdf_url(page)
                    page.close(); context.close()

                    if pdf_url and _download_pdf(pdf_url, path):
                        results[str(path)] = True; break
                    if attempt < 2:
                        time.sleep([1, 3, 9][attempt])
                except (PwTimeout, Exception):
                    try: page.close(); context.close()
                    except Exception: pass
                    if attempt < 2: time.sleep([1, 3, 9][attempt])
            else:
                results[str(path)] = False
        browser.close(); pw.stop()
    except Exception:
        for url, output in tasks:
            results[str(output)] = download(url, output)
    return results


# ── CLI ──

def main():
    import argparse
    parser = argparse.ArgumentParser(description="A股公告搜索与下载")
    sub = parser.add_subparsers(dest="cmd")

    sp = sub.add_parser("search", help="搜索公告")
    sp.add_argument("--code", required=True)
    sp.add_argument("--keywords", required=True, nargs="+", help="标题关键词(须同时命中)")
    sp.add_argument("--se-date", default="", help="日期范围 YYYY-MM-DD~YYYY-MM-DD")
    sp.add_argument("--json", action="store_true")

    sa = sub.add_parser("search-annual", help="搜索年报(按CNINFO分类，不依赖中文关键词)")
    sa.add_argument("--code", help="单股代码")
    sa.add_argument("--codes", nargs="+", help="多股代码（批量搜索+下载）")
    sa.add_argument("--year", type=int, help="年份(如2025)，不填则返回最新年报")
    sa.add_argument("--json", action="store_true")
    sa.add_argument("--download-dir", "-d", help="下载目录（可选，指定后自动下载第一个结果）")

    bp = sub.add_parser("batch", help="批量搜索并下载")
    bp.add_argument("--codes", required=True, nargs="+")
    bp.add_argument("--keywords", required=True, nargs="+", help="标题关键词(须同时命中)")
    bp.add_argument("--se-date", default="", help="日期范围 YYYY-MM-DD~YYYY-MM-DD")
    bp.add_argument("--json", action="store_true")
    bp.add_argument("--download-dir", "-d")
    bp.add_argument("--output-map", "-o")

    dl = sub.add_parser("download", help="下载PDF")
    dl.add_argument("--url", required=True)
    dl.add_argument("--output", required=True)

    args = parser.parse_args()

    if args.cmd == "search":
        results = search(args.code, args.keywords, se_date=args.se_date)
        if args.json:
            print(json.dumps(results, ensure_ascii=False, indent=2))
        elif results:
            for r in results:
                print(f"标题: {r['title']}\n时间: {r['announcementTime']}\nURL: {r['url']}\n")
        else:
            print("未找到", file=sys.stderr)
            sys.exit(1)

    elif args.cmd == "search-annual":
        codes = args.codes if args.codes else [args.code]
        if not codes or not codes[0]:
            print("请指定 --code 或 --codes", file=sys.stderr)
            sys.exit(1)

        year = args.year
        all_results: dict[str, list[dict]] = {}
        for code in codes:
            results = search_by_category(code, category=CATEGORY_ANNUAL, year=year)
            all_results[code] = results
            if not args.download_dir:
                if args.json:
                    continue
                elif results:
                    for r in results:
                        print(f"[{code}] 标题: {r['title']}\n时间: {r['announcementTime']}\nURL: {r['url']}\n")
                else:
                    yr = year if year else "最新"
                    print(f"[{code}] 未找到{yr}年报", file=sys.stderr)

        if args.json and not args.download_dir:
            print(json.dumps(all_results, ensure_ascii=False, indent=2))

        if args.download_dir:
            download_dir = Path(args.download_dir)
            tasks = []
            for code, results in all_results.items():
                if results:
                    r = results[0]
                    name = r.get("companyName", "").replace("<em>", "").replace("</em>", "")
                    fname = f"{code}_{name}_{year}年报.pdf" if name else f"{code}_{year}年报.pdf"
                    output = download_dir / fname
                    tasks.append((r["url"], str(output)))
                else:
                    yr = year if year else "最新"
                    print(f"[{code}] 未找到{yr}年报，跳过下载", file=sys.stderr)

            if tasks:
                print(f"批量下载 {len(tasks)} 个文件...")
                dl_results = download_batch(tasks)
                ok = sum(1 for v in dl_results.values() if v)
                print(f"下载完成: {ok}/{len(tasks)} 成功")
                for path, success in dl_results.items():
                    print(f"  {'OK' if success else 'FAIL'} -> {path}")
            if args.json:
                print(json.dumps(all_results, ensure_ascii=False, indent=2))

    elif args.cmd == "batch":
        results = {}
        for code in args.codes:
            code = code.strip()
            if code:
                results[code] = search(code, args.keywords, se_date=args.se_date)
        if args.json:
            print(json.dumps(results, ensure_ascii=False, indent=2))

        if args.download_dir:
            download_dir = Path(args.download_dir)
            tasks, mapping = [], {}
            for code, items in results.items():
                if items:
                    r = items[0]
                    name = r["companyName"].replace("<em>", "").replace("</em>", "")
                    fname = f"{code}_{name}.pdf" if name else f"{code}.pdf"
                    output = download_dir / fname
                    tasks.append((r["url"], str(output)))
                    mapping[code] = {"url": r["url"], "output": str(output)}
            if tasks:
                dl_results = download_batch(tasks)
                ok = sum(1 for v in dl_results.values() if v)
                print(f"成功: {ok}/{len(tasks)}")
            if args.output_map:
                Path(args.output_map).write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")

    elif args.cmd == "download":
        ok = download(args.url, args.output)
        print("OK" if ok else "FAIL")
        sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
