#!/usr/bin/env python3
"""港股公告搜索与下载 — 通过港交所披露易。"""

from __future__ import annotations

import asyncio
import json
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

# 确保能导入同目录下的 _common 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests
from playwright.async_api import async_playwright, TimeoutError as PwTimeout

from _common import (
    DEFAULT_UA, atomic_write_json, verify_pdf, download_pdf_stream, launch_browser,
    HKEX_PAGE_LOAD_TIMEOUT, HKEX_TABLE_TIMEOUT, HKEX_PDF_DOWNLOAD_TIMEOUT,
    RETRY_DELAYS_LONG, RETRY_DELAYS_SHORT, INTER_CODE_DELAY,
)

HKEX_SEARCH_URL = "https://www1.hkexnews.hk/search/titlesearch.xhtml?lang=zh"
HKEX_PREFIX_URL = "https://www1.hkexnews.hk/search/prefix.do"
_CACHE_FILE = Path(__file__).resolve().with_name(".hkex_stockid_cache.json")
_CACHE_TTL = 86400 * 90
_CACHE_LOCK = threading.Lock()

# 年报搜索重试次数
_ANNUAL_SEARCH_RETRIES = 3


# ── 缓存 ──

def _load_cache() -> dict:
    if _CACHE_FILE.exists():
        try:
            return json.loads(_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_cache(cache: dict):
    atomic_write_json(_CACHE_FILE, cache)


def _cache_get(code: str) -> dict | None:
    entry = _load_cache().get(code)
    if not entry:
        return None
    if time.time() - entry.get("ts", 0) > _CACHE_TTL:
        return None
    return entry.get("info")


def _cache_set(code: str, info: dict):
    with _CACHE_LOCK:
        cache = _load_cache()
        cache[code] = {"ts": time.time(), "info": info}
        _save_cache(cache)


# ── HTTP ──

_http: requests.Session | None = None


def _get_http() -> requests.Session:
    global _http
    if _http is None:
        _http = requests.Session()
        _http.headers.update({"User-Agent": DEFAULT_UA})
        _http.verify = False
    return _http


# ── Stock ID ──

def _normalize_code(code: str) -> str:
    """标准化股票代码：补零到5位。'688'→'00688', '1109'→'01109'。"""
    code = code.strip()
    if len(code) < 5:
        code = code.zfill(5)
    return code


def lookup_stock_id(code: str, timeout: int = 15) -> dict | None:
    """通过prefix.do将股票代码映射为HKEX内部stockId。
    自动标准化代码格式，对4位代码先尝试补零，失败再尝试原码。
    """
    normalized = _normalize_code(code)
    # 先查缓存（用标准化key）
    for key in (normalized, code):
        if key != normalized:
            continue
        cached = _cache_get(key)
        if cached:
            return cached

    # 尝试查找：先用标准化代码，不行再用原码
    for attempt_code in (normalized, code):
        if attempt_code == normalized and attempt_code == code:
            # 相同代码只试一次
            pass
        for retry_idx in range(len(RETRY_DELAYS_SHORT) + 1):
            try:
                resp = _get_http().get(
                    HKEX_PREFIX_URL,
                    params={"callback": "j", "lang": "ZH", "type": "A",
                            "name": attempt_code, "market": "SEHK"},
                    timeout=timeout,
                )
                resp.encoding = "utf-8"
                raw = resp.text.strip()
                a, b = raw.find("("), raw.rfind(")")
                if a == -1 or b == -1:
                    break
                data = json.loads(raw[a + 1:b])
                for s in data.get("stockInfo", []):
                    matched_code = s.get("code", "")
                    if matched_code == attempt_code or matched_code == _normalize_code(attempt_code):
                        info = {"code": matched_code, "name": s.get("name", ""),
                                "stockId": s.get("stockId")}
                        _cache_set(normalized, info)  # 统一用标准化key缓存
                        return info
                break  # 请求成功但未匹配，不重试
            except (requests.RequestException, json.JSONDecodeError):
                if retry_idx < len(RETRY_DELAYS_SHORT):
                    time.sleep(RETRY_DELAYS_SHORT[retry_idx])
    return None


def lookup_stock_ids(codes: list[str], max_workers: int = 5) -> dict[str, dict]:
    results: dict[str, dict] = {}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {ex.submit(lookup_stock_id, c): c for c in codes}
        for f in as_completed(futures):
            info = f.result()
            if info:
                results[futures[f]] = info
    return results


# ── 下载 ──

def download_pdf(url: str, output_path: str | Path, timeout: int = 180) -> bool:
    """从HKEX直链下载PDF，流式传输，带验证与重试。"""
    return download_pdf_stream(url, output_path, timeout=timeout)


# ── 异步搜索 ──

async def _do_search(page, stock_code, stock_name, from_date="", to_date="",
                      max_results=200):
    """在HKEX搜索公告。带重试、双通道日期设置、长超时。
    不再在JS侧做关键词过滤（回避bash中文编码问题）。
    """

    # ── 页面加载（带重试，覆盖 timeout / ERR_CONNECTION_CLOSED / ERR_TIMED_OUT 等）──
    for attempt in range(_ANNUAL_SEARCH_RETRIES):
        try:
            await page.goto(HKEX_SEARCH_URL, wait_until="load",
                          timeout=HKEX_PAGE_LOAD_TIMEOUT)
            break
        except Exception:
            if attempt == _ANNUAL_SEARCH_RETRIES - 1:
                raise  # 抛给外层 _search_one 处理
            await asyncio.sleep(RETRY_DELAYS_LONG[attempt])

    try:
        await page.wait_for_selector("#searchStockCode", state="visible", timeout=15000)
    except PwTimeout:
        pass

    # Cookie 弹窗
    try:
        accept = page.locator("#onetrust-accept-btn-handler")
        await accept.wait_for(state="visible", timeout=3000)
        await accept.click()
        await accept.wait_for(state="hidden", timeout=5000)
    except (PwTimeout, Exception):
        pass

    # 填写股票代码
    stock_input = page.locator("#searchStockCode")
    if not await stock_input.is_visible():
        return []
    await stock_input.click()
    await stock_input.fill(stock_code)

    # 选择自动补全建议
    try:
        await page.wait_for_selector("tr.autocomplete-suggestion", state="visible", timeout=10000)
    except PwTimeout:
        pass
    await page.evaluate(
        """(name) => {
            const rows = document.querySelectorAll("tr.autocomplete-suggestion");
            for (const r of rows) { if (r.textContent.includes(name)) { r.click(); return; } }
        }""", stock_name)
    await page.wait_for_timeout(1000)

    # ── 日期筛选（双通道：隐藏字段 + 可见控件）──
    if from_date or to_date:
        await page.evaluate(
            """([fd, td]) => {
                const f = document.forms["TitleSearchPanel"];
                // 通道1: 隐藏字段（表单提交值）
                if (fd) f.from.value = fd;
                if (td) f.to.value = td;
                // 通道2: 可见日期控件（JS datepicker 验证用）
                const fromEl = document.querySelector("#searchDate-From");
                const toEl = document.querySelector("#searchDate-To");
                if (fd && fromEl) {
                    fromEl.value = fd.replace(/-/g, '/');
                    fromEl.dispatchEvent(new Event('change', {bubbles: true}));
                }
                if (td && toEl) {
                    toEl.value = td.replace(/-/g, '/');
                    toEl.dispatchEvent(new Event('change', {bubbles: true}));
                }
            }""", [from_date, to_date])
        await page.wait_for_timeout(300)

    # 点击搜索
    await page.evaluate("""() => {
        for (const a of document.querySelectorAll("a")) {
            if (a.textContent.trim() === "搜尋") { a.click(); return; }
        }
    }""")

    # 等待结果
    try:
        await page.wait_for_load_state("networkidle", timeout=HKEX_TABLE_TIMEOUT)
    except PwTimeout:
        pass
    try:
        await page.wait_for_selector("table.table", state="visible", timeout=15000)
    except PwTimeout:
        pass

    # ── 提取数据（纯提取，不做文本过滤）──
    return await page.evaluate(
        """(maxRows) => {
            const table = document.querySelector("table.table");
            if (!table) return [];
            const rows = table.querySelectorAll("tbody tr");
            const out = [];
            for (const tr of rows) {
                if (out.length >= maxRows) break;
                const cells = tr.querySelectorAll("td");
                if (cells.length < 4) continue;
                const date = (cells[0].textContent || "").trim();
                const code = (cells[1].textContent || "").trim();
                const name = (cells[2].textContent || "").trim();
                const link = cells[3].querySelector("a");
                const title = (cells[3].textContent || "").replace(/\\s+/g, " ").trim();
                const href = link ? link.href : "";
                out.push({date, code, name, title, url: href});
            }
            return out;
        }""", max_results)


async def _search_one(browser, stock_code, stock_info, from_date="", to_date="",
                      max_results=200, max_retries=3):
    """搜索单只股票，带外层重试（覆盖 page.goto 连接错误等）。"""
    last_err = None
    for attempt in range(max_retries):
        context = await browser.new_context(
            user_agent=DEFAULT_UA, viewport={"width": 1920, "height": 1080}, locale="zh-CN")
        page = await context.new_page()
        await page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
            window.chrome = {runtime: {}};
        """)
        try:
            results = await _do_search(page, stock_code, stock_info["name"],
                                       from_date, to_date, max_results)
            return stock_code, results
        except Exception as e:
            last_err = e
            if attempt < max_retries - 1:
                await asyncio.sleep(RETRY_DELAYS_LONG[attempt])
        finally:
            await context.close()
    print(f"[_search_one] {stock_code}: 重试{max_retries}次后仍失败: {last_err}", file=sys.stderr)
    return stock_code, []


# ── 后处理过滤器 ──

def _is_annual_report(entry: dict) -> bool:
    """判断HKEX搜索结果是否为年报。

    数据来自浏览器的 page.evaluate（未经过bash），中文文本完整保留，
    因此可以直接做文本匹配。仅外部传入的CLI参数有编码损坏风险。
    """
    title = entry.get("title", "")
    url = entry.get("url", "")

    # 必须来自 listconews 路径
    if "listconews" not in url:
        return False

    # 排除 ESG / CSR 独立报告（子类别标签，非资料分类路径）
    if "ESG" in title or "esg" in title.lower():
        return False
    if "環境、社會及管治報告" in title:  # 独立的ESG报告
        return False

    # HKEX 年报的文件子类别标签
    # 注意：主分类"財務報表/環境、社會及管治資料"下包含两个子类：
    #   [年報] = 年报 ✓ , [環境、社會及管治資料/報告] = ESG报告 ✗
    if "[年報]" in title and "環境、社會及管治報告" not in title:
        return True
    if "Annual Report" in title:
        return True

    # 部分公司年报标题不含方括号，直接用"年報"
    if "年報" in title and "海外監管公告" not in title:
        return True

    return False


def find_annual_report(results: list[dict]) -> dict | None:
    """从搜索结果中定位年报条目。返回第一条匹配，无匹配返回None。"""
    # 先按日期倒序（最新的在前）
    for r in results:
        if _is_annual_report(r):
            return r
    return None


async def search_filings_async(stock_code: str, from_date: str = "", to_date: str = "",
                               max_results: int = 200) -> list[dict]:
    info = lookup_stock_id(stock_code)
    if not info:
        return []
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"])
        try:
            _, results = await _search_one(browser, stock_code, info,
                                           from_date, to_date, max_results)
            return results
        finally:
            await browser.close()


async def search_annual_report_async(stock_code: str, year: int = 2025) -> dict | None:
    """搜索港股公司某年年报，返回单条结果或None。

    自动搜索年报发布窗口期(当年3-4月)，使用后处理过滤器定位。
    """
    from_year = f"{year + 1}/03/01"
    to_year = f"{year + 1}/04/30"

    results = await search_filings_async(stock_code, from_year, to_year, max_results=200)
    if not results:
        return None
    return find_annual_report(results)


async def search_filings_multi(stock_codes: list[str], from_date: str = "",
                               to_date: str = "", max_results: int = 200
                               ) -> dict[str, list[dict]]:
    infos = lookup_stock_ids(stock_codes)
    result: dict[str, list[dict]] = {code: [] for code in stock_codes}
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"])
        try:
            codes_with_info = [(c, infos[c]) for c in stock_codes if c in infos]
            for idx, (code, info) in enumerate(codes_with_info):
                if idx > 0:
                    await asyncio.sleep(30)  # HKEX限流保护
                try:
                    _, data = await _search_one(browser, code, info,
                                                from_date, to_date, max_results)
                    result[code] = data
                except Exception as e:
                    print(f"[search_filings_multi] {code}: {e}", file=sys.stderr)
        finally:
            await browser.close()
    return result


# ── Sync wrappers ──

def search_filings(stock_code: str, from_date: str = "", to_date: str = "",
                   max_results: int = 200) -> list[dict]:
    return asyncio.run(search_filings_async(stock_code, from_date, to_date, max_results))


def search_annual_report(stock_code: str, year: int = 2025) -> dict | None:
    """搜索港股公司某年年报，同步封装。"""
    return asyncio.run(search_annual_report_async(stock_code, year))


def search_filings_multi_sync(stock_codes: list[str], from_date: str = "",
                              to_date: str = "", max_results: int = 200
                              ) -> dict[str, list[dict]]:
    return asyncio.run(search_filings_multi(stock_codes, from_date, to_date, max_results))


# ── CLI ──

def main():
    import argparse
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="港股公告搜索与下载")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("lookup", help="查股票内部ID")
    p.add_argument("--code", required=True)

    p = sub.add_parser("lookup-batch", help="批量查ID")
    p.add_argument("--codes", required=True, help="逗号分隔")

    p = sub.add_parser("search", help="搜索公告")
    p.add_argument("--code", required=True)
    p.add_argument("--from-date", default="")
    p.add_argument("--to-date", default="")
    p.add_argument("--max", type=int, default=200)

    p = sub.add_parser("search-annual", help="搜索年报(自动定位，不依赖中文关键词)")
    p.add_argument("--code", help="单股代码")
    p.add_argument("--codes", help="多股代码，逗号分隔")
    p.add_argument("--year", type=int, default=None, help="单年份(如2025)")
    p.add_argument("--years", type=int, nargs="+", help="多年份(如 2023 2024 2025)")
    p.add_argument("--download-dir", "-d", help="下载目录")
    p.add_argument("--quiet", "-q", action="store_true", help="静默模式，仅输出下载汇总")

    p = sub.add_parser("search-multi", help="并行搜索多只股票")
    p.add_argument("--codes", required=True, help="逗号分隔")
    p.add_argument("--from-date", default="")
    p.add_argument("--to-date", default="")
    p.add_argument("--max", type=int, default=200)

    p = sub.add_parser("download", help="下载PDF")
    p.add_argument("--url", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--timeout", type=int, default=180)

    args = parser.parse_args()

    if args.cmd == "lookup":
        info = lookup_stock_id(args.code)
        print(json.dumps(info, ensure_ascii=False, indent=2) if info else "FAIL")
    elif args.cmd == "lookup-batch":
        codes = [c.strip() for c in args.codes.split(",") if c.strip()]
        print(json.dumps(lookup_stock_ids(codes), ensure_ascii=False, indent=2))
    elif args.cmd == "search":
        print(json.dumps(search_filings(args.code, args.from_date, args.to_date,
                                         args.max), ensure_ascii=False, indent=2))
    elif args.cmd == "search-annual":
        codes = [c.strip() for c in args.codes.split(",") if c.strip()] if args.codes else []
        if args.code:
            codes.append(args.code.strip())
        if not codes:
            print("请指定 --code 或 --codes", file=sys.stderr)
            sys.exit(1)

        # 收集年份：--years 优先，其次 --year
        if args.years:
            years = args.years
        elif args.year:
            years = [args.year]
        else:
            print("请指定 --year 或 --years", file=sys.stderr)
            sys.exit(1)

        for year in years:
            from_date = f"{year + 1}/03/01"
            to_date = f"{year + 1}/04/30"

            all_filings = search_filings_multi_sync(codes, from_date, to_date, max_results=200)

            for code in codes:
                filings = all_filings.get(code, [])
                result = find_annual_report(filings)
                if result:
                    all_filings[code] = [result]
                    if not args.download_dir:
                        print(f"[{code}/{year}] {json.dumps(result, ensure_ascii=False, indent=2)}")
                else:
                    all_filings[code] = []
                    if not args.download_dir:
                        print(f"[{code}] 未找到{year}年报", file=sys.stderr)

            if args.download_dir:
                from pathlib import Path
                download_dir = Path(args.download_dir)
                ok_count = 0
                total = 0
                for i, code in enumerate(codes):
                    if i > 0:
                        time.sleep(INTER_CODE_DELAY)
                    results = all_filings.get(code, [])
                    if results:
                        r = results[0]
                        info = lookup_stock_id(code)
                        name = info["name"] if info else code
                        fname = f"{_normalize_code(code)}_{name}_{year}年报.pdf"
                        output = download_dir / fname
                        ok = download_pdf(r["url"], str(output), timeout=HKEX_PDF_DOWNLOAD_TIMEOUT)
                        total += 1
                        if ok:
                            ok_count += 1
                        if not args.quiet:
                            print(f"  {'OK' if ok else 'FAIL'} -> {output}")
                    else:
                        print(f"[{code}] 未找到{year}年报，跳过下载", file=sys.stderr)
                print(f"[{year}] 下载完成: {ok_count}/{total} 成功")

            # 年份间间隔（HKEX限流保护）
            if len(years) > 1:
                time.sleep(5)
    elif args.cmd == "search-multi":
        codes = [c.strip() for c in args.codes.split(",") if c.strip()]
        print(json.dumps(search_filings_multi_sync(codes, args.from_date, args.to_date,
                                                    args.max),
                         ensure_ascii=False, indent=2))
    elif args.cmd == "download":
        ok = download_pdf(args.url, args.output, timeout=args.timeout)
        print("OK" if ok else "FAIL")
        sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
