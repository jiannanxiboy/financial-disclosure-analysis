"""
财报下载工具共用模块：浏览器启动、PDF下载、验证、重试。
"""

import hashlib
import json
import os
import shutil
import tempfile
import threading
import time
from pathlib import Path

import requests
import urllib3

urllib3.disable_warnings()

DEFAULT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

# ── 超时与重试 ──

# 巨潮资讯网 API (内地直连)
CNINFO_API_TIMEOUT = 30

# 港交所披露易 (内地跨境较慢)
HKEX_PAGE_LOAD_TIMEOUT = 90000   # 搜索页面首次加载
HKEX_TABLE_TIMEOUT = 60000       # 搜索结果表格渲染
HKEX_PDF_DOWNLOAD_TIMEOUT = 300  # PDF流式下载

# 重试间隔(秒) — 指数退避，适配内地→HKEX的波动网络
RETRY_DELAYS_SHORT = [3, 9, 27]    # 3次重试，总等待约39s
RETRY_DELAYS_LONG = [5, 15, 30]    # 长间隔，用于页面加载等重操作

# 多代码搜索间隔(秒) — 避免触发服务端限流
INTER_CODE_DELAY = 2.0

_DOWNLOAD_LOCKS: dict[str, threading.Lock] = {}
_DOWNLOAD_LOCKS_GUARD = threading.Lock()


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def atomic_write_json(path: str | Path, payload: dict) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp", dir=target.parent)
    os.close(fd)
    temp = Path(temp_name)
    try:
        temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        temp.replace(target)
    finally:
        temp.unlink(missing_ok=True)


def _download_cache_dir() -> Path | None:
    if os.environ.get("FINANCIAL_DISCLOSURE_NO_CACHE", "").lower() in {"1", "true", "yes"}:
        return None
    configured = os.environ.get("FINANCIAL_DISCLOSURE_CACHE_DIR")
    return Path(configured).expanduser() if configured else Path.home() / ".cache" / "financial-disclosure-analysis" / "pdfs"


def _url_lock(url: str) -> threading.Lock:
    key = hashlib.sha256(url.encode("utf-8")).hexdigest()
    with _DOWNLOAD_LOCKS_GUARD:
        return _DOWNLOAD_LOCKS.setdefault(key, threading.Lock())


def _restore_cached_pdf(url: str, output: Path) -> bool:
    cache = _download_cache_dir()
    if cache is None:
        return False
    url_key = hashlib.sha256(url.encode("utf-8")).hexdigest()
    metadata_path = cache / "urls" / f"{url_key}.json"
    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        object_path = cache / "objects" / f"{metadata['sha256']}.pdf"
        if not object_path.is_file() or object_path.stat().st_size != metadata["size"]:
            return False
        if sha256_file(object_path) != metadata["sha256"] or not verify_pdf(object_path):
            return False
        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(object_path, output)
        return verify_pdf(output)
    except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError):
        return False


def _store_cached_pdf(url: str, downloaded: Path) -> None:
    cache = _download_cache_dir()
    if cache is None:
        return
    digest = sha256_file(downloaded)
    objects = cache / "objects"
    urls = cache / "urls"
    objects.mkdir(parents=True, exist_ok=True)
    urls.mkdir(parents=True, exist_ok=True)
    object_path = objects / f"{digest}.pdf"
    if not object_path.exists():
        fd, temp_name = tempfile.mkstemp(prefix=f".{digest}.", suffix=".tmp", dir=objects)
        os.close(fd)
        temp = Path(temp_name)
        try:
            shutil.copy2(downloaded, temp)
            if sha256_file(temp) != digest:
                raise OSError(f"Cache copy hash mismatch: {downloaded}")
            try:
                temp.replace(object_path)
            except FileExistsError:
                pass
        finally:
            temp.unlink(missing_ok=True)
    url_key = hashlib.sha256(url.encode("utf-8")).hexdigest()
    atomic_write_json(
        urls / f"{url_key}.json",
        {"url": url, "sha256": digest, "size": downloaded.stat().st_size},
    )


def verify_pdf(path: str | Path, min_pages: int = 1) -> bool:
    """快速验证PDF有效且可读。"""
    p = Path(path)
    if not p.exists() or p.stat().st_size < 5000:
        return False
    try:
        import pdfplumber
        with pdfplumber.open(str(p)) as pdf:
            return len(pdf.pages) >= min_pages
    except Exception:
        return False


def _download_pdf_cached(url: str, output: str | Path, timeout: int, stream: bool) -> bool:
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.is_file() and verify_pdf(path):
        return True

    with _url_lock(url):
        if _restore_cached_pdf(url, path):
            return True
        for i, delay in enumerate(RETRY_DELAYS_SHORT):
            temp = path.with_name(f".{path.name}.{os.getpid()}.{threading.get_ident()}.part")
            response = None
            try:
                response = requests.get(
                    url,
                    headers={"User-Agent": DEFAULT_UA, "Accept": "application/pdf,*/*"},
                    timeout=timeout,
                    verify=False,
                    stream=stream,
                )
                if response.status_code == 200:
                    if stream:
                        with temp.open("wb") as handle:
                            for chunk in response.iter_content(chunk_size=1024 * 1024):
                                if chunk:
                                    handle.write(chunk)
                    else:
                        temp.write_bytes(response.content)
                    if temp.stat().st_size > 10000 and verify_pdf(temp):
                        _store_cached_pdf(url, temp)
                        temp.replace(path)
                        return True
                if i < len(RETRY_DELAYS_SHORT) - 1:
                    time.sleep(delay)
            except (requests.RequestException, OSError):
                if i < len(RETRY_DELAYS_SHORT) - 1:
                    time.sleep(delay)
            finally:
                temp.unlink(missing_ok=True)
                if response is not None:
                    response.close()
        return False


def download_pdf(url: str, output: str | Path, timeout: int = 120) -> bool:
    """下载并验证PDF；按URL复用内容寻址缓存，原子写入目标。"""
    return _download_pdf_cached(url, output, timeout=timeout, stream=False)


def download_pdf_stream(url: str, output: str | Path, timeout: int = 180) -> bool:
    """流式下载大PDF，并复用相同的内容寻址缓存。"""
    return _download_pdf_cached(url, output, timeout=timeout, stream=True)


def launch_browser(headless: bool = True):
    """启动带反检测的Chromium浏览器和context/page。"""
    from playwright.sync_api import sync_playwright

    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        headless=headless,
        args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
    )
    context = browser.new_context(
        user_agent=DEFAULT_UA,
        viewport={"width": 1920, "height": 1080},
        locale="zh-CN",
    )
    page = context.new_page()
    page.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
        window.chrome = {runtime: {}};
    """)
    return pw, browser, context, page


def poll_for_button(page, texts: list[str], max_wait: int = 15) -> bool:
    """轮询等待页面按钮出现。"""
    from playwright.sync_api import TimeoutError as PwTimeout

    for _ in range(max_wait):
        page.wait_for_timeout(1000)
        for t in texts:
            btn = page.query_selector(f'button:has-text("{t}")')
            if btn:
                return True
    return False


def retry(func, max_attempts: int = 3):
    """装饰器：指数退避重试，默认3次间隔1s/3s/9s。"""
    def wrapper(*args, **kwargs):
        last_err = None
        for i, delay in enumerate(RETRY_DELAYS_SHORT[:max_attempts]):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_err = e
                if i < max_attempts - 1:
                    time.sleep(delay)
        raise last_err
    return wrapper
