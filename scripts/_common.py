"""
财报下载工具共用模块：浏览器启动、PDF下载、验证、重试。
"""

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

# 重试间隔(秒) — 指数递增适配内地→HKEX的波动网络
RETRY_DELAYS_SHORT = [1, 3, 9]
RETRY_DELAYS_LONG = [5, 15, 30]


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


def download_pdf(url: str, output: str | Path, timeout: int = 120) -> bool:
    """requests下载PDF，带验证与重试。"""
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)

    for i, delay in enumerate(RETRY_DELAYS_SHORT):
        try:
            resp = requests.get(
                url,
                headers={"User-Agent": DEFAULT_UA, "Accept": "application/pdf,*/*"},
                timeout=timeout,
                verify=False,
            )
            if resp.status_code == 200 and len(resp.content) > 10000:
                path.write_bytes(resp.content)
                if verify_pdf(path):
                    return True
                path.unlink(missing_ok=True)
            if i < len(RETRY_DELAYS_SHORT) - 1:
                time.sleep(delay)
        except requests.RequestException:
            if i < len(RETRY_DELAYS_SHORT) - 1:
                time.sleep(delay)
    return False


def download_pdf_stream(url: str, output: str | Path, timeout: int = 180) -> bool:
    """流式下载大PDF，带验证与重试。"""
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)

    for i, delay in enumerate(RETRY_DELAYS_SHORT):
        try:
            resp = requests.get(
                url,
                headers={"User-Agent": DEFAULT_UA, "Accept": "application/pdf,*/*"},
                timeout=timeout,
                verify=False,
                stream=True,
            )
            if resp.status_code == 200:
                with open(path, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=1024 * 1024):
                        if chunk:
                            f.write(chunk)
                if verify_pdf(path):
                    return True
                path.unlink(missing_ok=True)
            if i < len(RETRY_DELAYS_SHORT) - 1:
                time.sleep(delay)
        except (requests.RequestException, OSError):
            if i < len(RETRY_DELAYS_SHORT) - 1:
                time.sleep(delay)
    return False


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
