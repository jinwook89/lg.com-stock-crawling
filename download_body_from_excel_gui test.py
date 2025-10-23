# filename: download_body_from_excel_gui.py
import sys, re, html, json, time
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse
from collections import deque

import pandas as pd
from bs4 import BeautifulSoup

from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal, QTimer, QSettings, QRectF
from PyQt5.QtGui import QIcon, QColor, QFont, QPainter, QBrush, QPen, QPalette
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QMessageBox, QTextEdit, QCheckBox, QSpacerItem, QSizePolicy,
    QGroupBox, QTableWidget, QTableWidgetItem, QMenu
)

# ================== 앱 메타 ==================
APP_NAME      = "GMP3 lg.com Stock Crawler"
APP_ID        = "gmp3.crawler.gui"
APP_ICON_NAME = "stock crawling icon.ico"
APP_VERSION   = "1.5.4"  # QRectF/adjusted 수정 + QSS filter 제거

def resource_path(rel_path: str) -> str:
    base = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return str(Path(base) / rel_path)

# ====== Brand Colors (LG) ======
LG_RED          = "#A50034"  # LG Official Red
LG_RED_LIGHT    = "#C3003A"  # hover/gradient top
LG_RED_DARK     = "#7F0027"  # pressed/gradient bottom
LG_ACCENT_AMBER = "#a97100"  # WARN 텍스트
LG_ERROR_RED    = "#d11a2a"  # ERROR 텍스트

# ====== 기본 설정값(고정) ======
DEFAULT_SHEET_NAME = "url"
DEFAULT_OUTPUT_DIR_NAME = "html_body"
DEFAULT_SAVE_INNER_ONLY = True

# 안정성/성능(고정)
DEFAULT_HEADLESS         = False
DEFAULT_PAGE_TIMEOUT_SEC = 45
DEFAULT_WAIT_GRAPHQL_SEC = 15
DEFAULT_SCROLL_LOAD      = True
DEFAULT_SCROLL_LOOPS     = 6
DEFAULT_SCROLL_PAUSE_SEC = 0.7
DEFAULT_RETRY_ON_TIMEOUT = 1

# 대용량(고정)
DEFAULT_STREAM_TO_CSV    = True
DEFAULT_BATCH_SIZE       = 10
DEFAULT_JSON_MAX_BYTES   = 1_500_000  # 1.5MB

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/123.0.0.0 Safari/537.36")

# ====== 공통 유틸 ======
def sanitize_filename(name: str, max_len: int = 140) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", "_", name).strip("_")
    return name[:max_len] or "index"

def make_filename(url: str, idx: int) -> str:
    p = urlparse(url)
    host = p.netloc or "nohost"
    path = p.path or "/"
    stem_parts = [host] + [seg for seg in path.split("/") if seg]
    stem = "-".join(stem_parts) if stem_parts else host
    return f"{idx:04d}_{sanitize_filename(stem)}.html"

def wrap_html_with_base(body_html: str, base_href: str) -> str:
    return (
        "<!doctype html><html><head>"
        f'<base href="{base_href}"><meta charset="utf-8">'
        "</head><body>"
        f"{body_html}"
        "</body></html>"
    )

def extract_body_from_html_text(html_text: str, base_href: str) -> str:
    soup = BeautifulSoup(html_text or "", "lxml")
    if soup.body:
        body_inner = "".join(str(x) for x in soup.body.contents)
    else:
        body_inner = html_text or ""
    return wrap_html_with_base(body_inner, base_href)

def try_parse_json(text: str):
    try:
        return json.loads(text)
    except Exception:
        return None

def find_first_record_list(obj):
    stack, seen = [obj], set()
    while stack:
        cur = stack.pop()
        oid = id(cur)
        if oid in seen:
            continue
        seen.add(oid)
        if isinstance(cur, list) and cur and all(isinstance(x, dict) for x in cur):
            return cur
        if isinstance(cur, dict):
            stack.extend(cur.values())
        elif isinstance(cur, list):
            stack.extend(cur)
    return None

def normalize_graphql_jsons(json_objs, url_tag: str) -> pd.DataFrame:
    rows = []
    for jo in json_objs:
        if jo is None:
            continue
        recs = find_first_record_list(jo)
        if recs:
            for r in recs:
                if isinstance(r, dict):
                    rr = dict(r)
                    rr["_source_url"] = url_tag
                    rows.append(rr)
    return pd.DataFrame(rows) if rows else pd.DataFrame()

def _fmt_hms(seconds: float) -> str:
    seconds = max(0, int(seconds))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

# ====== 런타임 설정 ======
@dataclass
class RunConfig:
    sheet_name: str = DEFAULT_SHEET_NAME
    output_dir_name: str = DEFAULT_OUTPUT_DIR_NAME
    save_inner_only: bool = DEFAULT_SAVE_INNER_ONLY

    headless: bool = DEFAULT_HEADLESS
    page_timeout_sec: int = DEFAULT_PAGE_TIMEOUT_SEC
    wait_graphql_sec: int = DEFAULT_WAIT_GRAPHQL_SEC
    scroll_load: bool = DEFAULT_SCROLL_LOAD
    scroll_loops: int = DEFAULT_SCROLL_LOOPS
    scroll_pause_sec: float = DEFAULT_SCROLL_PAUSE_SEC
    retry_on_timeout: int = DEFAULT_RETRY_ON_TIMEOUT

    stream_to_csv: bool = DEFAULT_STREAM_TO_CSV
    batch_size: int = DEFAULT_BATCH_SIZE
    json_max_bytes: int = DEFAULT_JSON_MAX_BYTES

# ====== 캡슐형 진행바 ======
class CapsuleProgressBar(QWidget):
    def __init__(self, radius=8, bg="#FFFFFF", chunk=LG_RED, border="#CFD7E3", text_color=None, parent=None):
        super().__init__(parent)
        self._min, self._max, self._value = 0, 100, 0
        self._text_visible = True
        self._radius = radius
        self._bg = QColor(bg)
        self._chunk = QColor(chunk)
        self._border = QColor(border)
        self._text_color = QColor(text_color) if text_color else None
        self.setMinimumHeight(18)

    # API 유사성 제공
    def setRange(self, a, b): self._min, self._max = int(a), int(b); self.update()
    def setValue(self, v): self._value = int(v); self.update()
    def value(self): return self._value
    def minimum(self): return self._min
    def maximum(self): return self._max
    def setTextVisible(self, vis: bool): self._text_visible = bool(vis); self.update()

    def setColors(self, bg=None, chunk=None, border=None, text=None):
        if bg: self._bg = QColor(bg)
        if chunk: self._chunk = QColor(chunk)
        if border: self._border = QColor(border)
        if text: self._text_color = QColor(text)
        self.update()

    def setRadius(self, r): self._radius = int(r); self.update()

    def paintEvent(self, event):
        # 진행 비율
        minv, maxv = self._min, self._max
        pct = 0.0 if maxv == minv else max(0.0, min(1.0, (self._value - minv) / float(maxv - minv)))

        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)

        # ✅ QRectF로 변환 후 half-pixel 보정
        rectf = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        radius = min(self._radius, rectf.height() / 2.0)

        # 배경(캡슐)
        p.setPen(QPen(self._border, 1))
        p.setBrush(QBrush(self._bg))
        p.drawRoundedRect(rectf, radius, radius)

        # 진행 바(캡슐 유지)
        if pct > 0:
            w = max(2 * radius, rectf.width() * pct)
            chunk_rect = QRectF(rectf.left(), rectf.top(), w, rectf.height())
            p.setPen(Qt.NoPen)
            p.setBrush(QBrush(self._chunk))
            p.drawRoundedRect(chunk_rect, radius, radius)

        # 텍스트
        if self._text_visible:
            txt_color = self._text_color if self._text_color else self.palette().color(QPalette.Text)
            p.setPen(QPen(txt_color))
            p.drawText(rectf, Qt.AlignCenter, f"{int(pct*100)}%")

# ====== Selenium 드라이버 ======
def build_chrome_options(cfg: RunConfig, ua: str = UA):
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    chrome_opts = ChromeOptions()
    chrome_opts.page_load_strategy = "eager"
    if cfg.headless:
        chrome_opts.add_argument("--headless=new")
    chrome_opts.add_argument("--disable-gpu")
    chrome_opts.add_argument("--no-sandbox")
    chrome_opts.add_argument("--disable-dev-shm-usage")
    chrome_opts.add_argument("--disable-blink-features=AutomationControlled")
    chrome_opts.add_argument("--window-size=1920,1080")
    chrome_opts.add_argument(f"--user-agent={ua}")
    chrome_opts.add_experimental_option("prefs", {
        "profile.default_content_setting_values.images": 2,
        "profile.managed_default_content_settings.stylesheets": 1
    })
    return chrome_opts

def build_driver(cfg: RunConfig):
    from selenium import webdriver
    chrome_opts = build_chrome_options(cfg)
    driver = webdriver.Chrome(options=chrome_opts)
    driver.set_page_load_timeout(cfg.page_timeout_sec)
    driver.set_script_timeout(cfg.page_timeout_sec)
    return driver

def build_wire_driver(cfg: RunConfig):
    from seleniumwire import webdriver as wire_webdriver
    chrome_opts = build_chrome_options(cfg)
    driver = wire_webdriver.Chrome(options=chrome_opts)
    driver.set_page_load_timeout(cfg.page_timeout_sec)
    driver.set_script_timeout(cfg.page_timeout_sec)
    return driver

def detect_wire_available() -> bool:
    try:
        from seleniumwire.utils import decode as _  # noqa
        return True
    except Exception:
        return False

# ====== 쿠키 배너 처리 ======
def click_cookie_banners(driver):
    from selenium.webdriver.common.by import By
    candidates = [
        ("XPATH", "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]"),
        ("XPATH", "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'agree')]"),
        ("XPATH", "//button[contains(., '동의')]"),
        ("XPATH", "//button[contains(translate(., 'ÁÉÍÓÚÀÈÌÒÙÂÊÎÔÛÄËÏÖÜÇÑ', 'aeiouaeiouaeiouaeioucñ'), 'aceptar')]"),
        ("CSS", "button[aria-label*='Accept' i], button[aria-label*='동의' i]"),
        ("CSS", "[id*=consent i] button, [class*=cookie i] button, [data-testid*=consent i] button"),
    ]
    for typ, sel in candidates:
        try:
            els = driver.find_elements(By.XPATH, sel) if typ == "XPATH" else driver.find_elements(By.CSS_SELECTOR, sel)
            if els:
                els[0].click()
                time.sleep(0.2)
        except Exception:
            pass

# ====== 크롤 동작 ======
def crawl_urls(
    urls: pd.Series,
    output_dir: Path,
    cfg: RunConfig,
    progress_cb=None,
    progress_detail_cb=None,
    log_cb=None,
    stop_flag=lambda: False
):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, WebDriverException

    USE_WIRE = detect_wire_available()
    try:
        driver = build_wire_driver(cfg) if USE_WIRE else build_driver(cfg)
    except Exception:
        USE_WIRE = False
        driver = build_driver(cfg)

    driver.set_page_load_timeout(cfg.page_timeout_sec)
    driver.set_script_timeout(cfg.page_timeout_sec)
    output_dir.mkdir(parents=True, exist_ok=True)

    graphql_items_csv = output_dir / "graphql_items_part.csv"

    results, graphql_raw_records, graphql_norm_tables, network_json_rows = [], [], [], []
    total = len(urls)

    def restart_driver():
        nonlocal USE_WIRE, driver
        try:
            driver.quit()
        except Exception:
            pass
        try:
            driver = build_wire_driver(cfg) if USE_WIRE else build_driver(cfg)
        except Exception:
            USE_WIRE = False
            driver = build_driver(cfg)

    for i, url in enumerate(urls, start=1):
        if stop_flag():
            log_cb and log_cb("[INFO] 작업이 사용자에 의해 중단되었습니다.")
            break

        done_before = i - 1
        sub = 0.0
        def push(detail=None):
            nonlocal sub
            if detail is not None:
                sub = max(0.0, min(0.999, float(detail)))
            if progress_detail_cb:
                progress_detail_cb(done_before, total, sub)

        page_status = ""
        saved_path, size, err = "", 0, ""
        log_cb and log_cb(f"[INFO] [{i}/{total}] {url}")

        if progress_cb: progress_cb(i-1, total)
        push(0.0)

        try:
            # 단계 가중치
            w_load, w_scroll, w_save, w_wait, w_parse, w_wire = 0.25, 0.35, 0.10, 0.10, 0.10, 0.10

            # 1) 로딩
            tries = 0
            while True:
                try:
                    log_cb and log_cb("  • 페이지 로딩 중...")
                    driver.set_page_load_timeout(cfg.page_timeout_sec)
                    driver.set_script_timeout(cfg.page_timeout_sec)
                    driver.get(url)
                    WebDriverWait(driver, cfg.page_timeout_sec).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    try:
                        WebDriverWait(driver, 5).until(
                            lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
                        )
                    except Exception:
                        pass
                    break
                except TimeoutException as e:
                    tries += 1
                    log_cb and log_cb(f"[WARN] get() 타임아웃: {e}. 강제 중단/재시도")
                    try: driver.execute_script("window.stop();")
                    except Exception: pass
                    try:
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                        log_cb and log_cb("    └ body 감지됨. 계속 진행")
                        break
                    except Exception:
                        if tries <= cfg.retry_on_timeout:
                            log_cb and log_cb("    └ 새로고침 재시도")
                            try: driver.refresh()
                            except Exception: pass
                            continue
                        else:
                            raise
                except WebDriverException as e:
                    log_cb and log_cb(f"[WARN] 드라이버 오류: {e}. 드라이버 재시작")
                    restart_driver()
                    if tries <= cfg.retry_on_timeout:
                        tries += 1
                        continue
                    else:
                        raise

            push(w_load)

            # wire 가능 여부 갱신
            USE_WIRE = hasattr(driver, "requests")

            log_cb and log_cb("  • 쿠키 배너 확인...")
            click_cookie_banners(driver)
            push(w_load + 0.02)

            # 2) 스크롤
            if cfg.scroll_load:
                log_cb and log_cb("  • 스크롤 유도 중...")
                prev_h = 0
                for k in range(cfg.scroll_loops):
                    h = driver.execute_script("return document.body ? document.body.scrollHeight : 0")
                    if h == prev_h:
                        push(w_load + w_scroll); break
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(cfg.scroll_pause_sec)
                    prev_h = h
                    frac = (k + 1) / cfg.scroll_loops
                    push(w_load + w_scroll * min(1.0, frac))
            else:
                push(w_load + w_scroll)

            # 3) DOM 저장(항상 래핑)
            log_cb and log_cb("  • DOM 캡처 저장...")
            if cfg.save_inner_only:
                body_inner = driver.execute_script("return document.body ? document.body.innerHTML : ''") or ""
                minimal = wrap_html_with_base(body_inner, url)
            else:
                full_html = driver.page_source or ""
                minimal = extract_body_from_html_text(full_html, base_href=url)

            fpath = output_dir / make_filename(url, i)
            fpath.write_text(minimal, encoding="utf-8")
            saved_path = str(fpath.resolve())
            size = len(minimal.encode("utf-8"))
            push(w_load + w_scroll + w_save)

            # 4) data-graphql
            log_cb and log_cb("  • data-graphql 추출 대기/파싱...")
            try:
                WebDriverWait(driver, cfg.wait_graphql_sec).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-graphql]"))
                )
            except Exception:
                pass
            push(w_load + w_scroll + w_save + w_wait * 0.6)

            elems = driver.find_elements(By.CSS_SELECTOR, "[data-graphql]")
            vals = []
            for e in elems:
                try:
                    v = e.get_attribute("data-graphql")
                    if v: vals.append(html.unescape(v))
                except Exception:
                    pass
            log_cb and log_cb(f"    └ 발견: {len(vals)} 블록")
            push(w_load + w_scroll + w_save + w_wait)

            if len(vals) == 0:
                debug_path = output_dir / f"debug_{i:04d}.html"
                try:
                    debug_path.write_text(driver.page_source or "", encoding="utf-8")
                    log_cb and log_cb(f"    └ DEBUG 저장: {debug_path.name}")
                except Exception:
                    pass

            for j, raw in enumerate(vals, start=1):
                graphql_raw_records.append({"url": url, "block_index": j, "json_text": raw})

            # 5) 평탄화
            json_objs = [try_parse_json(t) for t in vals]
            norm_df = normalize_graphql_jsons(json_objs, url)
            if not norm_df.empty:
                graphql_norm_tables.append(norm_df)
                if cfg.stream_to_csv and len(graphql_norm_tables) >= cfg.batch_size:
                    pd.concat(graphql_norm_tables, ignore_index=True).to_csv(
                        graphql_items_csv, mode="a", header=not graphql_items_csv.exists(), index=False
                    )
                    graphql_norm_tables.clear()
                log_cb and log_cb(f"  • 평탄화 행 추가: {len(norm_df)}")
            push(w_load + w_scroll + w_save + w_wait + w_parse)

            # 6) 네트워크 JSON/상태코드
            if USE_WIRE:
                log_cb and log_cb("  • 네트워크 JSON 확인...")
                cnt, STEP, step_acc = 0, 5, 0
                base_host = urlparse(url).netloc
                page_status_found = False
                for req in getattr(driver, "requests", []):
                    try:
                        if not req.response:
                            continue
                        if not page_status_found:
                            rhost = urlparse(req.url or "").netloc
                            if rhost == base_host:
                                page_status = req.response.status_code
                                page_status_found = True

                        ctype = (req.response.headers or {}).get("Content-Type", "")
                        is_jsonish = "application/json" in ctype or "/graphql" in (req.url or "").lower()
                        if not is_jsonish:
                            continue
                        body = req.response.body
                        if not body:
                            continue
                        enc_len = len(body) if isinstance(body, (bytes, bytearray)) else len(str(body).encode("utf-8"))
                        if enc_len > cfg.json_max_bytes:
                            continue
                        try:
                            from seleniumwire.utils import decode
                            body = decode(body, req.response.headers.get('Content-Encoding', 'identity'))
                        except Exception:
                            pass
                        text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
                        if len(text.encode("utf-8")) > cfg.json_max_bytes:
                            continue

                        jo = try_parse_json(text)
                        if isinstance(jo, (dict, list)):
                            cnt += 1
                            network_json_rows.append({
                                "url": url, "index": cnt, "endpoint": req.url,
                                "status": req.response.status_code,
                                "json_text": json.dumps(jo, ensure_ascii=False)
                            })
                            ndf = normalize_graphql_jsons([jo], url)
                            if not ndf.empty:
                                graphql_norm_tables.append(ndf)
                                if cfg.stream_to_csv and len(graphql_norm_tables) >= cfg.batch_size:
                                    pd.concat(graphql_norm_tables, ignore_index=True).to_csv(
                                        graphql_items_csv, mode="a", header=not graphql_items_csv.exists(), index=False
                                    )
                                    graphql_norm_tables.clear()
                            step_acc += 1
                            if step_acc >= STEP:
                                step_acc = 0
                                cur_detail = w_load + w_scroll + w_save + w_wait + w_parse
                                cur_detail += w_wire * min(0.95, cnt / 50.0)
                                push(cur_detail)
                    except Exception:
                        continue
                log_cb and log_cb(f"    └ 네트워크 JSON 응답: {cnt}")
                push(w_load + w_scroll + w_save + w_wait + w_parse + w_wire)
            else:
                push(w_load + w_scroll + w_save + w_wait + w_parse + w_wire)

        except Exception as e:
            err = f"오류: {e}"
            log_cb and log_cb("[ERROR] " + err)

        results.append({
            "url": url,
            "status_code": page_status,
            "bytes": size if size else "",
            "saved_path": saved_path,
            "error": err
        })

        if progress_cb: progress_cb(i, total)

    try:
        driver.quit()
    except Exception:
        pass

    res_df = pd.DataFrame(results)
    graphql_raw_df = pd.DataFrame(graphql_raw_records)

    if cfg.stream_to_csv and graphql_norm_tables:
        pd.concat(graphql_norm_tables, ignore_index=True).to_csv(
            graphql_items_csv, mode="a", header=not graphql_items_csv.exists(), index=False
        )
        graphql_norm_tables.clear()

    graphql_items_df = pd.DataFrame()
    if graphql_items_csv.exists():
        try:
            graphql_items_df = pd.read_csv(graphql_items_csv).head(5000)
        except Exception:
            graphql_items_df = pd.DataFrame()

    network_json_df = pd.DataFrame(network_json_rows)

    if graphql_raw_df.empty:   graphql_raw_df = pd.DataFrame(columns=["url", "block_index", "json_text"])
    if graphql_items_df.empty: graphql_items_df = pd.DataFrame(columns=["_source_url"])
    if network_json_df.empty:  network_json_df = pd.DataFrame(columns=["url", "index", "endpoint", "status", "json_text"])
    return res_df, graphql_raw_df, graphql_items_df, network_json_df

# ====== 워커 ======
class CrawlWorker(QObject):
    progress = pyqtSignal(int, int)
    progress_detail = pyqtSignal(int, int, float)
    log = pyqtSignal(str)
    finished = pyqtSignal(object)
    errored = pyqtSignal(str)

    def __init__(self, excel_path: Path, cfg: RunConfig):
        super().__init__()
        self.excel_path = excel_path
        self.cfg = cfg
        self._stop = False

    def stop(self): self._stop = True

    def run(self):
        try:
            self.log.emit("[INFO] 엑셀 로딩 중...")
            df = pd.read_excel(self.excel_path, sheet_name=self.cfg.sheet_name, dtype=str)
            urls = df.iloc[:, 0].astype(str).str.strip()
            urls = urls[urls.ne("") & urls.str.startswith(("http://", "https://"), na=False)].reset_index(drop=True)
            if urls.empty: raise RuntimeError("유효한 URL이 없습니다.")

            base_dir = self.excel_path.parent
            output_dir = base_dir / self.cfg.output_dir_name

            self.log.emit(f"[INFO] 총 {len(urls)}건 처리 시작")
            def progress_cb(done, total): self.progress.emit(done, total)
            def progress_detail_cb(done, total, sub): self.progress_detail.emit(done, total, sub)
            def log_cb(msg): self.log.emit(msg)
            def stop_flag(): return self._stop

            res_df, graphql_raw_df, graphql_items_df, network_json_df = crawl_urls(
                urls, output_dir, self.cfg, progress_cb, progress_detail_cb, log_cb, stop_flag
            )
            self.progress.emit(len(urls), len(urls))
            self.finished.emit((res_df, graphql_raw_df, graphql_items_df, network_json_df, base_dir, df))
        except Exception as e:
            self.errored.emit(str(e))

# ====== GUI ======
class MainWindow(QWidget):
    ORG = "HSAD"
    APP = "GMP3Crawler"

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.resize(1120, 760)
        self.setAcceptDrops(True)

        self.settings = QSettings(self.ORG, self.APP)

        # 기본 테마
        self.dark = False

        self.excel_path: Path | None = None
        self.thread = None
        self.worker = None
        self.last_output_dir: Path | None = None
        self.last_result_file: Path | None = None

        # 진행 상태
        self._done = 0
        self._total = 0
        self._sub = 0.0
        self.t0 = None
        self.hist = deque()
        self.rolling_window = 20.0

        # 진행률 스무딩
        self._target_frac = 0.0
        self._display_frac = 0.0
        self.anim = QTimer(self)
        self.anim.setInterval(33)
        self.anim.timeout.connect(self.animate_progress)

        v = QVBoxLayout(self)
        v.setContentsMargins(18, 18, 18, 18)
        v.setSpacing(12)

        # 헤더
        header = QHBoxLayout()
        header.setSpacing(8)
        self.title = QLabel(f"{APP_NAME}  v{APP_VERSION}")
        self.title.setStyleSheet("font-size:18px; font-weight:700;")
        self.theme_btn = QPushButton("다크 테마"); self.theme_btn.setObjectName("secondary")
        self.theme_btn.clicked.connect(self.toggle_theme)
        header.addWidget(self.title)
        header.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        header.addWidget(self.theme_btn)
        v.addLayout(header)

        # 상단 컨트롤 카드
        ctrl = QGroupBox("작업")
        ctrl_l = QVBoxLayout(); ctrl_l.setSpacing(8); ctrl_l.setContentsMargins(12, 12, 12, 12)
        self.lbl_path = QLabel("엑셀 파일을 선택하거나 아래 영역에 드래그하세요.")
        self.lbl_path.setObjectName("hint")
        self.lbl_path.setStyleSheet("""
            QLabel#hint {
                border: 2px dashed rgba(0,0,0,0.15);
                border-radius: 12px;
                padding: 10px 12px;
            }
        """)
        hb = QHBoxLayout(); hb.setSpacing(8)
        self.btn_pick = QPushButton("엑셀 선택"); self.btn_pick.setObjectName("secondary")
        self.btn_start = QPushButton("실행"); self.btn_start.setObjectName("primary"); self.btn_start.setEnabled(False)
        self.btn_stop = QPushButton("중단"); self.btn_stop.setEnabled(False)
        self.btn_open_out = QPushButton("출력 폴더 열기"); self.btn_open_out.setEnabled(False)
        self.btn_open_result = QPushButton("결과 엑셀 열기"); self.btn_open_result.setEnabled(False)
        self.chk_auto_open = QCheckBox("완료 후 자동으로 폴더 열기"); self.chk_auto_open.setChecked(False)
        hb.addWidget(self.btn_pick); hb.addWidget(self.btn_start); hb.addWidget(self.btn_stop)
        hb.addWidget(self.btn_open_out); hb.addWidget(self.btn_open_result); hb.addWidget(self.chk_auto_open)
        ctrl_l.addWidget(self.lbl_path)
        ctrl_l.addLayout(hb)
        ctrl.setLayout(ctrl_l)
        v.addWidget(ctrl)

        # 진행 상황 카드
        prog_card = QGroupBox("진행 상황")
        prog_l = QVBoxLayout(); prog_l.setSpacing(8); prog_l.setContentsMargins(12, 12, 12, 12)
        self.progress = CapsuleProgressBar(radius=8)
        self.progress.setRange(0, 100); self.progress.setValue(0); self.progress.setTextVisible(True)
        self.stats = QLabel("0 / 0 (0%) | 경과: 00:00:00 | 남음: ~--:--:-- | 속도: -- urls/min")
        self.stats.setObjectName("hint")
        prog_l.addWidget(self.progress)
        prog_l.addWidget(self.stats)
        prog_card.setLayout(prog_l)
        v.addWidget(prog_card)

        # 로그 카드
        log_card = QGroupBox("로그")
        log_l = QHBoxLayout(); log_l.setSpacing(8); log_l.setContentsMargins(12, 12, 12, 12)
        self.log = QTextEdit(); self.log.setReadOnly(True)
        mono = QFont("Consolas"); mono.setPointSize(10)
        self.log.setFont(mono); self.log.setLineWrapMode(self.log.NoWrap)
        self.btn_save_log = QPushButton("로그 저장"); self.btn_save_log.setEnabled(True)
        log_l.addWidget(self.log, 1); log_l.addWidget(self.btn_save_log)
        log_card.setLayout(log_l)
        v.addWidget(log_card, 2)

        # 결과 미리보기 카드
        self.preview = QGroupBox("결과 미리보기 (download_log 상위 20행)")
        pvbox = QVBoxLayout(); pvbox.setSpacing(8); pvbox.setContentsMargins(12, 12, 12, 12)
        self.tbl_preview = QTableWidget(0, 0)
        self.tbl_preview.setAlternatingRowColors(True)
        self.tbl_preview.verticalHeader().setVisible(False)
        self.tbl_preview.setSortingEnabled(True)
        self.tbl_preview.setSelectionBehavior(self.tbl_preview.SelectRows)
        self.tbl_preview.setSelectionMode(self.tbl_preview.SingleSelection)
        pvbox.addWidget(self.tbl_preview)
        self.preview.setLayout(pvbox)
        v.addWidget(self.preview, 3)

        # 타이머
        self.timer = QTimer(self); self.timer.setInterval(1000); self.timer.timeout.connect(self.update_stats_tick)

        # 연결
        self.btn_pick.clicked.connect(self.pick_excel)
        self.btn_start.clicked.connect(self.start_work)
        self.btn_stop.clicked.connect(self.stop_work)
        self.btn_open_out.clicked.connect(self.open_output_dir)
        self.btn_open_result.clicked.connect(self.open_result_file)
        self.btn_save_log.clicked.connect(self.save_log_to_file)

        # 버튼 아이콘
        self._load_icons()

        # 테마(QSS) 적용 + 진행바 색상 연동
        self.apply_theme(self.dark)
        self._apply_progress_colors()

        # 마지막 경로 복원
        last_path = self.settings.value("last_excel_path", "", type=str)
        if last_path and Path(last_path).exists():
            self.set_excel_path(Path(last_path))

    # --------- 테마(QSS) ----------
    def apply_theme(self, dark: bool):
        common = """
            * { font-family: 'Malgun Gothic','Segoe UI',sans-serif; }
            QWidget { font-size: 12px; }
            QLabel#hint { color: palette(mid); }

            QPushButton {
                height: 36px; padding: 0 16px;
                border-radius: 10px; border: 1px solid transparent;
                font-weight: 600;
            }
            /* primary 버튼 그라데이션은 아래 primary_css에서 주입 */

            QPushButton#secondary { font-weight: 600; }
            QPushButton:disabled { opacity: .6; }

            QTextEdit { border-radius: 12px; padding: 10px; }

            QGroupBox {
                font-weight: 600;
                border: 1px solid rgba(0,0,0,0.08);
                border-radius: 12px; margin-top: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin; left: 12px; top: -2px;
                padding: 0 6px;
            }

            QTableWidget {
                border: 1px solid transparent;
                border-radius: 12px;
                gridline-color: rgba(0,0,0,0.08);
                selection-background-color: """ + LG_RED + """;
                selection-color: white;
                alternate-background-color: rgba(0,0,0,0.03);
            }
            QHeaderView::section {
                padding: 6px 10px; border: 0; border-bottom: 1px solid rgba(0,0,0,0.08);
                font-weight: 600;
            }
            QTableWidget::item:selected { outline: none; }
        """

        primary_css = f"""
            QPushButton#primary {{
                color: white;
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                            stop:0 {LG_RED_LIGHT}, stop:1 {LG_RED});
                border: 0;
            }}
            QPushButton#primary:hover {{
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                            stop:0 {LG_RED_LIGHT}, stop:1 {LG_RED});
            }}
            QPushButton#primary:pressed {{
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                            stop:0 {LG_RED_DARK}, stop:1 {LG_RED});
            }}
        """

        if dark:
            self.setStyleSheet(common + primary_css + """
                QWidget { color: #E6EAF2; background: #0C1220; }
                QLabel#hint { color: #9AA4B2; }

                QPushButton { color: #E6EAF2; background: #0F172A; border: 1px solid #1E293B; }
                QPushButton:hover { background: #14223B; }
                QPushButton#secondary { background: #0F172A; border: 1px solid #243449; color: #D2DBE7; }

                QTextEdit { background: #0F172A; border: 1px solid #1E293B; }

                QGroupBox { border: 1px solid rgba(255,255,255,0.08); }
                QGroupBox::title { color: #C7D1E1; background: #0C1220; }

                QTableWidget { background: #0F172A; color: #E6EAF2; alternate-background-color: rgba(255,255,255,0.03); }
                QHeaderView::section { background: #0F172A; color: #C7D1E1; border-bottom: 1px solid #1E293B; }
            """)
        else:
            self.setStyleSheet(common + primary_css + """
                QWidget { color: #1F2A44; background: #F6F8FB; }
                QLabel#hint { color: #6B7280; }

                QPushButton { color: #1F2A44; background: #FFFFFF; border: 1px solid #D6DDEB; }
                QPushButton:hover { background: #F1F5FF; }
                QPushButton#secondary { background: #E9EEF7; color: #1F2A44; border: 1px solid #CFD7E3; }

                QTextEdit { background: #FFFFFF; border: 1px solid #CFD7E3; }

                QGroupBox { border: 1px solid rgba(0,0,0,0.06); }
                QGroupBox::title { color: #334155; background: #F6F8FB; }

                QTableWidget { background: #FFFFFF; color: #1F2A44; alternate-background-color: #F8FAFE; }
                QHeaderView::section { background: #FFFFFF; color: #334155; border-bottom: 1px solid #CFD7E3; }
            """)

    def _apply_progress_colors(self):
        if isinstance(self.progress, CapsuleProgressBar):
            if self.dark:
                self.progress.setColors(bg="#0F172A", chunk=LG_RED, border="#1E293B", text="#E6EAF2")
            else:
                self.progress.setColors(bg="#FFFFFF", chunk=LG_RED, border="#CFD7E3", text="#1F2A44")

    def toggle_theme(self):
        self.dark = not self.dark
        self.apply_theme(self.dark)
        self._apply_progress_colors()
        self.theme_btn.setText("라이트 테마" if self.dark else "다크 테마")
        self._set_theme_icon()

    # ---- 버튼 아이콘 로드 ----
    def _load_icons(self):
        def icon_try(path):
            p = resource_path(path)
            return QIcon(p) if Path(p).exists() else QIcon()
        self.btn_pick.setIcon(icon_try("icons/folder-open.png"))
        self.btn_start.setIcon(icon_try("icons/play.png"))
        self.btn_stop.setIcon(icon_try("icons/stop.png"))
        self.btn_open_out.setIcon(icon_try("icons/folder.png"))
        self.btn_open_result.setIcon(icon_try("icons/file-excel.png"))
        self.btn_save_log.setIcon(icon_try("icons/save.png"))
        self._set_theme_icon()

    def _set_theme_icon(self):
        icon = "icons/sun.png" if self.dark else "icons/moon.png"
        p = resource_path(icon)
        if Path(p).exists():
            self.theme_btn.setIcon(QIcon(p))

    # ---- Drag & Drop ----
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            urls = e.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith((".xlsx", ".xls")):
                self.lbl_path.setStyleSheet(f"""
                    QLabel#hint {{
                        border: 2px dashed {LG_RED};
                        border-radius: 12px;
                        padding: 10px 12px;
                        background: rgba(165,0,52,0.06);
                    }}
                """)
                e.acceptProposedAction(); return
        e.ignore()

    def dragLeaveEvent(self, e):
        self.lbl_path.setStyleSheet("""
            QLabel#hint {
                border: 2px dashed rgba(0,0,0,0.15);
                border-radius: 12px;
                padding: 10px 12px;
            }
        """)

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        self.dragLeaveEvent(e)
        if urls:
            p = Path(urls[0].toLocalFile())
            if p.exists() and p.suffix.lower() in (".xlsx", ".xls"):
                self.set_excel_path(p)

    # ---- 경로 ----
    def set_excel_path(self, p: Path):
        self.excel_path = p
        self.lbl_path.setText(f"선택됨: {self.excel_path}")
        self.btn_start.setEnabled(True)
        self.last_output_dir = self.excel_path.parent / DEFAULT_OUTPUT_DIR_NAME
        self.settings.setValue("last_excel_path", str(p))

    def pick_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if not path: return
        self.set_excel_path(Path(path))

    # ---- 실행/중단 ----
    def build_config(self) -> RunConfig:
        return RunConfig()

    def start_work(self):
        if not self.excel_path:
            QMessageBox.warning(self, "알림", "엑셀 파일을 먼저 선택하세요."); return

        self.log.clear()
        self.progress.setRange(0, 100); self.progress.setValue(0)
        self.stats.setText("0 / 0 (0%) | 경과: 00:00:00 | 남음: ~--:--:-- | 속도: -- urls/min")
        self.btn_start.setEnabled(False); self.btn_stop.setEnabled(True)
        self.btn_pick.setEnabled(False);  self.btn_open_out.setEnabled(False); self.btn_open_result.setEnabled(False)

        # 상태 초기화
        self._done = self._total = 0; self._sub = 0.0
        self._target_frac = self._display_frac = 0.0
        self.t0 = time.time(); self.hist.clear(); self.hist.append((self.t0, 0.0))

        cfg = self.build_config()

        self.thread = QThread()
        self.worker = CrawlWorker(self.excel_path, cfg)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.progress_detail.connect(self.on_progress_detail)
        self.worker.log.connect(self.append_log_rich)
        self.worker.finished.connect(self.on_finished)
        self.worker.errored.connect(self.on_error)
        self.thread.start()
        self.timer.start()
        self.anim.start()

        self.append_log_rich("[INFO] 실행을 시작했습니다.")

    def stop_work(self):
        if self.worker:
            self.worker.stop()
            self.append_log_rich("[WARN] 중단 요청됨... (현재 URL 처리 마무리 후 중단)")

    # ---- 진행률/통계 ----
    def on_progress(self, done, total):
        self._done, self._total = done, total
        self._sub = 0.0
        self._record_progress_sample()
        self.update_progress_target()
        self.update_stats_label()

    def on_progress_detail(self, done, total, sub_pct):
        self._done, self._total = done, total
        self._sub = max(self._sub, max(0.0, min(0.999, sub_pct)))
        self._record_progress_sample()
        self.update_progress_target()
        self.update_stats_label()

    def update_stats_label(self):
        pct = int(((self._done + self._sub) / self._total) * 100) if self._total > 0 else 0
        now = time.time(); elapsed = now - (self.t0 or now)
        self.stats.setText(f"{self._done}/{self._total} ({pct}%) | 경과: {_fmt_hms(elapsed)} | 남음: ~--:--:-- | 속도: -- urls/min")

    def update_stats_tick(self):
        if self.t0 is None: return
        now = time.time(); elapsed = now - self.t0
        if len(self.hist) >= 2:
            t0, u0 = self.hist[0]; t1, u1 = self.hist[-1]
            dt = t1 - t0; du = u1 - u0
            speed_units_per_sec = (du / dt) if dt > 0 else 0.0
        else:
            speed_units_per_sec = 0.0
        remaining_units = max(0.0, self._total - (self._done + self._sub))
        pct = int(((self._done + self._sub) / self._total) * 100) if self._total > 0 else 0
        if speed_units_per_sec > 0:
            eta_sec = remaining_units / speed_units_per_sec
            speed_per_min = 60.0 * speed_units_per_sec
            self.stats.setText(f"{self._done}/{self._total} ({pct}%) | 경과: {_fmt_hms(elapsed)} | 남음: ~{_fmt_hms(eta_sec)} | 속도: {speed_per_min:.1f} urls/min")
        else:
            self.stats.setText(f"{self._done}/{self._total} ({pct}%) | 경과: {_fmt_hms(elapsed)} | 남음: ~--:--:-- | 속도: -- urls/min")

    def _record_progress_sample(self):
        if self.t0 is None: return
        now = time.time()
        units = (self._done + self._sub)
        self.hist.append((now, units))
        cutoff = now - self.rolling_window
        while self.hist and self.hist[0][0] < cutoff:
            self.hist.popleft()

    def update_progress_target(self):
        if self._total <= 0: new_target = 0.0
        else: new_target = max(0.0, min(1.0, (self._done + self._sub) / self._total))
        if new_target >= self._target_frac:
            self._target_frac = new_target

    def animate_progress(self):
        diff = self._target_frac - self._display_frac
        if abs(diff) < 0.002:
            self._display_frac = self._target_frac
        else:
            step = max(0.002, abs(diff) * 0.15)
            self._display_frac += step if diff > 0 else -step
        self.progress.setValue(int(self._display_frac * 100))

    # ---- 로그 ----
    def append_log_rich(self, msg: str):
        color = QColor("#1F2A44") if not self.dark else QColor("#E6EAF2")
        if msg.startswith("[ERROR]"):
            color = QColor(LG_ERROR_RED)
        elif msg.startswith("[WARN]"):
            color = QColor(LG_ACCENT_AMBER)
        elif msg.startswith("[INFO]"):
            color = QColor(LG_RED)  # LG 레드로 변경
        self.log.setTextColor(color)
        self.log.append(msg)
        self.log.moveCursor(self.log.textCursor().End)

    def save_log_to_file(self):
        if not self.log.toPlainText().strip():
            QMessageBox.information(self, "안내", "저장할 로그가 없습니다."); return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_dir = str(self.excel_path.parent if self.excel_path else Path.cwd())
        path, _ = QFileDialog.getSaveFileName(self, "로그 저장", f"{default_dir}/crawler_log_{ts}.txt", "Text Files (*.txt)")
        if not path: return
        try:
            Path(path).write_text(self.log.toPlainText(), encoding="utf-8")
            QMessageBox.information(self, "완료", f"로그 저장됨: {path}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"로그 저장 실패: {e}")

    # ---- 표 컨텍스트 메뉴(셀 복사) ----
    def contextMenuEvent(self, e):
        if self.tbl_preview.underMouse():
            menu = QMenu(self)
            copy = menu.addAction("셀 복사")
            act = menu.exec_(e.globalPos())
            if act == copy:
                items = self.tbl_preview.selectedItems()
                if items:
                    QApplication.clipboard().setText(items[0].text())

    # ---- 완료/에러/정리 ----
    def on_error(self, message: str):
        self.timer.stop(); self.anim.stop()
        self.stats.setText(f"{self._done}/{self._total} | 경과: {_fmt_hms(time.time() - (self.t0 or time.time()))} | 남음: ~--:--:-- | 속도: 중단")
        self.cleanup_thread()
        self.btn_start.setEnabled(True); self.btn_stop.setEnabled(False); self.btn_pick.setEnabled(True)
        QMessageBox.critical(self, "에러", message)

    def on_finished(self, payload):
        self.timer.stop(); self.anim.stop()
        self._done, self._total, self._sub = 1, 1, 0.0
        self._target_frac = self._display_frac = 1.0
        self.progress.setValue(100)
        self.stats.setText(f"{self._done}/{self._total} (100%) | 경과: {_fmt_hms(time.time() - (self.t0 or time.time()))} | 남음: ~00:00:00 | 속도: 완료")

        res_df, graphql_raw_df, graphql_items_df, network_json_df, base_dir, df = payload
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            result_xlsx_path = base_dir / f"{self.excel_path.stem}__results_{ts}.xlsx"
            errors_csv_path = base_dir / f"{self.excel_path.stem}__errors_{ts}.csv"

            try:
                err_only = res_df[res_df["error"].astype(str).str.len() > 0]
                if not err_only.empty:
                    err_only.to_csv(errors_csv_path, index=False, encoding="utf-8-sig")
                    self.append_log_rich(f"[INFO] 에러 URL CSV 저장: {errors_csv_path}")
            except Exception as e:
                self.append_log_rich(f"[WARN] 에러 CSV 저장 실패: {e}")

            try:
                with pd.ExcelWriter(result_xlsx_path, engine="openpyxl") as w:
                    df.to_excel(w, index=False, sheet_name=DEFAULT_SHEET_NAME)
                    res_df.to_excel(w, index=False, sheet_name="download_log")
                    graphql_raw_df.to_excel(w, index=False, sheet_name="graphql_raw")
                    graphql_items_df.to_excel(w, index=False, sheet_name="graphql_items_sample")
                    network_json_df.to_excel(w, index=False, sheet_name="network_json")
            except Exception as e:
                fallback = base_dir / f"{self.excel_path.stem}__results_{ts}"
                df.to_csv(fallback.with_suffix(".input_urls.csv"), index=False)
                res_df.to_csv(fallback.with_suffix(".download_log.csv"), index=False)
                graphql_raw_df.to_csv(fallback.with_suffix(".graphql_raw.csv"), index=False)
                graphql_items_df.to_csv(fallback.with_suffix(".graphql_items_sample.csv"), index=False)
                network_json_df.to_csv(fallback.with_suffix(".network_json.csv"), index=False)
                raise RuntimeError(f"Excel 저장 실패로 CSV로 대체 저장했습니다. 원인: {e}")

            self.append_log_rich(f"[INFO] 완료: {len(res_df)}건")
            self.append_log_rich(f"[INFO] 저장 폴더(HTML): {(base_dir / DEFAULT_OUTPUT_DIR_NAME).resolve()}")
            self.append_log_rich(f"[INFO] 결과 엑셀: {result_xlsx_path.resolve()}")
            self.append_log_rich(f"[INFO] DOM data-graphql 블록: {len(graphql_raw_df)}")
            self.append_log_rich(f"[INFO] 평탄화 샘플 행 수: {len(graphql_items_df)}")
            self.append_log_rich(f"[INFO] 네트워크 JSON 응답: {len(network_json_df)}")

            gcsv = base_dir / DEFAULT_OUTPUT_DIR_NAME / "graphql_items_part.csv"
            if gcsv.exists():
                self.append_log_rich(f"[INFO] 전체 평탄화 누적 CSV: {gcsv.resolve()}")

            self.last_output_dir = base_dir / DEFAULT_OUTPUT_DIR_NAME
            self.last_result_file = result_xlsx_path
            self.btn_open_out.setEnabled(True)
            self.btn_open_result.setEnabled(True)

            self.populate_preview(res_df.head(20))

            if self.chk_auto_open.isChecked():
                self.open_path_in_explorer(base_dir)
            QMessageBox.information(self, "완료", "결과 저장을 마쳤습니다.")
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", str(e))
        finally:
            self.cleanup_thread()
            self.btn_start.setEnabled(True); self.btn_stop.setEnabled(False); self.btn_pick.setEnabled(True)

    def populate_preview(self, df: pd.DataFrame):
        self.tbl_preview.clear()
        if df is None or df.empty:
            self.tbl_preview.setRowCount(0); self.tbl_preview.setColumnCount(0)
            return
        self.tbl_preview.setColumnCount(len(df.columns))
        self.tbl_preview.setRowCount(len(df))
        self.tbl_preview.setHorizontalHeaderLabels(df.columns.tolist())
        for r in range(len(df)):
            for c, col in enumerate(df.columns):
                val = "" if pd.isna(df.iloc[r, c]) else str(df.iloc[r, c])
                self.tbl_preview.setItem(r, c, QTableWidgetItem(val))
        self.tbl_preview.resizeColumnsToContents()
        self.tbl_preview.horizontalHeader().setStretchLastSection(True)

    def open_output_dir(self):
        if self.last_output_dir and self.last_output_dir.exists():
            self.open_path_in_explorer(self.last_output_dir)
        elif self.excel_path:
            self.open_path_in_explorer(self.excel_path.parent)

    def open_result_file(self):
        if self.last_result_file and self.last_result_file.exists():
            self.open_path_in_explorer(self.last_result_file)
        else:
            QMessageBox.information(self, "안내", "결과 파일이 아직 없습니다.")

    def open_path_in_explorer(self, p: Path):
        try:
            if sys.platform.startswith("win"):
                import subprocess; subprocess.Popen(["explorer", str(p)])
            else:
                import webbrowser; webbrowser.open(str(p))
        except Exception as e:
            QMessageBox.warning(self, "오류", f"열기 실패: {e}")

    def cleanup_thread(self):
        if self.thread:
            self.thread.quit(); self.thread.wait()
        self.thread = None; self.worker = None

def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except Exception: pass

    app = QApplication(sys.argv)
    icon_path = ""
    try:
        icon_path = resource_path(APP_ICON_NAME)
        if Path(icon_path).exists():
            app.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass

    w = MainWindow()
    try:
        if icon_path and Path(icon_path).exists():
            w.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass

    w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
