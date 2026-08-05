"""
extractors.py — Phase 1 內文抽取器

每個媒體集團一個 adapter，統一回傳:
    {media, date, reporter, title, body_paragraphs[], image_url}

用法:
    # 抽單一則
    from extractors import extract
    art = extract("https://money.udn.com/money/story/5613/9667796")

    # 跑完整份 urls.json，產出人工校對用的清單
    py extractors.py urls.json

設計原則（見 SPEC §6 Phase 1、§7）:
    - 抓圖一定帶 Referer（防盜連）
    - 抓不到圖 → image_url = None，不報錯
    - 每則之間 sleep(1.5)，不要拿掉
    - 任何抽取失敗都要看得見，不得靜默略過
"""
import json
import re
import sys
import time
from dataclasses import dataclass, field, asdict
from urllib.parse import urlparse

import requests
import trafilatura
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────────────
# 全域設定
# ─────────────────────────────────────────────────────────────

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Accept-Language": "zh-TW,zh;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
TIMEOUT = 20
POLITE_DELAY = 1.5  # SPEC §7.3：不要拿掉

# 無署名時的預設值（SPEC §4.2）
DEFAULT_REPORTER = "新聞中心"

# 表格第 1 列的媒體名稱。
# 預設取網頁的 og:site_name，但那不一定等於剪報上要寫的名字，
# 例如 chinatimes 的 og:site_name 是「中時新聞網」，紙本卻是「中國時報」。
# 跟前一天的成品對過之後，在這裡改就好，不要散在各個 adapter 裡。
MEDIA_OVERRIDES = {
    "www.chinatimes.com": "中國時報",   # /newspapers/ 是紙本稿
    "ec.ltn.com.tw": "自由時報",        # og:site_name 是「自由時報電子報」
}

# 這些圖不是新聞照，是網站自己的 logo / 佔位圖。
# ctee 每頁都會吐第二個 og:image 指向自家 logo，誤用的話
# 每則工商時報都會被插一張 logo，而且不會報錯（SPEC §7.4 要避免的正是這種靜默錯誤）。
IMAGE_BLACKLIST = (
    "ctee-logo",
    "static.ctee.com.tw",
    "favicon",
    "logo-main",
    "default.jpg",
    "placeholder",
)

# 這些不是人名，是「沒有署名」的意思，依 SPEC §4.2 要填「新聞中心」
NO_NAME_BYLINES = (
    "本報訊", "本報綜合報導", "綜合報導", "編輯部", "編譯", "新聞中心",
    "特派員", "記者", "本報", "整理",
)

# 通訊社供稿。剪報的媒體欄要寫「刊登的報紙」還是「供稿的通訊社」是編輯體例問題，
# 這裡一律保留刊登報紙（og:site_name），但會在 notes 標出來讓人校對時決定。
WIRE_SERVICES = ("中央社", "路透", "彭博", "法新社", "美聯社", "共同社", "新華社")

# udn.com 底下同時放聯合報系各報的稿子，署名開頭那個報名比 og:site_name 準
UDN_PAPERS = ("經濟日報", "聯合報", "聯合晚報", "udn 產經")

# 中央社供稿的署名格式是「中央社 記者張謙香港3日電」，姓名後面直接黏地點，
# 沒有任何分隔符。地點是有限集合，列出來才切得乾淨（不然會切成「張謙香港3日電」）。
CNA_PLACES = (
    "台北", "臺北", "新北", "桃園", "台中", "臺中", "台南", "臺南", "高雄",
    "香港", "東京", "北京", "上海", "首爾", "新加坡", "曼谷", "河內",
    "馬尼拉", "吉隆坡", "雅加達", "新德里", "紐約", "華盛頓", "舊金山",
    "洛杉磯", "芝加哥", "倫敦", "巴黎", "柏林", "法蘭克福", "布魯塞爾",
    "日內瓦", "維也納", "羅馬", "馬德里", "莫斯科", "雪梨", "墨爾本",
    "杜拜", "開羅", "特拉維夫",
)
# 記者{姓名}{地點}{N}日{電|專電}
CNA_BYLINE_RE = re.compile(
    r"記者([一-鿿]{2,4}?)(?:" + "|".join(CNA_PLACES) + r")\s*\d+\s*日[專]?電"
)
# 地點不在清單上時的退路：姓名取 2-3 字，後面吃掉到「N日電」為止
CNA_BYLINE_FALLBACK_RE = re.compile(
    r"記者([一-鿿]{2,3})[一-鿿]{0,5}\s*\d+\s*日[專]?電"
)

# trafilatura 會把這些網站雜訊當成內文吐出來，逐行濾掉
NOISE_PATTERNS = (
    re.compile(r"^本文共\s*\d+\s*字$"),
    re.compile(r"^字數：.*$"),
    re.compile(r"^\s*(延伸閱讀|相關新聞|更多相關新聞|推薦閱讀)\s*[:：]?\s*$"),
    re.compile(r"^※?\s*(本網站|本站|免責聲明|版權所有|不得轉載)"),
    re.compile(r"^（?示意圖|圖／|圖片來源|記者.*攝$"),
    re.compile(r"^廣告$"),
)

# 掃到這些就把「這一行以及後面全部」丟掉。
# 版權聲明、訂閱推銷這類用關鍵字擋就夠，但壹蘋的「點擊閱讀下一則新聞」後面接的是
# 別則新聞的標題，看起來跟正常內文一模一樣，只能整段截斷。
TAIL_CUTOFF_PATTERNS = (
    re.compile(r"^※\s*歡迎用"),              # 聯合報系版權聲明
    re.compile(r"點擊閱讀下一則新聞"),          # 壹蘋，後面全是相關新聞標題
    re.compile(r"^一手掌握經濟脈動"),           # 自由財經訂閱推銷
    re.compile(r"不用抽\s*不用搶"),            # 自由時報 APP 推銷
    re.compile(r"^(延伸閱讀|相關新聞|更多相關新聞|推薦閱讀|看更多)"),
    re.compile(r"(訂閱|加入).{0,8}(頻道|粉絲團|社群|LINE)"),
    re.compile(r"點我(下載|訂閱|看)"),
    re.compile(r"^(責任編輯|核稿編輯)[：:]"),
)


class ExtractError(Exception):
    """抽取失敗。呼叫端要把它印出來，不可以吞掉。"""


@dataclass
class Article:
    """一則新聞。欄位對應 SPEC §4.2 的表格與內文。"""
    url: str
    media: str = ""
    date: str = ""                       # YYYY.MM.DD
    reporter: str = DEFAULT_REPORTER
    title: str = ""
    body_paragraphs: list = field(default_factory=list)
    image_url: str = None                          # 網路新聞的主圖網址
    image_path: str = None                         # 集團新聞從 .docx 拆出來的本機圖檔
    warnings: list = field(default_factory=list)   # 可能有錯，要人工看（SPEC §7.4）
    notes: list = field(default_factory=list)      # 沒錯但值得知道，例如通訊社供稿

    @property
    def chars(self) -> int:
        return sum(len(p) for p in self.body_paragraphs)


# ─────────────────────────────────────────────────────────────
# 共用小工具
# ─────────────────────────────────────────────────────────────

def _meta(soup: BeautifulSoup, *keys: str) -> str:
    """依序找 <meta property=key> 或 <meta name=key>，回傳第一個非空的 content。

    注意：用 find_all 取「第一個」而不是最後一個。ctee 同一個 og:image
    出現兩次，第二個是 logo，取錯就會插到剪報裡。
    """
    for key in keys:
        for attr in ("property", "name"):
            for tag in soup.find_all("meta", attrs={attr: key}):
                content = (tag.get("content") or "").strip()
                if content:
                    return content
    return ""


def _jsonld_nodes(soup: BeautifulSoup):
    """吐出頁面裡所有 JSON-LD 物件（含 @graph 展開）。"""
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
        except (ValueError, TypeError):
            continue
        stack = data if isinstance(data, list) else [data]
        while stack:
            node = stack.pop(0)
            if not isinstance(node, dict):
                continue
            if "@graph" in node:
                graph = node["@graph"]
                stack.extend(graph if isinstance(graph, list) else [graph])
            yield node


def _jsonld_article(soup: BeautifulSoup) -> dict:
    """找出 NewsArticle / Article 那一個節點。"""
    for node in _jsonld_nodes(soup):
        t = str(node.get("@type", "")).lower()
        if "article" in t or "newsarticle" in t:
            return node
    return {}


def _jsonld_name(value) -> str:
    """JSON-LD 的 author/publisher 可能是字串、dict 或 list。"""
    if isinstance(value, dict):
        return str(value.get("name") or "").strip()
    if isinstance(value, list):
        names = [_jsonld_name(v) for v in value]
        return "、".join(n for n in names if n)
    return str(value or "").strip()


def normalize_date(raw: str) -> str:
    """把各站的日期格式統一成 SPEC §4.2 要求的 YYYY.MM.DD。

    吃得下 2026-08-04T03:00:00+08:00 / 2026-08-04 03:00:00 / 2026/08/04。
    """
    if not raw:
        return ""
    m = re.search(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", raw)
    if not m:
        return ""
    y, mo, d = m.groups()
    return f"{y}.{int(mo):02d}.{int(d):02d}"


def normalize_reporter(raw: str, media: str = "") -> str:
    """把各站五花八門的署名收斂成純姓名。

    實測要處理的格式:
        「經濟日報 記者林勁傑／台北即時報導」→ 林勁傑
        「記者戴玉翔／台北即時報導」        → 戴玉翔
        「工商時報 鄭妤安 呂欣芷」          → 鄭妤安、呂欣芷
        「呂淑美、鄭郁平」                  → 呂淑美、鄭郁平
        「廖家寧」                          → 廖家寧
        「中央社 記者張謙香港3日電」        → 張謙
    抽不出來回傳空字串，由呼叫端填 DEFAULT_REPORTER。
    """
    if not raw:
        return ""
    s = " ".join(raw.split())

    # 中央社體例要先攔，否則會被下面的通用規則切成「張謙香港3日電」
    m = CNA_BYLINE_RE.search(s) or CNA_BYLINE_FALLBACK_RE.search(s)
    if m:
        return m.group(1)

    # 去掉開頭的媒體名（ctee 和 money.udn 都會把媒體名黏在署名前面）
    for name in filter(None, [media, *MEDIA_OVERRIDES.values(),
                              "工商時報", "經濟日報", "聯合新聞網", "聯合報",
                              "中時新聞網", "中國時報", "自由時報電子報",
                              "自由時報", "壹蘋新聞網"]):
        if s.startswith(name):
            s = s[len(name):].strip()

    # 「記者XXX／地點報導」：取記者後面、斜線前面那段
    m = re.search(r"(?:記者|編譯|文／|特派員)\s*([^／/]+)", s)
    if m:
        s = m.group(1)
    else:
        # 沒有「記者」二字時，斜線後面通常是「台北報導」這類地點，砍掉
        s = re.split(r"[／/]", s)[0]

    # 收尾雜字
    s = re.sub(r"(即時)?(綜合)?報導$", "", s.strip()).strip()
    s = re.sub(r"^[／/、,，\s]+|[／/、,，\s]+$", "", s)

    # 多位記者的分隔符統一成頓號
    parts = [p for p in re.split(r"[、,，\s]+", s) if p]
    # 名字長度合理才收（避免把整段內文誤當成署名）
    parts = [p for p in parts if 1 < len(p) <= 6]
    # 「本報訊」「編輯部」這類等於沒署名，濾掉後由呼叫端填「新聞中心」
    parts = [p for p in parts if p not in NO_NAME_BYLINES]
    return "、".join(parts)


def _is_usable_image(url: str) -> bool:
    if not url:
        return False
    low = url.lower()
    return not any(bad in low for bad in IMAGE_BLACKLIST)


def clean_paragraphs(text: str, drop_lines=()) -> list:
    """把 trafilatura 的輸出切成段落，濾掉網站雜訊。

    drop_lines 傳入已知要剔除的整行內容（例如署名行、標題行）。
    """
    drop_norm = {" ".join(d.split()) for d in drop_lines if d}
    out = []
    for raw in text.splitlines():
        line = " ".join(raw.split())
        if not line:
            continue
        # 文末推銷／版權／相關新聞：這行以後全部不要
        if any(p.search(line) for p in TAIL_CUTOFF_PATTERNS):
            break
        if line in drop_norm:
            continue
        if any(p.search(line) for p in NOISE_PATTERNS):
            continue
        # 署名行有時前後多幾個字，用包含判斷再擋一次
        if any(d and d in line and len(line) < len(d) + 12 for d in drop_norm):
            continue
        if len(line) < 8:          # 太短的多半是標籤、按鈕
            continue
        out.append(line)
    return out


# ─────────────────────────────────────────────────────────────
# Adapter
# ─────────────────────────────────────────────────────────────

class BaseExtractor:
    """共用流程。子類只要覆寫 parse_media / parse_date / parse_reporter。"""

    DOMAINS: tuple = ()
    #  記者署名所在的 CSS selector，子類指定
    BYLINE_SELECTOR: str = None

    def fetch(self, url: str) -> str:
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        except requests.RequestException as e:
            raise ExtractError(f"連線失敗 {type(e).__name__}: {e}") from e
        if r.status_code != 200:
            raise ExtractError(f"HTTP {r.status_code}")
        r.encoding = r.apparent_encoding or r.encoding
        return r.text

    # -- 各站覆寫 --------------------------------------------------

    def parse_media(self, soup, url) -> str:
        host = urlparse(url).netloc
        if host in MEDIA_OVERRIDES:
            return MEDIA_OVERRIDES[host]
        return _meta(soup, "og:site_name", "publisher", "author")

    def parse_date(self, soup, url) -> str:
        raw = (_meta(soup, "article:published_time", "date", "date.available")
               or _jsonld_article(soup).get("datePublished", ""))
        return normalize_date(str(raw))

    def parse_reporter(self, soup, url, media) -> str:
        raw = ""
        if self.BYLINE_SELECTOR:
            el = soup.select_one(self.BYLINE_SELECTOR)
            if el:
                raw = el.get_text(" ", strip=True)
        if not raw:
            raw = _jsonld_name(_jsonld_article(soup).get("author"))
        if not raw:
            raw = _meta(soup, "dable:author", "author")
        return normalize_reporter(raw, media)

    def parse_title(self, soup, url) -> str:
        h1 = soup.find("h1")
        if h1 and h1.get_text(strip=True):
            return h1.get_text(strip=True)
        return _jsonld_article(soup).get("headline") or _meta(soup, "og:title", "title")

    # -- 共用 ------------------------------------------------------

    def image_candidates(self, soup, url) -> list:
        """依偏好順序列出候選圖，格式 (網址, 來源說明)。

        子類覆寫這個方法就能插隊，不必動下面的驗證流程。
        """
        return [(_meta(soup, "og:image", "twitter:image", "image"), "og:image")]

    def parse_image(self, soup, url, warnings) -> str:
        """依序試候選圖，回傳第一張真的抓得到的。

        全部落空就回 None，不是錯誤（SPEC §6 Phase 1）。
        """
        for candidate, source in self.image_candidates(soup, url):
            if not candidate:
                continue
            if not _is_usable_image(candidate):
                warnings.append(f"{source} 是 logo／佔位圖，已忽略：{candidate}")
                continue
            try:
                # 防盜連：一定要帶 Referer，否則多數站台回 403
                ir = requests.get(candidate, headers={**HEADERS, "Referer": url},
                                  timeout=TIMEOUT, stream=True)
                ok = (ir.status_code == 200
                      and ir.headers.get("Content-Type", "").startswith("image"))
                ir.close()
                if not ok:
                    warnings.append(
                        f"圖片下載失敗 HTTP {ir.status_code}：{candidate}")
                    continue
            except requests.RequestException as e:
                warnings.append(f"圖片下載失敗 {type(e).__name__}：{candidate}")
                continue
            if source != "og:image":
                warnings.append(f"改用{source}（og:image 是沒有資訊量的情境照）")
            return candidate
        return None

    def parse_body(self, html, soup, title, byline_text) -> list:
        text = trafilatura.extract(html, include_comments=False,
                                   include_tables=False) or ""
        return clean_paragraphs(text, drop_lines=(title, byline_text))

    def extract(self, url: str, title_hint: str = None) -> Article:
        html = self.fetch(url)
        soup = BeautifulSoup(html, "html.parser")
        art = Article(url=url)

        art.media = self.parse_media(soup, url)
        if not art.media:
            art.warnings.append("抓不到媒體名稱")

        art.date = self.parse_date(soup, url)
        if not art.date:
            art.warnings.append("抓不到日期")

        # 署名原文要留著，等下從內文裡剔除，免得重複出現在剪報上
        byline_text = ""
        if self.BYLINE_SELECTOR:
            el = soup.select_one(self.BYLINE_SELECTOR)
            byline_text = el.get_text(" ", strip=True) if el else ""

        reporter = self.parse_reporter(soup, url, art.media)
        if reporter:
            art.reporter = reporter
        else:
            art.reporter = DEFAULT_REPORTER
            art.warnings.append(f"抓不到記者，填「{DEFAULT_REPORTER}」")

        # 通訊社供稿：媒體欄仍填刊登的報紙，但標出來讓人校對時決定體例
        for wire in WIRE_SERVICES:
            if wire in byline_text:
                art.notes.append(f"{wire}供稿（媒體欄目前填「{art.media}」）")
                break

        # 標題以信件為準（同事可能已改過），信件沒給才用網頁的
        web_title = self.parse_title(soup, url)
        art.title = title_hint or web_title
        if not art.title:
            art.warnings.append("抓不到標題")

        art.body_paragraphs = self.parse_body(html, soup, web_title, byline_text)
        if art.chars < 300:
            art.warnings.append(f"內文只有 {art.chars} 字，疑似被擋或抽取失敗")

        art.image_url = self.parse_image(soup, url, art.warnings)
        return art


class UdnExtractor(BaseExtractor):
    """聯合報系：money.udn.com（經濟日報）、udn.com（聯合新聞網）。

    署名在內文容器上方的 div.article-body__info，內容像
    「經濟日報 記者林勁傑／台北即時報導」。
    udn.com 另外有乾淨的 dable:author，由 BaseExtractor 的 fallback 接手。
    """
    DOMAINS = ("money.udn.com", "udn.com")
    BYLINE_SELECTOR = "div.article-body__info span, .article-content__author"

    def parse_media(self, soup, url) -> str:
        """udn.com 的 og:site_name 一律是「聯合新聞網」，但底下同時放經濟日報和
        聯合報的稿子，署名開頭那個報名才是實際刊登的報紙。

        署名的寫法兩種都有，分隔符不一定:
            money.udn : 「經濟日報 記者林勁傑／台北即時報導」   ← 空白
            udn.com   : 「經濟日報／ 記者 戴玉翔 ／台北即時報導」← 全形斜線
        所以要用空白和斜線一起切，只切空白會得到「經濟日報／」而比對不到。
        """
        el = soup.select_one(self.BYLINE_SELECTOR)
        if el:
            text = " ".join(el.get_text(" ", strip=True).split())
            head = re.split(r"[／/\s]+", text)[0] if text else ""
            if head in UDN_PAPERS:
                return head
        return super().parse_media(soup, url)


class ChinaTimesExtractor(BaseExtractor):
    """中時集團：www.ctee.com.tw（工商時報）、www.chinatimes.com（中國時報）。

    兩者是不同 CMS，共用一個 adapter 但選擇器不同:
      - ctee 沒有 JSON-LD，記者在 <li class="publish-author">「工商時報 鄭妤安 呂欣芷」
      - chinatimes 有 JSON-LD，author already 乾淨（「呂淑美、鄭郁平」）
    ctee 的雙 og:image 陷阱由 _meta() 取第一個 + IMAGE_BLACKLIST 兩道防線處理。
    """
    DOMAINS = ("www.ctee.com.tw", "www.chinatimes.com")
    BYLINE_SELECTOR = "li.publish-author, .author"

    def image_candidates(self, soup, url) -> list:
        """資料圖表優先於 og:image。

        ctee 的圖檔名自己就分好了類:
            A02AA2_Table_Clipping_04_5.jpg        ← 記者製的資料圖表
            A02AA2_PictureItem_Clipping_04_4.jpg  ← 情境照（鈔票、大樓、人像）
        og:image 一律給情境照，但剪報要的是圖表 —— 一張「銀行看美日聯手出擊下
        的匯市」比一張日圓鈔票特寫有用得多。0804 那天 27 則裡有 4 則是這種情況。

        只掃 <article> 內，不然會撈到側欄推薦文章的圖。
        """
        article = soup.find("article")
        charts = []
        if article:
            for img in article.find_all("img"):
                src = img.get("src") or img.get("data-src") or ""
                if "_Table_" in src and src not in [c for c, _ in charts]:
                    charts.append((src, "內文資料圖表"))
        return charts + super().image_candidates(soup, url)


class LtnExtractor(BaseExtractor):
    """自由時報：ec.ltn.com.tw。記者在 JSON-LD 的 author，本來就是純姓名。"""
    DOMAINS = ("ec.ltn.com.tw", "news.ltn.com.tw")


class NextAppleExtractor(BaseExtractor):
    """壹蘋新聞網：news.nextapple.com。記者在 <meta name="author">，純姓名。"""
    DOMAINS = ("news.nextapple.com",)


# netloc → adapter。查不到的網域要明確報錯，不可以猜。
REGISTRY = {}
for _cls in (UdnExtractor, ChinaTimesExtractor, LtnExtractor, NextAppleExtractor):
    for _d in _cls.DOMAINS:
        REGISTRY[_d] = _cls()


def get_extractor(url: str) -> BaseExtractor:
    host = urlparse(url).netloc
    if host not in REGISTRY:
        raise ExtractError(f"UNSUPPORTED 未支援的網域：{host}")
    return REGISTRY[host]


def extract(url: str, title_hint: str = None) -> Article:
    """抽取單一則新聞。失敗會拋 ExtractError。"""
    return get_extractor(url).extract(url, title_hint=title_hint)


def extract_all(items: list, delay: float = POLITE_DELAY):
    """逐則抽取。回傳 [(item, Article|None, error|None)]。

    不會因為單一則失敗就中斷，但每一則的結果都要回報（SPEC §7.4）。
    """
    results = []
    total = len(items)
    for i, it in enumerate(items, 1):
        url, title = it["url"], it.get("title")
        try:
            art = extract(url, title_hint=title)
            results.append((it, art, None))
            flag = "OK  " if not art.warnings else "WARN"
            print(f"[{i:2d}/{total}] {flag} {urlparse(url).netloc:20s} "
                  f"{art.chars:4d}字 {'有圖' if art.image_url else '無圖'}  "
                  f"{(title or art.title)[:26]}")
            for w in art.warnings:
                print(f"          ! {w}")
            for n in art.notes:
                print(f"          · {n}")
        except ExtractError as e:
            results.append((it, None, str(e)))
            print(f"[{i:2d}/{total}] FAIL {urlparse(url).netloc:20s} "
                  f"{(title or '')[:26]}  ← {e}")
        if i < total:
            time.sleep(delay)
    return results


# ─────────────────────────────────────────────────────────────
# CLI：跑完整份 urls.json，產出人工校對清單
# ─────────────────────────────────────────────────────────────

def main(path: str):
    data = json.load(open(path, encoding="utf-8"))
    items = data["items"]
    results = extract_all(items)

    ok = [a for _, a, e in results if a and not a.warnings]
    warn = [a for _, a, e in results if a and a.warnings]
    fail = [(it, e) for it, a, e in results if a is None]
    noimg = [a for _, a, e in results if a and not a.image_url]

    print("\n" + "=" * 62)
    print(f"成功 {len(ok)} 則 / 有警告 {len(warn)} 則 / 失敗 {len(fail)} 則"
          f" / 無圖 {len(noimg)} 則  （共 {len(items)} 則）")
    print("=" * 62)
    if fail:
        print("\n失敗清單（這些要人工處理）:")
        for it, e in fail:
            print(f"  - {it['title'][:30]}\n    {it['url']}\n    {e}")

    # 給人目視檢查的清單（SPEC §6 Phase 1 驗收：各欄位正確率 ≥ 90%）
    lines = []
    for it, art, err in results:
        lines.append("=" * 70)
        lines.append(f"【{it.get('category', '?')}】{it['url']}")
        if err:
            lines.append(f"  !! 抽取失敗：{err}")
            continue
        lines.append(f"  媒體  : {art.media}")
        lines.append(f"  日期  : {art.date}")
        lines.append(f"  記者  : {art.reporter}")
        lines.append(f"  標題  : {art.title}")
        lines.append(f"  圖片  : {art.image_url or '（無）'}")
        lines.append(f"  內文  : {len(art.body_paragraphs)} 段 / {art.chars} 字")
        for w in art.warnings:
            lines.append(f"  警告  : {w}")
        for n in art.notes:
            lines.append(f"  備註  : {n}")
        lines.append("  ---- 內文前兩段 ----")
        for p in art.body_paragraphs[:2]:
            lines.append(f"    {p[:100]}")
    open("extract_review.txt", "w", encoding="utf-8").write("\n".join(lines))

    payload = []
    for it, art, err in results:
        row = {"category": it.get("category"), "url": it["url"], "error": err}
        if art:
            row.update(asdict(art))
        payload.append(row)
    json.dump(payload, open("extract_result.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print("\n校對清單已存 extract_review.txt")
    print("結構化結果已存 extract_result.json（Phase 3 產 Word 會吃這個）")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "urls.json")
