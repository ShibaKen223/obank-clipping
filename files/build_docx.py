"""
build_docx.py — Phase 3 產出 Word 剪報

作法：拿前一天的成品當底稿，保留封面、清掉舊內容，再把今天的新聞填進去。

為什麼不從空白文件建：
    實檔是「黑底白字」的版型（document.xml 有 <w:background w:color="000000"/>），
    表格是黑底、白框線、浮動定位，字色來自兩個自訂樣式「榮董新聞內文」和「榮董資訊」。
    從空白文件用 python-docx 重建會變成白底黑字的普通表格，完全不像。
    沿用底稿可以連同背景、樣式、頁面尺寸、頁尾一次繼承。

用法:
    py build_docx.py --template "20260803 每日新聞剪報.docx" \
                     --articles extract_result.json \
                     --excerpt "20260804 王道銀行集團相關新聞報導摘錄.docx" \
                     --urls urls.json \
                     --out "outputs/20260804 每日新聞剪報.docx"
"""
import argparse
import copy
import io
import json
import re
import sys
from pathlib import Path

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

from extractors import HEADERS, TIMEOUT, DEFAULT_REPORTER

# ─────────────────────────────────────────────────────────────
# 版面常數（全部從 20260803 實檔量測，不是從 SPEC 抄的）
# ─────────────────────────────────────────────────────────────

BODY_STYLE = "榮董新聞內文"     # 18pt 白字，標題與內文共用
CELL_STYLE = "榮董資訊"         # 16pt 白字，表格儲存格用
TITLE_PT = 24                   # 標題字級
TEXT_WIDTH_IN = 10.12           # 文字欄寬，圖片最寬就是這個
IMG_FAIL_COLOR = RGBColor(0xFF, 0x00, 0x00)   # 圖片待補用紅字（SPEC §7.4）

# 內文提到自家的地方要標起來（粗體＋紅色粗底線）。
# 底稿本來就帶著這個字元樣式，直接套它，格式細節就跟人工版同一個來源；
# 哪天要調整粗細或顏色，改 Word 裡的樣式就好，不用動程式。
HIGHLIGHT_STYLE = "榮董新聞醒目 字元"

# 要標哪些字。實檔標的是整串「王道銀行（O-Bank）」，所以括號那段要一起吃掉，
# 不然會斷成「王道銀行」有底線、「（O-Bank）」沒有。
# 只認完整的行名，不標單獨的「王道」—— 內文裡的「王道薪轉戶」「王道簽帳金融卡」
# 標起來會滿篇紅線，反而看不出重點在哪。
HIGHLIGHT_RE = re.compile(r"王道銀行\s*[（(]\s*O-?Bank\s*[）)]|王道銀行|O-Bank")

# 封面目錄的頁碼欄從第幾個半形字元開始。
# 實檔七行目錄的「類別名 + 空白」加起來一律是 25 個半形寬（全形算 2），
# 頁碼因此對得整整齊齊。不補這些空白的話，P.1~7 會緊貼在類別名後面，
# 七行長短不一，封面一眼就看得出不是同一份東西做出來的。
TOC_PAGE_COLUMN = 25

# urls.json 的類別 → 封面目錄上的寫法（實檔用詞跟信件不完全一樣）
CATEGORY_ORDER = [
    ("集團新聞", "集團新聞"),
    ("金融同業", "金融同業新聞"),
    ("主管機關與金融", "主管機關與金融政策"),
    ("其他政府機關及經濟新聞", "其他政府機關及經濟新聞"),
    ("產業新聞", "產業新聞"),
    ("大陸新聞", "大陸新聞"),
    ("國際新聞", "國際新聞"),
]


# ─────────────────────────────────────────────────────────────
# XML 小工具
# ─────────────────────────────────────────────────────────────

def _err(e: Exception) -> str:
    """有些例外的訊息是空字串（python-docx 的 UnrecognizedImageError 就是），
    只印 str(e) 會變成一片空白，看不出發生什麼事。"""
    return f"{type(e).__name__}: {e}" if str(e) else type(e).__name__


def _display_width(s: str) -> int:
    """字串在等寬字型下佔幾個半形位。中日韓字元算 2。"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in s)


def toc_line(label: str, pages: str = "__~__") -> str:
    """組一行封面目錄，頁碼對齊到 TOC_PAGE_COLUMN。

    產 Word 時 pages 是佔位的 "__~__"，paginate.py 量到真頁碼後會用
    同一個函式重寫一次，兩次的空白數才會一致。
    """
    pad = max(1, TOC_PAGE_COLUMN - _display_width(label))
    return f"{label}{' ' * pad}P.{pages}"


def _has_page_break(p_el) -> bool:
    return bool(p_el.findall(
        ".//" + qn("w:br") + "[@" + qn("w:type") + "='page']"))


def set_paragraph_text(para, text: str):
    """換掉段落文字，但保留原本第一個 run 的字型設定與段落屬性。

    直接改 para.text 會把 rPr 洗掉（字級、粗體、顏色全沒），所以要自己來。
    """
    # 第一個 run 可能不在段落底下，而是包在 w:hyperlink 之類的容器裡
    # （Word 的目錄項目就是超連結），所以要用 .// 全域找。
    runs = para._p.findall(".//" + qn("w:r"))
    proto_rpr = None
    if runs:
        rpr = runs[0].find(qn("w:rPr"))
        if rpr is not None:
            proto_rpr = copy.deepcopy(rpr)

    # 除了 pPr（段落樣式、對齊、框架定位）以外全部清掉。
    # 只刪 w:r 是不夠的：hyperlink / smartTag / ins 裡面的舊文字會留下來，
    # 新文字接上去就變成「主管機關與金融政策 P.主管機關與金融政策 P.__~__」。
    for el in list(para._p):
        if el.tag != qn("w:pPr"):
            para._p.remove(el)

    r = para._p.makeelement(qn("w:r"), {})
    if proto_rpr is not None:
        r.append(proto_rpr)
    t = para._p.makeelement(qn("w:t"), {})
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    para._p.append(r)


def append_block(doc, element):
    """把元素插到 body 尾端、但在 sectPr 之前（sectPr 帶著頁面設定，不能被推到後面）。"""
    sect_pr = doc.element.body.find(qn("w:sectPr"))
    if sect_pr is not None:
        sect_pr.addprevious(element)
    else:
        doc.element.body.append(element)


# ─────────────────────────────────────────────────────────────
# 底稿處理
# ─────────────────────────────────────────────────────────────

class Proto:
    """從底稿複製下來的版型零件。

    這些東西自己用 python-docx 的 API 造都會漏掉細節（黑底、白框線、浮動定位、
    看不見卻很關鍵的字級），複製底稿的 XML 最保險。
    """

    def __init__(self, table, gap, title):
        self.table = table    # 媒體／日期／記者的三列表格
        self.gap = gap        # 表格與標題之間的空段落
        self.title = title    # 24pt 粗體置中的標題段落


def _block_prototypes(tbl_el, path):
    """從底稿的一則新聞裡，抓出「空段落」與「標題段落」兩個原型。

    底稿的一則長這樣：表格 → 空段落 ×2 → 標題 → (圖) → 內文。

    那兩個空段落看起來只是留白，其實帶著 24pt 字級與 12pt 段前距，
    兩個疊起來約 1 吋。這高度是必要的：資訊表是浮動表格（tblpPr），
    Word 會讓後面的文字繞到它右邊，高度不夠標題就跑到表格旁邊去。
    自己 add_paragraph() 造的是 12pt 無段距，只有需要高度的三分之一。
    """
    gap = None
    el = tbl_el.getnext()
    while el is not None and el.tag == qn("w:p"):
        if el.findall(".//" + qn("w:t")):          # 有字了，這是標題
            if gap is None:
                break
            return copy.deepcopy(gap), copy.deepcopy(el)
        if gap is None:
            gap = el
        el = el.getnext()

    raise SystemExit(
        f"底稿 {path} 的表格後面不是「空段落＋標題」，版型可能被改過。\n"
        f"換一份確定正常的成品當底稿（--template）。")


def load_template(path):
    """開啟底稿，保留封面，清掉第一個分頁之後的所有內容。

    回傳 (doc, 版型零件)。零件要在清空前先複製起來，清空後就找不到了。
    """
    doc = Document(str(path))

    if not doc.tables:
        raise SystemExit(f"底稿 {path} 裡沒有表格，可能不是每日新聞剪報的成品檔")
    # 挑一個欄寬中庸的當原型（欄寬是自動調整的，複製哪個都會再依內容伸縮）
    tbl_el = doc.tables[len(doc.tables) // 2]._tbl
    gap_el, title_el = _block_prototypes(tbl_el, path)
    proto = Proto(table=copy.deepcopy(tbl_el), gap=gap_el, title=title_el)

    body = doc.element.body
    children = list(body.iterchildren())
    cut = next((i for i, ch in enumerate(children)
                if ch.tag == qn("w:p") and _has_page_break(ch)), None)
    if cut is None:
        raise SystemExit(f"底稿 {path} 裡找不到分頁符號，版型可能不同")

    for ch in children[cut:]:
        if ch.tag != qn("w:sectPr"):
            body.remove(ch)

    dropped = prune_unused_images(doc)
    if dropped:
        print(f"  清掉底稿殘留的 {dropped} 張孤兒圖")

    return doc, proto


def prune_unused_images(doc) -> int:
    """清掉底稿裡已經沒人引用的圖片本體。

    上面清內容只是把 XML 段落刪掉，圖檔本身還留在 package 裡照樣被存出去。
    不清的話會滾雪球：今天的成品是明天的底稿，孤兒圖一天疊一天，
    實測一天就從 4MB 漲到 16MB，幾天後 Gmail 就寄不出去了。

    只掃 document.xml，頁首頁尾是獨立的 part、有自己的 rels，不會被動到。
    """
    used = set()
    for el in doc.element.body.iter():
        for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
            rid = el.get(attr)
            if rid:
                used.add(rid)

    part = doc.part
    dropped = 0
    for rid, rel in list(part.rels.items()):
        if not rel.is_external and "image" in rel.reltype and rid not in used:
            part.drop_rel(rid)
            dropped += 1
    return dropped


def update_cover(doc, date_label: str, toc_lines: list):
    """更新封面的大標日期與目錄。

    封面結構（實檔）:
        圖片 → 「2026/8/3每日新聞剪報」40pt 置中 → 7 行目錄
    目錄用了 toc 1 和 List Paragraph 兩種樣式（原檔就不一致），
    這裡照原本的段落逐行換字，不動樣式。
    """
    paras = doc.paragraphs
    title_done = False
    toc_idx = []

    for p in paras:
        sizes = {r.font.size.pt for r in p.runs if r.font.size}
        text = " ".join(p.text.split())
        if not title_done and 40.0 in sizes and "每日新聞剪報" in text:
            set_paragraph_text(p, date_label)
            title_done = True
        elif text and re.search(r"P\.\s*\d+\s*[~～]\s*\d+", text):
            toc_idx.append(p)

    if not title_done:
        print("  ! 封面找不到 40pt 的大標，日期沒更新", file=sys.stderr)

    for p, line in zip(toc_idx, toc_lines):
        set_paragraph_text(p, line)
    # 今天類別比較少時，多出來的目錄行清空
    for p in toc_idx[len(toc_lines):]:
        set_paragraph_text(p, "")
    if len(toc_idx) < len(toc_lines):
        print(f"  ! 封面只有 {len(toc_idx)} 行目錄，今天有 {len(toc_lines)} 個類別，"
              f"後面 {len(toc_lines) - len(toc_idx)} 行要手動補", file=sys.stderr)


# ─────────────────────────────────────────────────────────────
# 一則新聞
# ─────────────────────────────────────────────────────────────

def add_page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


# 底稿 35 個表格後面全部都是 2 個空段落，這是版型的一部分，不是誰多按了 Enter。
TABLE_GAP_PARAS = 2


def add_table_gap(doc, proto_gap):
    """表格與標題之間的留白。少了它標題會被浮動表格繞排到右邊。

    複製底稿的原段落而不是自己造，理由見 _gap_prototype()：
    那個「空」段落的字級決定了高度，高度決定了標題落在哪。
    """
    for _ in range(TABLE_GAP_PARAS):
        append_block(doc, copy.deepcopy(proto_gap))


# 表格欄寬估算用（單位 twips，1 in = 1440）
_CJK_W = 330        # 16pt 全形字約佔的寬度
_ASCII_W = 180      # 半形字
_CELL_PAD = 300     # 左右內距
_MIN_CELL_W = 2830  # 底稿 35 個表格裡有 24 個是這個寬度（1.97in），拿來當下限


def _content_width_twips(*texts) -> int:
    """依最長的一格內容估算欄寬。

    寧可寬一點也不要讓文字換行 —— 表格是黑底配黑色頁面背景，
    框線又是白的（等於看不見），所以多出來的寬度不會被看到，
    但換行會讓那一則的框比別則高一截，一眼就看得出來。
    """
    widest = 0
    for t in texts:
        w = sum(_CJK_W if ord(ch) > 0x2E80 else _ASCII_W for ch in (t or ""))
        widest = max(widest, w)
    return max(_MIN_CELL_W, widest + _CELL_PAD)


def add_info_table(doc, proto_tbl, media, date, reporter):
    """複製底稿的表格原型，換掉三格文字。

    不自己 add_table 是因為原型帶著黑底、白框線、浮動定位（tblpPr）
    這些用 python-docx API 很難重建，複製 XML 最保險。
    """
    tbl_el = copy.deepcopy(proto_tbl)
    append_block(doc, tbl_el)
    tbl = Table(tbl_el, doc)

    for row, text in zip(tbl.rows, (media, date, reporter)):
        cell = row.cells[0]
        # 多餘的段落刪掉，只留第一段
        for extra in cell.paragraphs[1:]:
            extra._p.getparent().remove(extra._p)
        set_paragraph_text(cell.paragraphs[0], text or "")

    # 原型帶著它自己那則的固定欄寬，直接沿用的話每則的框都一樣寬，
    # 遇到長媒體名（例如「中央社財經訊息平台」9 個字）就會換行。
    # 底稿的 1.77~3.25in 是 Word 依內容自動算出來的，這裡照算一次。
    width = _content_width_twips(media, date, reporter)
    for tc_w in tbl_el.findall(".//" + qn("w:tcW")):
        tc_w.set(qn("w:w"), str(width))
        tc_w.set(qn("w:type"), "dxa")
    for grid_col in tbl_el.findall(".//" + qn("w:gridCol")):
        grid_col.set(qn("w:w"), str(width))
    return tbl


def add_title(doc, proto_title, title: str):
    """標題也複製底稿的原段落，只換文字。

    自己組出來的雖然也是 24pt 粗體置中，但少了 bCs／szCs（中日韓字型的粗體與
    字級是分開的屬性）和 rFonts hint，細看就不一樣。既然底稿在手上，用複製的。
    """
    el = copy.deepcopy(proto_title)
    append_block(doc, el)
    p = Paragraph(el, doc)
    set_paragraph_text(p, title)
    return p


def fetch_image(url: str, referer: str = None):
    """下載圖片。一定要帶 Referer，否則多數新聞站回 403。"""
    headers = dict(HEADERS)
    if referer:
        headers["Referer"] = referer
    r = requests.get(url, headers=headers, timeout=TIMEOUT)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}")
    if not r.headers.get("Content-Type", "").startswith("image"):
        raise RuntimeError(f"回傳的不是圖片（{r.headers.get('Content-Type')}）")
    return io.BytesIO(r.content)


def normalize_image(stream):
    """把 python-docx 不認得的圖轉成標準 JPEG。

    python-docx 判斷 JPEG 只認 JFIF(\\xff\\xd8\\xff\\xe0) 和 EXIF(\\xff\\xd8\\xff\\xe1)
    兩種開頭。壹蘋的 CDN 吐的是第一個標記為 DQT(\\xff\\xd8\\xff\\xdb) 的裸 JPEG，
    檔案完全合法但會被判成無法辨識（而且丟出的例外訊息是空字串）。
    用 Pillow 重存一次就會補上正規的檔頭。
    """
    from PIL import Image

    stream.seek(0)
    img = Image.open(stream)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=92)
    out.seek(0)
    return out


def add_image(doc, stream) -> bool:
    """插圖，置中，寬度不超過文字欄寬，維持長寬比。

    先試原始檔（不重新編碼、不損畫質），被拒才用 Pillow 轉一次。
    """
    try:
        pic = doc.add_picture(stream)
    except Exception:
        if isinstance(stream, (str, Path)):
            raise
        pic = doc.add_picture(normalize_image(stream))
    limit = Inches(TEXT_WIDTH_IN)
    if pic.width > limit:
        pic.height = int(pic.height * limit / pic.width)
        pic.width = limit
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    return True


def _mark_highlight(run):
    """套上「粗體＋紅色粗底線」。

    優先用底稿的字元樣式；底稿沒有那個樣式時（換了很舊的成品當底稿）
    退回自己畫，效果一樣，只是之後改樣式不會跟著動。
    """
    run.bold = True
    rpr = run._element.get_or_add_rPr()
    u = rpr.makeelement(qn("w:u"), {})
    u.set(qn("w:val"), "thick")
    u.set(qn("w:color"), "FF0000")
    rpr.append(u)


def add_body_paragraph(doc, text: str, highlight: bool, stats: dict):
    """加一段內文，把提到自家的地方標起來。

    整段一個 run 是不行的：底線只能標在 run 上，要標局部就得把段落切成
    「前面的字 / 要標的字 / 後面的字」三種 run 交錯排。
    """
    p = doc.add_paragraph(style=BODY_STYLE)
    if not highlight:
        p.add_run(text)
        return p

    pos = 0
    for m in HIGHLIGHT_RE.finditer(text):
        if m.start() > pos:
            p.add_run(text[pos:m.start()])
        run = p.add_run(m.group())
        try:
            run.style = doc.styles[HIGHLIGHT_STYLE]
        except KeyError:
            _mark_highlight(run)
        stats["標記"] += 1
        pos = m.end()
    if pos < len(text):
        p.add_run(text[pos:])
    return p


def add_image_placeholder(doc, url: str):
    """圖片抓失敗時留紅字，讓失敗看得見（SPEC §7.4），不要靜默跳過。"""
    p = doc.add_paragraph(style=BODY_STYLE)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"【圖片待補：{url}】")
    run.font.color.rgb = IMG_FAIL_COLOR
    run.bold = True


def add_news_block(doc, art: dict, proto: "Proto", stats: dict):
    """產生一則新聞：分頁 → 表格 → 留白 → 標題 → 圖片 → 內文（順序見實檔量測）。"""
    add_page_break(doc)
    add_info_table(doc, proto.table,
                   art.get("media", ""),
                   art.get("date", ""),
                   art.get("reporter") or DEFAULT_REPORTER)
    add_table_gap(doc, proto.gap)
    add_title(doc, proto.title, art.get("title", ""))

    # 圖片：集團新聞是本機檔，網路新聞是網址
    local = art.get("image_path")
    url = art.get("image_url")
    if local and Path(local).exists():
        try:
            add_image(doc, str(local))
            stats["圖片成功"] += 1
        except Exception as e:
            add_image_placeholder(doc, local)
            stats["圖片待補"] += 1
            print(f"    ! 本機圖片插入失敗 {local}：{_err(e)}")
    elif url:
        try:
            add_image(doc, fetch_image(url, referer=art.get("url")))
            stats["圖片成功"] += 1
        except Exception as e:
            add_image_placeholder(doc, url)
            stats["圖片待補"] += 1
            print(f"    ! 圖片失敗，已標紅字待補：{_err(e)}")
    else:
        stats["本來就無圖"] += 1

    for para in art.get("body_paragraphs", []):
        add_body_paragraph(doc, para, highlight=True, stats=stats)


# ─────────────────────────────────────────────────────────────
# 組裝
# ─────────────────────────────────────────────────────────────

def group_by_category(articles: list, group_articles: list) -> list:
    """依 SPEC §2.1 的固定順序排列，當天沒有的類別直接跳過。"""
    buckets = {label: [] for _, label in CATEGORY_ORDER}
    buckets["集團新聞"] = [dict(a, category="集團新聞") for a in group_articles]

    label_of = dict(CATEGORY_ORDER)
    for a in articles:
        if a.get("error"):
            continue
        label = label_of.get(a.get("category"), a.get("category"))
        buckets.setdefault(label, []).append(a)

    return [(label, buckets[label]) for _, label in CATEGORY_ORDER if buckets.get(label)]


def build(template, articles, group_articles, out_path, date_label):
    doc, proto = load_template(template)
    sections = group_by_category(articles, group_articles)

    toc_lines = [toc_line(label) for label, _ in sections]
    update_cover(doc, date_label, toc_lines)

    stats = {"則數": 0, "圖片成功": 0, "圖片待補": 0, "本來就無圖": 0, "標記": 0}
    for label, items in sections:
        print(f"\n【{label}】{len(items)} 則")
        for art in items:
            print(f"  - {art.get('title', '')[:34]}")
            add_news_block(doc, art, proto, stats)
            stats["則數"] += 1

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path, stats, sections


def main():
    ap = argparse.ArgumentParser(description="產出每日新聞剪報 Word 檔")
    ap.add_argument("--template", required=True, help="前一天的成品，當底稿用")
    ap.add_argument("--articles", default="extract_result.json",
                    help="extractors.py 的產出")
    ap.add_argument("--excerpt", help="集團新聞摘錄 .docx")
    ap.add_argument("--urls", default="urls.json", help="parse_email.py 的產出")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    articles = json.load(open(args.articles, encoding="utf-8"))

    group_articles = []
    if args.excerpt:
        from dataclasses import asdict
        from group_news import parse_excerpt, pick
        picks = json.load(open(args.urls, encoding="utf-8")
                          )["meta"]["group_news_picks"]
        all_group = parse_excerpt(args.excerpt)
        group_articles = [asdict(a) for a in pick(all_group, picks)]
        print(f"集團新聞：摘錄檔 {len(all_group)} 篇，取第 {picks} 則")

    meta = json.load(open(args.urls, encoding="utf-8"))["meta"]
    m = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})",
                  next((a["date"] for a in group_articles if a.get("date")), "")
                  .replace(".", "/")) or None
    subject = meta.get("subject", "")
    date_label = None
    if m:
        date_label = f"{m.group(1)}/{int(m.group(2))}/{int(m.group(3))}每日新聞剪報"
    else:
        mm = re.match(r"(\d{2})(\d{2})_", subject or "")
        if mm:
            date_label = f"2026/{int(mm.group(1))}/{int(mm.group(2))}每日新聞剪報"
    if not date_label:
        date_label = "每日新聞剪報"
        print("  ! 推不出日期，封面大標請手動改", file=sys.stderr)

    out, stats, sections = build(args.template, articles, group_articles,
                                 args.out, date_label)

    print("\n" + "=" * 60)
    print(f"已產出 {out}")
    print(f"  共 {stats['則數']} 則"
          f"（{' / '.join(f'{l} {len(i)}' for l, i in sections)}）")
    print(f"  圖片：成功 {stats['圖片成功']} / 待補 {stats['圖片待補']}"
          f" / 本來就沒圖 {stats['本來就無圖']}")
    print(f"  已標記「王道銀行／O-Bank」{stats['標記']} 處")
    if stats["圖片待補"]:
        print(f"  ! 有 {stats['圖片待補']} 處紅字【圖片待補】，請在 Word 裡搜尋補上")
    print("  ! 封面目錄的頁碼是 P.__~__ 佔位，請開檔後依實際頁數手填（SPEC §4.3）")
    print("=" * 60)


if __name__ == "__main__":
    main()
