"""
group_news.py — Phase 2 集團新聞抽取

從媒體監測廠商的「王道銀行集團相關新聞報導摘錄.docx」裡，
依信件指定的「第N則」取出對應文章。

用法:
    py group_news.py "20260804 王道銀行集團相關新聞報導摘錄.docx" --picks 1,3
    py group_news.py "20260804 ....docx" --urls urls.json     # 從信件解析結果讀 picks

重要：「第N則」對應的是**目錄編號**，也就是文章數（本例 7 篇），
不是「報導次數」。目錄第 2、6、7 則的標題後面寫「並轉載於 A、B、C」，
那些轉載各自算一次報導但只算一篇文章。編號一律以目錄為準。

摘錄檔結構（實測 20260804 那份）:
    [ 0] 2026/08/04王道銀行集團相關新聞報導摘錄
    [ 1] 文章總數：7
    [ 2] toc 1     1. 標題 (媒體)   2      ← 目錄，樣式 toc 1
    ...
    [10] 新聞內文：
    [11] a集團新聞標題   標題 (媒體)          ← 每篇的分界，樣式固定
    [12] Normal      經濟日報 2026/08/04 記者蔡穎青報導
    [13] Normal      字數：761 words 面積：210.7 cm2 廣告價值：NT42,306
    [14] Normal      （圖片）
    [15] Normal      內文第一段…
"""
import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from extractors import Article, DEFAULT_REPORTER, normalize_date, normalize_reporter

# 每篇文章的標題都用這個自訂樣式，是最可靠的分界點
TITLE_STYLE = "a集團新聞標題"

# 「經濟日報 2026/08/04 記者蔡穎青報導」
# 媒體名可能含空白（實測有「CNEWS 匯流新聞網」也有「CNEWS匯流新聞網」兩種寫法），
# 所以用非貪婪比對到日期為止，不能假設媒體名是單一 token。
BYLINE_RE = re.compile(
    r"^(?P<media>.+?)\s+(?P<date>\d{4}/\d{1,2}/\d{1,2})\s+(?P<reporter>.+?報導)\s*$"
)
# 「字數：761 words 面積：210.7 cm2 廣告價值：NT42,306」— 廠商的統計資訊，不是內文
STATS_RE = re.compile(r"^字數：|^面積：|^廣告價值：")
# 標題結尾的「 (經濟日報)」「 (中央社…，並轉載於…)」是廠商加註的出處，剪報標題不要
TITLE_SOURCE_RE = re.compile(r"\s*[（(][^（()]*[）)]\s*$")


def _para_image_blob(para, doc):
    """取出段落裡第一張圖的二進位內容，沒有就回 None。"""
    for blip in para._p.findall(".//" + qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid:
            try:
                return doc.part.related_parts[rid].blob
            except KeyError:
                continue
    return None


def _ext_from_blob(blob: bytes) -> str:
    if blob.startswith(b"\x89PNG"):
        return ".png"
    if blob.startswith(b"GIF8"):
        return ".gif"
    return ".jpg"


def parse_excerpt(path, image_dir="group_images") -> list:
    """把摘錄 .docx 解析成 Article 清單，順序等同目錄編號。

    回傳的第 i 篇對應信裡的「第 i+1 則」。
    """
    doc = Document(str(path))
    paras = doc.paragraphs

    # 廠商自己寫的文章總數，拿來跟實際解析出的篇數對帳
    declared = None
    for p in paras[:6]:
        m = re.search(r"文章總數：\s*(\d+)", p.text)
        if m:
            declared = int(m.group(1))
            break

    # 找出所有標題段落的位置，相鄰兩個標題之間就是一篇
    bounds = [i for i, p in enumerate(paras) if p.style.name == TITLE_STYLE]
    if not bounds:
        raise SystemExit(
            f"找不到樣式「{TITLE_STYLE}」的段落，摘錄檔格式可能改版了。\n"
            f"請確認 {path} 是媒體監測廠商產出的原始檔。"
        )

    image_dir = Path(image_dir)
    articles = []

    for n, start in enumerate(bounds, 1):
        end = bounds[n] if n < len(bounds) else len(paras)
        art = Article(url=f"（集團新聞第{n}則）")

        raw_title = paras[start].text.strip()
        art.title = TITLE_SOURCE_RE.sub("", raw_title).strip()

        body, byline_seen = [], False
        for idx in range(start + 1, end):
            para = paras[idx]
            text = " ".join(para.text.split())

            # 圖片：每篇只取第一張（SPEC §4.2）
            if art.image_path is None:
                blob = _para_image_blob(para, doc)
                if blob:
                    image_dir.mkdir(parents=True, exist_ok=True)
                    dest = image_dir / f"group_{n:02d}{_ext_from_blob(blob)}"
                    dest.write_bytes(blob)
                    art.image_path = str(dest)

            if not text:
                continue

            if not byline_seen:
                m = BYLINE_RE.match(text)
                if m:
                    art.media = m.group("media")
                    art.date = normalize_date(m.group("date"))
                    art.reporter = (normalize_reporter(m.group("reporter"), art.media)
                                    or DEFAULT_REPORTER)
                    byline_seen = True
                    continue

            if STATS_RE.match(text):        # 廠商統計資訊
                continue
            if text == art.title:           # 標題有時在內文重複一次
                continue
            body.append(text)

        art.body_paragraphs = body

        if not byline_seen:
            art.warnings.append("找不到「媒體 日期 記者報導」那一行，媒體／日期／記者是空的")
        if not body:
            art.warnings.append("解析不到內文")
        articles.append(art)

    if declared is not None and declared != len(articles):
        # 不擋流程，但要讓人看到（SPEC §7.4）
        print(f"  ! 注意：檔案寫「文章總數：{declared}」，實際解析出 {len(articles)} 篇",
              file=sys.stderr)

    return articles


def pick(articles: list, picks: list) -> list:
    """依信件的「第N則」取出文章。N 從 1 開始，超出範圍會明確報錯。"""
    out = []
    for n in picks:
        if not 1 <= n <= len(articles):
            raise SystemExit(
                f"信件指定「第{n}則」，但摘錄檔只有 {len(articles)} 篇。"
                f"請確認信件與附件是同一天的。"
            )
        out.append(articles[n - 1])
    return out


def main():
    ap = argparse.ArgumentParser(description="從摘錄 .docx 取出集團新聞")
    ap.add_argument("docx", help="王道銀行集團相關新聞報導摘錄.docx")
    ap.add_argument("--picks", help="逗號分隔的則數，例如 1,3")
    ap.add_argument("--urls", help="parse_email.py 產出的 urls.json，從中讀 picks")
    ap.add_argument("--json", dest="out_json", help="把結果寫成 JSON")
    args = ap.parse_args()

    articles = parse_excerpt(args.docx)

    print(f"摘錄檔共 {len(articles)} 篇：")
    for i, a in enumerate(articles, 1):
        print(f"  {i}. [{a.media} / {a.date} / {a.reporter}] {a.title[:44]}"
              f"  ({len(a.body_paragraphs)}段 {a.chars}字"
              f"{' 有圖' if a.image_path else ''})")

    picks = []
    if args.picks:
        picks = [int(x) for x in re.findall(r"\d+", args.picks)]
    elif args.urls:
        picks = json.load(open(args.urls, encoding="utf-8"))["meta"]["group_news_picks"]

    if not picks:
        print("\n（沒有指定 --picks 或 --urls，只列出全部，不挑選）")
        return

    chosen = pick(articles, picks)
    print(f"\n信件指定第 {picks} 則 → 取出 {len(chosen)} 篇：")
    for a in chosen:
        print("=" * 66)
        print(f"  媒體 : {a.media}")
        print(f"  日期 : {a.date}")
        print(f"  記者 : {a.reporter}")
        print(f"  標題 : {a.title}")
        print(f"  圖片 : {a.image_path or '（無）'}")
        print(f"  內文 : {len(a.body_paragraphs)} 段 / {a.chars} 字")
        for p in a.body_paragraphs[:2]:
            print(f"    {p[:76]}")
        for w in a.warnings:
            print(f"  ! {w}")

    if args.out_json:
        from dataclasses import asdict
        json.dump([asdict(a) for a in chosen],
                  open(args.out_json, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
        print(f"\n已寫出 {args.out_json}")


if __name__ == "__main__":
    main()
