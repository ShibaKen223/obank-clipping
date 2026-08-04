"""
main.py — 每日新聞剪報，一鍵從信件產出 Word 草稿

每天的標準流程:
    py main.py 0804_榮董新聞.eml

只想重產 Word（不重新抓網頁，省 2 分鐘也不打擾對方伺服器）:
    py main.py 0804_榮董新聞.eml --skip-fetch

流程:
    .eml
      ├─ parse_email.py  → urls.json + attachments/（兩個 .docx）
      ├─ extractors.py   → extract_result.json（27 則內文與圖）
      ├─ group_news.py   → 從摘錄 .docx 取出「第N則」
      └─ build_docx.py   → outputs/YYYYMMDD 每日新聞剪報.docx

產出是**草稿**。依 SPEC §7.1，本程式絕不寄信，校對後請自行寄出。
"""
import argparse
import json
import re
import sys
from dataclasses import asdict
from pathlib import Path

import build_docx
import extractors
import group_news
import parse_email

ATTACH_DIR = Path("attachments")
OUT_DIR = Path("outputs")

# 兩個附件靠檔名認：一個是媒體監測的摘錄，一個是前一天的成品當底稿
EXCERPT_HINT = "摘錄"
TEMPLATE_HINT = "每日新聞剪報"


def find_attachment(paths, hint, label):
    hits = [p for p in paths if hint in Path(p).name]
    if not hits:
        raise SystemExit(
            f"在附件裡找不到{label}（檔名要含「{hint}」）。\n"
            f"目前的附件：{[Path(p).name for p in paths] or '（一個都沒有）'}\n"
            f"可以用 --excerpt / --template 手動指定路徑。"
        )
    if len(hits) > 1:
        print(f"  ! 有 {len(hits)} 個檔名含「{hint}」，取第一個：{Path(hits[0]).name}")
    return hits[0]


def date_stamp(meta, excerpt_path) -> str:
    """決定產出檔名的日期。優先用摘錄檔名的 YYYYMMDD，其次用信件主旨的 MMDD。"""
    m = re.search(r"(20\d{6})", Path(excerpt_path).name if excerpt_path else "")
    if m:
        return m.group(1)
    m = re.match(r"(\d{2})(\d{2})_", meta.get("subject", "") or "")
    if m:
        from datetime import date
        return f"{date.today().year}{m.group(1)}{m.group(2)}"
    from datetime import date
    print("  ! 推不出日期，用今天的日期當檔名", file=sys.stderr)
    return date.today().strftime("%Y%m%d")


def step(n, total, msg):
    print(f"\n{'━' * 62}\n[{n}/{total}] {msg}\n{'━' * 62}")


def main():
    ap = argparse.ArgumentParser(
        description="從榮董新聞信件產出每日新聞剪報 Word 草稿")
    ap.add_argument("eml", nargs="?", help="信件檔，例如 0804_榮董新聞.eml")
    ap.add_argument("--urls", help="改用現成的 urls.json，跳過信件解析")
    ap.add_argument("--excerpt", help="手動指定集團新聞摘錄 .docx")
    ap.add_argument("--template", help="手動指定底稿（前一天的成品）")
    ap.add_argument("--skip-fetch", action="store_true",
                    help="沿用現有的 extract_result.json，不重新抓網頁")
    ap.add_argument("--out", help="輸出路徑，預設 outputs/YYYYMMDD 每日新聞剪報.docx")
    args = ap.parse_args()

    if not args.eml and not args.urls:
        ap.error("要給 .eml 檔，或用 --urls 指定現成的 urls.json")

    total = 4

    # ── 1. 解析信件 ────────────────────────────────────────
    step(1, total, "解析信件")
    if args.eml:
        eml = Path(args.eml)
        if not eml.exists():
            raise SystemExit(f"找不到信件檔 {eml}")
        result = parse_email.parse(eml)
        result["meta"]["attachments"] = parse_email.save_attachments(eml, ATTACH_DIR)
        Path("urls.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
        print(f"  已寫出 urls.json")
    else:
        result = json.load(open(args.urls, encoding="utf-8"))
        print(f"  沿用現成的 {args.urls}")

    meta = result["meta"]
    items = result["items"]
    picks = meta.get("group_news_picks", [])
    attachments = meta.get("attachments", [])

    print(f"  主旨    : {meta.get('subject')}")
    print(f"  網址    : {len(items)} 則")
    print(f"  集團新聞: 第 {picks} 則")
    print(f"  附件    : {len(attachments)} 個")

    excerpt = args.excerpt or (find_attachment(attachments, EXCERPT_HINT, "集團新聞摘錄檔")
                               if picks else None)
    template = args.template or find_attachment(attachments, TEMPLATE_HINT, "底稿")

    # ── 2. 抓網路新聞 ──────────────────────────────────────
    step(2, total, f"抓取 {len(items)} 則網路新聞")
    result_json = Path("extract_result.json")
    if args.skip_fetch and result_json.exists():
        articles = json.loads(result_json.read_text("utf-8"))
        print(f"  --skip-fetch：沿用現有的 {result_json}（{len(articles)} 則）")
    else:
        results = extractors.extract_all(items)
        articles = []
        for it, art, err in results:
            row = {"category": it.get("category"), "url": it["url"], "error": err}
            if art:
                row.update(asdict(art))
            articles.append(row)
        result_json.write_text(
            json.dumps(articles, ensure_ascii=False, indent=2), "utf-8")

    failed = [a for a in articles if a.get("error")]
    warned = [a for a in articles if not a.get("error") and a.get("warnings")]

    # ── 3. 集團新聞 ────────────────────────────────────────
    step(3, total, "抽取集團新聞")
    group_articles = []
    if picks and excerpt:
        all_group = group_news.parse_excerpt(excerpt)
        print(f"  摘錄檔 {len(all_group)} 篇，信件指定第 {picks} 則")
        group_articles = [asdict(a) for a in group_news.pick(all_group, picks)]
        for a in group_articles:
            print(f"    - [{a['media']}] {a['title'][:40]}")
    else:
        print("  今天沒有集團新聞（信件沒指定「第N則」）")

    # ── 4. 產出 Word ───────────────────────────────────────
    step(4, total, "產出 Word 剪報")
    stamp = date_stamp(meta, excerpt)
    out = Path(args.out) if args.out else OUT_DIR / f"{stamp} 每日新聞剪報.docx"
    date_label = f"{stamp[:4]}/{int(stamp[4:6])}/{int(stamp[6:8])}每日新聞剪報"

    out_path, stats, sections = build_docx.build(
        template, articles, group_articles, out, date_label)

    # ── 摘要 ───────────────────────────────────────────────
    print("\n" + "=" * 62)
    print(f"完成：{out_path}")
    print("=" * 62)
    print(f"  成功    {stats['則數'] - len(failed)} 則"
          f"（{' / '.join(f'{l} {len(i)}' for l, i in sections)}）")
    print(f"  待補圖  {stats['圖片待補']} 則")
    print(f"  失敗    {len(failed)} 則")
    if warned:
        print(f"\n  有警告的 {len(warned)} 則（欄位可能要人工看一眼）:")
        for a in warned:
            print(f"    - {a.get('title', '')[:34]}")
            for w in a["warnings"]:
                print(f"        ! {w}")
    if failed:
        print(f"\n  失敗的 {len(failed)} 則（Word 裡沒有這幾則，要手動補）:")
        for a in failed:
            print(f"    - {a['url']}\n        {a['error']}")

    print("\n  接下來要人工做的:")
    print("    1. 用 Word 開啟，填封面目錄的頁碼（現在是 P.__~__）")
    if stats["圖片待補"]:
        print(f"    2. 搜尋紅字「圖片待補」，補上 {stats['圖片待補']} 張圖")
    print(f"    {3 if stats['圖片待補'] else 2}. 校對後自行寄出（本程式不會寄信）")
    print("=" * 62)


if __name__ == "__main__":
    main()
