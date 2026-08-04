"""
parse_email.py — 解析 Duke 每天寄來的榮董新聞信

用法:
    python parse_email.py 0804_榮董新聞.eml
    python parse_email.py 0804_榮董新聞.eml --json out.json

輸出:
    1. 一份 JSON: [{category, title, url}, ...]
    2. 附件自動存到 ./attachments/
    3. 集團新聞的「第N則」指示會被抓出來放在 meta.group_news_picks
"""
import email
import email.policy
import json
import re
import sys
from pathlib import Path

# Outlook 把編號清單轉成純文字時，每一項都會變成 "  1.  類別名"
CATEGORY_RE = re.compile(r"^\s+\d+\.\s+(\S.*?)\s*$")
URL_RE = re.compile(r"^https?://\S+$")
# 「第1,3則」「第 1、3 則」等寫法
PICK_RE = re.compile(r"第\s*([\d\s,，、和及]+?)\s*則")

# 信尾雜訊，掃到就停
STOP_MARKERS = ("Best Regards", "-----------", "________")


def parse(eml_path: Path) -> dict:
    with open(eml_path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=email.policy.default)

    body = None
    for part in msg.walk():
        if part.get_content_type() == "text/plain":
            body = part.get_content()
            break
    if body is None:
        raise SystemExit("找不到純文字內文，改用 HTML 版本解析")

    items, picks = [], []
    category = None
    pending_title = None

    for raw in body.splitlines():
        line = raw.strip()
        if any(m in line for m in STOP_MARKERS):
            break
        if not line:
            continue

        m = CATEGORY_RE.match(raw)
        if m:
            category = m.group(1)
            pending_title = None
            continue

        if URL_RE.match(line):
            if pending_title:
                items.append(
                    {"category": category, "title": pending_title, "url": line}
                )
                pending_title = None
            continue

        # 集團新聞那段沒有網址，只有「第1,3則」
        if category and "集團" in category:
            p = PICK_RE.search(line)
            if p:
                picks = [int(n) for n in re.findall(r"\d+", p.group(1))]
                continue

        pending_title = line

    return {
        "meta": {
            "subject": msg.get("Subject"),
            "date": msg.get("Date"),
            "from": msg.get("From"),
            "group_news_picks": picks,
            "total_urls": len(items),
        },
        "items": items,
    }


def save_attachments(eml_path: Path, outdir: Path) -> list[str]:
    outdir.mkdir(parents=True, exist_ok=True)
    saved = []
    with open(eml_path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=email.policy.default)
    for part in msg.walk():
        fn = part.get_filename()
        if not fn:
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        dest = outdir / fn
        dest.write_bytes(payload)
        saved.append(str(dest))
    return saved


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    path = Path(sys.argv[1])
    result = parse(path)
    result["meta"]["attachments"] = save_attachments(path, Path("attachments"))

    if "--json" in sys.argv:
        out = Path(sys.argv[sys.argv.index("--json") + 1])
        out.write_text(json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
        print(f"已寫出 {out}")

    m = result["meta"]
    print(f"\n主旨: {m['subject']}")
    print(f"集團新聞挑選: 第 {m['group_news_picks']} 則")
    print(f"待抓網址: {m['total_urls']} 則")
    print(f"附件: {len(m['attachments'])} 個\n")

    last = None
    for it in result["items"]:
        if it["category"] != last:
            print(f"\n【{it['category']}】")
            last = it["category"]
        print(f"  - {it['title'][:40]}")
        print(f"    {it['url']}")
