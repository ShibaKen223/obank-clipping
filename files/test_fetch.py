"""
test_fetch.py — 反爬測試：確認 27 個網址能不能抓

這支只做「診斷」，不產生任何 Word 檔。
目的是在寫正式程式之前，先知道哪些網域會擋。

安裝:
    pip install requests beautifulsoup4 trafilatura

用法:
    python parse_email.py 0804_榮董新聞.eml --json urls.json
    python test_fetch.py urls.json

看什麼:
    OK   = 內文抽得到，可以全自動
    IMG? = 內文可以，但主圖抓不到 → 該則圖片手動補
    FAIL = 擋住了 → 這個網域要改用 Claude in Chrome 半自動
"""
import json
import sys
import time
from urllib.parse import urlparse

import requests
import trafilatura
from bs4 import BeautifulSoup

UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Accept-Language": "zh-TW,zh;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
TIMEOUT = 20


def probe(url: str) -> dict:
    out = {"url": url, "status": None, "chars": 0, "img": None,
           "img_ok": None, "verdict": "FAIL", "note": ""}
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        out["status"] = r.status_code
        if r.status_code != 200:
            out["note"] = f"HTTP {r.status_code}"
            return out

        # 內文抽取
        text = trafilatura.extract(r.text, include_comments=False,
                                   include_tables=False) or ""
        out["chars"] = len(text)

        # 主圖：og:image
        soup = BeautifulSoup(r.text, "html.parser")
        og = soup.find("meta", property="og:image")
        if og and og.get("content"):
            out["img"] = og["content"]
            # 實測防盜連：一定要帶 Referer
            try:
                ir = requests.get(out["img"], headers={**HEADERS, "Referer": url},
                                  timeout=TIMEOUT, stream=True)
                out["img_ok"] = (ir.status_code == 200
                                 and ir.headers.get("Content-Type", "").startswith("image"))
                if not out["img_ok"]:
                    out["note"] = f"圖片 HTTP {ir.status_code}"
            except Exception as e:
                out["img_ok"] = False
                out["note"] = f"圖片錯誤 {type(e).__name__}"

        # 判定：300 字以下多半是被擋或只吐出摘要
        if out["chars"] >= 300:
            out["verdict"] = "OK" if out["img_ok"] else "IMG?"
        else:
            out["note"] = out["note"] or f"只抽到 {out['chars']} 字，疑似被擋或付費牆"

    except Exception as e:
        out["note"] = f"{type(e).__name__}: {e}"
    return out


def main(path: str):
    items = json.load(open(path, encoding="utf-8"))["items"]
    results = []
    for i, it in enumerate(items, 1):
        r = probe(it["url"])
        r["title"] = it["title"]
        results.append(r)
        print(f"[{i:2d}/{len(items)}] {r['verdict']:4s} "
              f"{urlparse(r['url']).netloc:20s} {r['chars']:5d}字  "
              f"{it['title'][:24]}  {r['note']}")
        time.sleep(1.5)  # 有禮貌一點，別把對方伺服器打爆

    # 依網域彙總 —— 這是你真正要看的表
    print("\n" + "=" * 62)
    print("依網域彙總（決定哪些要走半自動）")
    print("=" * 62)
    by_domain: dict[str, list] = {}
    for r in results:
        by_domain.setdefault(urlparse(r["url"]).netloc, []).append(r)

    for dom, rs in sorted(by_domain.items(), key=lambda x: -len(x[1])):
        ok = sum(r["verdict"] == "OK" for r in rs)
        imgq = sum(r["verdict"] == "IMG?" for r in rs)
        fail = sum(r["verdict"] == "FAIL" for r in rs)
        avg = sum(r["chars"] for r in rs) / len(rs)
        flag = "✅ 可全自動" if fail == 0 else ("⚠️ 部分失敗" if ok + imgq else "❌ 要走 Chrome")
        print(f"{dom:22s} 共{len(rs):2d}則  OK {ok:2d} / 缺圖 {imgq:2d} / 失敗 {fail:2d}"
              f"  均{avg:.0f}字  {flag}")

    total_ok = sum(r["verdict"] in ("OK", "IMG?") for r in results)
    print(f"\n可自動處理：{total_ok}/{len(results)} 則 "
          f"({total_ok / len(results) * 100:.0f}%)")

    json.dump(results, open("test_result.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print("明細已存 test_result.json")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "urls.json")
