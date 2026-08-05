"""
paginate.py — 把封面目錄的 P.__~__ 換成真的頁碼

原本 SPEC §4.3 說頁碼算不出來、留給人工填。算不出來是因為要知道每一則佔幾頁，
就得真的排版一次 —— 而排版是 Word 的工作，python-docx 只碰得到 XML。

作法就是讓 Word 自己去排：把成品丟給 Word 匯出成 PDF，數出每個類別從第幾頁
開始，再把數字寫回封面。每一則都是分頁起頭，所以類別邊界很乾淨。

限制：要有 Microsoft Word，而且目前只支援 macOS（靠 AppleScript 驅動）。
沒有 Word 就跳過，封面維持 P.__~__ 讓人工填，不會讓整個流程失敗。
"""
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document

from build_docx import set_paragraph_text

# Word 排一份 55 頁的 A3 大約 20–40 秒，給寬一點免得慢的機器被砍掉
WORD_TIMEOUT = 300

_EXPORT_SCRIPT = """
on run argv
    set src to item 1 of argv
    set dst to item 2 of argv
    set srcName to do shell script "basename " & quoted form of src
    set dstName to do shell script "basename " & quoted form of dst
    -- AppleEvent 預設等 2 分鐘就放棄，但 Word 排一份 55 頁的 A3 會更久。
    tell application "Microsoft Word"
      with timeout of 600 seconds
        set theDoc to open file name src
        save as theDoc file name dst file format format PDF
        -- save as 之後這份文件的名字已經變成 PDF 的檔名，原本的 theDoc 參照
        -- 就失效了（close theDoc 會直接報錯），所以收尾改用檔名逐一比對。
        -- 倒著跑：關掉一份之後後面的索引會位移。
        -- 外層 try 是因為 Word 一份文件都沒開時，count of documents 會丟錯。
        try
            repeat with i from (count of documents) to 1 by -1
                try
                    set nm to name of document i
                    if nm is srcName or nm is dstName then
                        close document i saving no
                    end if
                end try
            end repeat
        end try
      end timeout
    end tell
    return "OK"
end run
"""


def _squash(s: str) -> str:
    """去掉所有空白再比對。PDF 抽出來的文字會在標題中間夾換行與空格。"""
    return re.sub(r"\s+", "", s or "")


def export_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """請 Word 把 docx 匯出成 PDF。失敗回 False，不丟例外。"""
    if platform.system() != "Darwin":
        print("  頁碼：只有 macOS 支援自動填（需要 AppleScript），跳過")
        return False
    if not Path("/Applications/Microsoft Word.app").exists():
        print("  頁碼：這台沒有 Microsoft Word，跳過")
        return False
    try:
        r = subprocess.run(
            ["osascript", "-", str(docx_path.resolve()), str(pdf_path.resolve())],
            input=_EXPORT_SCRIPT, capture_output=True, text=True,
            timeout=WORD_TIMEOUT)
    except subprocess.TimeoutExpired:
        print(f"  頁碼：Word 超過 {WORD_TIMEOUT} 秒沒回應，跳過")
        return False
    if r.returncode != 0 or not pdf_path.exists():
        msg = (r.stderr or "").strip().splitlines()
        print(f"  頁碼：Word 匯出失敗，跳過（{msg[-1] if msg else '沒有錯誤訊息'}）")
        return False
    return True


def page_of_each_section(pdf_path: Path, sections: list):
    """回傳 [(類別, 起頁, 迄頁)]，頁碼是讀者看到的頁碼（封面不算）。

    每個類別的起頁 = 該類第一則標題出現的那一頁。
    迄頁 = 下一類起頁的前一頁；最後一類就到全文結束。
    """
    try:
        import pypdfium2
    except ImportError:
        print("  頁碼：缺 pypdfium2（pip install pypdfium2），跳過")
        return None

    pdf = pypdfium2.PdfDocument(str(pdf_path))
    pages = [_squash(pdf[i].get_textpage().get_text_range())
             for i in range(len(pdf))]

    starts = []
    cursor = 0
    for label, items in sections:
        if not items:
            return None
        needle = _squash(items[0].get("title", ""))[:20]
        hit = next((i for i in range(cursor, len(pages)) if needle in pages[i]),
                   None)
        if hit is None:
            print(f"  頁碼：在 PDF 裡找不到「{label}」的第一則，跳過")
            return None
        starts.append(hit)
        cursor = hit

    # 讀者看到的第 1 頁 = 第一則新聞那一頁（封面不編號，跟人工版一致）
    offset = starts[0] - 1
    ranges = []
    for i, (label, _) in enumerate(sections):
        first = starts[i] - offset
        last = (starts[i + 1] - 1 if i + 1 < len(starts) else len(pages) - 1) - offset
        ranges.append((label, first, max(first, last)))
    return ranges


def write_into_cover(docx_path: Path, ranges: list) -> int:
    """把封面目錄的 P.__~__ 換成算出來的頁碼。回傳換掉幾行。"""
    doc = Document(str(docx_path))
    todo = list(ranges)
    changed = 0
    for p in doc.paragraphs:
        text = " ".join(p.text.split())
        if not todo or not re.search(r"P\.\s*(__|\d+)\s*[~～]", text):
            continue
        label, first, last = todo.pop(0)
        set_paragraph_text(p, f"{label} P.{first}~{last}")
        changed += 1
    if changed:
        doc.save(str(docx_path))
    return changed


def fill_page_numbers(docx_path, sections) -> bool:
    """主流程。任何一步失敗都只是保持 P.__~__，不會讓產出失敗。"""
    docx_path = Path(docx_path)
    with tempfile.TemporaryDirectory() as tmp:
        # 先複製成一個獨一無二的檔名再交給 Word。
        # 直接餵原檔的話，只要 Word 裡正好開著同名文件（上一次執行結束時會自動
        # 開起來校對，同一天重跑就會撞到），Word 會拿記憶體裡那份去匯出，
        # 量到的是舊內容而且完全不會報錯 —— 實測踩過這個坑。
        work = Path(tmp) / f"measure_{os.getpid()}.docx"
        shutil.copy2(docx_path, work)
        pdf_path = Path(tmp) / "measure.pdf"
        if not export_pdf(work, pdf_path):
            return False
        ranges = page_of_each_section(pdf_path, sections)
        if not ranges:
            return False

    changed = write_into_cover(docx_path, ranges)
    if not changed:
        print("  頁碼：封面找不到目錄行，跳過")
        return False
    for label, first, last in ranges:
        print(f"    {label} P.{first}~{last}")
    return True


if __name__ == "__main__":
    print("這支模組由 main.py 呼叫，沒有單獨的指令列用法。", file=sys.stderr)
