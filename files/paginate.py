"""
paginate.py — 把封面目錄的 P.__~__ 換成真的頁碼

原本 SPEC §4.3 說頁碼算不出來、留給人工填。算不出來是因為要知道每一則佔幾頁，
就得真的排版一次 —— 而排版是 Word 的工作，python-docx 只碰得到 XML。

作法就是讓 Word 自己去排，再數出每個類別從第幾頁開始。每一則都是分頁起頭，
所以類別邊界很乾淨。兩個平台各有一套量法：

    Windows  COM（pywin32）直接問 Word 每個段落在第幾頁 —— 最準，也不必轉檔
    macOS    AppleScript 叫 Word 匯出 PDF，再從 PDF 文字裡找標題落在哪一頁

Windows 走 COM 而不是也匯 PDF，是因為 COM 問得到的就是 Word 自己的頁碼，
不必再用文字比對去猜；Mac 沒有等價的 AppleScript 介面（Word for Mac 的
字典裡沒有「這段在第幾頁」），只好繞 PDF。

沒有 Word、或兩套都跑不動時就跳過，封面維持 P.__~__ 讓人工填，
不會讓整個流程失敗。
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

from build_docx import set_paragraph_text, toc_line

# Word 排一份 55 頁的 A3 大約 20–40 秒，給寬一點免得慢的機器被砍掉
WORD_TIMEOUT = 300

# 拿標題的前幾個字去比對就夠，取太長反而容易被換行、全半形差異弄壞
NEEDLE_LEN = 20

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
    """去掉所有空白再比對。PDF 抽出來的文字會在標題中間夾換行與空格，
    Word COM 讀到的段落尾端也會多一個 \\r。"""
    return re.sub(r"\s+", "", s or "")


def _needles(sections: list):
    """每個類別的第一則標題，當成「這一類從這裡開始」的定位點。"""
    out = []
    for label, items in sections:
        if not items:
            return None
        out.append(_squash(items[0].get("title", ""))[:NEEDLE_LEN])
    return out


# ─────────────────────────────────────────────────────────────
# Windows：COM 直接問 Word
# ─────────────────────────────────────────────────────────────

_WD_STATISTIC_PAGES = 2        # ComputeStatistics 的頁數
_WD_ACTIVE_END_PAGE = 3        # Range.Information 的「這裡是第幾頁」


def measure_windows(docx_path: Path, sections: list):
    """回傳 (每類起始頁, 總頁數)，頁碼是 Word 的頁碼（封面算第 1 頁）。"""
    try:
        import win32com.client as win32
    except ImportError:
        print("  頁碼：缺 pywin32（py -m pip install pywin32），跳過")
        return None

    needles = _needles(sections)
    if not needles:
        return None

    word = None
    doc = None
    try:
        # DispatchEx 會另外開一個 Word 行程，不去動使用者手上正開著的視窗
        # （校對到一半被程式關掉是很惱人的事）。
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(docx_path.resolve()),
                                  ReadOnly=True, AddToRecentFiles=False,
                                  Visible=False)
        doc.Repaginate()

        starts = []
        cursor = 0                      # 下一個要找的類別
        for para in doc.Paragraphs:
            if cursor >= len(needles):
                break
            text = _squash(para.Range.Text)
            if not text:
                continue
            if needles[cursor] in text:
                starts.append(int(para.Range.Information(_WD_ACTIVE_END_PAGE)))
                cursor += 1
        total = int(doc.ComputeStatistics(_WD_STATISTIC_PAGES))
    except Exception as e:
        print(f"  頁碼：Word 量測失敗，跳過（{type(e).__name__}: {e}）")
        return None
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass

    if len(starts) < len(needles):
        label = sections[len(starts)][0]
        print(f"  頁碼：在排版結果裡找不到「{label}」的第一則，跳過")
        return None
    return starts, total


# ─────────────────────────────────────────────────────────────
# macOS：AppleScript 叫 Word 匯 PDF，再讀 PDF
# ─────────────────────────────────────────────────────────────

def export_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """請 Word 把 docx 匯出成 PDF。失敗回 False，不丟例外。"""
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


def measure_macos(docx_path: Path, sections: list):
    """回傳 (每類起始頁, 總頁數)。頁碼從 1 起算，跟 Windows 那條路一致。"""
    try:
        import pypdfium2
    except ImportError:
        print("  頁碼：缺 pypdfium2（pip install pypdfium2），跳過")
        return None

    needles = _needles(sections)
    if not needles:
        return None

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / "measure.pdf"
        if not export_pdf(docx_path, pdf_path):
            return None
        pdf = pypdfium2.PdfDocument(str(pdf_path))
        try:
            pages = [_squash(pdf[i].get_textpage().get_text_range())
                     for i in range(len(pdf))]
        finally:
            # 一定要關。macOS 刪得掉開著的檔案，所以不關也看不出問題，
            # 但那是漏掉的 handle，而且在別的平台上 TemporaryDirectory
            # 收尾時會直接丟 PermissionError。
            pdf.close()

    starts = []
    cursor = 0
    for label, needle in zip([l for l, _ in sections], needles):
        hit = next((i for i in range(cursor, len(pages)) if needle in pages[i]),
                   None)
        if hit is None:
            print(f"  頁碼：在 PDF 裡找不到「{label}」的第一則，跳過")
            return None
        starts.append(hit + 1)          # 轉成 1 起算
        cursor = hit
    return starts, len(pages)


# ─────────────────────────────────────────────────────────────
# 共用
# ─────────────────────────────────────────────────────────────

def to_ranges(sections: list, starts: list, total: int) -> list:
    """把每類的起始頁換算成 [(類別, 起頁, 迄頁)]。

    讀者看到的第 1 頁 = 第一則新聞那一頁（封面不編號，跟人工版一致），
    所以要把封面那幾頁的偏移扣掉。
    """
    offset = starts[0] - 1
    ranges = []
    for i, (label, _) in enumerate(sections):
        first = starts[i] - offset
        last = (starts[i + 1] - 1 if i + 1 < len(starts) else total) - offset
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
        # 用 build_docx 的同一個函式排版，空白數才會跟產出當下一致
        set_paragraph_text(p, toc_line(label, f"{first}~{last}"))
        changed += 1
    if changed:
        doc.save(str(docx_path))
    return changed


def fill_page_numbers(docx_path, sections) -> bool:
    """主流程。任何一步失敗都只是保持 P.__~__，不會讓產出失敗。"""
    docx_path = Path(docx_path)
    system = platform.system()
    if system == "Windows":
        measure = measure_windows
    elif system == "Darwin":
        measure = measure_macos
    else:
        print(f"  頁碼：{system} 沒有可用的 Word，跳過")
        return False

    with tempfile.TemporaryDirectory() as tmp:
        # 先複製成一個獨一無二的檔名再交給 Word。
        # 直接餵原檔的話，只要 Word 裡正好開著同名文件（上一次執行結束時會自動
        # 開起來校對，同一天重跑就會撞到），Word 會拿記憶體裡那份去量，
        # 量到的是舊內容而且完全不會報錯 —— 實測踩過這個坑。
        work = Path(tmp) / f"measure_{os.getpid()}.docx"
        shutil.copy2(docx_path, work)
        measured = measure(work, sections)
        if not measured:
            return False
        starts, total = measured

    ranges = to_ranges(sections, starts, total)
    changed = write_into_cover(docx_path, ranges)
    if not changed:
        print("  頁碼：封面找不到目錄行，跳過")
        return False
    for label, first, last in ranges:
        print(f"    {label} P.{first}~{last}")
    return True


if __name__ == "__main__":
    print("這支模組由 main.py 呼叫，沒有單獨的指令列用法。", file=sys.stderr)
