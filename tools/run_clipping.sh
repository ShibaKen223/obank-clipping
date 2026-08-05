#!/bin/bash
# ─────────────────────────────────────────────────────────────
# 每日剪報 — 桌面工具的實際執行內容
#
# 用法（平常不用自己下，雙擊桌面的「每日剪報」就好）:
#     ./run_clipping.sh                  # 自動找 Downloads 裡最新的信件
#     ./run_clipping.sh 某封信.eml        # 指定信件
#
# 這支腳本負責「讓非工程師也能跑」的雜事：
#   顧好虛擬環境 → 找到信件 → 呼叫 main.py → 開啟成品
# 真正的邏輯都在 files/*.py，這裡不做任何抽取或排版。
# ─────────────────────────────────────────────────────────────
set -uo pipefail

PROJECT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CODE="$PROJECT/files"
VENV="$PROJECT/.venv"
PY="$VENV/bin/python"

RED=$'\033[31m'; GREEN=$'\033[32m'; YELLOW=$'\033[33m'; BOLD=$'\033[1m'; OFF=$'\033[0m'

# 不管成功失敗都讓視窗停住，不然使用者看不到摘要就被關掉了
finish() {
    echo
    echo "────────────────────────────────────────────────────────"
    echo "按 Enter 關閉視窗"
    read -r _
    exit "${1:-0}"
}

die() {
    echo
    echo "${RED}${BOLD}✗ $1${OFF}"
    shift
    for line in "$@"; do echo "  $line"; done
    finish 1
}

echo "${BOLD}每日新聞剪報${OFF}"
echo "專案：$PROJECT"
echo

# ── 1. 顧好執行環境 ──────────────────────────────────────────
# 第一次跑會自動建立，之後幾乎瞬間跳過
if [ ! -x "$PY" ]; then
    echo "${YELLOW}第一次執行，正在建立 Python 環境（約 1 分鐘，只有這次）…${OFF}"
    command -v python3 >/dev/null 2>&1 || die "這台電腦沒有 python3" \
        "到 https://www.python.org/downloads/ 裝一個，再重跑一次。"
    python3 -m venv "$VENV" || die "建立虛擬環境失敗"
fi

if ! "$PY" -c "import requests, bs4, trafilatura, docx, PIL" 2>/dev/null; then
    echo "${YELLOW}正在安裝／補齊套件…${OFF}"
    "$PY" -m pip install --quiet --upgrade pip
    "$PY" -m pip install --quiet -r "$PROJECT/requirements.txt" \
        || die "套件安裝失敗" "檢查一下網路連線，再重跑一次。"
    "$PY" -c "import requests, bs4, trafilatura, docx, PIL" \
        || die "套件裝完仍然叫不到" "把上面的訊息整段複製給我看。"
    echo "${GREEN}環境就緒${OFF}"
    echo
fi

# ── 2. 決定要處理哪封信 ──────────────────────────────────────
EML="${1:-}"

if [ -z "$EML" ]; then
    # 沒指定就抓 Downloads 裡最新的一封榮董新聞
    EML="$(ls -t "$HOME/Downloads/"*榮董新聞*.eml 2>/dev/null | head -1)"
    [ -n "$EML" ] && echo "自動選用 Downloads 裡最新的信件"
fi

if [ -z "$EML" ]; then
    # 還是找不到就跳出選檔視窗
    echo "Downloads 裡沒有榮董新聞的 .eml，請手動選一個…"
    EML="$(osascript -e 'POSIX path of (choose file with prompt "選擇今天的榮董新聞 .eml" of type {"eml", "public.data"})' 2>/dev/null)"
fi

[ -n "$EML" ] || die "沒有選擇信件，這次不做事" \
    "把信件從 Outlook／Gmail 另存成 .eml 放到「下載」資料夾，再點一次。"
[ -f "$EML" ] || die "找不到檔案：$EML"

case "$EML" in
    *.eml) ;;
    *) die "這不是 .eml 檔：$(basename "$EML")" \
           "要的是信件檔，不是附件的 .docx。" ;;
esac

echo "信件：$(basename "$EML")"
echo

# ── 3. 跑主流程 ──────────────────────────────────────────────
# CWD 設成 files/，中間產物（urls.json、attachments/、outputs/）都落在那裡，
# .gitignore 已經把它們全部排除，不會不小心把信件內容 commit 上去。
cd "$CODE" || die "進不去 $CODE"

"$PY" main.py "$EML"
STATUS=$?

if [ $STATUS -ne 0 ]; then
    die "執行失敗（結束代碼 $STATUS）" \
        "上面紅色或最後幾行的訊息就是原因，整段複製給我看。"
fi

# ── 4. 把成品端到使用者面前 ──────────────────────────────────
NEWEST="$(ls -t "$CODE/outputs/"*.docx 2>/dev/null | head -1)"
if [ -n "$NEWEST" ]; then
    echo
    echo "${GREEN}${BOLD}✓ 完成${OFF}  $(basename "$NEWEST")"
    echo "  位置：$CODE/outputs/"
    # 同一天重跑時 Word 可能還開著上一版。檔案在硬碟上已經換新，但 Word 手上
    # 那份還是舊的，直接 open 只會把舊視窗叫到前面 —— 校對到舊版是很難發現的錯。
    osascript -e "tell application \"Microsoft Word\" to close (every document whose name is \"$(basename "$NEWEST")\") saving no" >/dev/null 2>&1

    open -R "$NEWEST"      # 在 Finder 裡選取該檔
    open "$NEWEST"         # 順手用 Word 開起來校對
else
    echo "${YELLOW}! 程式回報成功，但 outputs/ 裡找不到檔案${OFF}"
fi

finish 0
