#!/bin/bash
# ─────────────────────────────────────────────────────────────
# 產生桌面上的「每日剪報.app」
#
#     ./tools/build_app.sh
#
# 只有在「第一次安裝」或「把專案資料夾搬家」之後才需要跑。
# app 本身不含程式邏輯，只是去呼叫 tools/run_clipping.sh，
# 所以之後改程式不用重建 app。
# ─────────────────────────────────────────────────────────────
set -euo pipefail

PROJECT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RUNNER="$PROJECT/tools/run_clipping.sh"
APP="${1:-$HOME/Desktop/每日剪報.app}"
BUNDLE_ID="tw.com.o-bank.daily-clipping"   # 用途見下面補 CFBundleIdentifier 那段

chmod +x "$RUNNER"

SCRATCH="$(mktemp -d)"
trap 'rm -rf "$SCRATCH"' EXIT
SRC="$SCRATCH/droplet.applescript"

# 用 Terminal 跑而不是背景執行，是刻意的：
# 這支程式要跑 2–3 分鐘，而且最後的摘要（待補圖幾張、哪幾則失敗）
# 是使用者一定要看到的資訊，藏起來等於白做。
cat > "$SRC" <<APPLESCRIPT
-- 每日剪報：把榮董新聞信件變成 Word 剪報草稿
-- 雙擊 = 自動抓「下載」資料夾裡最新的一封；把 .eml 拖上來 = 處理指定那封

on run
	launchWith("")
end run

on open theFiles
	launchWith(POSIX path of (item 1 of theFiles))
end open

on launchWith(emlPath)
	set runner to quoted form of "$RUNNER"
	if emlPath is "" then
		set cmd to "clear; " & runner
	else
		set cmd to "clear; " & runner & " " & quoted form of emlPath
	end if
	tell application "Terminal"
		activate
		do script cmd
	end tell
end launchWith
APPLESCRIPT

rm -rf "$APP"
osacompile -o "$APP" "$SRC"

# osacompile 不一定會宣告「這個 app 收得下拖進來的檔案」，補上去，
# 不然把 .eml 拖到圖示上會被系統擋掉。
PLIST="$APP/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Delete :CFBundleDocumentTypes" "$PLIST" 2>/dev/null || true
/usr/libexec/PlistBuddy \
    -c "Add :CFBundleDocumentTypes array" \
    -c "Add :CFBundleDocumentTypes:0 dict" \
    -c "Add :CFBundleDocumentTypes:0:CFBundleTypeName string '信件檔'" \
    -c "Add :CFBundleDocumentTypes:0:CFBundleTypeRole string Viewer" \
    -c "Add :CFBundleDocumentTypes:0:LSItemContentTypes array" \
    -c "Add :CFBundleDocumentTypes:0:LSItemContentTypes:0 string public.data" \
    -c "Add :CFBundleDocumentTypes:0:LSItemContentTypes:1 string public.item" \
    "$PLIST" >/dev/null

# osacompile 產生的 applet 沒有 CFBundleIdentifier，ad-hoc 簽章的識別碼會變成
# 中文的 app 名稱。macOS 的權限資料庫（TCC）是用 bundle identifier 認 app 的，
# 沒有 ID 就建不出授權紀錄 —— 症狀是「未獲授權來傳送 Apple Event 到 Terminal
# (-1743)」一直跳，而且「系統設定 → 隱私權與安全性 → 自動化」裡根本找不到這支
# 可以打開。補上固定的 ID 再重簽一次，授權才記得住。
/usr/libexec/PlistBuddy -c "Add :CFBundleIdentifier string $BUNDLE_ID" "$PLIST" \
    2>/dev/null \
    || /usr/libexec/PlistBuddy -c "Set :CFBundleIdentifier $BUNDLE_ID" "$PLIST"

# 改完 Info.plist 一定要重簽：簽章跟 ID 對不起來的話 TCC 會當成另一支 app，
# 之前給過的授權就失效了。
codesign --force --sign - --identifier "$BUNDLE_ID" "$APP"

touch "$APP"   # 逼 Finder 重讀 bundle 資訊

echo "✓ 已建立：$APP"
echo "  指向：$RUNNER"
