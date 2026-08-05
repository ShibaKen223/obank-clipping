# ─────────────────────────────────────────────────────────────
# 產生桌面上的「每日剪報」捷徑（Windows 版）
#
#     powershell -ExecutionPolicy Bypass -File .\tools\build_app.ps1
#
# 只有在「第一次安裝」或「把專案資料夾搬家」之後才需要跑。
# 捷徑本身不含程式邏輯，只是去呼叫 tools\run_clipping.ps1，
# 所以之後改程式不用重建捷徑。
# macOS 的對應版本是 build_app.sh，兩邊要一起改。
# ─────────────────────────────────────────────────────────────
param(
    # 想放到別的地方就自己指定，預設是桌面
    [string]$Shortcut = (Join-Path ([Environment]::GetFolderPath('Desktop')) '每日剪報.lnk')
)

$ErrorActionPreference = 'Stop'

$PROJECT = Split-Path -Parent $PSScriptRoot
$RUNNER  = Join-Path $PROJECT 'tools\run_clipping.ps1'

if (-not (Test-Path $RUNNER)) {
    Write-Host "X 找不到 $RUNNER" -ForegroundColor Red
    exit 1
}

# 用 powershell.exe 開一個看得見的視窗跑，而不是背景執行，是刻意的：
# 這支程式要跑 2–3 分鐘，而且最後的摘要（待補圖幾張、哪幾則失敗）
# 是使用者一定要看到的資訊，藏起來等於白做。
#
# -ExecutionPolicy Bypass 是因為預設政策（RemoteSigned）會擋掉沒簽章的
# .ps1；只對這一次執行放行，不會改到系統設定。
$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$args  = "-NoProfile -ExecutionPolicy Bypass -File `"$RUNNER`""

# 圖示借 Word 的，一眼就知道這東西產出的是什麼；沒裝 Word 就用系統的文件圖示
$icon = "$env:SystemRoot\System32\shell32.dll,70"
$appPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
if (Test-Path $appPaths) {
    $winword = (Get-ItemProperty $appPaths).'(default)'
    if ($winword -and (Test-Path $winword)) { $icon = "$winword,0" }
}

if (Test-Path $Shortcut) { Remove-Item $Shortcut -Force }

$shell = New-Object -ComObject WScript.Shell
$lnk = $shell.CreateShortcut($Shortcut)
$lnk.TargetPath       = $psExe
$lnk.Arguments        = $args
$lnk.WorkingDirectory = $PROJECT
$lnk.IconLocation     = $icon
$lnk.Description      = '把榮董新聞信件變成 Word 剪報草稿'
$lnk.Save()

Write-Host "V 已建立：$Shortcut" -ForegroundColor Green
Write-Host "  指向：$RUNNER"
Write-Host ''
Write-Host '用法：'
Write-Host '  雙擊         → 自動處理「下載」資料夾裡最新的一封榮董新聞'
Write-Host '  拖 .eml 上去 → 處理指定的那一封'
