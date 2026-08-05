# ─────────────────────────────────────────────────────────────
# 每日剪報 — 桌面工具的實際執行內容（Windows 版）
#
# 用法（平常不用自己下，雙擊桌面的「每日剪報」就好）:
#     .\run_clipping.ps1                  # 自動找「下載」裡最新的信件
#     .\run_clipping.ps1 某封信.eml        # 指定信件
#
# 這支腳本負責「讓非工程師也能跑」的雜事：
#   顧好虛擬環境 → 找到信件 → 呼叫 main.py → 開啟成品
# 真正的邏輯都在 files\*.py，這裡不做任何抽取或排版。
# macOS 的對應版本是 run_clipping.sh，兩邊要一起改。
# ─────────────────────────────────────────────────────────────
param(
    # 拖到桌面圖示上的檔案會變成這個參數
    [string]$Eml
)

$PROJECT = Split-Path -Parent $PSScriptRoot
$CODE    = Join-Path $PROJECT 'files'
$VENV    = Join-Path $PROJECT '.venv'
$PY      = Join-Path $VENV 'Scripts\python.exe'

# 主控台預設是 cp950，Python 印出的框線與部分字會變成問號。
# 這兩行讓中文輸出在任何 Windows 上都正常。
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$env:PYTHONIOENCODING = 'utf-8'

# 不管成功失敗都讓視窗停住，不然使用者看不到摘要就被關掉了
function Finish {
    param([int]$Code = 0)
    Write-Host ''
    Write-Host '────────────────────────────────────────────────────────'
    Write-Host '按 Enter 關閉視窗'
    Read-Host | Out-Null
    exit $Code
}

function Die {
    param([string]$Message, [string[]]$Hints = @())
    Write-Host ''
    Write-Host "X $Message" -ForegroundColor Red
    foreach ($h in $Hints) { Write-Host "  $h" }
    Finish 1
}

Write-Host '每日新聞剪報'
Write-Host "專案：$PROJECT"
Write-Host ''

# ── 1. 顧好執行環境 ──────────────────────────────────────────
# 第一次跑會自動建立，之後幾乎瞬間跳過
if (-not (Test-Path $PY)) {
    Write-Host '第一次執行，正在建立 Python 環境（約 1 分鐘，只有這次）…' -ForegroundColor Yellow
    # 這台機器上 python 是 Microsoft Store 的空殼，執行任何東西都不會有反應
    # 也不會報錯（見 README 一、1），所以一律用 py。
    if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
        Die '這台電腦沒有 Python（叫不到 py）' @(
            '到 https://www.python.org/downloads/ 裝一個，安裝時記得勾 py launcher，再重跑一次。')
    }
    py -m venv $VENV
    if ($LASTEXITCODE -ne 0) { Die '建立虛擬環境失敗' }
}

& $PY -c 'import requests, bs4, trafilatura, docx, PIL' 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host '正在安裝／補齊套件…' -ForegroundColor Yellow
    & $PY -m pip install --quiet --upgrade pip
    & $PY -m pip install --quiet -r (Join-Path $PROJECT 'requirements.txt')
    if ($LASTEXITCODE -ne 0) { Die '套件安裝失敗' @('檢查一下網路連線，再重跑一次。') }
    & $PY -c 'import requests, bs4, trafilatura, docx, PIL'
    if ($LASTEXITCODE -ne 0) { Die '套件裝完仍然叫不到' @('把上面的訊息整段複製給我看。') }
    Write-Host '環境就緒' -ForegroundColor Green
    Write-Host ''
}

# ── 2. 決定要處理哪封信 ──────────────────────────────────────
$downloads = Join-Path $env:USERPROFILE 'Downloads'

if (-not $Eml) {
    # 沒指定就抓「下載」裡最新的一封榮董新聞
    $newest = Get-ChildItem -Path $downloads -Filter '*榮董新聞*.eml' -File -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) {
        $Eml = $newest.FullName
        Write-Host '自動選用「下載」裡最新的信件'
    }
}

if (-not $Eml) {
    # 還是找不到就跳出選檔視窗
    Write-Host '「下載」裡沒有榮董新聞的 .eml，請手動選一個…'
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = '選擇今天的榮董新聞 .eml'
        $dlg.Filter = '信件檔 (*.eml)|*.eml|所有檔案 (*.*)|*.*'
        if (Test-Path $downloads) { $dlg.InitialDirectory = $downloads }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $Eml = $dlg.FileName
        }
    } catch {
        Write-Host "  選檔視窗開不起來：$($_.Exception.Message)"
    }
}

if (-not $Eml) {
    Die '沒有選擇信件，這次不做事' @(
        '把信件從 Outlook／Gmail 另存成 .eml 放到「下載」資料夾，再點一次。')
}
if (-not (Test-Path -LiteralPath $Eml -PathType Leaf)) {
    Die "找不到檔案：$Eml"
}
if ([IO.Path]::GetExtension($Eml).ToLower() -ne '.eml') {
    Die "這不是 .eml 檔：$([IO.Path]::GetFileName($Eml))" @('要的是信件檔，不是附件的 .docx。')
}

$Eml = (Resolve-Path -LiteralPath $Eml).Path
Write-Host "信件：$([IO.Path]::GetFileName($Eml))"
Write-Host ''

# ── 3. 跑主流程 ──────────────────────────────────────────────
# 工作目錄設成 files\，中間產物（urls.json、attachments\、outputs\）都落在那裡，
# .gitignore 已經把它們全部排除，不會不小心把信件內容 commit 上去。
if (-not (Test-Path $CODE)) { Die "進不去 $CODE" }
Push-Location $CODE
& $PY main.py $Eml
$status = $LASTEXITCODE
Pop-Location

if ($status -ne 0) {
    Die "執行失敗（結束代碼 $status）" @('上面紅色或最後幾行的訊息就是原因，整段複製給我看。')
}

# ── 4. 把成品端到使用者面前 ──────────────────────────────────
$out = Get-ChildItem -Path (Join-Path $CODE 'outputs') -Filter '*.docx' -File -ErrorAction SilentlyContinue |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($out) {
    Write-Host ''
    Write-Host "V 完成  $($out.Name)" -ForegroundColor Green
    Write-Host "  位置：$($out.DirectoryName)"

    # 同一天重跑時 Word 可能還開著上一版。檔案在硬碟上已經換新，但 Word 手上
    # 那份還是舊的，直接開只會把舊視窗叫到前面 —— 校對到舊版是很難發現的錯。
    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        # 倒著跑：關掉一份之後後面的索引會位移
        for ($i = $word.Documents.Count; $i -ge 1; $i--) {
            $d = $word.Documents.Item($i)
            if ($d.Name -eq $out.Name) { $d.Close(0) }   # 0 = 不存檔
        }
    } catch {
        # Word 沒開著時 GetActiveObject 會丟例外，那就沒事要做
    }

    explorer.exe "/select,$($out.FullName)"   # 在檔案總管裡選取該檔
    Invoke-Item -LiteralPath $out.FullName    # 順手用 Word 開起來校對
} else {
    Write-Host '! 程式回報成功，但 outputs\ 裡找不到檔案' -ForegroundColor Yellow
}

Finish 0
