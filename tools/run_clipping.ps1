# ─────────────────────────────────────────────────────────────
# 每日剪報 — 桌面工具的實際執行內容（Windows 版）
#
# 用法（平常不用自己下，雙擊桌面的「每日剪報」就好）:
#     .\run_clipping.ps1                  # 自動找「下載」裡最新的信件
#     .\run_clipping.ps1 某封信.eml        # 指定信件
#     .\run_clipping.ps1 某份剪報.docx     # 只重算封面頁碼（校對時改過圖片大小之後）
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
# 不要寫死 %USERPROFILE%\Downloads。公司／學校的 OneDrive 常把「桌面」「文件」
# 「下載」整組搬進 OneDrive 資料夾，寫死的話會變成「檔案明明就在下載裡，
# 程式卻說找不到」—— 這種錯很難自己看出原因。登錄檔記的才是真正的位置。
function Get-DownloadsPath {
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $guid = '{374DE290-123F-4565-9164-39C4925E467B}'   # KNOWNFOLDERID: Downloads
    try {
        $raw = (Get-ItemProperty -Path $key -Name $guid -ErrorAction Stop).$guid
        if ($raw) {
            # 值可能長成 %USERPROFILE%\Downloads，要自己展開環境變數
            $path = [Environment]::ExpandEnvironmentVariables($raw)
            if (Test-Path $path) { return $path }
        }
    } catch {}
    return (Join-Path $env:USERPROFILE 'Downloads')
}

$downloads = Get-DownloadsPath

# 拖進來的是剪報 .docx，代表「我在 Word 裡改過了，幫我重算頁碼」。
# 調圖片大小會讓後面的內容整個位移，產出當下量的頁碼就不準了，而且不會有提示。
if ($Eml -and [IO.Path]::GetExtension($Eml).ToLower() -eq '.docx') {
    if (-not (Test-Path -LiteralPath $Eml -PathType Leaf)) { Die "找不到檔案：$Eml" }
    $doc = (Resolve-Path -LiteralPath $Eml).Path
    Write-Host "重算封面頁碼  $([IO.Path]::GetFileName($doc))"
    Write-Host ''

    # Word 開著這個檔就會鎖成獨佔，寫不回去，先關掉那一份
    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        for ($i = $word.Documents.Count; $i -ge 1; $i--) {
            $d = $word.Documents.Item($i)
            if ($d.FullName -eq $doc) {
                Write-Host "  先關掉 Word 裡開著的「$($d.Name)」"
                $d.Close(0)
            }
        }
    } catch {}

    Push-Location $CODE
    & $PY 'paginate.py' $doc
    $code = $LASTEXITCODE
    Pop-Location
    if ($code -ne 0) { Die '重算失敗' @('上面的訊息就是原因。') }

    Write-Host ''
    Write-Host 'V 頁碼已更新' -ForegroundColor Green
    Invoke-Item -LiteralPath $doc
    Finish 0
}

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
    Die "認不得這個檔：$([IO.Path]::GetFileName($Eml))" @(
        '拖信件的 .eml 進來 = 產出剪報；',
        '拖剪報的 .docx 進來 = 重算封面頁碼。')
}

$Eml = (Resolve-Path -LiteralPath $Eml).Path
Write-Host "信件：$([IO.Path]::GetFileName($Eml))"
Write-Host ''

# ── 3. 跑主流程 ──────────────────────────────────────────────
$outDir = Join-Path $CODE 'outputs'

# 同一天重跑時，上一版通常還開在 Word 裡校對。Windows 的 Word 會把開著的檔案
# 鎖成獨佔，python-docx 存檔會直接 PermissionError 爆掉 —— 所以要在跑之前關掉，
# 不是跑完才關（Mac 版是跑完才關，因為 Word for Mac 不會鎖住檔案，
# 那邊要解決的只有「開起來是舊版」的問題）。
# 順帶也解決了舊版殘留：跑完再開一定是新的那份。
function Close-OutputDocs {
    param([string]$Folder)
    if (-not (Test-Path $Folder)) { return }
    $target = (Resolve-Path $Folder).Path.TrimEnd('\')
    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        # 倒著跑：關掉一份之後後面的索引會位移
        for ($i = $word.Documents.Count; $i -ge 1; $i--) {
            $d = $word.Documents.Item($i)
            if ($d.Path -and $d.Path.TrimEnd('\') -eq $target) {
                Write-Host "  先關掉 Word 裡開著的「$($d.Name)」（它會鎖住檔案）"
                $d.Close(0)   # 0 = 不存檔
            }
        }
    } catch {
        # Word 沒開著時 GetActiveObject 會丟例外，那就沒事要做
    }
}

Close-OutputDocs $outDir

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
$out = Get-ChildItem -Path $outDir -Filter '*.docx' -File -ErrorAction SilentlyContinue |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($out) {
    Write-Host ''
    Write-Host "V 完成  $($out.Name)" -ForegroundColor Green
    Write-Host "  位置：$($out.DirectoryName)"

    explorer.exe "/select,$($out.FullName)"   # 在檔案總管裡選取該檔
    Invoke-Item -LiteralPath $out.FullName    # 順手用 Word 開起來校對
} else {
    Write-Host '! 程式回報成功，但 outputs\ 裡找不到檔案' -ForegroundColor Yellow
}

Finish 0
