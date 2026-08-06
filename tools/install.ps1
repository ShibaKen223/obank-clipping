# ─────────────────────────────────────────────────────────────
# 每日剪報 — 一鍵安裝（Windows 版）
#
# 平常不用自己下，雙擊專案資料夾裡的「安裝.bat」就好。
# 這支負責把新人第一次會卡住的每一步都自動化：
#   確認 Python → 沒有就裝 → 建桌面捷徑 → 預先建好 Python 環境
#
# macOS 沒有對應版本：Mac 的 build_app.sh 就已經夠一步到位了，
# 系統本身就有 python3，不需要這層。
# ─────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'

$PROJECT = Split-Path -Parent $PSScriptRoot
$VENV    = Join-Path $PROJECT '.venv'
$PYEXE   = Join-Path $VENV 'Scripts\python.exe'

# 主控台預設是 cp950，中文會變問號。跟 run_clipping.ps1 同樣處理。
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

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

function Step { param([string]$Text) Write-Host ''; Write-Host $Text -ForegroundColor Cyan }

Write-Host ''
Write-Host '每日剪報 — 安裝' -ForegroundColor White
Write-Host "專案：$PROJECT"
Write-Host '這會花 2–3 分鐘，中途不用按任何東西。'

# ── 1. 找一個能用的 Python ───────────────────────────────────
# 直接打 python 會叫到 Microsoft Store 的空殼：不報錯、也不印版本，
# 所以不能只看「指令存不存在」，要真的執行看它印不印得出版本號。
function Find-Python {
    foreach ($candidate in @('py', 'python')) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        # WindowsApps 底下那個就是 Store 空殼本人，執行它會把應用程式商店
        # 叫起來卡住畫面。先認出來跳過，不要去碰。
        if ($cmd.Source -and $cmd.Source -like '*\WindowsApps\*') { continue }
        try {
            $out = & $candidate '--version' 2>&1 | Out-String
            if ($out -match 'Python 3\.(\d+)') {
                # trafilatura 等套件對 3.8 以下支援不完整，擋掉太舊的
                if ([int]$Matches[1] -ge 9) { return $candidate }
            }
        } catch { }
    }
    return $null
}

# winget 裝完之後 PATH 只有新開的視窗才看得到，這支還在跑的看不到，
# 所以要自己從登錄檔重讀一次，不然會誤判成「裝了還是找不到」。
function Update-PathFromRegistry {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user    = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = @($machine, $user | Where-Object { $_ }) -join ';'
}

Step '[1/3] 確認 Python'

$py = Find-Python
if ($py) {
    $ver = (& $py '--version' 2>&1 | Out-String).Trim()
    Write-Host "  已經有了：$ver"
} else {
    Write-Host '  這台還沒有 Python，正在自動安裝…' -ForegroundColor Yellow

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Die '這台電腦沒有 winget，沒辦法自動裝 Python' @(
            '請手動安裝，只要做一次：',
            '  1. 開 https://www.python.org/downloads/ 下載 Windows 版',
            '  2. 安裝畫面最下面兩個框「Add python.exe to PATH」和',
            '     「py launcher」都要勾起來（很容易漏，漏了要重裝）',
            '  3. 裝完重新雙擊「安裝.bat」')
    }

    # --scope user 是關鍵：裝到使用者自己的資料夾，不需要系統管理員權限，
    # 公司電腦沒有 admin 也裝得起來。
    winget install --id Python.Python.3.12 --exact --source winget `
        --scope user --accept-package-agreements --accept-source-agreements
    Update-PathFromRegistry

    $py = Find-Python
    if (-not $py) {
        Die 'Python 裝完了，但這個視窗還是叫不到它' @(
            '通常關掉這個視窗、重新雙擊一次「安裝.bat」就會好',
            '（新的視窗才吃得到剛剛裝好的路徑）。',
            '再一次還是不行的話，把上面整段訊息複製給管理者看。')
    }
    Write-Host "  裝好了：$((& $py '--version' 2>&1 | Out-String).Trim())" -ForegroundColor Green
}

# ── 2. 桌面捷徑 ──────────────────────────────────────────────
Step '[2/3] 建立桌面的「每日剪報」'

$shortcut = Join-Path ([Environment]::GetFolderPath('Desktop')) '每日剪報.lnk'
& (Join-Path $PSScriptRoot 'build_app.ps1')

# 用「捷徑到底存不存在」判斷成敗，不看 $LASTEXITCODE：呼叫 .ps1 時它只有在
# 對方真的 exit 才會更新，成功路徑上拿到的是上一個指令留下的舊值，會誤判。
if (-not (Test-Path $shortcut)) {
    Die '建立桌面捷徑失敗' @('把上面的訊息複製給管理者看。')
}

# ── 3. 先把 Python 環境建好 ──────────────────────────────────
# 不先建也能用（run_clipping.ps1 會自己補），但那樣新人第一次點桌面圖示
# 會莫名其妙卡一分鐘，還以為當掉了。在安裝階段等，心理上合理得多。
Step '[3/3] 建立 Python 環境（這步最久，約 1–2 分鐘）'

if (-not (Test-Path $PYEXE)) {
    & $py -m venv $VENV
    if (-not (Test-Path $PYEXE)) { Die '建立虛擬環境失敗' @('把上面的訊息複製給管理者看。') }
}

& $PYEXE -m pip install --quiet --upgrade pip
& $PYEXE -m pip install --quiet -r (Join-Path $PROJECT 'requirements.txt')
& $PYEXE -c 'import requests, bs4, trafilatura, docx, PIL' 2>$null
if ($LASTEXITCODE -ne 0) {
    Die '套件裝完仍然叫不到' @('檢查一下網路連線，再雙擊一次「安裝.bat」。')
}
Write-Host '  環境就緒' -ForegroundColor Green

# ── 完成 ─────────────────────────────────────────────────────
# 沒有 Word 不影響主流程，只有封面頁碼要人工填，所以是提醒不是錯誤。
$hasWord = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'

Write-Host ''
Write-Host '安裝完成' -ForegroundColor Green
Write-Host ''
Write-Host '桌面上已經有「每日剪報」了。以後每天兩步：'
Write-Host '  1. 把當天的「MMDD_榮董新聞」信件另存成 .eml，丟進「下載」資料夾'
Write-Host '  2. 雙擊桌面的「每日剪報」'
Write-Host ''
Write-Host '跑完會停在摘要畫面，「接下來要人工做的」列什麼就做什麼。'
Write-Host '產出是草稿，程式不會寄信，校對完請自己寄。'

if (-not $hasWord) {
    Write-Host ''
    Write-Host '! 這台沒有偵測到 Word' -ForegroundColor Yellow
    Write-Host '  其他功能都正常，只有封面目錄的頁碼要自己填（會維持 P.__~__）。'
}

Finish 0
