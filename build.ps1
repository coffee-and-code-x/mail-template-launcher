#Requires -Version 5.1
<#
.SYNOPSIS
    Mail Template Launcher ビルドスクリプト
    VBA ソースファイル (src/vba/*.bas, *.cls) から
    mail_template_launcher.xlsm を生成します。

.DESCRIPTION
    このスクリプトは Microsoft Excel の COM オートメーションを使用します。
    実行前に Excel がインストールされていることを確認してください。

    【前提条件】
    1. Microsoft Excel がインストールされていること
    2. Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を
       有効にすること（詳細は以下を参照）

    【Trust Center の設定方法】
    Excel を開く → ファイル → オプション → トラストセンター →
    トラストセンターの設定 → マクロの設定 →
    「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック → OK

.PARAMETER OutputPath
    出力する .xlsm ファイルのパス（省略時: スクリプトと同じフォルダ）

.EXAMPLE
    .\build.ps1
    .\build.ps1 -OutputPath "C:\Tools\mail_template_launcher.xlsm"
#>

param(
    [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ============================================================
# 設定
# ============================================================
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$VbaDir     = Join-Path $ScriptDir "src\vba"
$OutputFile = if ($OutputPath -ne "") { $OutputPath } else {
    Join-Path $ScriptDir "mail_template_launcher.xlsm"
}

# VBA モジュールのインポート順序（依存関係のある順に指定）
$ModuleFiles = @(
    "Module_Types.bas",
    "Module_Utils.bas",
    "Module_Init.bas",
    "Module_FileIO.bas",
    "Module_Search.bas",
    "Module_Template.bas",
    "Module_Outlook.bas",
    "Module_Launcher.bas",
    "Module_ButtonHandlers.bas"
)

# ============================================================
# ヘルパー関数
# ============================================================
function Write-Step {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Host "  [!]  $Message" -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Message)
    Write-Host "  [NG] $Message" -ForegroundColor Red
}

# ============================================================
# メイン処理
# ============================================================
Write-Host ""
Write-Host "=====================================" -ForegroundColor Blue
Write-Host "  Mail Template Launcher ビルド" -ForegroundColor Blue
Write-Host "=====================================" -ForegroundColor Blue
Write-Host ""

# --- 1. 前提条件チェック ---
Write-Host "[1/5] 前提条件チェック..."

# VBA ソースディレクトリの確認
if (-not (Test-Path $VbaDir)) {
    Write-Fail "VBA ソースディレクトリが見つかりません: $VbaDir"
    exit 1
}
Write-Success "VBA ソースディレクトリ: $VbaDir"

# 必要ファイルの確認
$missingFiles = @()
foreach ($file in $ModuleFiles) {
    $filePath = Join-Path $VbaDir $file
    if (-not (Test-Path $filePath)) {
        $missingFiles += $file
    }
}
$thisWbFile = Join-Path $VbaDir "ThisWorkbook.cls"
if (-not (Test-Path $thisWbFile)) {
    $missingFiles += "ThisWorkbook.cls"
}

if ($missingFiles.Count -gt 0) {
    Write-Fail "以下の VBA ファイルが見つかりません:"
    $missingFiles | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
    exit 1
}
Write-Success "全 VBA ソースファイル確認済み ($($ModuleFiles.Count + 1) ファイル)"

# Excel の確認
try {
    $testExcel = New-Object -ComObject Excel.Application
    $testExcel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testExcel) | Out-Null
    Write-Success "Microsoft Excel が見つかりました"
} catch {
    Write-Fail "Microsoft Excel が見つかりません。Excel をインストールしてください。"
    exit 1
}

# --- 2. 既存ファイルの確認 ---
Write-Host ""
Write-Host "[2/5] 出力ファイルの確認..."
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

if (Test-Path $OutputFile) {
    $response = Read-Host "  既存ファイルを上書きしますか？ $OutputFile`n  [Y] 上書き / [N] キャンセル (Y/N)"
    if ($response -notmatch "^[Yy]") {
        Write-Host "  キャンセルしました。" -ForegroundColor Yellow
        exit 0
    }
    Remove-Item $OutputFile -Force
    Write-Warn "既存ファイルを削除しました"
}
Write-Success "出力先: $OutputFile"

# --- 3. Excel ワークブック作成 ---
Write-Host ""
Write-Host "[3/5] Excel ワークブックを作成中..."

$excel = $null
$wb    = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible         = $false
    $excel.DisplayAlerts   = $false
    $excel.ScreenUpdating  = $false

    # 新規ワークブック作成
    $wb = $excel.Workbooks.Add()
    Write-Success "新規ワークブックを作成しました"

    # --- 4. VBA モジュールのインポート ---
    Write-Host ""
    Write-Host "[4/5] VBA モジュールをインポート中..."

    # Trust Center チェック
    $vbProject = $null
    try {
        $vbProject = $wb.VBProject
    } catch {
        Write-Host ""
        Write-Fail "VBA プロジェクトへのアクセスが拒否されました。"
        Write-Host ""
        Write-Host "  【解決方法】" -ForegroundColor Yellow
        Write-Host "  Excel を開き、以下の設定を行ってください:" -ForegroundColor Yellow
        Write-Host "    ファイル → オプション → トラストセンター → トラストセンターの設定" -ForegroundColor Yellow
        Write-Host "    → マクロの設定 → 「VBA プロジェクトオブジェクトモデルへの" -ForegroundColor Yellow
        Write-Host "      アクセスを信頼する」にチェック → OK" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  設定後、このスクリプトを再実行してください。" -ForegroundColor Yellow

        $wb.Close($false)
        $excel.Quit()
        exit 1
    }

    # 標準モジュールのインポート
    foreach ($file in $ModuleFiles) {
        $filePath = Join-Path $VbaDir $file
        Write-Step "インポート: $file"
        $vbProject.VBComponents.Import($filePath) | Out-Null
    }

    # ThisWorkbook モジュールの内容を置き換え
    Write-Step "インポート: ThisWorkbook.cls"
    $thisWbContent = Get-Content $thisWbFile -Raw -Encoding UTF8

    # ヘッダー行を除去（VERSION と Attribute で始まる行）
    $codeLines = $thisWbContent -split "`n" | Where-Object {
        $_ -notmatch "^VERSION\s" -and
        $_ -notmatch "^BEGIN\s*$" -and
        $_ -notmatch "^\s+MultiUse\s*=" -and
        $_ -notmatch "^END\s*$" -and
        $_ -notmatch "^Attribute VB_"
    }
    $codeOnly = ($codeLines -join "`n").Trim()

    $thisWbComp = $vbProject.VBComponents.Item("ThisWorkbook")
    $codeModule = $thisWbComp.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.InsertLines(1, $codeOnly)

    Write-Success "全 VBA モジュールのインポート完了"

    # --- 5. ファイルとして保存 ---
    Write-Host ""
    Write-Host "[5/5] ファイルを保存中..."

    # xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($OutputFile, 52)
    Write-Success "保存完了: $OutputFile"

    $wb.Close($false)
    $excel.Quit()

} catch {
    Write-Fail "ビルド中にエラーが発生しました: $_"
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    exit 1
} finally {
    if ($null -ne $wb) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($null -ne $excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ============================================================
# 完了メッセージ
# ============================================================
Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "  ビルド完了！" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host ""
Write-Host "  出力ファイル: $OutputFile" -ForegroundColor White
Write-Host ""
Write-Host "  【次のステップ】" -ForegroundColor Cyan
Write-Host "  1. $OutputFile を開く" -ForegroundColor White
Write-Host "  2. 「マクロを有効にする」を選択" -ForegroundColor White
Write-Host "  3. 初回起動時に自動的に初期化が実行されます" -ForegroundColor White
Write-Host "  4.「ファイル設定」シートで案件データファイルを登録してください" -ForegroundColor White
Write-Host ""
