#Requires -Version 5.1
<#
.SYNOPSIS
    Mail Template Launcher build script.
    Generates mail_template_launcher.xlsm from VBA source files (src/vba/).

.DESCRIPTION
    Requires Microsoft Excel to be installed.

    [Prerequisites]
    1. Microsoft Excel must be installed.
    2. Enable "Trust access to the VBA project object model" in Excel Trust Center:
       Excel -> File -> Options -> Trust Center -> Trust Center Settings
       -> Macro Settings -> check "Trust access to the VBA project object model" -> OK

.PARAMETER OutputPath
    Path for the output .xlsm file (default: same folder as this script)

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
# Configuration
# ============================================================
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$VbaDir     = Join-Path $ScriptDir "src\vba"
$OutputFile = if ($OutputPath -ne "") { $OutputPath } else {
    Join-Path $ScriptDir "mail_template_launcher.xlsm"
}

# VBA module import order (dependency order)
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
# Helper functions
# ============================================================
function Write-Step    { param([string]$m) Write-Host "  $m"       -ForegroundColor Cyan  }
function Write-Success { param([string]$m) Write-Host "  [OK] $m"  -ForegroundColor Green }
function Write-Warn    { param([string]$m) Write-Host "  [!]  $m"  -ForegroundColor Yellow }
function Write-Fail    { param([string]$m) Write-Host "  [NG] $m"  -ForegroundColor Red   }

# ============================================================
# Main
# ============================================================
Write-Host ""
Write-Host "======================================" -ForegroundColor Blue
Write-Host "  Mail Template Launcher - Build"       -ForegroundColor Blue
Write-Host "======================================" -ForegroundColor Blue
Write-Host ""

# --- Step 1: Prerequisites ---
Write-Host "[1/5] Checking prerequisites..."

if (-not (Test-Path $VbaDir)) {
    Write-Fail "VBA source directory not found: $VbaDir"
    exit 1
}
Write-Success "VBA source directory: $VbaDir"

$missingFiles = @()
foreach ($file in $ModuleFiles) {
    if (-not (Test-Path (Join-Path $VbaDir $file))) { $missingFiles += $file }
}
$thisWbFile = Join-Path $VbaDir "ThisWorkbook.cls"
if (-not (Test-Path $thisWbFile)) { $missingFiles += "ThisWorkbook.cls" }

if ($missingFiles.Count -gt 0) {
    Write-Fail "Missing VBA files:"
    $missingFiles | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
    exit 1
}
Write-Success "All VBA source files found ($($ModuleFiles.Count + 1) files)"

try {
    $testExcel = New-Object -ComObject Excel.Application
    $testExcel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testExcel) | Out-Null
    Write-Success "Microsoft Excel found"
} catch {
    Write-Fail "Microsoft Excel not found. Please install Excel and retry."
    exit 1
}

# --- Step 2: Output file check ---
Write-Host ""
Write-Host "[2/5] Checking output file..."
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

if (Test-Path $OutputFile) {
    $response = Read-Host "  File already exists. Overwrite? $OutputFile`n  [Y] Yes / [N] Cancel (Y/N)"
    if ($response -notmatch "^[Yy]") {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        exit 0
    }
    Remove-Item $OutputFile -Force
    Write-Warn "Existing file removed"
}
Write-Success "Output: $OutputFile"

# --- Step 3: Create workbook ---
Write-Host ""
Write-Host "[3/5] Creating Excel workbook..."

$excel = $null
$wb    = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible        = $false
    $excel.DisplayAlerts  = $false
    $excel.ScreenUpdating = $false

    $wb = $excel.Workbooks.Add()
    Write-Success "New workbook created"

    # --- Step 4: Import VBA modules ---
    Write-Host ""
    Write-Host "[4/5] Importing VBA modules..."

    $vbProject = $null
    try {
        $vbProject = $wb.VBProject
    } catch {
        Write-Host ""
        Write-Fail "Access to VBA project denied."
        Write-Host ""
        Write-Host "  [Fix] Enable the following setting in Excel:" -ForegroundColor Yellow
        Write-Host "    File -> Options -> Trust Center -> Trust Center Settings" -ForegroundColor Yellow
        Write-Host "    -> Macro Settings -> check:" -ForegroundColor Yellow
        Write-Host "       'Trust access to the VBA project object model'" -ForegroundColor Yellow
        Write-Host "    -> OK, then re-run this script." -ForegroundColor Yellow
        $wb.Close($false)
        $excel.Quit()
        exit 1
    }

    # VBComponents.Import reads files as Shift-JIS (cp932) on Japanese Windows.
    # Convert each UTF-8 source file to a Shift-JIS temp file before importing.
    $sjis     = [System.Text.Encoding]::GetEncoding(932)
    $utf8     = [System.Text.Encoding]::UTF8
    $tempFiles = @()

    foreach ($file in $ModuleFiles) {
        $filePath = Join-Path $VbaDir $file
        Write-Step "Importing: $file"

        # Read as UTF-8, write temp file as Shift-JIS
        $content  = [System.IO.File]::ReadAllText($filePath, $utf8)
        $tempPath = [System.IO.Path]::GetTempFileName()
        $tempPath = [System.IO.Path]::ChangeExtension($tempPath, [System.IO.Path]::GetExtension($file))
        [System.IO.File]::WriteAllText($tempPath, $content, $sjis)
        $tempFiles += $tempPath

        $vbProject.VBComponents.Import($tempPath) | Out-Null
    }

    # Replace ThisWorkbook module content (insert lines directly, no file import)
    Write-Step "Importing: ThisWorkbook.cls"
    $thisWbContent = [System.IO.File]::ReadAllText($thisWbFile, $utf8)
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

    # Clean up temp files
    $tempFiles | ForEach-Object { Remove-Item $_ -Force -ErrorAction SilentlyContinue }

    Write-Success "All VBA modules imported"

    # --- Step 5: Save ---
    Write-Host ""
    Write-Host "[5/5] Saving file..."

    # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.SaveAs($OutputFile, 52)
    Write-Success "Saved: $OutputFile"

    $wb.Close($false)
    $excel.Quit()

} catch {
    Write-Fail "Build failed: $_"
    if ($null -ne $wb)    { try { $wb.Close($false) }  catch {} }
    if ($null -ne $excel) { try { $excel.Quit() }       catch {} }
    exit 1
} finally {
    if ($null -ne $wb)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)    | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ============================================================
# Done
# ============================================================
Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "  Build complete!"                      -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Output: $OutputFile" -ForegroundColor White
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor Cyan
Write-Host "  1. Open $OutputFile" -ForegroundColor White
Write-Host "  2. Click 'Enable Macros'" -ForegroundColor White
Write-Host "  3. Workbook initializes automatically on first open" -ForegroundColor White
Write-Host "  4. Register your project data files in the [File Settings] sheet" -ForegroundColor White
Write-Host ""
