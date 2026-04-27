# ==============================================================================
# Setup-ADManager.ps1  —  AD Manager v2.0
# One-click installer / updater. Drop all files anywhere, run this. Done.
# Safe to re-run — updates in-place. Plugins\ never wiped.
# Files with shared names disambiguated by parent folder hint.
# ==============================================================================

#region ── CONFIG ───────────────────────────────────────────────────────────────
$BUILD_VERSION = "2.0"
$BUILD_DATE    = "2025-04-20"

$FILE_MAP = [ordered]@{

    # Root
    "Launch.ps1"                                    = @( "Launch.ps1",        "" )
    "PLUGINS_README.md"                             = @( "PLUGINS_README.md", "" )
    "Setup-ADManager.ps1"                           = @( "Setup-ADManager.ps1","" )

    # Core engine
    "Core\Logger.ps1"                               = @( "Logger.ps1",        "Core" )
    "Core\AppState.ps1"                             = @( "AppState.ps1",      "Core" )
    "Core\PluginLoader.ps1"                         = @( "PluginLoader.ps1",  "Core" )
    "Core\UIShell.ps1"                              = @( "UIShell.ps1",       "Core" )
    "Core\OUTreePicker.ps1"                         = @( "OUTreePicker.ps1",  "Core" )

    # Shell UI
    "UI\Shell.xaml"                                 = @( "Shell.xaml",        "UI"   )

    # Plugin: UserManagement
    "Plugins\UserManagement\plugin.json"            = @( "plugin.json",       "UserManagement" )
    "Plugins\UserManagement\Tab.xaml"               = @( "Tab.xaml",          "UserManagement" )
    "Plugins\UserManagement\Functions.ps1"          = @( "Functions.ps1",     "UserManagement" )
    "Plugins\UserManagement\Handlers.ps1"           = @( "Handlers.ps1",      "UserManagement" )

    # Plugin: SyncManager
    "Plugins\SyncManager\plugin.json"               = @( "plugin.json",       "SyncManager" )
    "Plugins\SyncManager\Tab.xaml"                  = @( "Tab.xaml",          "SyncManager" )
    "Plugins\SyncManager\Functions.ps1"             = @( "Functions.ps1",     "SyncManager" )
    "Plugins\SyncManager\Handlers.ps1"              = @( "Handlers.ps1",      "SyncManager" )

    # Plugin: ComputerMapper (Phase 1)
    "Plugins\ComputerMapper\plugin.json"            = @( "plugin.json",       "ComputerMapper" )
    "Plugins\ComputerMapper\Tab.xaml"               = @( "Tab.xaml",          "ComputerMapper" )
    "Plugins\ComputerMapper\Functions.ps1"          = @( "Functions.ps1",     "ComputerMapper" )
    "Plugins\ComputerMapper\Handlers.ps1"           = @( "Handlers.ps1",      "ComputerMapper" )

    # Plugin: OUMover (Phase 2)
    "Plugins\OUMover\plugin.json"                   = @( "plugin.json",       "OUMover" )
    "Plugins\OUMover\Tab.xaml"                      = @( "Tab.xaml",          "OUMover" )
    "Plugins\OUMover\Functions.ps1"                 = @( "Functions.ps1",     "OUMover" )
    "Plugins\OUMover\Handlers.ps1"                  = @( "Handlers.ps1",      "OUMover" )
}

$REQUIRED_FOLDERS = @(
    "Core", "UI", "Plugins",
    "Plugins\UserManagement",
    "Plugins\SyncManager",
    "Plugins\ComputerMapper",
    "Plugins\OUMover"
)
#endregion

#region ── HELPERS ──────────────────────────────────────────────────────────────
function Write-Header { param([string]$T)
    Write-Host ""; Write-Host ("=" * 64) -ForegroundColor Cyan
    Write-Host "  $T" -ForegroundColor White
    Write-Host ("=" * 64) -ForegroundColor Cyan
}
function Write-Step { param([string]$T) Write-Host "  >> $T"    -ForegroundColor Gray   }
function Write-OK   { param([string]$T) Write-Host "  [OK] $T"  -ForegroundColor Green  }
function Write-Warn { param([string]$T) Write-Host "  [!!] $T"  -ForegroundColor Yellow }
function Write-Fail { param([string]$T) Write-Host "  [XX] $T"  -ForegroundColor Red    }

function Find-SourceFile {
    param([string]$Root, [string]$FileName, [string]$ParentHint)
    $allMatches = Get-ChildItem -Path $Root -Filter $FileName -Recurse -File -ErrorAction SilentlyContinue
    if ($allMatches.Count -eq 0) { return $null }
    if ([string]::IsNullOrEmpty($ParentHint)) { return $allMatches[0].FullName }
    foreach ($f in $allMatches) {
        if ($f.DirectoryName -like "*$ParentHint*") { return $f.FullName }
    }
    return $allMatches[0].FullName
}
#endregion

#region ── MAIN ─────────────────────────────────────────────────────────────────
Clear-Host
Write-Header "AD Manager v$BUILD_VERSION — Setup / Updater  ($BUILD_DATE)"

$SetupDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$TargetDir = Join-Path $SetupDir "ADManager"

Write-Host ""
Write-Host "  Searching under : $SetupDir"  -ForegroundColor DarkCyan
Write-Host "  Installing to   : $TargetDir" -ForegroundColor DarkCyan

# ── Folders ───────────────────────────────────────────────────────────────────
Write-Host ""; Write-Step "Creating folder structure..."
if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    Write-OK "Created: ADManager\"
}
foreach ($folder in $REQUIRED_FOLDERS) {
    $fp = Join-Path $TargetDir $folder
    if (-not (Test-Path $fp)) {
        New-Item -ItemType Directory -Path $fp -Force | Out-Null
        Write-OK "Created: ADManager\$folder\"
    }
}

# ── Copy files ────────────────────────────────────────────────────────────────
Write-Host ""; Write-Step "Finding and installing files..."
$Copied  = [System.Collections.ArrayList]::new()
$Missing = [System.Collections.ArrayList]::new()

foreach ($destRel in $FILE_MAP.Keys) {
    $info       = $FILE_MAP[$destRel]
    $fileName   = $info[0]
    $parentHint = $info[1]
    $destFull   = Join-Path $TargetDir $destRel

    $srcFull = Find-SourceFile -Root $SetupDir -FileName $fileName -ParentHint $parentHint

    if ($null -ne $srcFull) {
        if ($srcFull -eq $destFull) {
            Write-OK "In place  : $destRel"
            [void]$Copied.Add($destRel)
        } else {
            try {
                Copy-Item -Path $srcFull -Destination $destFull -Force -ErrorAction Stop
                $srcFolder = Split-Path (Split-Path $srcFull -Parent) -Leaf
                Write-OK "Installed : $destRel  ← ...\$srcFolder\$fileName"
                [void]$Copied.Add($destRel)
            } catch {
                Write-Fail "Copy error : $destRel — $($_.Exception.Message)"
                [void]$Missing.Add($destRel)
            }
        }
    } else {
        Write-Warn "Not found : $fileName  (hint: $parentHint)"
        [void]$Missing.Add($destRel)
    }
}

# ── Verify ────────────────────────────────────────────────────────────────────
Write-Host ""; Write-Step "Verifying..."
$Verified = [System.Collections.ArrayList]::new()
$Failed   = [System.Collections.ArrayList]::new()
foreach ($destRel in $FILE_MAP.Keys) {
    $fp = Join-Path $TargetDir $destRel
    if (Test-Path $fp) { [void]$Verified.Add($destRel) } else { [void]$Failed.Add($destRel) }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Header "Summary"
Write-Host "  Installed : $($Copied.Count) / $($FILE_MAP.Count) files"   -ForegroundColor White
Write-Host "  Verified  : $($Verified.Count) / $($FILE_MAP.Count) files" -ForegroundColor White

if ($Failed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Missing files (re-download and re-run Setup):" -ForegroundColor Red
    foreach ($f in $Failed) { Write-Host "    - $f" -ForegroundColor Red }
    Write-Host ""; Write-Warn "Setup completed with warnings."
} else {
    Write-Host ""
    Write-Host "  =========================================================" -ForegroundColor Green
    Write-Host "   ✅  AD Manager v$BUILD_VERSION installed successfully!"    -ForegroundColor Green
    Write-Host "   📁  $TargetDir"                                             -ForegroundColor Green
    Write-Host "  =========================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Plugins installed:" -ForegroundColor White
    Write-Host "    • UserManagement  — search users, reset passwords"        -ForegroundColor Cyan
    Write-Host "    • SyncManager     — run Egnyte sync tasks remotely"        -ForegroundColor Cyan
    Write-Host "    • ComputerMapper  — match CSV assets to AD computers"     -ForegroundColor Cyan
    Write-Host "    • OUMover         — move users/computers to correct OUs"  -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Launch:" -ForegroundColor White
    Write-Host "    Right-click Launch.ps1 → Run with PowerShell"             -ForegroundColor Gray
    Write-Host "    OR: powershell -STA -ExecutionPolicy Bypass -File `"$TargetDir\Launch.ps1`"" -ForegroundColor DarkCyan
}

Write-Host ""
if ($Host.Name -eq 'ConsoleHost') {
    Write-Host "  Press any key to exit..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
#endregion
