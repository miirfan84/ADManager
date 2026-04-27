# =============================================================================
# AD Manager v2.0  —  Launch.ps1
# Entry point. Verbose terminal output at every step for easy troubleshooting.
#
# Usage:
#   Right-click → "Run with PowerShell"
#   OR: powershell.exe -STA -ExecutionPolicy Bypass -File ".\Launch.ps1"
# =============================================================================

#region ── STA enforcement ────────────────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $self = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe `
        -ArgumentList "-STA -ExecutionPolicy Bypass -NonInteractive -File `"$self`"" `
        -NoNewWindow
    exit
}
#endregion

#region ── WPF assemblies ─────────────────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
#endregion

#region ── Paths ──────────────────────────────────────────────────────────────
$Script:RootDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:CoreDir    = Join-Path $Script:RootDir "Core"
$Script:UIDir      = Join-Path $Script:RootDir "UI"
$Script:PluginsDir = Join-Path $Script:RootDir "Plugins"
#endregion

#region ── Logger (must load first) ──────────────────────────────────────────
try {
    . ([string](Join-Path $Script:CoreDir "Logger.ps1"))
    Write-LogBanner -Version "2.0"
} catch {
    Write-Host "FATAL: Cannot load Logger.ps1 — $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
#endregion

#region ── Core modules ───────────────────────────────────────────────────────
Write-LogHead "Loading Core Modules"

$coreModules = [ordered]@{
    'AppState.ps1'     = 'Application state + connection event bus'
    'PluginLoader.ps1' = 'Plugin discovery and sandboxed loader'
    'UIShell.ps1'      = 'Main window + connection bar'
    'OUTreePicker.ps1' = 'Shared OU tree picker component'
}

foreach ($module in $coreModules.Keys) {
    $desc = $coreModules[$module]
    $path = Join-Path $Script:CoreDir $module
    Write-LogStep "Loading $module — $desc"
    try {
        if (-not (Test-Path $path)) { throw "File not found: $path" }
        . ([string]$path)
        Write-LogStep $module -Status 'OK'
        Write-Log "Core loaded: $module" -Source 'Launch'
    } catch {
        $err = $_.Exception.Message
        Write-LogStep "$module FAILED: $err" -Status 'FAIL'
        Write-Log "FATAL — Core module '$module' failed: $err" -Level ERROR -Source 'Launch'
        [System.Windows.MessageBox]::Show(
            "Fatal error loading core module '$module':`n`n$err",
            "AD Manager — Startup Error", 'OK', 'Error') | Out-Null
        exit 1
    }
}
#endregion

#region ── Prerequisites note ────────────────────────────────────────────────
Write-LogHead "Checking Prerequisites"
Write-LogStep "RSAT check skipped — all AD operations run remotely on the DC" -Status 'OK'
Write-Log "No local RSAT required — AD cmdlets execute on DC via remote session" -Source 'Launch'
#endregion

#region ── Build shell window ─────────────────────────────────────────────────
Write-LogHead "Building Main Window"

$shellXaml = Join-Path $Script:UIDir "Shell.xaml"
Write-LogStep "Shell XAML path: $shellXaml"

try {
    $Script:Window = Initialize-Shell -XamlPath $shellXaml
    Write-LogStep "Main window created" -Status 'OK'
    Write-Log "Shell window built successfully" -Source 'Launch'
} catch {
    $err = $_.Exception.Message
    Write-LogStep "Shell build FAILED: $err" -Status 'FAIL'
    Write-Log "FATAL — Shell failed to load: $err" -Level ERROR -Source 'Launch'
    [System.Windows.MessageBox]::Show(
        "Cannot load the main window:`n`n$err",
        "AD Manager — Fatal Error", 'OK', 'Error') | Out-Null
    exit 1
}
#endregion

#region ── Load plugins ───────────────────────────────────────────────────────
Write-LogHead "Loading Plugins"

Write-LogStep "Plugin root: $Script:PluginsDir"

try {
    Import-Plugins -TabControl $Script:Shell['MainTabControl'] `
                   -PluginsRoot $Script:PluginsDir
} catch {
    $err = $_.Exception.Message
    Write-LogStep "Plugin loader error: $err" -Status 'FAIL'
    Write-Log "Plugin loader error: $err" -Level ERROR -Source 'Launch'
    [System.Windows.MessageBox]::Show(
        "Plugin loader error:`n$err`n`nApp will continue with no plugins.",
        "AD Manager — Plugin Warning", 'OK', 'Warning') | Out-Null
}

$total  = $Script:LoadedPlugins.Count
$ok     = 0
foreach ($p in $Script:LoadedPlugins) { if ($p.Loaded -eq $true) { $ok++ } }
$failed = $total - $ok
$stepStatus = if ($failed -eq 0) { 'OK' } else { 'WARN' }
Write-LogStep "Plugins: $ok loaded, $failed failed (total $total)" -Status $stepStatus
Write-Log "Plugin summary: $ok/$total loaded" -Source 'Launch'
#endregion

#region ── Show window ────────────────────────────────────────────────────────
Write-LogHead "Launching UI"
Write-LogStep "Showing main window — app ready" -Status 'OK'
Write-Log "AD Manager v$($Script:AppVersion) ready — showing window" -Source 'Launch'
Write-Host ""

[void]$Script:Window.ShowDialog()

Write-Log "Window closed — AD Manager exiting" -Source 'Launch'
Write-LogStep "Session ended" -Status 'OK'
#endregion