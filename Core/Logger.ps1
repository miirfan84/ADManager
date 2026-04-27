# =============================================================================
# Core\Logger.ps1  —  AD Manager v2.0
# Dual-output logger: writes to FILE and prints to TERMINAL simultaneously.
# Colour-coded by level. Aligned source column. Section headers + step helpers.
#
# Verbosity controlled by $Script:LogVerbose:
#   $true  → all INFO/WARN/ERROR print to terminal  (default — troubleshoot mode)
#   $false → only WARN + ERROR print to terminal    (quiet mode)
# =============================================================================

$Script:LogDir     = Join-Path $env:TEMP "ADManager"
$Script:LogFile    = Join-Path $Script:LogDir ("ADManager_" + (Get-Date -Format "yyyyMMdd") + ".log")
$Script:LogVerbose = $true          # flip to $false for quiet mode

if (-not (Test-Path $Script:LogDir)) {
    try   { New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null }
    catch { }   # silently skip — file logging will no-op
}

# =============================================================================
# Write-Log  —  core logging function (file + terminal)
# =============================================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level  = 'INFO',
        [string]$Source = 'Core',
        [switch]$NoFile
    )

    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ts    = Get-Date -Format "HH:mm:ss"

    # ── File ─────────────────────────────────────────────────────────────────
    if (-not $NoFile) {
        try { Add-Content -Path $Script:LogFile -Value "[$stamp][$Level][$Source] $Message" -Encoding UTF8 }
        catch { }
    }

    # ── Terminal ──────────────────────────────────────────────────────────────
    if ($Level -ne 'INFO' -or $Script:LogVerbose) {

        $msgColour = switch ($Level) {
            'INFO'  { 'Cyan'   }
            'WARN'  { 'Yellow' }
            'ERROR' { 'Red'    }
        }
        $badgeColour = switch ($Level) {
            'INFO'  { 'DarkCyan' }
            'WARN'  { 'Yellow'   }
            'ERROR' { 'Red'      }
        }
        $badge = switch ($Level) {
            'INFO'  { ' INFO ' }
            'WARN'  { ' WARN ' }
            'ERROR' { 'ERROR ' }
        }

        # Fixed-width source column (14 chars) keeps terminal tidy
        $src = "[$Source]".PadRight(18)

        Write-Host $ts    -NoNewline -ForegroundColor DarkGray
        Write-Host " $badge " -NoNewline -ForegroundColor $badgeColour
        Write-Host $src   -NoNewline -ForegroundColor DarkGray
        Write-Host $Message           -ForegroundColor $msgColour
    }
}

# =============================================================================
# Write-LogHead  —  section separator printed to terminal only
# =============================================================================
function Write-LogHead {
    param([string]$Title)
    $line = "─" * 58
    Write-Host ""
    Write-Host "  $line"  -ForegroundColor DarkCyan
    Write-Host "  ◆  $Title" -ForegroundColor White
    Write-Host "  $line"  -ForegroundColor DarkCyan
}

# =============================================================================
# Write-LogStep  —  numbered/bulleted step with optional OK/FAIL/SKIP badge
# =============================================================================
function Write-LogStep {
    param(
        [string]$Text,
        [ValidateSet('','OK','FAIL','SKIP','WAIT','WARN')]
        [string]$Status = ''
    )
    $ts = Get-Date -Format "HH:mm:ss"

    if ($Status -eq '') {
        Write-Host $ts      -NoNewline -ForegroundColor DarkGray
        Write-Host "   →  " -NoNewline -ForegroundColor DarkCyan
        Write-Host $Text               -ForegroundColor Gray
    } else {
        $col = switch ($Status) {
            'OK'   { 'Green'  }
            'FAIL' { 'Red'    }
            'SKIP' { 'Yellow' }
            'WAIT' { 'Cyan'   }
            'WARN' { 'Yellow' }
        }
        Write-Host $ts                          -NoNewline -ForegroundColor DarkGray
        Write-Host " [$Status]".PadRight(8)     -NoNewline -ForegroundColor $col
        Write-Host $Text                                   -ForegroundColor Gray
    }
}

# =============================================================================
# Write-LogBanner  —  startup splash (call once from Launch.ps1)
# =============================================================================
function Write-LogBanner {
    param([string]$Version = "2.0")
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║         AD Manager  v$Version                    ║" -ForegroundColor Cyan
    Write-Host "  ║         Plugin-based AD Management           ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Verbose logging : ON  (set LogVerbose=false to disable)" -ForegroundColor DarkGray
    Write-Host "  Log file        : $Script:LogFile"                       -ForegroundColor DarkGray
    Write-Host "  Started         : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Log "Logger ready — $Script:LogFile" -Source 'Logger'