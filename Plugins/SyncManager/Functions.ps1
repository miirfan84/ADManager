# =============================================================================
# Plugins\SyncManager\Functions.ps1  —  AD Manager v2.0
# Pure logic: CimSession management, task queries, task execution.
# No WPF references anywhere in this file.
#
# REMOTE ACCESS TECHNIQUE:
#   New-CimSession -ComputerName $Server -Credential $Cred -Authentication Negotiate
#   Negotiate auth auto-selects Kerberos (domain-joined) or NTLM as needed.
#   Get-ScheduledTask / Start-ScheduledTask via -CimSession — no WinRM PSSession needed.
#   CIM uses WS-Man (port 5985) by default, falls back to DCOM automatically.
# =============================================================================

# ── Task definitions ──────────────────────────────────────────────────────────
# TaskPath = "\" means root folder in Task Scheduler.
# All three tasks confirmed to be in the root folder.
$Script:SM_Tasks = @(
    @{ Name = 'ALEC_Egnyte_Sync_New';    Path = '\'; CardKey = 'ALEC' },
    @{ Name = 'ESA_Egnyte_Sync';         Path = '\'; CardKey = 'ESA'  },
    @{ Name = 'ESTIMATION_Egnyte_Sync';  Path = '\'; CardKey = 'EST'  }
)

# ── Active CimSession (stored here, cleaned up on disconnect) ─────────────────
$Script:SM_CimSession = $null

# ── Last task result code → human-readable string ────────────────────────────
# Source: Windows Task Scheduler HRESULT codes
function SM-DecodeLastResult {
    param([long]$Code)
    if ($null -eq $Code) { return "Unknown" }

    # WMI LastTaskResult might be a signed integer (e.g. -2147216625 for 0x8004130F)
    $evalCode = [int]$Code
    $hex = "0x{0:X8}" -f $evalCode

    $known = @{
        0x00000000 = "✅ Success (0x0)"
        0x00000001 = "⚠ Incorrect function (0x1)"
        0x00041300 = "✅ Task is ready"
        0x00041301 = "✅ Task is running"
        0x00041302 = "⚠ Task is disabled"
        0x00041303 = "✅ Task has not run yet"
        0x00041304 = "⚠ No valid triggers — will not run"
        0x00041305 = "⚠ Event triggers disabled"
        0x00041306 = "⚠ Task terminated (ran too long)"
        0x00041307 = "⚠ No valid triggers or start time"
        0x00041308 = "⚠ Task is already running"
        0x0004130B = "⚠ Could not start"
        0x0004130D = "⚠ Multiple instances not allowed"
        0x0004130E = "❌ Task stopped — no scheduled triggers left"
        0x0004131B = "⚠ Trigger will not fire (past end date)"
        0x0004131C = "⚠ Trigger will not fire (future start)"
        0x80041309 = "❌ User account not available"
        0x8004130F = "❌ Could not start — path not found"
        0x80070002 = "❌ File not found (0x80070002)"
        0x80070005 = "❌ Access denied (0x80070005)"
        0x80070057 = "❌ Invalid argument (0x80070057)"
        0xC000013A = "❌ Task terminated by user"
        0x800704DD = "❌ Service not available"
    }

    foreach ($key in $known.Keys) {
        if ($evalCode -eq [int]$key) { return "$($known[$key])  [$hex]" }
    }

    return "Exit code $hex"
}

# ── Connect to sync server via CimSession ─────────────────────────────────────
function SM-Connect {
    param(
        [string]$Server,
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Log "SyncManager: connecting to $Server via CimSession..." -Source 'SyncManager'

    # Disconnect existing session first
    SM-Disconnect

    try {
        # Negotiate: auto-picks Kerberos (domain) or NTLM.
        # SkipTestConnection: skip the WS-Identity probe — reduces connect time by ~1s.
        $session = New-CimSession `
            -ComputerName  $Server `
            -Credential    $Credential `
            -Authentication Negotiate `
            -OperationTimeoutSec 30 `
            -ErrorAction   Stop

        $Script:SM_CimSession = $session
        Write-Log "SyncManager: CimSession established to $Server (ID=$($session.Id))" -Source 'SyncManager'
        return @{ OK = $true; Msg = "Connected to $Server" }

    } catch {
        $err = $_.Exception.Message
        Write-Log "SyncManager: CimSession failed — $err" -Level ERROR -Source 'SyncManager'
        return @{ OK = $false; Msg = $err }
    }
}

# ── Disconnect and clean up CimSession ────────────────────────────────────────
function SM-Disconnect {
    if ($null -ne $Script:SM_CimSession) {
        try {
            Remove-CimSession -CimSession $Script:SM_CimSession -ErrorAction SilentlyContinue
        } catch {}
        $Script:SM_CimSession = $null
        Write-Log "SyncManager: CimSession removed" -Source 'SyncManager'
    }
}

# ── Get info for all 3 tasks in one shot ──────────────────────────────────────
# Returns array of hashtables: @{ Name; CardKey; State; LastRunTime; LastResult; ResultText; NextRunTime }
function SM-GetAllTaskInfo {
    if ($null -eq $Script:SM_CimSession) {
        return @()
    }

    $results = New-Object System.Collections.ArrayList

    foreach ($td in $Script:SM_Tasks) {
        $row = @{
            Name        = $td.Name
            Path        = $td.Path
            CardKey     = $td.CardKey
            State       = 'Unknown'
            LastRunTime = $null
            LastResult  = $null
            ResultText  = '—'
            NextRunTime = $null
            Error       = $null
        }

        try {
            $task = Get-ScheduledTask `
                -TaskName   $td.Name `
                -TaskPath   $td.Path `
                -CimSession $Script:SM_CimSession `
                -ErrorAction Stop

            # User explicitly requested NOT to get their status upon fetching
            $row.State       = 'Idle'
            $row.LastRunTime = $null
            $row.LastResult  = $null
            $row.ResultText  = 'Ready to deploy'
            $row.NextRunTime = $null

            Write-Log "SyncManager: $($td.Name) — State=$($row.State)  LastResult=$($row.ResultText)" -Source 'SyncManager'

        } catch {
            $row.State = 'Error'
            $row.Error = $_.Exception.Message
            Write-Log "SyncManager: Error querying $($td.Name) — $($_.Exception.Message)" -Level WARN -Source 'SyncManager'
        }

        [void]$results.Add($row)
    }

    return $results
}

# ── Start a task and return immediately ───────────────────────────────────────
# Returns @{ OK; Msg }
function SM-StartTask {
    param([string]$TaskName, [string]$TaskPath = '\')

    if ($null -eq $Script:SM_CimSession) {
        return @{ OK = $false; Msg = "No CimSession — connect first" }
    }

    Write-Log "SyncManager: starting task '$TaskName'..." -Source 'SyncManager'

    try {
        Start-ScheduledTask `
            -TaskName   $TaskName `
            -TaskPath   $TaskPath `
            -CimSession $Script:SM_CimSession `
            -ErrorAction Stop

        Write-Log "SyncManager: Start-ScheduledTask dispatched for '$TaskName'" -Source 'SyncManager'
        return @{ OK = $true; Msg = "Task started" }

    } catch {
        $err = $_.Exception.Message
        Write-Log "SyncManager: failed to start '$TaskName' — $err" -Level ERROR -Source 'SyncManager'
        return @{ OK = $false; Msg = $err }
    }
}

# ── Poll task state (called repeatedly by the UI timer) ───────────────────────
# Returns @{ State; LastResult; ResultText; LastRunTime }
function SM-PollTask {
    param([string]$TaskName, [string]$TaskPath = '\')

    if ($null -eq $Script:SM_CimSession) {
        return @{ State = 'NoSession'; LastResult = $null; ResultText = '—'; LastRunTime = $null }
    }

    try {
        $task = Get-ScheduledTask `
            -TaskName   $TaskName `
            -TaskPath   $TaskPath `
            -CimSession $Script:SM_CimSession `
            -ErrorAction Stop

        $info = Get-ScheduledTaskInfo `
            -TaskName   $TaskName `
            -TaskPath   $TaskPath `
            -CimSession $Script:SM_CimSession `
            -ErrorAction Stop

        return @{
            State       = [string]$task.State
            LastResult  = $info.LastTaskResult
            ResultText  = SM-DecodeLastResult -Code $info.LastTaskResult
            LastRunTime = $info.LastRunTime
        }

    } catch {
        Write-Log "SyncManager: poll error for '$TaskName' — $($_.Exception.Message)" -Level WARN -Source 'SyncManager'
        return @{ State = 'PollError'; LastResult = $null; ResultText = "Poll error: $($_.Exception.Message)"; LastRunTime = $null }
    }
}
