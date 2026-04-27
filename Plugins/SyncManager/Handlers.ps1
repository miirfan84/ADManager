# =============================================================================
# Plugins\SyncManager\Handlers.ps1  —  AD Manager v2.0
# Event wiring for the Sync Manager tab.
# =============================================================================

Write-Log "SyncManager: registering handlers..." -Source 'SyncManager'

# ── Grab controls ─────────────────────────────────────────────────────────────
$SM = @{}
$smControls = @(
    'SM_TxtServer', 'SM_ConnDot', 'SM_ConnLabel', 'SM_ConnMsg',
    'SM_BtnConnect', 'SM_BtnDisconnect',
    'SM_TaskGrid', 'SM_BtnRun', 'SM_StatusText'
)
foreach ($n in $smControls) {
    $ctrl = $Script:PluginRoot.FindName($n)
    if ($null -ne $ctrl) {
        $SM[$n] = $ctrl
        Write-Log "  Control: $n" -Source 'SyncManager'
    } else {
        Write-Log "  MISSING control: $n" -Level WARN -Source 'SyncManager'
    }
}

# ── Plugin state ──────────────────────────────────────────────────────────────
$Script:SM_Connected      = $false
$Script:SM_SyncServer     = "sync.alec.com"
$Script:SM_PollTimer      = $null
$Script:SM_PollTaskName   = $null
$Script:SM_RunStartTick   = 0

# Grid source - Using ObservableCollection for live UI updates
$Script:SM_TaskCollection = New-Object System.Collections.ObjectModel.ObservableCollection[psobject]
$SM['SM_TaskGrid'].ItemsSource = $Script:SM_TaskCollection

# Handle selection changes to enable Run button
$SM['SM_TaskGrid'].Add_SelectionChanged({
    if ($Script:SM_Connected -and $null -ne $SM['SM_TaskGrid'].SelectedItem) {
        $SM['SM_BtnRun'].IsEnabled = $true
    } else {
        $SM['SM_BtnRun'].IsEnabled = $false
    }
})

# ── Status helper ─────────────────────────────────────────────────────────────
function SM-SetStatus {
    param([string]$Text)
    $SM['SM_StatusText'].Text = $Text
    Write-Log "SyncManager: $Text" -Source 'SyncManager'
}

# ── React to main app disconnect ─────────────────────────────────────────────
Register-ConnectionHandler -Handler {
    param([bool]$Connected)
    if (-not $Connected) {
        if ($Script:SM_Connected) {
            SM-Disconnect
            $Script:SM_Connected = $false
            SM-SetConnectedUI -Connected $false
            SM-SetStatus "Main DC disconnected — Sync connection closed"
        }
        $SM['SM_BtnConnect'].IsEnabled = $false
    } else {
        $SM['SM_BtnConnect'].IsEnabled = $true
    }
}

# =============================================================================
# CONNECT Background Timer
# =============================================================================
$Script:SM_ConnTimer = New-Object System.Windows.Threading.DispatcherTimer
$Script:SM_ConnTimer.Interval = [TimeSpan]::FromMilliseconds(300)
$Script:SM_ConnTimer.Add_Tick({
    if (-not $Script:SMConnHandle) { return }

    if ($Script:SMConnHandle.IsCompleted) {
        $Script:SM_ConnTimer.Stop()
        try { $Script:SMConnPS.EndInvoke($Script:SMConnHandle) } catch {}

        $elapsed = ([Environment]::TickCount - $Script:SMConnStartTick) / 1000.0
        if ($elapsed -lt 0) { $elapsed = 0 }

        $res = $SM['__SMConnResult']
        $SM.Remove('__SMConnResult')

        if ($res -and $res.OK) {
            $Script:SM_CimSession = $res.Session
            $Script:SM_Connected  = $true
            SM-SetConnectedUI -Connected $true
            SM-SetStatus "Connected to $($Script:SM_SyncServer)"
            SM-RefreshGrid
        } else {
            $msg = if ($res) { $res.Msg } else { "Unknown error" }
            $SM['SM_BtnConnect'].IsEnabled = $true
            $SM['SM_ConnLabel'].Text       = "Failed"
            $SM['SM_ConnMsg'].Text         = "❌ $msg"
            $SM['SM_ConnMsg'].Foreground   = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xE7,0x4C,0x3C)
            SM-SetStatus "Connection failed: $msg"
            # Cleanup if failed
            try { $Script:SMConnPS.Dispose(); $Script:SMConnRS.Close(); $Script:SMConnRS.Dispose() } catch {}
        }
        $Script:SMConnHandle = $null
    }
})

# =============================================================================
# CONNECT button
# =============================================================================
$SM['SM_BtnConnect'].Add_Click({
    if (-not (Get-AppConnected)) {
        $SM['SM_ConnMsg'].Text = "⚠ Connect to the DC first"
        $SM['SM_ConnMsg'].Foreground = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xE7,0x4C,0x3C)
        return
    }

    $server = $SM['SM_TxtServer'].Text.Trim()
    if ([string]::IsNullOrWhiteSpace($server)) {
        $SM['SM_ConnMsg'].Text = "⚠ Enter a server name"
        return
    }

    $Script:SM_SyncServer = $server
    $SM['SM_BtnConnect'].IsEnabled    = $false
    $SM['SM_ConnMsg'].Text            = ""
    $SM['SM_ConnLabel'].Text          = "Connecting..."
    $SM['SM_ConnLabel'].Foreground    = [Windows.Media.Brushes]::Gray

    SM-SetStatus "Connecting to $server..."

    # Capture for runspace
    $rsServer   = $server
    $rsCred     = Get-AppCredential
    $rsWindow   = $Script:Window
    $rsSM       = $SM
    $rsLogFile  = $Script:LogFile

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $rs  = [RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'MTA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()

    foreach ($v in @('rsServer','rsCred','rsWindow','rsSM','rsLogFile')) {
        $rs.SessionStateProxy.SetVariable($v, (Get-Variable $v -ValueOnly))
    }

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        $result = $null
        try {
            $session = New-CimSession `
                -ComputerName   $rsServer `
                -Credential     $rsCred `
                -Authentication Negotiate `
                -OperationTimeoutSec 30 `
                -ErrorAction    Stop

            $result = @{ OK=$true; Msg="Connected"; Session=$session }
        } catch {
            $result = @{ OK=$false; Msg=$_.Exception.Message }
        }
        $r = $result
        $rsWindow.Dispatcher.Invoke([action]{ $rsSM['__SMConnResult'] = $r })
    })

    $Script:SMConnHandle    = $ps.BeginInvoke()
    $Script:SMConnStartTick = [Environment]::TickCount
    $Script:SMConnPS        = $ps
    $Script:SMConnRS        = $rs

    $Script:SM_ConnTimer.Start()
})

# =============================================================================
# DISCONNECT button
# =============================================================================
$SM['SM_BtnDisconnect'].Add_Click({
    SM-StopPollTimer
    SM-Disconnect
    $Script:SM_Connected    = $false
    SM-SetConnectedUI -Connected $false
    $Script:SM_TaskCollection.Clear()
    SM-SetStatus "Disconnected from $($Script:SM_SyncServer)"
})

# =============================================================================
# RUN button
# =============================================================================
$SM['SM_BtnRun'].Add_Click({
    $selected = $SM['SM_TaskGrid'].SelectedItem
    if ($null -eq $selected) { return }

    SM-SetStatus "Starting task: $($selected.Name)..."
    $SM['SM_BtnRun'].IsEnabled = $false

    $result = SM-StartTask -TaskName $selected.Name -TaskPath $selected.TaskPath

    if (-not $result.OK) {
        SM-SetStatus "Failed to start: $($result.Msg)"
        $SM['SM_BtnRun'].IsEnabled = $true
        return
    }

    SM-SetStatus "Task dispatched — watching live status..."
    $Script:SM_PollTaskName = $selected.Name
    $Script:SM_PollTaskPath = $selected.TaskPath
    $Script:SM_RunStartTick = [Environment]::TickCount

    SM-StartPollTimer
})

# =============================================================================
# POLL TIMER — watches task state every 3 seconds
# =============================================================================
function SM-StartPollTimer {
    SM-StopPollTimer

    $pollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $pollTimer.Interval = [TimeSpan]::FromSeconds(3)

    $pollTimer.Add_Tick({
        $elapsed = ([Environment]::TickCount - $Script:SM_RunStartTick) / 1000.0
        if ($elapsed -lt 0) { $elapsed = 0 }
        $elapsedStr = "$([Math]::Round($elapsed,0))s"

        $info = SM-PollTask -TaskName $Script:SM_PollTaskName -TaskPath $Script:SM_PollTaskPath
        $state = $info.State

        # Find row in collection and update it in place
        $row = $Script:SM_TaskCollection | Where-Object { $_.Name -eq $Script:SM_PollTaskName }
        if ($null -ne $row) {
            if ($state -eq 'Running') {
                $row.State = "Running ($elapsedStr)"
            } else {
                $row.State = $state
            }
            $row.LastRunTime = if ($null -ne $info.LastRunTime -and $info.LastRunTime -ne [DateTime]::MinValue) { 
                $info.LastRunTime.ToString("HH:mm:ss") 
            } else { "—" }
            $row.ResultText  = $info.ResultText
            
            # Refresh the grid item to show changes
            $SM['SM_TaskGrid'].Items.Refresh()
        }

        if ($state -eq 'Running') {
            SM-SetStatus "Task running... ($elapsedStr)"
        } elseif ($state -eq 'PollError') {
            SM-SetStatus "Error: $($info.ResultText)"
        } elseif ($state -eq 'NoSession') {
            SM-StopPollTimer
            SM-SetStatus "Session lost"
            $SM['SM_BtnRun'].IsEnabled = $true
        } else {
            # Finished
            SM-StopPollTimer
            SM-SetStatus "Task finished: $state"
            $SM['SM_BtnRun'].IsEnabled = $true
        }
    })

    $Script:SM_PollTimer = $pollTimer
    $pollTimer.Start()
}

function SM-StopPollTimer {
    if ($null -ne $Script:SM_PollTimer) {
        $Script:SM_PollTimer.Stop()
        $Script:SM_PollTimer = $null
    }
}

# =============================================================================
# UI STATE HELPERS
# =============================================================================
function SM-SetConnectedUI {
    param([bool]$Connected)
    if ($Connected) {
        $SM['SM_ConnDot'].Fill          = [Windows.Media.Brushes]::LimeGreen
        $SM['SM_ConnLabel'].Text        = "Connected: $($Script:SM_SyncServer)"
        $SM['SM_ConnLabel'].Foreground  = [Windows.Media.Brushes]::LightGreen
        $SM['SM_BtnConnect'].IsEnabled  = $false
        $SM['SM_BtnDisconnect'].IsEnabled = $true
        $SM['SM_TxtServer'].IsEnabled   = $false
        $SM['SM_ConnMsg'].Text          = ""
    } else {
        $SM['SM_ConnDot'].Fill          = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xE7,0x4C,0x3C)
        $SM['SM_ConnLabel'].Text        = "Not connected"
        $SM['SM_ConnLabel'].Foreground  = [Windows.Media.Brushes]::Gray
        $SM['SM_BtnConnect'].IsEnabled  = (Get-AppConnected)
        $SM['SM_BtnDisconnect'].IsEnabled = $false
        $SM['SM_TxtServer'].IsEnabled   = $true
        $SM['SM_BtnRun'].IsEnabled      = $false
        SM-SetStatus "Connect to sync.alec.com to begin"
    }
}

function SM-RefreshGrid {
    SM-SetStatus "Refreshing task list..."
    $Script:SM_TaskCollection.Clear()

    $allInfo = SM-GetAllTaskInfo

    foreach ($info in $allInfo) {
        $obj = New-Object PSObject -Property @{
            Name        = $info.Name
            TaskPath    = $info.Path
            State       = $info.State
            LastRunTime = if ($null -ne $info.LastRunTime -and $info.LastRunTime -ne [DateTime]::MinValue) { 
                $info.LastRunTime.ToString("HH:mm:ss") 
            } else { "Never" }
            ResultText  = $info.ResultText
        }
        $Script:SM_TaskCollection.Add($obj)
    }
    SM-SetStatus "Task list updated"
}

Write-Log "SyncManager handlers registered" -Source 'SyncManager'
