# =============================================================================
# Plugins\DnsManager\Handlers.ps1  —  AD Manager v2.0
# =============================================================================

Write-Log "DnsManager: registering handlers..." -Source 'DnsManager'

# ── Path to Functions.ps1 (used for dot-sourcing in click handlers + runspaces)
$Script:DNS_FnPath = Join-Path $PSScriptRoot "Functions.ps1"

# ── Grab controls via PluginRoot (NOT $Script:Window — controls are inside the tab)
$Script:DNS = @{}
foreach ($n in @(
    'DNS_CboZone','DNS_BtnLoadZones','DNS_TxtInput','DNS_BtnFetch',
    'DNS_BtnApply','DNS_BtnDelete','DNS_BtnExport',
    'DNS_WhatIfBadge','DNS_WhatIfLabel','DNS_WhatIfToggle',
    'DNS_Progress','DNS_ProgressLabel',
    'DNS_Grid','DNS_StatusBar','DNS_SelCount','DNS_LblWinRM'
)) {
    $ctrl = $Script:PluginRoot.FindName($n)
    if ($null -ne $ctrl) {
        $Script:DNS[$n] = $ctrl
        Write-Log "  Control: $n" -Source 'DnsManager'
    } else {
        Write-Log "  MISSING: $n" -Level WARN -Source 'DnsManager'
    }
}

# ── State
$Script:DNS_AllRows     = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:DNS_IsRunning   = $false
$Script:DNS_ApplyQueue  = $null
$Script:DNS_ApplyIdx    = 0
$Script:DNS_ApplyOK     = 0
$Script:DNS_ApplyFailed = 0

# Bind grid
$Script:DNS['DNS_Grid'].ItemsSource = $Script:DNS_AllRows

# =============================================================================
# HELPER SCRIPTBLOCKS (Script-scoped so visible inside all event closures)
# =============================================================================

$Script:DNS_SetStatus = {
    param([string]$Text)
    $Script:DNS['DNS_StatusBar'].Text = $Text
}

$Script:DNS_ShowProgress = {
    param([int]$Value = -1, [string]$Label = '')
    if ($Value -lt 0) {
        $Script:DNS['DNS_Progress'].Visibility      = 'Collapsed'
        $Script:DNS['DNS_ProgressLabel'].Visibility = 'Collapsed'
    } else {
        $Script:DNS['DNS_Progress'].Visibility      = 'Visible'
        $Script:DNS['DNS_ProgressLabel'].Visibility = 'Visible'
        $Script:DNS['DNS_Progress'].Value           = [Math]::Min($Value, 100)
        if ($Label -ne '') { $Script:DNS['DNS_ProgressLabel'].Text = $Label }
    }
}

$Script:DNS_UpdateSelCount = {
    $Script:DNS['DNS_SelCount'].Text = "$($Script:DNS['DNS_Grid'].SelectedItems.Count) selected"
}

# =============================================================================
# EVENT WIRING
# =============================================================================

# ── Connection handler ────────────────────────────────────────────────────────
Register-ConnectionHandler -Handler {
    param([bool]$Connected)
    if ($null -eq $Script:DNS -or $null -eq $Script:DNS['DNS_StatusBar']) { return }
    if ($Connected) {
        $Script:DNS['DNS_BtnLoadZones'].IsEnabled = $true
        $Script:DNS['DNS_TxtInput'].IsEnabled     = $true
        & $Script:DNS_SetStatus "Select a zone and load records"
    } else {
        $Script:DNS['DNS_BtnLoadZones'].IsEnabled = $false
        $Script:DNS['DNS_TxtInput'].IsEnabled     = $false
        $Script:DNS['DNS_BtnFetch'].IsEnabled     = $false
        $Script:DNS['DNS_BtnApply'].IsEnabled     = $false
        $Script:DNS['DNS_BtnDelete'].IsEnabled    = $false
        $Script:DNS['DNS_BtnExport'].IsEnabled    = $false
        $Script:DNS['DNS_CboZone'].ItemsSource    = $null
        $Script:DNS['DNS_CboZone'].IsEnabled      = $false
        $Script:DNS_AllRows.Clear()
        & $Script:DNS_SetStatus "Connect to a DC to begin"
    }
}

# ── WhatIf badge toggle ───────────────────────────────────────────────────────
$Script:DNS['DNS_WhatIfBadge'].Add_MouseLeftButtonUp({
    $current = Get-WhatIfMode
    if ($current) {
        $c = [Windows.MessageBox]::Show(
            "Switch to LIVE MODE?`n`nChanges WILL be written to DNS.",
            "WhatIf Warning", 'YesNo', 'Warning')
        if ($c -ne 'Yes') { return }
        Set-WhatIfMode -Mode $false
        $Script:DNS['DNS_WhatIfBadge'].Background = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xC0, 0x39, 0x2B)
        $Script:DNS['DNS_WhatIfLabel'].Text = "OFF — LIVE"
        if ($Script:DNS['DNS_WhatIfToggle']) { $Script:DNS['DNS_WhatIfToggle'].IsChecked = $false }
        Write-Log "DnsManager: WhatIf OFF — LIVE mode" -Level WARN -Source 'DnsManager'
    } else {
        Set-WhatIfMode -Mode $true
        $Script:DNS['DNS_WhatIfBadge'].Background = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0x27, 0xAE, 0x60)
        $Script:DNS['DNS_WhatIfLabel'].Text = "ON"
        if ($Script:DNS['DNS_WhatIfToggle']) { $Script:DNS['DNS_WhatIfToggle'].IsChecked = $true }
        Write-Log "DnsManager: WhatIf ON" -Source 'DnsManager'
    }
})

# ── LOAD ZONES button — background zone load ──────────────────────────────────
$Script:DNS['DNS_BtnLoadZones'].Add_Click({
    if (-not (Get-AppConnected)) {
        [System.Windows.MessageBox]::Show("Connect to a DC first.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    $dc      = Get-AppDCName
    $cred    = Get-AppCredential
    $window  = $Script:Window
    $dns     = $Script:DNS
    $fnPath  = $Script:DNS_FnPath
    $logFile = $Script:LogFile

    $dns['DNS_BtnLoadZones'].IsEnabled = $false
    & $Script:DNS_ShowProgress -Value 10 -Label "Loading zones..."
    & $Script:DNS_SetStatus "Loading zones from $dc..."

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $rs  = [RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()

    foreach ($v in @('dc','cred','window','dns','fnPath','logFile')) {
        $rs.SessionStateProxy.SetVariable($v, (Get-Variable $v -ValueOnly))
    }

    $ps = [PowerShell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript({
        Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
        function Write-Log { param($m, $l = 'INFO', $s = 'Log')
            try { if ($logFile) { Add-Content $logFile "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][$s] $m" -Encoding UTF8 } } catch {}
        }
        . ([string]$fnPath)

        try {
            $icParams = @{
                ComputerName = $dc
                ErrorAction  = 'Stop'
                ScriptBlock  = {
                    Get-DnsServerZone -ErrorAction Stop |
                        Where-Object {
                            $_.ZoneName -notmatch '\.in-addr\.arpa$' -and
                            $_.ZoneName -notmatch '\.ip6\.arpa$' -and
                            $_.IsReverseLookupZone -ne $true
                        } |
                        Select-Object -ExpandProperty ZoneName
                }
            }
            if ($null -ne $cred) { $icParams['Credential'] = $cred }

            $names = @(Invoke-Command @icParams)

            $dns['__ZoneNames'] = $names

            # Build the collection outside Dispatcher.Invoke — New-Object cannot run inside nested invoke
            $zoneList = New-Object System.Collections.ObjectModel.ObservableCollection[string]
            foreach ($z in $names) { $zoneList.Add([string]$z) }
            $dns['__ZoneList'] = $zoneList

            $window.Dispatcher.Invoke([action]{
                $dns['DNS_CboZone'].ItemsSource = $dns['__ZoneList']
                $dns['DNS_CboZone'].IsEnabled = $true
                $dns['DNS_BtnLoadZones'].IsEnabled = $true
                $dns['DNS_StatusBar'].Text         = "Loaded $($dns['__ZoneList'].Count) zone(s) — select one and click Fetch"
                $dns['DNS_Progress'].Visibility      = 'Collapsed'
                $dns['DNS_ProgressLabel'].Visibility = 'Collapsed'
                $dns['__ZoneDone'] = $true
            })
        } catch {
            $err = $_.Exception.Message
            $window.Dispatcher.Invoke([action]{
                $dns['DNS_BtnLoadZones'].IsEnabled   = $true
                $dns['DNS_StatusBar'].Text           = "Zone load failed: $err"
                $dns['DNS_Progress'].Visibility      = 'Collapsed'
                $dns['DNS_ProgressLabel'].Visibility = 'Collapsed'
                $dns['__ZoneDone'] = $true
            })
        }
    })

    $handle = $ps.BeginInvoke()

    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(300)
    $t.Tag = [PSCustomObject]@{ PS = $ps; RS = $rs; Handle = $handle }
    $t.Add_Tick({
        $state = $this.Tag
        if ($state.Handle.IsCompleted -or $Script:DNS.ContainsKey('__ZoneDone')) {
            $this.Stop()
            try { $state.PS.EndInvoke($state.Handle) } catch {}
            $state.PS.Dispose(); $state.RS.Close(); $state.RS.Dispose()
            $Script:DNS.Remove('__ZoneDone')
            $Script:DNS.Remove('__ZoneNames')
            $Script:DNS.Remove('__ZoneList')
            Write-Log "DnsManager: zone load complete" -Source 'DnsManager'
        }
    })
    $t.Start()
})

# ── FETCH RECORDS button — background record fetch ────────────────────────────
$Script:DNS['DNS_BtnFetch'].Add_Click({
    # Guard: zone must be selected
    if ($null -eq $Script:DNS['DNS_CboZone'].SelectedItem) {
        [System.Windows.MessageBox]::Show("Select a zone first.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    # Guard: input must not be blank
    $rawInput = $Script:DNS['DNS_TxtInput'].Text.Trim()
    if ([string]::IsNullOrWhiteSpace($rawInput)) {
        [System.Windows.MessageBox]::Show("Enter at least one hostname in the input box.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    $zone    = $Script:DNS['DNS_CboZone'].SelectedItem
    $dc      = Get-AppDCName
    $cred    = Get-AppCredential
    $wi      = Get-WhatIfMode
    $window  = $Script:Window
    $dns     = $Script:DNS
    $allRows = $Script:DNS_AllRows
    $fnPath  = $Script:DNS_FnPath
    $logFile = $Script:LogFile

    # ── Parse input lines on the UI thread before spawning the runspace ───────
    . ([string]$Script:DNS_FnPath)

    $Script:DNS_AllRows.Clear()
    $parsedLines = [System.Collections.Generic.List[hashtable]]::new()
    $hostnames   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($line in ($rawInput -split "`n")) {
        $parsed = DNS-ParseInputLine -Line $line -ZoneName $zone
        if (-not $parsed.OK) {
            # Blank/comment lines return OK=$false with empty ErrorMsg — skip silently
            if ($parsed.ErrorMsg -ne '') {
                $Script:DNS_AllRows.Add([PSCustomObject]@{
                    Status        = 'Failed'
                    StatusLabel   = '❌ Parse Error'
                    Name          = $line.Trim()
                    Type          = ''
                    ExistingValue = ''
                    NewValue      = ''
                    TTL           = ''
                    Match         = ''
                    ResultNote    = $parsed.ErrorMsg
                    IsWildcard    = $false
                    _Zone         = $zone
                    _MultiAction  = ''
                })
            }
            continue
        }
        [void]$parsedLines.Add($parsed)
        [void]$hostnames.Add($parsed.Hostname)
    }

    if ($hostnames.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No valid hostnames found in the input box.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    $hostnameList = @($hostnames)
    $parsedMap    = @{}
    foreach ($p in $parsedLines) { $parsedMap[$p.Hostname] = $p }

    $dns['DNS_BtnFetch'].IsEnabled = $false
    $dns['DNS_BtnApply'].IsEnabled = $false
    & $Script:DNS_ShowProgress -Value 5 -Label "Fetching records..."
    & $Script:DNS_SetStatus "Fetching DNS records for $($hostnameList.Count) hostname(s)..."

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $rs  = [RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()

    foreach ($v in @('dc','cred','wi','zone','window','dns','allRows','fnPath','logFile','hostnameList','parsedMap')) {
        $rs.SessionStateProxy.SetVariable($v, (Get-Variable $v -ValueOnly))
    }

    $ps = [PowerShell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript({
        Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
        function Write-Log { param($m, $l = 'INFO', $s = 'Log')
            try { if ($logFile) { Add-Content $logFile "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][$s] $m" -Encoding UTF8 } } catch {}
        }
        . ([string]$fnPath)

        # Fetch all records for the distinct hostname list
        $fetchResults = DNS-FetchRecords -Hostnames $hostnameList -Zone $zone -DC $dc -Credential $cred

        # Group results by hostname
        $byHost = @{}
        foreach ($rec in $fetchResults) {
            $h = $rec.Hostname
            if (-not $byHost.ContainsKey($h)) { $byHost[$h] = [System.Collections.Generic.List[object]]::new() }
            [void]$byHost[$h].Add($rec)
        }

        $localBatch = [System.Collections.Generic.List[object]]::new()

        # ── Process exact-match hostnames from the input ──────────────────────
        foreach ($h in $hostnameList) {
            $parsed    = $parsedMap[$h]
            $existing  = if ($byHost.ContainsKey($h)) { @($byHost[$h] | Where-Object { -not $_.IsWildcard }) } else { @() }

            # Check for fetch error
            $errRec = $existing | Where-Object { $_.Type -eq 'Error' } | Select-Object -First 1
            if ($errRec) {
                $row = [PSCustomObject]@{
                    Status        = 'Failed'
                    StatusLabel   = '❌ Failed'
                    Name          = $h
                    Type          = $parsed.Type
                    ExistingValue = ''
                    NewValue      = $parsed.Value
                    TTL           = ''
                    Match         = 'Exact'
                    ResultNote    = $errRec.ExistingValue
                    IsWildcard    = $false
                    _Zone         = $zone
                    _MultiAction  = ''
                }
                [void]$localBatch.Add($row)
                continue
            }

            $parsed    = $parsedMap[$h]
            $exactMatches = @($fetchResults | Where-Object { $_.Hostname -eq $h -and -not $_.IsWildcard })
            $status = DNS-ClassifyRow -ParsedRow $parsed -ExistingRecs $exactMatches

            # MultiConflict: ask user how to resolve — must be done on UI thread
            $multiAction = ''
            if ($status -eq 'MultiConflict') {
                $hCapture = $h
                $window.Dispatcher.Invoke([action]{
                    $choice = [Windows.MessageBox]::Show(
                        "Multiple A records exist for '$hCapture'.`nChoose an action:",
                        "DNS Conflict", 'YesNoCancel', 'Question')
                    $Script:DNS_MultiChoice = switch ($choice) {
                        'Yes'    { 'ReplaceAll' }
                        'No'     { 'AddAdditional' }
                        default  { 'NoAction' }
                    }
                })
                $multiAction = $Script:DNS_MultiChoice
            }

            $existingVal = ''
            $ttlVal      = ''
            $typeVal     = if ($parsed.Type) { $parsed.Type } else { '' }
            
            if ($exactMatches.Count -gt 0) {
                $firstExact = $exactMatches[0]
                $existingVal = $firstExact.ExistingValue
                $ttlVal      = $firstExact.TTL
                if (-not $typeVal) {
                    $typeVal = $firstExact.Type
                }
            }

            $statusLabel = switch ($status) {
                'New'           { '➕ New' }
                'NoAction'      { '✅ No Action' }
                'Update'        { '🔄 Update' }
                'Convert'       { '🔁 Convert' }
                'MultiConflict' { '⚠ Multi-Record' }
                'NotFound'      { '❌ Not Found' }
                default         { $status }
            }

            $row = [PSCustomObject]@{
                Status        = $status
                StatusLabel   = $statusLabel
                Name          = $h
                Type          = $typeVal
                ExistingValue = $existingVal
                NewValue      = $parsed.Value
                TTL           = $ttlVal
                Match         = 'Exact'
                ResultNote    = ''
                IsWildcard    = $false
                _Zone         = $zone
                _MultiAction  = $multiAction
            }
            [void]$localBatch.Add($row)
        }

        # ── Process wildcard rows ─────────────────────────────────────────────
        $allWildcards = @($fetchResults | Where-Object { $_.IsWildcard -eq $true })
        foreach ($wc in $allWildcards) {
            $row = [PSCustomObject]@{
                Status        = 'WildcardMatch'
                StatusLabel   = '❓ Wildcard Match'
                Name          = $wc.Hostname
                Type          = $wc.Type
                ExistingValue = $wc.ExistingValue
                NewValue      = ''
                TTL           = $wc.TTL
                Match         = 'Wildcard'
                ResultNote    = ''
                IsWildcard    = $true
                _Zone         = $zone
                _MultiAction  = ''
            }
            [void]$localBatch.Add($row)
        }

        # ── Batch-push all rows to the UI ─────────────────────────────────────
        $batch = $localBatch.ToArray()
        $window.Dispatcher.Invoke([action]{
            foreach ($r in $batch) { $allRows.Add($r) }
            $dns['DNS_BtnFetch'].IsEnabled = $true
            $dns['DNS_BtnApply'].IsEnabled = $true
            $dns['DNS_Progress'].Visibility      = 'Collapsed'
            $dns['DNS_ProgressLabel'].Visibility = 'Collapsed'
            $dns['DNS_StatusBar'].Text           = "Fetched — $($allRows.Count) row(s)"
            $dns['__FetchDone'] = $true
        })
    })

    $handle = $ps.BeginInvoke()

    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(300)
    $t.Tag = [PSCustomObject]@{ PS = $ps; RS = $rs; Handle = $handle }
    $t.Add_Tick({
        $state = $this.Tag
        if ($state.Handle.IsCompleted -or $Script:DNS.ContainsKey('__FetchDone')) {
            $this.Stop()
            try { $state.PS.EndInvoke($state.Handle) } catch {
                $errMsg = $_.Exception.Message
                $Script:DNS['DNS_BtnFetch'].IsEnabled = $true
                $Script:DNS['DNS_BtnApply'].IsEnabled = $true
                & $Script:DNS_ShowProgress -Value -1
                & $Script:DNS_SetStatus "Fetch failed: $errMsg"
            }
            $state.PS.Dispose(); $state.RS.Close(); $state.RS.Dispose()
            $Script:DNS.Remove('__FetchDone')
            Write-Log "DnsManager: fetch complete — $($Script:DNS_AllRows.Count) row(s)" -Source 'DnsManager'
        }
    })
    $t.Start()
})

# ── APPLY button — parallel apply with DispatcherTimer ────────────────────────
$Script:DNS['DNS_BtnApply'].Add_Click({
    if (-not (Get-AppConnected)) {
        [System.Windows.MessageBox]::Show("Connect to a DC first.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    # Build apply queue: rows that are actionable (not NoAction, WildcardMatch, Done, or Failed)
    $queue = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $Script:DNS_AllRows) {
        if ($r.Status -notin @('NoAction', 'WildcardMatch', 'Done', 'Failed')) {
            [void]$queue.Add($r)
        }
    }

    # Guard: nothing to do
    if ($queue.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No actionable rows.", "AD Manager", 'OK', 'Information') | Out-Null
        return
    }

    # Confirmation with current WhatIf mode shown
    $wi      = Get-WhatIfMode
    $modeStr = if ($wi) { 'WhatIf PREVIEW — no DNS changes will be written' } else { 'LIVE — changes WILL be written to DNS' }
    $confirm = [System.Windows.MessageBox]::Show(
        "Apply $($queue.Count) row(s)?`n`nMode: $modeStr",
        "Confirm Apply", 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }

    # Initialise apply state
    $Script:DNS_ApplyQueue   = $queue
    $Script:DNS_ApplyIdx     = 0
    $Script:DNS_ApplyOK      = 0
    $Script:DNS_ApplyFailed  = 0
    $Script:DNS_ApplyTotal   = $queue.Count
    $Script:DNS_ApplyDC      = Get-AppDCName
    $Script:DNS_ApplyCred    = Get-AppCredential
    $Script:DNS_ApplyWhatIf  = $wi
    $Script:DNS_ApplyThreads = [System.Collections.Generic.List[object]]::new()

    # Disable action buttons while running
    $Script:DNS['DNS_BtnApply'].IsEnabled  = $false
    $Script:DNS['DNS_BtnFetch'].IsEnabled  = $false
    $Script:DNS['DNS_BtnDelete'].IsEnabled = $false
    & $Script:DNS_ShowProgress -Value 0 -Label "Starting apply..."
    & $Script:DNS_SetStatus "Applying $($Script:DNS_ApplyTotal) row(s)..."

    Write-Log "DnsManager: apply started — $($Script:DNS_ApplyTotal) rows, WhatIf=$wi" -Source 'DnsManager'

    $applyTimer = New-Object System.Windows.Threading.DispatcherTimer
    $applyTimer.Interval = [TimeSpan]::FromMilliseconds(50)
    $applyTimer.Add_Tick({

        # ── Harvest completed runspaces ───────────────────────────────────────
        for ($i = $Script:DNS_ApplyThreads.Count - 1; $i -ge 0; $i--) {
            $thr = $Script:DNS_ApplyThreads[$i]
            if (-not $thr.Handle.IsCompleted) { continue }

            $res = $null
            try {
                $out = $thr.PS.EndInvoke($thr.Handle)
                $res = $out | Where-Object { $_ -is [hashtable] -and $null -ne $_.OK } | Select-Object -Last 1
            } catch {}
            $thr.PS.Dispose()
            if ($thr.RS) { $thr.RS.Close(); $thr.RS.Dispose() }

            if ($res -and $res.OK) {
                $Script:DNS_ApplyOK++
                $thr.Row.Status      = 'Done'
                $thr.Row.StatusLabel = '✅ Done'
                $thr.Row.ResultNote  = $res.Message
            } else {
                $Script:DNS_ApplyFailed++
                $thr.Row.Status      = 'Failed'
                $thr.Row.StatusLabel = '❌ Failed'
                $thr.Row.ResultNote  = if ($res) { $res.Message } else { 'Apply failed' }
            }

            # Refresh row in-place via RemoveAt/Insert on ObservableCollection
            $idx = $Script:DNS_AllRows.IndexOf($thr.Row)
            if ($idx -ge 0) {
                $Script:DNS_AllRows.RemoveAt($idx)
                $Script:DNS_AllRows.Insert($idx, $thr.Row)
            }
            $Script:DNS_ApplyThreads.RemoveAt($i)

            $done = $Script:DNS_ApplyOK + $Script:DNS_ApplyFailed
            & $Script:DNS_ShowProgress -Value ([int](($done / $Script:DNS_ApplyTotal) * 100)) `
                -Label "Applying $done / $Script:DNS_ApplyTotal"
        }

        # ── Spawn new runspaces (up to 5 concurrent) ──────────────────────────
        while ($Script:DNS_ApplyThreads.Count -lt 5 -and $Script:DNS_ApplyIdx -lt $Script:DNS_ApplyTotal) {
            $row  = $Script:DNS_ApplyQueue[$Script:DNS_ApplyIdx++]
            $tRow = $row
            $tDC  = $Script:DNS_ApplyDC
            $tCr  = $Script:DNS_ApplyCred
            $tWI  = $Script:DNS_ApplyWhatIf
            $tFn  = $Script:DNS_FnPath
            $tLog = $Script:LogFile

            $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
            $rs  = [RunspaceFactory]::CreateRunspace($iss)
            $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()

            $psA = [PowerShell]::Create(); $psA.Runspace = $rs
            [void]$psA.AddScript({
                param($tRow, $tDC, $tCr, $tWI, $tFn, $tLog)
                function Write-Log { param($m, $l = 'INFO', $s = 'Log')
                    try { if ($tLog) { Add-Content $tLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][$s] $m" -Encoding UTF8 } } catch {}
                }
                Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
                if (Test-Path ([string]$tFn)) { . ([string]$tFn) }
                DNS-ApplyRecord -Row $tRow -DC $tDC -Credential $tCr -WhatIf $tWI
            })
            $psA.AddParameter('tRow', $tRow)
            $psA.AddParameter('tDC',  $tDC)
            $psA.AddParameter('tCr',  $tCr)
            $psA.AddParameter('tWI',  $tWI)
            $psA.AddParameter('tFn',  $tFn)
            $psA.AddParameter('tLog', $tLog)

            $handle = $psA.BeginInvoke()
            [void]$Script:DNS_ApplyThreads.Add(@{ PS = $psA; RS = $rs; Handle = $handle; Row = $row })
        }

        # ── Stop when queue exhausted and all runspaces complete ──────────────
        if ($Script:DNS_ApplyIdx -ge $Script:DNS_ApplyTotal -and $Script:DNS_ApplyThreads.Count -eq 0) {
            $this.Stop()
            $Script:DNS['DNS_BtnApply'].IsEnabled  = $true
            $Script:DNS['DNS_BtnFetch'].IsEnabled  = $true
            $Script:DNS['DNS_BtnDelete'].IsEnabled = $true
            & $Script:DNS_ShowProgress -Value -1
            $summary = "Applied: $Script:DNS_ApplyOK done, $Script:DNS_ApplyFailed failed"
            & $Script:DNS_SetStatus $summary
            Write-Log "DnsManager: apply complete — OK=$Script:DNS_ApplyOK FAIL=$Script:DNS_ApplyFailed" -Source 'DnsManager'
        }
    })
    $applyTimer.Start()
})

# ── DELETE SELECTED button — parallel delete with DispatcherTimer ─────────────
$Script:DNS['DNS_BtnDelete'].Add_Click({
    if (-not (Get-AppConnected)) {
        [System.Windows.MessageBox]::Show("Connect to a DC first.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    # Collect selected rows; filter out wildcard rows (IsWildcard=$true)
    $selectedRows = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $Script:DNS['DNS_Grid'].SelectedItems) {
        if ($item.IsWildcard -ne $true) {
            [void]$selectedRows.Add($item)
        }
    }

    # Guard: nothing selected after filtering wildcards
    if ($selectedRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Select one or more records to delete.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    # Confirmation dialog
    $confirm = [System.Windows.MessageBox]::Show(
        "Delete $($selectedRows.Count) record(s)?`n`nThis cannot be undone.",
        "Confirm Delete", 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }

    # Initialise delete state
    $Script:DNS_DeleteQueue   = $selectedRows
    $Script:DNS_DeleteIdx     = 0
    $Script:DNS_DeleteOK      = 0
    $Script:DNS_DeleteFailed  = 0
    $Script:DNS_DeleteTotal   = $selectedRows.Count
    $Script:DNS_DeleteDC      = Get-AppDCName
    $Script:DNS_DeleteCred    = Get-AppCredential
    $Script:DNS_DeleteWhatIf  = Get-WhatIfMode
    $Script:DNS_DeleteThreads = [System.Collections.Generic.List[object]]::new()

    # Disable action buttons while running
    $Script:DNS['DNS_BtnDelete'].IsEnabled = $false
    $Script:DNS['DNS_BtnApply'].IsEnabled  = $false
    $Script:DNS['DNS_BtnFetch'].IsEnabled  = $false
    & $Script:DNS_ShowProgress -Value 0 -Label "Starting delete..."
    & $Script:DNS_SetStatus "Deleting $($Script:DNS_DeleteTotal) record(s)..."

    $wi = $Script:DNS_DeleteWhatIf
    Write-Log "DnsManager: delete started — $($Script:DNS_DeleteTotal) rows, WhatIf=$wi" -Source 'DnsManager'

    $deleteTimer = New-Object System.Windows.Threading.DispatcherTimer
    $deleteTimer.Interval = [TimeSpan]::FromMilliseconds(50)
    $deleteTimer.Add_Tick({

        # ── Harvest completed runspaces ───────────────────────────────────────
        for ($i = $Script:DNS_DeleteThreads.Count - 1; $i -ge 0; $i--) {
            $thr = $Script:DNS_DeleteThreads[$i]
            if (-not $thr.Handle.IsCompleted) { continue }

            $res = $null
            try {
                $out = $thr.PS.EndInvoke($thr.Handle)
                $res = $out | Where-Object { $_ -is [hashtable] -and $null -ne $_.OK } | Select-Object -Last 1
            } catch {}
            $thr.PS.Dispose()
            if ($thr.RS) { $thr.RS.Close(); $thr.RS.Dispose() }

            if ($res -and $res.OK) {
                $Script:DNS_DeleteOK++
                $thr.Row.Status      = 'Done'
                $thr.Row.StatusLabel = '✅ Done'
                $thr.Row.ResultNote  = $res.Message
            } else {
                $Script:DNS_DeleteFailed++
                $thr.Row.Status      = 'Failed'
                $thr.Row.StatusLabel = '❌ Failed'
                $thr.Row.ResultNote  = if ($res) { $res.Message } else { 'Delete failed' }
            }

            # Refresh row in-place via RemoveAt/Insert on ObservableCollection
            $idx = $Script:DNS_AllRows.IndexOf($thr.Row)
            if ($idx -ge 0) {
                $Script:DNS_AllRows.RemoveAt($idx)
                $Script:DNS_AllRows.Insert($idx, $thr.Row)
            }
            $Script:DNS_DeleteThreads.RemoveAt($i)

            $done = $Script:DNS_DeleteOK + $Script:DNS_DeleteFailed
            & $Script:DNS_ShowProgress -Value ([int](($done / $Script:DNS_DeleteTotal) * 100)) `
                -Label "Deleting $done / $Script:DNS_DeleteTotal"
        }

        # ── Spawn new runspaces (up to 5 concurrent) ──────────────────────────
        while ($Script:DNS_DeleteThreads.Count -lt 5 -and $Script:DNS_DeleteIdx -lt $Script:DNS_DeleteTotal) {
            $row  = $Script:DNS_DeleteQueue[$Script:DNS_DeleteIdx++]
            $tRow = $row
            $tDC  = $Script:DNS_DeleteDC
            $tCr  = $Script:DNS_DeleteCred
            $tWI  = $Script:DNS_DeleteWhatIf
            $tFn  = $Script:DNS_FnPath
            $tLog = $Script:LogFile

            $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
            $rs  = [RunspaceFactory]::CreateRunspace($iss)
            $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()

            $psD = [PowerShell]::Create(); $psD.Runspace = $rs
            [void]$psD.AddScript({
                param($tRow, $tDC, $tCr, $tWI, $tFn, $tLog)
                function Write-Log { param($m, $l = 'INFO', $s = 'Log')
                    try { if ($tLog) { Add-Content $tLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][$s] $m" -Encoding UTF8 } } catch {}
                }
                Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
                if (Test-Path ([string]$tFn)) { . ([string]$tFn) }
                DNS-DeleteRecord -Row $tRow -DC $tDC -Credential $tCr -WhatIf $tWI
            })
            $psD.AddParameter('tRow', $tRow)
            $psD.AddParameter('tDC',  $tDC)
            $psD.AddParameter('tCr',  $tCr)
            $psD.AddParameter('tWI',  $tWI)
            $psD.AddParameter('tFn',  $tFn)
            $psD.AddParameter('tLog', $tLog)

            $handle = $psD.BeginInvoke()
            [void]$Script:DNS_DeleteThreads.Add(@{ PS = $psD; RS = $rs; Handle = $handle; Row = $row })
        }

        # ── Stop when queue exhausted and all runspaces complete ──────────────
        if ($Script:DNS_DeleteIdx -ge $Script:DNS_DeleteTotal -and $Script:DNS_DeleteThreads.Count -eq 0) {
            $this.Stop()
            $Script:DNS['DNS_BtnDelete'].IsEnabled = $true
            $Script:DNS['DNS_BtnApply'].IsEnabled  = $true
            $Script:DNS['DNS_BtnFetch'].IsEnabled  = $true
            & $Script:DNS_ShowProgress -Value -1
            $summary = "Deleted: $Script:DNS_DeleteOK done, $Script:DNS_DeleteFailed failed"
            & $Script:DNS_SetStatus $summary
            Write-Log "DnsManager: delete complete — OK=$Script:DNS_DeleteOK FAIL=$Script:DNS_DeleteFailed" -Source 'DnsManager'
        }
    })
    $deleteTimer.Start()
})

# ── EXPORT button — export grid to XLSX ──────────────────────────────────────
$Script:DNS['DNS_BtnExport'].Add_Click({
    # Guard: nothing to export
    if ($Script:DNS_AllRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No data to export.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    # Show Save File dialog
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Title  = "Export DNS Records"
    $sfd.Filter = "Excel files (*.xlsx)|*.xlsx"
    $sfd.FileName = "DNS_Export_$((Get-Date -Format 'yyyyMMdd_HHmm')).xlsx"
    if ($sfd.ShowDialog() -ne $true) { return }
    $path = $sfd.FileName

    # Dot-source Functions.ps1 to ensure DNS-ExportToExcel is available
    . ([string]$Script:DNS_FnPath)

    # Call export function
    $result = DNS-ExportToExcel -Rows $Script:DNS_AllRows.ToArray() -FilePath $path

    if ($result.OK -eq $false) {
        & $Script:DNS_SetStatus "Export failed: $($result.ErrorMsg)"
        Write-Log "DnsManager: export failed — $($result.ErrorMsg)" -Level WARN -Source 'DnsManager'
    } else {
        $fileName = [System.IO.Path]::GetFileName($path)
        $count    = $Script:DNS_AllRows.Count
        & $Script:DNS_SetStatus "Exported $count rows to $fileName"
        Write-Log "DnsManager: exported $count rows to $path" -Source 'DnsManager'
    }
})

# ── Grid selection changed — update count and enable/disable Delete ───────────
$Script:DNS['DNS_Grid'].Add_SelectionChanged({
    & $Script:DNS_UpdateSelCount
    $nonWildcard = $Script:DNS['DNS_Grid'].SelectedItems | Where-Object { $_.IsWildcard -ne $true }
    $Script:DNS['DNS_BtnDelete'].IsEnabled = ($null -ne $nonWildcard -and @($nonWildcard).Count -gt 0)
})

# ── Zone combobox selection changed — enable/disable Fetch and input ──────────
$Script:DNS['DNS_CboZone'].Add_SelectionChanged({
    $zoneSelected = $null -ne $Script:DNS['DNS_CboZone'].SelectedItem
    $Script:DNS['DNS_BtnFetch'].IsEnabled  = $zoneSelected
    $Script:DNS['DNS_TxtInput'].IsEnabled  = $zoneSelected
})

Write-Log "DnsManager handlers registered" -Source 'DnsManager'

# ── Apply current connection state immediately (handles reload-while-connected) ──
if (Get-AppConnected) {
    $Script:DNS['DNS_BtnLoadZones'].IsEnabled = $true
    $Script:DNS['DNS_TxtInput'].IsEnabled     = $true
    & $Script:DNS_SetStatus "Select a zone and load records"
} else {
    & $Script:DNS_SetStatus "Connect to a DC to begin"
}
