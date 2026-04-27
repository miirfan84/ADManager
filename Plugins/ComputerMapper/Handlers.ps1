# =============================================================================
# Plugins\ComputerMapper\Handlers.ps1  —  AD Manager v2.0
# =============================================================================

Write-Log "ComputerMapper: registering handlers..." -Source 'ComputerMapper'

# ── Path to Functions.ps1 (used for dot-sourcing in click handlers + runspaces)
$Script:CM_FnPath = Join-Path $PSScriptRoot "Functions.ps1"

# ── Grab controls via PluginRoot (NOT $Script:Window — controls are inside the tab)
$Script:CM = @{}
foreach ($n in @(
    'CM_BtnImport','CM_BtnClearAll','CM_BtnFetchAD','CM_BtnDeepScan','CM_TxtFile',
    'CM_CntTotal','CM_CntOK','CM_CntPending','CM_CntWarn','CM_CntErr','CM_CntExtra',
    'CM_WhatIfBadge','CM_WhatIfLabel','CM_WhatIfToggle',
    'CM_Progress','CM_ProgressLabel',
    'CM_FAll','CM_FOK','CM_FPending','CM_FWarn','CM_FNotFound',
    'CM_FNoUser','CM_FMulti','CM_FUpdated','CM_FFailed','CM_FExtra',
    'CM_TxtSearch','CM_BtnSearch','CM_BtnClearSearch',
    'CM_Grid','CM_BtnRun','CM_BtnStop','CM_StatusBar','CM_SelCount','CM_BtnExport'
)) {
    $ctrl = $Script:PluginRoot.FindName($n)
    if ($null -ne $ctrl) {
        $Script:CM[$n] = $ctrl
        Write-Log "  Control: $n" -Source 'ComputerMapper'
    } else {
        Write-Log "  MISSING: $n" -Level WARN -Source 'ComputerMapper'
    }
}

# ── State
$Script:CM_AllRows      = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:CM_UserCompMap  = @{}
$Script:CM_ActiveFilter = 'All'
$Script:CM_SearchTerm   = ''
$Script:CM_LastImport   = $null   # saved import result — used by Fetch AD button
$Script:CM_ADComputers  = $null   # saved AD computers hashtable — reused by Deep Scan
$Script:CM_IsRunning    = $false

$Script:CM_FilterMap = @{
    'CM_FAll'='All'; 'CM_FOK'='NoAction'; 'CM_FPending'='Pending'
    'CM_FWarn'='Warning'; 'CM_FNotFound'='NotFound'; 'CM_FNoUser'='NoUser'
    'CM_FMulti'='Multi'; 'CM_FUpdated'='Updated'; 'CM_FFailed'='Failed'
    'CM_FExtra'='InADOnly'
}

# Bind grid — deferred to after window is loaded to avoid timing issues
$Script:CM['CM_Grid'].ItemsSource = $Script:CM_AllRows
$Script:CM['CM_Grid'].Items.Refresh()

# =============================================================================
# HELPER SCRIPTBLOCKS (Script-scoped so visible inside all event closures)
# =============================================================================

$Script:CM_SetStatus = {
    param([string]$Text)
    $Script:CM['CM_StatusBar'].Text = $Text
}

$Script:CM_ShowProgress = {
    param([int]$Value = -1, [string]$Label = '')
    if ($Value -lt 0) {
        $Script:CM['CM_Progress'].Visibility      = 'Collapsed'
        $Script:CM['CM_ProgressLabel'].Visibility = 'Collapsed'
    } else {
        $Script:CM['CM_Progress'].Visibility      = 'Visible'
        $Script:CM['CM_ProgressLabel'].Visibility = 'Visible'
        $Script:CM['CM_Progress'].Value           = [Math]::Min($Value, 100)
        if ($Label -ne '') { $Script:CM['CM_ProgressLabel'].Text = $Label }
    }
}

$Script:CM_UpdateCounters = {
    $t=0;$ok=0;$p=0;$w=0;$e=0;$x=0
    foreach ($r in $Script:CM_AllRows) {
        $t++
        switch ($r.Status) {
            'NoAction'  { $ok++ } 'Updated'    { $ok++ }
            'WouldUpdate'{ $p++ } 'Pending'    { $p++  }
            'Warning'   { $w++  }
            'NotFound'  { $e++  } 'NoUser'     { $e++  } 'Failed' { $e++ }
            'InADOnly'  { $x++  } 'InADStale'  { $x++  }
        }
    }
    $Script:CM['CM_CntTotal'].Text   = $t
    $Script:CM['CM_CntOK'].Text      = $ok
    $Script:CM['CM_CntPending'].Text = $p
    $Script:CM['CM_CntWarn'].Text    = $w
    $Script:CM['CM_CntErr'].Text     = $e
    if ($Script:CM['CM_CntExtra']) { $Script:CM['CM_CntExtra'].Text = $x }
}

$Script:CM_ApplyFilter = {
    $filter = $Script:CM_ActiveFilter
    $term   = $Script:CM_SearchTerm.ToLower()
    $view   = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:CM_AllRows)
    $view.Filter = {
        param($item)
        $ok = switch ($filter) {
            'All'      { $true; break }
            'Multi'    { ($item.IsMulti -eq $true); break }
            'Updated'  { ($item.Status -in @('Updated','WouldUpdate')); break }
            'InADOnly' { ($item.Status -in @('InADOnly','InADStale')); break }
            default    { ($item.Status -eq $filter) }
        }
        if (-not $ok) { return $false }
        if ([string]::IsNullOrWhiteSpace($term)) { return $true }
        $hay = "$($item.TAG) $($item.ComputerName) $($item.EmpID) $($item.DisplayName) $($item.SAM) $($item.Office) $($item.CurrentOU) $($item.ResultNote)".ToLower()
        return $hay.Contains($term)
    }
    $view.Refresh()
}

# =============================================================================
# BULK AD FETCH + MATCH  (runs only when user clicks Fetch AD Data)
# =============================================================================

$Script:CM_StartBulkAndMatch = {
    param($ImportData)
    if ($null -eq $ImportData) { return }

    $dc      = Get-AppDCName
    $cred    = Get-AppCredential
    $window  = $Script:Window
    $allRows = $Script:CM_AllRows
    $cm      = $Script:CM
    $logFile = $Script:LogFile
    $fnPath  = $Script:CM_FnPath
    $ucMap   = $Script:CM_UserCompMap

    $cm['CM_BtnFetchAD'].IsEnabled = $false
    $cm['CM_BtnImport'].IsEnabled  = $false
    $cm['CM_BtnRun'].IsEnabled     = $false
    $cm['CM_BtnExport'].IsEnabled  = $false

    & $Script:CM_ShowProgress -Value 5 -Label "Connecting to AD..."
    & $Script:CM_SetStatus "Fetching AD data..."

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $rs  = [RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()

    foreach ($v in @('dc','cred','window','allRows','cm','logFile','fnPath','ucMap')) {
        $rs.SessionStateProxy.SetVariable($v, (Get-Variable $v -ValueOnly))
    }
    $rs.SessionStateProxy.SetVariable('rows',     $ImportData.Rows)
    $rs.SessionStateProxy.SetVariable('tagCol',   $ImportData.TagCol)
    $rs.SessionStateProxy.SetVariable('empCol',   $ImportData.EmpCol)
    $rs.SessionStateProxy.SetVariable('mailCol',  $ImportData.EmailCol)
    $rs.SessionStateProxy.SetVariable('stsCol',   $ImportData.StatusCol)
    $rs.SessionStateProxy.SetVariable('total',    $ImportData.TotalRows)
    $rs.SessionStateProxy.SetVariable('CM_Status',      $Script:CM_Status)
    $rs.SessionStateProxy.SetVariable('CM_StatusLabel', $Script:CM_StatusLabel)

    $ps = [PowerShell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript({
        Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
        function Write-Log { param($m,$l='INFO',$s='Log')
            try { if ($LF) { Add-Content $LF "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][$s] $m" -Encoding UTF8 } } catch {}
        }
        . ([string]$fnPath)
        $Script:CM_Status      = $CM_Status
        $Script:CM_StatusLabel = $CM_StatusLabel

        $window.Dispatcher.Invoke([action]{
            $cm['CM_ProgressLabel'].Text       = 'Fetching AD data...'
            $cm['CM_ProgressLabel'].Visibility = 'Visible'
            $cm['CM_Progress'].Visibility      = 'Visible'
            $cm['CM_Progress'].Value           = 10
        })

        $adResult = CM-BulkFetchAD -DCName $dc -Credential $cred -ProgressCallback {
            param($s,$m)
            $msg = $m
            $window.Dispatcher.Invoke([action]{ $cm['CM_ProgressLabel'].Text = $msg })
        }

        if (-not $adResult.OK) {
            $err = $adResult.ErrorMsg
            $window.Dispatcher.Invoke([action]{
                $cm['CM_StatusBar'].Text       = "AD fetch failed: $err"
                $cm['CM_BtnImport'].IsEnabled  = $true
                $cm['CM_BtnFetchAD'].IsEnabled = $true
                $cm['CM_Progress'].Visibility  = 'Collapsed'
                $cm['CM_ProgressLabel'].Visibility = 'Collapsed'
            })
            return
        }

        $computers   = $adResult.Computers
        $users       = $adResult.Users
        $usersDup    = $adResult.UsersDup
        $usersByMail = $adResult.UsersByMail
        $usersBySAM  = $adResult.UsersBySAM

        # Store computers in script scope so Deep Scan can reuse without a second AD query
        # Use $cm hashtable as the bridge — it's a shared reference accessible on the main thread
        $window.Dispatcher.Invoke([action]{ $cm['__ADComputers'] = $computers })

        Write-Log "AD fetch done. Matching $total rows..." 'INFO' 'CM'

        # Clear the placeholder rows, then stream real results
        $window.Dispatcher.Invoke([action]{ $allRows.Clear() })

        $localBatch = [System.Collections.Generic.List[object]]::new()
        $cur = 0

        foreach ($csvRow in $rows) {
            $cur++
            $tag   = if ($tagCol -and $csvRow.$tagCol)   { [string]$csvRow.$tagCol }   else { '' }
            if ([string]::IsNullOrWhiteSpace($tag)) { continue }
            $empId = if ($empCol -and $csvRow.$empCol)   { [string]$csvRow.$empCol }   else { '' }
            $mail  = if ($mailCol -and $csvRow.$mailCol) { [string]$csvRow.$mailCol }  else { '' }
            $sts   = if ($stsCol -and $csvRow.$stsCol)   { [string]$csvRow.$stsCol }   else { '' }

            try {
                $built = CM-BuildRow -TAG $tag -EmpID $empId -SDStatus $sts -Email $mail `
                    -Computers $computers -Users $users -UsersByMail $usersByMail `
                    -UsersBySAM $usersBySAM -UsersDup $usersDup -UserCompMap $ucMap
                if ($null -ne $built) { [void]$localBatch.Add($built) }
            } catch {}

            # Push to UI in batches of 250 for speed
            if ($localBatch.Count -ge 250 -or $cur -eq $total) {
                $batch = $localBatch.ToArray()
                $localBatch.Clear()
                $pct = [int](($cur / $total) * 100)
                $c   = $cur
                $window.Dispatcher.Invoke([action]{
                    foreach ($r in $batch) { $allRows.Add($r) }
                    $cm['CM_Progress'].Value       = $pct
                    $cm['CM_ProgressLabel'].Text   = "Matching $c / $total"
                })
            }
        }

        # Fix multi-computer flags
        $window.Dispatcher.Invoke([action]{
            foreach ($r in $allRows) {
                if ($r._UserDN -ne '' -and $ucMap.ContainsKey($r._UserDN)) {
                    $r.IsMulti = ($ucMap[$r._UserDN].Count -gt 1)
                }
            }
            $cm['CM_Progress'].Visibility      = 'Collapsed'
            $cm['CM_ProgressLabel'].Visibility = 'Collapsed'
            $cm['CM_BtnImport'].IsEnabled      = $true
            $cm['CM_BtnFetchAD'].IsEnabled     = $true
            $cm['CM_BtnRun'].IsEnabled         = $true
            $cm['CM_BtnExport'].IsEnabled      = $true
            if ($cm['CM_BtnDeepScan']) { $cm['CM_BtnDeepScan'].IsEnabled = $true }
            $cm['CM_StatusBar'].Text           = "Match complete — $($allRows.Count) rows"
            $cm['__CMDone'] = $true
        })
    })

    $handle    = $ps.BeginInvoke()
    $startTick = [Environment]::TickCount

    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(400)
    $t.Tag = [PSCustomObject]@{ PS=$ps; RS=$rs; Handle=$handle; StartTick=$startTick }
    $t.Add_Tick({
        $state = $this.Tag
        if ($state.Handle.IsCompleted -or $Script:CM.ContainsKey('__CMDone')) {
            $this.Stop()
            try { $state.PS.EndInvoke($state.Handle) } catch {}
            $state.PS.Dispose(); $state.RS.Close(); $state.RS.Dispose()
            $Script:CM.Remove('__CMDone')
            # Retrieve the computers hashtable stored by the runspace via $cm bridge
            if ($Script:CM.ContainsKey('__ADComputers')) {
                $Script:CM_ADComputers = $Script:CM['__ADComputers']
                $Script:CM.Remove('__ADComputers')
            }
            $elapsed = [Math]::Max(0, ([Environment]::TickCount - $state.StartTick) / 1000.0)
            & $Script:CM_ShowProgress -Value -1
            & $Script:CM_UpdateCounters
            & $Script:CM_SetStatus "Ready — $($Script:CM_AllRows.Count) rows matched in $([Math]::Round($elapsed,1))s"
            Write-Log "ComputerMapper: match complete — $($Script:CM_AllRows.Count) rows in $([Math]::Round($elapsed,1))s" -Source 'ComputerMapper'
        }
    })
    $t.Start()
}

# =============================================================================
# EVENT WIRING
# =============================================================================

# Connection handler
Register-ConnectionHandler -Handler {
    param([bool]$Connected)
    if ($null -eq $Script:CM -or $null -eq $Script:CM['CM_StatusBar']) { return }
    if ($Connected) {
        $Script:CM['CM_BtnImport'].IsEnabled = $true
        & $Script:CM_SetStatus "Import a file to begin"
    } else {
        $Script:CM['CM_BtnImport'].IsEnabled  = $false
        $Script:CM['CM_BtnRun'].IsEnabled     = $false
        $Script:CM['CM_BtnFetchAD'].IsEnabled = $false
        & $Script:CM_SetStatus "Connect to a DC to begin"
    }
}

# ── WhatIf badge toggle ───────────────────────────────────────────────────────
$Script:CM['CM_WhatIfBadge'].Add_MouseLeftButtonUp({
    $current = Get-WhatIfMode
    if ($current) {
        $c = [Windows.MessageBox]::Show("Switch to LIVE MODE?`n`nChanges WILL be written to AD.", "WhatIf Warning", 'YesNo', 'Warning')
        if ($c -ne 'Yes') { return }
        Set-WhatIfMode -Mode $false
        $Script:CM['CM_WhatIfBadge'].Background = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xC0,0x39,0x2B)
        $Script:CM['CM_WhatIfLabel'].Text = "OFF — LIVE"
        if ($Script:CM['CM_WhatIfToggle']) { $Script:CM['CM_WhatIfToggle'].IsChecked = $false }
        Write-Log "ComputerMapper: WhatIf OFF — LIVE mode" -Level WARN -Source 'ComputerMapper'
    } else {
        Set-WhatIfMode -Mode $true
        $Script:CM['CM_WhatIfBadge'].Background = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0x27,0xAE,0x60)
        $Script:CM['CM_WhatIfLabel'].Text = "ON"
        if ($Script:CM['CM_WhatIfToggle']) { $Script:CM['CM_WhatIfToggle'].IsChecked = $true }
        Write-Log "ComputerMapper: WhatIf ON" -Source 'ComputerMapper'
    }
})

# ── Filter buttons ────────────────────────────────────────────────────────────
foreach ($btnKey in $Script:CM_FilterMap.Keys) {
    $btn = $Script:CM[$btnKey]
    if ($null -eq $btn) { continue }
    $btn.Add_Checked({
        param($s,$e)
        foreach ($k in $Script:CM_FilterMap.Keys) { if ($Script:CM[$k] -ne $s) { $Script:CM[$k].IsChecked = $false } }
        foreach ($k in $Script:CM_FilterMap.Keys) { if ($Script:CM[$k] -eq $s) { $Script:CM_ActiveFilter = $Script:CM_FilterMap[$k]; break } }
        & $Script:CM_ApplyFilter
    })
}

# ── Search ────────────────────────────────────────────────────────────────────
$Script:CM['CM_BtnSearch'].Add_Click({
    $Script:CM_SearchTerm = $Script:CM['CM_TxtSearch'].Text.Trim()
    & $Script:CM_ApplyFilter
})
$Script:CM['CM_TxtSearch'].Add_KeyDown({
    param($s,$e)
    if ($e.Key -eq 'Return') { $Script:CM_SearchTerm = $Script:CM['CM_TxtSearch'].Text.Trim(); & $Script:CM_ApplyFilter }
})
$Script:CM['CM_BtnClearSearch'].Add_Click({
    $Script:CM['CM_TxtSearch'].Clear()
    $Script:CM_SearchTerm = ''
    & $Script:CM_ApplyFilter
})

# ── Grid selection count ──────────────────────────────────────────────────────
$Script:CM['CM_Grid'].Add_SelectionChanged({
    $Script:CM['CM_SelCount'].Text = "$($Script:CM['CM_Grid'].SelectedItems.Count) selected"
})

# ── IMPORT button — loads file into grid immediately, NO AD query yet ─────────
$Script:CM['CM_BtnImport'].Add_Click({
    if (-not (Get-AppConnected)) {
        [System.Windows.MessageBox]::Show("Connect to a DC first.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }

    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Title  = "Select Asset File"
    $ofd.Filter = "Spreadsheet files (*.csv;*.xlsx)|*.csv;*.xlsx|CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx"
    if ($ofd.ShowDialog() -ne $true) { return }
    $fp = $ofd.FileName

    $Script:CM_AllRows.Clear()
    $Script:CM_UserCompMap  = @{}
    $Script:CM_LastImport   = $null
    # Re-bind after clear to ensure grid picks up the collection
    $Script:CM['CM_Grid'].ItemsSource = $null
    $Script:CM['CM_Grid'].ItemsSource = $Script:CM_AllRows
    # Reset any active filter so rows are visible immediately
    $Script:CM_ActiveFilter = 'All'
    $Script:CM_SearchTerm   = ''
    if ($Script:CM['CM_FAll']) { $Script:CM['CM_FAll'].IsChecked = $true }
    $Script:CM['CM_TxtFile'].Text = "Loading: $([System.IO.Path]::GetFileName($fp))..."
    & $Script:CM_ShowProgress -Value 5 -Label "Reading file..."
    & $Script:CM_SetStatus "Reading file..."

    try {
        . ([string]$Script:CM_FnPath)
        $res = CM-ImportFile -FilePath $fp
    } catch {
        [System.Windows.MessageBox]::Show("Parse error:`n$($_.Exception.Message)", "AD Manager", 'OK', 'Error') | Out-Null
        & $Script:CM_ShowProgress -Value -1
        & $Script:CM_SetStatus "Import failed"
        return
    }

    if (-not $res.OK) {
        [System.Windows.MessageBox]::Show("Import failed:`n$($res.ErrorMsg)", "AD Manager", 'OK', 'Error') | Out-Null
        $Script:CM['CM_TxtFile'].Text = "Import failed"
        & $Script:CM_ShowProgress -Value -1
        & $Script:CM_SetStatus "Import failed"
        return
    }

    # Save for Fetch AD button
    $Script:CM_LastImport = $res

    # Populate grid with placeholder rows so user sees data immediately
    foreach ($csvRow in $res.Rows) {
        $tag = if ($res.TagCol -and $csvRow.($res.TagCol)) { [string]$csvRow.($res.TagCol) } else { '' }
        if ([string]::IsNullOrWhiteSpace($tag)) { continue }
        $Script:CM_AllRows.Add([PSCustomObject]@{
            Status           = 'Pending'
            StatusLabel      = '⏳ Waiting for AD Fetch'
            TAG              = $tag
            EmpID            = if ($res.EmpCol -and $csvRow.($res.EmpCol)) { [string]$csvRow.($res.EmpCol) } else { '' }
            SDStatus         = if ($res.StatusCol -and $csvRow.($res.StatusCol)) { [string]$csvRow.($res.StatusCol) } else { '' }
            ComputerName     = ''
            ComputerEnabled  = ''
            DisplayName      = ''
            SAM              = ''
            DescCurrent      = ''
            DescNew          = ''
            ManagedByCurrent = ''
            Office           = ''
            CurrentOU        = ''
            ResultNote       = ''
            IsMulti          = $false
            _UserDN          = ''
            _CompDN          = ''
        })
    }

    $Script:CM['CM_TxtFile'].Text         = "$([System.IO.Path]::GetFileName($fp))  ($($res.TotalRows) rows)"
    $Script:CM['CM_BtnFetchAD'].IsEnabled = $true
    $Script:CM['CM_BtnExport'].IsEnabled  = $true

    & $Script:CM_ShowProgress -Value -1
    & $Script:CM_UpdateCounters
    $Script:CM['CM_Grid'].Items.Refresh()
    & $Script:CM_SetStatus "File loaded — $($res.TotalRows) rows. Click '☁ Fetch AD Data' to match."
    Write-Log "ComputerMapper: file loaded — $($res.TotalRows) rows" -Source 'ComputerMapper'
})

# ── FETCH AD DATA button — triggers AD query + matching ──────────────────────
$Script:CM['CM_BtnFetchAD'].Add_Click({
    if ($null -eq $Script:CM_LastImport) {
        [System.Windows.MessageBox]::Show("Import a file first.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }
    if (-not (Get-AppConnected)) {
        [System.Windows.MessageBox]::Show("Connect to a DC first.", "AD Manager", 'OK', 'Warning') | Out-Null
        return
    }
    $Script:CM_UserCompMap = @{}
    & $Script:CM_StartBulkAndMatch -ImportData $Script:CM_LastImport
})

# ── CLEAR ALL button ──────────────────────────────────────────────────────────
$Script:CM['CM_BtnClearAll'].Add_Click({
    if ($Script:CM_AllRows.Count -gt 0) {
        if ([Windows.MessageBox]::Show("Clear all data and start over?", "Confirm", 'YesNo', 'Question') -ne 'Yes') { return }
    }
    $Script:CM_AllRows.Clear()
    $Script:CM_UserCompMap  = @{}
    $Script:CM_LastImport   = $null
    $Script:CM_ADComputers  = $null
    $Script:CM['CM_TxtFile'].Text         = "No file loaded"
    $Script:CM['CM_BtnFetchAD'].IsEnabled = $false
    $Script:CM['CM_BtnRun'].IsEnabled     = $false
    $Script:CM['CM_BtnExport'].IsEnabled  = $false
    & $Script:CM_UpdateCounters
    & $Script:CM_SetStatus "Cleared."
    Write-Log "ComputerMapper: data cleared" -Source 'ComputerMapper'
})

# ── DEEP SCAN button ──────────────────────────────────────────────────────────
if ($Script:CM['CM_BtnDeepScan']) {
    $Script:CM['CM_BtnDeepScan'].Add_Click({
        if ($Script:CM_AllRows.Count -eq 0) { return }
        if ($null -eq $Script:CM_ADComputers) {
            [System.Windows.MessageBox]::Show("Run 'Fetch AD Data' first — Deep Scan reuses the already-loaded AD computer list.", "AD Manager", 'OK', 'Warning') | Out-Null
            return
        }

        $Script:CM['CM_BtnDeepScan'].IsEnabled = $false
        & $Script:CM_ShowProgress -Value 5 -Label "Scanning for orphans..."

        # Build set of computer names already in the grid (from the imported file)
        $matchedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($r in $Script:CM_AllRows) {
            if ($r.ComputerName -and $r.Status -notin @('InADOnly','InADStale')) {
                [void]$matchedNames.Add($r.ComputerName)
            }
        }

        # Reuse the already-fetched AD computers hashtable — no second AD query needed
        $adComputers = $Script:CM_ADComputers
        $orphans     = [System.Collections.Generic.List[object]]::new()
        $now         = Get-Date

        # Inline OU parser — Functions.ps1 not available on main thread here
        function Local-ParseOU {
            param([string]$DN)
            if ([string]::IsNullOrWhiteSpace($DN)) { return '' }
            $ouParts = ($DN -split ',') | Where-Object { $_ -match '^OU=' }
            if ($ouParts) { return ($ouParts -join ' > ').Replace('OU=','') }
            return ''
        }

        foreach ($key in $adComputers.Keys) {
            $comp = $adComputers[$key]
            if ($matchedNames.Contains($comp.Name)) { continue }

            # Exclude servers and DCs (primaryGroupId 516=DC, 521=RODC)
            if ($comp.OperatingSystem -like '*Server*') { continue }

            $isStale   = ($null -ne $comp.LastLogon -and $comp.LastLogon -lt $now.AddDays(-180))
            $status    = if ($isStale) { 'InADStale' } else { 'InADOnly' }
            $statusLbl = if ($isStale) { '🔴 In AD (STALE)' } else { '❓ In AD Only' }

            $orphans.Add([PSCustomObject]@{
                Status           = $status
                StatusLabel      = $statusLbl
                TAG              = $comp.Name
                ComputerName     = $comp.Name
                ComputerEnabled  = if ($comp.Enabled) { '✅ Enabled' } else { '❌ Disabled' }
                SDStatus         = ''
                EmpID            = ''
                DisplayName      = ''
                SAM              = ''
                DescCurrent      = $comp.Description
                DescNew          = ''
                ManagedByCurrent = $comp.ManagedBy
                ManagedByNew     = ''
                Office           = ''
                CurrentOU        = Local-ParseOU -DN $comp.DistinguishedName
                ResultNote       = if ($isStale) { "Stale — last logon: $($comp.LastLogon.ToString('yyyy-MM-dd'))" } else { 'Not in imported file' }
                IsMulti          = $false
                _UserDN          = ''
                _CompDN          = $comp.DistinguishedName
            })
        }

        & $Script:CM_ShowProgress -Value -1

        if ($orphans.Count -gt 0) {
            foreach ($o in $orphans) { $Script:CM_AllRows.Add($o) }
            & $Script:CM_UpdateCounters
            # Switch filter to In AD Only so user sees results immediately
            $Script:CM_ActiveFilter = 'InADOnly'
            if ($Script:CM['CM_FExtra']) { $Script:CM['CM_FExtra'].IsChecked = $true }
            & $Script:CM_ApplyFilter
            & $Script:CM_SetStatus "Deep Scan complete — $($orphans.Count) orphan(s) found"
            [System.Windows.MessageBox]::Show("Deep Scan complete!`nFound $($orphans.Count) orphan(s) not in your imported file.", "AD Manager", 'OK', 'Information') | Out-Null
        } else {
            & $Script:CM_SetStatus "Deep Scan complete — no orphans found"
            [System.Windows.MessageBox]::Show("Deep Scan complete.`nNo orphans found — all AD computers are in your imported file.", "AD Manager", 'OK', 'Information') | Out-Null
        }

        $Script:CM['CM_BtnDeepScan'].IsEnabled = $true
    })
}

# ── RUN UPDATES button ────────────────────────────────────────────────────────
$Script:CM['CM_BtnRun'].Add_Click({
    if (-not (Get-AppConnected)) { return }

    $pending = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $Script:CM_AllRows) {
        if ($r.Status -eq 'Pending') { [void]$pending.Add($r) }
    }

    if ($pending.Count -eq 0) {
        [Windows.MessageBox]::Show("No rows in 'Pending Update' status.", "AD Manager", 'OK', 'Information') | Out-Null
        return
    }

    $wi      = Get-WhatIfMode
    $modeStr = if ($wi) { 'WhatIf PREVIEW — no AD changes' } else { 'LIVE — will write to AD' }
    if ([Windows.MessageBox]::Show("Process $($pending.Count) row(s)?`n`nMode: $modeStr", "Confirm", 'YesNo', 'Question') -ne 'Yes') { return }

    $Script:CM_UpdateQueue   = $pending
    $Script:CM_UpdateTotal   = $pending.Count
    $Script:CM_UpdateIdx     = 0
    $Script:CM_UpdateOK      = 0
    $Script:CM_UpdateFailed  = 0
    $Script:CM_IsRunning     = $true
    $Script:CM_UpdateDC      = Get-AppDCName
    $Script:CM_UpdateCred    = Get-AppCredential
    $Script:CM_UpdateWhatIf  = $wi
    $Script:CM_UpdateThreads = [System.Collections.Generic.List[object]]::new()

    # Pre-build ISS with AD module so workers don't each load it cold
    $Script:CM_WorkerISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    [void]$Script:CM_WorkerISS.ImportPSModule('ActiveDirectory')

    $Script:CM['CM_BtnRun'].IsEnabled    = $false
    $Script:CM['CM_BtnStop'].IsEnabled   = $true
    $Script:CM['CM_BtnStop'].Visibility  = 'Visible'
    & $Script:CM_ShowProgress -Value 0 -Label "Starting..."

    $uT = New-Object System.Windows.Threading.DispatcherTimer
    $uT.Interval = [TimeSpan]::FromMilliseconds(50)
    $uT.Add_Tick({
        if (-not $Script:CM_IsRunning) {
            foreach ($thr in $Script:CM_UpdateThreads) { try { $thr.PS.Stop() } catch {} }
            $this.Stop(); return
        }

        # Harvest completed threads
        for ($i = $Script:CM_UpdateThreads.Count - 1; $i -ge 0; $i--) {
            $thr = $Script:CM_UpdateThreads[$i]
            if (-not $thr.Handle.IsCompleted) { continue }

            $res = $null
            try {
                $out = $thr.PS.EndInvoke($thr.Handle)
                $res = $out | Where-Object { $_ -is [hashtable] -and $null -ne $_.OK } | Select-Object -Last 1
            } catch {}
            $thr.PS.Dispose()
            if ($thr.RS) { $thr.RS.Close(); $thr.RS.Dispose() }

            if ($res -and $res.OK) {
                $Script:CM_UpdateOK++
                $thr.Row.Status      = if ($Script:CM_UpdateWhatIf) { 'WouldUpdate' } else { 'Updated' }
                $thr.Row.StatusLabel = if ($Script:CM_UpdateWhatIf) { '🔍 Would Update' } else { '✅ Updated' }
                $thr.Row.ResultNote  = $res.Message
            } else {
                $Script:CM_UpdateFailed++
                $thr.Row.Status      = 'Failed'
                $thr.Row.StatusLabel = '❌ Failed'
                $thr.Row.ResultNote  = if ($res) { $res.Message } else { 'Update failed' }
            }

            # Refresh row in-place
            $idx = $Script:CM_AllRows.IndexOf($thr.Row)
            if ($idx -ge 0) { $Script:CM_AllRows.RemoveAt($idx); $Script:CM_AllRows.Insert($idx, $thr.Row) }
            $Script:CM_UpdateThreads.RemoveAt($i)

            $done = $Script:CM_UpdateOK + $Script:CM_UpdateFailed
            & $Script:CM_ShowProgress -Value ([int](($done / $Script:CM_UpdateTotal) * 100)) -Label "Updating $done / $Script:CM_UpdateTotal"
            & $Script:CM_UpdateCounters
        }

        # Feed new threads (max 10 parallel)
        while ($Script:CM_UpdateThreads.Count -lt 10 -and $Script:CM_UpdateIdx -lt $Script:CM_UpdateTotal) {
            $row   = $Script:CM_UpdateQueue[$Script:CM_UpdateIdx++]
            $rDN   = $row._CompDN
            $rDesc = $row.DescNew
            $rMgBy = $row._UserDN
            $tDC   = $Script:CM_UpdateDC
            $tCr   = $Script:CM_UpdateCred
            $tWI   = $Script:CM_UpdateWhatIf
            $tFn   = $Script:CM_FnPath
            $tLog  = $Script:LogFile

            $rs = [RunspaceFactory]::CreateRunspace($Script:CM_WorkerISS)
            $rs.ApartmentState = 'MTA'; $rs.Open()

            $psU = [PowerShell]::Create(); $psU.Runspace = $rs
            [void]$psU.AddScript({
                param($rDN,$rDesc,$rMgBy,$tDC,$tCr,$tWI,$tFn,$tLog)
                function Write-Log { param($m,$l='INFO',$s='Log')
                    try { if ($tLog) { Add-Content $tLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][$s] $m" -Encoding UTF8 } } catch {}
                }
                if (Test-Path $tFn) { . ([string]$tFn) }
                CM-UpdateComputer -ComputerDN $rDN -NewDescription $rDesc -NewManagedBy $rMgBy `
                    -DCName $tDC -Credential $tCr -WhatIf $tWI
            })
            $psU.AddParameter('rDN',   $rDN)
            $psU.AddParameter('rDesc', $rDesc)
            $psU.AddParameter('rMgBy', $rMgBy)
            $psU.AddParameter('tDC',   $tDC)
            $psU.AddParameter('tCr',   $tCr)
            $psU.AddParameter('tWI',   $tWI)
            $psU.AddParameter('tFn',   $tFn)
            $psU.AddParameter('tLog',  $tLog)

            $handle = $psU.BeginInvoke()
            [void]$Script:CM_UpdateThreads.Add(@{ PS=$psU; RS=$rs; Handle=$handle; Row=$row })
        }

        # All done
        if ($Script:CM_UpdateIdx -ge $Script:CM_UpdateTotal -and $Script:CM_UpdateThreads.Count -eq 0) {
            $this.Stop()
            $Script:CM_IsRunning = $false
            $Script:CM['CM_BtnStop'].IsEnabled  = $false
            $Script:CM['CM_BtnStop'].Visibility = 'Collapsed'
            $Script:CM['CM_BtnRun'].IsEnabled   = $true
            & $Script:CM_ShowProgress -Value -1
            & $Script:CM_UpdateCounters
            & $Script:CM_SetStatus "Done — ✅ $Script:CM_UpdateOK updated   ❌ $Script:CM_UpdateFailed failed"
            Write-Log "ComputerMapper: bulk update done. OK=$Script:CM_UpdateOK FAIL=$Script:CM_UpdateFailed" -Source 'ComputerMapper'
        }
    })
    $uT.Start()
})

# ── STOP button ───────────────────────────────────────────────────────────────
$Script:CM['CM_BtnStop'].Add_Click({
    $Script:CM_IsRunning = $false
    & $Script:CM_SetStatus "Stopped by user"
    Write-Log "ComputerMapper: update stopped by user" -Level WARN -Source 'ComputerMapper'
})

# ── EXPORT button ─────────────────────────────────────────────────────────────
$Script:CM['CM_BtnExport'].Add_Click({
    if ($Script:CM_AllRows.Count -eq 0) {
        [Windows.MessageBox]::Show("No data to export.", "AD Manager", 'OK', 'Information') | Out-Null
        return
    }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Title    = "Export Computer Mapper Data"
    $sfd.Filter   = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv"
    $sfd.FileName = "ComputerMapper_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
    if ($sfd.ShowDialog() -ne $true) { return }

    & $Script:CM_SetStatus "Exporting..."
    try {
        . ([string]$Script:CM_FnPath)
        $result = CM-ExportToExcel -AllRows $Script:CM_AllRows -OutputPath $sfd.FileName
        if ($result.OK) {
            [Windows.MessageBox]::Show("Saved:`n$($result.Path)", "AD Manager", 'OK', 'Information') | Out-Null
            & $Script:CM_SetStatus "Exported: $([System.IO.Path]::GetFileName($result.Path))"
        } else {
            [Windows.MessageBox]::Show("Export failed:`n$($result.ErrorMsg)", "AD Manager", 'OK', 'Error') | Out-Null
            & $Script:CM_SetStatus "Export failed"
        }
    } catch {
        [Windows.MessageBox]::Show("Export error:`n$($_.Exception.Message)", "AD Manager", 'OK', 'Error') | Out-Null
        & $Script:CM_SetStatus "Export failed"
    }
})

Write-Log "ComputerMapper handlers registered" -Source 'ComputerMapper'
