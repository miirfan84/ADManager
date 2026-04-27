# =============================================================================
# Plugins\OUMover\Handlers.ps1  —  AD Manager v2.0
# All UI event wiring for the OU Mover tab.
# $Script:PluginRoot / $Script:PluginMeta set by PluginLoader.
# =============================================================================

Write-Log "OUMover: registering handlers..." -Source 'OUMover'

# ── Grab controls ─────────────────────────────────────────────────────────────
$CM = @{}
$cmControls = @(
    'CM_BtnImport', 'CM_TxtFile',
    'CM_CntTotal', 'CM_CntOK', 'CM_CntPending', 'CM_CntWarn', 'CM_CntErr',
    'CM_WhatIfBadge', 'CM_WhatIfLabel', 'CM_WhatIfToggle',
    'CM_Progress', 'CM_ProgressLabel',
    'CM_FAll','CM_FOK','CM_FPending','CM_FWarn','CM_FNotFound',
    'CM_FNoUser','CM_FMulti','CM_FUpdated','CM_FFailed',
    'CM_TxtSearch','CM_BtnSearch','CM_BtnClearSearch',
    'CM_Grid',
    'CM_BtnRun','CM_BtnStop','CM_StatusBar','CM_SelCount','CM_BtnExport'
)
foreach ($n in $cmControls) {
    $ctrl = $Script:PluginRoot.FindName($n)
    if ($null -ne $ctrl) {
        $CM[$n] = $ctrl
        Write-Log "  Control: $n" -Source 'ComputerMapper'
    } else {
        Write-Log "  MISSING control: $n" -Level WARN -Source 'ComputerMapper'
    }
}

# ── Plugin state ──────────────────────────────────────────────────────────────
$Script:CM_AllRows     = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:CM_ImportData  = $null   # result from CM-ImportFile
$Script:CM_ADData      = $null   # result from CM-BulkFetchAD
$Script:CM_UserCompMap = @{}     # user DN → [computer names] for multi-detection
$Script:CM_ActiveFilter = 'All'
$Script:CM_SearchTerm  = ''
$Script:CM_IsRunning   = $false

# Bind grid to collection
$CM['CM_Grid'].ItemsSource = $Script:CM_AllRows

# ── Counter helpers ───────────────────────────────────────────────────────────
function CM-UpdateCounters {
    $total   = $Script:CM_AllRows.Count
    $ok      = 0; $pending = 0; $warn = 0; $err = 0
    foreach ($r in $Script:CM_AllRows) {
        switch ($r.Status) {
            'NoAction'    { $ok++      }
            'Updated'     { $ok++      }
            'WouldUpdate' { $pending++ }
            'Pending'     { $pending++ }
            'Warning'     { $warn++    }
            'NotFound'    { $err++     }
            'NoUser'      { $err++     }
            'Failed'      { $err++     }
        }
    }
    $CM['CM_CntTotal'].Text   = $total
    $CM['CM_CntOK'].Text      = $ok
    $CM['CM_CntPending'].Text = $pending
    $CM['CM_CntWarn'].Text    = $warn
    $CM['CM_CntErr'].Text     = $err
}

# ── Progress helpers ──────────────────────────────────────────────────────────
function CM-ShowProgress {
    param([int]$Value = -1, [string]$Label = '')
    if ($Value -lt 0) {
        $CM['CM_Progress'].Visibility      = 'Collapsed'
        $CM['CM_ProgressLabel'].Visibility = 'Collapsed'
    } else {
        $CM['CM_Progress'].Visibility      = 'Visible'
        $CM['CM_ProgressLabel'].Visibility = 'Visible'
        $CM['CM_Progress'].Value           = [Math]::Min($Value, 100)
        $CM['CM_ProgressLabel'].Text       = $Label
    }
}

function CM-SetStatus { param([string]$Text) $CM['CM_StatusBar'].Text = $Text }

# ── WhatIf toggle ─────────────────────────────────────────────────────────────
# The badge area is clickable — clicking it toggles WhatIf
$CM['CM_WhatIfBadge'].Add_MouseLeftButtonUp({
    $current = Get-WhatIfMode

    if ($current) {
        # Switching OFF → live mode — ask confirmation
        $confirm = [Windows.MessageBox]::Show(
            "You are switching to LIVE MODE.`n`nChanges WILL be written to Active Directory.`n`nAre you sure?",
            "AD Manager — WhatIf Warning", 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') { return }
        Set-WhatIfMode -Mode $false
        $CM['CM_WhatIfBadge'].Background  = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xC0,0x39,0x2B)
        $CM['CM_WhatIfLabel'].Text        = "OFF — LIVE"
        $CM['CM_WhatIfToggle'].IsChecked  = $false
        Write-Log "ComputerMapper: WhatIf switched OFF — LIVE mode" -Level WARN -Source 'ComputerMapper'
    } else {
        Set-WhatIfMode -Mode $true
        $CM['CM_WhatIfBadge'].Background  = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0x27,0xAE,0x60)
        $CM['CM_WhatIfLabel'].Text        = "ON"
        $CM['CM_WhatIfToggle'].IsChecked  = $true
        Write-Log "ComputerMapper: WhatIf switched ON — preview mode" -Source 'ComputerMapper'
    }
})

# ── Filter buttons ─────────────────────────────────────────────────────────────
$filterMap = @{
    'CM_FAll'      = 'All'
    'CM_FOK'       = 'NoAction'
    'CM_FPending'  = 'Pending'
    'CM_FWarn'     = 'Warning'
    'CM_FNotFound' = 'NotFound'
    'CM_FNoUser'   = 'NoUser'
    'CM_FMulti'    = 'Multi'
    'CM_FUpdated'  = 'Updated'
    'CM_FFailed'   = 'Failed'
}

foreach ($btnKey in $filterMap.Keys) {
    $filterValue = $filterMap[$btnKey]
    $btn = $CM[$btnKey]
    if ($null -eq $btn) { continue }

    $btn.Add_Checked({
        param($s, $e)
        # Uncheck all others
        foreach ($k in $filterMap.Keys) {
            if ($CM[$k] -ne $s) { $CM[$k].IsChecked = $false }
        }
        # Find which filter value this is
        foreach ($k in $filterMap.Keys) {
            if ($CM[$k] -eq $s) {
                $Script:CM_ActiveFilter = $filterMap[$k]
                break
            }
        }
        CM-ApplyFilter
    })
}

function CM-ApplyFilter {
    $filter = $Script:CM_ActiveFilter
    $term   = $Script:CM_SearchTerm.ToLower()

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Script:CM_AllRows)
    $view.Filter = {
        param($item)
        $matchFilter = switch ($filter) {
            'All'      { $true }
            'Multi'    { $item.IsMulti -eq $true }
            'Updated'  { $item.Status -eq 'Updated' -or $item.Status -eq 'WouldUpdate' }
            default    { $item.Status -eq $filter }
        }
        if (-not $matchFilter) { return $false }
        if ([string]::IsNullOrWhiteSpace($term)) { return $true }
        # Search across visible text properties
        $haystack = "$($item.TAG) $($item.ComputerName) $($item.EmpID) $($item.DisplayName) $($item.UPN) $($item.Office) $($item.CurrentOU) $($item.ResultNote)".ToLower()
        return $haystack.Contains($term)
    }
    $view.Refresh()
}

# ── Search ────────────────────────────────────────────────────────────────────
$CM['CM_BtnSearch'].Add_Click({
    $Script:CM_SearchTerm = $CM['CM_TxtSearch'].Text.Trim()
    CM-ApplyFilter
})
$CM['CM_TxtSearch'].Add_KeyDown({
    param($s,$e)
    if ($e.Key -eq 'Return') {
        $Script:CM_SearchTerm = $CM['CM_TxtSearch'].Text.Trim()
        CM-ApplyFilter
    }
})
$CM['CM_BtnClearSearch'].Add_Click({
    $CM['CM_TxtSearch'].Clear()
    $Script:CM_SearchTerm = ''
    CM-ApplyFilter
})

# ── Grid selection count ──────────────────────────────────────────────────────
$CM['CM_Grid'].Add_SelectionChanged({
    $count = $CM['CM_Grid'].SelectedItems.Count
    $CM['CM_SelCount'].Text = "$count selected"
})

# ── React to main DC connection changes ──────────────────────────────────────
Register-ConnectionHandler -Handler {
    param([bool]$Connected)
    if (-not $Connected) {
        $CM['CM_BtnRun'].IsEnabled    = $false
        $CM['CM_BtnImport'].IsEnabled = $false
        CM-SetStatus "Connect to a DC to begin"
    } else {
        $CM['CM_BtnImport'].IsEnabled = $true
        CM-SetStatus "Import a file to begin"
    }
}

# =============================================================================
# IMPORT BUTTON
# =============================================================================
$CM['CM_BtnImport'].Add_Click({
    if (-not (Get-AppConnected)) {
        [Windows.MessageBox]::Show("Connect to a DC first.","AD Manager",'OK','Warning') | Out-Null
        return
    }

    # File picker
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Title  = "Select Asset File"
    $ofd.Filter = "Spreadsheet files (*.csv;*.xlsx)|*.csv;*.xlsx|CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx"
    if ($ofd.ShowDialog() -ne $true) { return }

    $filePath = $ofd.FileName
    Write-Log "ComputerMapper: file selected — $filePath" -Source 'ComputerMapper'

    $CM['CM_BtnImport'].IsEnabled = $false
    $CM['CM_BtnRun'].IsEnabled    = $false
    $CM['CM_BtnExport'].IsEnabled = $false
    $Script:CM_AllRows.Clear()
    $Script:CM_UserCompMap = @{}
    CM-ShowProgress -Value 0 -Label "Reading file..."
    CM-SetStatus "Reading file..."
    $CM['CM_TxtFile'].Text = "Loading: $([System.IO.Path]::GetFileName($filePath))..."

    # Import file in background
    $rsFile   = [RunspaceFactory]::CreateRunspace([System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2())
    $rsFile.ApartmentState = 'MTA'; $rsFile.ThreadOptions = 'ReuseThread'; $rsFile.Open()
    $rsFile.SessionStateProxy.SetVariable('FilePath', $filePath)
    $rsFile.SessionStateProxy.SetVariable('LF',       $Script:LogFile)
    # Dot-source Functions.ps1 inside runspace
    $rsFile.SessionStateProxy.SetVariable('FnPath',   (Join-Path $PSScriptRoot "Functions.ps1"))

    $psFile = [PowerShell]::Create(); $psFile.Runspace = $rsFile
    [void]$psFile.AddScript({
        . $FnPath
        $r = CM-ImportFile -FilePath $FilePath
        $r  # return
    })

    $hFile = $psFile.BeginInvoke()

    $timerFile = New-Object System.Windows.Threading.DispatcherTimer
    $timerFile.Interval = [TimeSpan]::FromMilliseconds(300)
    $timerFile.Add_Tick({
        if ($hFile.IsCompleted) {
            $timerFile.Stop()
            $importResult = $null
            try { $importResult = $psFile.EndInvoke($hFile) | Select-Object -Last 1 } catch {}
            $psFile.Dispose(); $rsFile.Close(); $rsFile.Dispose()

            if ($null -eq $importResult -or -not $importResult.OK) {
                $errMsg = if ($importResult) { $importResult.ErrorMsg } else { "Unknown import error" }
                CM-ShowProgress -Value -1
                CM-SetStatus "Import failed: $errMsg"
                $CM['CM_TxtFile'].Text = "Import failed"
                $CM['CM_BtnImport'].IsEnabled = $true
                Write-Log "ComputerMapper: import failed — $errMsg" -Level ERROR -Source 'ComputerMapper'
                [Windows.MessageBox]::Show("Import failed:`n$errMsg","AD Manager",'OK','Error') | Out-Null
                return
            }

            $Script:CM_ImportData = $importResult
            $rowCount = $importResult.TotalRows
            $CM['CM_TxtFile'].Text = "$([System.IO.Path]::GetFileName($filePath))  ($rowCount rows)"
            Write-Log "ComputerMapper: import OK — $rowCount rows, TAG='$($importResult.TagCol)', EmpID='$($importResult.EmpCol)'" -Source 'ComputerMapper'

            CM-ShowProgress -Value 10 -Label "File loaded. Fetching AD data..."
            CM-SetStatus "File loaded ($rowCount rows). Loading AD data..."

            # Now start bulk AD fetch + match
            CM-StartBulkAndMatch -ImportData $importResult
        }
    })
    $timerFile.Start()
})

# =============================================================================
# BULK FETCH + MATCH (streaming results to grid)
# =============================================================================
function CM-StartBulkAndMatch {
    param($ImportData)

    $dc      = Get-AppDCName
    $cred    = Get-AppCredential
    $rows    = $ImportData.Rows
    $tagCol  = $ImportData.TagCol
    $empCol  = $ImportData.EmpCol
    $total   = $ImportData.TotalRows

    $window  = $Script:Window
    $allRows = $Script:CM_AllRows
    $cm      = $CM
    $logFile = $Script:LogFile
    $fnPath  = Join-Path $PSScriptRoot "Functions.ps1"
    $ucMap   = $Script:CM_UserCompMap

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $rs  = [RunspaceFactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()

    foreach ($v in @('dc','cred','rows','tagCol','empCol','total','window',
                     'allRows','cm','logFile','fnPath','ucMap')) {
        $rs.SessionStateProxy.SetVariable($v, (Get-Variable $v -ValueOnly))
    }

    $ps = [PowerShell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript({

        . $fnPath   # load Functions.ps1 into runspace

        function RsLog { param($m,$l='INFO')
            try { Add-Content $logFile "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][CM-Match] $m" -Encoding UTF8 } catch {}
        }

        # ── Bulk AD fetch ─────────────────────────────────────────────────────
        RsLog "Starting bulk AD fetch..."
        $window.Dispatcher.Invoke([action]{ $cm['CM_ProgressLabel'].Text = "Fetching AD data..." })

        $adResult = CM-BulkFetchAD -DCName $dc -Credential $cred -ProgressCallback {
            param($step, $msg)
            $m = $msg
            $window.Dispatcher.Invoke([action]{
                $cm['CM_ProgressLabel'].Text = $m
                $cm['CM_ProgressLabel'].Visibility = 'Visible'
                $cm['CM_Progress'].Visibility      = 'Visible'
            })
        }

        if (-not $adResult.OK) {
            RsLog "AD fetch failed: $($adResult.ErrorMsg)" 'ERROR'
            $errMsg = $adResult.ErrorMsg
            $window.Dispatcher.Invoke([action]{
                $cm['CM_StatusBar'].Text = "AD fetch failed: $errMsg"
                $cm['CM_BtnImport'].IsEnabled = $true
            })
            $window.Dispatcher.Invoke([action]{ $cm['__CMDone'] = $true })
            return
        }

        $computers = $adResult.Computers
        $users     = $adResult.Users
        $usersDup  = $adResult.UsersDup

        RsLog "AD fetch complete. Starting match for $total rows..."

        # ── Match each row ────────────────────────────────────────────────────
        $current = 0
        foreach ($csvRow in $rows) {
            $current++

            $tag    = if ($csvRow.$tagCol)  { [string]$csvRow.$tagCol  } else { '' }
            $empId  = if ($csvRow.$empCol)  { [string]$csvRow.$empCol  } else { '' }

            $builtRow = CM-BuildRow `
                -TAG         $tag `
                -EmpID       $empId `
                -Computers   $computers `
                -Users       $users `
                -UsersDup    $usersDup `
                -UserCompMap $ucMap

            $pct = [int](($current / $total) * 100)
            $rowCapture = $builtRow
            $pctCapture = $pct
            $curCapture = $current

            $window.Dispatcher.Invoke([action]{
                $allRows.Add($rowCapture)

                # Update counters every 50 rows for performance
                if ($curCapture % 50 -eq 0 -or $curCapture -eq $total) {
                    $t=0;$ok=0;$p=0;$w=0;$e=0
                    foreach ($r in $allRows) {
                        $t++
                        switch ($r.Status) {
                            'NoAction'    { $ok++ }
                            'Updated'     { $ok++ }
                            'WouldUpdate' { $p++  }
                            'Pending'     { $p++  }
                            'Warning'     { $w++  }
                            'NotFound'    { $e++  }
                            'NoUser'      { $e++  }
                            'Failed'      { $e++  }
                        }
                    }
                    $cm['CM_CntTotal'].Text   = $t
                    $cm['CM_CntOK'].Text      = $ok
                    $cm['CM_CntPending'].Text = $p
                    $cm['CM_CntWarn'].Text    = $w
                    $cm['CM_CntErr'].Text     = $e
                    $cm['CM_Progress'].Value  = $pctCapture
                    $cm['CM_ProgressLabel'].Text = "Matching $curCapture / $total"
                }
            })
        }

        # ── Fix multi-computer flags ──────────────────────────────────────────
        # After all rows processed, update IsMulti flag on rows where user has 2+ computers
        RsLog "Applying multi-computer flags..."
        $window.Dispatcher.Invoke([action]{
            foreach ($r in $allRows) {
                if ($r._UserDN -ne '' -and $ucMap.ContainsKey($r._UserDN)) {
                    $r.IsMulti = $ucMap[$r._UserDN].Count -gt 1
                    if ($r.IsMulti -and $r.Status -notin @('Warning','NotFound','NoUser')) {
                        # Keep original status but flag as multi
                    }
                }
            }
        })

        RsLog "Match complete — $current rows processed"
        $window.Dispatcher.Invoke([action]{ $cm['__CMDone'] = $true })
    })

    $handle    = $ps.BeginInvoke()
    $startTick = [Environment]::TickCount

    $timerMatch = New-Object System.Windows.Threading.DispatcherTimer
    $timerMatch.Interval = [TimeSpan]::FromMilliseconds(400)
    $timerMatch.Add_Tick({
        if ($handle.IsCompleted -or $CM.ContainsKey('__CMDone')) {
            $timerMatch.Stop()
            try { $ps.EndInvoke($handle) } catch {}
            $ps.Dispose(); $rs.Close(); $rs.Dispose()
            $CM.Remove('__CMDone')

            $elapsed = ([Environment]::TickCount - $startTick) / 1000.0
            if ($elapsed -lt 0) { $elapsed = 0 }

            CM-ShowProgress -Value -1
            CM-SetStatus "Ready — $($Script:CM_AllRows.Count) rows matched in $([Math]::Round($elapsed,1))s"
            $CM['CM_BtnImport'].IsEnabled = $true
            $CM['CM_BtnRun'].IsEnabled    = $true
            $CM['CM_BtnExport'].IsEnabled = $true
            Write-Log "ComputerMapper: matching complete — $($Script:CM_AllRows.Count) rows in $([Math]::Round($elapsed,1))s" -Source 'ComputerMapper'
        }
    })
    $timerMatch.Start()
}

# =============================================================================
# RUN UPDATES BUTTON
# =============================================================================
$CM['CM_BtnRun'].Add_Click({
    if (-not (Get-AppConnected)) { return }

    $pendingRows = @($Script:CM_AllRows | Where-Object { $_.Status -eq 'Pending' })
    if ($pendingRows.Count -eq 0) {
        [Windows.MessageBox]::Show("No rows in 'Pending Update' status.`n`nImport a file first and wait for matching to complete.",
            "AD Manager",'OK','Information') | Out-Null
        return
    }

    $whatIf  = Get-WhatIfMode
    $modeStr = if ($whatIf) { 'WHATIF preview' } else { 'LIVE — will write to AD' }

    $confirm = [Windows.MessageBox]::Show(
        "Run updates for $($pendingRows.Count) pending row(s)?`n`nMode: $modeStr",
        "AD Manager — Confirm", 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }

    $CM['CM_BtnRun'].IsEnabled  = $false
    $CM['CM_BtnStop'].IsEnabled = $true
    $CM['CM_BtnStop'].Visibility = 'Visible'
    $Script:CM_IsRunning = $true

    $dc      = Get-AppDCName
    $cred    = Get-AppCredential
    $window  = $Script:Window
    $cm      = $CM
    $logFile = $Script:LogFile
    $fnPath  = Join-Path $PSScriptRoot "Functions.ps1"

    $total   = $pendingRows.Count
    $current = 0

    CM-ShowProgress -Value 0 -Label "Running updates..."
    CM-SetStatus "Running updates (0/$total)..."

    # Process updates row by row — update grid live
    # Run in dispatcher timer loop rather than one big runspace
    # so each row result appears immediately in grid

    $Script:CM_UpdateQueue   = [System.Collections.ArrayList]::new()
    foreach ($r in $pendingRows) { [void]$Script:CM_UpdateQueue.Add($r) }
    $Script:CM_UpdateIdx    = 0
    $Script:CM_UpdateOK     = 0
    $Script:CM_UpdateFailed = 0
    $Script:CM_UpdateTotal  = $total
    $Script:CM_UpdateDC     = $dc
    $Script:CM_UpdateCred   = $cred
    $Script:CM_UpdateWhatIf = $whatIf
    $Script:CM_UpdateFnPath = $fnPath
    $Script:CM_UpdateLogFile = $logFile

    $Script:CM_UpdateThreads = New-Object System.Collections.ArrayList
    $Script:CM_IsRunning     = $true
    
    $uT = New-Object System.Windows.Threading.DispatcherTimer
    $uT.Interval = [TimeSpan]::FromMilliseconds(50)
    $uT.Add_Tick({
        if (-not $Script:CM_IsRunning) { 
            foreach($thr in $Script:CM_UpdateThreads){ try{$thr.PS.Stop()}catch{} }
            $this.Stop(); return 
        }

        # ── 1. CLEANUP FINISHED THREADS ───────────────────────────────────────
        for ($i = $Script:CM_UpdateThreads.Count - 1; $i -ge 0; $i--) {
            $thr = $Script:CM_UpdateThreads[$i]
            if ($thr.Handle.IsCompleted) {
                $res = $null
                try { 
                    $out = $thr.PS.EndInvoke($thr.Handle)
                    $res = $out | Where-Object { $_ -is [hashtable] -and $_.OK -ne $null } | Select-Object -Last 1
                } catch {
                    Write-Log "OUMover: thread error - $($_.Exception.Message)" -Level ERROR
                }
                
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
                    $thr.Row.ResultNote  = if ($res) { $res.Message } else { 'Update failed or aborted' }
                }
                
                $idxGrid = $Script:CM_AllRows.IndexOf($thr.Row)
                if ($idxGrid -ge 0) { $Script:CM_AllRows.RemoveAt($idxGrid); $Script:CM_AllRows.Insert($idxGrid, $thr.Row) }

                $Script:CM_UpdateThreads.RemoveAt($i)
                CM-UpdateCounters
                $done = $Script:CM_UpdateOK + $Script:CM_UpdateFailed
                $pct  = [int](($done / $Script:CM_UpdateTotal) * 100)
                CM-ShowProgress -Value $pct -Label "Updating $done / $Script:CM_UpdateTotal"
                CM-SetStatus "Running: ✅ $($Script:CM_UpdateOK) updated, ❌ $($Script:CM_UpdateFailed) failed"
            }
        }

        # ── 2. FEED NEW THREADS (THROTTLE @ 10) ──────────────────────────────
        while ($Script:CM_UpdateThreads.Count -lt 10 -and $Script:CM_UpdateIdx -lt $Script:CM_UpdateTotal) {
            $row     = $Script:CM_UpdateQueue[$Script:CM_UpdateIdx++]
            $rowDN   = $row._CompDN; $rowDesc = $row.DescNew; $rowMgBy = $row._UserDN
            $thisDC  = $Script:CM_UpdateDC; $thisCred = $Script:CM_UpdateCred; $thisWI = $Script:CM_UpdateWhatIf; $thisFn = $Script:CM_UpdateFnPath

            $rs = [RunspaceFactory]::CreateRunspace([System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2())
            $rs.ApartmentState = 'MTA'; $rs.Open()

            $psU = [PowerShell]::Create(); $psU.Runspace = $rs
            [void]$psU.AddScript({ 
                param($rowDN, $rowDesc, $rowMgBy, $thisDC, $thisCred, $thisWI, $thisFn, $thisLog)
                Import-Module ActiveDirectory -ErrorAction SilentlyContinue
                function Write-Log { param($m,$l='INFO') try { if($thisLog){ $stamp=(Get-Date -f 'yyyy-MM-dd HH:mm:ss'); Add-Content $thisLog "[$stamp][$l][Worker] $m" -Encoding UTF8 } } catch {} }
                if (Test-Path $thisFn) { . ([string]$thisFn) }
                CM-UpdateComputer -ComputerDN $rowDN -NewDescription $rowDesc -NewManagedBy $rowMgBy -DCName $thisDC -Credential $thisCred -WhatIf $thisWI 
            })
            
            $psU.AddParameter('rowDN',   $rowDN)
            $psU.AddParameter('rowDesc', $rowDesc)
            $psU.AddParameter('rowMgBy', $rowMgBy)
            $psU.AddParameter('thisDC',   $thisDC)
            $psU.AddParameter('thisCred', $thisCred)
            $psU.AddParameter('thisWI',   $thisWI)
            $psU.AddParameter('thisFn',   $thisFn)
            $psU.AddParameter('thisLog',  $Script:LogFile)
            
            $handle = $psU.BeginInvoke()
            $Script:CM_UpdateThreads.Add(@{ PS=$psU; RS=$rs; Handle=$handle; Row=$row })
        }

        if ($Script:CM_UpdateIdx -ge $Script:CM_UpdateTotal -and $Script:CM_UpdateThreads.Count -eq 0) {
            $this.Stop(); $Script:CM_IsRunning = $false
            $CM['CM_BtnStop'].IsEnabled  = $false; $CM['CM_BtnStop'].Visibility = 'Collapsed'
            $CM['CM_BtnRun'].IsEnabled   = $true
            CM-ShowProgress -Value -1
            CM-SetStatus "Bulk update complete! ✅ $($Script:CM_UpdateOK)  ❌ $($Script:CM_UpdateFailed)"
            Write-Log "OUMover: bulk complete. OK=$($Script:CM_UpdateOK) FAIL=$($Script:CM_UpdateFailed)" -Source 'OUMover'
            return
        }
    })
    $uT.Start()
})

# ── Stop button ───────────────────────────────────────────────────────────────
$CM['CM_BtnStop'].Add_Click({
    $Script:CM_IsRunning = $false
    CM-SetStatus "Stopped by user"
    Write-Log "ComputerMapper: update run stopped by user" -Level WARN -Source 'ComputerMapper'
})

# =============================================================================
# EXPORT BUTTON
# =============================================================================
$CM['CM_BtnExport'].Add_Click({
    if ($Script:CM_AllRows.Count -eq 0) {
        [Windows.MessageBox]::Show("No data to export.","AD Manager",'OK','Information') | Out-Null
        return
    }

    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Title      = "Export Computer Mapper Data"
    $sfd.Filter     = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv"
    $sfd.FileName   = "ComputerMapper_Export_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
    if ($sfd.ShowDialog() -ne $true) { return }

    $outputPath = $sfd.FileName
    CM-SetStatus "Exporting..."

    # Ensure functions are loaded in this click scope
    . (Join-Path $PSScriptRoot "Functions.ps1")

    $result = CM-ExportToExcel -AllRows $Script:CM_AllRows -OutputPath $outputPath

    if ($result.OK) {
        $note = if ($result.Note) { "`n$($result.Note)" } else { '' }
        [Windows.MessageBox]::Show("Export saved:`n$($result.Path)$note","AD Manager",'OK','Information') | Out-Null
        CM-SetStatus "Export saved: $([System.IO.Path]::GetFileName($result.Path))"
    } else {
        [Windows.MessageBox]::Show("Export failed:`n$($result.ErrorMsg)","AD Manager",'OK','Error') | Out-Null
        CM-SetStatus "Export failed"
    }
})

Write-Log "ComputerMapper handlers registered" -Source 'ComputerMapper'
