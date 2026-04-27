# =============================================================================
# Plugins\ComputerMapper\Functions.ps1  —  AD Manager v2.0
# Pure logic — no WPF. Called from Handlers.ps1 and background runspaces.
#
# PERFORMANCE STRATEGY (10,000+ rows):
#   - One bulk Get-ADComputer query → hashtable by Name.ToLower()
#   - One bulk Get-ADUser query     → hashtable by Description
#   - Both queries run in parallel background runspaces
#   - Matching is pure in-memory hashtable lookups (~1s for 10k rows)
#   - Results streamed to ObservableCollection row-by-row (live grid update)
# =============================================================================

# ── Status constants ──────────────────────────────────────────────────────────
$Script:CM_Status = @{
    NoAction  = 'NoAction'    # already up to date
    Pending   = 'Pending'     # needs Description + ManagedBy update
    Warning   = 'Warning'     # non-numeric Emp. ID
    NotFound  = 'NotFound'    # TAG not in AD
    NoUser    = 'NoUser'      # computer found, no user match
    Multi     = 'Multi'       # user has 2+ computers
    Updated   = 'Updated'     # successfully updated
    Failed    = 'Failed'      # update failed
    WouldUpdate = 'WouldUpdate'  # WhatIf preview
}

$Script:CM_StatusLabel = @{
    NoAction    = '✅ No Action Needed'
    Pending     = '🔄 Pending Update'
    Warning     = '⚠ Warning'
    NotFound    = '❌ Not Found in AD'
    NoUser      = '👤 User Not Found'
    Multi       = '🖥 Multi-Computer'
    Updated     = '✅ Updated'
    Failed      = '❌ Failed'
    WouldUpdate = '🔍 Would Update'
}

# =============================================================================
# CM-ImportFile
#   Reads CSV or XLSX, finds TAG and Emp. ID columns by name (any position).
#   Returns: @{ OK; Rows=[]; TagCol; EmpCol; TotalRows; ErrorMsg }
#   Rows: array of hashtables with at minimum: TAG, EmpID
# =============================================================================
function CM-ImportFile {
    param([string]$FilePath)

    Write-Log "ComputerMapper: importing '$FilePath'" -Source 'ComputerMapper'

    if (-not (Test-Path $FilePath)) {
        return @{ OK = $false; ErrorMsg = "File not found: $FilePath" }
    }

    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
    $rows = $null

    try {
        if ($ext -eq '.csv') {
            # CSV — fast, native
            $rows = Import-Csv -Path $FilePath -Encoding UTF8 -ErrorAction Stop

        } elseif ($ext -eq '.xlsx' -or $ext -eq '.xls') {
            # XLSX — use COM Excel if available, else ImportExcel module, else openpyxl via python
            $rows = CM-ReadXlsx -FilePath $FilePath
        } else {
            return @{ OK = $false; ErrorMsg = "Unsupported file type: $ext  (use .csv or .xlsx)" }
        }

    } catch {
        return @{ OK = $false; ErrorMsg = "File read error: $($_.Exception.Message)" }
    }

    if ($null -eq $rows -or $rows.Count -eq 0) {
        return @{ OK = $false; ErrorMsg = "File is empty or could not be read" }
    }

    # ── Find TAG and Emp. ID columns by name (case-insensitive) ──────────────
    $headers = $rows[0].PSObject.Properties.Name

    $tagCol = $headers | Where-Object { $_.Trim() -eq 'TAG' } | Select-Object -First 1
    if (-not $tagCol) {
        $tagCol = $headers | Where-Object { $_.Trim().ToLower() -eq 'tag' } | Select-Object -First 1
    }

    $empCol = $headers | Where-Object { $_.Trim() -eq 'Emp. ID' } | Select-Object -First 1
    if (-not $empCol) {
        $empCol = $headers | Where-Object { $_.Trim().ToLower() -like '*emp*id*' -or $_.Trim().ToLower() -eq 'emp. id' -or $_.Trim().ToLower() -eq 'empid' } | Select-Object -First 1
    }

    if (-not $tagCol) {
        return @{ OK = $false; ErrorMsg = "Column 'TAG' not found in file.`nAvailable columns: $($headers -join ', ')" }
    }
    if (-not $empCol) {
        return @{ OK = $false; ErrorMsg = "Column 'Emp. ID' not found in file.`nAvailable columns: $($headers -join ', ')" }
    }

    Write-Log "ComputerMapper: TAG column='$tagCol'  EmpID column='$empCol'  Rows=$($rows.Count)" -Source 'ComputerMapper'

    return @{
        OK        = $true
        Rows      = $rows
        TagCol    = $tagCol
        EmpCol    = $empCol
        TotalRows = $rows.Count
        ErrorMsg  = ''
    }
}

# =============================================================================
# CM-ReadXlsx — read XLSX without RSAT/office dependency
#   Tries: ImportExcel module → COM Excel → embedded Python/openpyxl
# =============================================================================
function CM-ReadXlsx {
    param([string]$FilePath)

    # Method 1: ImportExcel PowerShell module (preferred)
    if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
        Write-Log "ComputerMapper: reading XLSX via ImportExcel module" -Source 'ComputerMapper'
        Import-Module ImportExcel -ErrorAction Stop
        return Import-Excel -Path $FilePath -ErrorAction Stop
    }

    # Method 2: COM Excel (if Excel is installed)
    try {
        Write-Log "ComputerMapper: reading XLSX via COM Excel" -Source 'ComputerMapper'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($FilePath)
        $ws = $wb.Sheets.Item(1)

        $lastRow = $ws.UsedRange.Rows.Count
        $lastCol = $ws.UsedRange.Columns.Count

        # Read headers
        $headers = @()
        for ($c = 1; $c -le $lastCol; $c++) {
            $headers += [string]$ws.Cells.Item(1, $c).Value2
        }

        # Read data rows
        $result = New-Object System.Collections.ArrayList
        for ($r = 2; $r -le $lastRow; $r++) {
            $rowObj = [ordered]@{}
            for ($c = 1; $c -le $lastCol; $c++) {
                $rowObj[$headers[$c-1]] = $ws.Cells.Item($r, $c).Value2
            }
            [void]$result.Add([PSCustomObject]$rowObj)
        }

        $wb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

        return $result

    } catch {
        Write-Log "ComputerMapper: COM Excel failed — $($_.Exception.Message)" -Level WARN -Source 'ComputerMapper'
    }

    # Method 3: Python + openpyxl (always available in this environment)
    try {
        Write-Log "ComputerMapper: reading XLSX via Python/openpyxl" -Source 'ComputerMapper'
        $tmpCsv = [System.IO.Path]::GetTempFileName() + ".csv"
        $pyScript = @"
import openpyxl, csv, sys
wb = openpyxl.load_workbook(r'$FilePath', read_only=True, data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
with open(r'$tmpCsv', 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    for row in rows:
        w.writerow(['' if v is None else str(v) for v in row])
print('OK')
"@
        $pyOut = python3 -c $pyScript 2>&1
        if ($pyOut -contains 'OK' -and (Test-Path $tmpCsv)) {
            $data = Import-Csv -Path $tmpCsv -Encoding UTF8
            Remove-Item $tmpCsv -ErrorAction SilentlyContinue
            return $data
        }
    } catch {
        Write-Log "ComputerMapper: Python XLSX read failed — $($_.Exception.Message)" -Level WARN -Source 'ComputerMapper'
    }

    throw "Cannot read XLSX file — install ImportExcel module: Install-Module ImportExcel"
}

# =============================================================================
# CM-BulkFetchAD
#   Fetches all computers AND all users in two parallel runspaces.
#   Returns: @{ OK; Computers=@{}; Users=@{}; ErrorMsg }
#   Computers hashtable key: Name.ToLower()
#   Users hashtable key:     Description (trimmed)
# =============================================================================
function CM-BulkFetchAD {
    param(
        [string]$DCName,
        [System.Management.Automation.PSCredential]$Credential,
        [scriptblock]$ProgressCallback   # called with (step, message)
    )

    Write-Log "ComputerMapper: starting bulk AD fetch" -Source 'ComputerMapper'
    if ($null -ne $ProgressCallback) { & $ProgressCallback 'start' 'Starting bulk AD fetch...' }

    # ── Shared result container ───────────────────────────────────────────────
    $shared = [System.Collections.Hashtable]::Synchronized(@{
        Computers = $null
        Users     = $null
        CompErr   = $null
        UserErr   = $null
        CompDone  = $false
        UserDone  = $false
    })

    $logFile = $Script:LogFile

    # ── Runspace A — All Computers ────────────────────────────────────────────
    $rsA = [RunspaceFactory]::CreateRunspace(
        [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2())
    $rsA.ApartmentState = 'MTA'; $rsA.ThreadOptions = 'ReuseThread'; $rsA.Open()
    $rsA.SessionStateProxy.SetVariable('DC',     $DCName)
    $rsA.SessionStateProxy.SetVariable('Cred',   $Credential)
    $rsA.SessionStateProxy.SetVariable('Shared', $shared)
    $rsA.SessionStateProxy.SetVariable('LF',     $logFile)

    $psA = [PowerShell]::Create(); $psA.Runspace = $rsA
    [void]$psA.AddScript({
        function Lg { param($m,$l='INFO') try { Add-Content $LF "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][CM-CompFetch] $m" -Encoding UTF8 } catch {} }
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            Lg "Fetching all AD computers..."
            $all = Get-ADComputer -Filter * -Server $DC -Credential $Cred `
                -Properties Name,Description,ManagedBy,Enabled,DistinguishedName `
                -ErrorAction Stop
            $ht = @{}
            foreach ($c in $all) { $ht[$c.Name.ToLower()] = $c }
            $Shared['Computers'] = $ht
            Lg "Computers fetched: $($ht.Count)"
        } catch {
            $Shared['CompErr'] = $_.Exception.Message
            Lg "Computer fetch failed: $($_.Exception.Message)" 'ERROR'
        }
        $Shared['CompDone'] = $true
    })
    $hA = $psA.BeginInvoke()

    # ── Runspace B — All Users ────────────────────────────────────────────────
    $rsB = [RunspaceFactory]::CreateRunspace(
        [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2())
    $rsB.ApartmentState = 'MTA'; $rsB.ThreadOptions = 'ReuseThread'; $rsB.Open()
    $rsB.SessionStateProxy.SetVariable('DC',     $DCName)
    $rsB.SessionStateProxy.SetVariable('Cred',   $Credential)
    $rsB.SessionStateProxy.SetVariable('Shared', $shared)
    $rsB.SessionStateProxy.SetVariable('LF',     $logFile)

    $psB = [PowerShell]::Create(); $psB.Runspace = $rsB
    [void]$psB.AddScript({
        function Lg { param($m,$l='INFO') try { Add-Content $LF "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][$l][CM-UserFetch] $m" -Encoding UTF8 } catch {} }
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            Lg "Fetching all AD users..."
            $all = Get-ADUser -Filter * -Server $DC -Credential $Cred `
                -Properties Description,DisplayName,UserPrincipalName,`
                            Office,DistinguishedName,Enabled `
                -ErrorAction Stop
            $ht  = @{}
            $dup = @{}
            foreach ($u in $all) {
                $desc = if ($u.Description) { $u.Description.Trim() } else { '' }
                if ($desc -ne '') {
                    if ($ht.ContainsKey($desc)) {
                        # Track duplicates — key = description, value = list of users
                        if (-not $dup.ContainsKey($desc)) { $dup[$desc] = [System.Collections.ArrayList]@($ht[$desc]) }
                        [void]$dup[$desc].Add($u)
                    } else {
                        $ht[$desc] = $u
                    }
                }
            }
            $Shared['Users']    = $ht
            $Shared['UsersDup'] = $dup
            Lg "Users fetched: $($ht.Count)  Duplicates: $($dup.Count)"
        } catch {
            $Shared['UserErr'] = $_.Exception.Message
            Lg "User fetch failed: $($_.Exception.Message)" 'ERROR'
        }
        $Shared['UserDone'] = $true
    })
    $hB = $psB.BeginInvoke()

    # ── Wait for both ─────────────────────────────────────────────────────────
    $startTick = [Environment]::TickCount
    while (-not ($shared['CompDone'] -and $shared['UserDone'])) {
        Start-Sleep -Milliseconds 300
        $elapsed = ([Environment]::TickCount - $startTick) / 1000.0
        if ($elapsed -lt 0) { $elapsed = 0 }

        $compStatus = if ($shared['CompDone']) { '✅ Computers' } else { "⏳ Computers ($([Math]::Round($elapsed,0))s)" }
        $userStatus = if ($shared['UserDone']) { '✅ Users'     } else { "⏳ Users ($([Math]::Round($elapsed,0))s)"     }
        if ($null -ne $ProgressCallback) {
            & $ProgressCallback 'loading' "$compStatus   $userStatus"
        }
    }

    try { $psA.EndInvoke($hA) } catch {}
    try { $psB.EndInvoke($hB) } catch {}
    $psA.Dispose(); $rsA.Close(); $rsA.Dispose()
    $psB.Dispose(); $rsB.Close(); $rsB.Dispose()

    if ($shared['CompErr']) {
        return @{ OK = $false; ErrorMsg = "Computer fetch failed: $($shared['CompErr'])" }
    }
    if ($shared['UserErr']) {
        return @{ OK = $false; ErrorMsg = "User fetch failed: $($shared['UserErr'])" }
    }

    $compCount = $shared['Computers'].Count
    $userCount = $shared['Users'].Count
    Write-Log "ComputerMapper: AD fetch complete — $compCount computers, $userCount users" -Source 'ComputerMapper'
    if ($null -ne $ProgressCallback) { & $ProgressCallback 'done' "AD loaded — $compCount computers, $userCount users" }

    return @{
        OK        = $true
        Computers = $shared['Computers']
        Users     = $shared['Users']
        UsersDup  = $shared['UsersDup']
        ErrorMsg  = ''
    }
}

# =============================================================================
# CM-BuildRow
#   Creates one grid data row (PSCustomObject) from matching logic.
#   Pure function — no AD calls, no UI calls.
# =============================================================================
function CM-BuildRow {
    param(
        [string]$TAG,
        [string]$EmpID,
        [hashtable]$Computers,     # from CM-BulkFetchAD
        [hashtable]$Users,         # from CM-BulkFetchAD
        [hashtable]$UsersDup,      # duplicate description tracking
        [hashtable]$UserCompMap    # user DN → list of computer names (built during matching)
    )

    $status = $null
    $note   = ''

    # ── Validate Emp. ID ──────────────────────────────────────────────────────
    $empStr = if ($null -ne $EmpID) { [string]$EmpID } else { '' }
    $isNumeric = $empStr -match '^\d+$'

    if (-not $isNumeric) {
        return [PSCustomObject]@{
            Status          = $Script:CM_Status.Warning
            StatusLabel     = $Script:CM_StatusLabel.Warning
            TAG             = $TAG
            ComputerName    = ''
            ComputerEnabled = ''
            EmpID           = $empStr
            DisplayName     = ''
            UPN             = ''
            DescCurrent     = ''
            DescNew         = ''
            ManagedByCurrent= ''
            Office          = ''
            CurrentOU       = ''
            ResultNote      = "Non-numeric Emp. ID: '$empStr'"
            _UserDN         = ''
            _CompDN         = ''
        }
    }

    # ── Match computer by TAG (case-insensitive) ──────────────────────────────
    $tagKey  = $TAG.ToLower().Trim()
    $adComp  = if ($Computers.ContainsKey($tagKey)) { $Computers[$tagKey] } else { $null }

    if ($null -eq $adComp) {
        return [PSCustomObject]@{
            Status          = $Script:CM_Status.NotFound
            StatusLabel     = $Script:CM_StatusLabel.NotFound
            TAG             = $TAG
            ComputerName    = ''
            ComputerEnabled = ''
            EmpID           = $empStr
            DisplayName     = ''
            UPN             = ''
            DescCurrent     = ''
            DescNew         = ''
            ManagedByCurrent= ''
            Office          = ''
            CurrentOU       = ''
            ResultNote      = "TAG '$TAG' not found in AD"
            _UserDN         = ''
            _CompDN         = ''
        }
    }

    $compName    = $adComp.Name
    $compEnabled = if ($adComp.Enabled) { '✅ Enabled' } else { '❌ Disabled' }
    $descCurrent = if ($adComp.Description) { $adComp.Description } else { '' }
    $managedBy   = if ($adComp.ManagedBy)   { $adComp.ManagedBy   } else { '' }
    $compOU      = CM-ParseOU -DN $adComp.DistinguishedName

    # ── Match user by Description = EmpID ────────────────────────────────────
    $adUser = if ($Users.ContainsKey($empStr)) { $Users[$empStr] } else { $null }

    if ($null -eq $adUser) {
        return [PSCustomObject]@{
            Status          = $Script:CM_Status.NoUser
            StatusLabel     = $Script:CM_StatusLabel.NoUser
            TAG             = $TAG
            ComputerName    = $compName
            ComputerEnabled = $compEnabled
            EmpID           = $empStr
            DisplayName     = ''
            UPN             = ''
            DescCurrent     = $descCurrent
            DescNew         = ''
            ManagedByCurrent= $managedBy
            Office          = ''
            CurrentOU       = $compOU
            ResultNote      = "No AD user with Description='$empStr'"
            _UserDN         = ''
            _CompDN         = $adComp.DistinguishedName
        }
    }

    $userDN      = $adUser.DistinguishedName
    $displayName = if ($adUser.DisplayName) { $adUser.DisplayName } else { '' }
    $upn         = if ($adUser.UserPrincipalName) { $adUser.UserPrincipalName } else { '' }
    $office      = if ($adUser.Office) { $adUser.Office } else { '' }

    # ── Track user → computers for multi-computer detection ──────────────────
    if (-not $UserCompMap.ContainsKey($userDN)) {
        $UserCompMap[$userDN] = [System.Collections.ArrayList]::new()
    }
    [void]$UserCompMap[$userDN].Add($compName)
    $isMulti = $UserCompMap[$userDN].Count -gt 1

    # ── Check if already up to date ───────────────────────────────────────────
    $alreadyCorrect = ($descCurrent -eq $displayName) -and ($managedBy -eq $userDN)

    if ($alreadyCorrect) {
        $status = $Script:CM_Status.NoAction
        $note   = 'Already up to date'
    } else {
        $status = $Script:CM_Status.Pending
        $changes = @()
        if ($descCurrent -ne $displayName) { $changes += "Description: '$descCurrent' → '$displayName'" }
        if ($managedBy   -ne $userDN)      { $changes += "ManagedBy will be set to user DN" }
        $note = $changes -join ' | '
    }

    return [PSCustomObject]@{
        Status          = $status
        StatusLabel     = $Script:CM_StatusLabel[$status]
        TAG             = $TAG
        ComputerName    = $compName
        ComputerEnabled = $compEnabled
        EmpID           = $empStr
        DisplayName     = $displayName
        UPN             = $upn
        DescCurrent     = $descCurrent
        DescNew         = $displayName
        ManagedByCurrent= $managedBy
        Office          = $office
        CurrentOU       = $compOU
        ResultNote      = $note
        IsMulti         = $isMulti
        _UserDN         = $userDN
        _CompDN         = $adComp.DistinguishedName
    }
}

# =============================================================================
# CM-ParseOU  —  extract clean OU path from DistinguishedName
# =============================================================================
function CM-ParseOU {
    param([string]$DN)
    if ([string]::IsNullOrWhiteSpace($DN)) { return '' }
    $parts   = $DN -split ','
    $ouParts = $parts | Where-Object { $_ -match '^OU=' }
    if ($ouParts) { return ($ouParts -join ' > ').Replace('OU=','') }
    return ''
}

# =============================================================================
# CM-UpdateComputer
#   Applies Description + ManagedBy to one computer.
#   WhatIf-aware. Returns @{ OK; Message }
# =============================================================================
function CM-UpdateComputer {
    param(
        [string]$ComputerDN,
        [string]$NewDescription,
        [string]$NewManagedBy,
        [string]$DCName,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]$WhatIf = $true
    )

    $isWhatIf = [bool]$WhatIf
    if ($isWhatIf) {
        Write-Log "OUMover [WHATIF]: would set '$ComputerDN' Description='$NewDescription' ManagedBy='$NewManagedBy'" -Source 'OUMover'
        return @{ OK = $true; Message = "[WhatIf] Success" }
    }

    try {
        Set-ADComputer -Identity $ComputerDN `
            -Description $NewDescription `
            -ManagedBy   $NewManagedBy `
            -Server      $DCName `
            -Credential  $Credential `
            -ErrorAction Stop `
            -WhatIf      $false   # Explicitly disable AD preference for WhatIf in Live mode

        Write-Log "OUMover: SUCCESS updated '$ComputerDN' on $DCName" -Source 'OUMover'
        return @{ OK = $true; Message = 'Updated successfully' }

    } catch {
        $err = $_.Exception.Message
        Write-Log "OUMover: FAILURE updating '$ComputerDN' - $err" -Level ERROR -Source 'OUMover'
        return @{ OK = $false; Message = "Failed: $err" }
    }
}

# =============================================================================
# CM-ExportToExcel
#   Exports grid data to XLSX with one sheet per UI filter.
#   Uses ImportExcel package-based approach for speed and reliability.
# =============================================================================
function CM-ExportToExcel {
    param(
        [System.Collections.ObjectModel.ObservableCollection[object]]$AllRows,
        [string]$OutputPath
    )

    Write-Log "OUMover: exporting to '$OutputPath'" -Source 'OUMover'

    # Convert ObservableCollection to a standard array
    $dataArray = @($AllRows)
    
    if ($dataArray.Count -eq 0) {
        return @{ OK = $false; ErrorMsg = "No data to export" }
    }

    # ── Define Groups matching UI Filters ─────────────────────────────────────
    $groups = [ordered]@{
        'All Data'              = $dataArray
        'No Action Needed'      = @($dataArray | Where-Object { $_.Status -eq $Script:CM_Status.NoAction })
        'Pending Updates'       = @($dataArray | Where-Object { $_.Status -eq $Script:CM_Status.Pending })
        'Warnings'              = @($dataArray | Where-Object { $_.Status -eq $Script:CM_Status.Warning })
        'Not Found in AD'       = @($dataArray | Where-Object { $_.Status -eq $Script:CM_Status.NotFound })
        'User Not Found'        = @($dataArray | Where-Object { $_.Status -eq $Script:CM_Status.NoUser })
        'Multi-Computer'        = @($dataArray | Where-Object { $_.IsMulti -eq $true })
        'Updated Successfully'  = @($dataArray | Where-Object { $_.Status -in @($Script:CM_Status.Updated, $Script:CM_Status.WouldUpdate) })
        'Failed Updates'        = @($dataArray | Where-Object { $_.Status -eq $Script:CM_Status.Failed })
    }

    # Columns to export (OUMover specific mapping)
    $exportProps = @('StatusLabel','TAG','ComputerName','ComputerEnabled','EmpID',
                     'DisplayName','UPN','DescCurrent','DescNew','ManagedByCurrent',
                     'Office','CurrentOU','ResultNote')

    try {
        # Ensure target file is NOT locked
        if (Test-Path $OutputPath) {
            try { Remove-Item $OutputPath -Force -ErrorAction Stop } catch { throw "Target file is in use. Please close it in Excel first." }
        }

        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction Stop

            foreach ($sheetName in $groups.Keys) {
                $rows = $groups[$sheetName]
                if ($rows.Count -eq 0) { continue }
                
                $safeName = $sheetName.Trim()
                if ($safeName.Length -gt 31) { $safeName = $safeName.Substring(0, 31) }

                $rows | Select-Object -Property $exportProps | 
                       Export-Excel -Path $OutputPath -WorksheetName $safeName -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium2
            }

            Write-Log "OUMover: export complete (ImportExcel) — $OutputPath" -Source 'OUMover'
            return @{ OK = $true; Path = $OutputPath }

        } else {
            # Fallback
            $csvPath = $OutputPath -replace '\.xlsx$','.csv'
            $dataArray | Select-Object -Property $exportProps | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Log "OUMover: exported as CSV (ImportExcel missing)" -Level WARN -Source 'OUMover'
            return @{ OK = $true; Path = $csvPath; Note = "Saved as CSV (ImportExcel not installed)" }
        }

    } catch {
        $err = $_.Exception.Message
        Write-Log "OUMover: export failed — $err" -Level ERROR -Source 'OUMover'
        return @{ OK = $false; ErrorMsg = $err }
    }
}
