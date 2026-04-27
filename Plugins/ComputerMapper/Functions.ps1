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
    InADOnly    = 'InADOnly'     # Orphan in AD (not in file)
    InADStale   = 'InADStale'    # Orphan in AD + hasn't logged in > 180d
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
    InADOnly    = '❓ In AD Only'
    InADStale   = '🔴 In AD (STALE)'
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

    # ── Find TAG, Emp. ID, and Status columns by name (case-insensitive) ─────
    $headers = $rows[0].PSObject.Properties.Name

    $tagCol = $headers | Where-Object { $_.Trim().ToLower() -eq 'tag' -or $_.Trim().ToLower() -eq 'asset' } | Select-Object -First 1
    $empCol = $headers | Where-Object { $_.Trim().ToLower() -like '*emp*id*' -or $_.Trim().ToLower() -eq 'empid' } | Select-Object -First 1
    $stsCol = $headers | Where-Object { $_.Trim().ToLower() -eq 'status' -or $_.Trim().ToLower() -eq 'sd status' } | Select-Object -First 1

    if (-not $tagCol) {
        return @{ OK = $false; ErrorMsg = "Column 'TAG' not found in file.`nAvailable columns: $($headers -join ', ')" }
    }
    if (-not $empCol) {
        # Don't fail if we have an Email column as fallback
        $emailCol = $headers | Where-Object { $_.Trim().ToLower() -eq 'email' -or $_.Trim().ToLower() -eq 'mail' } | Select-Object -First 1
        if (-not $emailCol) {
            return @{ OK = $false; ErrorMsg = "Column 'Emp. ID' or 'Email' not found in file." }
        }
    } else {
        # Check if email also exists for additional lookup
        $emailCol = $headers | Where-Object { $_.Trim().ToLower() -eq 'email' -or $_.Trim().ToLower() -eq 'mail' } | Select-Object -First 1
    }

    Write-Log "ComputerMapper: TAG='$tagCol' EmpID='$empCol' Status='$stsCol' Rows=$($rows.Count)" -Source 'ComputerMapper'

    return @{
        OK        = $true
        Rows      = $rows
        TagCol    = $tagCol
        EmpCol    = $empCol
        EmailCol  = $emailCol
        StatusCol = $stsCol
        TotalRows = $rows.Count
        ErrorMsg  = ''
    }
}

# =============================================================================
# CM-ReadXlsx — read XLSX fast
#   Priority order:
#   1. COM Excel bulk range read  — ONE call returns 2D array, ~1s for 10k rows
#   2. ImportExcel module         — if installed, very fast
#   3. Fallback error             — tells user to save as CSV
#
# PERFORMANCE NOTE:
#   The old approach called $ws.Cells.Item($r,$c).Value2 per cell — one COM
#   call per cell. 4500 rows x 12 cols = 54,000 COM calls = 15+ minutes.
#   The correct approach: $ws.UsedRange.Value2 — ONE call, returns 2D array.
#   4500 rows processed in under 2 seconds.
# =============================================================================
function CM-ReadXlsx {
    param([string]$FilePath)

    # Method 1: ImportExcel module (Best performance)
    try {
        Write-Log "ComputerMapper: reading XLSX via ImportExcel module" -Source 'ComputerMapper'
        Import-Module ImportExcel -ErrorAction Stop
        $data = Import-Excel -Path $FilePath -ErrorAction Stop
        Write-Log "ComputerMapper: ImportExcel read complete, $($data.Count) data rows" -Source 'ComputerMapper'
        return @($data)
    } catch {
        Write-Log "ComputerMapper: ImportExcel fast-load failed - $($_.Exception.Message)" -Level WARN -Source 'ComputerMapper'
    }

    # Method 3: Native Zip/XML parsing (Fast, NO dependencies, NO Excel required)
    try {
        Write-Log "ComputerMapper: reading XLSX via Native XML ZIP parsing" -Source 'ComputerMapper'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)

        # 1. Read Shared Strings (if exists)
        $sharedStrings = @()
        $ssEntry = $zip.GetEntry('xl/sharedStrings.xml')
        if ($ssEntry) {
            $stream = $ssEntry.Open()
            $reader = [System.IO.StreamReader]::new($stream)
            $xmlText = $reader.ReadToEnd()
            # Remove namespace strings to safely parse as simple XML
            $xmlText = $xmlText -replace 'xmlns(?:="[^"]*"|:[^=]+="[^"]*")', ''
            $xml = [xml]$xmlText
            $reader.Close()
            $stream.Close()
            
            # Extract texts safely
            if ($xml.sst -and $xml.sst.si) {
                foreach ($si in $xml.sst.si) {
                    if ($si.t) {
                        $sharedStrings += $si.t
                    } elseif ($si.r) {
                        # Rich text: concatenate all <t> nodes inside <r>
                        $text = ($si.r | ForEach-Object { if ($_.t) { $_.t } }) -join ''
                        $sharedStrings += $text
                    } else {
                        $sharedStrings += ''
                    }
                }
            }
        }

        # 2. Extract first worksheet
        $sheetEntry = $zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } | Select-Object -First 1
        if (-not $sheetEntry) { throw "No worksheet found in XLSX architecture" }

        $stream = $sheetEntry.Open()
        $reader = [System.IO.StreamReader]::new($stream)
        $xmlText = $reader.ReadToEnd()
        $xmlText = $xmlText -replace 'xmlns(?:="[^"]*"|:[^=]+="[^"]*")', ''
        $xml = [xml]$xmlText
        $reader.Close()
        $stream.Close()
        $zip.Dispose()

        $result = New-Object System.Collections.ArrayList
        $rows = $xml.worksheet.sheetData.row
        if (-not $rows -or $rows.Count -eq 0) { throw "Worksheet data is empty" }

        # Extract Headers correctly honoring the column letter index
        $headerRow = $rows[0]
        $colIndexToName = @{}
        $headers = @()

        foreach ($c in $headerRow.c) {
            $colRef = $c.r -replace '\d+',''
            $val = ''
            if ($c.v -ne $null) {
                $val = if ($c.t -eq 's') { $sharedStrings[[int]$c.v] } else { $c.v }
            } elseif ($c.is.t -ne $null) {
                $val = $c.is.t
            }
            if ($null -eq $val) { $val = "Col_$($colRef)" }
            $headers += $val
            $colIndexToName[$colRef] = $val
        }

        $blankStreak = 0
        for ($i = 1; $i -lt $rows.Count; $i++) {
            $r = $rows[$i]
            $rowObj = [ordered]@{}
            $isEmptyRow = $true
            
            # Initialize with empty strings based on header names
            foreach ($colRef in $colIndexToName.Keys) {
                $rowObj[$colIndexToName[$colRef]] = ''
            }

            foreach ($c in $r.c) {
                $colRef = $c.r -replace '\d+',''
                if (-not $colIndexToName.ContainsKey($colRef)) { continue }
                
                $val = ''
                if ($c.v -ne $null) {
                    $val = if ($c.t -eq 's') { $sharedStrings[[int]$c.v] } else { $c.v }
                } elseif ($c.is.t -ne $null) {
                    $val = $c.is.t
                }
                
                $strVal = if ($null -ne $val) { [string]$val } else { '' }
                if ($strVal -ne '') { $isEmptyRow = $false }
                
                $rowObj[$colIndexToName[$colRef]] = $strVal
            }

            if ($isEmptyRow) {
                $blankStreak++
                if ($blankStreak -ge 10) { break }
            } else {
                $blankStreak = 0
                [void]$result.Add([PSCustomObject]$rowObj)
            }
        }

        Write-Log "ComputerMapper: XML parsing complete, $($result.Count) data rows" -Source 'ComputerMapper'
        return @($result)

    } catch {
        Write-Log "ComputerMapper: XML parsing failed - $($_.Exception.Message)" -Level WARN -Source 'ComputerMapper'
        if ($zip) { try { $zip.Dispose() } catch {} }
    }

    throw @"
Cannot read XLSX file automatically. None of the backup methods succeeded.

Options:
  1. Save file as CSV from Excel and import it.
  2. Install ImportExcel module: Install-Module ImportExcel -Force
"@
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
            Lg "Connecting to LDAP://$DC..."
            $root = if ($Cred) { [ADSI]"LDAP://$DC" } else { [ADSI]"LDAP://$DC" }
            if ($Cred) {
                # Standard ADSI with explicit credentials
                $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DC", $Cred.UserName, $Cred.GetNetworkCredential().Password)
            }
            
            $searcher = [ADSISearcher]$root
            $searcher.Filter = "(objectCategory=computer)"
            $searcher.PageSize = 1000
            $searcher.ReferralChasing = [System.DirectoryServices.ReferralChasingOption]::All
            $searcher.PropertiesToLoad.AddRange(@('name','description','managedBy','distinguishedName','userAccountControl','operatingSystem','lastLogonTimestamp','primaryGroupId'))
            
            $results = $searcher.FindAll()
            $compHt = @{}
            
            foreach ($res in $results) {
                $p = $res.Properties
                $name = if ($p['name'].Count -gt 0) { [string]$p['name'][0] } else { '' }
                $dn   = if ($p['distinguishedName'].Count -gt 0) { [string]$p['distinguishedName'][0] } else { '' }
                if ($name -eq '' -or $dn -eq '') { continue }
                
                # Exclusion Logic: Strip Servers and Domain Controllers
                $os  = if ($p['operatingSystem'].Count -gt 0) { [string]$p['operatingSystem'][0] } else { '' }
                $pgi = if ($p['primaryGroupId'].Count -gt 0)  { [int]$p['primaryGroupId'][0] } else { 0 }
                
                if ($os -like "*Server*") { continue }
                if ($pgi -eq 516) { continue } # DC
                
                # Convert UAC to enabled status
                $uac = if ($p['userAccountControl'].Count -gt 0) { [int]$p['userAccountControl'][0] } else { 0 }
                $enabled = (($uac -band 2) -eq 0)

                # Flatten into a high-performance custom object using PascalCase
                $flat = [PSCustomObject]@{
                    Name              = $name
                    Description       = if ($p['description'].Count -gt 0) { [string]$p['description'][0] } else { '' }
                    ManagedBy         = if ($p['managedBy'].Count -gt 0)   { [string]$p['managedBy'][0] } else { '' }
                    DistinguishedName = $dn
                    OperatingSystem   = if ($p['operatingSystem'].Count -gt 0) { [string]$p['operatingSystem'][0] } else { '' }
                    LastLogon         = if ($p['lastLogonTimestamp'].Count -gt 0) { [DateTime]::FromFileTime($p['lastLogonTimestamp'][0]) } else { $null }
                    Enabled           = $enabled
                }
                $compHt[$name.ToLower()] = $flat
            }
            $Shared['Computers'] = $compHt
            Lg "Computers fetched: $($compHt.Count)"
        } catch {
            $Shared['CompErr'] = $_.Exception.Message
            Lg "Computer fetch failed: $($_.Exception.Message)" 'ERROR'
        } finally {
            $Shared['CompDone'] = $true
        }
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
            Lg "Connecting to LDAP://$DC for users..."
            $root = if ($Cred) { [ADSI]"LDAP://$DC" } else { [ADSI]"LDAP://$DC" }
            if ($Cred) {
                # Standard ADSI with explicit credentials
                $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DC", $Cred.UserName, $Cred.GetNetworkCredential().Password)
            }
            
            $searcher = [ADSISearcher]$root
            $searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
            $searcher.PageSize = 1000
            $searcher.PropertiesToLoad.AddRange(@('description','displayName','sAMAccountName','physicalDeliveryOfficeName','distinguishedName','userAccountControl','mail'))
            
            $results = $searcher.FindAll()
            $ht     = @{} # Index by Description
            $htMail = @{} # Index by Email
            $htSAM  = @{} # Index by sAMAccountName (for prefix mapping)
            $dup    = @{}
            
            foreach ($res in $results) {
                $p = $res.Properties
                
                $desc = if ($p['description'].Count -gt 0)    { [string]$p['description'][0].Trim() } else { '' }
                $mail = if ($p['mail'].Count -gt 0)           { [string]$p['mail'][0].ToLower().Trim() } else { '' }
                $sam  = if ($p['sAMAccountName'].Count -gt 0) { [string]$p['sAMAccountName'][0].ToLower().Trim() } else { '' }

                # Convert UAC to enabled status
                $uac = if ($p['userAccountControl'].Count -gt 0) { [int]$p['userAccountControl'][0] } else { 0 }
                $enabled = (($uac -band 2) -eq 0)

                $flat = [PSCustomObject]@{
                    Description                = $desc
                    DisplayName                = if ($p['displayName'].Count -gt 0) { [string]$p['displayName'][0].Trim() } else { '' }
                    SamAccountName             = if ($p['sAMAccountName'].Count -gt 0) { [string]$p['sAMAccountName'][0] } else { '' }
                    Email                      = $mail
                    physicalDeliveryOfficeName = if ($p['physicalDeliveryOfficeName'].Count -gt 0) { [string]$p['physicalDeliveryOfficeName'][0] } else { '' }
                    DistinguishedName          = [string]$p['distinguishedName'][0]
                    Enabled                    = $enabled
                }

                # 1. Index by Description (Primary)
                if ($desc -ne '') {
                    if ($ht.ContainsKey($desc)) {
                        if (-not $dup.ContainsKey($desc)) { $dup[$desc] = [System.Collections.ArrayList]@($ht[$desc]) }
                        [void]$dup[$desc].Add($flat)
                    } else { $ht[$desc] = $flat }
                }

                # 2. Index by Mail
                if ($mail -ne '') { $htMail[$mail] = $flat }

                # 3. Index by SAM (Support for matching email prefix)
                if ($sam -ne '') { $htSAM[$sam] = $flat }
            }
            $Shared['Users']      = $ht
            $Shared['UsersByMail']= $htMail
            $Shared['UsersBySAM'] = $htSAM
            $Shared['UsersDup']   = $dup
            Lg "Users fetched: $($ht.Count)  Duplicates: $($dup.Count)"
        } catch {
            $Shared['UserErr'] = $_.Exception.Message
            Lg "User fetch failed: $($_.Exception.Message)" 'ERROR'
        } finally {
            $Shared['UserDone'] = $true
        }
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
        OK          = $true
        Computers   = $shared['Computers']
        Users       = $shared['Users']
        UsersByMail = $shared['UsersByMail']
        UsersBySAM  = $shared['UsersBySAM']
        UsersDup    = $shared['UsersDup']
        ErrorMsg    = ''
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
        [string]$SDStatus = '',
        [string]$Email    = '',
        [hashtable]$Computers,     # from CM-BulkFetchAD
        [hashtable]$Users,         # from CM-BulkFetchAD
        [hashtable]$UsersDup,      # duplicate description tracking
        [hashtable]$UserCompMap,   # user DN → list of computer names
        [hashtable]$UsersByMail,   # lookup by exact mail
        [hashtable]$UsersBySAM,    # lookup by SamAccountName
        [object]$ADCompOverride = $null # Used for "Discovery" (Orphans) where no Excel row exists
    )

    $status = $null
    $note   = ''

    # ── Match computer ────────────────────────────────────────────────────────
    $adComp = $null
    try {
        if ($ADCompOverride) {
            $adComp = $ADCompOverride
        } else {
            if ([string]::IsNullOrWhiteSpace($TAG)) { return $null }
            $tagKey = $TAG.ToLower().Trim()
            $adComp = if ($Computers.ContainsKey($tagKey)) { $Computers[$tagKey] } else { $null }
        }
    } catch { return $null }

    if ($null -eq $adComp) {
        return [PSCustomObject]@{
            Status = $Script:CM_Status.NotFound; StatusLabel = $Script:CM_StatusLabel.NotFound
            TAG = $TAG; ComputerName = ''; ComputerEnabled = ''; SDStatus = $SDStatus; EmpID = $EmpID
            DisplayName = ''; SAM = ''; DescCurrent = ''; DescNew = ''; ManagedByCurrent = ''
            Office = ''; CurrentOU = ''; ResultNote = "TAG '$TAG' not found in AD"
            _UserDN = ''; _CompDN = ''
        }
    }

    $compName    = [string]$adComp.Name
    $compEnabled = if ($adComp.Enabled) { '✅ Enabled' } else { '❌ Disabled' }
    $descCurrent = if ($null -ne $adComp.Description) { [string]$adComp.Description } else { '' }
    $managedBy   = if ($null -ne $adComp.ManagedBy)   { [string]$adComp.ManagedBy } else { '' }
    $compOU      = CM-ParseOU -DN $adComp.DistinguishedName
    
    # ── User Match Logic ──────────────────────────────────────────────────────
    # Rule: if EmpID is present → match by AD Description only.
    #       if EmpID is blank   → match by Email (full UPN, full email, or prefix without @domain)
    $adUser      = $null
    $matchMethod = ""
    $empStr      = if ($null -ne $EmpID) { [string]$EmpID.Trim() } else { '' }

    if ($empStr -ne '') {
        # EmpID present — match against AD user Description field only
        if ($Users.ContainsKey($empStr)) {
            $adUser = $Users[$empStr]; $matchMethod = "EmpID→Description"
        }
    } else {
        # EmpID blank — fall back to Email matching
        $emailStr = if ($null -ne $Email) { $Email.Trim() } else { '' }

        if ($emailStr -ne '') {
            # 1. Try full email/UPN as-is (e.g. john.smith@domain.com)
            $mailKey = $emailStr.ToLower()
            if ($UsersByMail.ContainsKey($mailKey)) {
                $adUser = $UsersByMail[$mailKey]; $matchMethod = "Email"
            }

            # 2. Try as UPN against AD UPN index
            if ($null -eq $adUser -and $UsersByMail.ContainsKey($mailKey)) {
                $adUser = $UsersByMail[$mailKey]; $matchMethod = "UPN"
            }

            # 3. Strip @domain suffix and match prefix against sAMAccountName
            #    Handles both "john.smith@domain.com" and plain "john.smith"
            if ($null -eq $adUser) {
                $prefix = if ($emailStr -match '^([^@]+)@') { $Matches[1].ToLower() } else { $emailStr.ToLower() }
                if ($UsersBySAM.ContainsKey($prefix)) {
                    $adUser = $UsersBySAM[$prefix]; $matchMethod = "Email Prefix→SAM"
                }
            }
        }
    }

    # Handling the "In AD Only" Orphan case
    if ($ADCompOverride) {
        $isStale = $null -ne $adComp.LastLogon -and $adComp.LastLogon -lt (Get-Date).AddDays(-180)
        $status = if ($isStale) { $Script:CM_Status.InADStale } else { $Script:CM_Status.InADOnly }
        
        # Determine why it's an orphan
        $reason = "Not found in imported file"
        if ($descCurrent -eq '') { $reason += " + Description missing" }
        if ($managedBy -eq '')   { $reason += " + ManagedBy missing" }

        return [PSCustomObject]@{
            Status = $status; StatusLabel = $Script:CM_StatusLabel[$status]
            TAG = $adComp.Name; ComputerName = $compName; ComputerEnabled = $compEnabled
            SDStatus = ''; EmpID = ''; DisplayName = $managedBy; SAM = ''
            DescCurrent = $descCurrent; DescNew = ''; ManagedByCurrent = $managedBy
            Office = ''; CurrentOU = $compOU; ResultNote = $reason
            _UserDN = ''; _CompDN = $adComp.DistinguishedName
        }
    }

    # Standard Excel-Row Result
    if ($null -eq $adUser) {
        return [PSCustomObject]@{
            Status = $Script:CM_Status.NoUser; StatusLabel = $Script:CM_StatusLabel.NoUser
            TAG = $TAG; ComputerName = $compName; ComputerEnabled = $compEnabled; SDStatus = $SDStatus
            EmpID = $empStr; DisplayName = ''; SAM = ''; DescCurrent = $descCurrent; DescNew = ''
            ManagedByCurrent = $managedBy; Office = ''; CurrentOU = $compOU
            ResultNote = "No AD user found for ID='$empStr' or Email='$Email'"; _UserDN = ''; _CompDN = $adComp.DistinguishedName
        }
    }

    $userDN      = [string]$adUser.DistinguishedName
    $displayName = [string]$adUser.DisplayName
    $sam         = [string]$adUser.SamAccountName
    $office      = [string]$adUser.physicalDeliveryOfficeName
    
    # Back-populate EmpID if found via email
    $finalEmpID = if ($matchMethod -match 'Email') { $sam } else { $empStr }

    # Track multi-computer
    if (-not $UserCompMap.ContainsKey($userDN)) { $UserCompMap[$userDN] = [System.Collections.ArrayList]::new() }
    [void]$UserCompMap[$userDN].Add($compName)

    # Final logic for update necessity
    # Compare DN to DN — most reliable, no CN/DisplayName parsing needed
    $needsUpdate = $false
    if ($descCurrent -ne $displayName) { $needsUpdate = $true }
    if ($managedBy -ne $userDN) { $needsUpdate = $true }

    $statusKey = if ($needsUpdate) { $Script:CM_Status.Pending } else { $Script:CM_Status.NoAction }
    
    return [PSCustomObject]@{
        Status           = $statusKey
        StatusLabel      = $Script:CM_StatusLabel[$statusKey]
        TAG              = $TAG
        ComputerName     = $compName
        ComputerEnabled  = $compEnabled
        SDStatus         = $SDStatus
        EmpID            = $finalEmpID
        DisplayName      = $displayName
        SAM              = $sam
        DescCurrent      = $descCurrent
        DescNew          = $displayName
        ManagedByCurrent = $managedBy
        ManagedByNew     = $sam
        Office           = $office
        CurrentOU        = $compOU
        ResultNote       = if ($matchMethod -ne "Description") { "Matched via $matchMethod" } else { '' }
        _UserDN          = $userDN
        _CompDN          = $adComp.DistinguishedName
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
        Write-Log "ComputerMapper [WHATIF]: would set '$ComputerDN' Description='$NewDescription' ManagedBy='$NewManagedBy'" -Source 'ComputerMapper'
        return @{ OK = $true; Message = "[WhatIf] Success" }
    }

    try {
        if ([string]::IsNullOrWhiteSpace($ComputerDN)) { throw "Missing Computer DN" }
        
        $params = @{
            Identity    = $ComputerDN
            Server      = $DCName
            Credential  = $Credential
            ErrorAction = 'Stop'
            WhatIf      = $false  # Explicitly disable AD preference for WhatIf in Live mode
        }
        
        # Only add parameters if they have actual values
        if ($null -ne $NewDescription) { $params['Description'] = $NewDescription }
        if ($null -ne $NewManagedBy)   { $params['ManagedBy']   = $NewManagedBy   }

        Set-ADComputer @params
        Write-Log "ComputerMapper: SUCCESS updated '$ComputerDN' on $DCName" -Source 'ComputerMapper'
        return @{ OK = $true; Message = 'Updated successfully' }

    } catch {
        $err = $_.Exception.Message
        Write-Log "ComputerMapper: FAILURE updating '$ComputerDN' - $err" -Level ERROR -Source 'ComputerMapper'
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

    Write-Log "ComputerMapper: exporting to '$OutputPath'" -Source 'ComputerMapper'

    # Convert ObservableCollection to a standard array for speed and grouping
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

    # Columns to export (exclude internal _ properties)
    $exportProps = @('StatusLabel','TAG','ComputerName','SDStatus','ComputerEnabled','EmpID',
                     'DisplayName','SAM','DescCurrent','DescNew','ManagedByCurrent',
                     'Office','CurrentOU','ResultNote')

    try {
        # Ensure target file is NOT locked and clear it for a fresh export
        if (Test-Path $OutputPath) {
            try { 
                Remove-Item $OutputPath -Force -ErrorAction Stop 
            } catch {
                throw "Target file is in use. Please close it in Excel first."
            }
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

            Write-Log "ComputerMapper: export complete (ImportExcel) — $OutputPath" -Source 'ComputerMapper'
            return @{ OK = $true; Path = $OutputPath }

        } else {
            # Fallback: export as CSV with sheet name prefix
            $csvPath = $OutputPath -replace '\.xlsx$','.csv'
            $groups['All Data'] | Select-Object $exportProps |
                Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Log "ComputerMapper: export as CSV (ImportExcel not installed) — $csvPath" -Level WARN -Source 'ComputerMapper'
            return @{ OK = $true; Path = $csvPath; Note = 'Saved as CSV — install ImportExcel for XLSX' }
        }

    } catch {
        $err = $_.Exception.Message
        Write-Log "ComputerMapper: export failed — $err" -Level ERROR -Source 'ComputerMapper'
        return @{ OK = $false; ErrorMsg = $err }
    }
}