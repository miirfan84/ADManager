# =============================================================================
# Plugins\DnsManager\Functions.ps1  —  AD Manager v2.0
# Pure logic — no WPF. Called from Handlers.ps1 and background runspaces.
#
# All DNS operations target the DC via -ComputerName $DC over WinRM (port 5985).
# Requires the DnsServer PowerShell module to be installed on the DC.
# =============================================================================

# ── Status constants ──────────────────────────────────────────────────────────
$Script:DNS_Status = @{
    New           = 'New'
    NoAction      = 'NoAction'
    Update        = 'Update'
    Convert       = 'Convert'
    WildcardMatch = 'WildcardMatch'
    MultiConflict = 'MultiConflict'
    Done          = 'Done'
    Failed        = 'Failed'
    NotFound      = 'NotFound'
}

$Script:DNS_StatusLabel = @{
    New           = '➕ New'
    NoAction      = '✅ No Action'
    Update        = '🔄 Update'
    Convert       = '🔁 Convert'
    WildcardMatch = '❓ Wildcard Match'
    MultiConflict = '⚠ Multi-Record'
    Done          = '✅ Done'
    Failed        = '❌ Failed'
    NotFound      = '❌ Not Found'
}

# ── DNS-ParseInputLine ────────────────────────────────────────────────────────
# Parses one raw line from the Input_Box into a structured record specification.
#
# Parameters:
#   $Line      — one raw line from the Input_Box
#   $ZoneName  — currently selected zone (e.g. "alec.com")
#
# Returns a hashtable:
#   @{ OK=[bool]; Hostname=[string]; Type=[string]; Value=[string]; ErrorMsg=[string] }
# =============================================================================
function DNS-ParseInputLine {
    param(
        [string]$Line,
        [string]$ZoneName
    )

    # Step 1: Trim; skip blank lines and comment lines
    $trimmed = $Line.Trim()
    if ($trimmed -eq '' -or $trimmed.StartsWith('#')) {
        return @{ OK = $false; Hostname = ''; Type = ''; Value = ''; ErrorMsg = '' }
    }

    # Step 2: Split on whitespace
    $tokens = $trimmed -split '\s+'
    if ($tokens.Count -eq 0) {
        return @{ OK = $false; Hostname = ''; Type = ''; Value = ''; ErrorMsg = 'Parse error' }
    }

    # Step 3: Token[0] = raw name; strip zone suffix if present (case-insensitive)
    $rawName = $tokens[0]
    $suffix  = ".$ZoneName"
    if ($rawName -match [regex]::Escape($suffix) + '$' -and
        $rawName.ToLower().EndsWith($suffix.ToLower())) {
        $hostname = $rawName.Substring(0, $rawName.Length - $suffix.Length)
    } else {
        $hostname = $rawName
    }

    # Step 4: Token[1] = value (if provided)
    $value = ''
    $type  = ''
    if ($tokens.Count -ge 2) {
        $value = $tokens[1]
        if ($value -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $type = 'A'
        } else {
            $type = 'CNAME'
        }
    }

    # Step 5: Return success
    return @{
        OK       = $true
        Hostname = $hostname
        Type     = $type
        Value    = $value
        ErrorMsg = ''
    }
}

# ── DNS-ClassifyRow ───────────────────────────────────────────────────────────
# Classifies a parsed input row against existing DNS records and returns a
# single status key string.
#
# Parameters:
#   $ParsedRow    — hashtable output of DNS-ParseInputLine
#   $ExistingRecs — array of record objects from DNS-FetchRecords
#
# Returns one of: "New" | "WildcardMatch" | "MultiConflict" | "Convert" |
#                 "NoAction" | "Update"
#
# Priority order is fixed — see logic below.
# =============================================================================
function DNS-ClassifyRow {
    param(
        [hashtable]$ParsedRow,
        [object[]]$ExistingRecs
    )

    # 1. No existing records → brand-new hostname
    if (-not $ExistingRecs -or $ExistingRecs.Count -eq 0) {
        if ([string]::IsNullOrWhiteSpace($ParsedRow.Value)) {
            return 'NotFound'
        }
        return 'New'
    }

    # 2. Wildcard-match row (found by prefix search, not in input) → reference only
    if ($ParsedRow.IsWildcard -eq $true) {
        return 'WildcardMatch'
    }

    # 3. Multiple existing records and all are A records → conflict requiring user decision
    if ($ExistingRecs.Count -gt 1) {
        $allA = ($ExistingRecs | Where-Object { $_.Type -ne 'A' }).Count -eq 0
        if ($allA) {
            if ([string]::IsNullOrWhiteSpace($ParsedRow.Value)) {
                return 'NoAction'
            }
            return 'MultiConflict'
        }
    }

    # From here on we deal with a single existing record
    $existing = $ExistingRecs[0]

    # If no value was provided (search only), then no action required
    if ([string]::IsNullOrWhiteSpace($ParsedRow.Value)) {
        return 'NoAction'
    }

    # 4. Different record type → conversion required
    if ($existing.Type -ne $ParsedRow.Type) {
        return 'Convert'
    }

    # 5. Same type, same value → nothing to do
    if ($existing.ExistingValue -eq $ParsedRow.Value) {
        return 'NoAction'
    }

    # 6. Same type, different value → update required
    return 'Update'
}

# ── DNS-FetchRecords ──────────────────────────────────────────────────────────
# Queries the DC for existing DNS records matching the supplied hostnames.
# Also performs a zone-wide wildcard (prefix) pass to surface related records.
#
# Parameters:
#   $Hostnames   — distinct list of short hostnames from parsed input
#   $Zone        — DNS forward lookup zone name (e.g. "alec.com")
#   $DC          — domain controller name (ComputerName for DNS cmdlets)
#   $Credential  — PSCredential (may be $null for current-user context)
#
# Returns an array of record objects:
#   @{
#       Hostname      = [string]   — short hostname (zone-relative)
#       Type          = [string]   — "A" or "CNAME"
#       ExistingValue = [string]   — IP address or CNAME target
#       TTL           = [string]   — TTL display string
#       IsWildcard    = [bool]     — $true if found by prefix match only
#   }
#
# On per-hostname error: returns a failure record with Type="Error" and
# ExistingValue containing the exception message; processing continues.
#
# Requirements: 4.1–4.5, 4.7
# =============================================================================
function DNS-FetchRecords {
    param(
        [string[]]$Hostnames,
        [string]$Zone,
        [string]$DC,
        [PSCredential]$Credential
    )

    $results = [System.Collections.ArrayList]::new()

    $icParams = @{
        ComputerName = $DC
        ErrorAction  = 'Stop'
        ArgumentList = @($Zone, $Hostnames)
        ScriptBlock  = {
            param($zone, [string[]]$hostnames)
            
            $out = [System.Collections.ArrayList]::new()
            $queried = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($h in $hostnames) {
                if (-not [string]::IsNullOrWhiteSpace($h)) { [void]$queried.Add($h) }
            }

            if ($queried.Count -eq 0) { return @() }

            try {
                $allRecs = Get-DnsServerResourceRecord -ZoneName $zone -ErrorAction Stop
                
                foreach ($r in $allRecs) {
                    $rName = $r.HostName
                    $isMatch = $false
                    $isExact = $false

                    foreach ($h in $queried) {
                        if ($rName -eq $h) {
                            $isMatch = $true
                            $isExact = $true
                            break
                        }
                        if ($rName -like "*$h*") {
                            $isMatch = $true
                        }
                    }

                    if ($isMatch) {
                        $type = $r.RecordType
                        $val  = ''
                        if ($type -eq 'A')         { $val = [string]$r.RecordData.IPv4Address }
                        elseif ($type -eq 'CNAME') { $val = $r.RecordData.HostNameAlias.TrimEnd('.') }
                        elseif ($type -eq 'TXT')   { $val = [string]$r.RecordData.DescriptiveText }
                        elseif ($type -eq 'MX')    { $val = "$($r.RecordData.Preference) $($r.RecordData.MailExchange.TrimEnd('.'))" }
                        elseif ($type -eq 'NS')    { $val = $r.RecordData.NameServer.TrimEnd('.') }
                        elseif ($type -eq 'SOA')   { $val = $r.RecordData.PrimaryServer.TrimEnd('.') }
                        elseif ($type -eq 'AAAA')  { $val = [string]$r.RecordData.IPv6Address }
                        else                       { $val = "[$type Record]" }
                        $ttl = if ($r.TimeToLive) { $r.TimeToLive.ToString() } else { '' }

                        [void]$out.Add(@{
                            Hostname   = $rName
                            Type       = $type
                            Value      = $val
                            TTL        = $ttl
                            IsWildcard = (-not $isExact)
                            ErrorMsg   = ''
                        })
                    }
                }
            } catch {
                [void]$out.Add(@{ ErrorMsg = $_.Exception.Message })
            }

            return $out.ToArray()
        }
    }
    if ($null -ne $Credential) { $icParams['Credential'] = $Credential }

    try {
        $rawRecords = Invoke-Command @icParams

        foreach ($rec in $rawRecords) {
            if ($null -ne $rec.ErrorMsg -and $rec.ErrorMsg -ne '') {
                # Add an error row for each queried hostname
                foreach ($h in $Hostnames) {
                    if ([string]::IsNullOrWhiteSpace($h)) { continue }
                    [void]$results.Add(@{
                        Hostname      = $h
                        Type          = 'Error'
                        ExistingValue = $rec.ErrorMsg
                        TTL           = ''
                        IsWildcard    = $false
                    })
                }
                break
            } else {
                [void]$results.Add(@{
                    Hostname      = $rec.Hostname
                    Type          = $rec.Type
                    ExistingValue = $rec.Value
                    TTL           = $rec.TTL
                    IsWildcard    = [bool]$rec.IsWildcard
                })
            }
        }
    } catch {
        # WinRM failure or overall invoke failure
        foreach ($h in $Hostnames) {
            if ([string]::IsNullOrWhiteSpace($h)) { continue }
            [void]$results.Add(@{
                Hostname      = $h
                Type          = 'Error'
                ExistingValue = $_.Exception.Message
                TTL           = ''
                IsWildcard    = $false
            })
        }
    }

    return $results.ToArray()
}

# ── DNS-ApplyRecord ───────────────────────────────────────────────────────────
# Applies a single planned DNS change to the DC.
#
# Parameters:
#   $Row        — grid row PSCustomObject (Status, Type, Name, NewValue, _Zone,
#                 _MultiAction, ExistingValue, etc.)
#   $DC         — domain controller name (ComputerName for DNS cmdlets)
#   $Credential — PSCredential (may be $null for current-user context)
#   $WhatIf     — when $true, returns a preview message without executing any cmdlets
#
# Returns: @{ OK=[bool]; Message=[string] }
#
# Dispatch table (by $Row.Status and $Row._MultiAction):
#   New    + Type=A       → Add-DnsServerResourceRecordA
#   New    + Type=CNAME   → Add-DnsServerResourceRecordCName
#   Update + Type=A       → Remove-DnsServerResourceRecord → Add-DnsServerResourceRecordA
#   Update + Type=CNAME   → Set-DnsServerResourceRecord
#   Convert (A→CNAME)     → Remove-DnsServerResourceRecord (A) → Add-DnsServerResourceRecordCName
#   Convert (CNAME→A)     → Remove-DnsServerResourceRecord (CNAME) → Add-DnsServerResourceRecordA
#   MultiConflict ReplaceAll    → Remove all existing A records → Add-DnsServerResourceRecordA
#   MultiConflict AddAdditional → Add-DnsServerResourceRecordA (existing preserved)
#   MultiConflict NoAction      → @{ OK=$true; Message="Skipped" }
#
# All cmdlets use -ZoneName $Row._Zone -ComputerName $DC; no -TimeToLive (zone default).
# Requirements: 7.1–7.6
# =============================================================================
function DNS-ApplyRecord {
    param(
        [PSCustomObject]$Row,
        [string]$DC,
        [PSCredential]$Credential,
        [bool]$WhatIf
    )

    # ── WhatIf path — no cmdlets called ──────────────────────────────────────
    if ($WhatIf) {
        $preview = switch ($Row.Status) {
            'New' {
                if ($Row.Type -eq 'A') {
                    "[WhatIf] Would Add A record: $($Row.Name) → $($Row.NewValue) in zone $($Row._Zone)"
                } else {
                    "[WhatIf] Would Add CNAME record: $($Row.Name) → $($Row.NewValue) in zone $($Row._Zone)"
                }
            }
            'Update' {
                if ($Row.Type -eq 'A') {
                    "[WhatIf] Would Update A record: $($Row.Name) $($Row.ExistingValue) → $($Row.NewValue) in zone $($Row._Zone)"
                } else {
                    "[WhatIf] Would Update CNAME record: $($Row.Name) → $($Row.NewValue) in zone $($Row._Zone)"
                }
            }
            'Convert' {
                if ($Row.Type -eq 'CNAME') {
                    "[WhatIf] Would Convert A→CNAME: $($Row.Name) in zone $($Row._Zone)"
                } else {
                    "[WhatIf] Would Convert CNAME→A: $($Row.Name) → $($Row.NewValue) in zone $($Row._Zone)"
                }
            }
            'MultiConflict' {
                switch ($Row._MultiAction) {
                    'ReplaceAll'    { "[WhatIf] Would Replace All A records for $($Row.Name) with $($Row.NewValue) in zone $($Row._Zone)" }
                    'AddAdditional' { "[WhatIf] Would Add Additional A record: $($Row.Name) → $($Row.NewValue) in zone $($Row._Zone)" }
                    'NoAction'      { "[WhatIf] Would Skip (NoAction) $($Row.Name) in zone $($Row._Zone)" }
                    default         { "[WhatIf] Would process $($Row.Name) (MultiConflict/$($Row._MultiAction)) in zone $($Row._Zone)" }
                }
            }
            default { "[WhatIf] Would process $($Row.Name) (Status=$($Row.Status)) in zone $($Row._Zone)" }
        }
        return @{ OK = $true; Message = $preview }
    }

    # ── Live path — wrapped in try/catch ─────────────────────────────────────
    try {
        $zone     = $Row._Zone
        $name     = $Row.Name
        $newVal   = $Row.NewValue
        $status   = $Row.Status
        $type     = $Row.Type
        $multiAct = $Row._MultiAction

        $icParams = @{
            ComputerName = $DC
            ErrorAction  = 'Stop'
            ArgumentList = @($zone, $name, $newVal, $status, $type, $multiAct)
            ScriptBlock  = {
                param($zone, $name, $newVal, $status, $type, $multiAct)

                switch ($status) {

                    'New' {
                        if ($type -eq 'A') {
                            Add-DnsServerResourceRecordA -Name $name -ZoneName $zone -IPv4Address $newVal -ErrorAction Stop
                            return "Added A record $name → $newVal"
                        } else {
                            Add-DnsServerResourceRecordCName -Name $name -ZoneName $zone -HostNameAlias $newVal -ErrorAction Stop
                            return "Added CNAME record $name → $newVal"
                        }
                    }

                    'Update' {
                        if ($type -eq 'A') {
                            Remove-DnsServerResourceRecord -ZoneName $zone -Name $name -RRType 'A' -Force -ErrorAction Stop
                            Add-DnsServerResourceRecordA -Name $name -ZoneName $zone -IPv4Address $newVal -ErrorAction Stop
                            return "Updated A record $name → $newVal"
                        } else {
                            $old = Get-DnsServerResourceRecord -Name $name -ZoneName $zone -RRType 'CNAME' -ErrorAction Stop
                            $new = $old.Clone()
                            $new.RecordData.HostNameAlias = "$newVal."
                            Set-DnsServerResourceRecord -ZoneName $zone -OldInputObject $old -NewInputObject $new -ErrorAction Stop
                            return "Updated CNAME record $name → $newVal"
                        }
                    }

                    'Convert' {
                        if ($type -eq 'CNAME') {
                            Remove-DnsServerResourceRecord -ZoneName $zone -Name $name -RRType 'A' -Force -ErrorAction Stop
                            Add-DnsServerResourceRecordCName -Name $name -ZoneName $zone -HostNameAlias $newVal -ErrorAction Stop
                            return "Converted A→CNAME: $name → $newVal"
                        } else {
                            Remove-DnsServerResourceRecord -ZoneName $zone -Name $name -RRType 'CNAME' -Force -ErrorAction Stop
                            Add-DnsServerResourceRecordA -Name $name -ZoneName $zone -IPv4Address $newVal -ErrorAction Stop
                            return "Converted CNAME→A: $name → $newVal"
                        }
                    }

                    'MultiConflict' {
                        switch ($multiAct) {
                            'ReplaceAll' {
                                $existing = Get-DnsServerResourceRecord -Name $name -ZoneName $zone -RRType 'A' -ErrorAction Stop
                                foreach ($rec in $existing) {
                                    Remove-DnsServerResourceRecord -ZoneName $zone -InputObject $rec -Force -ErrorAction Stop
                                }
                                Add-DnsServerResourceRecordA -Name $name -ZoneName $zone -IPv4Address $newVal -ErrorAction Stop
                                return "Replaced all A records for $name with $newVal"
                            }
                            'AddAdditional' {
                                Add-DnsServerResourceRecordA -Name $name -ZoneName $zone -IPv4Address $newVal -ErrorAction Stop
                                return "Added additional A record $name → $newVal"
                            }
                            'NoAction' { return 'Skipped' }
                            default    { throw "Unknown _MultiAction: $multiAct" }
                        }
                    }

                    default { throw "Unhandled row status: $status" }
                }
            }
        }
        if ($null -ne $Credential) { $icParams['Credential'] = $Credential }

        $msg = Invoke-Command @icParams
        Write-Log "DnsManager: SUCCESS — $msg" -Source 'DnsManager'
        return @{ OK = $true; Message = $msg }

    } catch {
        return @{ OK = $false; Message = $_.Exception.Message }
    }
}

# ── DNS-DeleteRecord ──────────────────────────────────────────────────────────
# Deletes a single DNS record from the DC.
#
# Parameters:
#   $Row        — grid row PSCustomObject (Name, Type, _Zone)
#   $DC         — domain controller name (ComputerName for DNS cmdlets)
#   $Credential — PSCredential (may be $null for current-user context)
#   $WhatIf     — when $true, returns a preview message without executing any cmdlets
#
# Returns: @{ OK=[bool]; Message=[string] }
#
# Requirements: 8.1–8.5
# =============================================================================
function DNS-DeleteRecord {
    param(
        [PSCustomObject]$Row,
        [string]$DC,
        [PSCredential]$Credential,
        [bool]$WhatIf
    )

    # ── WhatIf path — no cmdlets called ──────────────────────────────────────
    if ($WhatIf) {
        return @{ OK = $true; Message = "[WhatIf] Would Delete $($Row.Name) ($($Row.Type))" }
    }

    # ── Live path — wrapped in try/catch ─────────────────────────────────────
    try {
        $icParams = @{
            ComputerName = $DC
            ErrorAction  = 'Stop'
            ArgumentList = @($Row._Zone, $Row.Name, $Row.Type)
            ScriptBlock  = {
                param($zone, $name, $type)
                Remove-DnsServerResourceRecord -ZoneName $zone -Name $name -RRType $type -Force -ErrorAction Stop
            }
        }
        if ($null -ne $Credential) { $icParams['Credential'] = $Credential }

        Invoke-Command @icParams
        return @{ OK = $true; Message = "Deleted $($Row.Type) record $($Row.Name) from zone $($Row._Zone)" }
    } catch {
        return @{ OK = $false; Message = $_.Exception.Message }
    }
}

# ── DNS-ExportToExcel ─────────────────────────────────────────────────────────
# Exports DNS Manager grid data to XLSX using the ImportExcel module.
#
# Parameters:
#   $Rows      — array of grid row objects (PSCustomObject)
#   $FilePath  — full path to the output XLSX file
#
# Returns: @{ OK=[bool]; ErrorMsg=[string] }
#
# Columns exported: StatusLabel, Name, Type, ExistingValue, NewValue, TTL, Match, ResultNote
#
# On failure: returns OK=$false with ErrorMsg; does not leave a partial file.
# Requirements: 9.1–9.3
# =============================================================================
function DNS-ExportToExcel {
    param(
        [object[]]$Rows,
        [string]$FilePath
    )

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return @{ OK = $false; ErrorMsg = "No data to export" }
    }

    # Columns to export (all visible grid columns)
    $exportProps = @('StatusLabel', 'Name', 'Type', 'ExistingValue', 'NewValue', 'TTL', 'Match', 'ResultNote')

    try {
        # Ensure target file is NOT locked and clear it for a fresh export
        if (Test-Path $FilePath) {
            try { 
                Remove-Item $FilePath -Force -ErrorAction Stop 
            } catch {
                throw "Target file is in use. Please close it in Excel first."
            }
        }

        # Import ImportExcel module
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction Stop

            # Export all rows to a single sheet
            $Rows | Select-Object -Property $exportProps | 
                   Export-Excel -Path $FilePath -WorksheetName 'DNS Records' -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium2

            return @{ OK = $true; ErrorMsg = '' }

        } else {
            # Fallback: export as CSV (ImportExcel not available)
            $csvPath = [System.IO.Path]::ChangeExtension($FilePath, '.csv')
            $Rows | Select-Object -Property $exportProps |
                Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            return @{ OK = $true; ErrorMsg = "Saved as CSV (install ImportExcel module for XLSX support): $csvPath" }
        }

    } catch {
        # Clean up any partial file that may have been written
        if (Test-Path $FilePath) {
            try { Remove-Item $FilePath -Force -ErrorAction SilentlyContinue } catch {}
        }
        return @{ OK = $false; ErrorMsg = $_.Exception.Message }
    }
}

# ── DNS-FilterZones ───────────────────────────────────────────────────────────
# Filters a list of DNS zone objects to return only forward lookup zones.
#
# Parameters:
#   $Zones — array of zone objects from Get-DnsServerZone
#
# Returns only zones where:
#   - Name does NOT end with '.in-addr.arpa' or '.ip6.arpa'
#   - IsReverseLookupZone is NOT $true
#
# Requirements: 2.2
# =============================================================================
function DNS-FilterZones {
    param(
        [object[]]$Zones
    )

    if ($null -eq $Zones -or $Zones.Count -eq 0) {
        return @()
    }

    return @($Zones | Where-Object {
        $_.ZoneName -notmatch '\.in-addr\.arpa$' -and
        $_.ZoneName -notmatch '\.ip6\.arpa$' -and
        $_.IsReverseLookupZone -ne $true
    })
}
