# =============================================================================
# Plugins\DnsManager\Tests\DNS-WhatIf.Property.Tests.ps1
#
# Property-Based Tests for DNS-ApplyRecord and DNS-DeleteRecord (Pester)
#
# Property 6: WhatIf mode prevents all DNS writes and prefixes all messages
#   Validates: Requirements 7.1, 8.3, 10.1, 10.4
#
# Run with:
#   Invoke-Pester -Path .\Plugins\DnsManager\Tests\DNS-WhatIf.Property.Tests.ps1
# =============================================================================

# Dot-source the function under test
$FunctionsPath = Join-Path $PSScriptRoot '..' 'Functions.ps1'
. $FunctionsPath

# ---------------------------------------------------------------------------
# Stub definitions for DnsServer cmdlets
# Pester 3.x Mock requires the command to exist in the session before it can
# be mocked. Because the DnsServer module is not installed on the test machine,
# we define no-op stubs here so that Mock can intercept them.
# ---------------------------------------------------------------------------
if (-not (Get-Command Add-DnsServerResourceRecordA    -ErrorAction SilentlyContinue)) {
    function Add-DnsServerResourceRecordA    { param([Parameter(ValueFromRemainingArguments)]$args) }
}
if (-not (Get-Command Add-DnsServerResourceRecordCName -ErrorAction SilentlyContinue)) {
    function Add-DnsServerResourceRecordCName { param([Parameter(ValueFromRemainingArguments)]$args) }
}
if (-not (Get-Command Set-DnsServerResourceRecord     -ErrorAction SilentlyContinue)) {
    function Set-DnsServerResourceRecord     { param([Parameter(ValueFromRemainingArguments)]$args) }
}
if (-not (Get-Command Remove-DnsServerResourceRecord  -ErrorAction SilentlyContinue)) {
    function Remove-DnsServerResourceRecord  { param([Parameter(ValueFromRemainingArguments)]$args) }
}

# ---------------------------------------------------------------------------
# Generator helpers
# ---------------------------------------------------------------------------

function New-RandomLabel {
    <#
    .SYNOPSIS
        Returns a random short DNS label (letters and digits only, 1–12 chars).
    #>
    $chars  = 'abcdefghijklmnopqrstuvwxyz0123456789'
    $length = (Get-Random -Minimum 1 -Maximum 13)
    $label  = -join (1..$length | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
    return $label
}

function New-RandomIPv4 {
    <#
    .SYNOPSIS
        Returns a random valid IPv4 address string.
    #>
    $o1 = Get-Random -Minimum 1   -Maximum 256
    $o2 = Get-Random -Minimum 0   -Maximum 256
    $o3 = Get-Random -Minimum 0   -Maximum 256
    $o4 = Get-Random -Minimum 0   -Maximum 256
    return "$o1.$o2.$o3.$o4"
}

function New-RandomCNAMETarget {
    <#
    .SYNOPSIS
        Returns a random CNAME target (hostname or FQDN).
    #>
    $label1 = New-RandomLabel
    $label2 = New-RandomLabel
    $tlds   = @('com', 'net', 'org', 'local', 'internal', 'corp', 'io')
    $tld    = $tlds[(Get-Random -Minimum 0 -Maximum $tlds.Count)]
    return "$label1.$label2.$tld"
}

function New-RandomZoneName {
    <#
    .SYNOPSIS
        Returns a random zone name in the form "label.tld" (e.g. "contoso.com").
    #>
    $tlds  = @('com', 'net', 'org', 'local', 'internal', 'corp', 'io')
    $label = New-RandomLabel
    $tld   = $tlds[(Get-Random -Minimum 0 -Maximum $tlds.Count)]
    return "$label.$tld"
}

function New-RandomApplyRow {
    <#
    .SYNOPSIS
        Creates a random grid row PSCustomObject suitable for DNS-ApplyRecord.
        Covers all Status/Type/MultiAction combinations that DNS-ApplyRecord handles.
    #>

    # All actionable statuses (excludes NoAction, WildcardMatch, Done, Failed)
    $statuses = @('New', 'Update', 'Convert', 'MultiConflict')
    $status   = $statuses[(Get-Random -Minimum 0 -Maximum $statuses.Count)]

    $name          = New-RandomLabel
    $zone          = New-RandomZoneName
    $existingValue = New-RandomIPv4
    $newValue      = New-RandomIPv4
    $type          = 'A'
    $multiAction   = $null

    switch ($status) {

        'New' {
            # New A or CNAME
            $type = @('A', 'CNAME')[(Get-Random -Minimum 0 -Maximum 2)]
            if ($type -eq 'A') {
                $newValue = New-RandomIPv4
            } else {
                $newValue = New-RandomCNAMETarget
            }
            $existingValue = ''
        }

        'Update' {
            # Update A or CNAME
            $type = @('A', 'CNAME')[(Get-Random -Minimum 0 -Maximum 2)]
            if ($type -eq 'A') {
                $existingValue = New-RandomIPv4
                $newValue      = New-RandomIPv4
            } else {
                $existingValue = New-RandomCNAMETarget
                $newValue      = New-RandomCNAMETarget
            }
        }

        'Convert' {
            # Convert A→CNAME or CNAME→A
            # $Row.Type is the TARGET type (what we're converting TO)
            $type = @('A', 'CNAME')[(Get-Random -Minimum 0 -Maximum 2)]
            if ($type -eq 'CNAME') {
                # Converting A → CNAME
                $existingValue = New-RandomIPv4
                $newValue      = New-RandomCNAMETarget
            } else {
                # Converting CNAME → A
                $existingValue = New-RandomCNAMETarget
                $newValue      = New-RandomIPv4
            }
        }

        'MultiConflict' {
            # MultiConflict with ReplaceAll, AddAdditional, or NoAction
            $type        = 'A'
            $multiActions = @('ReplaceAll', 'AddAdditional', 'NoAction')
            $multiAction  = $multiActions[(Get-Random -Minimum 0 -Maximum $multiActions.Count)]
            $existingValue = New-RandomIPv4
            $newValue      = New-RandomIPv4
        }
    }

    $row = [PSCustomObject]@{
        Status        = $status
        StatusLabel   = $status
        Name          = $name
        Type          = $type
        ExistingValue = $existingValue
        NewValue      = $newValue
        TTL           = '01:00:00'
        Match         = 'Exact'
        ResultNote    = ''
        IsWildcard    = $false
        _Zone         = $zone
        _MultiAction  = $multiAction
    }

    return $row
}

function New-RandomDeleteRow {
    <#
    .SYNOPSIS
        Creates a random grid row PSCustomObject suitable for DNS-DeleteRecord.
        Covers A and CNAME record types.
    #>
    $type = @('A', 'CNAME')[(Get-Random -Minimum 0 -Maximum 2)]
    $existingValue = if ($type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }

    return [PSCustomObject]@{
        Status        = 'Done'
        StatusLabel   = '✅ Done'
        Name          = New-RandomLabel
        Type          = $type
        ExistingValue = $existingValue
        NewValue      = ''
        TTL           = '01:00:00'
        Match         = 'Exact'
        ResultNote    = ''
        IsWildcard    = $false
        _Zone         = New-RandomZoneName
        _MultiAction  = $null
    }
}

# ---------------------------------------------------------------------------
# Property 6: WhatIf mode prevents all DNS writes and prefixes all messages
# Validates: Requirements 7.1, 8.3, 10.1, 10.4
# ---------------------------------------------------------------------------

Describe 'DNS-ApplyRecord — Property 6: WhatIf mode prevents all DNS writes and prefixes all messages' -Tag 'Feature: dns-manager', 'Property 6: WhatIf mode prevents all DNS writes and prefixes all messages' {

    # Mock all four DNS write cmdlets before each test so we can assert zero calls
    BeforeEach {
        Mock Add-DnsServerResourceRecordA    { }
        Mock Add-DnsServerResourceRecordCName { }
        Mock Set-DnsServerResourceRecord     { }
        Mock Remove-DnsServerResourceRecord  { }
    }

    It 'returns a message starting with "[WhatIf]" and never calls any DNS write cmdlet (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            $row = New-RandomApplyRow

            # Reset mock call counts before each call
            # (Pester tracks cumulative calls within a BeforeEach scope;
            #  we track failures manually and assert at the end)

            $result = DNS-ApplyRecord -Row $row -DC 'dc01' -Credential $null -WhatIf $true

            # ----------------------------------------------------------------
            # Assert 1: result must be a hashtable with OK and Message keys
            # ----------------------------------------------------------------
            if ($null -eq $result) {
                $failures += "Iteration $i (Status='$($row.Status)', Type='$($row.Type)', MultiAction='$($row._MultiAction)') — DNS-ApplyRecord returned `$null"
                continue
            }

            # ----------------------------------------------------------------
            # Assert 2: Message must start with "[WhatIf]"
            # Use -match with regex (escape brackets) instead of -like
            # because -like treats [ ] as wildcard metacharacters.
            # ----------------------------------------------------------------
            if (-not ($result.Message -match '^\[WhatIf\]')) {
                $failures += "Iteration $i (Status='$($row.Status)', Type='$($row.Type)', MultiAction='$($row._MultiAction)') — Message does not start with '[WhatIf]': '$($result.Message)'"
            }

            # ----------------------------------------------------------------
            # Assert 3: OK must be $true (WhatIf path always succeeds)
            # ----------------------------------------------------------------
            if ($result.OK -ne $true) {
                $failures += "Iteration $i (Status='$($row.Status)', Type='$($row.Type)', MultiAction='$($row._MultiAction)') — OK was `$false in WhatIf mode (Message: '$($result.Message)')"
            }
        }

        # ----------------------------------------------------------------
        # Assert 4: None of the four DNS write cmdlets were ever called
        # ----------------------------------------------------------------
        Assert-MockCalled Add-DnsServerResourceRecordA    -Times 0 -Scope It
        Assert-MockCalled Add-DnsServerResourceRecordCName -Times 0 -Scope It
        Assert-MockCalled Set-DnsServerResourceRecord     -Times 0 -Scope It
        Assert-MockCalled Remove-DnsServerResourceRecord  -Times 0 -Scope It

        if ($failures.Count -gt 0) {
            $message = "Property 6 (WhatIf prevents DNS writes — ApplyRecord) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'covers all Status/Type combinations and never calls DNS write cmdlets (100 iterations)' {

        $failures = @()

        # Explicitly cycle through all status+type combinations to ensure full coverage
        $combinations = @(
            @{ Status = 'New';           Type = 'A';    MultiAction = $null          }
            @{ Status = 'New';           Type = 'CNAME'; MultiAction = $null         }
            @{ Status = 'Update';        Type = 'A';    MultiAction = $null          }
            @{ Status = 'Update';        Type = 'CNAME'; MultiAction = $null         }
            @{ Status = 'Convert';       Type = 'CNAME'; MultiAction = $null         }  # A→CNAME
            @{ Status = 'Convert';       Type = 'A';    MultiAction = $null          }  # CNAME→A
            @{ Status = 'MultiConflict'; Type = 'A';    MultiAction = 'ReplaceAll'   }
            @{ Status = 'MultiConflict'; Type = 'A';    MultiAction = 'AddAdditional'}
            @{ Status = 'MultiConflict'; Type = 'A';    MultiAction = 'NoAction'     }
            @{ Status = 'MultiConflict'; Type = 'A';    MultiAction = 'Unknown'      }  # default branch
        )

        for ($i = 0; $i -lt 100; $i++) {

            # Cycle through combinations, then use random rows for remaining iterations
            $combo = $combinations[$i % $combinations.Count]

            $name  = New-RandomLabel
            $zone  = New-RandomZoneName

            $newValue = if ($combo.Type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }
            $existingValue = if ($combo.Type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }

            $row = [PSCustomObject]@{
                Status        = $combo.Status
                StatusLabel   = $combo.Status
                Name          = $name
                Type          = $combo.Type
                ExistingValue = $existingValue
                NewValue      = $newValue
                TTL           = '01:00:00'
                Match         = 'Exact'
                ResultNote    = ''
                IsWildcard    = $false
                _Zone         = $zone
                _MultiAction  = $combo.MultiAction
            }

            $result = DNS-ApplyRecord -Row $row -DC 'dc01' -Credential $null -WhatIf $true

            if ($null -eq $result) {
                $failures += "Iteration $i (Status='$($combo.Status)', Type='$($combo.Type)', MultiAction='$($combo.MultiAction)') — DNS-ApplyRecord returned `$null"
                continue
            }

            if (-not ($result.Message -match '^\[WhatIf\]')) {
                $failures += "Iteration $i (Status='$($combo.Status)', Type='$($combo.Type)', MultiAction='$($combo.MultiAction)') — Message does not start with '[WhatIf]': '$($result.Message)'"
            }

            if ($result.OK -ne $true) {
                $failures += "Iteration $i (Status='$($combo.Status)', Type='$($combo.Type)', MultiAction='$($combo.MultiAction)') — OK was `$false in WhatIf mode"
            }
        }

        # Assert no DNS write cmdlets were called across all 100 iterations
        Assert-MockCalled Add-DnsServerResourceRecordA    -Times 0 -Scope It
        Assert-MockCalled Add-DnsServerResourceRecordCName -Times 0 -Scope It
        Assert-MockCalled Set-DnsServerResourceRecord     -Times 0 -Scope It
        Assert-MockCalled Remove-DnsServerResourceRecord  -Times 0 -Scope It

        if ($failures.Count -gt 0) {
            $message = "Property 6 (all combinations — ApplyRecord) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}

Describe 'DNS-DeleteRecord — Property 6: WhatIf mode prevents all DNS writes and prefixes all messages' -Tag 'Feature: dns-manager', 'Property 6: WhatIf mode prevents all DNS writes and prefixes all messages' {

    BeforeEach {
        Mock Add-DnsServerResourceRecordA    { }
        Mock Add-DnsServerResourceRecordCName { }
        Mock Set-DnsServerResourceRecord     { }
        Mock Remove-DnsServerResourceRecord  { }
    }

    It 'returns a message starting with "[WhatIf]" and never calls Remove-DnsServerResourceRecord (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            $row = New-RandomDeleteRow

            $result = DNS-DeleteRecord -Row $row -DC 'dc01' -Credential $null -WhatIf $true

            # ----------------------------------------------------------------
            # Assert 1: result must not be null
            # ----------------------------------------------------------------
            if ($null -eq $result) {
                $failures += "Iteration $i (Name='$($row.Name)', Type='$($row.Type)') — DNS-DeleteRecord returned `$null"
                continue
            }

            # ----------------------------------------------------------------
            # Assert 2: Message must start with "[WhatIf]"
            # Use -match with regex (escape brackets) instead of -like
            # because -like treats [ ] as wildcard metacharacters.
            # ----------------------------------------------------------------
            if (-not ($result.Message -match '^\[WhatIf\]')) {
                $failures += "Iteration $i (Name='$($row.Name)', Type='$($row.Type)') — Message does not start with '[WhatIf]': '$($result.Message)'"
            }

            # ----------------------------------------------------------------
            # Assert 3: OK must be $true
            # ----------------------------------------------------------------
            if ($result.OK -ne $true) {
                $failures += "Iteration $i (Name='$($row.Name)', Type='$($row.Type)') — OK was `$false in WhatIf mode (Message: '$($result.Message)')"
            }
        }

        # ----------------------------------------------------------------
        # Assert 4: None of the four DNS write cmdlets were ever called
        # ----------------------------------------------------------------
        Assert-MockCalled Add-DnsServerResourceRecordA    -Times 0 -Scope It
        Assert-MockCalled Add-DnsServerResourceRecordCName -Times 0 -Scope It
        Assert-MockCalled Set-DnsServerResourceRecord     -Times 0 -Scope It
        Assert-MockCalled Remove-DnsServerResourceRecord  -Times 0 -Scope It

        if ($failures.Count -gt 0) {
            $message = "Property 6 (WhatIf prevents DNS writes — DeleteRecord) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'covers both A and CNAME record types and never calls DNS write cmdlets (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # Alternate between A and CNAME to ensure both types are covered
            $type = if ($i % 2 -eq 0) { 'A' } else { 'CNAME' }
            $existingValue = if ($type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }

            $row = [PSCustomObject]@{
                Status        = 'Done'
                StatusLabel   = '✅ Done'
                Name          = New-RandomLabel
                Type          = $type
                ExistingValue = $existingValue
                NewValue      = ''
                TTL           = '01:00:00'
                Match         = 'Exact'
                ResultNote    = ''
                IsWildcard    = $false
                _Zone         = New-RandomZoneName
                _MultiAction  = $null
            }

            $result = DNS-DeleteRecord -Row $row -DC 'dc01' -Credential $null -WhatIf $true

            if ($null -eq $result) {
                $failures += "Iteration $i (Type='$type') — DNS-DeleteRecord returned `$null"
                continue
            }

            if (-not ($result.Message -match '^\[WhatIf\]')) {
                $failures += "Iteration $i (Type='$type', Name='$($row.Name)') — Message does not start with '[WhatIf]': '$($result.Message)'"
            }

            if ($result.OK -ne $true) {
                $failures += "Iteration $i (Type='$type', Name='$($row.Name)') — OK was `$false in WhatIf mode"
            }
        }

        # Assert no DNS write cmdlets were called across all 100 iterations
        Assert-MockCalled Add-DnsServerResourceRecordA    -Times 0 -Scope It
        Assert-MockCalled Add-DnsServerResourceRecordCName -Times 0 -Scope It
        Assert-MockCalled Set-DnsServerResourceRecord     -Times 0 -Scope It
        Assert-MockCalled Remove-DnsServerResourceRecord  -Times 0 -Scope It

        if ($failures.Count -gt 0) {
            $message = "Property 6 (all types — DeleteRecord) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}
