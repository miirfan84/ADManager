# =============================================================================
# Plugins\DnsManager\Tests\DNS-ClassifyRow.Property.Tests.ps1
#
# Property-Based Tests for DNS-ClassifyRow (Pester)
#
# Property 5: Record status classification is exhaustive and consistent
#   Validates: Requirements 5.1–5.5, 6.1
#
# Run with:
#   Invoke-Pester -Path .\Plugins\DnsManager\Tests\DNS-ClassifyRow.Property.Tests.ps1
# =============================================================================

# Dot-source the function under test
$FunctionsPath = Join-Path $PSScriptRoot '..' 'Functions.ps1'
. $FunctionsPath

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

function New-RandomRecordType {
    <#
    .SYNOPSIS
        Returns either "A" or "CNAME" randomly.
    #>
    $types = @('A', 'CNAME')
    return $types[(Get-Random -Minimum 0 -Maximum $types.Count)]
}

function New-RandomValue {
    <#
    .SYNOPSIS
        Returns a random DNS value — either IPv4 or CNAME target.
    #>
    if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) {
        return New-RandomIPv4
    } else {
        return New-RandomCNAMETarget
    }
}

function New-ParsedRow {
    <#
    .SYNOPSIS
        Creates a random ParsedRow hashtable.
    .PARAMETER Type
        Optional: force the Type to "A" or "CNAME".
    .PARAMETER Value
        Optional: force the Value.
    .PARAMETER IsWildcard
        Optional: force the IsWildcard flag.
    #>
    param(
        [string]$Type = $null,
        [string]$Value = $null,
        [bool]$IsWildcard = $false
    )

    $hostname = New-RandomLabel

    if (-not $Type) {
        $Type = New-RandomRecordType
    }

    if (-not $Value) {
        if ($Type -eq 'A') {
            $Value = New-RandomIPv4
        } else {
            $Value = New-RandomCNAMETarget
        }
    }

    return @{
        OK         = $true
        Hostname   = $hostname
        Type       = $Type
        Value      = $Value
        ErrorMsg   = ''
        IsWildcard = $IsWildcard
    }
}

function New-ExistingRecord {
    <#
    .SYNOPSIS
        Creates a random existing DNS record object.
    .PARAMETER Type
        Optional: force the Type to "A" or "CNAME".
    .PARAMETER ExistingValue
        Optional: force the ExistingValue.
    #>
    param(
        [string]$Type = $null,
        [string]$ExistingValue = $null
    )

    $hostname = New-RandomLabel

    if (-not $Type) {
        $Type = New-RandomRecordType
    }

    if (-not $ExistingValue) {
        if ($Type -eq 'A') {
            $ExistingValue = New-RandomIPv4
        } else {
            $ExistingValue = New-RandomCNAMETarget
        }
    }

    return [PSCustomObject]@{
        Hostname      = $hostname
        Type          = $Type
        ExistingValue = $ExistingValue
        TTL           = '01:00:00'
        IsWildcard    = $false
    }
}

# ---------------------------------------------------------------------------
# Property 5: Record status classification is exhaustive and consistent
# Validates: Requirements 5.1–5.5, 6.1
# ---------------------------------------------------------------------------

Describe 'DNS-ClassifyRow — Property 5: Record status classification is exhaustive and consistent' -Tag 'Feature: dns-manager', 'Property 5: Record status classification is exhaustive and consistent' {

    It 'returns exactly one correct status key for all six branches with no overlap and no gaps (100 iterations)' {

        $failures = @()

        # We will run 100 iterations, ensuring all six branches are covered across the iterations
        # Branch distribution: we'll cycle through all six branches multiple times
        $branches = @(
            'New',           # Empty ExistingRecs
            'WildcardMatch', # ParsedRow.IsWildcard=$true
            'MultiConflict', # ExistingRecs.Count > 1 and all A records
            'Convert',       # Single record, different type
            'NoAction',      # Single record, same type, same value
            'Update'         # Single record, same type, different value
        )

        for ($i = 0; $i -lt 100; $i++) {

            # Select which branch to test in this iteration (cycle through all branches)
            $targetBranch = $branches[$i % $branches.Count]

            $parsedRow    = $null
            $existingRecs = $null
            $expectedStatus = $targetBranch

            switch ($targetBranch) {

                'New' {
                    # Branch 1: Empty ExistingRecs → "New"
                    $parsedRow    = New-ParsedRow
                    $existingRecs = @()   # Empty array
                }

                'WildcardMatch' {
                    # Branch 2: ParsedRow.IsWildcard=$true → "WildcardMatch"
                    $parsedRow    = New-ParsedRow -IsWildcard $true
                    # ExistingRecs can be anything (even empty), but we'll add one record for realism
                    $existingRecs = @( (New-ExistingRecord) )
                }

                'MultiConflict' {
                    # Branch 3: ExistingRecs.Count > 1 and all A records → "MultiConflict"
                    $parsedRow    = New-ParsedRow -Type 'A'
                    $count        = Get-Random -Minimum 2 -Maximum 6   # 2–5 existing A records
                    $existingRecs = @()
                    for ($j = 0; $j -lt $count; $j++) {
                        $existingRecs += New-ExistingRecord -Type 'A'
                    }
                }

                'Convert' {
                    # Branch 4: Single record, different type → "Convert"
                    $parsedType   = New-RandomRecordType
                    $existingType = if ($parsedType -eq 'A') { 'CNAME' } else { 'A' }
                    $parsedRow    = New-ParsedRow -Type $parsedType
                    $existingRecs = @( (New-ExistingRecord -Type $existingType) )
                }

                'NoAction' {
                    # Branch 5: Single record, same type, same value → "NoAction"
                    $type         = New-RandomRecordType
                    $value        = if ($type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }
                    $parsedRow    = New-ParsedRow -Type $type -Value $value
                    $existingRecs = @( (New-ExistingRecord -Type $type -ExistingValue $value) )
                }

                'Update' {
                    # Branch 6: Single record, same type, different value → "Update"
                    $type         = New-RandomRecordType
                    $parsedValue  = if ($type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }
                    $existingValue = if ($type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }
                    # Ensure values are different
                    while ($parsedValue -eq $existingValue) {
                        $existingValue = if ($type -eq 'A') { New-RandomIPv4 } else { New-RandomCNAMETarget }
                    }
                    $parsedRow    = New-ParsedRow -Type $type -Value $parsedValue
                    $existingRecs = @( (New-ExistingRecord -Type $type -ExistingValue $existingValue) )
                }
            }

            # ----------------------------------------------------------------
            # Call DNS-ClassifyRow and verify the result
            # ----------------------------------------------------------------

            $actualStatus = DNS-ClassifyRow -ParsedRow $parsedRow -ExistingRecs $existingRecs

            # Assert: the returned status must match the expected status for this branch
            if ($actualStatus -ne $expectedStatus) {
                $failures += "Iteration $i (branch '$targetBranch') — expected status '$expectedStatus' but got '$actualStatus' (ParsedRow.Type='$($parsedRow.Type)', ParsedRow.Value='$($parsedRow.Value)', ParsedRow.IsWildcard='$($parsedRow.IsWildcard)', ExistingRecs.Count=$($existingRecs.Count))"
            }

            # ----------------------------------------------------------------
            # Additional validation: ensure the status is one of the six valid keys
            # ----------------------------------------------------------------

            $validStatuses = @('New', 'WildcardMatch', 'MultiConflict', 'Convert', 'NoAction', 'Update')
            if ($actualStatus -notin $validStatuses) {
                $failures += "Iteration $i (branch '$targetBranch') — returned status '$actualStatus' is not one of the six valid status keys"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 5 (classification is exhaustive and consistent) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'always returns exactly one status key (no null, no empty string) (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # Generate completely random inputs (any combination)
            $parsedRow    = New-ParsedRow -IsWildcard ((Get-Random -Minimum 0 -Maximum 2) -eq 0)
            $recCount     = Get-Random -Minimum 0 -Maximum 6   # 0–5 existing records
            $existingRecs = @()
            for ($j = 0; $j -lt $recCount; $j++) {
                $existingRecs += New-ExistingRecord
            }

            # Call DNS-ClassifyRow
            $status = DNS-ClassifyRow -ParsedRow $parsedRow -ExistingRecs $existingRecs

            # Assert: status must be a non-null, non-empty string
            if ([string]::IsNullOrEmpty($status)) {
                $failures += "Iteration $i — DNS-ClassifyRow returned null or empty status (ParsedRow.Type='$($parsedRow.Type)', ParsedRow.IsWildcard='$($parsedRow.IsWildcard)', ExistingRecs.Count=$($existingRecs.Count))"
            }

            # Assert: status must be one of the six valid keys
            $validStatuses = @('New', 'WildcardMatch', 'MultiConflict', 'Convert', 'NoAction', 'Update')
            if ($status -notin $validStatuses) {
                $failures += "Iteration $i — DNS-ClassifyRow returned invalid status '$status' (ParsedRow.Type='$($parsedRow.Type)', ParsedRow.IsWildcard='$($parsedRow.IsWildcard)', ExistingRecs.Count=$($existingRecs.Count))"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 5 (always returns exactly one status key) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}
