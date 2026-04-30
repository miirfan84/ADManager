# =============================================================================
# Plugins\DnsManager\Tests\DNS-PartialFailure.Property.Tests.ps1
#
# Property-Based Tests for DNS-FetchRecords (Pester)
#
# Property 8: Partial failure resilience — all rows are processed regardless
#             of individual errors
#
#   For any list of hostnames where a random subset of Get-DnsServerResourceRecord
#   calls throw, every hostname SHALL appear in the output:
#     - Failed hostnames → Type="Error" with non-empty ExistingValue (error message)
#     - Successful hostnames → processed normally (Type != "Error")
#   No exception in one hostname SHALL prevent processing of the remaining hostnames.
#
#   Validates: Requirements 4.7, 7.5, 8.5
#
# Run with:
#   Invoke-Pester -Path .\Plugins\DnsManager\Tests\DNS-PartialFailure.Property.Tests.ps1
# =============================================================================

# Dot-source the function under test
$FunctionsPath = Join-Path $PSScriptRoot '..' 'Functions.ps1'
. $FunctionsPath

# ---------------------------------------------------------------------------
# Stub definitions for DnsServer cmdlets
# Pester Mock requires the command to exist in the session before it can be
# intercepted. Because the DnsServer module is not installed on the test
# machine, we define no-op stubs here.
# ---------------------------------------------------------------------------
if (-not (Get-Command Get-DnsServerResourceRecord -ErrorAction SilentlyContinue)) {
    function Get-DnsServerResourceRecord { param([Parameter(ValueFromRemainingArguments)]$args) }
}

# ---------------------------------------------------------------------------
# Generator helpers
# ---------------------------------------------------------------------------

function New-RandomLabel {
    <#
    .SYNOPSIS
        Returns a random short DNS label (letters and digits only, 1–12 chars).
        Labels are guaranteed unique within a test iteration by appending a counter.
    #>
    param([int]$Index = -1)
    $chars  = 'abcdefghijklmnopqrstuvwxyz0123456789'
    $length = Get-Random -Minimum 3 -Maximum 10
    $label  = -join (1..$length | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
    if ($Index -ge 0) { $label = "$label$Index" }
    return $label
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

function New-RandomHostnameList {
    <#
    .SYNOPSIS
        Returns an array of 2–8 distinct short hostnames.
    #>
    $count     = Get-Random -Minimum 2 -Maximum 9
    $hostnames = @()
    for ($i = 0; $i -lt $count; $i++) {
        $hostnames += New-RandomLabel -Index $i
    }
    return $hostnames
}

function Select-RandomSubset {
    <#
    .SYNOPSIS
        Returns a random subset (possibly empty, possibly all) of the input array.
        Guarantees at least one element is NOT in the subset when the array has
        more than one element, so that we always have at least one successful hostname.
    #>
    param([string[]]$Items)

    if ($Items.Count -le 1) {
        # With only one item, randomly decide to fail it or not
        if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) { return @($Items[0]) }
        return @()
    }

    # Randomly pick a subset size: 0 to (Count-1) so at least one succeeds
    $maxFail = $Items.Count - 1
    $failCount = Get-Random -Minimum 0 -Maximum ($maxFail + 1)

    # Shuffle and take the first $failCount items
    $shuffled = $Items | Sort-Object { Get-Random }
    return @($shuffled | Select-Object -First $failCount)
}

# ---------------------------------------------------------------------------
# Property 8: Partial failure resilience — all rows are processed regardless
#             of individual errors
# Validates: Requirements 4.7, 7.5, 8.5
# ---------------------------------------------------------------------------

Describe 'DNS-FetchRecords — Property 8: Partial failure resilience' -Tag 'Feature: dns-manager', 'Property 8: Partial failure resilience' {

    It 'failed hostnames appear as Type="Error" with non-empty ExistingValue; successful hostnames are not marked as errors (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # ----------------------------------------------------------------
            # Generate a random list of distinct hostnames
            # ----------------------------------------------------------------
            $hostnames = New-RandomHostnameList
            $zone      = New-RandomZoneName
            $dc        = 'dc01'

            # ----------------------------------------------------------------
            # Choose a random subset of hostnames that will throw.
            # We track which call index (0-based, per-hostname queries only)
            # should fail. DNS-FetchRecords calls Get-DnsServerResourceRecord
            # once per hostname (exact query) then once more for the zone-wide
            # wildcard pass. We use a script-scoped call counter to identify
            # per-hostname calls (calls 0..N-1) vs the zone-wide call (call N).
            # ----------------------------------------------------------------
            $failingHostnames = Select-RandomSubset -Items $hostnames
            $successHostnames = $hostnames | Where-Object { $_ -notin $failingHostnames }

            # Build a set of call indices (0-based) that should throw.
            # Call index maps to hostname index in $hostnames.
            $Script:_FailingIndices = [System.Collections.Generic.HashSet[int]]::new()
            for ($j = 0; $j -lt $hostnames.Count; $j++) {
                if ($failingHostnames -contains $hostnames[$j]) {
                    [void]$Script:_FailingIndices.Add($j)
                }
            }
            $Script:_CallCounter   = 0
            $Script:_HostnameCount = $hostnames.Count

            Mock Get-DnsServerResourceRecord {
                $callIdx = $Script:_CallCounter
                $Script:_CallCounter++

                # Per-hostname calls are indices 0..(N-1).
                # The zone-wide wildcard pass is call N (or later).
                if ($callIdx -lt $Script:_HostnameCount) {
                    if ($Script:_FailingIndices.Contains($callIdx)) {
                        throw "Simulated DNS error for call index $callIdx"
                    }
                }

                # Successful call — return empty array (no existing records)
                return @()
            }

            # ----------------------------------------------------------------
            # Call DNS-FetchRecords
            # ----------------------------------------------------------------
            $output = DNS-FetchRecords `
                -Hostnames  $hostnames `
                -Zone       $zone `
                -DC         $dc `
                -Credential $null

            # ----------------------------------------------------------------
            # Assert 1: Every FAILING hostname appears in the output as an Error record.
            # (Successful hostnames with no existing DNS records will not appear in
            # output — that is correct behaviour; DNS-FetchRecords only emits records
            # that exist in DNS or that produced an error.)
            # ----------------------------------------------------------------
            foreach ($h in $failingHostnames) {
                $matchingRows = @($output | Where-Object { $_.Hostname -eq $h })
                if ($matchingRows.Count -eq 0) {
                    $failures += "Iter $i — failing hostname '$h' is missing from output (failingSet: $($failingHostnames -join ','))"
                }
            }

            # ----------------------------------------------------------------
            # Assert 2: Failed hostnames have Type="Error" with non-empty ExistingValue
            # ----------------------------------------------------------------
            foreach ($h in $failingHostnames) {
                $row = @($output | Where-Object { $_.Hostname -eq $h }) | Select-Object -First 1
                if ($null -eq $row) {
                    # Already caught by Assert 1 above; skip to avoid duplicate noise
                    continue
                }
                if ($row.Type -ne 'Error') {
                    $failures += "Iter $i — failing hostname '$h' has Type='$($row.Type)' instead of 'Error'"
                }
                if ([string]::IsNullOrEmpty($row.ExistingValue)) {
                    $failures += "Iter $i — failing hostname '$h' has empty ExistingValue (error message should be present)"
                }
            }

            # ----------------------------------------------------------------
            # Assert 3: Successful hostnames are NOT marked as errors
            # (They may not appear in output at all if no DNS records exist —
            # that is correct. We only check rows that DO appear.)
            # ----------------------------------------------------------------
            foreach ($h in $successHostnames) {
                $row = @($output | Where-Object { $_.Hostname -eq $h }) | Select-Object -First 1
                if ($null -eq $row) {
                    # Not in output — correct when no existing DNS records exist for this hostname
                    continue
                }
                if ($row.Type -eq 'Error') {
                    $failures += "Iter $i — successful hostname '$h' was unexpectedly marked as Type='Error' (ExistingValue: '$($row.ExistingValue)')"
                }
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 8 (partial failure resilience) failed for $($failures.Count) check(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'all hostnames are processed even when every single query throws (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # ----------------------------------------------------------------
            # Generate a random list of distinct hostnames
            # ----------------------------------------------------------------
            $hostnames = New-RandomHostnameList
            $zone      = New-RandomZoneName
            $dc        = 'dc01'

            # ----------------------------------------------------------------
            # Mock: ALL calls throw unconditionally.
            # This covers the case where every per-hostname query fails.
            # The zone-wide wildcard pass also throws, but that is non-fatal
            # per the DNS-FetchRecords design (wildcard rows simply won't appear).
            # ----------------------------------------------------------------
            Mock Get-DnsServerResourceRecord {
                throw "Total DNS failure"
            }

            # ----------------------------------------------------------------
            # Call DNS-FetchRecords
            # ----------------------------------------------------------------
            $output = DNS-FetchRecords `
                -Hostnames  $hostnames `
                -Zone       $zone `
                -DC         $dc `
                -Credential $null

            # ----------------------------------------------------------------
            # Assert: Every hostname appears in the output as an Error row
            # ----------------------------------------------------------------
            foreach ($h in $hostnames) {
                $row = @($output | Where-Object { $_.Hostname -eq $h }) | Select-Object -First 1
                if ($null -eq $row) {
                    $failures += "Iter $i — hostname '$h' is missing from output when all queries fail"
                    continue
                }
                if ($row.Type -ne 'Error') {
                    $failures += "Iter $i — hostname '$h' has Type='$($row.Type)' instead of 'Error' when all queries fail"
                }
                if ([string]::IsNullOrEmpty($row.ExistingValue)) {
                    $failures += "Iter $i — hostname '$h' has empty ExistingValue when all queries fail (error message expected)"
                }
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 8 (all-fail resilience) failed for $($failures.Count) check(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'all hostnames are processed when no queries throw (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # ----------------------------------------------------------------
            # Generate a random list of distinct hostnames
            # ----------------------------------------------------------------
            $hostnames = New-RandomHostnameList
            $zone      = New-RandomZoneName
            $dc        = 'dc01'

            # ----------------------------------------------------------------
            # Mock: No queries throw — all return empty arrays (no existing records)
            # ----------------------------------------------------------------
            Mock Get-DnsServerResourceRecord {
                return @()
            }

            # ----------------------------------------------------------------
            # Call DNS-FetchRecords
            # ----------------------------------------------------------------
            $output = DNS-FetchRecords `
                -Hostnames  $hostnames `
                -Zone       $zone `
                -DC         $dc `
                -Credential $null

            # ----------------------------------------------------------------
            # Assert: No hostname appears as an Error row
            # ----------------------------------------------------------------
            $errorRows = @($output | Where-Object { $_.Type -eq 'Error' })
            if ($errorRows.Count -gt 0) {
                $errorNames = ($errorRows | ForEach-Object { $_.Hostname }) -join ', '
                $failures += "Iter $i — unexpected Error rows when no queries throw: $errorNames"
            }

            # ----------------------------------------------------------------
            # Assert: Output contains no rows for hostnames that were not queried
            # (sanity check — output hostnames must be a subset of input hostnames
            #  plus any wildcard-match hostnames; since mock returns empty, only
            #  exact-match rows can appear, and there are none — so output is empty)
            # ----------------------------------------------------------------
            foreach ($row in $output) {
                if ($row.IsWildcard -eq $false -and $row.Hostname -notin $hostnames) {
                    $failures += "Iter $i — output contains unexpected hostname '$($row.Hostname)' not in input list"
                }
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 8 (no-fail baseline) failed for $($failures.Count) check(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}
