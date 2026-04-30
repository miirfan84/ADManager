# =============================================================================
# Plugins\DnsManager\Tests\DNS-FilterZones.Property.Tests.ps1
#
# Property-Based Tests for DNS-FilterZones (Pester)
#
# Property 7: Forward-only zone filtering
#   Validates: Requirements 2.2
#
# Run with:
#   Invoke-Pester -Path .\Plugins\DnsManager\Tests\DNS-FilterZones.Property.Tests.ps1
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

function New-ForwardZone {
    <#
    .SYNOPSIS
        Creates a PSCustomObject representing a forward lookup zone.
        ZoneName is a plain domain (e.g. "contoso.com"), IsReverseLookupZone=$false.
    #>
    $tlds  = @('com', 'net', 'org', 'local', 'internal', 'corp', 'io', 'test', 'example')
    $label = New-RandomLabel
    $tld   = $tlds[(Get-Random -Minimum 0 -Maximum $tlds.Count)]
    $name  = "$label.$tld"

    return [PSCustomObject]@{
        ZoneName             = $name
        IsReverseLookupZone  = $false
    }
}

function New-InAddrArpaZone {
    <#
    .SYNOPSIS
        Creates a PSCustomObject representing an IPv4 reverse lookup zone
        whose ZoneName ends with '.in-addr.arpa'.
    #>
    $octet1 = Get-Random -Minimum 1   -Maximum 256
    $octet2 = Get-Random -Minimum 0   -Maximum 256
    $octet3 = Get-Random -Minimum 0   -Maximum 256
    $name   = "$octet3.$octet2.$octet1.in-addr.arpa"

    return [PSCustomObject]@{
        ZoneName             = $name
        IsReverseLookupZone  = $true
    }
}

function New-Ip6ArpaZone {
    <#
    .SYNOPSIS
        Creates a PSCustomObject representing an IPv6 reverse lookup zone
        whose ZoneName ends with '.ip6.arpa'.
    #>
    # Generate a random partial IPv6 reverse zone name (nibble format)
    $nibbles = 1..(Get-Random -Minimum 4 -Maximum 9) | ForEach-Object {
        '{0:x}' -f (Get-Random -Minimum 0 -Maximum 16)
    }
    $name = ($nibbles -join '.') + '.ip6.arpa'

    return [PSCustomObject]@{
        ZoneName             = $name
        IsReverseLookupZone  = $true
    }
}

function New-IsReverseFlagZone {
    <#
    .SYNOPSIS
        Creates a PSCustomObject representing a zone that has IsReverseLookupZone=$true
        but whose ZoneName does NOT end with '.in-addr.arpa' or '.ip6.arpa'.
        This tests that the IsReverseLookupZone flag alone is sufficient to exclude a zone.
    #>
    $label = New-RandomLabel
    $tlds  = @('com', 'net', 'org', 'local', 'internal')
    $tld   = $tlds[(Get-Random -Minimum 0 -Maximum $tlds.Count)]
    $name  = "$label.$tld"

    return [PSCustomObject]@{
        ZoneName             = $name
        IsReverseLookupZone  = $true
    }
}

function New-RandomReverseZone {
    <#
    .SYNOPSIS
        Returns a random reverse zone using one of the three reverse zone types.
    #>
    $kind = Get-Random -Minimum 0 -Maximum 3
    switch ($kind) {
        0 { return New-InAddrArpaZone }
        1 { return New-Ip6ArpaZone }
        2 { return New-IsReverseFlagZone }
    }
}

# ---------------------------------------------------------------------------
# Property 7: Forward-only zone filtering
# Validates: Requirements 2.2
# ---------------------------------------------------------------------------

Describe 'DNS-FilterZones — Property 7: Forward-only zone filtering' -Tag 'Feature: dns-manager', 'Property 7: Forward-only zone filtering' {

    It 'filtered result contains only forward zones — no reverse zones appear (100 iterations)' {
        <#
        **Validates: Requirements 2.2**

        For any list of zone objects returned by Get-DnsServerZone, the filtered result
        SHALL contain only zones whose name does not end with '.in-addr.arpa' or '.ip6.arpa'
        and whose IsReverseLookupZone property is not $true.
        #>

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # ----------------------------------------------------------------
            # Build a random mixed zone list:
            #   - 0–8 forward zones
            #   - 0–8 reverse zones (mix of in-addr.arpa, ip6.arpa, IsReverseLookupZone=$true)
            # Ensure at least one zone total so the list is non-trivial.
            # ----------------------------------------------------------------

            $forwardCount = Get-Random -Minimum 0 -Maximum 9   # 0–8 forward zones
            $reverseCount = Get-Random -Minimum 0 -Maximum 9   # 0–8 reverse zones

            # Guarantee at least one zone in the list
            if ($forwardCount -eq 0 -and $reverseCount -eq 0) {
                $forwardCount = 1
            }

            $allZones      = @()
            $forwardZones  = @()
            $reverseZones  = @()

            for ($j = 0; $j -lt $forwardCount; $j++) {
                $z = New-ForwardZone
                $allZones     += $z
                $forwardZones += $z
            }

            for ($j = 0; $j -lt $reverseCount; $j++) {
                $z = New-RandomReverseZone
                $allZones    += $z
                $reverseZones += $z
            }

            # ----------------------------------------------------------------
            # Call DNS-FilterZones
            # ----------------------------------------------------------------

            $result = DNS-FilterZones -Zones $allZones

            # ----------------------------------------------------------------
            # Assert 1: No reverse zone appears in the result
            # ----------------------------------------------------------------

            foreach ($rev in $reverseZones) {
                $found = $result | Where-Object { $_.ZoneName -eq $rev.ZoneName -and $_.IsReverseLookupZone -eq $rev.IsReverseLookupZone }
                if ($found) {
                    $failures += "Iteration $i — reverse zone '$($rev.ZoneName)' (IsReverseLookupZone=$($rev.IsReverseLookupZone)) appeared in filtered result (forwardCount=$forwardCount, reverseCount=$reverseCount)"
                }
            }

            # ----------------------------------------------------------------
            # Assert 2: Every zone in the result has a forward-zone name
            #           (does not end with .in-addr.arpa or .ip6.arpa)
            #           and IsReverseLookupZone is not $true
            # ----------------------------------------------------------------

            foreach ($z in $result) {
                if ($z.ZoneName -match '\.in-addr\.arpa$') {
                    $failures += "Iteration $i — result contains zone '$($z.ZoneName)' which ends with '.in-addr.arpa'"
                }
                if ($z.ZoneName -match '\.ip6\.arpa$') {
                    $failures += "Iteration $i — result contains zone '$($z.ZoneName)' which ends with '.ip6.arpa'"
                }
                if ($z.IsReverseLookupZone -eq $true) {
                    $failures += "Iteration $i — result contains zone '$($z.ZoneName)' with IsReverseLookupZone=`$true"
                }
            }

            # ----------------------------------------------------------------
            # Assert 3: All forward zones are preserved in the result
            # ----------------------------------------------------------------

            foreach ($fwd in $forwardZones) {
                $found = $result | Where-Object { $_.ZoneName -eq $fwd.ZoneName }
                if (-not $found) {
                    $failures += "Iteration $i — forward zone '$($fwd.ZoneName)' was incorrectly excluded from the filtered result (forwardCount=$forwardCount, reverseCount=$reverseCount)"
                }
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 7 (forward-only zone filtering) failed for $($failures.Count) failure(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'returns an empty result when all zones are reverse zones (100 iterations)' {
        <#
        **Validates: Requirements 2.2**

        When the input contains only reverse zones, the filtered result must be empty.
        #>

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # Build a list of only reverse zones (1–8 zones)
            $reverseCount = Get-Random -Minimum 1 -Maximum 9
            $allZones     = @()

            for ($j = 0; $j -lt $reverseCount; $j++) {
                $allZones += New-RandomReverseZone
            }

            $result = DNS-FilterZones -Zones $allZones

            if ($result.Count -ne 0) {
                $zoneNames = ($result | ForEach-Object { $_.ZoneName }) -join ', '
                $failures += "Iteration $i — expected empty result for all-reverse input but got $($result.Count) zone(s): $zoneNames"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 7 (empty result for all-reverse input) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'returns all zones when all zones are forward zones (100 iterations)' {
        <#
        **Validates: Requirements 2.2**

        When the input contains only forward zones, all zones must be preserved in the result.
        #>

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # Build a list of only forward zones (1–8 zones)
            $forwardCount = Get-Random -Minimum 1 -Maximum 9
            $allZones     = @()

            for ($j = 0; $j -lt $forwardCount; $j++) {
                $allZones += New-ForwardZone
            }

            $result = DNS-FilterZones -Zones $allZones

            if ($result.Count -ne $forwardCount) {
                $failures += "Iteration $i — expected $forwardCount forward zone(s) in result but got $($result.Count)"
            }

            # Verify no reverse-zone indicators appear
            foreach ($z in $result) {
                if ($z.ZoneName -match '\.in-addr\.arpa$' -or $z.ZoneName -match '\.ip6\.arpa$' -or $z.IsReverseLookupZone -eq $true) {
                    $failures += "Iteration $i — result contains unexpected reverse zone '$($z.ZoneName)'"
                }
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 7 (all forward zones preserved) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}
