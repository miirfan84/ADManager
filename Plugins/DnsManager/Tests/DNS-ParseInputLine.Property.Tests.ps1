# =============================================================================
# Plugins\DnsManager\Tests\DNS-ParseInputLine.Property.Tests.ps1
#
# Property-Based Tests for DNS-ParseInputLine (Pester)
#
# Property 1: Zone suffix stripping is idempotent
#   Validates: Requirements 3.4
#
# Run with:
#   Invoke-Pester -Path .\Plugins\DnsManager\Tests\DNS-ParseInputLine.Property.Tests.ps1
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

function New-RandomValue {
    <#
    .SYNOPSIS
        Returns a random DNS value — either a valid IPv4 address or a CNAME target.
        This ensures the line has 2 tokens and parses as OK=$true.
    #>
    if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) {
        # IPv4
        $o1 = Get-Random -Minimum 1   -Maximum 256
        $o2 = Get-Random -Minimum 0   -Maximum 256
        $o3 = Get-Random -Minimum 0   -Maximum 256
        $o4 = Get-Random -Minimum 0   -Maximum 256
        return "$o1.$o2.$o3.$o4"
    } else {
        # CNAME target — a simple hostname (no spaces)
        return "$(New-RandomLabel).$(New-RandomZoneName)"
    }
}

# ---------------------------------------------------------------------------
# Property 1: Zone suffix stripping is idempotent
# Validates: Requirements 3.4
# ---------------------------------------------------------------------------

Describe 'DNS-ParseInputLine — Property 1: Zone suffix stripping is idempotent' -Tag 'Feature: dns-manager', 'Property 1: Zone suffix stripping is idempotent' {

    It 'produces the same Hostname whether the zone suffix is present or absent (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # Generate random inputs
            $hostname = New-RandomLabel          # e.g. "app"
            $zone     = New-RandomZoneName       # e.g. "contoso.com"
            $value    = New-RandomValue          # e.g. "192.168.1.1" or "target.contoso.com"

            # Line WITH zone suffix appended to the hostname
            $lineWithSuffix    = "$hostname.$zone $value"

            # Line WITHOUT zone suffix (short hostname only)
            $lineWithoutSuffix = "$hostname $value"

            # Parse both lines with the same zone
            $resultWith    = DNS-ParseInputLine -Line $lineWithSuffix    -ZoneName $zone
            $resultWithout = DNS-ParseInputLine -Line $lineWithoutSuffix -ZoneName $zone

            # Both must parse successfully
            if (-not $resultWith.OK) {
                $failures += "Iteration $i — lineWithSuffix '$lineWithSuffix' parsed as OK=`$false (ErrorMsg: '$($resultWith.ErrorMsg)')"
                continue
            }
            if (-not $resultWithout.OK) {
                $failures += "Iteration $i — lineWithoutSuffix '$lineWithoutSuffix' parsed as OK=`$false (ErrorMsg: '$($resultWithout.ErrorMsg)')"
                continue
            }

            # The Hostname output must be identical (the short label, no zone suffix)
            if ($resultWith.Hostname -ne $resultWithout.Hostname) {
                $failures += "Iteration $i — Hostname mismatch: with-suffix='$($resultWith.Hostname)' vs without-suffix='$($resultWithout.Hostname)' (hostname='$hostname', zone='$zone', value='$value')"
            }

            # Additionally, the Hostname must equal the original short label (not the FQDN)
            if ($resultWith.Hostname -ne $hostname) {
                $failures += "Iteration $i — Hostname '$($resultWith.Hostname)' does not equal expected short label '$hostname' (zone='$zone', value='$value')"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 1 (zone suffix stripping is idempotent) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}

# ---------------------------------------------------------------------------
# Property 2: IPv4 detection determines record type
# Validates: Requirements 3.2, 3.3
# ---------------------------------------------------------------------------

Describe 'DNS-ParseInputLine — Property 2: IPv4 detection determines record type' -Tag 'Feature: dns-manager', 'Property 2: IPv4 detection determines record type' {

    It 'returns Type="A" for every valid IPv4 second token (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            # Generate a random valid IPv4 address (each octet 0–255)
            $o1  = Get-Random -Minimum 0 -Maximum 256
            $o2  = Get-Random -Minimum 0 -Maximum 256
            $o3  = Get-Random -Minimum 0 -Maximum 256
            $o4  = Get-Random -Minimum 0 -Maximum 256
            $ipv4 = "$o1.$o2.$o3.$o4"

            $hostname = New-RandomLabel
            $zone     = New-RandomZoneName
            $line     = "$hostname $ipv4"

            $result = DNS-ParseInputLine -Line $line -ZoneName $zone

            if (-not $result.OK) {
                $failures += "Iteration $i — line '$line' parsed as OK=`$false (ErrorMsg: '$($result.ErrorMsg)')"
                continue
            }

            if ($result.Type -ne 'A') {
                $failures += "Iteration $i — expected Type='A' for IPv4 '$ipv4' but got Type='$($result.Type)' (line='$line')"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 2 (IPv4 → Type='A') failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }

    It 'returns Type="CNAME" for every non-IP second token (100 iterations)' {

        $failures = @()

        # A set of non-IP string patterns to draw from
        $nonIpPatterns = @(
            # Plain hostname labels
            { New-RandomLabel },
            # FQDN-style targets
            { "$(New-RandomLabel).$(New-RandomZoneName)" },
            # Strings that look almost like IPs but are not (extra octet, letters, etc.)
            { "$(Get-Random -Min 0 -Max 256).$(Get-Random -Min 0 -Max 256).$(Get-Random -Min 0 -Max 256)" },   # only 3 octets
            { "$(New-RandomLabel).$(Get-Random -Min 0 -Max 256).$(Get-Random -Min 0 -Max 256).$(Get-Random -Min 0 -Max 256)" },  # leading label
            { "$(Get-Random -Min 0 -Max 256).$(Get-Random -Min 0 -Max 256).$(Get-Random -Min 0 -Max 256).$(New-RandomLabel)" }   # trailing label
        )

        for ($i = 0; $i -lt 100; $i++) {

            # Pick a random non-IP generator
            $generator = $nonIpPatterns[(Get-Random -Minimum 0 -Maximum $nonIpPatterns.Count)]
            $nonIp     = & $generator

            # Ensure the generated value is not accidentally a valid IPv4 pattern
            # (i.e. does not match ^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$)
            if ($nonIp -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                # Skip this iteration — the generator accidentally produced a valid IP
                continue
            }

            $hostname = New-RandomLabel
            $zone     = New-RandomZoneName
            $line     = "$hostname $nonIp"

            $result = DNS-ParseInputLine -Line $line -ZoneName $zone

            if (-not $result.OK) {
                $failures += "Iteration $i — line '$line' parsed as OK=`$false (ErrorMsg: '$($result.ErrorMsg)')"
                continue
            }

            if ($result.Type -ne 'CNAME') {
                $failures += "Iteration $i — expected Type='CNAME' for non-IP '$nonIp' but got Type='$($result.Type)' (line='$line')"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 2 (non-IP → Type='CNAME') failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}

# ---------------------------------------------------------------------------
# Property 3: Parse error on insufficient tokens
# Validates: Requirements 3.6
# ---------------------------------------------------------------------------

Describe 'DNS-ParseInputLine — Property 3: Parse error on insufficient tokens' -Tag 'Feature: dns-manager', 'Property 3: Parse error on insufficient tokens' {

    It 'returns OK=$false with a non-empty ErrorMsg for every single-token non-blank non-comment line (100 iterations)' {

        $failures = @()

        # Generator: produce a single token that is non-blank and does not start with '#'
        function New-SingleToken {
            do {
                # Build a token from printable non-whitespace characters, length 1–20
                $chars  = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.'
                $length = Get-Random -Minimum 1 -Maximum 21
                $token  = -join (1..$length | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
            } while ($token -eq '' -or $token.StartsWith('#'))
            return $token
        }

        for ($i = 0; $i -lt 100; $i++) {

            $token = New-SingleToken
            $zone  = New-RandomZoneName

            # The line is exactly one token — no whitespace, so fewer than 2 tokens
            $result = DNS-ParseInputLine -Line $token -ZoneName $zone

            # Must return OK=$false
            if ($result.OK -ne $false) {
                $failures += "Iteration $i — single-token line '$token' returned OK=`$true (expected OK=`$false)"
                continue
            }

            # ErrorMsg must be non-empty
            if ([string]::IsNullOrEmpty($result.ErrorMsg)) {
                $failures += "Iteration $i — single-token line '$token' returned OK=`$false but ErrorMsg was empty"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 3 (parse error on insufficient tokens) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}

# ---------------------------------------------------------------------------
# Property 4: Blank and comment lines produce no output
# Validates: Requirements 3.5
# ---------------------------------------------------------------------------

Describe 'DNS-ParseInputLine — Property 4: Blank and comment lines produce no output' -Tag 'Feature: dns-manager', 'Property 4: Blank and comment lines produce no output' {

    It 'output count equals count of valid two-token lines only (100 iterations)' {

        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {

            $zone = New-RandomZoneName

            # ----------------------------------------------------------------
            # Build a random mixed collection of lines:
            #   - blank lines (empty string or whitespace-only)
            #   - comment lines (start with '#')
            #   - valid two-token lines (hostname + value)
            # ----------------------------------------------------------------

            $lines          = @()
            $expectedCount  = 0
            $totalLineCount = Get-Random -Minimum 5 -Maximum 31   # 5–30 lines per collection

            for ($j = 0; $j -lt $totalLineCount; $j++) {

                $kind = Get-Random -Minimum 0 -Maximum 3   # 0=blank, 1=comment, 2=valid

                switch ($kind) {
                    0 {
                        # Blank line — either empty or whitespace-only
                        $spaces = ' ' * (Get-Random -Minimum 0 -Maximum 5)
                        $lines += $spaces
                    }
                    1 {
                        # Comment line — starts with '#', may have trailing text
                        $comment = '#' + (New-RandomLabel)
                        $lines += $comment
                    }
                    2 {
                        # Valid two-token line — hostname + value (A or CNAME)
                        $hostname = New-RandomLabel
                        $value    = New-RandomValue
                        $lines += "$hostname $value"
                        $expectedCount++
                    }
                }
            }

            # ----------------------------------------------------------------
            # Parse every line and collect results where OK=$true
            # ----------------------------------------------------------------

            $parsedOkCount = 0
            foreach ($line in $lines) {
                $result = DNS-ParseInputLine -Line $line -ZoneName $zone
                if ($result.OK) {
                    $parsedOkCount++
                }
            }

            # ----------------------------------------------------------------
            # Assert: only the valid two-token lines produced OK=$true output
            # ----------------------------------------------------------------

            if ($parsedOkCount -ne $expectedCount) {
                $failures += "Iteration $i — expected $expectedCount OK results but got $parsedOkCount (zone='$zone', totalLines=$totalLineCount)"
            }
        }

        if ($failures.Count -gt 0) {
            $message = "Property 4 (blank and comment lines produce no output) failed for $($failures.Count) iteration(s).`nFailures:`n" + ($failures -join "`n")
            throw $message
        }
        $failures.Count | Should Be 0
    }
}
