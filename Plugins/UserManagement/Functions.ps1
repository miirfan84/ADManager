# =============================================================================
# Plugins\UserManagement\Functions.ps1
# Pure AD logic — no WPF references.
# Called from Handlers.ps1 and from background runspaces.
# =============================================================================

function UM-GetOUFromDN {
    param([string]$DN)
    if ([string]::IsNullOrWhiteSpace($DN)) { return "" }
    $parts   = $DN -split ","
    $ouParts = $parts | Where-Object { $_ -match "^OU=" }
    if ($ouParts) { return ($ouParts -join " > ").Replace("OU=","") }
    return ""
}

function UM-GetMustChange {
    param($User)
    try {
        if ($User.PasswordNeverExpires) { return $false }
        $raw = $User.pwdLastSet
        if ($null -eq $raw) { return $false }
        return ($raw -eq 0)
    } catch { return $false }
}

function UM-BuildUserRow {
    <#
    .SYNOPSIS
        Converts a raw ADUser object into a flat ordered hashtable for the DataGrid.
        Safe to call from any runspace.
    #>
    param([object]$ADUser, [string]$InputAccount)

    if ($null -eq $ADUser) {
        return [ordered]@{
            OpStatus    = ""
            SAM         = $InputAccount
            DisplayName = "NOT FOUND"
            UPN         = ""; Email = ""; Enabled = ""; PwNeverExp = ""; PwdLastSet = ""
            MustChange  = ""; LastLogon = ""; Department = ""; Title = ""; OU = ""
            NotFound    = $true
        }
    }

    $mustChange = UM-GetMustChange -User $ADUser
    $enabled    = if ($ADUser.Enabled -eq $true) { "✅" } else { "❌" }
    $neverExp   = if ($ADUser.PasswordNeverExpires -eq $true) { "✅" } else { "❌" }
    $pwdSet     = if ($null -ne $ADUser.PasswordLastSet) { $ADUser.PasswordLastSet.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
    $logon      = if ($null -ne $ADUser.LastLogonDate)   { $ADUser.LastLogonDate.ToString("yyyy-MM-dd HH:mm") }   else { "Never" }
    $mustStr    = if ($mustChange) { "✅" } else { "❌" }
    $dept       = if ($null -ne $ADUser.Department) { $ADUser.Department } else { "" }
    $title      = if ($null -ne $ADUser.Title)      { $ADUser.Title }      else { "" }

    return [ordered]@{
        OpStatus    = ""
        SAM         = $ADUser.SamAccountName
        DisplayName = $ADUser.DisplayName
        UPN         = $ADUser.UserPrincipalName
        Email       = $ADUser.EmailAddress
        Enabled     = $enabled
        PwNeverExp  = $neverExp
        PwdLastSet  = $pwdSet
        MustChange  = $mustStr
        LastLogon   = $logon
        Department  = $dept
        Title       = $title
        OU          = UM-GetOUFromDN -DN $ADUser.DistinguishedName
        NotFound    = $false
    }
}

function UM-ResetPasswords {
    <#
    .SYNOPSIS
        Resets password for one or more SAM accounts.
        Returns ArrayList of @{ SAM; Success; Message }.
    #>
    param(
        [string[]]$SAMNames,
        [System.Security.SecureString]$NewPassword,
        [bool]$MustChange,
        [string]$DCName,
        [System.Management.Automation.PSCredential]$Credential
    )

    $results = New-Object System.Collections.ArrayList

    foreach ($sam in $SAMNames) {
        $r = @{ SAM = $sam; Success = $false; Message = "" }
        try {
            Set-ADAccountPassword -Identity $sam -NewPassword $NewPassword -Reset `
                                  -Server $DCName -Credential $Credential -ErrorAction Stop
            Set-ADUser -Identity $sam -ChangePasswordAtLogon $MustChange `
                       -Server $DCName -Credential $Credential -ErrorAction Stop
            $r.Success = $true
            $r.Message = "Password reset OK"
            Write-Log "Password reset: $sam (MustChange=$MustChange)" -Source 'UserManagement'
        } catch {
            $r.Message = $_.Exception.Message
            Write-Log "Password reset FAILED: $sam — $($_.Exception.Message)" -Level ERROR -Source 'UserManagement'
        }
        [void]$results.Add($r)
    }
    return $results
}

function UM-SetForceChange {
    <#
    .SYNOPSIS
        Sets ChangePasswordAtLogon=$true for one or more accounts without touching the password.
        Returns ArrayList of @{ SAM; Success; Message }.
    #>
    param(
        [string[]]$SAMNames,
        [string]$DCName,
        [System.Management.Automation.PSCredential]$Credential
    )

    $results = New-Object System.Collections.ArrayList

    foreach ($sam in $SAMNames) {
        $r = @{ SAM = $sam; Success = $false; Message = "" }
        try {
            Set-ADUser -Identity $sam -ChangePasswordAtLogon $true `
                       -Server $DCName -Credential $Credential -ErrorAction Stop
            $r.Success = $true
            $r.Message = "Flag set OK"
            Write-Log "ForceChange set: $sam" -Source 'UserManagement'
        } catch {
            $r.Message = $_.Exception.Message
            Write-Log "ForceChange FAILED: $sam — $($_.Exception.Message)" -Level ERROR -Source 'UserManagement'
        }
        [void]$results.Add($r)
    }
    return $results
}

function UM-ClearNeverExp {
    <#
    .SYNOPSIS
        Sets PasswordNeverExpires=$false for one or more accounts.
        Returns ArrayList of @{ SAM; Success; Message }.
    #>
    param(
        [string[]]$SAMNames,
        [string]$DCName,
        [System.Management.Automation.PSCredential]$Credential
    )

    $results = New-Object System.Collections.ArrayList

    foreach ($sam in $SAMNames) {
        $r = @{ SAM = $sam; Success = $false; Message = "" }
        try {
            Set-ADUser -Identity $sam -PasswordNeverExpires $false `
                       -Server $DCName -Credential $Credential -ErrorAction Stop
            $r.Success = $true
            $r.Message = "Cleared OK"
            Write-Log "Password Never Expires cleared: $sam" -Source 'UserManagement'
        } catch {
            $r.Message = $_.Exception.Message
            Write-Log "Clear Password Never Expires FAILED: $sam — $($_.Exception.Message)" -Level ERROR -Source 'UserManagement'
        }
        [void]$results.Add($r)
    }
    return $results
}

function UM-RefreshUserRow {
    <#
    .SYNOPSIS
        Re-queries AD for a single user and returns a fresh row hashtable.
        Returns $null if the user cannot be retrieved.
    #>
    param(
        [string]$SAM,
        [string]$DCName,
        [System.Management.Automation.PSCredential]$Credential
    )
    try {
        $adUser = Get-ADUser -Identity $SAM -Server $DCName -Credential $Credential `
                             -Properties $Script:ADProperties -ErrorAction Stop
        return UM-BuildUserRow -ADUser $adUser -InputAccount $SAM
    } catch {
        Write-Log "Refresh failed for ${SAM}: $($_.Exception.Message)" -Level WARN -Source 'UserManagement'
        return $null
    }
}
