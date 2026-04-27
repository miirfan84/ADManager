# =============================================================================
# Core\AppState.ps1
# Single source of truth for all shared runtime state.
# Plugins READ from here; they never write directly — they call the helpers.
# =============================================================================

# ── Connection state ─────────────────────────────────────────────────────────
$Script:AppState = @{
    Connected  = $false
    DCName     = $null
    Credential = $null         # PSCredential
    WhatIf     = $true         # Default to safety
}

# ── Plugin registry — populated by PluginLoader ──────────────────────────────
# Each entry: @{ Meta=@{...}; TabItem=$wpfTabItem; Loaded=$true/$false; Error="" }
$Script:LoadedPlugins = [System.Collections.ArrayList]::new()

# ── App metadata ─────────────────────────────────────────────────────────────
$Script:AppName    = "AD Manager"
$Script:AppVersion = "2.0"

# ── AD properties fetched on every user search ───────────────────────────────
$Script:ADProperties = @(
    'SamAccountName','DisplayName','UserPrincipalName','EmailAddress',
    'Enabled','PasswordLastSet','PasswordExpired','PasswordNeverExpires',
    'pwdLastSet','LastLogonDate','Department','Title','DistinguishedName'
)

# =============================================================================
# State helpers — call these instead of writing $Script:AppState directly
# =============================================================================

function Set-AppConnected {
    param(
        [string]$DCName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $Script:AppState.DCName     = $DCName
    $Script:AppState.Credential = $Credential
    $Script:AppState.Connected  = $true
    Write-Log "AppState: connected to $DCName"

    # Notify all loaded plugins that connection changed
    Invoke-PluginConnectionEvent -Connected $true
}

function Set-AppDisconnected {
    $Script:AppState.DCName     = $null
    $Script:AppState.Credential = $null
    $Script:AppState.Connected  = $false
    Write-Log "AppState: disconnected"

    Invoke-PluginConnectionEvent -Connected $false
}

function Get-AppConnected  { return $Script:AppState.Connected  }
function Get-AppDCName     { return $Script:AppState.DCName     }
function Get-AppCredential { return $Script:AppState.Credential }

function Get-WhatIfMode    { return $Script:AppState.WhatIf    }
function Set-WhatIfMode    {
    param([bool]$Mode)
    $Script:AppState.WhatIf = $Mode
    Write-Log "AppState: WhatIf mode set to $($Mode)"
}

# =============================================================================
# Plugin event bus — plugins register an OnConnect/OnDisconnect scriptblock
# in their plugin.json or Handlers.ps1 via Register-PluginConnectionHandler
# =============================================================================

$Script:ConnectionHandlers = [System.Collections.ArrayList]::new()

function Register-ConnectionHandler {
    param([scriptblock]$Handler)
    [void]$Script:ConnectionHandlers.Add($Handler)
}

function Clear-ConnectionHandlers {
    $Script:ConnectionHandlers.Clear()
    Write-Log "AppState: Connection handlers cleared" -Source 'AppState'
}

function Invoke-PluginConnectionEvent {
    param([bool]$Connected)
    foreach ($handler in $Script:ConnectionHandlers) {
        try {
            & $handler $Connected
        } catch {
            Write-Log "ConnectionHandler error: $($_.Exception.Message)" -Level WARN -Source 'AppState'
        }
    }
}
