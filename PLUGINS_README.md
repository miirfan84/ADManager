# AD Manager v2.0 — Plugin Developer Guide
# ============================================

## Folder Structure
```
ADManager\
├── Launch.ps1              <- Entry point (don't modify)
├── Core\
│   ├── Logger.ps1          <- Write-Log function
│   ├── AppState.ps1        <- Shared state + connection helpers
│   ├── PluginLoader.ps1    <- Discovers and sandboxes plugins
│   └── UIShell.ps1         <- Main window + connection bar
├── UI\
│   └── Shell.xaml          <- Main window XAML
└── Plugins\
    └── <YourPlugin>\       <- Drop a folder here to add a tab
        ├── plugin.json     <- Required: metadata
        ├── Tab.xaml        <- Required: tab UI (root must be a Panel/Grid)
        ├── Handlers.ps1    <- Required: event wiring
        └── Functions.ps1   <- Optional: pure logic (no UI refs)
```

---

## Creating a New Plugin

### 1. plugin.json
```json
{
  "Id":          "GroupManagement",
  "TabTitle":    "👥 Group Management",
  "Description": "Manage AD group memberships",
  "Version":     "1.0",
  "Author":      "Your Name",
  "Order":       20
}
```
- `Id`       — unique, no spaces, used internally
- `TabTitle` — shown on the tab header (emojis OK)
- `Order`    — lower number = further left tab

### 2. Tab.xaml
Root element must be a **Grid** (or other Panel). PluginLoader wraps it in a TabItem.
Name all controls with a plugin prefix to avoid clashes:
```xml
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      Background="#1A252F">
    <Button x:Name="GM_BtnSearch" Content="Search"/>
    <DataGrid x:Name="GM_Grid" .../>
</Grid>
```

### 3. Handlers.ps1
PluginLoader sets these before dot-sourcing Handlers.ps1:
- `$Script:PluginRoot` — the loaded root element from Tab.xaml
- `$Script:PluginMeta` — parsed plugin.json as PSCustomObject

```powershell
# Grab your controls
$GM = @{}
foreach ($n in @('GM_BtnSearch','GM_Grid')) {
    $GM[$n] = $Script:PluginRoot.FindName($n)
}

# React to DC connect/disconnect via the event bus
Register-ConnectionHandler -Handler {
    param([bool]$Connected)
    $GM['GM_BtnSearch'].IsEnabled = $Connected
}

# Wire buttons
$GM['GM_BtnSearch'].Add_Click({
    $dc   = Get-AppDCName       # from AppState
    $cred = Get-AppCredential   # from AppState
    # ... your AD logic
})
```

### 4. Functions.ps1 (optional)
Pure PowerShell — no WPF imports. Keep AD cmdlets here, call from Handlers.ps1:
```powershell
function GM-GetGroupMembers {
    param([string]$GroupName, [string]$DC, $Cred)
    Get-ADGroupMember -Identity $GroupName -Server $DC -Credential $Cred
}
```

---

## Core APIs Available to All Plugins

### AppState helpers
```powershell
Get-AppConnected   # → $true/$false
Get-AppDCName      # → "dc01.corp.local"
Get-AppCredential  # → PSCredential
```

### Shell helpers (update the main window)
```powershell
Set-ShellStatus   -Text "Searching..."
Set-ShellProgress -Value 50   # 0-100, or -1 to hide
```

### Connection event bus
```powershell
Register-ConnectionHandler -Handler {
    param([bool]$Connected)
    # runs when user connects or disconnects
}
```

### Logging
```powershell
Write-Log "Something happened"             -Source 'MyPlugin'
Write-Log "Something bad" -Level ERROR     -Source 'MyPlugin'
Write-Log "Watch this"    -Level WARN      -Source 'MyPlugin'
# Log file: %TEMP%\ADManager\ADManager_YYYYMMDD.log
```

---

## Plugin Isolation Guarantee
If your plugin throws any error during load (bad XAML, missing file, syntax error),
PluginLoader catches it, shows a red "⚠ Plugin failed" tab, and continues loading
all other plugins. The core app and other plugins are unaffected.

---

## Log File Location
`%TEMP%\ADManager\ADManager_YYYYMMDD.log`
