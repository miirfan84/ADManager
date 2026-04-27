# =============================================================================
# Core\PluginLoader.ps1
# Discovers plugins under Plugins\ and loads each one in a try/catch sandbox.
# A broken plugin gets a red error tab — the rest of the app keeps running.
#
# Plugin contract (each folder under Plugins\ must contain):
#   plugin.json   — metadata  (required)
#   Tab.xaml      — tab UI    (required)
#   Handlers.ps1  — event wiring (required)
#   Functions.ps1 — helper functions (optional but conventional)
#
# plugin.json schema:
# {
#   "Id":          "UserManagement",      // unique, no spaces
#   "TabTitle":    "👤 User Management",
#   "Description": "Search and manage AD users",
#   "Version":     "1.0",
#   "Author":      "AD Manager",
#   "Order":       10                     // tab sort order (lower = left)
# }
# =============================================================================

function Reset-Plugins {
    param([System.Windows.Controls.TabControl]$TabControl)
    
    Write-Log "Resetting plugin system..." -Source 'PluginLoader'
    
    # 1. Clear TabControl items
    $TabControl.Items.Clear()
    
    # 2. Clear connection event registry to prevent handler leaks
    Clear-ConnectionHandlers

    # 3. Reset the internal loaded tracking list
    if ($Script:LoadedPlugins) {
        $Script:LoadedPlugins.Clear()
    } else {
        $Script:LoadedPlugins = [System.Collections.ArrayList]::new()
    }
}

function Import-Plugins {
    param(
        [System.Windows.Controls.TabControl]$TabControl,
        [string]$PluginsRoot
    )

    if (-not (Test-Path $PluginsRoot)) {
        Write-Log "Plugins folder not found: $PluginsRoot" -Level WARN -Source 'PluginLoader'
        return
    }

    # Discover plugin folders (any subfolder containing plugin.json)
    $pluginFolders = Get-ChildItem -Path $PluginsRoot -Directory -ErrorAction SilentlyContinue |
                     Where-Object { Test-Path (Join-Path $_.FullName "plugin.json") }

    if ($pluginFolders.Count -eq 0) {
        Write-Log "No plugins found in $PluginsRoot" -Level WARN -Source 'PluginLoader'
        return
    }

    # Load metadata first so we can sort by Order
    $pluginMetas = [System.Collections.ArrayList]::new()
    foreach ($folder in $pluginFolders) {
        try {
            $jsonPath = Join-Path $folder.FullName "plugin.json"
            $json     = Get-Content $jsonPath -Raw -Encoding UTF8
            $meta     = ConvertFrom-Json $json
            $meta | Add-Member -NotePropertyName 'FolderPath' -NotePropertyValue $folder.FullName -Force
            [void]$pluginMetas.Add($meta)
        } catch {
            Write-Log "Failed to read plugin.json in '$($folder.Name)': $($_.Exception.Message)" -Level ERROR -Source 'PluginLoader'
            # Add an error entry so we still show a broken-plugin tab
            $errMeta = [PSCustomObject]@{
                Id         = $folder.Name
                TabTitle   = "⚠ $($folder.Name)"
                Order      = 9999
                FolderPath = $folder.FullName
                _LoadError = "plugin.json parse failed: $($_.Exception.Message)"
            }
            [void]$pluginMetas.Add($errMeta)
        }
    }

    # Sort by Order field
    $sorted = $pluginMetas | Sort-Object { if ($_.Order) { $_.Order } else { 9999 } }

    foreach ($meta in $sorted) {
        Load-SinglePlugin -TabControl $TabControl -Meta $meta
    }
}

function Load-SinglePlugin {
    param(
        [System.Windows.Controls.TabControl]$TabControl,
        [PSCustomObject]$Meta
    )

    $pluginId = $Meta.Id
    Write-Log "Loading plugin: $pluginId" -Source 'PluginLoader'

    # Guard: skip if this plugin ID was already loaded
    $alreadyLoaded = $Script:LoadedPlugins | Where-Object { $_.Id -eq $pluginId }
    if ($alreadyLoaded) {
        Write-Log "  Plugin '$pluginId' already loaded — skipping duplicate" -Level WARN -Source 'PluginLoader'
        return
    }

    # If meta already has a load error (from JSON parse failure)
    if ($Meta.PSObject.Properties['_LoadError']) {
        Add-ErrorTab -TabControl $TabControl -Title $Meta.TabTitle -ErrorMsg $Meta._LoadError
        return
    }

    try {
        $folderPath   = $Meta.FolderPath
        $xamlPath     = Join-Path $folderPath "Tab.xaml"
        $handlersPath = Join-Path $folderPath "Handlers.ps1"
        $functionsPath = Join-Path $folderPath "Functions.ps1"

        # Validate required files exist
        if (-not (Test-Path $xamlPath)) {
            throw "Missing required file: Tab.xaml"
        }
        if (-not (Test-Path $handlersPath)) {
            throw "Missing required file: Handlers.ps1"
        }

        # Load optional Functions.ps1 first (no UI dependency)
        if (Test-Path $functionsPath) {
            try {
                . ([string]$functionsPath)
                Write-Log "  Loaded Functions.ps1 for $pluginId" -Source 'PluginLoader'
            } catch {
                # Functions.ps1 failing is non-fatal for tab display
                Write-Log "  Functions.ps1 error in $pluginId — $($_.Exception.Message)" -Level WARN -Source 'PluginLoader'
            }
        }

        # Load XAML for the tab content
        $xamlContent = Get-Content $xamlPath -Raw -Encoding UTF8
        $xmlDoc      = [xml]$xamlContent
        # Strip x:Class so WPF XAML loader accepts it
        $xmlDoc.DocumentElement.RemoveAttribute("x:Class")
        $reader      = New-Object System.Xml.XmlNodeReader $xmlDoc
        $tabContent  = [Windows.Markup.XamlReader]::Load($reader)

        # Build the TabItem
        $tabItem = New-Object System.Windows.Controls.TabItem
        $tabItem.Header  = $Meta.TabTitle
        $tabItem.Content = $tabContent
        $tabItem.Tag     = $pluginId

        # Apply tab header style (inherit from parent TabControl)
        $TabControl.Items.Add($tabItem) | Out-Null

        # Load Handlers.ps1 — pass the tab content root so it can FindName controls
        # We expose the root element as $PluginRoot so handlers can reference it
        $Script:PluginRoot = $tabContent
        $Script:PluginMeta = $Meta
        . ([string]$handlersPath)

        # Register the plugin in the registry
        $entry = @{
            Id       = $pluginId
            Meta     = $Meta
            TabItem  = $tabItem
            Root     = $tabContent
            Loaded   = $true
            Error    = ""
        }
        [void]$Script:LoadedPlugins.Add($entry)

        Write-Log "  Plugin '$pluginId' loaded OK" -Source 'PluginLoader'

    } catch {
        $errMsg = $_.Exception.Message
        Write-Log "  Plugin '$pluginId' FAILED: $errMsg" -Level ERROR -Source 'PluginLoader'
        Add-ErrorTab -TabControl $TabControl -Title "⚠ $($Meta.TabTitle)" -ErrorMsg $errMsg

        $entry = @{
            Id      = $pluginId
            Meta    = $Meta
            TabItem = $null
            Root    = $null
            Loaded  = $false
            Error   = $errMsg
        }
        [void]$Script:LoadedPlugins.Add($entry)
    }
}

# ---------------------------------------------------------------------------
# Add a red "broken plugin" tab — visible but harmless
# ---------------------------------------------------------------------------
function Add-ErrorTab {
    param(
        [System.Windows.Controls.TabControl]$TabControl,
        [string]$Title,
        [string]$ErrorMsg
    )

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [Windows.Media.Brushes]::Transparent

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.HorizontalAlignment = 'Center'
    $sp.VerticalAlignment   = 'Center'

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.Text              = "⚠"
    $icon.FontSize          = 36
    $icon.Foreground        = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xE7, 0x4C, 0x3C)
    $icon.HorizontalAlignment = 'Center'
    $icon.Margin            = [System.Windows.Thickness]::new(0,0,0,10)

    $msg = New-Object System.Windows.Controls.TextBlock
    $msg.Text               = "Plugin failed to load"
    $msg.Foreground         = [Windows.Media.Brushes]::White
    $msg.FontSize           = 14
    $msg.FontWeight         = [System.Windows.FontWeights]::SemiBold
    $msg.HorizontalAlignment = 'Center'
    $msg.Margin             = [System.Windows.Thickness]::new(0,0,0,8)

    $detail = New-Object System.Windows.Controls.TextBlock
    $detail.Text            = $ErrorMsg
    $detail.Foreground      = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xBD, 0xC3, 0xC7)
    $detail.FontSize        = 11
    $detail.TextWrapping    = [System.Windows.TextWrapping]::Wrap
    $detail.MaxWidth        = 500
    $detail.HorizontalAlignment = 'Center'

    [void]$sp.Children.Add($icon)
    [void]$sp.Children.Add($msg)
    [void]$sp.Children.Add($detail)
    $border.Child = $sp

    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header  = $Title
    $tabItem.Content = $border
    $tabItem.Tag     = "__error"

    # Red foreground on the header
    $tabItem.Foreground = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xE7, 0x4C, 0x3C)

    $TabControl.Items.Add($tabItem) | Out-Null
}

# ---------------------------------------------------------------------------
# Helper for plugins — find a named control inside their own tab root
# ---------------------------------------------------------------------------
function Find-PluginControl {
    param(
        [string]$Name,
        [System.Windows.FrameworkElement]$Root
    )
    $ctrl = $Root.FindName($Name)
    if ($null -eq $ctrl) {
        Write-Log "Control '$Name' not found in plugin root" -Level WARN -Source 'PluginLoader'
    }
    return $ctrl
}