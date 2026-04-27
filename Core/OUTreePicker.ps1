# =============================================================================
# Core\OUTreePicker.ps1  —  AD Manager v2.0
# Shared lazy-loaded WPF TreeView OU picker.
# Used by: ComputerMapper (single-select), OUMover (multi-select)
#
# USAGE:
#   $result = Show-OUPicker -Owner $Script:Window -MultiSelect $false
#   Returns: single DN string (single) or [string[]] array (multi)
#   Returns: $null if user cancelled
#
# LAZY LOAD:
#   Root nodes load immediately on open.
#   Child nodes load on first expand of each node — no upfront full-tree load.
#   Handles domains with 1000+ OUs efficiently.
# =============================================================================

function Show-OUPicker {
    param(
        [System.Windows.Window]$Owner,
        [bool]$MultiSelect = $false,
        [string]$Title = "Select Organisational Unit"
    )

    Write-Log "OUTreePicker: opening ($( if($MultiSelect){'multi'}else{'single'} )-select)" -Source 'OUTreePicker'

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Height="560" Width="520"
        WindowStartupLocation="CenterOwner"
        Background="#1A252F" ResizeMode="CanResize"
        MinHeight="400" MinWidth="380"
        FontFamily="Segoe UI">
    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0"
                   Text="$Title"
                   Foreground="White" FontSize="14" FontWeight="SemiBold"
                   Margin="0,0,0,10"/>

        <!-- Search box -->
        <Grid Grid.Row="1" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="TxtSearch" Grid.Column="0"
                     Background="#243447" Foreground="#ECF0F1"
                     BorderBrush="#4A6278" BorderThickness="1"
                     Padding="6,5" CaretBrush="White"
                     ToolTip="Type to filter OU names"/>
            <Button x:Name="BtnSearch" Grid.Column="1"
                    Content="🔍" Width="34" Height="30" Margin="4,0,0,0"
                    Background="#2980B9" Foreground="White"
                    BorderThickness="0" Cursor="Hand"/>
        </Grid>

        <!-- Tree -->
        <Border Grid.Row="2" Background="#1E2D3D"
                BorderBrush="#34495E" BorderThickness="1" CornerRadius="4">
            <TreeView x:Name="OUTree"
                      Background="Transparent"
                      BorderThickness="0"
                      Foreground="#ECF0F1"
                      Padding="4"/>
        </Border>

        <!-- Selected display -->
        <Border Grid.Row="3" Background="#243447" CornerRadius="3"
                Padding="8,5" Margin="0,8,0,0">
            <TextBlock x:Name="TxtSelected"
                       Foreground="#7F8C8D" FontSize="11"
                       TextWrapping="Wrap"
                       Text="No selection"/>
        </Border>

        <!-- Buttons -->
        <Grid Grid.Row="4" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="80"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0"
                       x:Name="TxtStatus"
                       Foreground="#7F8C8D" FontSize="11"
                       VerticalAlignment="Center"
                       Text="Loading..."/>
            <Button x:Name="BtnOK" Grid.Column="2"
                    Content="✅ Select"
                    Background="#27AE60" Foreground="White"
                    FontWeight="SemiBold" Height="34"
                    BorderThickness="0" Cursor="Hand"
                    IsEnabled="False"/>
            <Button x:Name="BtnCancel" Grid.Column="4"
                    Content="Cancel"
                    Background="#7F8C8D" Foreground="White"
                    FontWeight="SemiBold" Height="34"
                    BorderThickness="0" Cursor="Hand"/>
        </Grid>
    </Grid>
</Window>
"@

    $xmlDoc = [xml]$xaml
    $reader  = New-Object System.Xml.XmlNodeReader $xmlDoc
    $dlg     = [Windows.Markup.XamlReader]::Load($reader)
    $dlg.Owner = $Owner

    # Controls
    $tree       = $dlg.FindName('OUTree')
    $txtSearch  = $dlg.FindName('TxtSearch')
    $btnSearch  = $dlg.FindName('BtnSearch')
    $txtSel     = $dlg.FindName('TxtSelected')
    $txtStatus  = $dlg.FindName('TxtStatus')
    $btnOK      = $dlg.FindName('BtnOK')
    $btnCancel  = $dlg.FindName('BtnCancel')

    # Result storage
    $Script:OUPicker_Result     = $null
    $Script:OUPicker_MultiSel   = [System.Collections.ArrayList]::new()
    $Script:OUPicker_IsMulti    = $MultiSelect
    $Script:OUPicker_AllNodes   = [System.Collections.ArrayList]::new()  # for search

    # ── TreeView styling ─────────────────────────────────────────────────────
    $itemStyle = New-Object System.Windows.Style([System.Windows.Controls.TreeViewItem])
    $itemStyle.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.TreeViewItem]::ForegroundProperty,
        [Windows.Media.Brushes]::White)))
    $itemStyle.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.TreeViewItem]::BackgroundProperty,
        [Windows.Media.Brushes]::Transparent)))

    # ── Helper: make a TreeViewItem for an OU ────────────────────────────────
    function New-OUTreeItem {
        param([string]$Name, [string]$DN, [bool]$HasChildren)

        $item = New-Object System.Windows.Controls.TreeViewItem
        $item.Tag = $DN

        # Header: icon + name
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = 'Horizontal'

        $icon = New-Object System.Windows.Controls.TextBlock
        $icon.Text     = "🗂 "
        $icon.FontSize = 12
        $icon.VerticalAlignment = 'Center'

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text              = $Name
        $lbl.FontSize          = 12
        $lbl.VerticalAlignment = 'Center'
        $lbl.ToolTip           = $DN

        [void]$sp.Children.Add($icon)
        [void]$sp.Children.Add($lbl)
        $item.Header = $sp

        # If has children, add a dummy node so the expand arrow shows
        if ($HasChildren) {
            $dummy = New-Object System.Windows.Controls.TreeViewItem
            $dummy.Tag    = '__loading__'
            $dummy.Header = "Loading..."
            $dummy.Foreground = [Windows.Media.Brushes]::Gray
            [void]$item.Items.Add($dummy)
        }

        # Multi-select: add checkbox to header
        if ($Script:OUPicker_IsMulti) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Margin          = [System.Windows.Thickness]::new(0,0,6,0)
            $cb.VerticalAlignment = 'Center'
            $cb.Tag             = $DN

            $cb.Add_Checked({
                param($s,$e)
                $dn = $s.Tag
                if (-not $Script:OUPicker_MultiSel.Contains($dn)) {
                    [void]$Script:OUPicker_MultiSel.Add($dn)
                }
                Update-PickerSelection
            })
            $cb.Add_Unchecked({
                param($s,$e)
                [void]$Script:OUPicker_MultiSel.Remove($s.Tag)
                Update-PickerSelection
            })

            # Insert checkbox before icon in header
            $sp2 = New-Object System.Windows.Controls.StackPanel
            $sp2.Orientation = 'Horizontal'
            [void]$sp2.Children.Add($cb)
            [void]$sp2.Children.Add($icon)
            [void]$sp2.Children.Add($lbl)
            $item.Header = $sp2
        }

        [void]$Script:OUPicker_AllNodes.Add(@{ Item=$item; Name=$Name; DN=$DN })
        return $item
    }

    # ── Update selection display ──────────────────────────────────────────────
    function Update-PickerSelection {
        if ($Script:OUPicker_IsMulti) {
            $count = $Script:OUPicker_MultiSel.Count
            if ($count -eq 0) {
                $txtSel.Text = "No OUs selected"
                $btnOK.IsEnabled = $false
            } else {
                $names = $Script:OUPicker_MultiSel | ForEach-Object {
                    ($_ -split ',')[0].Replace('OU=','').Replace('CN=','')
                }
                $txtSel.Text = "$count selected: $($names -join ', ')"
                $btnOK.IsEnabled = $true
            }
        } else {
            if ($null -ne $Script:OUPicker_Result) {
                $name = ($Script:OUPicker_Result -split ',')[0].Replace('OU=','').Replace('CN=','')
                $txtSel.Text     = $name
                $txtSel.ToolTip  = $Script:OUPicker_Result
                $btnOK.IsEnabled = $true
            }
        }
    }

    # ── Load root OUs ─────────────────────────────────────────────────────────
    function Load-RootOUs {
        $txtStatus.Text = "Loading root OUs..."
        $tree.Items.Clear()
        [void]$Script:OUPicker_AllNodes.Clear()

        try {
            $dc   = Get-AppDCName
            $cred = Get-AppCredential

            # Get domain root DN
            $domain    = Get-ADDomain -Server $dc -Credential $cred -ErrorAction Stop
            $rootDN    = $domain.DistinguishedName
            $domainName = $domain.DNSRoot

            # Add domain root node (always top)
            $rootItem = New-OUTreeItem -Name "🌐 $domainName" -DN $rootDN -HasChildren $true
            $rootItem.IsExpanded = $false
            [void]$tree.Items.Add($rootItem)

            # Load first-level OUs
            $topOUs = Get-ADOrganizationalUnit `
                -Filter * `
                -SearchBase $rootDN `
                -SearchScope OneLevel `
                -Server $dc `
                -Credential $cred `
                -Properties Name,DistinguishedName `
                -ErrorAction Stop |
                Sort-Object Name

            foreach ($ou in $topOUs) {
                # Check if this OU has children (sub-OUs)
                $childCount = @(Get-ADOrganizationalUnit `
                    -Filter * `
                    -SearchBase $ou.DistinguishedName `
                    -SearchScope OneLevel `
                    -Server $dc `
                    -Credential $cred `
                    -ErrorAction SilentlyContinue).Count

                $node = New-OUTreeItem -Name $ou.Name -DN $ou.DistinguishedName -HasChildren ($childCount -gt 0)
                [void]$tree.Items.Add($node)
            }

            $total = $tree.Items.Count
            $txtStatus.Text = "$total top-level OUs loaded"
            Write-Log "OUTreePicker: loaded $total root items" -Source 'OUTreePicker'

        } catch {
            $txtStatus.Text = "Error: $($_.Exception.Message)"
            Write-Log "OUTreePicker: root load failed — $($_.Exception.Message)" -Level ERROR -Source 'OUTreePicker'
        }
    }

    # ── Lazy load children when a node expands ────────────────────────────────
    $tree.Add_Expanded({
        param($sender, $e)
        $item = $e.Source

        # Check if first child is the dummy loading node
        if ($item.Items.Count -eq 1 -and $item.Items[0].Tag -eq '__loading__') {
            $item.Items.Clear()
            $parentDN = $item.Tag

            try {
                $dc   = Get-AppDCName
                $cred = Get-AppCredential

                $children = Get-ADOrganizationalUnit `
                    -Filter * `
                    -SearchBase $parentDN `
                    -SearchScope OneLevel `
                    -Server $dc `
                    -Credential $cred `
                    -Properties Name,DistinguishedName `
                    -ErrorAction Stop |
                    Sort-Object Name

                foreach ($ou in $children) {
                    $childCount = @(Get-ADOrganizationalUnit `
                        -Filter * `
                        -SearchBase $ou.DistinguishedName `
                        -SearchScope OneLevel `
                        -Server $dc `
                        -Credential $cred `
                        -ErrorAction SilentlyContinue).Count

                    $node = New-OUTreeItem -Name $ou.Name -DN $ou.DistinguishedName -HasChildren ($childCount -gt 0)
                    [void]$item.Items.Add($node)
                }

                if ($item.Items.Count -eq 0) {
                    $empty = New-Object System.Windows.Controls.TreeViewItem
                    $empty.Header    = "(no sub-OUs)"
                    $empty.Foreground = [Windows.Media.Brushes]::Gray
                    $empty.IsEnabled = $false
                    [void]$item.Items.Add($empty)
                }

                Write-Log "OUTreePicker: expanded '$parentDN' — $($item.Items.Count) children" -Source 'OUTreePicker'

            } catch {
                $errItem = New-Object System.Windows.Controls.TreeViewItem
                $errItem.Header     = "Error loading: $($_.Exception.Message)"
                $errItem.Foreground = [Windows.Media.Brushes]::OrangeRed
                [void]$item.Items.Add($errItem)
                Write-Log "OUTreePicker: expand error '$parentDN' — $($_.Exception.Message)" -Level WARN -Source 'OUTreePicker'
            }
        }
    })

    # ── Single-select: click to select ───────────────────────────────────────
    if (-not $MultiSelect) {
        $tree.Add_SelectedItemChanged({
            $sel = $tree.SelectedItem
            if ($null -ne $sel -and $sel.Tag -ne '__loading__' -and $sel.Tag -ne $null) {
                $Script:OUPicker_Result = $sel.Tag
                Update-PickerSelection
            }
        })
    }

    # ── Search / filter ───────────────────────────────────────────────────────
    function Apply-OUSearch {
        $term = $txtSearch.Text.Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($term)) {
            $txtStatus.Text = "Type to filter"
            return
        }

        $matches = $Script:OUPicker_AllNodes | Where-Object { $_.Name.ToLower() -like "*$term*" }
        $txtStatus.Text = "$($matches.Count) match(es) for '$($txtSearch.Text)'"

        # Highlight matched items by expanding their parents
        foreach ($m in $matches) {
            $m.Item.BringIntoView()
            $m.Item.IsSelected = $true
        }
    }

    $btnSearch.Add_Click({ Apply-OUSearch })
    $txtSearch.Add_KeyDown({
        param($s,$e)
        if ($e.Key -eq 'Return') { Apply-OUSearch }
    })

    # ── OK / Cancel ───────────────────────────────────────────────────────────
    $btnOK.Add_Click({
        Write-Log "OUTreePicker: selection confirmed" -Source 'OUTreePicker'
        $dlg.DialogResult = $true
        $dlg.Close()
    })
    $btnCancel.Add_Click({
        $Script:OUPicker_Result   = $null
        $Script:OUPicker_MultiSel.Clear()
        $dlg.Close()
    })

    # ── Load on open ─────────────────────────────────────────────────────────
    $dlg.Add_Loaded({ Load-RootOUs })

    # ── Show dialog ───────────────────────────────────────────────────────────
    $dlgResult = $dlg.ShowDialog()

    if ($dlgResult -eq $true) {
        if ($MultiSelect) {
            $selected = @($Script:OUPicker_MultiSel)
            Write-Log "OUTreePicker: returning $($selected.Count) selected OUs" -Source 'OUTreePicker'
            return $selected
        } else {
            Write-Log "OUTreePicker: returning '$($Script:OUPicker_Result)'" -Source 'OUTreePicker'
            return $Script:OUPicker_Result
        }
    }

    Write-Log "OUTreePicker: cancelled" -Source 'OUTreePicker'
    return $null
}
