# =============================================================================
# Plugins\UserManagement\Handlers.ps1  —  AD Manager v2.0
# Event wiring for the User Management tab.
# $Script:PluginRoot  — tab root Grid (set by PluginLoader)
# $Script:PluginMeta  — parsed plugin.json
# =============================================================================

Write-Log "UserManagement: registering handlers..." -Source 'UserManagement'

# ── Grab controls ─────────────────────────────────────────────────────────────
$UM = @{}
$umControls = @(
    'UM_TxtInput','UM_BtnSearch','UM_BtnClear',
    'UM_BtnResetPwd','UM_BtnForceChange','UM_BtnClearNeverExp','UM_BtnCopySAM','UM_BtnImport',
    'UM_TxtHint','UM_Grid',
    'UM_TxtStatus','UM_TxtSelected','UM_TxtTotal'
)
foreach ($n in $umControls) {
    $ctrl = $Script:PluginRoot.FindName($n)
    if ($null -ne $ctrl) {
        $UM[$n] = $ctrl
        Write-Log "  Control found: $n" -Source 'UserManagement'
    } else {
        Write-Log "  Control MISSING: $n" -Level WARN -Source 'UserManagement'
    }
}

$Script:UM_Results = $null

# ── Action button helper ──────────────────────────────────────────────────────
function UM-SetActionState {
    param([bool]$Enabled)
    $UM['UM_BtnResetPwd'].IsEnabled    = $Enabled
    $UM['UM_BtnForceChange'].IsEnabled = $Enabled
    $UM['UM_BtnClearNeverExp'].IsEnabled = $Enabled
    $UM['UM_BtnCopySAM'].IsEnabled     = $Enabled
}

# ── Connection event handler ──────────────────────────────────────────────────
Register-ConnectionHandler -Handler {
    param([bool]$Connected)
    Write-Log "UserManagement: connection state changed → Connected=$Connected" -Source 'UserManagement'
    $UM['UM_BtnSearch'].IsEnabled = $Connected
    if (-not $Connected) {
        UM-SetActionState -Enabled $false
        $UM['UM_TxtHint'].Text   = "Connect to a DC to enable actions"
        $UM['UM_TxtStatus'].Text = "Not connected"
    } else {
        $UM['UM_TxtHint'].Text   = "Select rows in the grid to enable actions"
        $UM['UM_TxtStatus'].Text = "Ready — paste accounts and click Search"
    }
}

# =============================================================================
# CLEAR
# =============================================================================
$UM['UM_BtnClear'].Add_Click({
    Write-Log "Clear clicked" -Source 'UserManagement'
    $UM['UM_TxtInput'].Clear()
    $UM['UM_Grid'].ItemsSource = $null
    $UM['UM_TxtStatus'].Text   = "Cleared"
    $UM['UM_TxtSelected'].Text = "0 selected"
    $UM['UM_TxtTotal'].Text    = ""
    UM-SetActionState -Enabled $false
    $UM['UM_TxtHint'].Text     = "Select rows in the grid to enable actions"
})

# =============================================================================
# IMPORT CSV / XLSX
# =============================================================================
$UM['UM_BtnImport'].Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = "Import Files (*.csv;*.xlsx)|*.csv;*.xlsx|CSV Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    $dialog.Title  = "Select User List (CSV or Excel)"

    if ($dialog.ShowDialog()) {
        $path = $dialog.FileName
        $ext  = [System.IO.Path]::GetExtension($path).ToLower()
        Write-Log "Importing $ext file: $path" -Source 'UserManagement'
        
        try {
            $data = $null
            if ($ext -eq ".xlsx") {
                if (Get-Module -ListAvailable ImportExcel) {
                    $data = Import-Excel -Path $path -ErrorAction Stop
                } else {
                    [Windows.MessageBox]::Show("Excel import requires the 'ImportExcel' module.", "Module Missing", 'OK', 'Error') | Out-Null
                    return
                }
            } else {
                $data = Import-Csv -Path $path -ErrorAction Stop
            }

            if (-not $data -or $data.Count -eq 0) {
                [Windows.MessageBox]::Show("The selected file is empty.", "Import Warning", 'OK', 'Warning') | Out-Null
                return
            }

            $headers = $data[0].PSObject.Properties.Name
            if ($headers -notcontains "Name") {
                [Windows.MessageBox]::Show("The file must have a column named 'Name'.", "Import Error", 'OK', 'Error') | Out-Null
                return
            }

            $userList = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($row in $data) {
                $raw = $row.Name
                if ([string]::IsNullOrWhiteSpace($raw)) { continue }
                $sanitized = if ($raw.Contains('\')) { $raw.Split('\')[-1] } else { $raw }
                if (-not [string]::IsNullOrWhiteSpace($sanitized)) { [void]$userList.Add($sanitized.Trim()) }
            }

            if ($userList.Count -gt 0) {
                $UM['UM_TxtInput'].Text = ($userList | Sort-Object) -join "`r`n"
                Set-ShellStatus -Text "✅ Imported $($userList.Count) unique users"
            } else {
                [Windows.MessageBox]::Show("No valid user names found.", "Import Warning", 'OK', 'Warning') | Out-Null
            }
        } catch {
            Write-Log "Import failed: $($_.Exception.Message)" -Level ERROR -Source 'UserManagement'
            [Windows.MessageBox]::Show("Error reading file:`n$($_.Exception.Message)", "Import Error", 'OK', 'Error') | Out-Null
        }
    }
})

# =============================================================================
# SEARCH TIMER (Background Runspace poller)
# =============================================================================
$Script:UMSearchTimer = New-Object System.Windows.Threading.DispatcherTimer
$Script:UMSearchTimer.Interval = [TimeSpan]::FromMilliseconds(350)
$Script:UMSearchTimer.Add_Tick({
    if (-not $Script:UMSearchHandle) { return }

    if ($Script:UMSearchHandle.IsCompleted -or $UM.ContainsKey('__Done')) {
        $Script:UMSearchTimer.Stop()
        try { $Script:UMSearchPS.EndInvoke($Script:UMSearchHandle) } catch {}
        $Script:UMSearchPS.Dispose(); $Script:UMSearchRunspace.Close(); $Script:UMSearchRunspace.Dispose()
        if ($UM.ContainsKey('__Done')) { $UM.Remove('__Done') }
        $Script:UMSearchHandle = $null

        $cnt = if ($Script:UM_Results) { $Script:UM_Results.Count } else { 0 }
        Write-Log "Search finished — $cnt result(s) in grid" -Source 'UserManagement'
        $UM['UM_TxtStatus'].Text      = "$cnt user(s) found"
        $UM['UM_TxtTotal'].Text       = "$cnt found"
        $UM['UM_BtnSearch'].IsEnabled = $true
        Set-ShellProgress -Value -1
        Set-ShellStatus   -Text "Search complete — $cnt result(s)"
    }
})

# =============================================================================
# SEARCH (background runspace) — RESTORED FROM BACKUP
# =============================================================================
$UM['UM_BtnSearch'].Add_Click({

    if (-not (Get-AppConnected)) {
        Write-Log "Search clicked but not connected" -Level WARN -Source 'UserManagement'
        [Windows.MessageBox]::Show("Not connected to a DC.","AD Manager",'OK','Warning') | Out-Null
        return
    }

    $rawInput = $UM['UM_TxtInput'].Text
    $accounts = $rawInput -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    if ($accounts.Count -eq 0) {
        Write-Log "Search clicked with empty input" -Level WARN -Source 'UserManagement'
        [Windows.MessageBox]::Show("Enter at least one account name.","AD Manager",'OK','Warning') | Out-Null
        return
    }

    Write-Log "Search started — $($accounts.Count) account(s) queued" -Source 'UserManagement'
    
    $Script:UM_Results = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $UM['UM_Grid'].ItemsSource = $Script:UM_Results
    $UM['UM_TxtStatus'].Text   = "Searching..."
    $UM['UM_TxtSelected'].Text = "0 selected"
    $UM['UM_TxtTotal'].Text    = ""
    UM-SetActionState -Enabled $false
    $UM['UM_BtnSearch'].IsEnabled = $false
    Set-ShellProgress -Value 0
    Set-ShellStatus   -Text "Searching $($accounts.Count) account(s)..."

    $rsAccounts = $accounts
    $rsDC       = Get-AppDCName
    $rsCred     = Get-AppCredential
    $rsTotal    = $accounts.Count
    $rsResults  = $Script:UM_Results
    $rsWindow   = $Script:Window
    $rsUM       = $UM
    $rsLogFile  = $Script:LogFile
    $rsADProps  = $Script:ADProperties

    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions  = 'ReuseThread'
    $runspace.Open()

    foreach ($v in @('rsAccounts','rsDC','rsCred','rsTotal','rsResults','rsWindow','rsUM','rsLogFile','rsADProps')) {
        $runspace.SessionStateProxy.SetVariable($v, (Get-Variable $v -ValueOnly))
    }

    $ps = [PowerShell]::Create(); $ps.Runspace = $runspace
    [void]$ps.AddScript({

        function RsLog {
            param([string]$m, [string]$l = 'INFO')
            $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            try { Add-Content -Path $rsLogFile -Value "[$stamp][$l][UM-Search] $m" -Encoding UTF8 } catch {}
            $msg = $m; $lvl = $l
            try {
                $rsWindow.Dispatcher.Invoke([action]{
                    $col = switch ($lvl) { 'ERROR'{'Red'} 'WARN'{'Yellow'} default{'DarkCyan'} }
                    $ts = Get-Date -Format "HH:mm:ss"
                    Write-Host "$ts   [UM-Search]      " -NoNewline -ForegroundColor DarkGray
                    Write-Host $msg -ForegroundColor $col
                })
            } catch {}
        }

        function GetOU {
            param([string]$DN)
            if ([string]::IsNullOrWhiteSpace($DN)) { return "" }
            $p = $DN -split ","; $o = $p | Where-Object { $_ -match "^OU=" }
            if ($o) { return ($o -join " > ").Replace("OU=","") }; return ""
        }

        function MustChange {
            param($u)
            try {
                if ($u.PasswordNeverExpires) { return $false }
                return ($u.pwdLastSet -eq 0)
            } catch { return $false }
        }

        RsLog "Runspace started — importing module..."
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            RsLog "ActiveDirectory module ready"
        } catch {
            RsLog "Failed to import AD: $($_.Exception.Message)" 'ERROR'
            $rsWindow.Dispatcher.Invoke([action]{ $rsUM['__Done'] = $true }); return
        }

        $current = 0; $batchSize = 250
        $validAccounts = @($rsAccounts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
        $totalValid = $validAccounts.Count
        $userDict = @{}

        RsLog "Starting bulk search for $totalValid valid account(s)"
        
        for ($i = 0; $i -lt $totalValid; $i += $batchSize) {
            $batchIndex = [Math]::Floor($i / $batchSize) + 1
            $batch      = $validAccounts[$i..[Math]::Min($i + $batchSize - 1, $totalValid - 1)]
            
            $filterParts = [System.Collections.ArrayList]::new()
            foreach ($acc in $batch) {
                $escaped = $acc.Replace('\','\5c').Replace('*','\2a').Replace('(','\28').Replace(')','\29')
                if ($acc -like "*@*") { [void]$filterParts.Add("(userPrincipalName=$escaped)") }
                else { [void]$filterParts.Add("(sAMAccountName=$escaped)") }
            }
            
            $ldapFilter = "(|" + ($filterParts -join "") + ")"
            try {
                $results = Get-ADUser -LDAPFilter $ldapFilter -Server $rsDC -Credential $rsCred -Properties $rsADProps -ErrorAction Stop
                if ($null -ne $results) {
                    foreach ($u in @($results)) {
                        if ($u.SamAccountName)    { $userDict[$u.SamAccountName.ToLower()] = $u }
                        if ($u.UserPrincipalName) { $userDict[$u.UserPrincipalName.ToLower()] = $u }
                    }
                }
            } catch {
                RsLog "Batch LDAP query failed: $($_.Exception.Message)" 'ERROR'
            }
        }

        RsLog "LDAP queries complete — rendering UI..."
        foreach ($account in $rsAccounts) {
            $current++
            $account = $account.Trim()
            if ([string]::IsNullOrWhiteSpace($account)) { continue }

            $rsWindow.Dispatcher.Invoke([action]{ $rsUM['UM_TxtStatus'].Text = "Processing $current / $rsTotal — $account" })

            $row = [ordered]@{
                OpStatus=""; SAM="$account"; DisplayName=""; UPN=""; Email=""; Enabled=""; PwNeverExp=""; PwdLastSet=""; MustChange=""; LastLogon=""; Department=""; Title=""; OU=""; NotFound=$false
            }

            $u = $userDict[$account.ToLower()]
            if ($null -eq $u) {
                RsLog "  → NOT FOUND: $account" 'WARN'
                $row.DisplayName = "NOT FOUND"; $row.NotFound = $true
            } else {
                $row.SAM         = $u.SamAccountName; $row.DisplayName = $u.DisplayName; $row.UPN = $u.UserPrincipalName; $row.Email = $u.EmailAddress
                $row.Enabled     = if ($u.Enabled) { "✅" } else { "❌" }
                $row.PwNeverExp  = if ($u.PasswordNeverExpires) { "✅" } else { "❌" }
                $row.PwdLastSet  = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
                $row.MustChange  = if (MustChange $u) { "✅" } else { "❌" }
                $row.LastLogon   = if ($u.LastLogonDate) { $u.LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
                $row.Department  = $u.Department; $row.Title = $u.Title; $row.OU = GetOU $u.DistinguishedName
            }

            $capRow = $row
            $rsWindow.Dispatcher.Invoke([action]{
                $rsResults.Add((New-Object PSObject -Property $capRow))
                $rsUM['UM_TxtTotal'].Text = "$($rsResults.Count) found"
            })
        }
        $rsWindow.Dispatcher.Invoke([action]{ $rsUM['__Done'] = $true })
    })
    $Script:UMSearchHandle = $ps.BeginInvoke(); $Script:UMSearchPS = $ps; $Script:UMSearchRunspace = $runspace; $Script:UMSearchTimer.Start()
})

# =============================================================================
# GRID SELECTION CHANGED
# =============================================================================
$UM['UM_Grid'].Add_SelectionChanged({
    $sel = $UM['UM_Grid'].SelectedItems; $count = if ($sel) { $sel.Count } else { 0 }
    $UM['UM_TxtSelected'].Text = "$count selected"
    if ($count -gt 0 -and (Get-AppConnected)) { UM-SetActionState -Enabled $true } else { UM-SetActionState -Enabled $false }
})

# =============================================================================
# COPY SAM NAMES
# =============================================================================
$UM['UM_BtnCopySAM'].Add_Click({
    $sel = $UM['UM_Grid'].SelectedItems; if ($null -eq $sel -or $sel.Count -eq 0) { return }
    [Windows.Clipboard]::SetText((@($sel | ForEach-Object { $_.SAM }) -join "`r`n"))
    Set-ShellStatus -Text "✅ Copied SAM name(s)"
})

# =============================================================================
# FORCE CHANGE / RESET DIALOG
# =============================================================================
$UM['UM_BtnForceChange'].Add_Click({
    $sel = $UM['UM_Grid'].SelectedItems; if ($null -eq $sel -or $sel.Count -eq 0) { return }; $sams = @($sel | ForEach-Object { $_.SAM })
    $confirm = [Windows.MessageBox]::Show("Set 'Must change password' for selected?", "Confirm", 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }
    $null = UM-SetForceChange -SAMNames $sams -DCName (Get-AppDCName) -Credential (Get-AppCredential)
    UM-RefreshGridRows -SAMNames $sams -Status "✅"
})

$UM['UM_BtnClearNeverExp'].Add_Click({
    $sel = $UM['UM_Grid'].SelectedItems; if ($null -eq $sel -or $sel.Count -eq 0) { return }; $sams = @($sel | ForEach-Object { $_.SAM })
    $confirm = [Windows.MessageBox]::Show("Clear 'Password never expires' for selected users?", "Confirm", 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }
    $null = UM-ClearNeverExp -SAMNames $sams -DCName (Get-AppDCName) -Credential (Get-AppCredential)
    UM-RefreshGridRows -SAMNames $sams -Status "✅"
})

$UM['UM_BtnResetPwd'].Add_Click({
    $sel = $UM['UM_Grid'].SelectedItems; if ($null -eq $sel -or $sel.Count -eq 0) { return }; $sams = @($sel | ForEach-Object { $_.SAM })
    UM-ShowResetDialog -SAMNames $sams
})

function UM-GeneratePassword {
    param([int]$Length = 14)
    $u='ABCDEFGHJKLMNPQRSTUVWXYZ'; $l='abcdefghjkmnpqrstuvwxyz'; $d='23456789'; $s='!@#$%^&*()-_=+[]{}|;:,.<>?'
    $all = $u+$l+$d+$s; $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider; $chars = [System.Collections.ArrayList]::new()
    foreach($p in @($u,$l,$d,$s)){$buf=New-Object byte[] 1;$rng.GetBytes($buf);[void]$chars.Add($p[$buf[0]%$p.Length])}
    for($i=$chars.Count;$i -lt $Length;$i++){$buf=New-Object byte[] 1;$rng.GetBytes($buf);[void]$chars.Add($all[$buf[0]%$all.Length])}
    for($i=$chars.Count-1;$i-gt 0;$i--){$buf=New-Object byte[] 4;$rng.GetBytes($buf);$j=[BitConverter]::ToUInt32($buf,0)%($i+1);$t=$chars[$i];$chars[$i]=$chars[$j];$chars[$j]=$t}
    $rng.Dispose(); return -join $chars
}

function UM-ShowResetDialog {
    param([string[]]$SAMNames)
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Reset Password" Height="680" Width="480" WindowStartupLocation="CenterOwner" Background="#1A252F" FontFamily="Segoe UI">
    <Grid><ScrollViewer VerticalScrollBarVisibility="Auto"><Grid Margin="20">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="80"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="🔑 Reset Password" Foreground="White" FontSize="16" FontWeight="Bold" Margin="0,0,0,12"/>
        <Border Grid.Row="2" Background="#243447" CornerRadius="3" Padding="8" Margin="0,0,0,12"><ScrollViewer><TextBlock x:Name="DlgUserList" Foreground="#ECF0F1" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap"/></ScrollViewer></Border>
        <Border Grid.Row="3" Background="#1E2D3D" CornerRadius="4" Padding="12" Margin="0,0,0,12"><StackPanel>
            <TextBlock Text="🎲 AUTO-GENERATE" Foreground="#5DADE2" FontSize="10" FontWeight="Bold" Margin="0,0,0,10"/>
            <Grid Margin="0,0,0,8"><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <Slider x:Name="DlgSlider" Grid.Column="1" Minimum="12" Maximum="24" Value="14" IsSnapToTickEnabled="True"/><TextBlock x:Name="DlgLengthLabel" Grid.Column="2" Text="14" Foreground="White" Margin="10,0,0,0"/>
            </Grid>
            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="5"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="5"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <Border Background="#2C3E50" Padding="8,4"><TextBlock x:Name="DlgGenPwd" Text="---" Foreground="#7F8C8D" FontFamily="Consolas"/></Border>
                <Button x:Name="DlgBtnGenerate" Grid.Column="2" Content="⚡ Gen" Background="#2980B9" Foreground="White" Width="60" Height="30"/>
                <Button x:Name="DlgBtnCopy" Grid.Column="4" Content="📋" Background="#566573" Foreground="White" Width="32" Height="30"/>
            </Grid>
            <Button x:Name="DlgBtnGenApply" Content="⚡ Generate &amp; Apply Now" Background="#27AE60" Foreground="White" FontWeight="Bold" Height="34" Margin="0,10,0,0"/>
        </StackPanel></Border>
        <StackPanel Grid.Row="5" Margin="0,0,0,12">
            <PasswordBox x:Name="DlgPwd1" Background="#2C3E50" Foreground="White" Height="30" Margin="0,0,0,6"/><PasswordBox x:Name="DlgPwd2" Background="#2C3E50" Foreground="White" Height="30"/>
        </StackPanel>
        <CheckBox x:Name="DlgMustChange" Grid.Row="7" Content="Must change at logon" Foreground="#ECF0F1" IsChecked="True" Margin="0,0,0,12"/>
        <TextBlock x:Name="DlgValidation" Grid.Row="8" Foreground="#E74C3C" FontSize="11" Margin="0,0,0,12" TextWrapping="Wrap"/>
        <Grid Grid.Row="9"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions>
            <Button x:Name="DlgApply" Content="✅ Update" Background="#E67E22" Foreground="White" Height="36"/><Button x:Name="DlgCancel" Grid.Column="2" Content="Cancel" Height="36"/>
        </Grid>
    </Grid></ScrollViewer></Grid>
</Window>
"@
    $xmlDoc = [xml]$xaml; $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    $dlg = [Windows.Markup.XamlReader]::Load($reader); $dlg.Owner = $Script:Window
    $controls = @('DlgUserList','DlgSlider','DlgLengthLabel','DlgGenPwd','DlgBtnGenerate','DlgBtnCopy','DlgBtnGenApply','DlgPwd1','DlgPwd2','DlgMustChange','DlgValidation','DlgApply','DlgCancel')
    $dlgCtrl = @{}; foreach($c in $controls) { $dlgCtrl[$c] = $dlg.FindName($c) }
    $dlgCtrl['DlgUserList'].Text = $SAMNames -join "`n"; $Script:LastGen = ""

    function Internal-Reset {
        param([System.Security.SecureString]$p1, [System.Security.SecureString]$p2)
        if ($p1.Length -lt 8) { $dlgCtrl['DlgValidation'].Text = "⚠ Too short"; return $false }
        $results = UM-ResetPasswords -SAMNames $SAMNames -NewPassword $p1 -MustChange ($dlgCtrl['DlgMustChange'].IsChecked -eq $true) -DCName (Get-AppDCName) -Credential (Get-AppCredential)
        if (($results | Where-Object { -not $_.Success }).Count -eq 0) { $dlg.Close(); return $true }
        else { [Windows.MessageBox]::Show("Errors occurred.", "Error"); $dlg.Close(); return $false }
    }
    $dlgCtrl['DlgSlider'].Add_ValueChanged({ $dlgCtrl['DlgLengthLabel'].Text = "$([int]$dlgCtrl['DlgSlider'].Value)" })
    $dlgCtrl['DlgBtnGenerate'].Add_Click({ $new = UM-GeneratePassword -Length ([int]$dlgCtrl['DlgSlider'].Value); $Script:LastGen = $new; $dlgCtrl['DlgGenPwd'].Text = $new; $dlgCtrl['DlgGenPwd'].Foreground = [Windows.Media.Brushes]::White })
    $dlgCtrl['DlgBtnCopy'].Add_Click({ if ($Script:LastGen) { [Windows.Clipboard]::SetText($Script:LastGen) } })
    $dlgCtrl['DlgBtnGenApply'].Add_Click({
        $new = UM-GeneratePassword -Length ([int]$dlgCtrl['DlgSlider'].Value); $Script:LastGen = $new; [Windows.Clipboard]::SetText($new); $sec = ConvertTo-SecureString $new -AsPlainText -Force
        if (Internal-Reset -p1 $sec -p2 $sec) { UM-RefreshGridRows -SAMNames $SAMNames -Status "✅" }
    })
    $dlgCtrl['DlgApply'].Add_Click({ if (Internal-Reset -p1 $dlgCtrl['DlgPwd1'].SecurePassword -p2 $dlgCtrl['DlgPwd2'].SecurePassword) { UM-RefreshGridRows -SAMNames $SAMNames -Status "✅" } })
    $dlgCtrl['DlgCancel'].Add_Click({ $dlg.Close() }); [void]$dlg.ShowDialog()
}

function UM-RefreshGridRows {
    param([string[]]$SAMNames, [string]$Status)
    if (-not (Get-AppConnected))      { return }
    if ($null -eq $Script:UM_Results) { return }
    foreach ($sam in $SAMNames) {
        $fresh = UM-RefreshUserRow -SAM $sam -DCName (Get-AppDCName) -Credential (Get-AppCredential)
        if ($null -eq $fresh) { continue }
        for ($i = 0; $i -lt $Script:UM_Results.Count; $i++) {
            if ($Script:UM_Results[$i].SAM -eq $sam) {
                $obj = New-Object PSObject -Property $fresh
                if ($Status) { $obj.OpStatus = $Status }
                $Script:UM_Results[$i] = $obj; break
            }
        }
    }
}
Write-Log "UserManagement handlers registered successfully" -Source 'UserManagement'
