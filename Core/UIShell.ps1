# =============================================================================
# Core\UIShell.ps1  —  AD Manager v2.0
# Loads Shell.xaml, wires the connection bar, hosts the plugin TabControl.
#
# CONNECTION FIX:
#   Old: Get-ADUser -Filter {SamAccountName -eq "krbtgt"}  ← full LDAP search, slow
#   New: Get-ADDomain -Server $DC -Credential $Cred         ← single LDAP bind, fast
#   This matches how AD Sentinel connects — typically <2 seconds vs 10-30 seconds.
# =============================================================================

$Script:Shell = @{}

# =============================================================================
# Initialize-Shell  —  load XAML, harvest controls, wire events
# =============================================================================
function Initialize-Shell {
    param([string]$XamlPath)

    Write-LogStep "Loading Shell XAML: $XamlPath"

    if (-not (Test-Path $XamlPath)) {
        throw "Shell XAML not found: $XamlPath"
    }

    try {
        $xaml = [xml](Get-Content $XamlPath -Raw -Encoding UTF8)
        $xaml.DocumentElement.RemoveAttribute("x:Class")
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-LogStep "XAML parsed and loaded" -Status 'OK'
    } catch {
        throw "XAML load failed: $($_.Exception.Message)"
    }

    # ── Harvest named controls ────────────────────────────────────────────────
    Write-LogStep "Harvesting shell controls..."
    $names = @(
        'ConnStatusDot','ConnStatusLabel',
        'TxtDCName','TxtUsername','PwdPassword',
        'BtnConnect', 'BtnDisconnect', 'BtnReloadPlugins', 'TxtConnMessage',
        'MainTabControl',
        'TxtStatusBar','ProgressBar'
    )
    $missing = 0
    foreach ($n in $names) {
        $ctrl = $window.FindName($n)
        if ($null -ne $ctrl) {
            $Script:Shell[$n] = $ctrl
        } else {
            Write-Log "Shell control '$n' not found in XAML" -Level WARN -Source 'UIShell'
            $missing++
        }
    }
    if ($missing -eq 0) {
        Write-LogStep "All $($names.Count) shell controls found" -Status 'OK'
    } else {
        Write-LogStep "$missing control(s) missing from XAML" -Status 'FAIL'
    }

    Register-ShellConnectionHandlers
    Write-Log "Shell initialised successfully" -Source 'UIShell'
    return $window
}

# =============================================================================
# CONNECTION HANDLERS
# =============================================================================
function Register-ShellConnectionHandlers {

    Write-LogStep "Registering connection bar handlers..."

    # ── Connection Background Timer ───────────────────────────────────────────
    $Script:ConnTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:ConnTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $Script:ConnTimer.Add_Tick({
        if (-not $Script:ConnHandle) { return }

        $elapsed = ([Environment]::TickCount - $Script:ConnStartTick) / 1000.0
        if ($elapsed -lt 0) { $elapsed = 0 }

        if ($Script:ConnHandle.IsCompleted) {
            $Script:ConnTimer.Stop()
            try { $Script:ConnPS.EndInvoke($Script:ConnHandle) } catch {}
            $Script:ConnPS.Dispose(); $Script:ConnRS.Close(); $Script:ConnRS.Dispose()
            
            $Script:ConnHandle = $null

            $res = $Script:Shell['__ConnResult']
            $Script:Shell.Remove('__ConnResult')

            if ($res -and $res.OK) {
                Write-Log "Connection established in $([Math]::Round($elapsed,1))s — Domain: $($res.DomainDNS)  PDC: $($res.PDC)" -Source 'Connection'
                $Script:AppState['DomainDNS']  = $res.DomainDNS
                $Script:AppState['NetBIOS']    = $res.NetBIOS
                $Script:AppState['DomainMode'] = $res.DomainMode
                $Script:AppState['PDC']        = $res.PDC
                Set-AppConnected -DCName $Script:ConnDC -Credential $Script:ConnCred
                Update-ShellConnectedState -Connected $true
                Set-ShellMessage -Text "✅ Connected in $([Math]::Round($elapsed,1))s" -IsError $false
            } else {
                $msg = if ($res) { $res.Msg } else { "Unknown error — check logs" }
                Write-Log "Connection FAILED after $([Math]::Round($elapsed,1))s — $msg" -Level ERROR -Source 'Connection'
                Set-ShellMessage -Text "❌ $msg" -IsError $true
                Set-ShellStatus  -Text "Connection failed — check DC name and credentials"
                $Script:Shell['BtnConnect'].IsEnabled = $true
                $Script:Shell['BtnConnect'].Content   = "🔌 Connect"
            }

        } elseif ($elapsed -gt 2) {
            Set-ShellStatus -Text "Connecting to $($Script:ConnDC) ... $([Math]::Round($elapsed,0))s"
        }
    })

    # ── CONNECT ───────────────────────────────────────────────────────────────
    $Script:Shell['BtnConnect'].Add_Click({

        $dcName   = $Script:Shell['TxtDCName'].Text.Trim()
        $username = $Script:Shell['TxtUsername'].Text.Trim()
        $password = $Script:Shell['PwdPassword'].SecurePassword

        Write-Log "Connect clicked — DC='$dcName'  User='$username'" -Source 'Connection'

        # ── Input validation ──────────────────────────────────────────────────
        $validationMsg = ""
        if ([string]::IsNullOrWhiteSpace($dcName))      { $validationMsg = "⚠ Enter a DC name or IP address" }
        elseif ([string]::IsNullOrWhiteSpace($username)) { $validationMsg = "⚠ Enter a username (DOMAIN\user)" }
        elseif ($password.Length -eq 0)                  { $validationMsg = "⚠ Enter your password" }

        if ($validationMsg -ne "") {
            Write-Log "Validation failed: $validationMsg" -Level WARN -Source 'Connection'
            Set-ShellMessage -Text $validationMsg -IsError $true
            return
        }

        # ── Build PSCredential ────────────────────────────────────────────────
        try {
            $cred = New-Object System.Management.Automation.PSCredential($username, $password)
            Write-Log "PSCredential built for '$username'" -Source 'Connection'
        } catch {
            Write-Log "PSCredential build failed: $($_.Exception.Message)" -Level ERROR -Source 'Connection'
            Set-ShellMessage -Text "⚠ Invalid credential format" -IsError $true
            return
        }

        # ── Lock UI ───────────────────────────────────────────────────────────
        $Script:Shell['BtnConnect'].IsEnabled = $false
        $Script:Shell['BtnConnect'].Content   = "⏳ Connecting..."
        Set-ShellMessage -Text "" -IsError $false
        Set-ShellStatus  -Text "Connecting to $dcName ..."

        # ── Capture vars for runspace ─────────────────────────────────────────
        $dcCapture   = $dcName
        $credCapture = $cred
        $window      = $Script:Window
        $shell       = $Script:Shell
        $logFile     = $Script:LogFile

        Write-Log "Starting connection runspace for DC: $dcCapture" -Source 'Connection'
        $Script:ConnDC   = $dcCapture
        $Script:ConnCred = $credCapture

        # ── Background thread — inherits current session state (modules already loaded)
        # KEY: We use InitialSessionState::CreateDefault2() so the ActiveDirectory
        # module that was already imported in Launch.ps1 is available immediately.
        # No cold module load = connection in 1-3 seconds instead of 2+ minutes.
        # ─────────────────────────────────────────────────────────────────────────
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
        $rs  = [RunspaceFactory]::CreateRunspace($iss)
        $rs.ApartmentState = 'MTA'   # MTA is fine for AD calls (no WPF here)
        $rs.ThreadOptions  = 'ReuseThread'
        $rs.Open()

        $rs.SessionStateProxy.SetVariable('DC',      $dcCapture)
        $rs.SessionStateProxy.SetVariable('Cred',    $credCapture)
        $rs.SessionStateProxy.SetVariable('Window',  $window)
        $rs.SessionStateProxy.SetVariable('Shell',   $shell)
        $rs.SessionStateProxy.SetVariable('LogFile', $logFile)

        $ps = [PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript({

            function RsLog {
                param([string]$m, [string]$l = 'INFO')
                $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                try { Add-Content -Path $LogFile -Value "[$stamp][$l][Conn] $m" -Encoding UTF8 } catch {}
                $msg = $m; $lvl = $l
                $Window.Dispatcher.Invoke([action]{
                    $col = switch ($lvl) { 'ERROR'{'Red'} 'WARN'{'Yellow'} default{'DarkCyan'} }
                    $ts  = Get-Date -Format "HH:mm:ss"
                    Write-Host "$ts  [Conn]          " -NoNewline -ForegroundColor DarkGray
                    Write-Host $msg -ForegroundColor $col
                })
            }

            $result = $null

            # ── STEP 1: Ensure AD module is available ─────────────────────────
            # CreateDefault2 inherits loaded modules — this is usually instant.
            RsLog "Step 1/3 — Checking ActiveDirectory module..."
            try {
                if (-not (Get-Module -Name ActiveDirectory)) {
                    Import-Module ActiveDirectory -ErrorAction Stop
                    RsLog "ActiveDirectory module imported"
                } else {
                    RsLog "ActiveDirectory module already loaded — skipping import"
                }
            } catch {
                $result = @{ OK = $false; Msg = "Cannot load ActiveDirectory module: $($_.Exception.Message)" }
                $r = $result
                $Window.Dispatcher.Invoke([action]{ $Shell['__ConnResult'] = $r })
                return
            }

            # ── STEP 2: Quick TCP port 389 check (faster + more reliable than ICMP)
            # ICMP is often blocked on DC networks. Port 389 (LDAP) must be open.
            RsLog "Step 2/3 — TCP port 389 (LDAP) check on $DC ..."
            $tcpOk = $false
            try {
                $tcp = New-Object System.Net.Sockets.TcpClient
                $iar = $tcp.BeginConnect($DC, 389, $null, $null)
                $waited = $iar.AsyncWaitHandle.WaitOne(3000, $false)   # 3s timeout
                if ($waited -and $tcp.Connected) {
                    $tcpOk = $true
                    RsLog "Port 389 open — DC is reachable"
                } else {
                    RsLog "Port 389 not reachable in 3s — will attempt LDAP bind anyway" 'WARN'
                }
                $tcp.Close()
            } catch {
                RsLog "TCP check exception: $($_.Exception.Message) — continuing anyway" 'WARN'
            }

            # ── STEP 3: LDAP bind via Get-ADDomain ───────────────────────────
            # Single lightweight LDAP bind — returns domain info we store for plugins.
            RsLog "Step 3/3 — LDAP bind: Get-ADDomain -Server $DC ..."
            try {
                $domain = Get-ADDomain -Server $DC -Credential $Cred -ErrorAction Stop
                RsLog "Connected — Domain: $($domain.DNSRoot)  NetBIOS: $($domain.NetBIOSName)  PDC: $($domain.PDCEmulator)"
                $result = @{
                    OK         = $true
                    Msg        = "OK"
                    DomainDNS  = $domain.DNSRoot
                    NetBIOS    = $domain.NetBIOSName
                    DomainMode = [string]$domain.DomainMode
                    PDC        = $domain.PDCEmulator
                }
            } catch {
                $errMsg = $_.Exception.Message
                RsLog "LDAP bind failed: $errMsg" 'ERROR'
                $result = @{ OK = $false; Msg = $errMsg }
            }

            $r = $result
            $Window.Dispatcher.Invoke([action]{ $Shell['__ConnResult'] = $r })
        })

        $Script:ConnHandle    = $ps.BeginInvoke()
        $Script:ConnStartTick = [Environment]::TickCount
        $Script:ConnPS        = $ps
        $Script:ConnRS        = $rs
        
        $Script:ConnTimer.Start()
    })

    # ── DISCONNECT ────────────────────────────────────────────────────────────
    $Script:Shell['BtnDisconnect'].Add_Click({
        Write-Log "Disconnect clicked — was connected to $(Get-AppDCName)" -Source 'Connection'
        Set-AppDisconnected
        Update-ShellConnectedState -Connected $false
        $Script:Shell['PwdPassword'].Clear()
        Set-ShellMessage -Text "" -IsError $false
        Set-ShellStatus  -Text "Disconnected — enter DC details and click Connect"
    })

    # ── RELOAD PLUGINS ────────────────────────────────────────────────────────
    if ($Script:Shell['BtnReloadPlugins']) {
        $Script:Shell['BtnReloadPlugins'].Add_Click({
            Write-Log "Reload plugins clicked" -Source 'UIShell'
            Set-ShellStatus -Text "Reloading plugins..."
            
            try {
                # 1. Reset
                Reset-Plugins -TabControl $Script:Shell['MainTabControl']
                
                # 2. Re-import
                Import-Plugins -TabControl $Script:Shell['MainTabControl'] -PluginsRoot $Script:PluginsDir
                
                Set-ShellStatus -Text "✅ Plugins reloaded successfully"
            } catch {
                Write-Log "Reload FAILED: $($_.Exception.Message)" -Level ERROR -Source 'UIShell'
                Set-ShellStatus -Text "❌ Reload failed: $($_.Exception.Message)"
            }
        })
    }

    # ── ENTER KEY in password box triggers Connect ────────────────────────────
    $Script:Shell['PwdPassword'].Add_KeyDown({
        param($s, $e)
        if ($e.Key -eq 'Return' -and $Script:Shell['BtnConnect'].IsEnabled) {
            $Script:Shell['BtnConnect'].RaiseEvent(
                [System.Windows.RoutedEventArgs]::new(
                    [System.Windows.Controls.Button]::ClickEvent))
        }
    })

    Write-LogStep "Connection handlers registered" -Status 'OK'
}

# =============================================================================
# SHELL UI HELPERS
# =============================================================================

function Update-ShellConnectedState {
    param([bool]$Connected)

    if ($Connected) {
        $Script:Shell['ConnStatusDot'].Fill         = [Windows.Media.Brushes]::LimeGreen
        $Script:Shell['ConnStatusLabel'].Text       = "Connected to: $(Get-AppDCName)"
        $Script:Shell['ConnStatusLabel'].Foreground = [Windows.Media.Brushes]::LightGreen
        $Script:Shell['BtnConnect'].IsEnabled       = $false
        $Script:Shell['BtnDisconnect'].IsEnabled    = $true
        $Script:Shell['TxtDCName'].IsEnabled        = $false
        $Script:Shell['TxtUsername'].IsEnabled      = $false
        $Script:Shell['PwdPassword'].IsEnabled      = $false
        $Script:Window.Title = "$($Script:AppName) v$($Script:AppVersion)  —  $(Get-AppDCName)"
        Set-ShellStatus -Text "Connected to $(Get-AppDCName) — ready"
    } else {
        $Script:Shell['ConnStatusDot'].Fill         = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xE7,0x4C,0x3C)
        $Script:Shell['ConnStatusLabel'].Text       = "Not Connected"
        $Script:Shell['ConnStatusLabel'].Foreground = [Windows.Media.Brushes]::Gray
        $Script:Shell['BtnConnect'].IsEnabled       = $true
        $Script:Shell['BtnConnect'].Content         = "🔌 Connect"
        $Script:Shell['BtnDisconnect'].IsEnabled    = $false
        $Script:Shell['TxtDCName'].IsEnabled        = $true
        $Script:Shell['TxtUsername'].IsEnabled      = $true
        $Script:Shell['PwdPassword'].IsEnabled      = $true
        $Script:Window.Title = "$($Script:AppName) v$($Script:AppVersion)"
        Set-ShellStatus -Text "Ready — connect to a domain controller to begin"
    }
}

function Set-ShellMessage {
    param([string]$Text, [bool]$IsError = $false)
    $Script:Shell['TxtConnMessage'].Text = $Text
    if ($IsError) {
        $Script:Shell['TxtConnMessage'].Foreground = [Windows.Media.SolidColorBrush][Windows.Media.Color]::FromRgb(0xE7,0x4C,0x3C)
    } else {
        $Script:Shell['TxtConnMessage'].Foreground = [Windows.Media.Brushes]::LightGreen
    }
}

function Set-ShellStatus {
    param([string]$Text)
    if ($Script:Shell.ContainsKey('TxtStatusBar')) {
        $Script:Shell['TxtStatusBar'].Text = $Text
    }
}

function Set-ShellProgress {
    param([int]$Value)   # -1 = hide
    if (-not $Script:Shell.ContainsKey('ProgressBar')) { return }
    if ($Value -lt 0) {
        $Script:Shell['ProgressBar'].Visibility = 'Collapsed'
    } else {
        $Script:Shell['ProgressBar'].Visibility = 'Visible'
        $Script:Shell['ProgressBar'].Value = [Math]::Min($Value, 100)
    }
}