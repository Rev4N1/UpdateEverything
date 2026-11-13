#Requires -Version 5.1
<#
.SYNOPSIS
    Updates user apps, machine apps, Microsoft Store apps and Windows (including
    drivers) without installing any PowerShell module.

.DESCRIPTION
    The script self-elevates, but per-user software is never updated from the
    elevated (possibly different) admin account:

    - User apps + Store apps: for every user logged on to the machine, a temporary
      scheduled task runs this same script as THAT user with a normal (non-elevated)
      token and runs "winget upgrade --scope user". Store apps are per-user MSIX
      packages, so they are covered here too (msstore source). The output of each
      user run is shown live in the elevated window.
    - Machine apps: "winget upgrade --scope machine" from the elevated session.
    - Microsoft Store: additionally triggers the Store's own update scan, which
      catches Store apps winget does not know about (installs in the background).
    - Windows Update: Windows Update Agent COM API, including drivers.

    winget in the elevated account: if App Installer was never registered for that
    profile, it is registered; if that fails, winget.exe is run directly from its
    install folder with its VC++ runtime added to PATH.

    No logs are written. A temporary folder in %SystemRoot%\Temp is used only to
    hand the script to the user tasks and stream their output back, and it is
    deleted at the end.

.PARAMETER Exclude
    Windows Update titles to skip (matches any part of the title).

.PARAMETER SkipUserApps
    Do not update per-user apps / Store apps of the logged-on users.

.PARAMETER SkipMachineApps
    Do not run the machine-wide winget upgrade.

.PARAMETER SkipStore
    Do not trigger the Store's own update scan.

.PARAMETER SkipWindowsUpdate
    Do not run Windows Update.

.PARAMETER SkipDrivers
    Windows Update: software updates only, no drivers. Drivers are also skipped
    automatically when a policy or the Device installation settings turn off
    drivers from Windows Update.

.PARAMETER TargetRelease
    Windows Update: temporarily target this feature release (e.g. '26H2') so the
    feature update is offered even when it is optional or a policy pins an older
    release. The previous TargetReleaseVersion* policy values are restored at the
    end (also on errors or Ctrl+C). Note: a management agent may re-apply its pin.

.PARAMETER IncludeUnknown
    Also let winget upgrade packages whose installed version it cannot determine.

.PARAMETER UserTimeoutMinutes
    Maximum time for each user's app update run (default 90).

.PARAMETER Reboot
    Restart automatically (60 s countdown) if any update needs it.

.PARAMETER NoPause
    Do not wait for Enter at the end of the elevated window.

.EXAMPLE
    .\UpdateEverything.ps1

.EXAMPLE
    .\UpdateEverything.ps1 -Exclude 'Silverlight','Apple, Inc.' -Reboot

.EXAMPLE
    .\UpdateEverything.ps1 -TargetRelease 26H2 -SkipUserApps -SkipMachineApps -SkipStore

.NOTES
    Exit codes: 0 = everything succeeded, 1 = something failed or was partial,
    3010 = succeeded but a reboot is pending (and -Reboot was not given).
#>
[CmdletBinding()]
param(
    [string[]]$Exclude = @(),
    [switch]$SkipUserApps,
    [switch]$SkipMachineApps,
    [switch]$SkipStore,
    [switch]$SkipWindowsUpdate,
    [switch]$SkipDrivers,
    [ValidatePattern('^(\d{2}H\d)?$')][string]$TargetRelease,
    [switch]$IncludeUnknown,
    [int]$UserTimeoutMinutes = 90,
    [switch]$Reboot,
    [switch]$NoPause,
    [Parameter(DontShow)][switch]$Elevated,
    [Parameter(DontShow)][switch]$UserPhase,
    [Parameter(DontShow)][string]$UserLog
)

$ErrorActionPreference = 'Continue'
$script:PhaseLog       = $null
$script:RebootRequired = $false
$script:Summary        = New-Object System.Collections.Generic.List[object]
$script:WuCodes        = @{ 0 = 'Not started'; 1 = 'In progress'; 2 = 'Succeeded'
                            3 = 'Succeeded with errors'; 4 = 'Failed'; 5 = 'Aborted' }
$script:WinPs          = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$script:SourceUrl      = 'https://raw.githubusercontent.com/Rev4N1/UpdateEverything/main/UpdateEverything.ps1'

# ===========================================================================
# Output helpers
# ===========================================================================
function Write-Out {
    param([string]$Text, [string]$Color = 'Gray')
    Write-Host $Text -ForegroundColor $Color
    if ($script:PhaseLog) {
        try { Add-Content -Path $script:PhaseLog -Value $Text -Encoding UTF8 -ErrorAction Stop } catch { }
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error', 'Section')][string]$Level = 'Info'
    )
    if ($Level -eq 'Section') {
        Write-Out ''
        Write-Out "==== $Message ====" 'Cyan'
        return
    }
    $colors = @{ Info = 'Gray'; Ok = 'Green'; Warn = 'Yellow'; Error = 'Red' }
    Write-Out ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message) $colors[$Level]
}

function Add-Summary {
    param([string]$Area, [string]$Status, [string]$Detail)
    $script:Summary.Add([pscustomobject]@{ Area = $Area; Status = $Status; Detail = $Detail })
}

# ===========================================================================
# winget helpers (used by both the user phase and the elevated phase)
# ===========================================================================
function Resolve-Winget {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    # App Installer is on the machine but not registered for this profile
    # (typical for an admin account that never used it, or a brand-new profile)
    Write-Log "winget is not registered for '$env:USERNAME'. Registering App Installer for this profile..." Warn
    try {
        Add-AppxPackage -RegisterByFamilyName -MainPackage 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe' -ErrorAction Stop
    } catch {
        Write-Log "Registration failed: $($_.Exception.Message)" Warn
    }

    $alias = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
    for ($i = 0; $i -lt 15; $i++) {
        if (Test-Path $alias) { return $alias }
        Start-Sleep -Seconds 1
    }

    # Fallback (needs admin): run winget.exe straight from the package folder
    Write-Log 'Falling back to running winget.exe from its install folder...' Warn
    $pkg = Get-AppxPackage -AllUsers -Name 'Microsoft.DesktopAppInstaller' -ErrorAction SilentlyContinue |
        Where-Object { $_.InstallLocation -and (Test-Path (Join-Path $_.InstallLocation 'winget.exe')) } |
        Sort-Object { [version]$_.Version } -Descending |
        Select-Object -First 1
    if (-not $pkg) { return $null }

    # Outside the package context winget.exe cannot find its VC++ runtime by itself
    $vclibs = Get-AppxPackage -AllUsers -Name 'Microsoft.VCLibs.140.00.UWPDesktop' -ErrorAction SilentlyContinue |
        Where-Object { $_.Architecture -eq $pkg.Architecture } |
        Sort-Object { [version]$_.Version } -Descending |
        Select-Object -First 1
    if ($vclibs) { $env:Path = "$($vclibs.InstallLocation);$env:Path" }

    return (Join-Path $pkg.InstallLocation 'winget.exe')
}

function Invoke-Winget {
    param([string]$Path, [string[]]$Arguments)
    # Pipe through PowerShell so output reaches the logs; drop spinner/progress-bar noise
    & $Path @Arguments 2>&1 | ForEach-Object {
        $line = "$_".TrimEnd()
        if ($line -and $line -notmatch '^\s*[-\\|/]\s*$' -and $line -notmatch '[\u2588\u2592]') {
            Write-Out "    $line"
        }
    }
    return $LASTEXITCODE
}

function Invoke-WingetUpgrade {
    param([string]$Winget, [ValidateSet('user', 'machine')][string]$Scope)
    Write-Log "Using $Winget"
    Write-Log ("winget version: {0}" -f (& $Winget --version 2>$null))

    [void](Invoke-Winget $Winget @('source', 'update'))

    $wgArgs = @('upgrade', '--all', '--scope', $Scope, '--silent',
                '--accept-source-agreements', '--accept-package-agreements',
                '--disable-interactivity')
    if ($IncludeUnknown) { $wgArgs += '--include-unknown' }

    Write-Log "Upgrading all $Scope-scope packages..."
    $code = Invoke-Winget $Winget $wgArgs
    if ($code -eq 0) {
        Write-Log 'winget finished.' Ok
    } else {
        Write-Log ("winget finished with exit code 0x{0:X8} (some packages may have failed, see above)." -f $code) Warn
    }
    return $code
}

# ===========================================================================
# USER PHASE: runs as the logged-on user, non-elevated, via a scheduled task
# ===========================================================================
if ($UserPhase) {
    if ($UserLog) { $script:PhaseLog = $UserLog }
    try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

    Write-Log "User phase running as $env:USERDOMAIN\$env:USERNAME (non-elevated)"
    $winget = Resolve-Winget
    if (-not $winget) {
        Write-Log 'winget is not available for this user.' Error
        exit 2
    }
    # --scope user covers per-user installers and Store (MSIX) apps via the msstore source
    $code = Invoke-WingetUpgrade -Winget $winget -Scope user
    if ($code -eq 0) { exit 0 } else { exit 1 }
}

# ===========================================================================
# Self-elevation (always into Windows PowerShell 5.1, the Appx/task cmdlets need it)
# ===========================================================================
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin   = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Run from memory (irm ... | iex, or [scriptblock]::Create): there is no file to
# elevate or hand to the user tasks, and "exit" would close the caller's window.
# Save the script to %TEMP%, run it from there, then clean up and "return".
$fromMemory = -not $PSCommandPath

if ($fromMemory -or -not $isAdmin -or $PSVersionTable.PSEdition -eq 'Core') {
    if ($Elevated -and -not $isAdmin) {
        Write-Error 'Elevation did not produce an administrator session. Aborting.'
        exit 1
    }

    $launchPath = $PSCommandPath
    if ($fromMemory) {
        $launchPath = Join-Path $env:TEMP 'UpdateEverything.ps1'
        $self = "$($MyInvocation.MyCommand.ScriptBlock)"
        try {
            # [scriptblock]::Create keeps the full text; under iex it is not available
            if ($self -notmatch 'function Invoke-UserPhase') { $self = Invoke-RestMethod $script:SourceUrl -UseBasicParsing }
            Set-Content -Path $launchPath -Value $self -Encoding UTF8 -ErrorAction Stop
        } catch {
            Write-Error "Could not save the script to ${launchPath}: $($_.Exception.Message)"
            return
        }
    }

    $quote = { param($s) "'" + ($s -replace "'", "''") + "'" }
    # Already an admin in Windows PowerShell (only possible from memory): run in this window
    $inline = $fromMemory -and $isAdmin -and $PSVersionTable.PSEdition -ne 'Core'
    $parts = @('&', (& $quote $launchPath))
    if (-not $inline) { $parts += '-Elevated' }
    foreach ($kv in $PSBoundParameters.GetEnumerator()) {
        if ($kv.Key -in 'Elevated', 'UserPhase', 'UserLog') { continue }
        if ($kv.Value -is [switch]) {
            if ($kv.Value.IsPresent) { $parts += "-$($kv.Key)" }
        } else {
            $parts += "-$($kv.Key)"
            $parts += (@($kv.Value) | ForEach-Object { & $quote "$_" }) -join ','
        }
    }

    $parts += '; exit $LASTEXITCODE'
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($parts -join ' '))
    $code = 1
    try {
        if ($inline) {
            & $script:WinPs -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded
            $code = $LASTEXITCODE
        } else {
            # Wait so the caller gets the elevated run's exit code
            $proc = Start-Process -FilePath $script:WinPs -Verb RunAs -Wait -PassThru `
                -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded"
            $code = $proc.ExitCode
        }
    } catch {
        Write-Error "Elevation was cancelled or failed: $($_.Exception.Message)"
    }

    if ($fromMemory) {
        Remove-Item -Path $launchPath -Force -ErrorAction SilentlyContinue
        $global:LASTEXITCODE = $code
        return
    }
    exit $code
}

# ===========================================================================
# ELEVATED PHASE
# ===========================================================================
function Get-InteractiveUsers {
    # Everyone with a desktop session (console, RDP, fast user switching)
    $users = foreach ($p in (Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue)) {
        $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction SilentlyContinue
        if ($o -and $o.ReturnValue -eq 0 -and $o.User) { "$($o.Domain)\$($o.User)" }
    }
    $users | Sort-Object -Unique
}

function Show-NewLines {
    param([string]$Path, [ref]$Position, [string]$Prefix)
    if (-not (Test-Path $Path)) { return }
    $lines = @(Get-Content -Path $Path -Encoding UTF8 -ErrorAction SilentlyContinue)
    for ($i = $Position.Value; $i -lt $lines.Count; $i++) { Write-Host "$Prefix$($lines[$i])" }
    $Position.Value = $lines.Count
}

function Invoke-UserPhase {
    param([string]$User)

    $safe     = $User -replace '[\\/:*?"<>|]', '_'
    $taskPath = '\UpdateAll\'
    $taskName = "UserApps-$safe"
    $userLog  = Join-Path $script:WorkDir "User_$safe.out"
    New-Item -ItemType File -Path $userLog -Force | Out-Null
    & icacls.exe $userLog /grant "${User}:(M)" | Out-Null

    Write-Log "Updating apps for $User (runs as that user, not elevated)..."

    $taskArgs = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script:ScriptCopy`" -UserPhase -UserLog `"$userLog`""
    if ($IncludeUnknown) { $taskArgs += ' -IncludeUnknown' }

    try {
        $action   = New-ScheduledTaskAction -Execute $script:WinPs -Argument $taskArgs
        $taskUser = New-ScheduledTaskPrincipal -UserId $User -LogonType Interactive -RunLevel Limited
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                        -ExecutionTimeLimit (New-TimeSpan -Minutes $UserTimeoutMinutes)
        Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action `
            -Principal $taskUser -Settings $settings -Force -ErrorAction Stop | Out-Null
        Start-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction Stop
    } catch {
        Write-Log "Could not start the user task for ${User}: $($_.Exception.Message)" Error
        Add-Summary "User apps ($User)" 'Failed' 'Task could not start'
        return
    }

    # 267009 = task is running, 267011 = task has not run yet
    $pos          = 0
    $started      = $false
    $startLimit   = (Get-Date).AddSeconds(60)
    $deadline     = (Get-Date).AddMinutes($UserTimeoutMinutes + 1)
    $timedOut     = $false
    do {
        Start-Sleep -Seconds 2
        Show-NewLines -Path $userLog -Position ([ref]$pos) -Prefix "  [$User] "
        $state = (Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath).State
        $info  = Get-ScheduledTaskInfo -TaskName $taskName -TaskPath $taskPath
        if ($state -eq 'Running') { $started = $true }
        $finished = ($state -ne 'Running') -and ($info.LastTaskResult -notin 267009, 267011)
        if (-not $started -and -not $finished -and (Get-Date) -gt $startLimit) { $timedOut = $true }
        if ((Get-Date) -gt $deadline) { $timedOut = $true }
    } until ($finished -or $timedOut)

    Start-Sleep -Seconds 1
    Show-NewLines -Path $userLog -Position ([ref]$pos) -Prefix "  [$User] "

    if ($timedOut) {
        Stop-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue
        Write-Log "User run for $User did not finish in time and was stopped." Error
        Add-Summary "User apps ($User)" 'Failed' 'Timed out / did not start'
    } else {
        switch ($info.LastTaskResult) {
            0       { Add-Summary "User apps ($User)" 'OK'      'All upgrades applied' }
            1       { Add-Summary "User apps ($User)" 'Partial' 'Some packages failed' }
            2       { Add-Summary "User apps ($User)" 'Failed'  'winget not available for user' }
            default { Add-Summary "User apps ($User)" 'Failed'  ('Task result 0x{0:X8}' -f $info.LastTaskResult) }
        }
    }

    Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false -ErrorAction SilentlyContinue
}

function Update-UserApps {
    Write-Log 'User apps and Store apps (per user, non-elevated)' Section
    $users = @(Get-InteractiveUsers)
    if ($users.Count -eq 0) {
        Write-Log 'No user is logged on, per-user apps skipped.' Warn
        Add-Summary 'User apps' 'Skipped' 'Nobody logged on'
        return
    }
    Write-Log ("Logged-on users: {0}" -f ($users -join ', '))
    foreach ($u in $users) { Invoke-UserPhase -User $u }
}

function Update-MachineApps {
    Write-Log 'Machine-wide apps (winget, elevated)' Section
    $winget = Resolve-Winget
    if (-not $winget) {
        Write-Log 'App Installer (winget) is not installed on this machine. Install "App Installer" from the Microsoft Store and run again.' Error
        Add-Summary 'Machine apps' 'Skipped' 'App Installer not installed'
        return
    }
    $code = Invoke-WingetUpgrade -Winget $winget -Scope machine
    if ($code -eq 0) { Add-Summary 'Machine apps' 'OK' 'All upgrades applied' }
    else             { Add-Summary 'Machine apps' 'Partial' ('Exit code 0x{0:X8}' -f $code) }
}

function Update-StoreApps {
    Write-Log 'Microsoft Store update scan' Section
    try {
        $r = Get-CimInstance -Namespace 'root\cimv2\mdm\dmmap' `
                             -ClassName 'MDM_EnterpriseModernAppManagement_AppManagement01' -ErrorAction Stop |
             Invoke-CimMethod -MethodName UpdateScanMethod -ErrorAction Stop
        if ($r.ReturnValue -eq 0) {
            Write-Log 'Store update scan triggered. Remaining Store apps will update in the background.' Ok
            Add-Summary 'Store scan' 'OK' 'Triggered (installs in background)'
        } else {
            Write-Log "Store scan returned $($r.ReturnValue)." Warn
            Add-Summary 'Store scan' 'Warning' "ReturnValue $($r.ReturnValue)"
        }
    } catch {
        Write-Log "Could not trigger the Store scan: $($_.Exception.Message)" Warn
        Add-Summary 'Store scan' 'Failed' $_.Exception.Message
    }
}

# Settings that say "no drivers from Windows Update". Windows' own update
# honours them, but a search through the Update Agent API does not reliably, so
# without this check the script would install drivers a policy blocks.
function Get-DriverBlockReason {
    $checks = @(
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
           Name = 'ExcludeWUDriversInQualityUpdate'; Value = 1
           Why  = 'policy "Do not include drivers with Windows Updates"' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update'
           Name = 'ExcludeWUDriversInQualityUpdate'; Value = 1
           Why  = 'MDM policy ExcludeWUDriversInQualityUpdate' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching'
           Name = 'DontSearchWindowsUpdate'; Value = 1
           Why  = 'policy "Turn off Windows Update device driver searching"' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching'
           Name = 'SearchOrderConfig'; Value = 0
           Why  = 'policy "Specify search order for device driver source locations" (do not search Windows Update)' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching'
           Name = 'SearchOrderConfig'; Value = 0
           Why  = 'Device installation settings: automatic driver downloads are off' }
    )
    foreach ($c in $checks) {
        $v = (Get-ItemProperty -Path $c.Path -Name $c.Name -ErrorAction SilentlyContinue).($c.Name)
        if ($null -ne $v -and [int]$v -eq $c.Value) { return $c.Why }
    }
    return $null
}

function Update-WindowsUpdates {
    Write-Log 'Windows Update' Section
    $noDrivers = [bool]$SkipDrivers
    if (-not $noDrivers) {
        $why = Get-DriverBlockReason
        if ($why) {
            Write-Log "Drivers skipped: $why." Warn
            $noDrivers = $true
        }
    }
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $session.ClientApplicationID = 'UpdateEverything'
        $searcher = $session.CreateUpdateSearcher()
        $criteria = "IsInstalled=0 and IsHidden=0"
        if ($noDrivers) { $criteria += " and Type='Software'" }
        Write-Log 'Searching for updates (this can take a few minutes)...'
        $result = $searcher.Search($criteria)
    } catch {
        Write-Log ("Search failed: {0} (0x{1:X8})" -f $_.Exception.Message, $_.Exception.HResult) Error
        Add-Summary 'Windows Update' 'Failed' 'Search error'
        return
    }

    $toDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($u in $result.Updates) {
        $title = $u.Title
        # Literal, case-insensitive match (no wildcard surprises with [ ] * ?)
        if ($Exclude | Where-Object { $_ -and $title.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0 }) {
            Write-Log "Skipped:  $title"
            continue
        }
        if (-not $u.EulaAccepted) { try { $u.AcceptEula() } catch { } }
        [void]$toDownload.Add($u)
        Write-Log "Queued:   $title"
    }

    if ($toDownload.Count -eq 0) {
        Write-Log 'No pending updates.' Ok
        Add-Summary 'Windows Update' 'OK' 'Nothing to install'
        return
    }

    Write-Log "Downloading $($toDownload.Count) update(s)..."
    try {
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $toDownload
        [void]$downloader.Download()
    } catch {
        Write-Log ("Download error: {0} (0x{1:X8})" -f $_.Exception.Message, $_.Exception.HResult) Warn
    }

    $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($u in $toDownload) {
        if ($u.IsDownloaded) { [void]$toInstall.Add($u) }
        else { Write-Log "Not downloaded: $($u.Title)" Warn }
    }
    if ($toInstall.Count -eq 0) {
        Write-Log 'Nothing could be downloaded.' Error
        Add-Summary 'Windows Update' 'Failed' 'Download failed'
        return
    }

    Write-Log "Installing $($toInstall.Count) update(s)..."
    try {
        $installer = $session.CreateUpdateInstaller()
        $installer.ForceQuiet = $true
        $installer.AllowSourcePrompts = $false
        $installer.Updates = $toInstall
        $res = $installer.Install()
    } catch {
        Write-Log ("Install error: {0} (0x{1:X8})" -f $_.Exception.Message, $_.Exception.HResult) Error
        Add-Summary 'Windows Update' 'Failed' 'Install error'
        return
    }

    $ok = 0; $fail = 0
    for ($i = 0; $i -lt $toInstall.Count; $i++) {
        $r     = $res.GetUpdateResult($i)
        $code  = [int]$r.ResultCode
        $title = $toInstall.Item($i).Title
        if ($code -eq 2) {
            $ok++
            Write-Log ("{0,-22} {1}" -f $script:WuCodes[$code], $title) Ok
        } else {
            $fail++
            Write-Log ("{0,-22} {1} (0x{2:X8})" -f $script:WuCodes[$code], $title, $r.HResult) Error
        }
    }
    if ($res.RebootRequired) { $script:RebootRequired = $true }

    $status = if ($fail -gt 0) { 'Partial' } else { 'OK' }
    Add-Summary 'Windows Update' $status "$ok installed, $fail failed"
}

# ---------------------------------------------------------------------------
# Feature release targeting (temporary, restored afterwards)
# ---------------------------------------------------------------------------
$script:WuPolicyKey   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$script:WuPolicyNames = 'TargetReleaseVersion', 'TargetReleaseVersionInfo', 'ProductVersion'
$script:WuPolicySaved = $null

function Set-FeatureTarget {
    param([string]$Release)
    $saved = @{ KeyExisted = (Test-Path $script:WuPolicyKey); Values = @{} }
    if ($saved.KeyExisted) {
        $key = Get-Item $script:WuPolicyKey
        foreach ($n in $script:WuPolicyNames) {
            if ($key.GetValueNames() -contains $n) {
                $saved.Values[$n] = @{ Value = $key.GetValue($n); Kind = $key.GetValueKind($n) }
            }
        }
    } else {
        New-Item -Path $script:WuPolicyKey -Force | Out-Null
    }
    $script:WuPolicySaved = $saved

    $build   = [Environment]::OSVersion.Version.Build
    $product = if ($build -ge 22000) { 'Windows 11' } else { 'Windows 10' }
    New-ItemProperty -Path $script:WuPolicyKey -Name 'TargetReleaseVersion'     -Value 1        -PropertyType DWord  -Force | Out-Null
    New-ItemProperty -Path $script:WuPolicyKey -Name 'TargetReleaseVersionInfo' -Value $Release -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $script:WuPolicyKey -Name 'ProductVersion'           -Value $product -PropertyType String -Force | Out-Null
    Write-Log "Temporarily targeting feature release $product $Release."
}

function Restore-FeatureTarget {
    $saved = $script:WuPolicySaved
    if (-not $saved) { return }
    $script:WuPolicySaved = $null
    try {
        if (-not $saved.KeyExisted) {
            Remove-Item -Path $script:WuPolicyKey -Recurse -Force -ErrorAction Stop
        } else {
            foreach ($n in $script:WuPolicyNames) {
                if ($saved.Values.ContainsKey($n)) {
                    $v = $saved.Values[$n]
                    New-ItemProperty -Path $script:WuPolicyKey -Name $n -Value $v.Value -PropertyType $v.Kind -Force | Out-Null
                } else {
                    Remove-ItemProperty -Path $script:WuPolicyKey -Name $n -ErrorAction SilentlyContinue
                }
            }
        }
        Write-Log 'Feature release targeting restored to its previous state.'
    } catch {
        Write-Log "Could not restore the feature release policy: $($_.Exception.Message)" Error
    }
}

# ---------------------------------------------------------------------------
# Leftovers from an interrupted earlier run
# ---------------------------------------------------------------------------
function Remove-TaskFolder {
    foreach ($t in @(Get-ScheduledTask -TaskPath '\UpdateAll\' -ErrorAction SilentlyContinue)) {
        $t | Stop-ScheduledTask -ErrorAction SilentlyContinue
        $t | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
    }
    try {
        $svc = New-Object -ComObject Schedule.Service
        $svc.Connect()
        $svc.GetFolder('\').DeleteFolder('UpdateAll', 0)
    } catch { }
}

function Clear-Leftovers {
    Remove-TaskFolder
    Get-ChildItem -Path (Join-Path $env:SystemRoot 'Temp') -Directory -Filter 'UpdateAll_*' -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# Main (elevated)
# ---------------------------------------------------------------------------
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

$script:WorkDir    = $null
$script:ScriptCopy = $null
$exitCode          = 0

# One run at a time (a second run would clean up the first one's tasks and files)
$mutex = New-Object System.Threading.Mutex($false, 'Global\UpdateEverything')
try { $owned = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
if (-not $owned) {
    Write-Log 'Another UpdateEverything run is already in progress. Exiting.' Error
    if ($Elevated -and -not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
    exit 1
}

try {
    Clear-Leftovers

    # Temporary folder to hand the script to the user tasks (the original may sit in a
    # profile other users cannot read) and to stream their output back. Deleted at the end.
    if (-not $SkipUserApps) {
        $script:WorkDir    = Join-Path $env:SystemRoot ('Temp\UpdateAll_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $script:ScriptCopy = Join-Path $script:WorkDir 'UpdateEverything.ps1'
        New-Item -ItemType Directory -Path $script:WorkDir -Force | Out-Null
        & icacls.exe $script:WorkDir /grant '*S-1-5-32-545:(OI)(CI)RX' | Out-Null
        Copy-Item -Path $PSCommandPath -Destination $script:ScriptCopy -Force
    }

    Write-Log "UpdateEverything started as $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME"

    # User apps first: if the Store/winget updates App Installer itself, it happens
    # before the elevated winget run starts, not in the middle of it.
    if (-not $SkipUserApps)    { Update-UserApps }
    if (-not $SkipMachineApps) { Update-MachineApps }
    if (-not $SkipStore)       { Update-StoreApps }
    if (-not $SkipWindowsUpdate) {
        if ($TargetRelease) { Set-FeatureTarget -Release $TargetRelease }
        try { Update-WindowsUpdates } finally { Restore-FeatureTarget }
    }

    try {
        if ((New-Object -ComObject Microsoft.Update.SystemInfo).RebootRequired) { $script:RebootRequired = $true }
    } catch { }
}
catch {
    Write-Log "Unexpected error: $($_.Exception.Message)" Error
    Add-Summary 'Script' 'Failed' $_.Exception.Message
}
finally {
    # Runs on errors and Ctrl+C too
    Restore-FeatureTarget
    if (-not $SkipUserApps) { Remove-TaskFolder }
    if ($script:WorkDir -and (Test-Path $script:WorkDir)) {
        Remove-Item -Path $script:WorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}

Write-Log 'Summary' Section
$script:Summary | Format-Table -AutoSize | Out-String | Write-Host

if ($script:Summary | Where-Object { $_.Status -in 'Failed', 'Partial' }) { $exitCode = 1 }

if ($script:RebootRequired) {
    if ($Reboot) {
        Write-Log 'Reboot required. Restarting in 60 seconds (run "shutdown /a" to cancel)...' Warn
        shutdown.exe /r /t 60 /c "UpdateEverything: restarting to finish installing updates"
        exit $exitCode
    }
    Write-Log 'A reboot is required to finish installing updates.' Warn
    if ($exitCode -eq 0) { $exitCode = 3010 }
}

if ($Elevated -and -not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
exit $exitCode