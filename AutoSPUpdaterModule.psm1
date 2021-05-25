#Author
#https://github.com/brianlala/AutoSPUpdater

#region Install Updates
function InstallUpdatesFromPatchPath
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()]
        [string]$patchPath,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()]
        [string]$spVer
    )
    $spVer,$spYear = Get-SPYear
    Write-Host -ForegroundColor White " - Looking for SharePoint updates to install in $patchPath..."
    # Result codes below are from http://technet.microsoft.com/en-us/library/cc179058(v=office.14).aspx
    $oPatchInstallResultCodes = @{"17301" = "Error: General Detection error";
                                  "17302" = "Error: Applying patch";
                                  "17303" = "Error: Extracting file";
                                  "17021" = "Error: Creating temp folder";
                                  "17022" = "Success: Reboot flag set";
                                  "17023" = "Error: User cancelled installation";
                                  "17024" = "Error: Creating folder failed";
                                  "17025" = "Patch already installed";
                                  "17026" = "Patch already installed to admin installation";
                                  "17027" = "Installation source requires full file update";
                                  "17028" = "No product installed for contained patch";
                                  "17029" = "Patch failed to install";
                                  "17030" = "Detection: Invalid CIF format";
                                  "17031" = "Detection: Invalid baseline";
                                  "17034" = "Error: Required patch does not apply to the machine";
                                  "17038" = "You do not have sufficient privileges to complete this installation for all users of the machine. Log on as administrator and then retry this installation";
                                  "17044" = "Installer was unable to run detection for this package"}

    # Get all CUs and PUs
    $updatesToInstall = Get-ChildItem -Path "$patchPath" -Include office2010*.exe,ubersrv*.exe,ubersts*.exe,*pjsrv*.exe,sharepointsp2013*.exe,coreserver201*.exe,sts201*.exe,wssloc201*.exe,svrproofloc201*.exe,oserver*.exe,wac*.exe,oslpksp*.exe -Recurse -ErrorAction SilentlyContinue | Sort-Object -Descending
    # Look for Server Update installers
    if ($updatesToInstall)
    {
        Write-Host -ForegroundColor White " - Starting local install..."
        <#
        # Display warning about missing March 2013 PU only if we are actually installing SP2013 and SP1 isn't already installed and the SP1 installer isn't found
        $sp2013SP1 = Get-ChildItem -Path "$bits\$spYear\Updates" -Name -Include "officeserversp2013-kb2880552-fullfile-x64-en-us.exe" -Recurse -ErrorAction SilentlyContinue
        if ($spYear -eq "2013" -and !($sp2013SP1 -or (CheckFor2013SP1)) -and !$marchPublicUpdate)
        {
            Write-Host -ForegroundColor Yellow "  - Note: the March 2013 PU package wasn't found in ..\$spYear\Updates; it may need to be installed first if it wasn't slipstreamed."
        }
        #>
        # Now attempt to install any other CUs found in the \Updates folder
        Write-Host -ForegroundColor White "  - Installing SharePoint Updates on " -NoNewline
        Write-Host -ForegroundColor Black -BackgroundColor Green "$env:COMPUTERNAME"
        ForEach ($updateToInstall in $updatesToInstall)
        {
            # Get the file name only, in case $updateToInstall includes part of a path (e.g. is in a subfolder)
            $splitUpdate = Split-Path -Path $updateToInstall -Leaf
            Write-Verbose -Message "Running `"Start-Process -FilePath `"$updateToInstall`" -ArgumentList `"/passive /norestart`" -LoadUserProfile`""
            Write-Host -ForegroundColor Cyan "   - Installing $splitUpdate from `"$($updateToInstall.Directory.Name)`"..." -NoNewline
            $startTime = Get-Date
            Start-Process -FilePath "$updateToInstall" -ArgumentList "/passive /norestart" -LoadUserProfile
            Show-Progress -Process $($splitUpdate -replace ".exe", "") -Color Cyan -Interval 5
            $delta,$null = (New-TimeSpan -Start $startTime -End (Get-Date)).ToString() -split "\."
            $oPatchInstallLog = Get-ChildItem -Path (Get-Item $env:TEMP).FullName | Where-Object {$_.Name -like "opatchinstall*.log"} | Sort-Object -Descending -Property "LastWriteTime" | Select-Object -first 1
            # Get install result from log
            $oPatchInstallResultMessage = $oPatchInstallLog | Select-String -SimpleMatch -Pattern "OPatchInstall: Property 'SYS.PROC.RESULT' value" | Select-Object -Last 1
            If (!($oPatchInstallResultMessage -like "*value '0'*")) # Anything other than 0 means unsuccessful but that's not necessarily a bad thing
            {
                $null,$oPatchInstallResultCode = $oPatchInstallResultMessage.Line -split "OPatchInstall: Property 'SYS.PROC.RESULT' value '"
                $oPatchInstallResultCode = $oPatchInstallResultCode.TrimEnd("'")
                # OPatchInstall: Property 'SYS.PROC.RESULT' value '17028' means the patch was not needed or installed product was newer
                if ($oPatchInstallResultCode -eq "17028") {Write-Host -ForegroundColor Yellow "   - Patch not required; installed product is same or newer."}
                elseif ($oPatchInstallResultCode -eq "17031")
                {
                    Write-Warning "Error 17031: Detection: Invalid baseline"
                    Write-Warning "A baseline patch (e.g. March 2013 PU for SP2013, SP1 for SP2010) is missing!"
                    Write-Host -ForegroundColor Yellow "   - Either slipstream the missing patch first, or include the patch package in the ..\$spYear\Updates folder."
                    Pause "continue"
                }
                else
                {
                    Write-Host -ForegroundColor Yellow "   - $($oPatchInstallResultCodes.$oPatchInstallResultCode)"
                    if ($oPatchInstallResultCode -ne "17025") # i.e. "Patch already installed"
                    {
                        Write-Host -ForegroundColor Yellow "   - Please log on to this server ($env:COMPUTERNAME) now, and install the update manually."
                        Pause "continue once the update has been successfully installed manually" "y"
                    }
                }
            }
            Write-Host -ForegroundColor White "   - $splitUpdate install completed in $delta."
        }
        Write-Host -ForegroundColor White "  - Update installation complete."
    }
    Write-Host -ForegroundColor White " - Finished installing SharePoint updates on " -NoNewline
    Write-Host -ForegroundColor Black -BackgroundColor Green "$env:COMPUTERNAME"
    WriteLine

}
#endregion

#region Remote Install
function Install-Remote
{
    [CmdletBinding()]
    param
    (
        [bool]$skipParallelInstall = $false,
        [array]$remoteFarmServers,
        [System.Management.Automation.PSCredential]$credential,
        [string]$launchPath,
        [string]$patchPath
    )
    if ($VerbosePreference -eq "Continue")
    {
        $verboseParameter = @{Verbose = $true}
        $verboseSwitch = "-Verbose"
    }
    else
    {
        $verboseParameter = @{}
        $verboseSwitch = ""
    }

    if (!$RemoteStartDate) {$RemoteStartDate = Get-Date}
    if ($null -eq $spVer)
    {
        [string]$spVer = (Get-SPFarm).BuildVersion.Major
        if (!$?)
        {
            Start-Sleep 10
            throw "Could not determine version of farm."
        }
    }
    Write-Host -ForegroundColor White " - Starting remote installs..."
    Enable-CredSSP $remoteFarmServers
    foreach ($server in $remoteFarmServers)
    {
        if (!($skipParallelInstall)) # Launch each farm server install simultaneously
        {
            # Add the -Version 2 switch in case we are installing SP2010 on Windows Server 2012 or 2012 R2
            if (((Get-CimInstance -ClassName Win32_OperatingSystem).Version -like "6.2*" -or (Get-CimInstance -ClassName Win32_OperatingSystem).Version -like "6.3*") -and ($spVer -eq "14"))
            {
                $versionSwitch = "-Version 2"
            }
            else {$versionSwitch = ""}
            Start-Process -FilePath "$PSHOME\powershell.exe" -ArgumentList "$versionSwitch `
                                                                            -ExecutionPolicy Bypass Invoke-Command -ScriptBlock {
                                                                            Import-Module -Name `"$launchPath\AutoSPUpdaterModule.psm1`" -DisableNameChecking -Global -Force `
                                                                            StartTracing -Server $server; `
                                                                            Test-ServerConnection -Server $server; `
                                                                            Enable-RemoteSession -Server $server -plainPass $(ConvertFrom-SecureString $($credential.Password)) -launchPath `"$launchPath`"; `
                                                                            Backup-IISConfiguration; `
                                                                            Start-RemoteUpdate -Server $server -plainPass $(ConvertFrom-SecureString $($credential.Password)) -launchPath `"$launchPath`" -patchPath `"$patchPath`" -spVer $spver $verboseSwitch; `
                                                                            Pause `"exit`"; `
                                                                            Stop-Transcript -ErrorAction SilentlyContinue}" -Verb Runas
            Start-Sleep 10
        }
        else # Launch each farm server install in sequence, one-at-a-time, or run these steps on the current $targetServer
        {
            WriteLine
            Write-Host -ForegroundColor Green " - Server: $server"
            Import-Module -Name "$launchPath\AutoSPUpdaterModule.psm1" -DisableNameChecking -Global -Force
            Test-ServerConnection -Server $server
            Enable-RemoteSession -Server $server -Password $(ConvertFrom-SecureString $($credential.Password)) -launchPath $launchPath; `
            InstallUpdatesFromPatchPath `
        }
    }
}
function Start-RemoteUpdate
{
    [CmdletBinding()]
    param
    (
        [String]$server,
        [String]$plainPass,
        [String]$launchPath,
        [String]$patchPath,
        [String]$spVer
    )
    if ($VerbosePreference -eq "Continue")
    {
        $verboseParameter = @{Verbose = $true}
    }
    else
    {
        $verboseParameter = @{}
    }
    If ($plainPass) {$credential = New-Object System.Management.Automation.PsCredential $env:USERDOMAIN\$env:USERNAME,$(ConvertTo-SecureString $plainPass)}
    If (!$credential) {$credential = $host.ui.PromptForCredential("AutoSPInstaller - Remote Install", "Re-Enter Credentials for Remote Authentication:", "$env:USERDOMAIN\$env:USERNAME", "NetBiosUserName")}
    If ($session.Name -ne "AutoSPUpdaterSession-$server")
    {
        Write-Host -ForegroundColor White " - Starting remote session to $server..."
        $session = New-PSSession -Name "AutoSPUpdaterSession-$server" -Authentication Credssp -Credential $credential -ComputerName $server
    }
    # Set some remote variables that we will need...
    Invoke-Command -ScriptBlock {param ($value) Set-Variable -Name launchPath -Value $value} -ArgumentList $launchPath -Session $session
    Invoke-Command -ScriptBlock {param ($value) Set-Variable -Name spVer -Value $value} -ArgumentList $spVer -Session $session
    Invoke-Command -ScriptBlock {param ($value) Set-Variable -Name patchPath -Value $value} -ArgumentList $patchPath -Session $session
    Invoke-Command -ScriptBlock {param ($value) Set-Variable -Name credential -Value $value} -ArgumentList $credential -Session $session
    Invoke-Command -ScriptBlock {param ($value) Set-Variable -Name verboseParameter -Value $value} -ArgumentList $verboseParameter -Session $session
    Write-Host -ForegroundColor White " - Launching AutoSPUpdater..."
    Invoke-Command -ScriptBlock {& "$launchPath\AutoSPUpdaterLaunch.ps1" -patchPath $patchPath -remoteAuthPassword $(ConvertFrom-SecureString $($credential.Password)) @verboseParameter} -Session $session
    Write-Host -ForegroundColor White " - Removing session `"$($session.Name)...`""
    Remove-PSSession $session
}
#endregion

#region Utility Functions
function Pause($action, $key)
{
    # From http://www.microsoft.com/technet/scriptcenter/resources/pstips/jan08/pstip0118.mspx
    if ($key -eq "any" -or ([string]::IsNullOrEmpty($key)))
    {
        $actionString = " - Press any key to $action..."
        if (-not $unattended)
        {
            Write-Host -ForegroundColor White $actionString
            $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        else
        {
            Write-Verbose -Message "Skipping pause due to -unattended switch: $actionString"
        }
    }
    else
    {
        $actionString = " - Enter `"$key`" to $action"
        $continue = Read-Host -Prompt $actionString
        if ($continue -ne $key) {pause $action $key}

    }
}
function Import-SharePointPowerShell
{
    [CmdletBinding()]
    param ()
    if ($null -eq (Get-PsSnapin | Where-Object {$_.Name -eq "Microsoft.SharePoint.PowerShell"}))
    {
        Write-Host -ForegroundColor White " - (Re-)Loading SharePoint PowerShell Snapin..."
        # Added the line below to match what the SharePoint.ps1 file implements (normally called via the SharePoint Management Shell Start Menu shortcut)
        if (Confirm-LocalSession) {$Host.Runspace.ThreadOptions = "ReuseThread"}
        Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
    }
}
function Confirm-LocalSession
{
    if ($Host.Name -eq "ServerRemoteHost") {return $false}
    else {return $true}
}
function Enable-CredSSP
{
    [CmdletBinding()]
    param
    (
        [array]$remoteFarmServers
    )
    Write-Verbose -Message "Remote farm servers: $remoteFarmServers"
    foreach ($server in $remoteFarmServers)
    {
        Write-Host -ForegroundColor White " - Enabling WSManCredSSP for `"$server`""
        Enable-WSManCredSSP -Role Client -Force -DelegateComputer $server | Out-Null
        if (!$?) {Pause "exit"; throw $_}
    }
}
function Test-ServerConnection
{
    [CmdletBinding()]
    param
    (
        [string]$server
    )
    Write-Verbose -Message "Running `"Test-Connection -ComputerName $server -Count 1 -Quiet`""
    Write-Host -ForegroundColor White " - Testing connection (via Ping) to `"$server`"..." -NoNewline
    $canConnect = Test-Connection -ComputerName $server -Count 1 -Quiet
    if ($canConnect) {Write-Host -ForegroundColor Cyan -BackgroundColor Black $($canConnect.ToString() -replace "True","Success.")}
    if (!$canConnect)
    {
        Write-Host -ForegroundColor Yellow -BackgroundColor Black $($canConnect.ToString() -replace "False","Failed.")
        Write-Host -ForegroundColor Yellow " - Check that `"$server`":"
        Write-Host -ForegroundColor Yellow "  - Is online"
        Write-Host -ForegroundColor Yellow "  - Has the required Windows Firewall exceptions set (or turned off)"
        Write-Host -ForegroundColor Yellow "  - Has a valid DNS entry for $server.$($env:USERDNSDOMAIN)"
        throw "Ping connectivity test failed for `"$server`""
    }
}
function Enable-RemoteSession
{
    [CmdletBinding()]
    param
    (
        [String]$server,
        [String]$plainPass,
        [String]$launchPath
    )
    If ($plainPass) {$credential = New-Object System.Management.Automation.PsCredential $env:USERDOMAIN\$env:USERNAME,$(ConvertTo-SecureString $plainPass)}
    If (!$credential) {$credential = $host.ui.PromptForCredential("AutoSPUpdater - Remote Install", "Re-Enter Credentials for Remote Authentication:", "$env:USERDOMAIN\$env:USERNAME", "NetBiosUserName")}
    $username = $credential.Username
    $password = ConvertTo-PlainText $credential.Password
    $configureTargetScript = "$launchPath\AutoSPUpdaterConfigureRemoteTarget.ps1"
    $psExec = $launchPath+"\PsExec.exe"
    If (!(Get-Item ($psExec) -ErrorAction SilentlyContinue))
    {
        Write-Host -ForegroundColor White " - PsExec.exe not found; downloading..."
        $psExecUrl = "http://live.sysinternals.com/PsExec.exe"
        Import-Module BitsTransfer | Out-Null
        Start-BitsTransfer -Source $psExecUrl -Destination $psExec -DisplayName "Downloading Sysinternals PsExec..." -Priority Foreground -Description "From $psExecUrl..." -ErrorVariable err
        If ($err) {Write-Warning "Could not download PsExec!"; Pause "exit"; break}
    }
    Write-Host -ForegroundColor White " - Updating PowerShell execution policy on `"$server`" via PsExec..."
    Start-Process -FilePath "$psExec" `
                  -ArgumentList "/acceptEula \\$server -h powershell.exe -Command `"try {Set-ExecutionPolicy Bypass -Force} catch {}; Stop-Process -Id `$PID`"" `
                  -Wait -NoNewWindow
    # Another way to exit powershell when running over PsExec from http://www.leeholmes.com/blog/2007/10/02/using-powershell-and-PsExec-to-invoke-expressions-on-remote-computers/
    # PsExec \\server cmd /c "echo . | powershell {command}"
    Write-Host -ForegroundColor White " - Enabling PowerShell remoting on `"$server`" via PsExec..."
    Write-Verbose -Message "Running '$psexec /acceptEula \\$server -u $username -p $password -h powershell.exe -Command `"$configureTargetScript`"..."
    Start-Process -FilePath "$psExec" `
                  -ArgumentList "/acceptEula \\$server -u $username -p $password -h powershell.exe -Command `"$configureTargetScript`"" `
                  -Wait -NoNewWindow
}
function StartTracing
{
    [CmdletBinding()]
    param
    (
        [string]$server
    )
    if (!$isTracing)
    {
        If ([string]::IsNullOrEmpty($logtime)) {$script:Logtime = Get-Date -Format yyyy-MM-dd_h-mm}
        If ($server) {$script:LogFile = Join-Path -Path $([Environment]::GetFolderPath("Desktop")) -ChildPath "\AutoSPUpdater-$server-$script:Logtime.log"}
        else {$script:LogFile = Join-Path -Path $([Environment]::GetFolderPath("Desktop")) -ChildPath "\AutoSPUpdater-$script:Logtime.log"}
        Start-Transcript -Path $logFile -Append -Force
        If ($?) {$global:isTracing = $true}
    }
}
function UnblockFiles ($path)
{
    # Ensure that if we're running from a UNC path, the host portion is added to the Local Intranet zone so we don't get the "Open File - Security Warning"
    If ($path -like "\\*")
    {
        WriteLine
        if (Get-Command -Name "Unblock-File" -ErrorAction SilentlyContinue)
        {
            Write-Host -ForegroundColor White " - Unblocking executable files in $path to prevent security prompts..." -NoNewline
            # Leverage the Unblock-File cmdlet, if available to prevent security warnings when working with language packs, CUs etc.
            Get-ChildItem -Path $path -Recurse | Where-Object {($_.Name -like "*.exe") -or ($_.Name -like "*.ms*") -or ($_.Name -like "*.zip") -or ($_.Name -like "*.cab")} | Unblock-File -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host -ForegroundColor White "Done."
        }
        $safeHost = ($path -split "\\")[2]
        Write-Host -ForegroundColor White " - Adding location `"$safeHost`" to local Intranet security zone to prevent security prompts..." -NoNewline
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains" -Name $safeHost -ItemType Leaf -Force | Out-Null
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$safeHost" -Name "file" -value "1" -PropertyType dword -Force | Out-Null
        Write-Host -ForegroundColor White "Done."
        WriteLine
    }
}
function WriteLine
{
    Write-Host -ForegroundColor White "--------------------------------------------------------------"
}
<#
# ===================================================================================
# Func: ConvertTo-PlainText
# Desc: Convert string to secure phrase
#       Used (for example) to get the Farm Account password into plain text as input to provision the User Profile Sync Service
#       From http://www.vistax64.com/powershell/159190-read-host-assecurestring-problem.html
# ===================================================================================
#>
function ConvertTo-PlainText( [security.securestring]$secure )
{
    $marshal = [Runtime.InteropServices.Marshal]
    $marshal::PtrToStringAuto( $marshal::SecureStringToBSTR($secure) )
}
<#
# ====================================================================================
# Func: Show-Progress
# Desc: Shows a row of dots to let us know that $process is still running
# From: Brian Lalancette, 2012
# ====================================================================================
#>
function Show-Progress ($process, $color, $interval)
{
    While (Get-Process -Name $process -ErrorAction SilentlyContinue)
    {
        Write-Host -ForegroundColor $color "." -NoNewline
        Start-Sleep $interval
    }
    Write-Host -ForegroundColor Green "Done."
}
<#
# ====================================================================================
# Func: Test-UpgradeRequired
# Desc: Returns $true if the server or farm requires an upgrade (i.e. requires PSConfig or the corresponding PowerShell commands to be run)
# ====================================================================================
#>
Function Test-UpgradeRequired
{
    if ($null -eq $spVer)
    {
        $spVer = (Get-SPFarm).BuildVersion.Major
        if (!$?)
        {
            throw "Could not determine version of farm."
        }
    }
    $setupType = (Get-Item -Path "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\$spVer.0\WSS\").GetValue("SetupType")
    If ($setupType -ne "CLEAN_INSTALL") # For example, if the value is "B2B_UPGRADE"
    {
        Return $true
    }
    Else
    {
        Return $false
    }
}
function Test-PSConfig
{
    [CmdletBinding()]
    param ()
    $PSConfigLogLocation = $((Get-SPDiagnosticConfig).LogLocation) -replace "%CommonProgramFiles%","$env:CommonProgramFiles"
    $PSConfigLog = Get-ChildItem -Path $PSConfigLogLocation | Where-Object {$_.Name -like "PSCDiagnostics*"} | Sort-Object -Descending -Property "LastWriteTime" | Select-Object -first 1
    If ($null -eq $PSConfigLog)
    {
        Write-Warning " - Could not find PSConfig log file!"
    }
    Else
    {
        # Get error(s) from log
        $PSConfigLastError = $PSConfigLog | select-string -SimpleMatch -CaseSensitive -Pattern "ERR" | Select-Object -Last 1
        return $PSConfigLastError
    }
}
function Request-SPSearchServiceApplicationStatus
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]
        [ValidateSet("Paused","Online")]
        [String]$desiredStatus
    )

    # From https://technet.microsoft.com/en-ca/library/dn745901.aspx
    <#
($ssa.IsPaused() -band 0x01) -ne 0 #A change in the number of crawl components or crawl databases is in progress.
($ssa.IsPaused() -band 0x02) -ne 0 #A backup or restore procedure is in progress.
($ssa.IsPaused() -band 0x04) -ne 0 #A backup of the Volume Shadow Copy Service (VSS) is in progress.
($ssa.IsPaused() -band 0x08) -ne 0 #One or more servers in the search topology that host query components are offline.
($ssa.IsPaused() -band 0x20) -ne 0 #One or more crawl databases in the search topology are being rebalanced.
($ssa.IsPaused() -band 0x40) -ne 0 #One or more link databases in the search topology are being rebalanced.
($ssa.IsPaused() -band 0x80) -ne 0 #An administrator has manually paused the Search service application.
($ssa.IsPaused() -band 0x100) -ne 0 #The search index is being deleted.
($ssa.IsPaused() -band 0x200) -ne 0 #The search index is being repartitioned.
#>
    [array]$farmServers = (Get-SPFarm).Servers | Where-Object {$_.Role -ne "Invalid"}
    Write-Verbose -Message "$($farmservers.Count) farm server(s) detected."

    switch ($desiredStatus)
    {
        "Paused" {$actionWord = "Pausing"; $color = "Yellow"; $action = "Pause"; $cmdlet = "Suspend-SPEnterpriseSearchServiceApplication"; $statusCheck = "((Get-SPEnterpriseSearchServiceApplication -Identity `$searchServiceApplication -ErrorAction SilentlyContinue).IsPaused() -band 0x80) -ne 0"}
        "Online" {$actionWord = "Resuming"; $color = "Green"; $action = "Resume"; $cmdlet = "Resume-SPEnterpriseSearchServiceApplication"; $statusCheck = "(Get-SPEnterpriseSearchServiceApplication -Identity `$searchServiceApplication -ErrorAction SilentlyContinue).IsPaused() -eq 0"}
    }
    if (Get-SPEnterpriseSearchServiceApplication -ErrorAction SilentlyContinue)
    {
        Write-Host -ForegroundColor White " - $actionWord Search Service Application(s)..."
        foreach ($searchServiceApplication in (Get-SPEnterpriseSearchServiceApplication))
        {
            try
            {
                $status = (Invoke-Expression -Command "$statusCheck")
                if ($null -eq $status) {throw}
                if (Invoke-Expression -Command "$statusCheck")
                {
                    Write-Host -ForegroundColor White "  - `"$($searchServiceApplication.Name)`" is already $desiredStatus."
                }
                else
                {
                    # Only pause if we are resuming, and if there are multiple farm servers
                    if ($action -eq "Resume" -and $farmServers.Count -gt 1)
                    {
                        Pause "$($action.ToLower()) `"$($searchServiceApplication.Name)`" after all installs have completed" "y"
                    }
                    Write-Host -ForegroundColor White "  - $actionWord `"$($searchServiceApplication.Name)`"; this can take several minutes..."
                    try
                    {
                        Invoke-Expression -Command "`$searchServiceApplication | $cmdlet"
                        if (!$?) {throw}
                        Invoke-Expression -Command "$statusCheck"
                        if (!$?) {throw}
                        if (Invoke-Expression -Command "$statusCheck")
                        {
                            Write-Host -ForegroundColor White "  - `"$($searchServiceApplication.Name)`" is now " -NoNewline
                            Write-Host -ForegroundColor $color "$desiredStatus"
                        }
                        else
                        {
                            throw
                        }
                    }
                    catch
                    {
                        Write-Warning "Could not $action `"$($searchServiceApplication.Name)`""
                    }
                }
            }
            catch
            {
             Write-Warning "Could not get status of `"$($searchServiceApplication.Name)`""
            }
        }
        Write-Host -ForegroundColor White " - Done $($actionWord.ToLower()) Search Service Application(s)."
    }
}
function Update-ContentDatabases
{
    [CmdletBinding()]
    param
    (
        [string]$spVer,
        [Switch]$useSqlSnapshot = $false
    )
    $upgradeContentDBScriptBlock = {
        ##$Host.UI.RawUI.WindowTitle = "-- Upgrading Content Databases --"
        ##$Host.UI.RawUI.BackgroundColor = "Black"
        # Only allow use of SQL snapshots when updating content databases if we are on SP2013 or earlier, as there is no benefit with SP2016+ per https://blog.stefan-gossner.com/2016/04/29/sharepoint-2016-zero-downtime-patching-demystified/
        if ($useSqlSnapshot -and $spVer -le "15")
        {
            $UseSnapshotParameter = @{UseSnapshot = $true}
        }
        else
        {
            $UseSnapshotParameter = @{}
            Write-Verbose -Message " - Not using SQL snapshots to upgrade content databases, either because useSQLSnapshot not specified or the SharePoint farm is 2016 or newer."
        }
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
        # Updated to include all content databases, including ones that are "stopped"
        [array]$contentDatabases = Get-SPDatabase | Where-Object {$null -ne $_.WebApplication} | Sort-Object Name
        Write-Host -ForegroundColor White " - Upgrading SharePoint content databases:"
        foreach ($contentDatabase in $contentDatabases)
        {
            Write-Host -ForegroundColor White "  - $($contentDatabase.Name) ($($contentDatabases.IndexOf($contentDatabase)+1) of $($contentDatabases.Count))..."
            $contentDatabase | Upgrade-SPContentDatabase -Confirm:$false @UseSnapshotParameter
            Write-Host -ForegroundColor White "  - Done upgrading $($contentDatabase.Name)."
        }
    }
    # Kick off a separate PowerShell process to update content databases prior to running PSConfig
    Write-Host -ForegroundColor White " - Upgrading content databases in a separate process..."
    # Some special accomodations for older OSes and PowerShell versions
    if (((Get-CimInstance -ClassName Win32_OperatingSystem).Version -like "6.1*" -or (Get-CimInstance -ClassName Win32_OperatingSystem).Version -like "6.2*" -or (Get-CimInstance -ClassName Win32_OperatingSystem).Version -like "6.3*") -and ($spVer -eq "14"))
    {
        $upgradeContentDBJob = Start-Job -Name "UpgradeContentDBJob" -ScriptBlock $upgradeContentDBScriptBlock
        Write-Host -ForegroundColor Cyan " - Waiting for content databases to finish upgrading..." -NoNewline
        While ($upgradeContentDBJob.State -eq "Running")
        {
            # Wait for job to complete
            Write-Host -ForegroundColor Cyan "." -NoNewline
            Start-Sleep -Seconds 1
        }
        Write-Host -ForegroundColor Green "$($upgradeContentDBJob.State)."
    }
    else
    {
        Start-Job -Name "UpgradeContentDBJob" -ScriptBlock $upgradeContentDBScriptBlock | Receive-Job -Wait
    }
    Write-Host -ForegroundColor White " - Done upgrading databases."
}
function Clear-SPConfigurationCache
{
    [CmdletBinding()]
    param ()
    # Based on manual steps provided here:
    # http://blogs.msdn.com/b/jamesway/archive/2011/05/23/sharepoint-2010-clearing-the-configuration-cache.aspx
    Try
    {
            Write-Host -ForegroundColor White " - Clearing SP configuration cache for SharePoint 2010..."
        if ((Get-Service -Name SPTimerV4).Status -eq "Running")
        {
            # Stop SP Timer Service
            Write-Host -ForegroundColor White "  - Stopping timer service..."
            Stop-Service SPTimerV4
        }
        # Get the location of the cache files; if there is more than one folder, grab the latest one
        $cacheParentDir = "$env:SystemDrive\ProgramData\Microsoft\SharePoint\Config"
        $cacheSubDir = Get-ChildItem -Path $cacheParentDir -Filter "*-*-*-*-*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $cacheDir = Join-Path -Path $cacheParentDir -ChildPath $cacheSubDir
        # Grab the cache.ini file
        $cacheIni = Get-Content "$cacheDir\cache.ini"
        # Replace the contents of the cache.ini file with a single '1'
        Write-Host -ForegroundColor White "  - Modifying cache.ini file..."
        If ($cacheIni -ne "1")
        {
            Set-Content -Path "$cacheDir\cache.ini" -Value "1" -Force
        }
        # Delete all the XML files in the folder
        Write-Host -ForegroundColor White "  - Purging XML files from $cacheDir..."
        ForEach ($xmlFile in (Get-ChildItem -Path $cacheDir -Filter "*.XML"))
        {
            Remove-Item -Path (Join-Path -Path $cacheDir -ChildPath $xmlFile)
        }
    }
    Catch
    {
        Write-Warning $_
    }

    Finally
    {
        if ((Get-Service -Name SPTimerV4).Status -ne "Running")
        {
            # Restart the SP Timer Service
            Write-Host -ForegroundColor White "  - Attempting to start timer service..."
            Start-Service SPTimerV4 -ErrorAction SilentlyContinue
        }
        Write-Host -ForegroundColor White " - Done clearing configuration cache."
    }
}


function Get-SPYear
{
    $spYears = @{"14" = "2010"; "15" = "2013"; "16" = "2016"} # Can't use this hashtable to map SharePoint 2019 versions because it uses version 16 as well
    $farm = Get-SPFarm -ErrorAction SilentlyContinue
    [string]$spVer = $farm.BuildVersion.Major
    [string]$spBuild = $farm.BuildVersion.Build
    if (!$spVer -or !$spBuild)
    {
        Start-Sleep 10
        throw "Could not determine version of farm."
    }
    $spYear = $spYears.$spVer
    # Accomodate SharePoint 2019 (uses the same major version number, but 5-digit build numbers)
    if ($spBuild.Length -eq 5)
    {
        $spYear = "2019"
    }
    return $spVer, $spYear
}
#endregion



function DistributedCache(){

    [CmdletBinding()]
    param
    (
        #[String]$server,
        [bool]$enable
    )

	if($enable){
        Write-Host "`n`nProvisioning Distributed Cache on $env:COMPUTERNAME`n" -ForegroundColor Cyan
       $DCattempts = 0
        Do {
            $instanceName ="SPDistributedCacheService Name=AppFabricCachingService"
            $serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq $env:computername}
            #only valid when one Servcie running on farm
            #$serviceInstance = Get-SPServiceInstance | ? {$_.TypeName -like "Distributed Cache"}
            $serviceInstance.Provision()

            Start-Sleep -Seconds 10
            write-host " - Attempt $DCattempts on $env:COMPUTERNAME with instance `"$($serviceInstance)`" current status is: `"$($serviceInstance.Status)`": ." -ForegroundColor White
            $DCattempts++

            if($DCattempts -eq 3){
                Write-host " - Adding SPDistributedCacheServiceInstance intead." -ForegroundColor Cyan
                Add-SPDistributedCacheServiceInstance
            } 

        }
        While (($serviceInstance.Status -ne "Online") -and ($DCattempts -ne '4'))
        if($DCattempts -eq '4'){
        write-host "FAILED at $DCattempts try. Distribute DC manually!!" -ForegroundColor Red
        }else{
        write-host "Distributed Cache on $env:COMPUTERNAME has been Provisioned..." -ForegroundColor Green
        }
	}else{
        Write-Host "`n`nUnprovisioning Distributed Cache on $env:COMPUTERNAME`n" -ForegroundColor Cyan
        $instanceName ="SPDistributedCacheService Name=AppFabricCachingService"
        $serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq $env:computername}
        $serviceInstance.Unprovision()
	}
	
}


function Backup-IISConfiguration
{
  
  $SiteManager = Get-IISServerManager
  $SMs = $SiteManager.Sites

  foreach ($SM in $SMs)
  {
    
        $GWs = Get-Website $SM.name
            foreach ($GW in $GWs){
                if($sm.name -like '*GetOrganized*' -or $sm.name -like '*GO*' -or $sm.name -like '*SharePoint*'){
                    $Filename = 'web.config'
                    $PathToFile = $GW.physicalpath +'\'+ $Filename
                    if($PathToFile -ne $null){
                    if(get-item -path $PathToFile){
                        #Write-host ' - Copying file' $PathToFile
                        $date = Get-Date -format "yyyy_MM_dd-H_mm"
                        Copy-Item $PathToFile -Destination $PathToFile-$date.bak -force 

                        if(get-item -path $PathToFile-$date.bak){
                            write-host "  - IIS Configuration on $env:COMPUTERNAME has been completed successfully with: $PathToFile" -ForegroundColor white
                        }

                    }

                }else{
                        write-host "  - There is some issues please check site name...." -ForegroundColor Red
                        }
                }
                }

  }
}


function Backup-SPFarmConfiguration
{
    [CmdletBinding()]
    param
    (
        [String]$SPFarmBackupDirPath
    )
$SPFarmBackupDirName = "SPFarmBackup"
if(!(Test-Path "$SPFarmBackupDirPath\$SPFarmBackupDirName")) {
    New-Item -Path $SPFarmBackupDirPath -Name $SPFarmBackupDirName -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
    Write-Host " - Backing up the SPFarm..." -ForegroundColor White
    Backup-SPFarm -Directory $SPFarmBackupDirPath\$SPFarmBackupDirName -BackupMethod full -Percentage 5 -Force -ErrorAction Stop -ConfigurationOnly
    Write-Host "  - Congratulations! The FULL backup completed successfully!" -ForegroundColor Cyan
}else{
    Write-Host " - Folder with SPFarmConfiguration already exist uder: `"$SPFarmBackupDirPath`"" -ForegroundColor White
}
}


function Test-PendingReboot
{
if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { return $true }
if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { return $true }
if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA Ignore) { return $true }
try { 
   $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
   $status = $util.DetermineIfRebootPending()
   if(($status -ne $null) -and $status.RebootPending){
     return $true
   }
}catch{}

 return $false
}

function Set-WindowsFirewallRuleAction
{
    [CmdletBinding()]
    param
    (
        [bool]$Action,
        [String]$FirewallPortNumber

    )
	if($Action){
        Write-Host " - Blocking traffic on port: $FirewallPortNumber ..." -ForegroundColor White
        Get-NetFirewallPortFilter | ? LocalPort -eq $FirewallPortNumber | Get-NetFirewallRule | Set-NetFirewallRule -Action Block 
        }else{
        Write-Host " - Allowing traffic on port: $FirewallPortNumber ..." -ForegroundColor White
        Get-NetFirewallPortFilter | ? LocalPort -eq $FirewallPortNumber | Get-NetFirewallRule | Set-NetFirewallRule -Action Allow 
    }
}

function Test-InstalledFeatures
{
$missings = Get-SPContentDatabase | Test-SPContentDatabase | ? { $_.Category -eq "MissingFeature" }
if($missings -ne $null){
    write-host ' - Please review missing features in your installation:' -ForegroundColor White
    $missings | ft -AutoSize
    write-host 'Details:' -ForegroundColor White
        foreach($Missing in $missings){
            write-host '' -ForegroundColor White
            write-host $missing.message -ForegroundColor Red 
            return $true
        }
    }else{
    return $false
    }
}


function Test-FeaturesStatus
{
  $SPWebApps = get-SPWebApplication  

  foreach ($SPWebApp in $SPWebApps)
  {
  $ContentDBColl = (Get-SPWebApplication -Identity $SPWebApp.URL).ContentDatabases
      foreach ($contentDB in $ContentDBColl)
      {
           if($ContentDB.name -ne $null){
             $TestDBs = Test-SPContentDatabase -Name $ContentDB.name -WebApplication $SPWebApp.URL
                foreach ($TestDB in $TestDBs){
                    if($TestDB.UpgradeBlocking -eq $true){
                        Write-host 'Feature blockling upgrade?: ' $TestDB.UpgradeBlocking
                        return $true
                        }else{
                        return $false

                    }
                }
            } 

      }
  }
}

# Author
#https://gist.github.com/ShauneDonohue/3d4f8cc78771c2bd35e5daa0c17ba41c

function Update-ContentDatabasesMultiThreads{
[CmdletBinding()]
Param(
[Parameter(Mandatory=$false, HelpMessage="The Farm account to connect to SharePoint")] [string] $farmaccount = "itlup\sharepoint",
[Parameter(Mandatory=$false, HelpMessage="Web application URL to the use for upgrading")][string] $weburl = "https://test-partner.example.com/"
)


#$creds = (Get-Credential $farmaccount )

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

# Number of current threads to upgrade content databases in parallel.
# This is limited to a max of logical processors on the box, but the really stress will be on the SQL server
$maxUpgradeThreads = 4

function Get-ProcessorCount() {
    return [int](@(Get-WmiObject -Class Win32_Processor) | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
}

function New-ContentDatabaseUpgradeJob() {
    param(
        [Parameter(Mandatory=$true)][string]$ContentDatabaseName,
        [Parameter(Mandatory=$true)][System.Management.Automation.Runspaces.RunspacePool]$RunspacePool
    )
    $ErrorActionPreference = "Stop"
    try {
        # create an upgrade job
        $upgradeJob = [System.Management.Automation.PowerShell]::Create()
        # set the runspace pool
        $upgradeJob.RunspacePool = $RunspacePool
        # add the script block and the parameters to execute
        $powershell = $upgradeJob.AddScript( $UpgradeContentDatabaseScriptBlock )
        $powershell = $upgradeJob.AddParameter( "ContentDatabaseName", $ContentDatabaseName )
 
        return $upgradeJob   
    }
    catch {
        throw $_.Exception
    }
    finally {
        $ErrorActionPreference = "Continue"
    }
}

function New-RunspacePool() {
    param([Parameter(Mandatory=$true)][int]$MaxRunspaces)
    
    $ErrorActionPreference = "Stop"
    try
    {
        $warning = ""
        # create a shared session state that imports the SharePoint Snap-In
        $defaultSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $snapInInfo = $defaultSessionState.ImportPSSnapIn( "Microsoft.SharePoint.PowerShell", [ref]$warning )
        $defaultSessionState.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread

        # create the runspace pool that will be unique for all of the upgrade jobs
        $runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool($defaultSessionState)
        $added = $runspacePool.SetMinRunspaces(1)
        $added = $runspacePool.SetMaxRunspaces($MaxRunspaces)
        $runspacePool.Open()    
        return $runspacePool
    }
    catch
    {
        throw $_.Exception
    }
    finally
    {
        $ErrorActionPreference = "Continue"
    }
}

filter Write-ColorCodedTable {
    param
    (
        [HashTable]$ColorMappings
    )

    $ForegroundColor = "White" # default color for table headers
    if( $Host.UI.RawUI.ForegroundColor -eq "-1" )
    {
        $ForegroundColor = $Host.UI.RawUI.ForegroundColor
    }

    $lines = $_ -split '\r\n'

    foreach( $line in $lines )
    {
        $ColorMappings.GetEnumerator() | % {
            if($line -match $_.Key)
            {
                 $ForegroundColor = $_.Value
            }
        }

        Write-Host $line -ForegroundColor $ForegroundColor
    }
}

function updateContentDatabases($contentDatabaseNamesToUpgrade){
    $resultColorHash = @{ 
        # message regex                               = display color 
        "No Upgrade Required|Completed|DEBUG" = "Green" 
        "Error|Invalid|Failed|Unknown"        = "Red" 
    }

    $UpgradeContentDatabaseScriptBlock = {
        param([string]$ContentDatabaseName)
        $upgradeResultMessage = "Completed"

        try
        {
            $contentDatabase = Get-SPContentDatabase | ? { $_.Name -eq $ContentDatabaseName }
            if($contentDatabase -and $contentDatabase.NeedsUpgrade) {
                $upgradeErrorMessage = $upgradeWarningMessage = ""
                # we need to stagger start time to prevent upgrades from starting at the same time 
                # because they will have a session naming conflict and cause the upgrade to fail
                Start-Sleep -Seconds $( Get-Random -Minimum 5.0 -Maximum 60.0 )
                $databaseUpgradeStopWatch = Measure-Command {
                    $contentDatabase | Upgrade-SPContentDatabase -Confirm:$false -ErrorVariable upgradeErrorMessage -WarningVariable upgradeWarningMessage
                }
            }
            elseif($contentDatabase -and !$contentDatabase.NeedsUpgrade) {
                # return the object showing the db doesn't need to be upgraded
                return New-Object PSObject -Property @{
                    ContentDatabaseName = $ContentDatabaseName;
                    ExecutionTime       = "0";
                    Result              = "No Upgrade Required";
                }
            }
        }
        catch {
            $upgradeResultMessage = $_.Exception.ToString()
        }

        if($upgradeWarningMessage) {
            $upgradeResultMessage = $upgradeWarningMessage
        }

        if($upgradeErrorMessage) {
            $upgradeResultMessage = $upgradeErrorMessage
        }

        return New-Object PSObject -Property @{
            ContentDatabaseName = $ContentDatabaseName;
            ExecutionTime       = $databaseUpgradeStopWatch.TotalSeconds.ToString("N1")
            Result              = $upgradeResultMessage;
        }
    }
    
    # limit the number of threads to a max of the number of logical processsors on the box
    if($maxUpgradeThreads -gt $(Get-ProcessorCount)) {
        $maxUpgradeThreads = $(Get-ProcessorCount)
        Write-Host "`n`tLimiting max threads to the number of logical processsors `"$($maxUpgradeThreads)`".`n" -ForegroundColor Yellow
    }
    # start the upgrades
    $databaseUpgradeJobs = $lastCheckedCompletedUpgradeJobs = $databaseUpgradeResults = @()
    $databaseUpgradeProgressCounter = 0
    $warning = $null

    # create the run space
    $runspacePool = New-RunspacePool -MaxRunspaces $maxUpgradeThreads
    Write-Host "`n`nCreating Content Database Upgrade Jobs`n" -ForegroundColor Green

    # enumerate all the content databases to upgrade
    $contentDatabaseNamesToUpgrade | % {
        $databaseName = $_
        try {
            # Make sure the database exists
            $contentDatabase = Get-SPContentDatabase -Identity $databaseName -ErrorAction SilentlyContinue
            if($contentDatabase) {
                Write-Host "`tQueuing upgrade job for database: $databaseName"
                # don't crush the machine with starting jobs.
                Start-Sleep -Seconds 5
                # create the upgrade job
                $contentDatabaseUpgradeJob = New-ContentDatabaseUpgradeJob -ContentDatabaseName $databaseName -RunspacePool $runspacePool

                # add the job to our upgrade job list and start the job
                $databaseUpgradeJobs += New-Object PSObject -Property @{
                    PowerShell       = $contentDatabaseUpgradeJob;
                    Runspace         = $contentDatabaseUpgradeJob.BeginInvoke();
                    ContentDatabase  = $databaseName;
                    UpgradeJobResult = $null;
                    NeedsUpgrade     = $null;
                }
            }
        }
        catch {
            Write-Host "`t`tError:   Error creating or starting upgrade job for database: $databaseName." -ForegroundColor Red
            Write-host "`t`tDetails: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if( $databaseUpgradeJobs.Count -gt 0 ) {
        Write-Host "`nDatabase Upgrade Progress`n" -ForegroundColor Green
        # keep looping until all the installation jobs are complete
        do {
            try {
                # pause for a few seconds between checks for completed jobs
                Start-Sleep -Seconds 5
                # get a list of all the completed database jobs
                $completedDatabaseUpgradeJobs = @($databaseUpgradeJobs | ? { $_.Runspace.IsCompleted })

                # check if any more jobs completed since last check
                if($lastCheckedCompletedUpgradeJobs.Count -ne $completedDatabaseUpgradeJobs.Count) {
                    # print out the computer names for each job that completed since the last check
                    $completedDatabaseUpgradeJobs | ? { $lastCheckedCompletedUpgradeJobs -notcontains $_ } | % {
                        Write-Output "`t$(Get-Date) - $($_.ComputerName): Upgrade of $($_.ContentDatabase) has completed."
                    }

                    # update the list of completed jobs
                    $lastCheckedCompletedUpgradeJobs = $completedDatabaseUpgradeJobs

                    # update the progress bar
                    Write-Progress `
                        -Activity       "Content Database Upgrade Progress" `
                        -Status         "Progress: $($completedDatabaseUpgradeJobs.Count) of $($databaseUpgradeJobs.Count) Upgrades Completed" `
                        -PercentComplete $(($($Progress.Count)/$($databaseUpgradeJobs.Count))*100)
                }
            }
            catch {
                Write-Host "`t`tError:   Error checking for job completion status" -ForegroundColor Red
                Write-host "`t`tDetails: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        while( $databaseUpgradeJobs | ? { !$_.Runspace.IsCompleted } )

        try {
            # collect the job results
            $databaseUpgradeJobs | % {
                $_.UpgradeJobResult = $_.PowerShell.EndInvoke($_.Runspace)
            }

            # pull the NeedsUpgrade property from the attached database
            $databaseUpgradeJobs | % {
                $_.NeedsUpgrade = $(Get-SPContentDatabase -Identity $_.ContentDatabase).NeedsUpgrade
            }

            # build an output object collection we can display and write to csv
            $databaseUpgradeJobs | % {
                $databaseUpgradeResults += New-Object PSObject -Property @{
                    "Database"       = $_.ContentDatabase;
                    "Execution Time" = $_.UpgradeJobResult | % { $_.ExecutionTime }; # hack for PowerShell 2.0
                    "Upgrade Result" = $_.UpgradeJobResult | % { $_.Result };        # hack for PowerShell 2.0
                    "Needs Upgrade"  = $_.NeedsUpgrade;
                }
            }

            Write-Host "`nContent Database Upgrade Results`n" -ForegroundColor Green
            # write out the results to the screen and .csv file
            $databaseUpgradeResults | Sort Database | FT Database, "Upgrade Result", "Execution Time", "Needs Upgrade" -AutoSize | Out-String | Write-ColorCodedTable -ColorMappings $resultColorHash
            $databaseUpgradeResults | Export-Csv -Path $outputFile -NoTypeInformation

            if( $outputFile.IndexOfAny("\") -gt 0 ) {
                Write-Output "`nUpgrade results written to $outputFile`n"    
            }
            else {
                Write-Output "`nUpgrade results written to $((Get-Item -Path ".\" -Verbose).FullName)\$outputFile`n"
            }
        }
        catch {
            Write-Host "`t`tError:   Error upgrading content database: $databaseName." -ForegroundColor Red
            Write-host "`t`tDetails: $($_.Exception.Message)" -ForegroundColor Red
            return
        }
        finally {
            if( $runspacePool -and $runspacePool.RunspacePoolStateInfo -and $runspacePool.RunspacePoolStateInfo.State -eq [System.Management.Automation.Runspaces.RunspaceState]::Opened) {
                $runspacePool.Close()
                $runspacePool.Dispose()
            }
        }
    }
}

# Example: upgrade specific databases
#$contentDatabaseNamesToUpgrade = @("SP2010_CONTENT_001", "SP2010_CONTENT_002", "SP2010_CONTENT_003" )

# Example: upgrade all the content databases hosted on SQL01 
#$contentDatabaseNamesToUpgrade = Get-SPContentDatabase | ? { $_.Server -eq "SQL01" } | % { $_.Name }

# Example: upgrade all the content databases on the finance web application 
#$contentDatabaseNamesToUpgrade = Get-SPContentDatabase | ? { $_.WebApplication.Url -match "finance.contoso.com" } | % { $_.Name }

# Example: upgrade all the content databases attached to the SharePoint farm
$contentDatabaseNamesToUpgrade = Get-SPContentDatabase | % { $_.Name }

# log file to write the results to
$outputFile = "Upgrade-ContentDatabaseB2B.results.{0}.csv" -f [DateTime]::Now.ToString("yyyy-MM-dd_hh-mm-ss")

$startTime = Get-Date

# hack for powershell 2.0, so we always have an array of database names
if($contentDatabaseNamesToUpgrade.GetType().Name -eq "String") {
    $contentDatabaseNamesToUpgrade = @($contentDatabaseNamesToUpgrade)
}

updateContentDatabases $contentDatabaseNamesToUpgrade

$TimeSpan = New-TimeSpan -Start $startTime -End (Get-Date)
Write-Output "`n`nExection Time: $($TimeSpan.Hours):$($TimeSpan.Minutes):$($TimeSpan.Seconds)`n"

}

