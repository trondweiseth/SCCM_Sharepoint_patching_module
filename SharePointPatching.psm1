<#
.SYNOPSIS
Toolset to aid in SharePoint patching through SCCM and Software Center.

.DESCRIPTION
PS:Need to create server lists by running Initiate_Serverlists.ps1 firt time before using this script if you haven't already.

To get a list of all the available commands run: Get-Command SP-*

All fuctions are created to loop through $servers variable.

All the functions that require pscredentials validates that pscredentials are set with TestPSCredentials function and prompt for it if missing.
This can be set with the function pscredentials

Any functions that changes a state or value have validation before running. 
Any function that is only informational are not validated before running.

.EXAMPLE
    SP-ClearSCCMCache
.EXAMPLE
    SP-GetCheckpoints
.EXAMPLE
    SP-CreateCheckpoints
.EXAMPLE
    SP-RemoveCheckpoints
.EXAMPLE
    SP-ForceStopServers
.EXAMPLE
    SP-GetApplications
    SP-GetApplications [[-AppName] <string>]
.EXAMPLE
    SP-InstallationStatus
    SP-InstallationStatus [[-AppName] <string>] [-Time <int>] [-Wait]
.EXAMPLE
    SP-ListSCCMPackages
    SP-ListSCCMPackages [[-Application] <string>]
.EXAMPLE
    SP-RunSCCMClientAction
    SP-RunSCCMClientAction [[-ClientAction] {MachinePolicy | DiscoveryData | ComplianceEvaluation | AppDeployment | HardwareInventory | UpdateDeployment | UpdateScan | SoftwareInventory}]
.EXAMPLE
    SP-Servers
    SP-Servers [[-Global:Enviroment] {YT01 | AT05 | TT02 | TUT01 | TUL | PROD}] [-EditFile]
.EXAMPLE
    SP-StartServers
.EXAMPLE
    SP-StopServers
.EXAMPLE
    SP-TestConnection
    SP-TestConnection [[-Time] <int>] [-Wait]
.EXAMPLE
    SP-TriggerInstallation
    SP-TriggerInstallation [-AppName] <string>
    SP-TriggerInstallation [-Method] {Install | Uninstall}
.EXAMPLE
    SP-VMConnect
    SP-VMConnect [-Wait]
.EXAMPLE
    SP-VMLog
    SP-VMLog [[-Newest] <int>]
.EXAMPLE
    SP-VMStatus

.NOTES
Author: Trond Weiseth 
#>

$scriptlocation = $MyInvocation.MyCommand.Path
Clear-Variable servers -ErrorAction SilentlyContinue

Function SP-Servers() {

    [CmdletBinding()]
    param
    (
        [ValidateSet('YT01', 'AT05', 'TT02', 'TUT01', 'TUL', 'PROD')]
        [string]$Global:Enviroment,
        [switch]$EditFile
    )
    
    function helpmsg {
        Write-Host -ForegroundColor Yellow "SYNTAX: SP-Servers [{YT01 | AT05 | TT02 | TUT01 | TUL | PROD}] [-EditFile]"
    }

    if (! $Enviroment) { helpmsg; break }
    $serverlistfolder = "$HOME\Documents\SharepointHosts"
    if ($EditFile) { notepad.exe $serverlistfolder\$Enviroment.txt | Out-Null }
    $Global:servers = Get-Content $serverlistfolder\$Enviroment.txt
    Write-Host -ForegroundColor Yellow "Servers are set to:"
    $servers | ForEach-Object { Write-Host -ForegroundColor Cyan "$_" }
}

Function pscredentials{
    $uname = ("$env:USERDOMAIN\$env:USERNAME")
    [pscredential]$Global:cred = Get-Credential $uname
}

Function TestPSCredentials {
    if (!$cred) {pscredentials}
}

Function validation() {

    Write-Host -ForegroundColor Red "Do you want to run $command on enviroment ${Enviroment}? (y/n):" -NoNewline
    $validate = Read-Host -InformationAction SilentlyContinue
    if ($validate -notmatch 'y') { break }
}

Function errormsg {

    if ($null -eq $Servers) {
        Write-Warning "No servers are selected. Run SP-Servers [{YT01 | AT05 | TT02 | TUT01 | TUL | PROD}]"
        break
    }
}

Function ServerHeader {
    write-host -BackgroundColor DarkCyan (" "*$Length) -NoNewline
    Write-Host -ForegroundColor Yellow -BackgroundColor DarkCyan  $server -NoNewline
    write-host -BackgroundColor DarkCyan (" "*$Length)
}

Function LoadSCCMModule {

    Begin
    {
        # Site configuration
        $SiteCode = "001" # Site code 
        $ProviderMachineName = "SCCMserver.contoso.local" # SMS Provider machine name
    }
    Process
    {
        # Customizations
        $initParams = @{}
        if ((Get-Module ConfigurationManager) -eq $null) {
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
        }

        # Connect to the site's drive if it is not already present
        if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
        }

        # Set the current location to be the site code.
        Set-Location "$($SiteCode):\" @initParams
    }
    End{}
}

Function SP-StartServers {

    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
    }
    Process
    {
        foreach ($server in $servers) {
            Get-SCVirtualMachine $server |  start-SCVirtualMachine -RunAsynchronously | Select-Object Name, Status | Format-Table -AutoSize
        }

        while ((SP-VMStatus | Select-Object Status) -notmatch "Running") { Start-Sleep 2 }
        SP-VMStatus
    }
    End{}
}

Function SP-StopServers {

    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
    }
    Process
    {
        foreach ($server in $servers) {
            Get-SCVirtualMachine $server |  Stop-SCVirtualMachine -Shutdown -RunAsynchronously | Select-Object Name, Status | Format-Table -AutoSize
        }

        while ((SP-VMStatus | Select-Object Status) -notmatch "PowerOff") { Start-Sleep 2 }
        SP-VMStatus
    }
    End{}
}

Function SP-ForceStopServers {

    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
    }
    Process
    {
        foreach ($server in $servers) {
            try {
                stop-computer $server -force -ErrorAction stop
            }
            catch {
                Write-Error $_.Exception.Message
            }
        }

        while ((SP-VMStatus | Select-Object Status) -notmatch "PowerOff") { Start-Sleep 2 }
        SP-VMStatus
    }
    End{}
}

Function SP-VMStatus {
    
    Begin
    {
        errormsg
    }
    Process
    {
            foreach ($server in $servers) {
            Get-SCVirtualMachine $server | select-object name, status
        }
    }
    End {}
}

Function SP-CreateCheckpoints {

    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
        $spchkpointdescription = {BFC-2337928 - SharePoint Patching}
        $validate = Read-Host -Prompt "Current description is: '$spchkpointdescription'. Do you want you change the description? (y/n) (Default n)" -ErrorAction SilentlyContinue

        if ($validate -imatch 'y') {
            $response = Read-Host -Prompt "Description "
            [string]$newspchkpointdescription = $response.Trim('"')
            $savedescription = Read-Host -Prompt "Do you want to save the description? (y/n) (Default n)" -ErrorAction SilentlyContinue
            if ($savedescription -imatch 'y') {
                $currentdescription = ((Get-Content $scriptlocation | Where-Object { $_ -match "spchkpointdescription" }).Split('{')[1]).Trim("}")
                (Get-Content $scriptlocation).Replace("$currentdescription", "$newspchkpointdescription") | Set-Content -Path $scriptlocation
            }
        }
    }
    Process
    {
        foreach ($server in $servers) {
            Get-SCVirtualMachine $server | New-SCVMCheckpoint -Description $spchkpointdescription -RunAsynchronously | Select-Object Name,MostRecentTaskIfLocal,Description | Format-Table -AutoSize -Wrap
        }
    }
    End{}
}

Function SP-GetCheckpoints() {

    Begin
    {
        errormsg
    }
    Process
    {
        foreach ($server in $servers) {
            $checkpoints = Get-SCVirtualMachine $server | Get-SCVMCheckpoint
            if ($checkpoints) {
                $checkpoints | ForEach-Object {
                    Write-Host -ForegroundColor Green 'Name :' $_.Name -NoNewline
                    Write-Host -ForegroundColor Yellow '  Description :' $_.Description
                }
            }
        }
    }
    End{}
}

Function SP-RemoveCheckpoints {

    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
    }
    Process
    {
        foreach ($server in $servers) {
            $Checkpoints = Get-SCVMCheckpoint -VM $server
            foreach ($checkpoint in $Checkpoints) {
                Remove-SCVMCheckpoint -VMCheckpoint $Checkpoint -RunAsynchronously | Select-Object Name,MostRecentTaskIfLocal,Description | Format-Table -AutoSize -Wrap
            }
        }
    }
    End{}
}

Function SP-VMConnect() {

    [CmdletBinding()]
    param
    (
        [switch]$Wait
    )

    Begin
    {
        errormsg
    }
    Process
    {
        foreach ($server in $servers) {
            if ($wait) {
                while (! (Test-Connection $server -ErrorAction SilentlyContinue)) { Start-Sleep 3 }
            }
            mstsc /w:1024 /h:800 /v:$server
        }
    }
    End{}
}

Function SP-TestConnection {

    [CmdletBinding()]
    param
    (
        [switch]$Wait,
        [int]$Time=3
    )

    Begin
    {
        errormsg
        $resultlist = [System.Collections.ArrayList]@()
    }
    Process
    {
        foreach ($server in $servers) {
            if ($Wait) { while (! (Test-Connection $server -Count 1 -ErrorAction SilentlyContinue)) { Start-Sleep $Time } }
            try {
                $result = Test-Connection $server -Count 1 -ErrorAction Stop
                $Prop=[ordered]@{
                    'Server'=$result.Address
                    'IPV4Address'=$result.IPV4Address
                    'Bytes'=$result.ReplySize
                    'Time(ms)'=$result.ResponseTime
                    'Status'='OK'
                }
                $obj=New-Object -TypeName psobject -Property $Prop
                [void]($resultlist.Add($obj))
            }
            catch {
                Write-Host -ForegroundColor Red "$server is not responding."
            }
        }
    }
    End
    {
        Write-Output $resultlist | ft
    }
}

Function SP-ClearSCCMCache {

    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
        TestPSCredentials
    }
    Process
    {
        $servers | foreach-object {
            Invoke-Command -ComputerName $_ -ScriptBlock {
                ## Initialize the CCM resource manager com object
                [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'
                ## Get the CacheElementIDs to delete
                $CacheInfo = $CCMComObject.GetCacheInfo().GetCacheElements()
                ## Remove cache items
                ForEach ($CacheItem in $CacheInfo) {
                    $null = $CCMComObject.GetCacheInfo().DeleteCacheElement([string]$($CacheItem.CacheElementID))
                }
                return $true
            } -Credential $cred
        }
    }
    End{} 
}

Function SP-RunSCCMClientAction {

    [CmdletBinding()]
    param
    (  
        [ValidateSet('MachinePolicy', 
            'DiscoveryData', 
            'ComplianceEvaluation', 
            'AppDeployment',  
            'HardwareInventory', 
            'UpdateDeployment', 
            'UpdateScan', 
            'SoftwareInventory')] 
        [string[]]$ClientAction
    )
    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
        TestPSCredentials
        $ActionResults = @()
    }
    Process
    {
        $Global:servers | ForEach-Object {
            Try {
                $ActionResults = Invoke-Command -ComputerName $_ -Credential $cred -ArgumentList (, $ClientAction) -ErrorAction Stop -ScriptBlock { param($ClientAction)
 
                    Foreach ($Item in $ClientAction) {
                        $Object = @{} | Select-Object "Action name", Status
                        Try {
                            $ScheduleIDMappings = @{ 
                                'MachinePolicy'        = '{00000000-0000-0000-0000-000000000021}'; 
                                'DiscoveryData'        = '{00000000-0000-0000-0000-000000000003}'; 
                                'ComplianceEvaluation' = '{00000000-0000-0000-0000-000000000071}'; 
                                'AppDeployment'        = '{00000000-0000-0000-0000-000000000121}'; 
                                'HardwareInventory'    = '{00000000-0000-0000-0000-000000000001}'; 
                                'UpdateDeployment'     = '{00000000-0000-0000-0000-000000000108}'; 
                                'UpdateScan'           = '{00000000-0000-0000-0000-000000000113}'; 
                                'SoftwareInventory'    = '{00000000-0000-0000-0000-000000000002}'; 
                            }
                            $ScheduleID = $ScheduleIDMappings[$item]
                            Write-Verbose "Processing $Item - $ScheduleID"
                            [void]([wmiclass] "root\ccm:SMS_Client").TriggerSchedule($ScheduleID);
                            $Status = "Success"
                            Write-Verbose "Operation status - $status"
                        }
                        Catch {
                            $Status = "Failed"
                            Write-Verbose "Operation status - $status"
                        }
                        $Object."Action name" = $item
                        $Object.Status = $Status
                        $Object
                    }
 
                } | Select-Object @{n = 'ServerName'; e = { $_.pscomputername } }, "Action name", Status
            }  
            Catch {
                Write-Error $_.Exception.Message
            }
            Return $ActionResults
        }
    }
    End{}
}

Function SP-GetApplications {

    [CmdletBinding()]
    param
    (
        # AppName parameter
        [Parameter(Mandatory = $false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   Position = 0,
                   ParameterSetName='Parameter Set AppName',
                   HelpMessage='Name of application you want to search for on remote computer system center.')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("a")] 
        [String]$AppName
    )
    Begin
    {
        errormsg
        TestPSCredentials
        $Length=54
        $evalstates = @(
            "No state information is available"
            "Application is enforced to desired/resolved state"
            "Application is not required on the client"
            "Application is available for enforcement (install or uninstall based on resolved state). Content may/may not have been downloaded"
            "Application last failed to enforce (install/uninstall)"
            "Application is currently waiting for content download to complete"
            "Application is currently waiting for content download to complete"
            "Application is currently waiting for its dependencies to download"
            "Application is currently waiting for a service (maintenance) window"
            "Application is currently waiting for a previously pending reboot"
            "Application is currently waiting for serialized enforcement"
            "Application is currently enforcing dependencies"
            "Application is currently enforcing"
            "Application install/uninstall enforced and soft reboot is pending"
            "Application installed/uninstalled and hard reboot is pending"
            "Update is available but pending installation"
            "Application failed to evaluate"
            "Application is currently waiting for an active user session to enforce"
            "Application is currently waiting for all users to logoff"
            "Application is currently waiting for a user logon"
            "Application in progress, waiting for retry"
            "Application is waiting for presentation mode to be switched off"
            "Application is pre-downloading content (downloading outside of install job)"
            "Application is pre-downloading dependent content (downloading outside of install job)"
            "Application download failed (downloading during install job)"
            "Application pre-downloading failed (downloading outside of install job)"
            "Download success (downloading during install job)"
            "Post-enforce evaluation"
            "Waiting for network connectivity"
        )
    }
    Process
    {
        foreach ($server in $servers) {
            ServerHeader
            Invoke-Command -ComputerName $server -Credential $cred -ArgumentList $AppName, $evalstates -ScriptBlock {

                param
                (
                    $AppName,
                    $evalstates
                )

                (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" | Where-Object { $_.Name -match "$AppName" }) |
                Select-Object Name, InstallState, LastInstallTime, ResolvedState, AllowedActions, InProgressActions, @{ Name = 'EvaluationState'; Expression = { ($Evalstates[$_.EvaluationState]) } } | Format-Table -AutoSize -Wrap
            }
        }
    }
    End{}
}

Function SP-TriggerInstallation {

    [CmdletBinding()]
    Param
    (
        # AppName parameter
        [Parameter(Mandatory = $True,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position = 0,
                   ParameterSetName='Parameter Set AppName',
                   HelpMessage='Name of application to be installed.')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("a")] 
        [String]$AppName,

        # Method parameter
        [ValidateSet("Install", "Uninstall")]
        [Parameter(Mandatory = $true,
                   Position = 1,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set Method',
                   HelpMessage='Valid methods are Install/Uninstall. Use tab to autocomplete.')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("m")]
        [String]$Method
    )

    Begin
    {
        errormsg
        $command = (Get-PSCallStack).Command | Select-Object -First 1
        validation
        TestPSCredentials
    }
    Process
    {
        foreach ($server in $servers) {
            Invoke-Command -ComputerName $server -Credential $cred -ArgumentList $AppName, $Method -ScriptBlock {

                Param
                (
                    [String]$AppName,
                    [String]$Method
                )
                Begin {
                    $Application = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" | Where-Object { $_.Name -like $AppName })
                    $CCMApplicationName = $Application.FullName
                    if (!('' -eq $Application)) { $readval = Read-Host -Prompt "Do you want to install ${CCMApplicationName}? (y/n) (Default n)" }
                    if (!($readval -imatch 'y')) { break }
                    $Args = @{EnforcePreference = [UINT32] 0
                        Id                      = "$($Application.id)"
                        IsMachineTarget         = $Application.IsMachineTarget
                        IsRebootIfNeeded        = $False
                        Priority                = 'High'
                        Revision                = "$($Application.Revision)" 
                    } 
                }
                Process { 
                    Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Application -MethodName $Method -Arguments $Args 
                }
                End {}
            }
        }  
    }
    End{}
}

Function SP-InstallationStatus {

    [CmdletBinding()]
    param
    (
        # AppName parameter
        [Parameter(Mandatory = $false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   Position = 0,
                   ParameterSetName='Parameter Set AppName',
                   HelpMessage='Type a name for the application you want to validate installation for.')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("a")] 
        [String]$AppName,

        # Time parameter
        [Parameter(HelpMessage='Number in seconds to wait for next check. Default is 30.')]
        [ValidateNotNullOrEmpty()]
        [int]$Time=30,

        # Wait parameter
        [Parameter()]
        [switch]$Wait
    )
    Begin
    {
        errormsg
        TestPSCredentials
        if (!$AppName) {
            [string]$InstallStatusAppName = {Sharepoint 2013 CU 2021 September}
            $validate = Read-Host -Prompt "AppName is: '$InstallStatusAppName'. Do you want you change the AppName? (y/n) (Default n)" -ErrorAction SilentlyContinue
            if ($validate -imatch 'y') {
                $response = Read-Host -Prompt "AppName "
                $newInstallStatusAppName = $response.Trim('"')
                $saveAppName = Read-Host -Prompt "Do you want to save AppName? (y/n) (Default n)" -ErrorAction SilentlyContinue
                if ($saveAppName -imatch 'y') {
                    $currentAppName = ((Get-Content $scriptlocation | Where-Object { $_ -match "InstallStatusAppName" }).Split('{')[1]).Trim("}")
                    (Get-Content $scriptlocation).Replace("$currentAppName", "$newInstallStatusAppName") | Set-Content -Path $scriptlocation
                }
            }
        }
        else {
            $InstallStatusAppName = $AppName
            $saveAppName = Read-Host -Prompt "Do you want to save AppName? (y/n) (Default n)" -ErrorAction SilentlyContinue
            if ($saveAppName -imatch 'y') {
                $currentAppName = ((Get-Content $scriptlocation | Where-Object { $_ -imatch "InstallStatusAppName" }).Split('{')[1]).Trim("}")
                (Get-Content $scriptlocation).Replace("$currentAppName", "$InstallStatusAppName") | Set-Content -Path $scriptlocation
            }
        }

        Write-Host -ForegroundColor Yellow "Checking installation staus for $InstallStatusAppName..."
        LoadSCCMModule
        $xmloutput = Get-CMApplication -Name $InstallStatusAppName | Select-Object SDMPackageXML

        if (!($null -eq $xmloutput)) {
            $xmlitems = ([xml]($xmloutput.SDMPackageXML)).AppMgmtDigest.DeploymentType.Installer.DetectAction.InnerText
            $NewHash = ($xmlitems).Split('"')[1]
            $NewVersion = ($xmlitems).Split("'")[5]
        }
        else {
            Write-Warning "Could not find any application with the name $InstallStatusAppName"
        }

        function validatescript {

            param
            (
                [string]$NewHash,
                [string]$NewVersion
            )

            $installed = $null
            $FileHash = Get-FileHash -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.dll' -ErrorAction SilentlyContinue
            $version = (Get-Item -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.dll').Versioninfo.FileVersion
            if (($version -eq "$NewVersion" -and $FileHash.Hash -eq "$NewHash")) {
                $installed = "$env:COMPUTERNAME SP patch ready $(get-date)"
                Write-Host $installed -ForegroundColor Green
            }
            else {
                write-host "Not yet installed on $env:COMPUTERNAME $(get-date)" -ForegroundColor Red
            }
        }

        function waitvalidatescript {

            param
            (
                [string]$NewHash,
                [string]$NewVersion,
                [int]$Time
            )

            while (1) {
                $installed = $null
                $FileHash = Get-FileHash -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.dll' -ErrorAction SilentlyContinue
                $version = (Get-Item -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.dll').Versioninfo.FileVersion
                if (($version -eq "$NewVersion" -and $FileHash.Hash -eq "$NewHash")) {
                    $installed = "$env:COMPUTERNAME SP patch ready $(get-date)"
                    Write-Host $installed -ForegroundColor Green
                    break
                }
                else {
                    Write-Host "Still not installed on $env:COMPUTERNAME. Checking $server again in $Time seconds..." -ForegroundColor Red
                    Start-Sleep $Time
                }
            }
        }
    }
    Process
    {
        if ($Wait) {
            $servers | foreach-object {
                Invoke-Command -ComputerName $_ -Credential $cred -ArgumentList $NewHash, $NewVersion, $Time -ScriptBlock ${function:waitvalidatescript}
            }
        }
        else {
            $servers | foreach-object {
                Invoke-Command -ComputerName $_ -Credential $cred -ArgumentList $NewHash, $NewVersion -ScriptBlock ${function:validatescript}
            }
        }
    }
    End{}
}

Function SP-VMLog {

    [CmdletBinding()]
    param
    (
        [Parameter(HelpMessage='Type in the number of log lines to retrieve. Default 100')]
        [int]$Newest=100
    )

    Begin
    {
        errormsg
        [int]$Length=25
        Get-SCVirtualMachine | Out-Null
    }
    Process
    {
        foreach ($server in $servers) {
            ServerHeader
            Get-SCJob -Full -Newest $Newest | Where-Object { $_.ResultName -imatch "$server" } | Select-Object Name, Status, Progress, StartTime, Owner | Format-Table -AutoSize -Wrap
        }
    }
    End{}

}

Function SP-ListSCCMPackages {

    [CmdletBinding()]
    param
    (
        # Application parameter
        [Parameter(Mandatory = $false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   Position = 0,
                   ParameterSetName='Parameter Set Application',
                   HelpMessage='Name of application to search for. I.e. -Application "sharepoint 2013".')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Application
    )

    Begin
    {
        LoadSCCMModule
    }
    Process
    {
        if (!$Application) {
            $res = Get-CMApplication -Fast | select LocalizedDisplayName, CreatedBy, DateCreated, DateLastModified, IsDeployable, IsEnabled, IsDeployed, NumberOfDevicesWithApp, NumberOfDevicesWithFailure, SecuredScopeNames | Out-GridView -PassThru
        }
        else {
            $res = Get-CMApplication -Fast | select LocalizedDisplayName, CreatedBy, DateCreated, DateLastModified, IsDeployable, IsEnabled, IsDeployed, NumberOfDevicesWithApp, NumberOfDevicesWithFailure, SecuredScopeNames | where { $_.LocalizedDisplayName -imatch "$Application" } | Out-GridView -PassThru
        }
    }
    End{
    return $res | Format-Table -AutoSize -Wrap
    }
}
