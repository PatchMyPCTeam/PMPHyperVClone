<#
.SYNOPSIS
    Clone local Hyper-V VMs to a remote Hyper-V host.
.DESCRIPTION
    Clone local Hyper-V VMs to a remote Hyper-V host. 
    When invoked, you will asked which VMs (locally) you would like to clone from a multi-selection Out-GridView window.
    You will then be asked which volume on the host specified for -ComputerName you would like to store these VMs from another Out-GridView window.
    You can choose drive letter but you cannot choose folder. The VMs will always be stored in a folder at the root of the volume named "VMs". 
    For example, choosing the "D" volume will store the VMs in "D:\VMs" on the target host. 
    If a folder exists on the target host with the same name as the VM(s) you're cloning, the script will skip the VM. You must delete or rename the folder on the target host. This is to prevent accidental clobbering.
    If VMs are running, you will be asked if you want to shut them down. Not necessary but it is recommended.
    Requirements:
    - SYSTEM account of source host must be local admin on target host: VMMS service runs as SYSTEM by default and that's used for Export-VM. 
        - This could be negated by exporting locally first, then copy to target host
    - Invoking user be in the Hyper-V Administrators group on the source host or execute the script as administrator / elevated
.EXAMPLE
    PS C:\> .\Invoke-PMPHyperVClone.ps1 -ComputerName "TRX2.contoso.local"
    
    After selecting the VMs you want to export and the volume drive letter you wish to store the VMs on the remote host, the script will export the VMs and import them to Hyper-V on the target host.
.INPUTS
    This script does not accept input from the pipeline
.OUTPUTS
    A PSObject with two properties: VM and Result
.NOTES
    Author: Adam Cook
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        $Ports = 139,5985 # SMB and WinRM
        foreach ($Port in $Ports) {
            if (-not (Test-NetConnection -ComputerName $_ -Port $Port -ErrorAction "Stop" -WarningAction "Stop")) { 
                throw ("Unable to reach {0} port {1}" -f $_, $Port)
            }
        }
        $true
    })]
    [String]$ComputerName
)

$JobId = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$StorageThresholdGB = 5
$CimSessionOpts = New-CimSessionOption -Protocol "DCOM"
$CimSession = New-CimSession -ComputerName $ComputerName -SessionOption $CimSessionOpts -ErrorAction "Stop"

#region Determine which VMs to clone
$Title = "Choose the VMs you wish to clone to {0}" -f $ComputerName
try {
    $VMs = Get-VM -ErrorAction "Stop" | Out-GridView -Title $Title -PassThru
}
catch {
    Write-Error -ErrorRecord $_
    Write-Warning "Are you a member of the Hyper-V Administrators group or running the script as administrator?"
    Pause
}
if (-not $VMs) {
    Write-Warning ("Did not choose a VM to clone to {0}, quitting" -f $ComputerName)
    Pause
    return
}
#endregion

#region Determine which volume to copy the exported Hyper-V VMs to on the target host
$Title = "Choose the remote volume to store the VMs on '{0}'" -f $ComputerName
$TargetVolume = Get-Volume -CimSession $CimSession -ErrorAction "Stop" | Where-Object { $_.DriveLetter -match "\w" } | Out-GridView -Title $Title -OutputMode "Single"
if (-not $TargetVolume) {
    Write-Warning ("Did not choose a remote volume to store the VMs on '{0}', quitting" -f $ComputerName)
    Pause
    return
}
$TargetVolumeFreeGB = [Math]::Round($TargetVolume.SizeRemaining / 1GB, 2)
$TargetPath = '\\{0}\{1}$\VMs' -f $ComputerName, $TargetVolume.DriveLetter

if (-not (Test-Path $TargetPath -ErrorAction "Stop")) {
    $Message = "Target directory '{0}' does not exist, creating it" -f $TargetPath
    Write-Verbose $Message -Verbose
    $null = New-Item -Path $TargetPath -ItemType "Directory" -ErrorAction "Stop"
}
#endregion

#region Determine if the target volume on the remote host has enough capacity to store all the VMs in scope for cloning
$AllVMsTotalBytes = Get-Item -Path $VMs.HardDrives.Path | Measure-Object -Property Length -Sum
$AllVMsTotalGB = [Math]::Round($AllVMsTotalBytes.Sum / 1GB, 2)
# The volume must have at least $StorageThresholdGB free after the copy
if ($AllVMsTotalGB -ge ($TargetVolumeFreeGB - $StorageThresholdGB)) {
    $Message = "The selected remote volume '{0}' is not big enough to store all the VMs. Please choose another volume with at least '{1}GB' free." -f $TargetPath, ($AllVMsTotalGB + $StorageThresholdGB)
    Write-Error $Message -Category "InvalidOperation" -ErrorAction "Stop"
}
#endregion

#region Prompt if selected VMs are running, not essential but probably safer
$RunningVMs = $VMs.Where{$_.State -eq "Running"}

if ($RunningVMs.Count -gt 0) {
    $Message = "The below VM(s) are still running, do you wish to shut them down?`n - {0}" -f [String]::Join("`n - ", $RunningVMs.Name)
    $Options = "&Yes", "&No"
    $Response = $Host.UI.PromptForChoice($null, $Message, $Options, 0)

    if ($Response -eq 0) {
        $Jobs = foreach ($RunningVM in $RunningVMs) {
            Start-Job -Name $RunningVM.Name -ArgumentList $RunningVM.Name -ScriptBlock { 
                param(
                    [String]$VMName
                )
                Stop-VM -Name $VMName -Force -ErrorAction "Stop"  
            }
        }

        do {
            $CompleteJobs = $Jobs.Where{ $_.State -match "Completed|Failed" }
            Write-Progress -Activity "Waiting for VMs to shut down" -PercentComplete (($CompleteJobs.Count / $Jobs.Count) * 100)
            Start-Sleep -Seconds 2
        } while ($CompleteJobs.Count -ne $Jobs.Count)

        $FailedJobs = $Jobs.Where{ $_.State -eq "Failed" }

        if ($FailedJobs.Count -gt 0) {
            $Message = "The below VM(s) did not shut down within reasonable time, do you wish to continue?`n{0}" -f [String]::Join("`n - ", $FailedJobs.Name)
            $Response = $Host.UI.PromptForChoice($null, $Message, $Options, 1)

            if ($Response -eq 1) {
                return
            }
        }
        
        # Remove the jobs, so we can create new ones for Export-VM using the same names - helpful for logging
        Remove-Job -Job $Jobs -ErrorAction "Stop"
    }
}
#endregion

#region Export the VMs
$Jobs = foreach ($VM in $VMs) {

    # Determine if the path already exists on the target Hyper-V host
    # You can apparently register two VMs, with the same name, in the same directory
    # Therefore to be safe, bail if the path already exists and notify the user
    $Path = "{0}\{1}" -f $TargetPath, $VM.Name
    if (Test-Path $Path) {
        Write-Warning ("Path '{0}' already exists, skipping VM '{1}'" -f $Path, $VM.Name)
        continue
    }

    Start-Job -Name $VM.Name -ArgumentList @{ VMName = $VM.Name; ExportPath = $TargetPath; JobId = $JobId } -ScriptBlock {
        param (
            [hashtable]$ht
        )
        Export-VM -Name $ht["VMName"] -Path $ht["ExportPath"] -CaptureLiveState "CaptureDataConsistentState" -ErrorAction "Stop"
    }
}

if ($Jobs) {
    do {
        $CompleteJobs = $Jobs.Where{ $_.State -match "Completed|Failed" }
        Write-Progress -Activity "Waiting for VMs to finish exporting" -PercentComplete (($CompleteJobs.Count / $Jobs.Count) * 100)
        Start-Sleep -Seconds 5
    } while ($CompleteJobs.Count -ne $Jobs.Count)    
}
else {
    return
}

$FailedJobs = $Jobs.Where{ $_.State -eq "Failed" }

if ($FailedJobs.Count -gt 0) {
    # Give the error records back to the user of what the issue was within the job(s)
    $FailedJobs | Receive-Job
    $Message = "The below VMs failed to export, quitting`n - {0}" -f [String]::Join("`n - ", $FailedJobs.Name)
    # Produce a terminating error
    Write-Error $Message -Category "InvalidOperation" -ErrorId ([String]::Join(",", $FailedJobs.Name)) -TargetObject ([String]::Join(",", $FailedJobs.Name)) -ErrorAction "Stop"
}
#endregion

#region Import only the exported VMs
$ImportJobs = foreach ($Job in $Jobs) {

    Start-Job -Name $Job.Name -ArgumentList @{ VMName = $Job.Name; DriveLetter = $TargetVolume.DriveLetter; ComputerName = $ComputerName } -ScriptBlock {
        param (
            [hashtable]$ht
        )

        $CimSessionOpts = New-CimSessionOption -Protocol "DCOM"
        $CimSession = New-CimSession -ComputerName $ht["ComputerName"] -SessionOption $CimSessionOpts -ErrorAction "Stop"

        $VMRootPath = "{0}:\VMs\{1}" -f $ht["DriveLetter"], $ht["VMName"]
        $VMPath = "{0}\Virtual Machines" -f $VMRootPath
        $VMCXPath = "{0}\Virtual Machines\{1}.vmcx" -f $VMRootPath, (Get-VM -Name $ht["VMName"]).VMId
        $Params = @{
            Path = $VMCXPath
            VhdDestinationPath = "{0}\Virtual Hard Disks" -f $VMRootPath
            VirtualMachinePath = $VMPath
            SnapshotFilePath = "{0}\Snapshots" -f $VMRootPath
            Copy = $true
            GenerateNewId = $true
            # AsJob = $true
            CimSession = $CimSession
            ErrorAction = "Stop"
        }
    
        Import-VM @Params
    }

    # For some reason, the below didn't work. Literally given:
    #   Cannot load a virtual machine configuration: Unspecified error (0x80004005)
    #Import-VM -Path $VMCXPath -ErrorAction "Stop" # -GenerateNewId
}

if ($ImportJobs) {
    do {
        $CompleteJobs = $ImportJobs.Where{ $_.State -match "Completed|Failed" }
        Write-Progress -Activity "Waiting for VMs to finish importing" -PercentComplete (($CompleteJobs.Count / $ImportJobs.Count) * 100)
        Start-Sleep -Seconds 2
    } while ($CompleteJobs.Count -ne $ImportJobs.Count)
}
#endregion

#region Show results
# Controlled output to the console: show error records from jobs (if any) and then produce the pscustomobject giving overall result
$FailedJobs = @()

$Result = foreach ($Job in $ImportJobs) {
    [PSCustomObject]@{
        VM = $Job.Name
        Result = $Job.State
    }

    if ($Job.State -eq "Failed") { 
        # So shoot me
        $FailedJobs += $Job
    }
}

if ($FailedJobs.Count -gt 0) {
    Write-Output $FailedJobs | Receive-Job
} 

Write-Output $Result
#endregion

Remove-CimSession $CimSession -ErrorAction "Stop"
Get-Job | Remove-Job -ErrorAction "Stop"
Pause
