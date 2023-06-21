[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [string]$Packageroot,
    [Parameter(Mandatory = $true)]
    [string]$AppName,
    [Parameter(Mandatory = $true)]
    [string]$AppRelease,
    [Parameter(Mandatory = $true)]
    [string]$LogFile,
    [Parameter(Mandatory = $true)]
    [string]$TempFolder,
    [Parameter(Mandatory = $true)]
    [string]$DllPath,
    [Parameter(Mandatory = $false)]
    [Object]$InputObject = $null
)

# DO NOT CHANGE
$Script:ScriptFailed = 0
# Change as needed
$CapaServer = 'CISRVKURSUS'
$Database = 'CapaInstaller'
$Script:LastRunDate = 90

$Script:Test = $true # If $true no changes will be made

$Script:InactiveFolderStructure = "Inaktiv i mindst $Script:LastRunDate dage"
$Script:InactiveActiveFolderStructure = "$Script:InactiveFolderStructure\Aktive igen" # If $null nothing happens else will move the units to the folder structure spiciefied
$Script:InactiveLastRunTimeFolderStructure = "$Script:InactiveFolderStructure\Bad agent" # If $null nothing happens else will move the units to the folder structure spiciefied

function Set-UnitInactive {
    param (
        $oCMS
    )
    $cs.Log_SectionHeader('Set-UnitInactive', 'o')

    $InactiveCount = 0
    $InactiveActiveCount = 0
    $InactiveLastRunTimeCount = 0
    $Date = (Get-Date).AddDays(-$Script:LastRunDate)

    $AllUnits = Get-CapaUnits -CapaSDK $oCMS -Type Computer
    foreach ($Unit in $AllUnits) {
        #region Units to skip
        if ($Unit.IsMobileDevice -eq $true) {
            # Skip mobile devices
            continue
        }
        #endregion

        #region Variables
        $LastRunTime = (Get-CapaUnitLastRuntime -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer).Split(' ')
        $LastRunTime = $LastRunTime[0].Split('-')
        $LastRunTime = Get-Date -Day $LastRunTime[0] -Month $LastRunTime[1] -Year $LastRunTime[2]
        
        $UnitLastExecuted = $Unit.LastExecuted.Split(' ')
        $UnitLastExecuted = $UnitLastExecuted[0].Split('-')
        $UnitLastExecuted = Get-Date -Day $UnitLastExecuted[0] -Month $UnitLastExecuted[1] -Year $UnitLastExecuted[2]

        $UnitFolder = Get-CapaUnitFolder -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer
        #endregion

        #region Code
        if ($null -ne $Script:InactiveLastRunTimeFolderStructure -and $UnitLastExecuted -lt $Date -and $LastRunTime -gt $Date) {
            $cs.Job_WriteLog("The unit $($Unit.Name) has not been executed for $Script:LastRunDate days but the last run time is $LastRunTime")

            if ($Test -ne $true) {
                [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer -FolderStructure $Script:InactiveLastRunTimeFolderStructure -CreateFolder true
                if ($Status -ne $true) {
                    $cs.Job_WriteLog("Failed to add unit $($Unit.Name) to folder $Script:InactiveLastRunTimeFolderStructure")
                    
                    $Script:ScriptFailed = 1
                }
            }

            $InactiveLastRunTimeCount++
        } elseif ($UnitLastExecuted -lt $Date) {
            $cs.Job_WriteLog("Unit to set inactive: $($Unit.Name) - LastRunTime: $LastRunTime - UnitLastExecuted: $UnitLastExecuted")

            if ($Test -ne $true) {
                [bool]$Status = Set-CapaUnitStatus -CapaSDK $oCMS -UnitName $Unit.Name -Status Inactive
                if ($Status -ne $true) {
                    $cs.Job_WriteLog("Failed to set unit $($Unit.Name) to inactive")

                    $Script:ScriptFailed = 1
                }

                if ($UnitFolder -ne "$Script:InactiveFolderStructure\") {
                    [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer -FolderStructure $Script:InactiveFolderStructure -CreateFolder true
                    if ($Status -ne $true) {
                        $cs.Job_WriteLog("Failed to add unit $($Unit.Name) to folder $Script:InactiveFolderStructure")
                        $Script:ScriptFailed = 1
                    }
                }
            }

            $InactiveCount++
        } elseif ($null -ne $Script:InactiveActiveFolderStructure -and $UnitFolder -eq "$Script:InactiveFolderStructure\" -and $Unit.Status -eq 'Active') {
            $cs.Job_WriteLog("The unit $($Unit.Name) has become active again")

            if ($Test -ne $true) {
                [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer -FolderStructure $Script:InactiveActiveFolderStructure -CreateFolder true
                if ($Status -ne $true) {
                    $cs.Job_WriteLog("Failed to add unit $($Unit.Name) to folder $Script:InactiveActiveFolderStructure")
                    $Script:ScriptFailed = 1
                }
            }

            $InactiveActiveCount++
        }
        #endregion
    }

    $cs.Job_WriteLog("Inactive: $InactiveCount - InactiveActive: $InactiveActiveCount - InactiveLastRunTime: $InactiveLastRunTimeCount")
}

try {
    ### Download package kit
    [bool]$global:DownloadPackage = $true

    ##############################################
    #load core PS lib - don't mess with this!
    if ($InputObject) { $pgkit = '' }else { $pgkit = 'kit' }
    Import-Module (Join-Path $Packageroot $pgkit 'PSlib.psm1') -ErrorAction stop
    #load Library dll
    $cs = Add-PSDll
    ##############################################

    #Begin
    $cs.Job_Start('WS', $AppName, $AppRelease, $LogFile, 'INSTALL')
    $cs.Job_WriteLog("[Init]: Starting package: '" + $AppName + "' Release: '" + $AppRelease + "'")
    if (!$cs.Sys_IsMinimumRequiredDiskspaceAvailable('c:', 1500)) { Exit-PSScript 3333 }
    if ($global:DownloadPackage -and $InputObject) { Start-PSDownloadPackage }

    $cs.Job_WriteLog("[Init]: `$PackageRoot:` '" + $Packageroot + "'")
    $cs.Job_WriteLog("[Init]: `$AppName:` '" + $AppName + "'")
    $cs.Job_WriteLog("[Init]: `$AppRelease:` '" + $AppRelease + "'")
    $cs.Job_WriteLog("[Init]: `$LogFile:` '" + $LogFile + "'")
    $cs.Job_WriteLog("[Init]: `$TempFolder:` '" + $TempFolder + "'")
    $cs.Job_WriteLog("[Init]: `$DllPath:` '" + $DllPath + "'")
    $cs.Job_WriteLog("[Init]: `$global:DownloadPackage`: '" + $global:DownloadPackage + "'")

    $cs.Log_SectionHeader('Import custom modules', 'o')
    Import-Module (Join-Path $Packageroot $pgkit 'Capa.PowerShell.Module.SDK.Authentication') -ErrorAction stop
    Import-Module (Join-Path $Packageroot $pgkit 'Capa.PowerShell.Module.SDK.Unit') -ErrorAction stop

    $cs.Log_SectionHeader('Initialize Capa SDK', 'o')
    $cs.Job_WriteLog($CapaServer)
    $cs.Job_WriteLog($Database)
    $oCMS = Initialize-CapaSDK -Server $CapaServer -Database $Database

    Set-UnitInactive -oCMS $oCMS

    If ($Script:ScriptFailed -ne 0) {
        $Text = "Script $AppName $AppRelease failed, check the log for more information on server $($env:COMPUTERNAME)"
        $cs.Job_WriteLog($Text)
        if ($Test -ne $true) {
            #TODO: Send mail with error message
        }
    }

    Exit-PSScript $Error

} catch {
    $line = $_.InvocationInfo.ScriptLineNumber
    $cs.Job_WriteLog('*****************', "Something bad happend at line $($line): $($_.Exception.Message)")
    Exit-PSScript $_.Exception.HResult
}
