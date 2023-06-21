Import-Module Capa.PowerShell.Module.SDK
Import-Module Mark.Boennelykke
<#
    .NOTES
    ===========================================================================
    Created with:	Visual Studio Code
    Created on:     20-06-2023
    Created by:		Mark5900
    Organization:	
    Filename:       Capa_Set_Units_Inactive.ps1
    ===========================================================================
    .DESCRIPTION
        TODO: A description of the file.
#>
##################
### PARAMETERS ###
##################
# DO NOT CHANGE
$ScriptName = 'Capa_Set_Units_Inactive'
$Script:ScriptFailed = 0
# Change as needed
$CapaServer = 'CISRVKURSUS'
$Database = 'CapaInstaller'
$DefaultManagementPointDev = '1'
$DefaultManagementPointProd = $null #Keep null if you don't have two enviroments
$Script:LastRunDate = 90

$Script:InactiveFolderStructure = "Inaktiv i mindst $Script:LastRunDate dage"
$Script:InactiveActiveFolderStructure = "$Script:InactiveFolderStructure\Aktive igen" # If $null nothing happens else will move the units to the folder structure spiciefied
$Script:InactiveLastRunTimeFolderStructure = "$Script:InactiveFolderStructure\Bad agent" # If $null nothing happens else will move the units to the folder structure spiciefied

#################
### FUNCTIONS ###
#################
function Set-UnitInactive {
    param (
        $oCMS
    )
    $ScriptPart = 'Set-UnitInactive'
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
            Write-MBLogLine -ScriptPart $ScriptPart -Text "The unit $($Unit.Name) has not been executed for $Script:LastRunDate days but the last run time is $LastRunTime"

            if ($Status -ne $true) {
                Write-MBLogLine -ScriptPart $ScriptPart -Text "Failed to add unit $($Unit.Name) to folder $Script:InactiveActiveFolderStructure" -ForegroundColor Red
                $Script:ScriptFailed = 1
            }
        } elseif ($UnitLastExecuted -lt $Date) {
            Write-MBLogLine -ScriptPart $ScriptPart -Text "Unit to set inactive: $($Unit.Name) - LastRunTime: $LastRunTime - UnitLastExecuted: $UnitLastExecuted"

            [bool]$Status = Set-CapaUnitStatus -CapaSDK $oCMS -UnitName $Unit.Name -Status Inactive
            if ($Status -ne $true) {
                Write-MBLogLine -ScriptPart $ScriptPart -Text "Failed to set unit $($Unit.Name) to inactive" -ForegroundColor Red
                $Script:ScriptFailed = 1
            }

            if ($UnitFolder -ne "$Script:InactiveFolderStructure\") {
                [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer -FolderStructure $Script:InactiveFolderStructure -CreateFolder true
                if ($Status -ne $true) {
                    Write-MBLogLine -ScriptPart $ScriptPart -Text "Failed to add unit $($Unit.Name) to folder $Script:InactiveFolderStructure" -ForegroundColor Red
                    $Script:ScriptFailed = 1
                }
            }
        } elseif ($null -ne $Script:InactiveActiveFolderStructure -and $UnitFolder -eq "$Script:InactiveFolderStructure\" -and $Unit.Status -eq 'Active') {
            Write-MBLogLine -ScriptPart $ScriptPart -Text "The unit $($Unit.Name) has become active again"

            [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer -FolderStructure $Script:InactiveActiveFolderStructure -CreateFolder true
            if ($Status -ne $true) {
                Write-MBLogLine -ScriptPart $ScriptPart -Text "Failed to add unit $($Unit.Name) to folder $Script:InactiveActiveFolderStructure" -ForegroundColor Red
                $Script:ScriptFailed = 1
            }
        }
        #endregion
    }
}

##############
### SCRIPT ###
##############

Start-MBScriptLogging -Path $PSScriptRoot -LogName $ScriptName -DeleteDaysOldLogs 30

try {
    If ($null -eq $DefaultManagementPointProd) {
        $oCMSDev = Initialize-CapaSDK -Server $CapaServer -Database $Database
        $oCMSProd = $oCMSDev
        Set-UnitInactive -oCMS $oCMSDev
    } else {
        $oCMSDev = Initialize-CapaSDK -Server $CapaServer -Database $Database -DefaultManagementPoint $DefaultManagementPointDev
        $oCMSProd = Initialize-CapaSDK -Server $CapaServer -Database $Database -DefaultManagementPoint $DefaultManagementPointProd

        Set-UnitInactive -oCMS $oCMSDev
        Set-UnitInactive -oCMS $oCMSProd
    }
} catch {
    $Error[0]
    $Script:ScriptFailed = 1
}

If ($Script:ScriptFailed -ne 0) {
    $Text = "Script $ScriptName failed, check the log for more information on server $($env:COMPUTERNAME)"
    Write-MBLogLine -ScriptPart $ScriptName -Text $Text -ForegroundColor Red
    if ($Test -ne $true) {
        #TODO: Send mail with error message
    }
}

Stop-MBScriptLogging