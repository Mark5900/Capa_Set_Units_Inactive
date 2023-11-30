Import-Module Capa.PowerShell.Module.SDK.Authentication
Import-Module Capa.PowerShell.Module.SDK.Unit
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
        This script will set units inactive if they have not been executed for a given amount of days.
        If desired units that have become active again will be moved to a folder structure.
        Likewise units that have not been executed for a given amount of days but have an run time newer than the given amount of days will be moved to a folder structure,
        this will also apply to units where the name is reused (a good bug).

        REMEMBER TO CHANGE THE PARAMETERS BELOW AND LOOK AT THE TODO'S
#>
##################
### PARAMETERS ###
##################
# DO NOT CHANGE
$ScriptName = 'Capa_SetUnitsInactive'
$Script:ScriptFailed = 0
# Change as needed
#TODO: $CapaServer = ''
#TODO: $Database = ''
#TODO: $InstanceManagementPoint = '2'
$Script:LastRunDate = 90

$Script:Test = $false # If $true no changes will be made

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
            #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "The unit $($Unit.Name) has not been executed for $Script:LastRunDate days but the last run time is $LastRunTime" -ForegroundColor Yellow

            if ($Test -ne $true) {
                [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.UUID -UnitType Computer -FolderStructure $Script:InactiveLastRunTimeFolderStructure -CreateFolder true
                if ($Status -ne $true) {
                    #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "Failed to add unit $($Unit.Name) to folder $Script:InactiveLastRunTimeFolderStructure" -ForegroundColor Red
                    $Script:ScriptFailed = 1
                }
            } 

            $InactiveLastRunTimeCount++
        } elseif ($UnitLastExecuted -lt $Date) {
            #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "Unit to set inactive: $($Unit.Name) - LastRunTime: $LastRunTime - UnitLastExecuted: $UnitLastExecuted" -ForegroundColor Red

            if ($Test -ne $true) {
                [bool]$Status = Set-CapaUnitStatus -CapaSDK $oCMS -UnitName $Unit.Name -Status Inactive
                if ($Status -ne $true) {
                    #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "Failed to set unit $($Unit.Name) to inactive" -ForegroundColor Red
                    $Script:ScriptFailed = 1
                }

                if ($UnitFolder -ne "$Script:InactiveFolderStructure\") {
                    [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.UUID -UnitType Computer -FolderStructure $Script:InactiveFolderStructure -CreateFolder true
                    if ($Status -ne $true) {
                        #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "Failed to add unit $($Unit.Name) to folder $Script:InactiveFolderStructure" -ForegroundColor Red
                        $Script:ScriptFailed = 1
                    }
                }
            }

            $InactiveCount++
        } elseif ($null -ne $Script:InactiveActiveFolderStructure -and $UnitFolder -eq "$Script:InactiveFolderStructure\" -and $Unit.Status -eq 'Active') {
            #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "The unit $($Unit.Name) has become active again" -ForegroundColor Magenta

            if ($Test -ne $true) {
                [bool]$Status = Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.UUID -UnitType Computer -FolderStructure $Script:InactiveActiveFolderStructure -CreateFolder true
                if ($Status -ne $true) {
                    #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "Failed to add unit $($Unit.Name) to folder $Script:InactiveActiveFolderStructure" -ForegroundColor Red
                    $Script:ScriptFailed = 1
                }
            }

            $InactiveActiveCount++
        }
        #endregion
    }

    #TODO: ITCE-WriteLogLine -ScriptPart $ScriptPart -Text "Inactive: $InactiveCount - InactiveActive: $InactiveActiveCount - InactiveLastRunTime: $InactiveLastRunTimeCount"
}

##############
### SCRIPT ###
##############

#TODO: ITCE-StartScriptLoggin -Path $PSScriptRoot -LogName $ScriptName -DeleteDaysOldLogs 30

try {
    $oCMS = Initialize-CapaSDK -Server $CapaServer -Database $Database -InstanceManagementPoint $InstanceManagementPoint
    Set-UnitInactive -oCMS $oCMS
} catch {
    $Error[0]
    $Script:ScriptFailed = 1
}

If ($Global:ItceFejl -ne 0) {
    $Text = "Script $ScriptName failed, check the log for more information on server $$($env:COMPUTERNAME)"
    if ($Test -ne $true) {
        #TODO: ITCE-SendMail -Subject $Text -Body $Text -To $MailTo
    }
}

#TODO: ITCE-StopScriptLoggin
