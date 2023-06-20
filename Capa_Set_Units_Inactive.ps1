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
$Global:ScriptFailed = 0
# Change as needed
$CapaServer = 'CISRVKURSUS'
$Database = 'CapaInstaller'
$DefaultManagementPointDev = '1'
$DefaultManagementPointProd = $null #Keep null if you don't have two enviroments
$Global:LastRunDate = 90
$Global:InactiveFolderName = "Inaktiv i mindst $Global:LastRunDate dage"

#################
### FUNCTIONS ###
#################
function Set-UnitInactive {
    param (
        $oCMS
    )
    $Date = (Get-Date).AddDays(-$Global:LastRunDate)

    $AllUnits = Get-CapaUnits -CapaSDK $oCMS -Type Computer
    foreach ($Unit in $AllUnits) {
        $LastRunTime = (Get-CapaUnitLastRuntime -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer).Split(' ')
        $LastRunTime = $LastRunTime[0].Split('-')
        $LastRunTime = Get-Date -Day $LastRunTime[0] -Month $LastRunTime[1] -Year $LastRunTime[2]
        
        $UnitLastExecuted = $Unit.LastExecuted.Split(' ')
        $UnitLastExecuted = $UnitLastExecuted[0].Split('-')
        $UnitLastExecuted = Get-Date -Day $UnitLastExecuted[0] -Month $UnitLastExecuted[1] -Year $UnitLastExecuted[2]
        
        if ($LastRunTime -lt $Date -or $UnitLastExecuted -lt $Date) {
            Write-MBLogLine "Unit to set inactive: $($Unit.Name) - LastRunTime: $LastRunTime - UnitLastExecuted: $UnitLastExecuted"
            Set-CapaUnitStatus -CapaSDK $oCMS -UnitName $Unit.Name -Status Inactive | Out-Null
            Add-CapaUnitToFolder -CapaSDK $oCMS -UnitName $Unit.Name -UnitType Computer -FolderStructure $Global:InactiveFolderName -CreateFolder true | Out-Null
        }
    }
}

##############
### SCRIPT ###
##############

Start-MBScriptLogging -Path $PSScriptRoot -LogName $ScriptName -DeleteDaysOldLogs 30

try{
If ($null -eq $DefaultManagementPointProd){
    $oCMSDev = Initialize-CapaSDK -Server $CapaServer -Database $Database
    $oCMSProd = $oCMSDev
}
else{
    $oCMSDev = Initialize-CapaSDK -Server $CapaServer -Database $Database -DefaultManagementPoint $DefaultManagementPointDev
    $oCMSProd = Initialize-CapaSDK -Server $CapaServer -Database $Database -DefaultManagementPoint $DefaultManagementPointProd
}
}catch {
    $Error[0]
    $Global:ScriptFailed = 1
}

If ($Global:ScriptFailed -ne 0) {
    $Text = "Script $ScriptName failed, check the log for more information on server $($env:COMPUTERNAME)"
    if ($Test -ne $true) {
        #TODO: Send mail with error message
    }
}

Stop-MBScriptLogging