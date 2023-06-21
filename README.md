# Capa_Set_Units_Inactive

PowerShell script to set units inactive in CapaInstaller Console after x-days where the agen haven't executed and moves the units to a folder specified by the parameter `$Script:InactiveFolderStructure`.

If the parameter `$Script:InactiveActiveFolderStructure` in not `$null` units who wil become active again will be moved to the folder specified by the parameter `$Script:InactiveActiveFolderStructure`.

If the parameter `$Script:InactiveLastRunTimeFolderStructure` is not `$null`, units who have not executed for x-days but have newer run time than x-days will be moved to the folder specified by the parameter `$Script:InactiveLastRunTimeFolderStructure`. A little "good bug" is if you have multiple units with the same name, after x-days whre one of the units have not executed, both units will be moved to the folder specified by the parameter.

## Used modules 📦

 - [Capa.PowerShell.Module.SDK.Authentication](https://github.com/Mark5900/Capa.PowerShell.Module/tree/main/Modules/Capa.PowerShell.Module.SDK.Authentication)
 - [Capa.PowerShell.Module.SDK.Unit](https://github.com/Mark5900/Capa.PowerShell.Module/tree/main/Modules/Capa.PowerShell.Module.SDK.Unit)
 - [Mark.Boennelykke](https://github.com/Mark5900/Mark.Boennelykke)

### How to install the modules

At the path `C:\Program Files\PowerShell\Modules` creat a folder named `Capa.PowerShell.Module.SDK.Authentication`, `Capa.PowerShell.Module.SDK.Unit` and `Mark.Boennelykke` and copy the .psm1 & .psd1 files into their respective folders.

## How to use 📋

### [Capa_Set_Units_Inactive.ps1](Capa_Set_Units_Inactive.ps1)

Create a task in Windows Task Scheduler to run the script every day. The task need to run as an user with proper rights to the CapaInstaller Console.

* NOTE: The script needs to be run as administrator.
* NOTE: Remember to change parameters in the script to match your environment.
* NOTE: Look throug the script for #TODO

### [ServerTask_Capa_Set_Units_Inactive.ps1](ServerTask_Capa_Set_Units_Inactive.ps1)

Create a new PowerPack packaget and copy the script into the install script.

* NOTE: Remember to change parameters in the script to match your environment.
* NOTE: Look throug the script for #TODO
