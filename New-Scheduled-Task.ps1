<#
Author: Stan Crider
Date: 8Mar2019
What this crap does:
Creates a one-time scheduled task on selected servers. If task exists, delete and re-register.
### Minimum Windows Server version is Server 2012! Will not work on 2008R2 or below! ###
Reference:
https://docs.microsoft.com/en-us/powershell/module/scheduledtasks/?view=win10-ps
#>

# User Input below
$ADOU = "DC=acme,DC=com"
$TaskName = "Scripted Server Reboot"
$Action = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /f /t 010 /d P:2:4"
$ExecuteTime = "2019-03-09 11:00pm"
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest -LogonType ServiceAccount

# Script below
Import-Module ActiveDirectory

$Servers = Get-ADComputer -SearchBase $ADOU -Filter * -Properties Name,Operatingsystem | Where-Object{$_.OperatingSystem -match "Windows Server"} | Sort-Object Name

ForEach($Server in $Servers){
    If(Test-Connection $Server.Name -Quiet){
        Invoke-Command -ComputerName $Server.Name -ScriptBlock {
            $TaskExists = Get-ScheduledTask | Where-Object{$_.TaskName -eq $Using:TaskName}
            $Trigger = New-ScheduledTaskTrigger -Once -At $Using:ExecuteTime
            $Task = New-ScheduledTask -Action $Using:Action -Trigger $Trigger -Principal $Using:Principal
            If($TaskExists){
                Unregister-ScheduledTask -TaskName $Using:TaskName -Confirm:$false
            }
            Register-ScheduledTask -TaskName $Using:TaskName -InputObject $Task
        }
    }
}
