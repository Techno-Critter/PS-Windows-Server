#
Author: Stan Crider
Date: 2May2017
What this crap does:
Reports back drive size and usage for each hard drive of a specified computer
Script will ping for connectivity and use WMI to check for properties
#>

# Enter computer name
$Computer = "computer01"

#===============================================================================================
# Function: Change data sizes to legible values; converts number to string
#===============================================================================================
Function Get-DataSize($FSize){
    If($FSize -lt 1KB){
        $FValue =  "$FSize B"
    }
    ElseIf(($FSize -ge 1KB) -and ($FSize -lt 1MB)){
        $FValue = "{0:N2}" -f ($FSize/1KB) + " KB"
    }
    ElseIf(($FSize -ge 1MB) -and ($FSize -lt 1GB)){
        $FValue = "{0:N2}" -f ($FSize/1MB) + " MB"
    }
    ElseIf(($FSize -ge 1GB) -and ($FSize -lt 1TB)){
        $FValue = "{0:N2}" -f ($FSize/1GB) + " GB"
    }
    ElseIf(($FSize -ge 1TB) -and ($FSize -lt 1PB)){
        $FValue = "{0:N2}" -f ($FSize/1TB) + " TB"
    }
    Else{
        $FValue = "{0:N2}" -f ($FSize/1PB) + " PB"
    }
    $FValue
}

#===============================================================================================
# Script: Call computer and retrieve WMI data
#===============================================================================================
If(Test-Connection $Computer -Quiet){
    $Drives = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter "DriveType='3'"
    $OpSys = Get-WmiObject Win32_Operatingsystem -ComputerName $Computer
    $WriteLine = ("=" * 95)
    $DriveList = @()

    ForEach($Drive in $Drives){

        If($Drive.DeviceID -eq $OpSys.SystemDrive){
            $SysDrive = $true
        }
        Else{
            $SysDrive = $false
        }
        $DriveList += $Drive | Select DeviceID, VolumeName,
            @{Name = 'DriveSize'; Expression = {(Get-DataSize $_.Size)}},
            @{Name = 'FreeSpace'; Expression = {(Get-DataSize $_.FreeSpace)}},
            @{Name = 'PercentFree'; Expression = {("{0:P2}" -f ($_.FreeSpace/$_.Size))}},
            @{Name = 'UsedSpace'; Expression = {(Get-DataSize ($_.Size - $_.FreeSpace))}},
            @{Name = 'PercentUsed'; Expression = {("{0:P2}" -f (($_.Size - $_.FreeSpace)/$_.Size))}},
            @{Name = 'SystemDrive'; Expression = {$SysDrive}}
    }

    Write $WriteLine
    Write-Host("Drive properties for computer: " + $Computer) -ForegroundColor Green
    Write $WriteLine
    $DriveList | FT
    Write $WriteLine
}
Else{
    Write $WriteLine
    Write ("Computer " + $Computer + " is not available. Script is terminated.")
    Write $WriteLine
}
