<#
Author: Stan Crider
Date: 21Nov2019
Crap:
Gets list of Hyper-V hosts and their VM's and outputs to Excel
Divides report into 7 worksheets
Host sheet: Physical host properties and number of VM's
Host CPU: Physical CPU properties and count
Host RAM: Physical RAM properties and count
Host HD: Each host hard drive letter, size and free space
VMs: VM properties and status
VM HD's: Each VM-HD properties, sizes and file location
Errors: Report unresponsive or in-error hosts
### Must have ImportExcel module installed! ###
### https://github.com/dfinke/ImportExcel  ###
#>

#Requires -Module ImportExcel

#region User specified variables
# Create Excel filename based on date
$DateName = Get-Date -Format yyyyMMdd
$LogFile = "C:\HyperV\HyperV Report $DateName.xlsx"

# Call AD and search for computers
$DomainName = "acme.com"
$ServerList = Get-ADComputer -Filter {Enabled -eq $true} -SearchBase "OU=Hypervisors,DC=acme,DC=com" -Server $DomainName -Properties IPv4Address,OperatingSystem | Sort-Object Name
#endregion

#region Do not overwrite existing logfiles
If(Test-Path $LogFile){
    Write-Output "The file $LogFile already exists. Script terminated."
    Return $null
}
#endregion

#region Functions
# Change data sizes to legible values; converts number to string
Function Get-Size($FSize){
    Switch($Fsize){
        {$_ -lt 1KB}{
            $FValue = "$FSize B"
        }
        {($_ -ge 1KB) -and ($_ -lt 1MB)}{
            $FValue = "{0:N2}" -f ($FSize/1KB) + " KiB"
        }
        {($_ -ge 1MB) -and ($_ -lt 1GB)}{
            $FValue = "{0:N2}" -f ($FSize/1MB) + " MiB"
        }
        {($_ -ge 1GB) -and ($_ -lt 1TB)}{
            $FValue = "{0:N2}" -f ($FSize/1GB) + " GiB"
        }
        {($_ -ge 1TB) -and ($_ -lt 1PB)}{
            $FValue = "{0:N2}" -f ($FSize/1TB) + " TiB"
        }
        Default{
            $FValue = "{0:N2}" -f ($FSize/1PB) + " PiB"
        }
    }
    $FValue
}

# Convert number of object items into Excel column headers
Function Get-ColumnName ([int]$ColumnCount){
    If(($ColumnCount -le 702) -and ($ColumnCount -ge 1)){
        $ColumnCount = [Math]::Floor($ColumnCount)
        $CharStart = 64
        $FirstCharacter = $null

        # Convert number into double letter column name (AA-ZZ)
        If($ColumnCount -gt 26){
            $FirstNumber = [Math]::Floor(($ColumnCount)/26)
            $SecondNumber = ($ColumnCount) % 26

            # Reset increment for base-26
            If($SecondNumber -eq 0){
                $FirstNumber--
                $SecondNumber = 26
            }

            # Left-side column letter (first character from left to right)
            $FirstLetter = [int]($FirstNumber + $CharStart)
            $FirstCharacter = [char]$FirstLetter

            # Right-side column letter (second character from left to right)
            $SecondLetter = $SecondNumber + $CharStart
            $SecondCharacter = [char]$SecondLetter

            # Combine both letters into column name
            $CharacterOutput = $FirstCharacter + $SecondCharacter
        }

        # Convert number into single letter column name (A-Z)
        Else{
            $CharacterOutput = [char]($ColumnCount + $CharStart)
        }
    }
    Else{
        $CharacterOutput = "ZZ"
    }

    # Output column name
    $CharacterOutput
}
#endregion

#region Prepare script
# Import modules
Import-Module Hyper-V,ActiveDirectory

# Create Arrays
$HostArray = @()
$HostCPUArray = @()
$HostRAMArray = @()
$HostHDArray = @()
$VMArray = @()
$VMHDArray = @()
$ErrorArray = @()
#endregion

ForEach($Server in $ServerList){
    Write-Output ("[$(Get-Date)] Processing $($Server.DNSHostName) ...")
    #region Reset host variables
    $HostCPUCores = 0
    $HostTotalHDSpace = 0
    $HostTotalHDFreeSpace = 0
    $TotalVMAllocatedRAM = 0
    $TotalVMAllocatedHDSpace = 0
    #$HostADProps = $null
    $Hostmanufacturer = $null
    $HostBIOS = $null
    $HostOSProps = $null
    $HostProcProps = $null
    $HostRAMModules = $null
    $HostHardDrives = $null
    $HostHyperVProps = $null
    $VirtualMachines = $null
    $TimeZone = $null
    $CurrentTime = $null
    #endregion

    Try{
        $CimSession = New-CimSession -ComputerName $Server.DNSHostName -ErrorAction Stop
    }
    Catch{
        $CimSession = $null
        $ErrorArray += [PSCustomObject]@{
            "Host Name" = $Server.DNSHostName
            "Error"     = $_.Exception.Message
        }
        Continue
    }

    Try{
        $PSSession = New-PSSession -ComputerName $Server.DNSHostName -ErrorAction Stop
    }
    Catch{
        $PSSession = $null
        $ErrorArray += [PSCustomObject]@{
            "Host Name" = $Server.DNSHostName
            "Error"     = $_.Exception.Message
        }
        Continue
    }

    Try{
        $HVRole = Invoke-Command -Session $PSSession -ScriptBlock {Get-WindowsFeature -Name Hyper-V} -ErrorAction Stop
    }
    Catch{
        $HVRole = $null
        $ErrorArray += [PSCustomObject]@{
            "Host Name" = $Server.DNSHostName
            "Error"     = $_.Exception.Message
        }
        Continue
    }

    If(-Not $HVRole.Installed){
        Continue
    }

    Try{
        $HVPSRole = Invoke-Command -Session $PSSession -ScriptBlock {Get-WindowsFeature -Name Hyper-V-PowerShell} -ErrorAction Stop
    }
    Catch{
        $HVPSRole = $null
        $ErrorArray += [PSCustomObject]@{
            "Host Name" = $Server.DNSHostName
            "Error"     = $_.Exception.Message
        }
        Continue
    }

    If(-Not $HVPSRole.Installed){
        $ErrorArray += [PSCustomObject]@{
            "Host Name" = $Server.DNSHostName
            "Error"     = 'Hyper-V-PowerShell feature not installed'
        }
        Continue
    }

    #region Gather host data
    $HostManufacturer = Get-CimInstance -Class Win32_ComputerSystem -CimSession $CimSession
    $TimeZone = Get-CimInstance -Class Win32_TimeZone -CimSession $CimSession
    $CurrentTime = Invoke-Command -Session $PSSession -ScriptBlock {Get-Date -Format "MM/dd/yyyy HH:mm:ss"}
    $HostBIOS = Get-CimInstance -Class Win32_BIOS -CimSession $CimSession
    $HostOSProps = Get-CimInstance -Class Win32_OperatingSystem -CimSession $CimSession
    $HostProcProps = Get-CimInstance -Class Win32_Processor -CimSession $CimSession
    $HostRAMModules = Get-CimInstance -Class Win32_PhysicalMemory -CimSession $CimSession
    $HostHardDrives = Get-CimInstance -Class Win32_LogicalDisk -CimSession $CimSession -Filter "DriveType = 3"
    $HostHyperVProps = Invoke-Command -Session $PSSession {Get-VMHost | Select-Object *}
    $VirtualMachines = Invoke-Command -Session $PSSession {Get-VM | Select-Object *}
    #endregion

    #region Host CPU Properties
    ForEach($HostProc in $HostProcProps){
        $HostCPUCores = ($HostProc.NumberOfCores + $HostCPUCores)
        $HostCPUArray += [PSCustomObject]@{
            "Host Name"  = $Server.DNSHostName
            "CPU"        = $HostProc.Name
            "Speed"      = $HostProc.MaxClockSpeed
            "Socket"     = $HostProc.SocketDesignation
            "Core Count" = $HostProc.NumberOfCores
        }
    }
    #endregion

    #region Host RAM Properties
    ForEach($HostRAMModule in $HostRAMModules){
        $HostRAMArray += [PSCustomObject]@{
            "Host Name" = $Server.DNSHostName
            "RAM Bank"  = $HostRAMModule.DeviceLocator
            "Size"      = Get-Size $HostRAMModule.Capacity
            "Bus Speed" = $HostRAMModule.Speed
        }
    }
    #endregion

    #region Host hard drives
    ForEach($HostHD in $HostHardDrives){
        $HostHDArray += [PSCustomObject]@{
            "Host Name" = $Server.DNSHostName
            "Drive"     = $HostHD.DeviceID
            "Size"      = Get-Size $HostHD.Size
            "Free"      = Get-Size $HostHD.FreeSpace
            "% Free"    = ("{0:N4}" -f ($HostHD.Freespace/$HostHD.Size))
        }
        $HostTotalHDSpace += $HostHD.Size
        $HostTotalHDFreeSpace += $HostHD.FreeSpace
    }
    #Endregion

    #region Virtual Machines
    ForEach($VirtualMachine in $VirtualMachines){
        # Reset VM variables
        $TotHDSpace = 0
        $TotHDUsedSpace = 0
        $VMIPAddress = $null
        # Gather VM data
        $VMIntegratedServices = Invoke-Command  -Session $PSSession -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMIntegrationService -VMName $VirtualMachine} -ErrorAction Stop
        $VMProc = Invoke-Command  -Session $PSSession -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMProcessor -VMName $VirtualMachine} -ErrorAction Stop
        $VMMemory = Invoke-Command  -Session $PSSession -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMMemory -VMName $VirtualMachine} -ErrorAction Stop
        $VMHardDrives = Invoke-Command  -Session $PSSession -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMHardDiskDrive -VMName $VirtualMachine} -ErrorAction Stop
        $SnapShots = Invoke-Command  -Session $PSSession -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMSnapshot -VMName $VirtualMachine} -ErrorAction Stop
        If($VirtualMachine.State.ToString() -eq "Running"){
            $VMIPAddress = Invoke-Command  -Session $PSSession -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMNetworkAdapter -VMName $VirtualMachine} -ErrorAction Stop
        }

        # Get VM Hard Drive data
        ForEach($VMHardDrive in $VMHardDrives){
            $VHDUsedPct = 0
            If($null -eq $VMHardDrive.DiskNumber){
                $VHDObj = Invoke-Command  -Session $PSSession -ArgumentList $VMHardDrive.Path {$VMHardDrivepath = [string]$Args[0]; Get-VHD -Path $VMHardDrivepath} -ErrorAction Stop
                $VHDType = $VHDObj.VhdType.ToString()
                $VHDUsedPct = $VHDObj.FileSize/$VHDObj.Size
            }
            Else{
                $VHDObj = $null
                $VHDType = "Physical"
            }
            $VMHDArray += [PSCustomObject]@{
                "VM"              = $VirtualMachine.Name
                "Host Name"       = $Server.DNSHostName
                "Used Space"      = Get-Size $VHDObj.FileSize
                "Drive Size"      = Get-Size $VHDObj.Size
                "% Used"          = ("{0:N4}" -f $VHDUsedPct)
                "Min Size"        = Get-Size $VHDObj.MinimumSize
                "Block"           = Get-Size $VHDObj.BlockSize
                "Logical Sector"  = Get-Size $VHDObj.LogicalSectorSize
                "Physical Sector" = Get-Size $VHDObj.PhysicalSectorSize
                "Controller"      = $VMHardDrive.ControllerType.ToString()
                "Controller No"   = $VMHardDrive.ControllerNumber
                "Controller ID"   = $VMHardDrive.ControllerLocation
                "VHD Type"        = $VHDType
                "HD File Path"    = $VMHardDrive.Path
            }
            $TotHDSpace += $VHDObj.Size
            $TotHDUsedSpace += $VHDObj.FileSize
        }


        $VMArray += [PSCustomObject]@{
            "VM"                    = $VirtualMachine.Name
            "State"                 = $VirtualMachine.State.ToString()
            "Host Name"             = $Server.DNSHostName
            "VM IP"                 = $VMIPAddress.IPAddresses -join ", "
            "Generation"            = $VirtualMachine.Generation
            "CPU"                   = $VMProc.Count
            "CPU Migration Enabled" = $VMProc.CompatibilityForMigrationEnabled
            "RAM"                   = Get-Size $VMMemory.Startup
            "Dynamic RAM"           = $VMMemory.DynamicMemoryEnabled
            "HD Count"              = ($VMHardDrives | Measure-Object).Count
            "Total HD Space"        = Get-Size $TotHDSpace
            "Total HD Used"         = Get-Size $TotHDUsedSpace
            "AutoStart"             = $VirtualMachine.AutomaticStartAction.ToString()
            "Time Sync"             = ($VMIntegratedServices | Where-Object{$_.Name -eq "Time Synchronization"}).Enabled
            "Heartbeat"             = ($VMIntegratedServices | Where-Object{$_.Name -eq "Heartbeat"}).Enabled
            "Data Exchange"         = ($VMIntegratedServices | Where-Object{$_.Name -eq "Key-Value Pair Exchange"}).Enabled
            "OS Shutdown"           = ($VMIntegratedServices | Where-Object{$_.Name -eq "Shutdown"}).Enabled
            "Backup"                = ($VMIntegratedServices | Where-Object{$_.Name -eq "VSS"}).Enabled
            "Guest Services"        = ($VMIntegratedServices | Where-Object{$_.Name -eq "Guest Service Interface"}).Enabled
            "Snapshots"             = ($SnapShots | Measure-Object).Count
            "Config File Path"      = $VirtualMachine.ConfigurationLocation
        }
        $TotalVMAllocatedRAM += $VMMemory.Startup
        $TotalVMAllocatedHDSpace += $TotHDSpace
    }
    #endregion

    #region Host Properties
    $HostArray += [PSCustomObject]@{
        "Host Name"      = $Server.DNSHostName
        "Host IP"        = $Server.IPv4Address
        "Host OS"        = $Server.OperatingSystem
        "Host CPUs"      = $HostProcProps.Count
        "Host Cores"     = $HostCPUCores
        "Host OS RAM"    = Get-Size $HostManufacturer.TotalPhysicalMemory
        "RAM Allocated"  = Get-Size $TotalVMAllocatedRAM
        "RAM Banks Used" = ($HostRAMModules | Measure-Object).Count
        "HD Count"       = ($HostHardDrives | Measure-Object).Count
        "HD Space"       = Get-Size $HostTotalHDSpace
        "HD Free"        = Get-Size $HostTotalHDFreeSpace
        "% Free"         = "{0:N4}" -f ($HostTotalHDFreeSpace/$HostTotalHDSpace)
        "HD Allocated"   = Get-Size $TotalVMAllocatedHDSpace
        "Total VMs"      = ($VirtualMachines | Measure-Object).Count
        "VMs On"         = ($VirtualMachines | Where-Object {$_.State.ToString() -eq "Running"} | Measure-Object).Count
        "Manufacturer"   = $HostManufacturer.Manufacturer
        "Model"          = $HostManufacturer.Model
        "Serial No"      = $HostBIOS.SerialNumber
        "Last Boot"      = ($HostOSProps.LastBootUpTime).ToString("d'/'MM'/'yyyy hh':'mm")
        "Uptime"         = ((Get-Date) - $HostOSProps.LastBootUpTime).ToString("d' days 'hh':'mm':'ss")
        "Current Time"   = $CurrentTime
        "Time Zone"      = $TimeZone.Caption
        "VM Drive Path"  = $HostHyperVProps.VirtualHardDiskPath
        "VM Config Path" = $HostHyperVProps.VirtualMachinePath
        "AD Description" = $Server.Description
    }

    Remove-CimSession -CimSession $CimSession
    Remove-PSSession -Session $PSSession
    #endregion
}

#region Export to file
# Create Excel standard configuration properties
$ExcelProps = @{
    Autosize = $true;
    FreezeTopRow = $true;
    BoldTopRow = $true;
}

$ExcelProps.Path = $LogFile

# Hosts sheet
$HostArrayLastRow = ($HostArray | Measure-Object).Count + 1
If($HostArrayLastRow -gt 1){
    $HostArrayHeaderCount = Get-ColumnName ($HostArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostArrayHeaderRow = "`$A`$1:`$$HostArrayHeaderCount`$1"
    $HostArrayPctFreeColumn = "'HypV Hosts'!`$L`$2:`$L`$$HostArrayLastRow"
    $TotalVMsColumn = "'HypV Hosts'!`$N`$2:`$N`$$HostArrayLastRow"
    $TotalVMsRunningColumn = "'HypV Hosts'!`$O`$2:`$O`$$HostArrayLastRow"
    $HostArrayStyle = @()
    $HostArrayStyle += New-ExcelStyle -Range "'HypV Hosts'$HostArrayHeaderRow" -HorizontalAlignment Center
    $HostArrayStyle += New-ExcelStyle -Range $HostArrayPctFreeColumn -NumberFormat '0.00%'
    $HostArrayConditionalText = @()
    $HostArrayConditionalText += New-ConditionalText -Range $HostArrayPctFreeColumn -ConditionalType LessThan 0.10 -ConditionalTextColor Maroon -BackgroundColor Pink
    $HostArrayConditionalText += New-ConditionalText -Range $HostArrayPctFreeColumn -ConditionalType LessThan 0.20 -ConditionalTextColor Brown -BackgroundColor Wheat
    $HostArrayConditionalText += New-ConditionalText -Range $TotalVMsColumn -ConditionalType GreaterThan 2 -ConditionalTextColor Brown -BackgroundColor Wheat
    $HostArrayConditionalText += New-ConditionalText -Range $TotalVMsRunningColumn -ConditionalType GreaterThan 2 -ConditionalTextColor Maroon -BackgroundColor Pink
    $HostArray | Sort-Object "Host Name" | Export-Excel @ExcelProps -WorksheetName "HypV Hosts" -Style $HostArrayStyle -ConditionalText $HostArrayConditionalText
}

# Host CPU sheet
$HostCPUArrayLastRow = ($HostCPUArray | Measure-Object).Count + 1
If($HostCPUArrayLastRow -gt 1){
    $HostCPUHeaderCount = Get-ColumnName ($HostCPUArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostCPUHeaderRow = "`$A`$1:`$$HostCPUHeaderCount`$1"
    $HostCPUArrayStyle = New-ExcelStyle -Range "'Host CPU'$HostCPUHeaderRow" -HorizontalAlignment Center    
    $HostCPUArray | Sort-Object "Host Name","Socket" | Export-Excel @ExcelProps -WorksheetName "Host CPU" -Style $HostCPUArrayStyle
}

# Host RAM sheet
$HostRAMArrayLastRow = ($HostRAMArray | Measure-Object).Count + 1
If($HostRAMArrayLastRow -gt 1){
    $HostRAMHeaderCount = Get-ColumnName ($HostRAMArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostRAMHeaderRow = "`$A`$1:`$$HostRAMHeaderCount`$1"
    $HostRAMArrayStyle = New-ExcelStyle -Range "'Host RAM'$HostRAMHeaderRow" -HorizontalAlignment Center
    $HostRAMArray | Sort-Object "Host Name","RAM Bank" | Export-Excel @ExcelProps -WorksheetName "Host RAM" -Style $HostRAMArrayStyle
}

# Host HD sheet
$HostHDArrayLastRow = ($HostHDArray | Measure-Object).Count + 1
If($HostHDArrayLastRow -gt 1){
    $HostHDHeaderCount = Get-ColumnName ($HostHDArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostHDHeaderRow = "`$A`$1:`$$HostHDHeaderCount`$1"
    $PctFreeColumn = "'Host HD'!`$E`2:`$E`$$HostHDArrayLastRow"
    $HostHDArrayStyle = @()
    $HostHDArrayStyle += New-ExcelStyle -Range "'Host HD'$HostHDHeaderRow" -HorizontalAlignment Center
    $HostHDArrayStyle += New-ExcelStyle -Range $PctFreeColumn -NumberFormat '0.00%'
    $HostHDArrayConditionalText = @()
    $HostHDArrayConditionalText += New-ConditionalText -Range $PctFreeColumn -ConditionalType LessThanOrEqual 0.10 -ConditionalTextColor Maroon -BackgroundColor Pink
    $HostHDArrayConditionalText += New-ConditionalText -Range $PctFreeColumn -ConditionalType LessThanOrEqual 0.20 -ConditionalTextColor Brown -BackgroundColor Wheat
    $HostHDArray | Sort-Object "Host Name","Drive" | Export-Excel @ExcelProps -WorksheetName "Host HD" -Style $HostHDArrayStyle -ConditionalText $HostHDArrayConditionalText
}

# VM sheet
$VMArrayLastRow = ($VMArray | Measure-Object).Count + 1
If($VMArrayLastRow -gt 1){
    $VMArrayHeaderCount = Get-ColumnName ($VMArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $VMArrayHeaderRow = "`$A`$1:`$$VMArrayHeaderCount`$1"
    $VMStateColumn = "VMs!`$B`$2:`$B`$$VMArrayLastRow"
    $VMProcColumn = "VMs!`$F`$2:`$F`$$VMArrayLastRow"
    $VMRAMColumn = "VMs!`$H`$2:`$H`$$VMArrayLastRow"
    $VMDynamicRAMColumn = "VMs!`$I`$2:`$I`$$VMArrayLastRow"
    $VMSnapshotColumn = "VMs!`$T`$2:`$T`$$VMArrayLastRow"
    $VMArrayStyle = New-ExcelStyle -Range "'VMs'$VMArrayHeaderRow" -HorizontalAlignment Center
    $VMArrayConditionalText = @()
    $VMArrayConditionalText += New-ConditionalText -Range $VMStateColumn -ConditionalType ContainsText "Off" -ConditionalTextColor Brown -BackgroundColor Wheat
    $VMArrayConditionalText += New-ConditionalText -Range $VMProcColumn -ConditionalType NotEqual 2 -ConditionalTextColor Brown -BackgroundColor Wheat
    $VMArrayConditionalText += New-ConditionalText -Range $VMRAMColumn -ConditionalType NotContainsText "4.00 GiB" -ConditionalTextColor Brown -BackgroundColor Wheat
    $VMArrayConditionalText += New-ConditionalText -Range $VMDynamicRAMColumn -ConditionalType NotContainsText "False" -ConditionalTextColor Brown -BackgroundColor Wheat
    $VMArrayConditionalText += New-ConditionalText -Range $VMSnapshotColumn -ConditionalType GreaterThan 0 -ConditionalTextColor Brown -BackgroundColor Wheat
    $VMArray | Sort-Object "Host Name","VM" | Export-Excel @ExcelProps -WorksheetName "VMs" -Style $VMArrayStyle -ConditionalText $VMArrayConditionalText
}

# VM HD sheet
$VMHDArrayLastRow = ($VMHDArray | Measure-Object).Count + 1
If($VMHDArrayLastRow -gt 1){
    $VMHDHeaderCount = Get-ColumnName ($VMHDArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $VMHDHeaderRow = "`$A`$1:`$$VMHDHeaderCount`$1"
    $VMHDUsedColumn = "'VM HDs'!`$E`$2:`$E`$$VMHDArrayLastRow"
    $VMHDTypeColumn = "'VM HDs'!`$M`$2:`$M`$$VMHDArrayLastRow"
    $VMHDArrayStyle = @()
    $VMHDArrayStyle += New-ExcelStyle -Range "'VM HDs'$VMHDHeaderRow" -HorizontalAlignment Center
    $VMHDArrayStyle += New-ExcelStyle -Range $VMHDUsedColumn -NumberFormat "0.00%"
    $VMHDArrayConditionalText = @()
    $VMHDArrayConditionalText += New-ConditionalText -Range $VMHDUsedColumn -ConditionalType GreaterThanOrEqual 0.90 -ConditionalTextColor Maroon -BackgroundColor Pink
    $VMHDArrayConditionalText += New-ConditionalText -Range $VMHDUsedColumn -ConditionalType GreaterThanOrEqual 0.80 -ConditionalTextColor Brown -BackgroundColor Wheat
    $VMHDArrayConditionalText += New-ConditionalText -Range $VMHDTypeColumn -ConditionalType NotContainsText "Dynamic" -ConditionalTextColor Brown -BackgroundColor Wheat
    $VMHDArray | Sort-Object "Host Name","VM","HD File Path" | Export-Excel @ExcelProps -WorksheetName "VM HDs" -Style $VMHDArrayStyle -ConditionalText $VMHDArrayConditionalText
}

# Error sheet
If($ErrorArray){
    $ErrorHeaderCount = Get-ColumnName ($ErrorArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ErrorHeaderRow = "`$A`$1:`$$ErrorHeaderCount`$1"
    $ErrorArrayStyle = New-ExcelStyle -Range "'Errors'$ErrorHeaderRow" -HorizontalAlignment Center
    $ErrorArray | Sort-Object "Host Name" | Export-Excel @ExcelProps -WorksheetName "Errors" -Style $ErrorArrayStyle
}
    #endregion
