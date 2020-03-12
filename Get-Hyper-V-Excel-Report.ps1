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
$LogFile = "C:\Temp\HyperV Report $DateName.xlsx"

# Call AD and search for computers
$ServerList = Get-ADComputer -Filter {Enabled -eq $true} -SearchBase "DC=ACME,DC=COM" | Sort-Object Name
#endregion

#region Do not overwrite existing logfiles
If(Test-Path $LogFile){
    Write-Output "The file $LogFile already exists. Script terminated."
}
Else{
#endregion

#region Function: Change data sizes to legible values; converts number to string
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
        Write-Output ("Processing " + $Server.Name + " at " + (Get-Date) + "...")
#region Reset host variables
        $HostCPUCores = 0
        $HostTotalHDSpace = 0
        $HostTotalHDFreeSpace = 0
        $TotalVMAllocatedRAM = 0
        $TotalVMAllocatedHDSpace = 0
        $HostADProps = $null
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

        If(Test-Connection $Server.Name -Quiet -Count 2){
            Try{
                $HVRole = Get-WindowsFeature -ComputerName $Server.Name -Name Hyper-V -ErrorAction Stop
                If($HVRole.Installed){

#region Gather host data
                    $CimSession = New-CimSession -ComputerName $Server.Name -ErrorAction Stop
                    $HostADProps = Get-ADComputer -Identity $Server.Name -Properties IPv4Address,OperatingSystem -ErrorAction Stop
                    $HostManufacturer = Get-CimInstance -Class Win32_ComputerSystem -CimSession $CimSession -ErrorAction Stop
                    $TimeZone = Get-CimInstance -Class Win32_TimeZone -CimSession $CimSession -ErrorAction Stop
                    $CurrentTime = Invoke-Command -ComputerName $Server.Name -ScriptBlock {Get-Date -Format "MM/dd/yyyy HH:mm:ss"} -ErrorAction Stop
                    $HostBIOS = Get-CimInstance -Class Win32_BIOS -ComputerName $Server.Name -ErrorAction Stop
                    $HostOSProps = Get-CimInstance -Class Win32_OperatingSystem -CimSession $CimSession -ErrorAction Stop
                    $HostProcProps = Get-CimInstance -Class Win32_Processor -CimSession $CimSession -ErrorAction Stop
                    $HostRAMModules = Get-CimInstance -Class Win32_PhysicalMemory -CimSession $CimSession -ErrorAction Stop
                    $HostHardDrives = Get-CimInstance -Class Win32_LogicalDisk -CimSession $CimSession -Filter "DriveType = 3" -ErrorAction Stop
                    $HostHyperVProps = Invoke-Command -ComputerName $Server.Name {Get-VMHost | Select-Object *} -ErrorAction Stop
                    $VirtualMachines = Invoke-Command -ComputerName $Server.Name {Get-VM | Select-Object *} -ErrorAction Stop
                    Remove-CimSession -CimSession $CimSession -ErrorAction Stop
#endregion

#region Host CPU Properties
                    ForEach($HostProc in $HostProcProps){
                        $HostCPUCores = ($HostProc.NumberOfCores + $HostCPUCores)
                        $HostCPUArray += [PSCustomObject]@{
                            "Host Name"  = $Server.Name
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
                            "Host Name" = $Server.Name
                            "RAM Bank"  = $HostRAMModule.DeviceLocator
                            "Size"      = Get-Size $HostRAMModule.Capacity
                            "Bus Speed" = $HostRAMModule.Speed
                        }
                    }
#endregion

# Host hard drives
                                                                ForEach($HostHD in $HostHardDrives){
                    $HostHDArray += [PSCustomObject]@{
                        "Host Name" = $Server.Name
                        "Drive"     = $HostHD.DeviceID
                        "Size"      = Get-Size $HostHD.Size
                        "Free"      = Get-Size $HostHD.FreeSpace
                        "% Free"    = ("{0:N4}" -f ($HostHD.Freespace/$HostHD.Size))
                    }
                    $HostTotalHDSpace += $HostHD.Size
                    $HostTotalHDFreeSpace += $HostHD.FreeSpace
                }

# Virtual Machines
                    ForEach($VirtualMachine in $VirtualMachines){
# Reset VM variables
                        $TotHDSpace = 0
                        $TotHDUsedSpace = 0
                        $VMIPAddress = $null
# Gather VM data
                        $VMIntegratedServices = Invoke-Command -ComputerName $Server.Name -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMIntegrationService -VMName $VirtualMachine} -ErrorAction Stop
                        $VMProc = Invoke-Command -ComputerName $Server.Name -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMProcessor -VMName $VirtualMachine} -ErrorAction Stop
                        $VMMemory = Invoke-Command -ComputerName $Server.Name -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMMemory -VMName $VirtualMachine} -ErrorAction Stop
                        $VMHardDrives = Invoke-Command -ComputerName $Server.Name -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMHardDiskDrive -VMName $VirtualMachine} -ErrorAction Stop
                        $SnapShots = Invoke-Command -ComputerName $Server.Name -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMSnapshot -VMName $VirtualMachine} -ErrorAction Stop
                        If($VirtualMachine.State.ToString() -eq "Running"){
                            $VMIPAddress = Invoke-Command -ComputerName $Server.Name -ArgumentList $VirtualMachine.Name {$VirtualMachine = [string]$Args[0]; Get-VMNetworkAdapter -VMName $VirtualMachine} -ErrorAction Stop
                        }

# Get VM Hard Drive data
                        ForEach($VMHardDrive in $VMHardDrives){
                            $VHDUsedPct = 0
                            If($null -eq $VMHardDrive.DiskNumber){
                                $VHDObj = Invoke-Command -ComputerName $Server.Name -ArgumentList $VMHardDrive.Path {$VMHardDrivepath = [string]$Args[0]; Get-VHD -Path $VMHardDrivepath} -ErrorAction Stop
                                $VHDType = $VHDObj.VhdType.ToString()
                                $VHDUsedPct = $VHDObj.FileSize/$VHDObj.Size
                            }
                            Else{
                                $VHDObj = $null
                                $VHDType = "Physical"
                            }
                            $VMHDArray += [PSCustomObject]@{
                                "VM"              = $VirtualMachine.Name
                                "Host Name"       = $Server.Name
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
                            "Host Name"             = $Server.Name
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

# Host Properties
                    $HostArray += [PSCustomObject]@{
                        "Host Name"      = $Server.Name
                        "Host IP"        = $HostADProps.IPv4Address
                        "Host OS"        = $HostADProps.OperatingSystem
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
                    }
                }
            }
            Catch{
                $ErrorArray += [PSCustomObject]@{
                    "Host Name" = $Server.Name
                    "Error"     = $_.Exception.Message
                }
            }
        }
        Else{
            $ErrorArray += [PSCustomObject]@{
                "Host Name" = $Server.Name
                "Error"     = "Not responding to ping"
            }
        }
    }

# Export to file
    $HeaderRow = ("!`$A`$1:`$ZZ`$1")

# Hosts sheet
    $HostArrayLastRow = ($HostArray | Measure-Object).Count + 1
    If($HostArrayLastRow -gt 1){
        $HostArrayPctFreeColumn = "'HypV Hosts'!`$L`$2:`$L`$$HostArrayLastRow"
        $HostArrayStyle = @()
        $HostArrayStyle += New-ExcelStyle -Range "'HypV Hosts'$HeaderRow" -HorizontalAlignment Center
        $HostArrayStyle += New-ExcelStyle -Range $HostArrayPctFreeColumn -NumberFormat '0.00%'
        $TotalVMsColumn = "'HypV Hosts'!`$N`$2:`$N`$$HostArrayLastRow"
        $TotalVMsRunningColumn = "'HypV Hosts'!`$O`$2:`$O`$$HostArrayLastRow"
        $HostArray | Sort-Object "Host Name" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "HypV Hosts" -Style $HostArrayStyle -ConditionalText $(
            New-ConditionalText -Range $HostArrayPctFreeColumn -ConditionalType LessThan 0.10 -ConditionalTextColor Maroon -BackgroundColor Pink
            New-ConditionalText -Range $HostArrayPctFreeColumn -ConditionalType LessThan 0.20 -ConditionalTextColor Brown -BackgroundColor Wheat
            New-ConditionalText -Range $TotalVMsColumn -ConditionalType GreaterThan 2 -ConditionalTextColor Brown -BackgroundColor Wheat
            New-ConditionalText -Range $TotalVMsRunningColumn -ConditionalType GreaterThan 2 -ConditionalTextColor Maroon -BackgroundColor Pink
        )
    }

# Host CPU sheet
    $HostCPUArrayStyle = New-ExcelStyle -Range "'Host CPU'$HeaderRow" -HorizontalAlignment Center    
    $HostCPUArray | Sort-Object "Host Name","Socket" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Host CPU" -Style $HostCPUArrayStyle

# Host RAM sheet
    $HostRAMArrayStyle = New-ExcelStyle -Range "'Host RAM'$HeaderRow" -HorizontalAlignment Center
    $HostRAMArray | Sort-Object "Host Name","RAM Bank" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Host RAM" -Style $HostRAMArrayStyle

# Host HD sheet
    $HostHDArrayLastRow = ($HostHDArray | Measure-Object).Count + 1
    If($HostHDArrayLastRow -gt 1){
        $PctFreeColumn = "'Host HD'!`$E`2:`$E`$$HostHDArrayLastRow"
        $HostHDArrayStyle = @()
        $HostHDArrayStyle += New-ExcelStyle -Range "'Host HD'$HeaderRow" -HorizontalAlignment Center
        $HostHDArrayStyle += New-ExcelStyle -Range $PctFreeColumn -NumberFormat '0.00%'
        $HostHDArray | Sort-Object "Host Name","Drive" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Host HD" -Style $HostHDArrayStyle -ConditionalText $(
            New-ConditionalText -Range $PctFreeColumn -ConditionalType LessThanOrEqual 0.10 -ConditionalTextColor Maroon -BackgroundColor Pink
            New-ConditionalText -Range $PctFreeColumn -ConditionalType LessThanOrEqual 0.20 -ConditionalTextColor Brown -BackgroundColor Wheat
        )
    }

# VM sheet
    $VMArrayLastRow = ($VMArray | Measure-Object).Count + 1
    $VMArrayStyle = New-ExcelStyle -Range "'VMs'$HeaderRow" -HorizontalAlignment Center
    If($VMArrayLastRow -gt 1){
        $VMStateColumn = "VMs!`$B`$2:`$B`$$VMArrayLastRow"
        $VMProcColumn = "VMs!`$F`$2:`$F`$$VMArrayLastRow"
        $VMRAMColumn = "VMs!`$H`$2:`$H`$$VMArrayLastRow"
        $VMDynamicRAMColumn = "VMs!`$I`$2:`$I`$$VMArrayLastRow"
        $VMSnapshotColumn = "VMs!`$T`$2:`$T`$$VMArrayLastRow"
        $VMArray | Sort-Object "Host Name","VM" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "VMs" -Style $VMArrayStyle -ConditionalText $(
            New-ConditionalText -Range $VMStateColumn -ConditionalType ContainsText "Off" -ConditionalTextColor Brown -BackgroundColor Wheat
            New-ConditionalText -Range $VMProcColumn -ConditionalType NotEqual 2 -ConditionalTextColor Brown -BackgroundColor Wheat
            New-ConditionalText -Range $VMRAMColumn -ConditionalType NotContainsText "4.00 GiB" -ConditionalTextColor Brown -BackgroundColor Wheat
            New-ConditionalText -Range $VMDynamicRAMColumn -ConditionalType NotContainsText "False" -ConditionalTextColor Brown -BackgroundColor Wheat
            New-ConditionalText -Range $VMSnapshotColumn -ConditionalType GreaterThan 0 -ConditionalTextColor Brown -BackgroundColor Wheat
        )
    }

# VM HD sheet
    $VMHDArrayLastRow = ($VMHDArray | Measure-Object).Count + 1
    $VMHDUsedColumn = "'VM HDs'!`$E`$2:`$E`$$VMHDArrayLastRow"
    $VMHDTypeColumn = "'VM HDs'!`$M`$2:`$M`$$VMHDArrayLastRow"
    $VMHDArrayStyle = @()
    $VMHDArrayStyle += New-ExcelStyle -Range "'VM HDs'$HeaderRow" -HorizontalAlignment Center
    $VMHDArrayStyle += New-ExcelStyle -Range $VMHDUsedColumn -NumberFormat "0.00%"
    If($VMHDArrayLastRow -gt 1){
        $VMHDArray | Sort-Object "Host Name","VM","HD File Path" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "VM HDs" -Style $VMHDArrayStyle -ConditionalText $(
            New-ConditionalText -Range $VMHDUsedColumn -ConditionalType GreaterThanOrEqual 0.90 -ConditionalTextColor Maroon -BackgroundColor Pink
            New-ConditionalText -Range $VMHDUsedColumn -ConditionalType GreaterThanOrEqual 0.80 -ConditionalTextColor Brown -BackgroundColor Wheat
            New-ConditionalText -Range $VMHDTypeColumn -ConditionalType NotContainsText "Dynamic" -ConditionalTextColor Brown -BackgroundColor Wheat
        )
    }

    # Error sheet
    If($ErrorArray){
        $ErrorArrayStyle = New-ExcelStyle -Range "'Errors'$HeaderRow" -HorizontalAlignment Center
        $ErrorArray | Sort-Object "Host Name" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Errors" -Style $ErrorArrayStyle
    }
}