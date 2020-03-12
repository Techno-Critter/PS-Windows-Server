<#
Author: Stan Crider
Date: 17Dec2019
Crap: gets servers from specified AD Group and retreives file server roles and features
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

# Function: Convert drive sizes to legible strings
Function Get-Size([double]$DataSize) {
    Switch ($DataSize) {
        { $_ -lt 1KB } {
            $DataValue = "$DataSize B"
        }
        { ($_ -ge 1KB) -and ($_ -lt 1MB) } {
            $DataValue = "{0:N2}" -f ($DataSize / 1KB) + " KiB"
        }
        { ($_ -ge 1MB) -and ($_ -lt 1GB) } {
            $DataValue = "{0:N2}" -f ($DataSize / 1MB) + " MiB"
        }
        { ($_ -ge 1GB) -and ($_ -lt 1TB) } {
            $DataValue = "{0:N2}" -f ($DataSize / 1GB) + " GiB"
        }
        { ($_ -ge 1TB) -and ($_ -lt 1PB) } {
            $DataValue = "{0:N2}" -f ($DataSize / 1TB) + " TiB"
        }
        Default {
            $DataValue = "{0:N2}" -f ($DataSize / 1PB) + " PiB"
        }
    }
    $DataValue
}

# Configure variables
$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\Reports\File Servers\File_Servers_$Date.xlsx"

If (Test-Path $LogFile) {
    Write-Warning "The file $LogFile already exists. Script terminated."
}
Else {
    # Configure parameters and arrays
    $ServerList = Get-ADGroupMember -Identity "CN=File Servers,DC=acme,DC=com" | Select-Object * | Sort-Object Name
    $ServerFeatureArray = @()
    $ServerDriveArray = @()
    $ErrorArray = @()

    # Search through list
    ForEach ($Member in $ServerList) {
        $Features = $null
        $Server = $null
        $OpSys = $null
        $Drives = $null
        $Volumes = $null
        $Dedup = $null
        $ScheduledTasks = $null
        $ShadowCopyVolumes = $null
        Write-Output ("Processing $($Member.Name)...")

        $Server = Get-ADComputer -Identity $Member.DistinguishedName -Properties *
        If (Test-Connection $Server.Name -Quiet) {
            Try {
                $Features = Get-WindowsFeature -ComputerName $Server.Name -ErrorAction SilentlyContinue
                If (($Features | Where-Object { $_.DisplayName -eq "File Server" }).InstallState -eq "Installed") {
                    $ServerFeatureArray += [PSCustomObject]@{
                        "Server"         = $Server.Name
                        "IP Address"     = $Server.IPv4Address
                        "OS"             = $Server.OperatingSystem
                        "Enabled"        = $Server.Enabled
                        "Online"         = $true #Responds to ping
                        "File Server"    = ($Features | Where-Object { $_.Name -eq "FS-FileServer" }).InstallState
                        "Deduplication"  = ($Features | Where-Object { $_.Name -eq "FS-Data-Deduplication" }).InstallState
                        "DFSR"           = ($Features | Where-Object { $_.Name -eq "FS-DFS-Replication" }).InstallState
                        "DFS Namespace"  = ($Features | Where-Object { $_.Name -eq "FS-DFS-Namespace" }).InstallState
                        "FSRM"           = ($Features | Where-Object { $_.Name -eq "FS-Resource-Manager" }).InstallState
                        "Search Service" = ($Features | Where-Object { $_.Name -eq "Search-Service" }).InstallState
                    }
                }
            }
            Catch {
                $ErrorArray += [PSCustomObject]@{
                    "Server" = $Server.Name
                    "Command" = $_.Exception.Command
                    "Error"  = $_.Exception.Message
                }
            }
            Try {
                $ServerSession = New-CimSession -ComputerName $Server.Name -ErrorAction Stop
                $OpSys = Get-CimInstance Win32_OperatingSystem -CimSession $ServerSession -ErrorAction SilentlyContinue
                $Drives = Get-CimInstance Win32_LogicalDisk -CimSession $ServerSession -Filter "DriveType='3'" -ErrorAction SilentlyContinue
                $Volumes = Get-CimInstance Win32_Volume -CimSession $ServerSession -ErrorAction SilentlyContinue
                $Dedup = Invoke-Command -ComputerName $Server.Name -ScriptBlock{Get-DeDupStatus} -ErrorAction SilentlyContinue
                $ScheduledTasks = Get-ScheduledTask -CimSession $ServerSession -TaskName ShadowCopyVolume* -ErrorAction SilentlyContinue
                Remove-CimSession -CimSession $ServerSession

                $ShadowCopyVolumes = $null
                If ($ScheduledTasks) {
                    $ShadowCopyVolumes = $ScheduledTasks.Actions.Arguments | ForEach-Object { $_.Split()[3].Split("=")[1] }
                }

                ForEach ($Drive in $Drives) {
                    $DedupStatus = "Disabled"
                    $VolumeID = ($Volumes | Where-Object { $_.DriveLetter -eq $Drive.DeviceID }).DeviceID

                    $DedupSavings = 0
                    If ($Dedup) {
                        ForEach ($Vol in $Dedup) {
                            If ($Vol.Volume -eq $Drive.DeviceID) {
                                $DedupStatus = "Enabled"
                                $DedupSavings = $Vol.SavedSpace
                            }
                        }
                    }

                    $IsOpSysDrive = $false
                    If ($Drive.DeviceID -eq $OpSys.SystemDrive) {
                        $IsOpSysDrive = $true
                    }

                    $ShadowEnabled = $false
                    If ($VolumeID -in $ShadowCopyVolumes) {
                        $ShadowEnabled = $true
                    }

                    $ServerDriveArray += [PSCustomObject]@{
                        "Server"            = $Server.Name
                        "Drive"             = $Drive.DeviceID
                        "Label"             = $Drive.VolumeName
                        "System Drive"      = $IsOpSysDrive
                        "Size"              = Get-Size $Drive.Size
                        "Free"              = Get-Size $Drive.FreeSpace
                        "Dedup Savings"     = Get-Size $DedupSavings
                        "Raw Size"          = $Drive.Size
                        "Raw Free"          = $Drive.FreeSpace
                        "Raw Dedup Savings" = $DedupSavings
                        "Dedup"             = $DedupStatus
                        "Shadow Copy"       = $ShadowEnabled
                        "Volume ID"         = $VolumeID
                    }
                }
            }
            Catch {
                $ErrorArray += [PSCustomObject]@{
                    "Server" = $Server.Name
                    "Command" = $_.Exception.Command
                    "Error"  = $_.Exception.Message
                }
            }
        }
        Else {
            Write-Warning "Server $($Server.Name) failed to respond."
            $ServerFeatureArray += [PSCustomObject]@{
                "Server"         = $Server.Name
                "IP Address"     = $Server.IPv4Address
                "OS"             = $Server.OperatingSystem
                "System Drive"   = $OpSys.SystemDrive
                "Enabled"        = $Server.Enabled
                "Online"         = $false #No response to ping
                "File Server"    = "N/A"
                "Deduplication"  = "N/A"
                "DFSR"           = "N/A"
                "DFS Namespace"  = "N/A"
                "FSRM"           = "N/A"
                "Search Service" = "N/A"
            }
        }
    }

    # Export to Excel
    $HeaderRow = ("!`$A`$1:`$Z`$1")

    # File Servers worksheet
    $ServerSheetLastRow = ($ServerFeatureArray | Measure-Object).Count + 1
    If($ServerSheetLastRow -gt 1){
        $FileServerColumn = "`$F`$2:`$F`$$ServerSheetLastRow"
        $DeduplicationColumn = "`$G`$2:`$G`$$ServerSheetLastRow"
        $DFSRColumn = "`$H`$2:`$H`$$ServerSheetLastRow"
        $DFSNameSpaceColumn = "`$I`$2:`$I`$$ServerSheetLastRow"
        $FSRMColumn = "`$J`$2:`$J`$$ServerSheetLastRow"
        $SearchServiceColumn = "`$K`$2:`$K`$$ServerSheetLastRow"

        $ServerFeatureStyle = @()
        $ServerFeatureStyle += New-ExcelStyle -Range "'File Servers'$HeaderRow" -HorizontalAlignment Center

        $ServerFeatureCondText = @()
        $ServerFeatureCondText += New-ConditionalText -Range $FileServerColumn -ConditionalType ContainsText "Installed" -ConditionalTextColor Green -BackgroundColor LightGreen
        $ServerFeatureCondText += New-ConditionalText -Range $DeduplicationColumn -ConditionalType ContainsText "Installed" -ConditionalTextColor Green -BackgroundColor LightGreen
        $ServerFeatureCondText += New-ConditionalText -Range $DFSRColumn -ConditionalType ContainsText "Installed" -ConditionalTextColor Green -BackgroundColor LightGreen
        $ServerFeatureCondText += New-ConditionalText -Range $DFSNameSpaceColumn -ConditionalType ContainsText "Installed" -ConditionalTextColor Green -BackgroundColor LightGreen
        $ServerFeatureCondText += New-ConditionalText -Range $FSRMColumn -ConditionalType ContainsText "Installed" -ConditionalTextColor Green -BackgroundColor LightGreen
        $ServerFeatureCondText += New-ConditionalText -Range $SearchServiceColumn -ConditionalType ContainsText "Installed" -ConditionalTextColor Green -BackgroundColor LightGreen

        $ServerFeatureArray | Sort-Object "Server" | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "File Servers" -ConditionalText $ServerFeatureCondText -Style $ServerFeatureStyle
    }
    # Server Drives worksheet
    $DriveSheetLastRow = ($ServerDriveArray | Measure-Object).Count + 1
    If ($DriveSheetLastRow -gt 1) {
        $DriveRawSizeColumn = "'Drive!'`$H`$2:`$H`$$DriveSheetLastRow"
        $DriveRawFreeColumn = "'Drive!'`$I`$2:`$I`$$DriveSheetLastRow"
        $DriveRawSavingsColumn = "'Drive!'`$J`$2:`$J`$$DriveSheetLastRow"
        #$DriveSystemDrColumn = "`$D`$2:`$D`$$DriveSheetLastRow"
        $DriveDedupColumn = "`$K`$2:`$K`$$DriveSheetLastRow"
        $DriveShadowColumn = "`$L`$2:`$L`$$DriveSheetLastRow"

        $DriveCondText = @()
        $DriveCondText += New-ConditionalText -Range $DriveDedupColumn -ConditionalType Expression "=AND(`$D2=TRUE,`$L2=TRUE)" -ConditionalTextColor Maroon -BackgroundColor Pink
        $DriveCondText += New-ConditionalText -Range $DriveShadowColumn -ConditionalType Expression "=AND(`$D2=FALSE,`$L2=FALSE)" -ConditionalTextColor Maroon -BackgroundColor Pink

        $DriveStyle = @()
        $DriveStyle += New-ExcelStyle -Range "'Server Drives'$HeaderRow" -HorizontalAlignment Center
        $DriveStyle += New-ExcelStyle -Range $DriveRawSizeColumn -NumberFormat 0
        $DriveStyle += New-ExcelStyle -Range $DriveRawFreeColumn -NumberFormat 0
        $DriveStyle += New-ExcelStyle -Range $DriveRawSavingsColumn -NumberFormat 0

        $ServerDriveArray | Sort-Object "Server", "Drive" | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Server Drives" -ConditionalText $DriveCondText -Style $DriveStyle
    }

    # Errors worksheet
    If ($ErrorArray -ne "") {
        $ErrorStyle = New-ExcelStyle -Range "'Errors'$HeaderRow" -HorizontalAlignment Center
        $ErrorArray | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Errors" -Style $ErrorStyle
    }
}
