<#
Author: Stan Crider
Date: 17Dec2019
Crap: gets servers from specified AD Group and retreives file server roles and features
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

# Configure variables
$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\Reports\File Servers\File_Servers_$Date.xlsx"
$ADGroup = "CN=ServerGroup,DC=ACME,DC=COM"

# Exclude specific shares based on names
$ShareNameExclusions = "SMS_DP$",
    "SMSPKGC$",
    "SMSPKGD$",
    "SMSPKGE$",
    "SCCMContentLib$",
    "SMSSIG$",
    "MTATempStore$",
    "print$",
    "prnproc$"

# Exclude default shares by description
$ShareDescriptionExclusions = "Default share",
    "Remote IPC",
    "Remote Admin",
    "RemoteInstallation"

## Functions
#Convert drive sizes to legible strings
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

## Script below
If (Test-Path $LogFile) {
    Write-Warning "The file $LogFile already exists. Script terminated."
}
Else {
    # Configure parameters and arrays
    $ServerList = Get-ADGroupMember -Identity $ADGroup | Select-Object * | Sort-Object Name
    $ServerFeatureArray = @()
    $ServerDriveArray = @()
    $ServerShareArray = @()
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
        $Shares = $null
        $SMBAccess = $null
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
                        "DHCP"           = ($Features | Where-Object {$_.Name -eq "DHCP"}).InstallState
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
                # Call server properties
                $ServerSession = New-CimSession -ComputerName $Server.Name -ErrorAction Stop
                $OpSys = Get-CimInstance Win32_OperatingSystem -CimSession $ServerSession -ErrorAction SilentlyContinue
                $Drives = Get-CimInstance Win32_LogicalDisk -CimSession $ServerSession -Filter "DriveType='3'" -ErrorAction SilentlyContinue
                $Volumes = Get-CimInstance Win32_Volume -CimSession $ServerSession -ErrorAction SilentlyContinue
                $Dedup = Invoke-Command -ComputerName $Server.Name -ScriptBlock{Get-DeDupStatus} -ErrorAction SilentlyContinue
                $ScheduledTasks = Get-ScheduledTask -CimSession $ServerSession -TaskName ShadowCopyVolume* -ErrorAction SilentlyContinue
                $Shares = Get-CimInstance Win32_Share -CimSession $ServerSession | Where-Object{($_.Description -notin $ShareDescriptionExclusions) -and ($_.Name -notin $ShareNameExclusions)}

                # Server Volume Shadow Copies
                $ShadowCopyVolumes = $null
                If ($ScheduledTasks) {
                    $ShadowCopyVolumes = $ScheduledTasks.Actions.Arguments | ForEach-Object { $_.Split()[3].Split("=")[1] }
                }

                # Server Drives
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
                        "Used"              = Get-Size ($Drive.Size - $Drive.FreeSpace)
                        "Dedup Savings"     = Get-Size $DedupSavings
                        "Raw Size"          = $Drive.Size
                        "Raw Free"          = $Drive.FreeSpace
                        "Raw Used"          = ($Drive.Size - $Drive.FreeSpace)
                        "Raw Dedup Savings" = $DedupSavings
                        "Dedup"             = $DedupStatus
                        "Shadow Copy"       = $ShadowEnabled
                        "Volume ID"         = $VolumeID
                    }
                }
                
                # Server Shares
                ForEach($Share in $Shares){
                    $SMBAccess = Get-SmbShareAccess -Name $Share.Name -CimSession $ServerSession
                    ForEach($SMBObject in $SMBAccess){
                        $ServerShareArray += [PSCustomObject]@{
                            "Server"       = $Server.Name
                            "Share Name"   = $Share.Name
                            "Share Path"   = $Share.Path
                            "Description"  = $Share.Description
                            "Account Name" = $SMBObject.AccountName
                            "Access Right" = $SMBObject.AccessRight
                            "Control Type" = $SMBObject.AccessControlType
                        }
                    }
                }
                
                # Close server connection
                Remove-CimSession -CimSession $ServerSession
            }
            Catch {
                $ErrorArray += [PSCustomObject]@{
                    "Server"  = $Server.Name
                    "Command" = $_.Exception.Command
                    "Error"   = $_.Exception.Message
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
                "DHCP"           = "N/A"
            }
        }
    }

    ## Export to Excel
    # File Servers worksheet
    $ServerSheetHeaderCount = Get-ColumnName ($ServerFeatureArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ServerSheetHeaderRow = "`$A`$1:`$$ServerSheetHeaderCount`$1"
    $ServerSheetLastRow = ($ServerFeatureArray | Measure-Object).Count + 1
    If($ServerSheetLastRow -gt 1){
        $FileServerColumn = "'File Servers'!`$F`$2:`$F`$$ServerSheetLastRow"
        $DeduplicationColumn = "'File Servers'!`$G`$2:`$G`$$ServerSheetLastRow"
        $DFSRColumn = "'File Servers'!`$H`$2:`$H`$$ServerSheetLastRow"
        $DFSNameSpaceColumn = "'File Servers'!`$I`$2:`$I`$$ServerSheetLastRow"
        $FSRMColumn = "'File Servers'!`$J`$2:`$J`$$ServerSheetLastRow"
        $SearchServiceColumn = "'File Servers'!`$K`$2:`$K`$$ServerSheetLastRow"

        $ServerFeatureStyle = @()
        $ServerFeatureStyle += New-ExcelStyle -Range "'File Servers'!$ServerSheetHeaderRow" -HorizontalAlignment Center

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
    $DriveSheetHeaderCount = Get-ColumnName ($ServerDriveArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $DriveSheetHeaderRow = "`$A`$1:`$$DriveSheetHeaderCount`$1"
    $DriveSheetLastRow = ($ServerDriveArray | Measure-Object).Count + 1
    If($DriveSheetLastRow -gt 1){
        $DriveSystemDrColumn = "'Server Drives'!`$D`$2:`$D`$$DriveSheetLastRow"
        $DriveRawSizeColumn = "'Server Drives'!`$I`$2:`$I`$$DriveSheetLastRow"
        $DriveRawFreeColumn = "'Server Drives'!`$J`$2:`$J`$$DriveSheetLastRow"
        $DriveRawUsedColumn = "'Server Drives'!`$K`$2:`$K`$$DriveSheetLastRow"
        $DriveRawSavingsColumn = "'Server Drives'!`$L`$2:`$L`$$DriveSheetLastRow"
        $DriveDedupColumn = "'Server Drives'!`$M`$2:`$M`$$DriveSheetLastRow"
        $DriveShadowColumn = "'Server Drives'!`$N`$2:`$N`$$DriveSheetLastRow"

        $DriveCondText = @()
        $DriveCondText += New-ConditionalText -Range $DriveSystemDrColumn -ConditionalType ContainsText "TRUE" -ConditionalTextColor Navy -BackgroundColor Cyan
        $DriveCondText += New-ConditionalText -Range $DriveDedupColumn -ConditionalType Expression "=AND(`$D2=TRUE,`$N2=TRUE)" -ConditionalTextColor Maroon -BackgroundColor Pink
        $DriveCondText += New-ConditionalText -Range $DriveShadowColumn -ConditionalType Expression "=AND(`$D2=FALSE,`$N2=FALSE)" -ConditionalTextColor Maroon -BackgroundColor Pink

        $DriveStyle = @()
        $DriveStyle += New-ExcelStyle -Range "'Server Drives'!$DriveSheetHeaderRow" -HorizontalAlignment Center
        $DriveStyle += New-ExcelStyle -Range $DriveRawSizeColumn -NumberFormat 0
        $DriveStyle += New-ExcelStyle -Range $DriveRawFreeColumn -NumberFormat 0
        $DriveStyle += New-ExcelStyle -Range $DriveRawUsedColumn -NumberFormat 0
        $DriveStyle += New-ExcelStyle -Range $DriveRawSavingsColumn -NumberFormat 0

        $ServerDriveArray | Sort-Object "Server", "Drive" | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Server Drives" -ConditionalText $DriveCondText -Style $DriveStyle
    }

    # Server Share worksheet
    $ServerSharesHeaderCount = Get-ColumnName ($ServerShareArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ServerSharesHeaderRow = "`$A`$1:`$$ServerSharesHeaderCount`$1"
    $ServerSharesLastRow = ($ServerShareArray | Measure-Object).Count + 1
    $SharesControlColumn = "'Server Shares'!`$G`$2:`$G`$$ServerSharesLastRow"
    $ServerSharesStyle = @()
    $ServerSharesStyle += New-ExcelStyle -Range "'Server Shares'!$ServerSharesHeaderRow" -HorizontalAlignment Center
    $ServerSharesConditionalText = @()
    $ServerSharesConditionalText += New-ConditionalText -Range $SharesControlColumn -ConditionalType ContainsText "Deny" -ConditionalTextColor Maroon -BackgroundColor Pink
    $ServerShareArray | Sort-Object "Server","Share Path","Account Name" | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Server Shares" -ConditionalText $ServerSharesConditionalText -Style $ServerSharesStyle

    # Errors worksheet
    If($ErrorArray -ne ""){
        $ErrorSheetHeaderCount = Get-ColumnName ($ErrorArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $ErrorSheetHeaderRow = "`$A`$1:`$$ErrorSheetHeaderCount`$1"
        $ErrorStyle = New-ExcelStyle -Range "Errors!$ErrorSheetHeaderRow" -HorizontalAlignment Center
        $ErrorArray | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Errors" -Style $ErrorStyle
    }
}
