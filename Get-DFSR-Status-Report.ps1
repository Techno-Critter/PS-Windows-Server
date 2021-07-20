<#
Author: Stan Crider
Date: 7Jul2021
Crap:
Lists DFSR folders and status on specified servers and outputs to Excel
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

## Functions and script prep
# Configure variables
$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\DFS\DFS_Report_$Date.xlsx"
$RepServers =  "ServerFS01","ServerFS02"

# Class: worksheet properties
Class DfsrInfo{
    [string]$Server
    [string]$GroupName
    [string]$FolderName
    [string]$ContentPath
    [string]$PrimaryMember
    [string]$StagingQuota
    [string]$State
    [string]$Error
}

# Function: Convert raw data size to legible output
Function Get-Size([double]$DataSize){
    Switch($DataSize){
        {$_ -lt 1KB}{
            $DataValue =  "$DataSize B"
        }
        {($_ -ge 1KB) -and ($_ -lt 1MB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1KB) + " KiB"
        }
        {($_ -ge 1MB) -and ($_ -lt 1GB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1MB) + " MiB"
        }
        {($_ -ge 1GB) -and ($_ -lt 1TB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1GB) + " GiB"
        }
        {($_ -ge 1TB) -and ($_ -lt 1PB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1TB) + " TiB"
        }
        Default{
            $DataValue = "{0:N2}" -f ($DataSize/1PB) + " PiB"
        }
    }
    $DataValue
}

# Function: Set column count for spreadsheet output
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

# Function: format and output to Excel
Function Invoke-ExcelOutput{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $WorksheetData,
        [Parameter(Mandatory = $true)]
        [string]
        $WorksheetName,
        [Parameter(Mandatory = $true)]
        [string]
        $WorkbookName
    )
    $ExcelProps = @{
        Autosize = $true;
        FreezeTopRow = $true;
        BoldTopRow = $true;
    }
    $ExcelProps.Path = $WorkbookName
    $DataLastRow = ($WorksheetData | Measure-Object).Count + 1
    If($DataLastRow -gt 1){
        $DataHeaderCount = Get-ColumnName ($WorksheetData | Get-Member | Where-Object{$_.MemberType -match "Property"} | Measure-Object).Count
        $DataHeaderRow = "'Clusters'!`$A`$1:`$$DataHeaderCount`$1"
        $DataStyle = New-ExcelStyle -Range $DataHeaderRow -HorizontalAlignment Center
        $WorksheetData | Export-Excel @ExcelProps -WorkSheetname $WorksheetName -Style $DataStyle
    }
}

# Convert DFSR status to legible output
$DFSRStatus = @{
    [byte]0="Uninitialized";
    [byte]1="Initialized";
    [byte]2="Initial Sync";
    [byte]3="Auto Recovery";
    [byte]4="Normal";
    [byte]5="In Error"
}

$Errors = @()

## Script below
If(Test-Path $LogFile){
    Write-Warning "The file $LogFile already exists. Script terminated."
}
Else{
    ForEach($RepServer in $RepServers){
# Prep output array
        $DfsrInfo = @()

        If(-Not(Test-Connection -ComputerName $RepServer -Quiet -Count 2)){
            $Errors += [DfsrInfo]@{
                "Server" = $RepServer
                "Error"  = "Server $RepServer is not responding to ping."
            }
            Continue
        }
        Else{
            Try{
                $DfsFldrs = Get-CimInstance -Namespace "Root\MicrosoftDFS" DfsrReplicatedFolderInfo -ComputerName $RepServer -ErrorAction Stop
                $DfsrMembershipGroups = Get-DfsrMembership -ComputerName $RepServer -ErrorAction Stop

                ForEach($DfsrGroup in $DfsrMembershipGroups){
                    $DfsFldr = $null
                    $DfsFldr = $DfsFldrs | Where-Object{$_.ReplicationGroupName -eq $DfsrGroup.GroupName}

                    $DfsrInfo += [DfsrInfo]@{
                        "Server"        = $DfsrGroup.ComputerName
                        "GroupName"     = $DfsrGroup.GroupName
                        "FolderName"    = $DfsrGroup.FolderName
                        "ContentPath"   = $DfsrGroup.ContentPath
                        "PrimaryMember" = $DfsrGroup.PrimaryMember
                        "StagingQuota"  = Get-Size ($DfsrGroup.StagingPathQuotaInMB*1MB)
                        "State"         = $DFSRStatus[$DfsFldr.State]
                    }
                }
                Invoke-ExcelOutput -WorksheetData ($DfsrInfo | Sort-Object "GroupName") -WorksheetName $RepServer -WorkbookName $LogFile
            }
            Catch{
                $Errors += [DfsrInfo]@{
                    "Server" = $RepServer
                    "Error"  = $_.Exception.Message
                }
            }
        }
    }


# Create error worksheet and add to Excel workbook if necessary
    If($Errors){
        Invoke-ExcelOutput -WorksheetData $Errors -WorksheetName "Errors" -WorkbookName $LogFile
    }
}