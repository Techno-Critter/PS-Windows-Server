<#
Author: Stan Crider
Date: 10Mar2020
Crap:
Get Home directories and ACL settings from specified AD users
### Must have ImportExcel and NTFSSecurity modules installed!!!
### https://github.com/dfinke/ImportExcel
### https://github.com/raandree/NTFSSecurity
#>

# Configure variables
$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\Reports\Home Drives\HomeDrives_$Date.xlsx"
$ADLocation = "OU=Company Users,DC=acme,DC=com"

### Script below

## FUNCTION: Convert number of object items into Excel column headers
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

# Check if logfile exists and terminate if it does
If(Test-Path $LogFile){
    Write-Output "The file $LogFile already exists. Script terminated."
}
Else{

    # Configure array objects
    $UserHomeArray = @()
    $PermissionsArray = @()
    $ErrorArray = @()

    # Get users from specified AD location
    $UserProps = Get-ADUser -Filter * -SearchBase $ADLocation -Properties * -ErrorAction Stop | Sort-Object Name

    # List user properties and home directory properties
    ForEach($User in $UserProps){
        $ACLMisMatch = $true
        $OwnerMisMatch = $true
        $FolderNameMismatch = $true
        If($null -ne $User.HomeDirectory){
            If(Test-Path $User.HomeDirectory){
                $Access = Get-NTFSAccess -Path $User.HomeDirectory -ExcludeInherited
                $Owner = Get-NTFSOwner -Path $User.HomeDirectory
                $Inherited = Get-NTFSInheritance -Path $User.HomeDirectory

                # Verify owner of folder is assigned user
                If($User.SID -eq $Owner.Owner.Sid){
                    $OwnerMisMatch = $false
                }

                # Verify home folder name matches user account name (not case sensitive)
                If($User.SamAccountName -eq (($User.HomeDirectory).Split("\\")[-1])){
                    $FolderNameMismatch = $false
                }

                # List each AD object with explicit permissions and what permissions are granted
                ForEach($ACLObject in $Access){
                    If($ACLObject.Account.Sid -eq $User.SID){
                        $ACLMisMatch = $false
                    }

                    $PermissionsArray += [PSCustomObject]@{
                        "Name"      = $User.Name
                        "Path"      = $User.HomeDirectory
                        "AD Object" = $ACLObject.Account
                        "Rights"    = $ACLObject.AccessRights
                        "Type"      = $ACLObject.AccessControlType
                    }
                }

                $UserHomeArray += [PSCustomObject]@{
                    "Name"                 = $User.Name
                    "Account Name"         = $User.SamAccountName
                    "Enabled"              = $User.Enabled
                    "Drive"                = $User.HomeDrive
                    "FolderNameMismatch"   = $FolderNameMismatch
                    "Path"                 = $User.HomeDirectory
                    "Explicit Permissions" = $Access.Account.AccountName -join ", "
                    "Inherit Permissions"  = $Inherited.AccessInheritanceEnabled
                    "Folder Owner"         = $Owner.Owner.AccountName
                    "ACL Mismatch"         = $ACLMisMatch
                    "Owner Mismatch"       = $OwnerMisMatch
                }
            }
            # Error handling
            Else{
                $ErrorArray += [PSCustomObject]@{
                    "Name"  = $User.Name
                    "Error" = "Home directory does not exist"
                }
            }
        }
        # If user has no home directory mark N/A
        Else{
            $UserHomeArray += [PSCustomObject]@{
                "Name"                 = $User.Name
                "Account Name"         = $User.SamAccountName
                "Enabled"              = $User.Enabled
                "Drive"                = "N/A"
                "Path"                 = "N/A"
                "FolderNameMismatch"   = "N/A"
                "Explicit Permissions" = "N/A"
                "Inherit Permissions"  = "N/A"
                "Folder Owner"         = "N/A"
                "ACL Mismatch"         = "N/A"
                "Owner Mismatch"       = "N/A"
            }
        }
    }

    ### Output to Excel
    ## UserHome worksheet
    # Call Excel column count function based on number of NoteProperty members in UserHomeArray
    $UserSheetLastRow = ($UserHomeArray | Measure-Object).Count + 1
    If($UserSheetLastRow -gt 1){
        $UserHomeHeaderCount = Get-ColumnName ($UserHomeArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $UserHeaderRow = "`$A`$1:`$$UserHomeHeaderCount`$1"
        $UserEnabledColumn = "`$C`$2:`$C`$$UserSheetLastRow"
        $FolderMismatchColumn = "`$D`$2:`$D`$$UserSheetLastRow"
        $UserPermissionsColumn = "`$H`$2:`$H`$$UserSheetLastRow"
        $UserOwnerColumn = "`$I`$2:`$I`$$UserSheetLastRow"

        # Format style for User sheet
        $UserSheetStyle = @()
        $UserSheetStyle += New-ExcelStyle -Range "'User Homes'$UserHeaderRow" -HorizontalAlignment Center

        # Format conditions for User sheet
        $UserSheetConditionalText = @()
        $UserSheetConditionalText += New-ConditionalText -Range $UserEnabledColumn -ConditionalType ContainsText "FALSE" -ConditionalTextColor Maroon -BackgroundColor Pink
        $UserSheetConditionalText += New-ConditionalText -Range $FolderMismatchColumn -ConditionalType ContainsText "TRUE" -ConditionalTextColor Maroon -BackgroundColor Pink
        $UserSheetConditionalText += New-ConditionalText -Range $UserPermissionsColumn -ConditionalType Expression "=AND(`$J2=TRUE)" -ConditionalTextColor Maroon -BackgroundColor Pink
        $UserSheetConditionalText += New-ConditionalText -Range $UserOwnerColumn -ConditionalType Expression "=AND(`$K2=TRUE)" -ConditionalTextColor Maroon -BackgroundColor Pink

        $UserHomeArray | Export-Excel -Path $LogFile  -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "User Homes" -ConditionalText $UserSheetConditionalText -Style $UserSheetStyle
    }
    ## Permissions worksheet
    # Call Excel column count function based on number of NoteProperty members in PermissionsArray
    $PermissionsSheetLastRow = ($PermissionsArray | Measure-Object).Count + 1
    If($PermissionsSheetLastRow -gt 1){
        $PermissionsHeaderCount = Get-ColumnName ($PermissionsArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $PermissionsHeaderRow = "`$A`$1:`$$PermissionsHeaderCount`$1"
        $PermissionsRightsColumn = "`$D`$2:`$D`$$PermissionsSheetLastRow"
        $PermissionsInheritedColumn = "`$F`$2:`$F`$$PermissionsSheetLastRow"

        # Format style for Permissions sheet
        $PermissionsStyle = @()
        $PermissionsStyle += New-ExcelStyle -Range "'Permissions'$PermissionsHeaderRow" -HorizontalAlignment Center

        # Format conditions for Permissions sheet
        $PermissionsConditionalText = @()
        $PermissionsConditionalText += New-ConditionalText -Range $PermissionsRightsColumn -ConditionalType NotContainsText "Modify" -ConditionalTextColor Maroon -BackgroundColor Pink
        $PermissionsConditionalText += New-ConditionalText -Range $PermissionsInheritedColumn -ConditionalType ContainsText "FALSE"  -ConditionalTextColor Maroon -BackgroundColor Pink

        $PermissionsArray | Export-Excel -Path $LogFile  -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Permissions" -ConditionalText $PermissionsConditionalText -Style $PermissionsStyle
    }

    ## Error worksheet
    If($ErrorArray){
        # Call Excel column count function based on number of NoteProperty members in ErrorArray
        $ErrorHeaderCount = Get-ColumnName ($ErrorArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count + 1
        $ErrorHeaderRow = "`$A`$1:`$$ErrorHeaderCount`$1"
        $ErrorSheetLastRow = ($ErrorArray | Measure-Object).Count
        If($ErrorSheetLastRow -gt 1){
            # Format style for Error sheet
            $ErrorStyle = @()
            $ErrorStyle += New-ExcelStyle -Range "'Errors'$ErrorHeaderRow" -HorizontalAlignment Center
            $ErrorArray | Export-Excel -Path $LogFile  -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Errors" -Style $ErrorStyle
        }
    }
}
