<#
Author: Stan Crider
Date: 20Mar2020
Crap:
Gets the members of specified groups from specified computers
Used for validating group memberships against desired configuration, i.e, GPO, AD groups, etc
#>

## Input variables below
$LocalGroups = "Administrators","DHCP Administrators"
$Computers = "DHCPSrvr01","DHCPSrvr02"

## Script below; do not cross
# Prep Arrays
$MemberArray = @()
$ErrorArray = @()

# Run through each computer
ForEach($Computer in $Computers){
    Write-Output "Processing $Computer ..." # used for troubleshooting purposes
    # Call local groups once and reuse for each group variable
    $AllGroupMembers = $null
    Try{
        $AllGroupMembers = Get-CimInstance Win32_GroupUser -ComputerName $Computer -ErrorAction Stop
    }
    # Error handling, duh
    Catch{
        $ErrorArray += [PSCustomObject]@{
            "Computer Name" = $Computer
            "Error" = $_.Exception.Message
        }
    }

    # Run through each group variable
    ForEach($LocalGroup in $LocalGroups){
        # Find matches
        $GroupMembers = $AllGroupMembers | Where-Object {$_.GroupComponent.Name -eq $LocalGroup}
        # Run through each match
        ForEach($GroupMember in $GroupMembers){
            # Run through each match object
            ForEach($PartComponent in $GroupMember.PartComponent){
                # Output matches to custom object
                $MemberArray += [PSCustomObject]@{
                    "Computer" = $GroupMember.PSComputerName
                    "Member of" = $GroupMember.GroupComponent.Name
                    "Full Name" = ("" + $PartComponent.Domain + "\" + $PartComponent.Name)
                    "Domain" = $PartComponent.Domain
                    "Name" = $PartComponent.Name
                }
            }
        }
    }
}

## Output
$MemberArray
$ErrorArray
