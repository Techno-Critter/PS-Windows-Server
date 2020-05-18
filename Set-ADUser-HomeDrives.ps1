<#
Author: Stan Crider
Date: 17April2020
Crap: Assigns home directories for users in specified OU. This will also
set the permissions on the folder to "modify" for the user account with
inheritance and set the ownership to the user account. If folder does not
exist it will be created.
#>

## Set variables
$ADLocation = "OU=Users,DC=ACME,DC=COM"
$HomeDirectory = "\\ACME.COM\Home_Folders"
$HomeDrive = "H:"
$Domain = "ACME"
# Exclude SAMAccountNames from creating or overwriting existing home directories; set to null if not used
$CreationExceptions = "Bob.Smith","Printer.ServiceAccount"

## Script below
$ADUsers = Get-ADUser -Filter * -SearchBase $ADLocation -Properties HomeDirectory,HomeDrive,SamAccountName,SID | Where-Object{$_.SamAccountName -notin $CreationExceptions}| Sort-Object SamAccountName

$UserPermissions = [System.Security.AccessControl.FileSystemRights]"Modify"
$Inheritance = [System.Security.AccessControl.InheritanceFlags]::"ContainerInherit", "ObjectInherit"
$Propagation = [System.Security.AccessControl.PropagationFlags]::None
$AccessControl =[System.Security.AccessControl.AccessControlType]::Allow

ForEach($User in $ADUsers){
    $NewHomeFolder = $null
    $FolderUser = $null
    $AccessRule = $null
    $FolderOwner = $null
    $FolderACL = $null

    If($User.Enabled -eq $true){
        $NewHomeFolder = ($HomeDirectory + $User.SamAccountName)
        $FolderUser = "$Domain\$($User.SamAccountName)"
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($FolderUser,$UserPermissions,$Inheritance,$Propagation,$AccessControl)
        $FolderOwner = New-Object System.Security.Principal.NTAccount($FolderUser)

        If((Test-Path $NewHomeFolder) -eq $false){
            New-Item -ItemType Directory -Path $NewHomeFolder
        }

        $FolderACL = Get-Acl -Path $NewHomeFolder
        $FolderACL.SetAccessRule($AccessRule)
        $FolderACL.SetOwner($FolderOwner)
        Set-Acl -Path $NewHomeFolder -AclObject $FolderACL
        Set-ADUser -Identity $User -HomeDrive $HomeDrive -HomeDirectory $NewHomeFolder
    }
}
