<#
Author: Stan Crider
Date: 17April2020
Crap: Assigns home directories for users in specified OU. This will also
set the permissions on the folder to "modify" for the user account with
inheritanc and set the ownership to the user account. If folder does not
exist it will be created.
### Must have NTFSSecurity module installed!!!
### https://github.com/raandree/NTFSSecurity
#>

$ADLocation = "OU=Users,DC=ACME,DC=COM"
$HomeDirectory = "\\ACME.COM\Home_Folders"
$HomeDrive = "H:"
$Domain = "ACME"

$ADUsers = Get-ADUser -Filter * -SearchBase $ADLocation -Properties HomeDirectory,HomeDrive,SamAccountName,SID

ForEach($User in $ADUsers){
    If($User.Enabled -eq $true){
        $NewHomeFolder = ($HomeDirectory + $User.SamAccountName)
        $FolderUser = "$Domain\$($User.SamAccountName)"
        $UserPermissions = [System.Security.AccessControl.FileSystemRights]"Modify"
        $Inheritance = [System.Security.AccessControl.InheritanceFlags]::"ContainerInherit", "ObjectInherit"
        $Propagation = [System.Security.AccessControl.PropagationFlags]::None
        $AccessControl =[System.Security.AccessControl.AccessControlType]::Allow
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($FolderUser,$UserPermissions,$Inheritance,$Propagation,$AccessControl)

        If((Test-Path $NewHomeFolder) -eq $false){
            New-Item -ItemType Directory -Path $NewHomeFolder
        }

        $FolderOwner = Get-NTFSOwner -Path $NewHomeFolder
        If($FolderOwner.Account.Sid -ne $User.SID){
            Set-NTFSOwner -Path $NewHomeFolder -Account $User.SID
        }

        $FolderACL = Get-Acl -Path $NewHomeFolder
        $FolderACL.SetAccessRule($AccessRule)
        Set-Acl -Path $NewHomeFolder -AclObject $FolderACL
        Set-ADUser -Identity $User -HomeDrive $HomeDrive -HomeDirectory $NewHomeFolder
    }
}
