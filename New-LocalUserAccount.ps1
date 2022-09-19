<#
Author: Stan Crider
Date: 19Sept2022
Crap: Create a local user on specified remote computers and add user to local groups
### Requires administrative access to specified computers!!!
#>

## Variables
# List of computers FQDN (computer.domain)
$Computers = @(
    "computer-1.acme.com"
    "computer-2.acme.com"
)

# List of local accounts to create on remote computer; Accounts should have 3 properties: Name, FullName, Description
# Create an account custom-object in the $AccountName array for each local user that should be created
$AccountNames = @()
$AccountNames += [PSCustomObject]@{
    "Name"        = "newuser1"
    "FullName"    = "NewUser1"
    "Description" = "new user account for testing"
}

$AccountNames += [PSCustomObject]@{
    "Name"        = "newuser2"
    "FullName"    = "NewUser2"
    "Description" = "new user account for testing"
}

# List of local groups to add local accounts to
$LocalGroups = @(
    "Administrators"
    "Remote Users"
)

## Script
ForEach($Computer in $Computers){
    # Ping to ensure computer is online
    If(Test-Connection -ComputerName $Computer -Quiet -Count 2){
        $ConnectionError = $null
        Write-Host "Processing $Computer..."
        # Try to open remote session
        Try{
            $RemoteSession = New-PSSession -ComputerName $Computer -ErrorAction Stop
        }
        Catch{
            $ConnectionError = "Remote session to computer $Computer refused. No attempt to add local user(s) will be made."
        }
        If($ConnectionError){
            Write-Warning $ConnectionError
        }
        # If remote session open, run script blocks
        Else{
            ForEach($AccountObject in $AccountNames){
                # Disseminate local user properties to separate variables for script block use
                $AccountName =  $AccountObject.Name
                $AccountFullName = $AccountObject.FullName
                $AccountDescription = $AccountObject.Description

                # Run local user check script block
                Invoke-Command -Session $RemoteSession -ScriptBlock{
                    $ErrorMessage = $null
                    # Check if local user already exists; create if not
                    Try{
                        $UserExists = Get-LocalUser -Name $using:AccountName -ErrorAction Stop
                    }
                    Catch{
                        $UserExists = $null
                    }
                    If($UserExists){
                        Write-Warning ("The local account " + $using:AccountName + " already exists on computer " + $using:Computer + ".")
                    }
                    Else{
                        Try{
                            New-LocalUser -Name $using:AccountName -FullName $using:AccountFullName -Description $using:AccountDescription -NoPassword -ErrorAction Stop | Out-Null
                        }
                        Catch{
                            $ErrorMessage = ("The local account " + $using:AccountName + " was not created on " + $using:Computer + ".")
                            Write-Warning $ErrorMessage
                        }
                    }
                }
                # If no errors in user creation block, run group membership block
                If(-Not($ErrorMessage)){
                    ForEach($LocalGroup in $LocalGroups){
                        # Check if group exists
                        Invoke-Command -Session $RemoteSession -ScriptBlock{
                            Try{
                                $GroupExists = Get-LocalGroup -Name $using:LocalGroup -ErrorAction Stop
                            }
                            Catch{
                                $GroupExists = $null
                            }
                            If($GroupExists){
                                # Get membership of group
                                $LocalGroupMembers = Get-LocalGroupMember -Group $using:LocalGroup
                                # Add local user to group if not already a member
                                $FullLocalUserName = ("" + (($using:Computer).Split(".")[0]) + "\" + $using:AccountName)
                                If($LocalGroupMembers.Name -contains $FullLocalUserName){
                                    Write-Warning ("The user " + $using:AccountName + " is already in the group " + $using:LocalGroup + " on computer " + $using:Computer + ".")
                                }
                                Else{
                                    Try{
                                        Add-LocalGroupMember -Group $using:LocalGroup -Member $using:AccountName -ErrorAction Stop
                                    }
                                    Catch{
                                        Write-Warning ("Failed to add user " + $using:AccountName + " to group " + $using:LocalGroup + " on " + $using:Computer + ".")
                                    }
                                }
                            }
                            Else{
                                Write-Warning ("The group " + $using:LocalGroup + " does not exist on " + $using:Computer + ".")
                            }
                        }
                    }
                }
            }
            # Close session
            Exit-PSSession
        }
    }
    Else{
        Write-Warning "The computer $Computer is not responding to ping."
    }
}
