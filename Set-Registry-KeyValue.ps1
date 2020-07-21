<#
Author: Stan Crider
Date: 14July2020
Crap:
Sets the specified registry key on all domain controllers in specified domain and restarts specified service
#>

## Set user variables; specify the specifics
$DesiredRegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\DNS\Parameters\"
$DesiredKey = "TcpReceivePacketSize"
$DesiredValue = 65280
# Specify the service object name to be restarted
$DesiredService = "DNS"
# Specify domain or set to (Get-ADDomain).DNSRoot for current domain
$Domain = (Get-ADDomain).DNSRoot

## Script below
$Servers = Get-ADDomainController -Filter * -Server $Domain | Sort-Object HostName

ForEach($Server in $Servers){
    # Reset reusable variables for each round
    $CurrentValue = $null
    If($env:COMPUTERNAME -match $Server.name){
        $ServerName = "localhost"
    }
    Else{
        $ServerName = $Server.Name
    }

    # Test connectivity to server
    $Online = Test-Connection -ComputerName $ServerName -Quiet -Count 2
    If($Online){
        # Test registry path
        Try{
            $TestPath = Invoke-Command -ComputerName $ServerName -ScriptBlock {Test-Path -Path $using:DesiredRegistryPath} -ErrorAction Stop
        }
        Catch{
            $TestPath = $false
        }

        If($TestPath -eq $true){
            # Test registry key
            Try{
                Invoke-Command -ComputerName $ServerName -ScriptBlock {Get-ItemProperty -Path $using:DesiredRegistryPath  | Select-Object -ExpandProperty $using:DesiredKey} -ErrorAction Stop | Out-Null
                $TestKey = $true
            }
            Catch{
                $TestKey = $false
            }

            If($TestKey -eq $true){
                # Check value of registry key
                $CurrentValue = Invoke-Command -ComputerName $ServerName -ScriptBlock {(Get-ItemProperty -Path $using:DesiredRegistryPath  -Name $using:DesiredKey).$using:DesiredKey}

                If($CurrentValue -eq $DesiredValue){
                    # No changes necessary; abort
                    Write-Output ("The registry key $DesiredKey on " + $ServerName + " is already set to " + $DesiredValue)
                }
                Else{
                    # Change value of registry key
                    Write-Output ("The registry key $DesiredKey on " + $ServerName + " will be changed from " + $CurrentValue + " to " + $DesiredValue + ".")
                    Try{
                        Invoke-Command -ComputerName $ServerName -ScriptBlock {
                            Set-ItemProperty -Path $using:DesiredRegistryPath  -Name $using:DesiredKey -Value $using:DesiredValue
                            Restart-Service -Name $using:DesiredService
                        } -ErrorAction Stop
                    }
                    Catch{
                        Write-Error $_.Exception.Message
                    }

                }
            }
            Else{
                # Create key and set value
                Write-Output ("The registry key $DesiredKey will be created on " + $ServerName + ".")
                Try{
                    Invoke-Command -ComputerName $ServerName -ScriptBlock {
                        New-ItemProperty -Path $using:DesiredRegistryPath  -Name $using:DesiredKey -Value $using:DesiredValue
                        Restart-Service -Name $using:DesiredService
                    } -ErrorAction Stop
                }
                Catch{
                    Write-Error $_.Exception.Message
                }
            }
        }
        Else{
            # If registry path does not exist, abort
            Write-Warning ("The registry path $DesiredRegistryPath does not exist on " + $ServerName + ".")
        }
    }
    Else{
        # If server does not respond, abort
        Write-Warning ("The server " + $ServerName + " is not online.")
    }
}
