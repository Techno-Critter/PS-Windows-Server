<#
Author: Stan Crider
Date: 12 Sept 2022
Crap: starts and/or restrarts a specified service on specified computers
#>

## Variables
# Service name (not display name) of service
$ServiceName = "AGPM Service"
# list of computers to run against
$Computers = @(
    "test1.acme.com"
    "test2.acme.com"
)

## Script
ForEach($Computer in $Computers){
    # Check if computer is online
    If(Test-Connection -ComputerName $Computer -Quiet -Count 2){
        # Connect to computer and run script
        $PSSession = New-PSSession -ComputerName $Computer
        Invoke-Command -Session $PSSession -ScriptBlock{
            $ErrorMessage = $null
            Try{
                # Check if service is installed and report status
                $ServiceStatus = Get-Service -Name $using:ServiceName -ErrorAction Stop
            }
            Catch{
                $ErrorMessage = "The service $using:ServiceName on $using:Computer does not exist."
            }
            If($ErrorMessage){
                Write-Warning $ErrorMessage
            }
            Else{
                Switch($ServiceStatus.Status){
                    # Start service if stopped
                    {$_ -eq "Stopped"}{
                        Write-Host "The service `"$using:ServiceName`" on $using:Computer is starting..."
                        Try{
                            Start-Service -Name $using:ServiceName -ErrorAction Stop
                        }
                        Catch{
                            Write-Warning "The service `"$using:ServiceName`" on $using:Computer failed to start."
                        }
                    }
                    # Restart service if running
                    {$_ -eq "Running"}{
                        Write-Host "The service `"$using:ServiceName`" on $using:Computer is restarting..."
                        Try{
                            Restart-Service -Name $using:ServiceName -ErrorAction Stop
                        }
                        Catch{
                            Write-Warning "The service `"$using:ServiceName`" on $using:Computer failed to restart."
                        }
                    }
                    # Report status if not stopped or running
                    Default{
                        Write-Warning ("The status of service `"$using:ServiceName`" on $using:Computer is: " + $ServiceStatus.Status)
                    }
                }
            }
        }
        # Close connection
        Exit-PSSession
    }
    Else{
        Write-Warning "The computer $Computer is not responding to pings. No connection attempt made."
    }
}
