<#
Run specified programs and verifies network drive connectivity upon PS open
Copy contents of this script into PS profile
From PS window, type "notepad $PROFILE", insert contents, save, and close
Remove or remark out any programs or consoles that are not needed; add as necessary
#>

# Map drive
$DriveMappings = @(
    @{Letter="H";Path="\\fileserver\home\username"},
    @{Letter="O";Path="\\fileserver\operations"},
    @{Letter="S";Path="\\fileserver\share"}
    )

ForEach($DriveMapping in $DriveMappings){
    $CurrentDrives = Get-PSDrive | Where-Object{$_.Name -eq $DriveMapping.Letter}
    If($null -eq $CurrentDrives){
        New-PSDrive -Name $DriveMapping.Letter -Root $DriveMapping.Path -Persist -PSProvider "FileSystem"
    }
    Else{
        Write-Output "Drive $($DriveMapping.Letter) already in use."
    }
}

# Open  MMC programs
$MMCProcess = Get-Process -Name "MMC" | Select-Object *

# DFS
If(($null -eq $MMCProcess) -or ("DFS Management" -notin $MMCProcess.MainWindowTitle)){
	dfsmgmt.msc
}

# ADUC
If(($null -eq $MMCProcess) -or ("Active Directory Users and Computers" -notin $MMCProcess.MainWindowTitle)){
	dsa.msc
}

# DNS
If(($null -eq $MMCProcess) -or ("DNS Manager" -notin $MMCProcess.MainWindowTitle)){
	dnsmgmt.msc
}

# DHCP
If(($null -eq $MMCProcess) -or ("DHCP" -notin $MMCProcess.MainWindowTitle)){
	dhcpmgmt.msc
}

# Sites and Services
If(($null -eq $MMCProcess) -or ("Active Directory Sites and Services" -notin $MMCProcess.MainWindowTitle)){
	dssite.msc
}

# GPO
If(($null -eq $MMCProcess) -or ("Group Policy Management" -notin $MMCProcess.MainWindowTitle)){
	gpmc.msc
}

# Hyper-V Manager
If(($null -eq $MMCProcess) -or ("Hyper-V Manager" -notin $MMCProcess.MainWindowTitle)){
	virtmgmt.msc
}

# Domains and Trusts
If(($null -eq $MMCProcess) -or "Active Directory Domains and Trusts" -notin $MMCProcess.MainWindowTitle){
	domain.msc
}

# VS Code
$VSCodeProcess = Get-Process -Name "Code" -ErrorAction SilentlyContinue
If($null -eq $VSCodeProcess){	
	& "$ENV:PROGRAMFILES\Microsoft VS Code\bin\Code.cmd" "U:\Scripts\PowerShell"
}

# SCCM Console
$SCCMProcess = Get-Process -Name "Microsoft.ConfigurationManagement" -ErrorAction SilentlyContinue
If($null -eq $SCCMProcess){
	& "${ENV:PROGRAMFILES(x86)}\Microsoft Endpoint Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe"
}

# Volume Activation Tools
$VATProcess = Get-Process -Name "Volume Activation Tools" -ErrorAction SilentlyContinue
If($null -eq $VATProcess){
	& "$ENV:WINDIR\system32\vmw.exe"
}
