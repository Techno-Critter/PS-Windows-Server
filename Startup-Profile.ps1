<#
Run specified programs and verifies network drive connectivity upon PS open
Copy contents of this script into PS profile
From PS window, type "notepad $PROFILE", insert contents, save, and close
#>

# Map drive
$CurrentDrives = Get-PSDrive | Where-Object{$_.Name -eq "H"}
If($null -eq $CurrentDrives){
	New-PSDrive -Name "H" -Root "\\fileserver\home\username" -Persist -PSProvider "FileSystem"
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

# VS Code
$VSCodeProcess = Get-Process -Name "Code" -ErrorAction SilentlyContinue
If($null -eq $VSCodeProcess){	
	& "$ENV:PROGRAMFILES\Microsoft VS Code\bin\Code.cmd" "U:\Scripts\PowerShell"
}
