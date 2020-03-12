<#
Author: Stan Crider
Date: 28Jan2019
What this crap does:
Get the DFS-R replication schedule for replication groups in a specified domain. Highlight green if custom schedule
is applied, highlight yellow if schedule is full-time-no-throttle and highlight red if schedule does not fit either.
### The 'BandwidthDetail' property of the 'Get-DfsrGroupSchedule' object is in a single string format. Each hexadecimal
digit represents a 15 minute block of time that begins with Sunday morning at 12am and ends on Saturday night at 11:59pm.
The hexadecimal codes translate as follows:
0 = no replication
1 = 16 Kbps
2 = 64 Kbps
3 = 128 Kbps
4 = 256 Kbps
5 = 512 Kbps
6 = 1 Mbps
7 = 2 Mbps
8 = 4 Mbps
9 = 8 Mbps
A = 16 Mbps
B = 32 Mbps
C = 64 Mbps
D = 128 Mbps
E = 256 Mbps
F = Full; no throttle
#>

# ====================================================================================================================
# Enter variables below
# ====================================================================================================================

# Specify domain to gather DFSR groups from
$Domain = "acme.com"

# Custom built schedule to check against; Weekends: no throttle, weekdays: 6am-6pm 256Kbps, 6pm-6am no throttle
$CustomSchedule = (("F"*120) + (("4"*48) + ("F"*48))*5 + ("F"*72))

# ====================================================================================================================
# End of variables; no customizable variables beyond this point!
# ====================================================================================================================

# Full/no throttle schedule
$FullSchedule = ("F"*672)

# Create dividers and schedule headers
$WriteLine = ("-"*107)
$AMPMMarker = ("|" + (" "*23) + "AM" + (" "*22) + "|" + (" "*23) + "PM")
$TimeMarker = (("|" + " "*3 + "1" + " "*7 + "3" + " "*7 + "5" + " "*7 + "7" + " "*7 + "9" + " "*7 + "11" + " "*2)*2)
$HourMarker = ("|   "*24)

# Get DFSR Groups and schedules
$Groups = Get-DfsReplicationGroup -GroupName * -DomainName $Domain | Sort-Object GroupName
ForEach($Group in $Groups){
    $GroupSchedule = Get-DfsrGroupSchedule -GroupName $Group.GroupName
    Switch($GroupSchedule.BandwidthDetail){
        $CustomSchedule{
            $FGColor = "Green"
        }
        $FullSchedule{
            $FGColor = "Yellow"
        }
        Default{
            $FGColor = "Red"
        }
    }
    Write-Output $WriteLine
    Write-Host("Schedule for group " + $Group.GroupName + ":") -ForegroundColor $FGColor
    Write-Output ((" "*11) + $AMPMMarker)
    Write-Output ((" "*11) + $TimeMarker)
    Write-Output ((" "*11) + $HourMarker)
    Write-Output ("Sunday:    " + ($GroupSchedule.BandwidthDetail).Substring(0,96))
    Write-Output ("Monday:    " + ($GroupSchedule.BandwidthDetail).Substring(96,96))
    Write-Output ("Tuesday:   " + ($GroupSchedule.BandwidthDetail).Substring(192,96))
    Write-Output ("Wednesday: " + ($GroupSchedule.BandwidthDetail).Substring(288,96))
    Write-Output ("Thursday:  " + ($GroupSchedule.BandwidthDetail).Substring(384,96))
    Write-Output ("Friday:    " + ($GroupSchedule.BandwidthDetail).Substring(480,96))
    Write-Output ("Saturday:  " + ($GroupSchedule.BandwidthDetail).Substring(576,96))
    Write-Output $WriteLine
}

# Write group count
Write-Output ("Total groups: " + ($Groups | Measure-Object).Count)
