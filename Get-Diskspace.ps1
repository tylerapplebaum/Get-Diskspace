[CmdletBinding(DefaultParameterSetName = "Set1")]
    param(
        [Parameter(mandatory=$true, parametersetname="Set1", HelpMessage="Specify the path to the list of computer names (C:\Scripts\list.txt)")]
		[Alias("l")]
        [string]$Complist,

        [Parameter(mandatory=$true, parametersetname="Set2", HelpMessage="Specify the Get-ADComputer name filter to apply (Use * for wildcard")]
		[Alias("f")]
        [string]$Filter,
		
		[Parameter(mandatory=$true, parametersetname="Set3", HelpMessage="If 'servers' is present, all computers with a Windows Server OS will be scanned.")]
		[Alias("servers")]
        [switch]$Svrs,
		
		[Parameter(mandatory=$true, parametersetname="Set4", HelpMessage="If 'VDI' is present, all computers prefixed with 'TC' will be scanned.")]
		[Alias("vdi")]
        [switch]$thin,
		
		[Parameter(mandatory=$false, HelpMessage="If -e is specified, the report will be emailed to the recipients.")]
		[Alias("e")]
        [switch]$Email
		
	)

# Continue even if there are errors
$ErrorActionPreference = "Continue";
$Company = "Your company here"
$date = get-date
$i = 0

function script:Input {
	If ($Complist){
	#Get content of file specified, trim any trailing spaces and blank lines
	$script:Computers = gc ($Complist) | where {$_ -notlike $null } | foreach { $_.trim() }
	}
	Elseif ($Filter) {
		If (!(Get-Module ActiveDirectory)) {
		Import-Module ActiveDirectory
		} #include AD module
	#Filter out AD computer objects with ESX in the name
	$script:Computers = Get-ADComputer -Filter {SamAccountName -notlike "*esx*" -AND Name -Like $Filter} | select -ExpandProperty Name | sort
	}
	Elseif ($Svrs) {
		If (!(Get-Module ActiveDirectory)) {
		Import-Module ActiveDirectory
		} #include AD module
	#Filter out AD computer objects with ESX in the name and includes computers with a Windows Server OS
	$script:Computers = Get-ADComputer -Filter {SamAccountName -notlike "*esx*" -AND OperatingSystem -Like "*Server*"} | select -ExpandProperty Name | sort
	}
	Elseif ($thin) {
		If (!(Get-Module ActiveDirectory)) {
		Import-Module ActiveDirectory
		} #include AD module
	#Filter out AD computer objects with ESX in the name and includes TC* for VDIs
	$script:Computers = Get-ADComputer -Filter {SamAccountName -notlike "*esx*" -AND Name -Like "TC*"} | select -ExpandProperty Name | sort
	}
}#end Input

function script:PingTest {
$script:TestedComps = @()
	foreach ($WS in $Computers){
		If (Test-Connection -count 1 -computername $WS -quiet){
		$script:TestedComps += $WS
		}
		Else {
		Write-Host "Cannot connect to $WS" -ba black -fo yellow
		}
	}#end foreach
}#end PingTest

function script:Report {
# Path to the report
$reportPath = "$env:userprofile\desktop\";
# Report name
	If ($Svrs){
	$reportName = "DiskSpaceRpt_Servers_$(get-date -format ddMMyyyy).html";
	}
	Elseif ($thin){
	$reportName = "DiskSpaceRpt_VirtualDesktops_$(get-date -format ddMMyyyy).html";
	}
	Else {
	$reportName = "DiskSpaceRpt_$(get-date -format ddMMyyyy).html";
	}
	$script:diskReport = $reportPath + $reportName
	If (Test-Path $diskReport){
    Remove-Item $diskReport
    }
}#end Report

. Input #Call input function
. Pingtest #Call PingTest function
. Report #Call report function

#Begin main processing
$container = @() #Empty array to hold data
foreach($computer in $testedcomps){
$disks = Get-WmiObject -ComputerName $computer -Class Win32_LogicalDisk -Filter "DriveType = 3"
$IP = gwmi -query "SELECT IPAddress FROM win32_NetworkAdapterConfiguration where IPEnabled='True'" -computer $computer
$VMPhy = get-wmiobject -class "Win32_ComputerSystem" -namespace "root\CIMV2" -computername $computer
$computer = $computer.toupper()
	foreach($disk in $disks){      
	$deviceID = $disk.DeviceID;
	$volName = $disk.VolumeName;
	[float]$size = $disk.Size;
	[float]$freespace = $disk.FreeSpace; 
	$percentFree = [Math]::Round(($freespace / $size) * 100, 2);
	$sizeGB = [Math]::Round($size / 1GB, 2);
	$freeSpaceGB = [Math]::Round($freespace / 1GB, 2);
	$usedSpaceGB = [Math]::Round($sizeGB - $freeSpaceGB, 2);
	$IPAdd = [string]$IP.IPAddress[0] 
		If ($VMPhy.Manufacturer -eq "VMware, Inc.") {
		$Type = "Virtual"
		}
		Else {
		$Type = "Physical"
		}
	$newobj = $null
	$newobj = new-object psobject
	$newobj | add-member -membertype noteproperty -name "Computer" -value $computer
	$newobj | add-member -membertype noteproperty -name "IP" -value $IPAdd
	$newobj | add-member -membertype noteproperty -name "Platform" -value $Type
	$newobj | add-member -membertype noteproperty -name "Drive" -value $deviceID
	$newobj | add-member -membertype noteproperty -name "Drive Label" -value $volName
	$newobj | add-member -membertype noteproperty -name "Total Capacity" -value $sizeGB
	$newobj | add-member -membertype noteproperty -name "Used Capacity" -value $usedSpaceGB
	$newobj | add-member -membertype noteproperty -name "Free Space" -value $freeSpaceGB
	$newobj | add-member -membertype noteproperty -name 'Free Space %' -value $percentFree`%	
	$container += $newobj
	}#end foreach disk in disks
$i++
Write-Progress -Activity "Gathering data" -status "Processed computer $i of $($testedcomps.count)" -percentComplete ($i / $testedcomps.count*100)
}#end main foreach loop

#Path to CSS file
$CSS = "Style.css"

#Report title and legend
$PreContent = @"
<h2>$Company Disk Space Report $(get-date -uformat "%m-%d-%Y - %A")</h2>
<h3><table><tr><td>Legend:</td><td class="key7">15 percent free</td><td class="key6">10 percent free</td><td class="key5">5 percent free</td></tr></table></h3>
"@

#Javascript w/ conditional formatting
$PostContent = @'
<script type="text/javascript">
var t = document.querySelectorAll('td'), l = t.length;
for (var i = 0; i < l; i++) {
	var td = t[i];
	if (td.textContent.indexOf('%') == -1) continue;
	var val = parseInt(td.textContent.replace('%', ''), 10);
	if (val < 5) td.style.backgroundColor = 'red';
	else if (val < 10) td.style.backgroundColor = 'yellow';
	else if (val < 15) td.style.backgroundColor = 'orange';
}
</script>
'@

#Convert the PSCustomObject $container to HTML
$container | ConvertTo-HTML -CSSUri $CSS -Title "$Company Disk Report" -PreContent $PreContent -PostContent $PostContent | SC $diskReport

#Sends the message and attachment.	
If ($email) {
#Declare our message properties using splatting. Attach the file and specify recipients.
$MailMessage = @{
To = "ITDEPT <ITDEPT@yourcompany.com>"
From = "ITDEPT <ITDEPT@yourcompany.com>"
Subject = "Disk Space Report"
Body = "Disk Space Report $date"	
Attachments = $diskreport
Smtpserver = "EXCH.yourcomapny.corp.com"
}
Send-MailMessage @MailMessage
Write-Host "Report complete; e-mail sent." -fore green -back black
}