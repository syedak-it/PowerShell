################################################################################
# Author          : Syed Abdul Khader 
# Description     : Get hardware inventory and send as HTML report in email
################################################################################
#param( [string] $auditlist)
Function Get-CustomHTML ($Header){
$Report = @"
<!DOCTYPE HTML>
<html><head><title>$($Header)</title>
<META http-equiv=Content-Type content='text/html; charset=windows-1252'>
<meta name="save" content="history">
<style type="text/css">
DIV .expando {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 8pt; RIGHT: 8px; COLOR: #ffffff; FONT-FAMILY: Arial; POSITION: absolute; TEXT-DECORATION: underline}
TABLE {TABLE-LAYOUT: fixed; FONT-SIZE: 100%; WIDTH: 100%}
*{margin:0}
.dspcont { display:none; BORDER-RIGHT: #B1BABF 1px solid; BORDER-TOP: #B1BABF 1px solid; PADDING-LEFT: 16px; FONT-SIZE: 8pt;MARGIN-BOTTOM: -1px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; WIDTH: 95%; COLOR: #000000; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; BACKGROUND-COLOR: #f9f9f9}
.filler {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 43px; BORDER-LEFT: medium none; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative}
.save{behavior:url(#default#savehistory);}
.dspcont1{ display:none}
a.dsphead0 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #FFFFFF; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #CC0000}
a.dsphead1 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead2 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead1 span.dspchar{font-family:monospace;font-weight:normal;}
td {VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma}
th {VERTICAL-ALIGN: TOP; COLOR: #CC0000; TEXT-ALIGN: left}
BODY {margin-left: 4pt} 
BODY {margin-right: 4pt} 
BODY {margin-top: 6pt} 
</style>
<script type="text/javascript">
function dsp(loc){
   if(document.getElementById){
      var foc=loc.firstChild;
      foc=loc.firstChild.innerHTML?
         loc.firstChild:
         loc.firstChild.nextSibling;
      foc.innerHTML=foc.innerHTML=='hide'?'show':'hide';
      foc=loc.parentNode.nextSibling.style?
         loc.parentNode.nextSibling:
         loc.parentNode.nextSibling.nextSibling;
      foc.style.display=foc.style.display=='block'?'none':'block';}}  
if(!document.getElementById)
   document.write('<style type="text/css">\n'+'.dspcont{display:block;}\n'+ '</style>');
</script>
</head>
<body>
<b><font face="Arial" size="5">$($Header)</font></b><hr size="8" color="#CC0000">
<font face="Arial" size="1">Report created on $(Get-Date)</br></font>
<div class="filler"></div>
<div class="filler"></div>
<div class="filler"></div>
<div class="save">
"@
Return $Report
}

Function Get-CustomHeader0 ($Title){
$Report = @"
		<h1><a class="dsphead0">$($Title)</a></h1>
	<div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader ($Num, $Title){
$Report = @"
	<h2><a href="javascript:void(0)" class="dsphead$($Num)" onclick="dsp(this)">
	<span class="expando">show</span>$($Title)</a></h2>
	<div class="dspcont">
"@
Return $Report
}

Function Get-CustomHeaderClose{

	$Report = @"
		</DIV>
		<div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader0Close{

	$Report = @"
</DIV>
"@
Return $Report
}

Function Get-CustomHTMLClose{

	$Report = @"
</div>
</body>
</html>
"@
Return $Report
}

Function Get-HTMLTable{
	param([array]$Content)
	$HTMLTable = $Content | ConvertTo-Html
	$HTMLTable = $HTMLTable -replace '<!DOCTYPE html>', ""
	$HTMLTable = $HTMLTable -replace '<head>', ""
	$HTMLTable = $HTMLTable -replace '<title>HTML TABLE</title>', ""
	$HTMLTable = $HTMLTable -replace '</head><body>', ""
	$HTMLTable = $HTMLTable -replace '</body></html>', ""
	Return $HTMLTable
}

Function Get-HTMLGeneral ($Heading, $Detail){
$Report = @"
<TABLE>
    <style>
    table, th, td { }
    </style>
	<tr>
	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%'>$($Detail)</td>
	</tr>
</TABLE>
"@
Return $Report
}

Function Get-HTMLDetail ($Heading, $Detail){
$Report = @"
<TABLE>
	<tr>
	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%'>$($Detail)</td>
	</tr>
</TABLE>
"@
Return $Report
}

Function Get-HTMLRolesDetail ($Heading, $Detail){
$Report = @"
<TABLE>
	<tr>
   	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%>$($Detail)</td>
	</tr>
</TABLE>
"@
Return $Report
}

$target = $env:computername
$ComputerSystem = Get-WmiObject -computername $Target Win32_ComputerSystem
$ComputerBIOS = Get-WmiObject Win32_Bios
$ComputerProcessor = Get-WmiObject Win32_Processor
$OperatingSystems = Get-WmiObject -computername $Target Win32_OperatingSystem
$TimeZone = Get-WmiObject -computername $Target Win32_Timezone
$Keyboards = Get-WmiObject -computername $Target Win32_Keyboard
$SchedTasks = Get-WmiObject -computername $Target Win32_ScheduledJob
$BootINI = $OperatingSystems.SystemDrive + "boot.ini"
$RecoveryOptions = Get-WmiObject -computername $Target Win32_OSRecoveryConfiguration
$Computeruser=(Get-WMIObject -ClassName Win32_ComputerSystem).Username

switch ($ComputerRole){
	"Member Workstation" { $CompType = "Computer Domain"; break }
	"Domain Controller" { $CompType = "Computer Domain"; break }
	"Member Server" { $CompType = "Computer Domain"; break }
	default { $CompType = "Computer Domain"; break }
}
$Adapters = Get-WmiObject -ComputerName $Target Win32_NetworkAdapterConfiguration
$IP = $null
Foreach ($Adapter in ($Adapters | Where {$_.IPEnabled -eq $True})) 
{
    If ($IP -eq $null)
    {
        $IP = "$($Adapter.IPAddress)"
    }
    Else
    {
        $IP = -Join($IP, ", ","$($Adapter.IPAddress)")
    }
}
$LBTime=$OperatingSystems.ConvertToDateTime($OperatingSystems.Lastbootuptime)
####################################################################################
# General Settings
####################################################################################
$MyReport = Get-CustomHTML "$Target Laptop Info"
$MyReport += Get-CustomHeader0  "$Target Details"
$MyReport += Get-CustomHeader "2" "General"
	$MyReport += Get-HTMLGeneral "Computer Name" ($ComputerSystem.Name)
	$MyReport += Get-HTMLGeneral $CompType ($ComputerSystem.Domain)
	$MyReport += Get-HTMLGeneral "Operating System" ($OperatingSystems.Caption)
	$MyReport += Get-HTMLGeneral "Service Pack" ($OperatingSystems.CSDVersion)
	$MyReport += Get-HTMLGeneral "Manufacturer" ($ComputerSystem.Manufacturer)
	$MyReport += Get-HTMLGeneral "Model" ($ComputerSystem.Model)
	$MyReport += Get-HTMLGeneral "Service Tag" ($ComputerBIOS.SerialNumber)
    $ServiceTag = $ComputerBIOS.SerialNumber
    $MyReport += Get-HTMLGeneral "Processor" ($ComputerProcessor.Name)
    $TotPhysicalMemory = (([Math]::Round($ComputerSystem.TotalPhysicalMemory/1GB,0)).tostring()) + " GB"
	$MyReport += Get-HTMLGeneral "RAM" $TotPhysicalMemory
    $MyReport += Get-HTMLGeneral "Login User" $Computeruser
    $MyReport += Get-CustomHeaderClose
################################ End of General Settings############################
    
####################################################################################
# Logical Disks
####################################################################################
$Disks = Get-WmiObject -ComputerName $Target Win32_LogicalDisk
$MyReport += Get-CustomHeader "2" "Storage"
	$LogicalDrives = @()
	Foreach ($LDrive in ($Disks | Where {$_.DriveType -eq 3}))
    {
		$Details = "" | Select "Drive Letter", Label, "File System", "Disk Size (GB)", "Disk Free Space (GB)"
		$Details."Drive Letter" = $LDrive.DeviceID
		$Details.Label = $LDrive.VolumeName
		$Details."File System" = $LDrive.FileSystem
		$Details."Disk Size (GB)" = [math]::Round($LDrive.size / 1GB,2)
		$Details."Disk Free Space (GB)" = [math]::round($LDrive.FreeSpace / 1GB,2)
		$LogicalDrives += $Details
	}
	$MyReport += Get-HTMLTable ($LogicalDrives)
	$MyReport += Get-CustomHeaderClose
################################ End of Disks#####################################

####################################################################################
# Software
####################################################################################
If ((get-wmiobject -ComputerName $Target -namespace "root/cimv2" -list) | Where-Object {$_.name -match "Win32_Product"})
{
	$MyReport += Get-CustomHeader "2" "Software"
		$MyReport += Get-HTMLTable (get-wmiobject -ComputerName $Target Win32_Product | Select Name,Version,Vendor)
	$MyReport += Get-CustomHeaderClose
}
################################ End of Software ####################################

####################################################################################
# Format Report
####################################################################################


$MyReport += Get-CustomHeader0Close
$MyReport += Get-CustomHTMLClose
$MyReport += Get-CustomHTMLClose

$reportime=get-date -uFormat "%d-%m-%Y %H:%M"
$Path="C:\Temp\"
If ((Test-Path $Path) -eq $false)
{
    New-Item $Path -type directory
}
$fdate=get-Date -format "dd-MM-yy_HHmm"
$Filename = $Path+"$($Target)_$fdate.html"
$MyReport | out-file -encoding ASCII -filepath $Filename

####################################################################################
# Email Report
####################################################################################
$Computeruser=((Get-WMIObject -ClassName Win32_ComputerSystem).username)

# Mail the output file
$mailParams = @{
SmtpServer                 = 'smtpserver.com'
Port                       = '587' 
From                       = 'sender@smtpserver.com'
To                         = 'receiver@smtpserver.com'
Subject                    = "Inventory for $ServiceTag"
Body                       = "Hi, `n`nPlease find attachment laptop details of $ServiceTag `n`nThanks & Regards"
Attachment                 = $Filename # Here you can change your attachment
#DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}
Send-MailMessage @mailParams
