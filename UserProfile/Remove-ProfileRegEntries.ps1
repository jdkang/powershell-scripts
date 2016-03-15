param(
	$sid,
	[switch]$list
)

if (!$sid) {
	write-warning "no SID specified. you can specify a partial SID"
	exit
}

write-host "WMI win32_userprofile" -foregroundcolor green
$profileWmi = Get-WmiObject win32_userprofile | ? { $_.Sid -like "*$sid*" }
$profileWmi
if ($profileWmi) {
	if (!$list) {
		$profileWmi.Delete()
	}
}

write-host "HKLM:\Software\microsoft\windows nt\CurrentVersion\ProfileList" -foregroundcolor green

$profile = gci "HKLM:\Software\microsoft\windows nt\CurrentVersion\ProfileList" | ?{ $_.Name -like "*$sid*" }
$profile
if ($profile) {
	if (!$list) {
		$profile | remove-item -force -confirm:$false
	}
}

write-host "HKLM:\Software\microsoft\windows nt\CurrentVersion\ProfileGUID" -foregroundcolor green

$profileguid = gci "HKLM:\Software\microsoft\windows nt\CurrentVersion\ProfileGUID" | Get-ItemProperty | ?{ $_.SidString -like "*$sid*" }
$profileguid
if ($profileguid) {
	if (!$list) {
		$profileguid | remove-item -force -confirm:$false
	}
}

write-host "HKLM:\Software\microsoft\windows nt\CurrentVersion\ProfileGUID" -foregroundcolor green
$policyguid = gci "HKLM:\Software\microsoft\windows nt\CurrentVersion\ProfileGUID" | Get-ItemProperty | ?{ $_.SidString -like "*$sid*" }
$policyguid
if ($policyguid) {
	if (!$list) {
		$policyguid | remove-item -force -confirm:$false
	}
}