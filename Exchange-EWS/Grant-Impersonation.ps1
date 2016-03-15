param(
	[string]$ewsProxyUser,
	[string[]]$targetsToImpersonate
)

$casses = Get-ExchangeServer | ? { $_.IsClientAccessServer }
write-host "Checking ews proxy user ... $ewsProxyUser" -foregroundcolor yellow
$user = Get-User -Identity $ewsProxyUser
if (!$user) {
	write-warning "Could not find $user"
	exit
} else {
	write-host "$user" -foregroundcolor green
	write-host "Object = OK" -foregroundcolor green
}

write-host "Checking target impersonation users ...." -foregroundcolor yellow
$targetImpersonators = @()
foreach ($targetToImpersonate in $targetsToImpersonate) {
	$targeIimpersonator = $null
	$targeIimpersonator = Get-User -Identity $targetToImpersonate
	if (!$targeIimpersonator) {
		write-warning "Could not find $targetToImpersonate"
		exit
	} else {
		$targetImpersonators += $targeIimpersonator
		write-host "$targeIimpersonator" -foregroundcolor green
		write-host "object = OK" -foregroundcolor green
	}
}

write-host ""
write-host "Press any key to continue...." -foregroundcolor magenta
read-host

write-host "Setting impersonation permissions for $user" -foregroundcolor green
# Grant rights on each CAS
write-host "Granting ms-Exch-EPI-Impersonation right on CAS servers..." -foregroundcolor green
foreach ($CAS in $casses) {
	write-host "Adding permissions on $CAS" -foregroundcolor green
	Add-ADPermission -Identity $CAS.DistinguishedName -User $user.Identity -extendedRight ms-Exch-EPI-Impersonation
}

# Grant rights on user
write-host "Granting ms-Exch-EPI-May-Impersonate rights per user... "
foreach ($targeIimpersonator in $targetImpersonators) {
	write-host "Granting rights for $targeIimpersonator ..." -foregroundcolor green
	Add-ADPermission $targeIimpersonator.DistinguishedName -user $user.Identity -extendedRight ms-Exch-EPI-May-Impersonate
}

