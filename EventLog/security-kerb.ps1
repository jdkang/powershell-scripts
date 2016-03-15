<#
    .SYNOPSIS
    Parses a domain controller's security log for user events.
    
    .DESCRIPTION
    This can be useful when you have extra auditing events enabled on your domain controller policy. This was written to parse down kerberos ticket issues from locked out accounts. 
#>
param(
    $username,
	$after,
	[switch]$ft
)

$a = Get-EventLog -LogName security -After $after | ? { $_.Message -match $username } | Select Index,TimeGenerated,InstanceID,Message | % {
	$regex = 'Client Address:\s+.*?(?<IP>\d\d\d.\d\d\d?.\d\d\d\d?.\d\d\d?)'
	$ipMatches = $_.Message | Select-String -Pattern $regex
	$ip = 'n/a'
	$ipMatches | select -ExpandProperty matches | % {		
		if ($_.groups["IP"].value) {
			$ip = $_.groups["IP"].value
			if ($ip) {
				$hostName = [System.Net.Dns]::GetHostbyAddress($ip).HostName
			}
		}
	}
	$_ | Select *,@{N='hostname';E={$hostName}}
}

if ($ft) {
$a | ft @{N="Message";E={$_.Message};Width=30},Index,@{N='Time';E={$_.TimeGenerated};Width=18},InstanceID
} else {
$a
}