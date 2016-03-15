param(
	[string[]]$computers=@(),
	[string]$textFile=""
)
ipmo activedirectory

if ($textFile) {
	if (!(Test-Path $textFile)) {
		write-warning "$textFile does not exist."
		exit
	}
}

if ($textFile) {
	$additionalComputers = gc $textFile
	
	foreach ($additionalComputer in $additionalComputers) {
		if ($additionalComputer) {
			$computers += $additionalComputer.trim()
		}
	}
}

foreach ($computer in $computers) {
	Get-ADComputer $computer -Properties LastLogonTimeStamp,description |
	Select Name,description,@{Name="lastlogon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}
}

