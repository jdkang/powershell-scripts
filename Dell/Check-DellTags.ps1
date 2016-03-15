param(
	[string[]]$tags,
	[string]$tagTextFile=""
)


function Get-DellWarrenty
{
param(
	[string]$serial=""
)

$service = New-WebServiceProxy -Uri http://143.166.84.118/services/assetservice.asmx?WSDL
if(!$serial)
{
    $system = Get-WmiObject win32_SystemEnclosure
    $serial = $system.serialnumber
}

$isDell = $false
if (!($serial.trim().tolower() -eq "n/a"))
{
	$guid = [guid]::NewGuid()
	$info = $service.GetAssetInformation($guid,'check_warranty.ps1',$serial)
	if($info.count -eq 0)
	{
		$isDell = $false
	}
	else
	{
		$isDell = $true
		$warranty = $info[0].Entitlements[0]
		$expires = $warranty.EndDate
		$days_left = $warranty.DaysLeft
		if($days_left -eq 0)
		{
			$expired = $true
		}
		else{
			$expired = $false
		}
	}
}

$dellObjProp = @{
	tag = $serial
	isDell = $isDell
	isExpired = $expired
	daysLeft = $days_left
	expires = $expires
}
New-Object PsObject -Property $dellObjProp

}


if ($tagTextFile)
{
	if(Test-Path $tagTextFile)
	{
		$content = gc $tagTextFile
		foreach($line in $content)
		{
			if($line.trim() -ne "")
			{
				$tags += $line.trim()
			}
		}
	}
	else
	{
		write-warning "i/o error with find $tagTextFile"
	}
}

Foreach($tag in $tags)
{
	Get-DellWarrenty -serial $tag

}