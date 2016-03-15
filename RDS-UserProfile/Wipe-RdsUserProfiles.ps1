<#
    .SYNOPSIS
    Wipe RDS user profiles from a UNC path and locally on each Termianl Server.

    .DESCRIPTION
    This was written in an env without psremoting enabled, ergo the PsExec wrapper.
#>
param(
	[string[]]$users,
    $roamingProfilesBase = @(New-Object PsObject -Property @{
        uncServer = 'XYZ-NAS42'
        localPath = 'P:\foo\bar\'
        UNCPath = '\\XYZ-NAS42\PROFILES$'
        takeown = $false
    }),
    [string]$farmName = "rdsfarm",
    [string[]]$farmServers = @('10.128.0.2','10.128.0.3'),
	[switch]$skipTests,
	[switch]$test
)

###########################################################################################
# generic functions
###########################################################################################
function Invoke-RemoteExpressionWithPsExec
{
    param(
        ## The computer on which to invoke the command.
        $ComputerName = "\\$ENV:ComputerName",
 
        ## The scriptblock to invoke on the remote machine.
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock,
 
        ## The username / password to use in the connection
        $Credential,
 
        ## Determines if PowerShell should load the user's PowerShell profile
        ## when invoking the command.
        [switch] $NoProfile
    )
 
    Set-StrictMode -Version Latest
 
    ## Prepare the command line for PsExec. We use the XML output encoding so
    ## that PowerShell can convert the output back into structured objects.
    ## PowerShell expects that you pass it some input when being run by PsExec
    ## this way, so the 'echo .' statement satisfies that appetite.
    $commandLine = "echo . | powershell -Output XML "
 
    if($noProfile)
    {
        $commandLine += "-NoProfile "
    }
 
    ## Convert the command into an encoded command for PowerShell
    $commandBytes = [System.Text.Encoding]::Unicode.GetBytes($scriptblock)
    $encodedCommand = [Convert]::ToBase64String($commandBytes)
    $commandLine += "-EncodedCommand $encodedCommand"
 
    ## Collect the output and error output
    $errorOutput = [IO.Path]::GetTempFileName()
 
    if($Credential)
    {
        ## This lets users pass either a username, or full credential to our
        ## credential parameter
        $credential = Get-Credential $credential
        $networkCredential = $credential.GetNetworkCredential()
        $username = $networkCredential.Username
        $password = $networkCredential.Password
 
        $output = psexec $computername /user $username /password $password `
            /accepteula cmd /c $commandLine 2>$errorOutput
    }
    else
    {
        $output = psexec /acceptEula $computername `
            cmd /c $commandLine 2>$errorOutput
    }
 
    ## Check for any errors
    $errorContent = Get-Content $errorOutput
    Remove-Item $errorOutput
    if($errorContent -match "(Access is denied)|(failure)|(Couldn't)")
    {
        $OFS = "`n"
        $errorMessage = "Could not execute remote expression. "
        $errorMessage += "Ensure that your account has administrative " +
            "privileges on the target machine.`n"
        $errorMessage += ($errorContent -match "psexec.exe :")
 
        Write-Error $errorMessage
    }
 
    ## Return the output to the user
    $output
}

function wipe-location 
{
param(
	[string]$path
)

	# null
	if (Test-Path "C:\null") {
		rm C:\null -r -force
	}
	md C:\null | write-debug
	
	rm $path -r -force
	
	if (Test-Path $path) {
		robocopy "C:\null" "$path" /B /Purge | write-debug
	}
	rm $path -r -force
	if (Test-Path $path) {
		$path2 = $path + ".deleted"
		move-item $path $path2
	}
}

###########################################################################################
# Init
###########################################################################################
if (!$users) {
	write-warning "no users inputed"
	exit
}

$psexec = psexec 2>$foo
if (!$psexec) {
	write-warning "PsExec not available. Put it in your PATH env var or system32 folder"
}

# pull IPs from farm IP
write-host "Checking farm IP $farmName for IPs" -foregroundcolor green
$farmIPs = [System.Net.Dns]::GetHostAddresses($farmName) | ? { $_.IPAddressToString} | Select -expand IPAddressToString
if ($farmIPs) {
	write-host "Using IPs $farmIPs" -foregroundcolor green
	$farmServers = $farmIPs
} else {
	write-host "Using hard-coded entries $farmServers" -foregroundcolor yellow
}

# Verify
$serverList = @()

foreach ($server in $farmServers) {
	$hostname = [System.Net.Dns]::GetHostEntry($server) | Select -expand Hostname
	write-host "Server: $server (hostname: $hostname)" -foregroundcolor cyan
	if (!$skipTests) {
		if (Test-Connection $server -quiet -count 2) {
			write-host "Connection OK" -foregroundcolor green
			
			$serverName = Invoke-RemoteExpressionWithPsExec \\$server { $env:computername }
			write-host "Local Server Name: $serverName" -foregroundcolor green
			if ($serverName) {
				$serverList += New-Object PsObject -Property @{
					name = $serverName
					IP = $server
				}
			} else {
				write-warning "Unable to get ENV variable COMPUTERNAME from $server"
			}
			
		} else {
			write-warning "Unable to connect (ICMP) $server"
		}
	} else {
		$serverList += New-Object PsObject -Property @{
			IP = $server
			name = $hostname
		}
	}
}
$serverList = $serverList | Sort -Property IP


###########################################################################################
# script functions
###########################################################################################
function Wipe-Profiles
{
param(
	[string]$userName
)

	write-host "Processing: $userName" -foregroundcolor cyan -backgroundcolor black

	# Local profiles
	write-host "Deleting LOCAL profiles" -foregroundcolor black -backgroundcolor green
	foreach ($server in $serverList) {
		write-host "Deleting user profile on $($server.IP) $($server.name)" -foregroundcolor green
		$profile = $null
		$profile = (Get-WmiObject Win32_UserProfile -ComputerName $server.IP | ? {$_.LocalPath -like "*$userName*"})
		if ($profile) {
			if (!$test) {
				$profile.Delete()
			}
		} else {
			write-warning "No profile found for $userName on $($server.IP) / $($server.name)"
		}
	}

	foreach ($roamingProfile in $roamingProfilesBase) {
		write-host "checking roaming profile: $($roamingProfile.server) > $($roamingProfile.UNCPath) / $($roamingProfile.localPath)" -foregroundcolor black -backgroundcolor green
		# sometimes you get a "v2" etc profile name
		$roamingProfileNames = @()
		$standardRoamingProfileName = "$($roamingProfile.UNCPath)\$userName"
		$possibleRoamingProfileNames = gci $roamingProfile.UNCPath | ? { $_.Name -like "$userName*" } | Select -Expand Name
		
		if ($possibleRoamingProfileNames) {
			if ( ($possibleRoamingProfileNames | Measure).Count -eq 1 ) {
				if ($possibleRoamingProfileNames -eq $userName) {
					$roamingProfileNames = @( $userName )
				}
			} 
			
			if (!$roamingProfileNames) {
				$possibleRoamingProfileNames
				$answer = Read-Host "Is this list of Profile(s) correct for $($userName)?"
				if (!($answer -like "y*")) {
					exit
				} else {
					$roamingProfileNames = $possibleRoamingProfileNames
				}
			}
		} else {
			write-warning "No potential roaming profile names found."
		}
		
		
		if ($roamingProfileNames) {
			write-host "Starting profile deletions" -foregroundcolor green
			foreach ($profileName in $roamingProfileNames) {
				if ($roamingProfile.takeown) {
					# takeown
					write-host "taking ownership of roaming $($roamingProfile.server) > $($roamingProfile.localPath)\$profileName" -foregroundcolor green
					if (!$test) {
						psexec -s "\\$($roamingProfile.uncServer)" takeown /f "$($roamingProfile.localPath)\$profileName" /a /r /D Y | write-debug
					}
				}
				
				# delete roaming profile (redirected)
				write-host "deleting roaming profile $($roamingProfile.UNCPath)\$profileName" -foregroundcolor green
				if (!$test) {
					wipe-location -path "$($roamingProfile.UNCPath)\$profileName"
				}
			}
		}
		
		#appdata
		if ($roamingProfile.appdataUNC) {
			write-host "clearing $($roamginProfile.appdataUNC)" -foregroundcolor green
			if (!$tst) {
				wipe-location -path "$($roamingProfile.appdataUNC)\$userName"
			}
		}
	}
}

###########################################################################################
# main
###########################################################################################
foreach ($user in $users) {
	Wipe-Profiles -userName $user
}


