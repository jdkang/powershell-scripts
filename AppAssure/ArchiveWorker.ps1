#############################
# Loads Assembly 			#
#############################

function LoadAssembly($AssemblyName)
{
	try {
	$coreProcess = Get-Process "Core.Service"
	$exeInfo = New-Object System.IO.FileInfo($coreProcess.Path)	
	[system.reflection.assembly]::loadfrom([system.IO.Path]::Combine($exeInfo.Directory, $AssemblyName))  |  out-null
	}
	catch {
			$_ | slog -error -buffer
	}
}

function FormatJobRate($trasferRate)
{
	LoadAssembly "Containers.Implementation.dll"
	
	if($trasferRate -eq 0)	{
		
		$rate = "Rate hasn't calculated yet"
	}
	else {
		try {
			$rate = [Replay.Common.Containers.Implementation.Buffers.ByteCount]::Format([System.UInt64]$trasferRate) + "/s"
		}
		catch {
			"Problems with formatting rate value" | slog -error
			$rate = "N/A"
		}
	}
	
	return $rate
}

function FormatJobPercentDone($progress, $totalWork)
{
	if($totalWork -le 0) {		
		return "0%"
	}
	else {
		return [string]::Format("{0:F1}%", ($progress / $totalWork) * 100)
	}
}

####################################################
# Determines whether the service is running or not #
####################################################

function IsCoreServiceRunning()
{
	$coreProcess = Get-Process "Core.Service*"

	if($coreProcess -eq $null) {
		"Core process is not running." | slog -error -buffer
		return $false;
	}
	
	return $true
}

###################################
# Gets Core Client For Connection #
###################################

function GetCoreClient
{
param(
	[string]$coreUri,
	[string]$coreAccessUserName="",
	[string]$coreAccessPass="",
	[switch]$useWindowsAuth
)
	try {
    LoadAssembly "Core.Client.dll"
	LoadAssembly "WCFClientBase.dll"
	LoadAssembly "Common.Contracts.dll"
	
	$uri = New-Object System.Uri($coreUri)
	$coreClient = New-Object Replay.Core.Client.CoreClient($uri)

	if ($useWindowsAuth -eq $true) {		
		$coreClient.UseDefaultCredentials()
	}
	else {
        if (-not [System.String]::IsNullOrEmpty($coreAccessUserName) -and $coreAccessUserName.Contains("\\"))
        {
            $slashIndex = $coreAccessUserName.IndexOf("\\", [System.StringComparison]::OrdinalIgnoreCase)
            $credentials = New-Object System.Net.NetworkCredential($coreAccessUserName.Substring($slashIndex + 1), $coreAccessPass, $coreAccessUserName.Substring(0, $slashIndex));
        }
        else
        {
            $credentials = New-Object System.Net.NetworkCredential($coreAccessUserName, $coreAccessPass);
        }

		$coreClient.UseSpecificCredentials($credentials)
	}
	 
    if ($coreClient -ne $null) {
        "Core object was created." | slog -buffer
    }
    else {
    	"Cannot create core client object" | slog -error -buffer
	    throw "Cannot create core client object"
    }
   
   #Load certificate
   
   LoadAssembly "ServiceHost.Implementation.dll"
   LoadAssembly "ServiceHost.Contracts.dll"
   
   $cryptoMethods = New-Object Replay.ServiceHost.Implementation.Certificates.CryptoApiMethods
   $registryConfigurationService = New-Object Replay.ServiceHost.Implementation.Configuration.RegistryConfigurationService("Software\AppRecovery\Core")
   $certificateService = New-Object Replay.ServiceHost.Implementation.Certificates.CertificateService($cryptoMethods, $registryConfigurationService, "AppRecoveryCore", "Root");
   $coreClient.ClientCertificates.Add($certificateService.GetClientCertificate()) |  out-null
   
   return $coreClient
   }
   catch {
			$_ | slog -error -buffer
	}
}

#############################
# Force backup creation		#
#############################

function StartArchiveCreation
{
param(
	[Replay.Core.Client.CoreClient]$coreClient,
	[string[]]$agentsDisplayName=@(),
	[string]$basePath,
	[string]$username="",
	[string]$pass="",
	[switch]$all
)
	try {
	"Starting Archive Creation" | slog -debug -buffer
	LoadAssembly "Core.Contracts.dll"
	
	$agentsManagement = $coreClient.AgentsManagement
	$backupManagement = $coreClient.BackupManagement
	
	$agents = $agentsManagement.GetAgents()
	$agentIds = New-Object Replay.Core.Contracts.Agents.AgentIdsCollection
	
	if ($all)
	{
		" -all switch, adding all agent ids to archive job" | slog -buffer
		foreach ($agent in $agents) {
			$agentIds.Add($agent.Id)
		}
	} elseif ($agentsDisplayName.Length -ge 1) {
		"Adding selected agent ids to archive job" | slog -buffer
		foreach ($displayName in $agentsDisplayName)
		{
			$displayName = $displayName.ToLowerInvariant().Trim()
		}
		
		foreach($agent in $agents) {
			if ($agentsDisplayName -contains ($agent.DisplayName.ToLowerInvariant()) ){
				"Adding $($agent.DisplayName) / ID: $($agent.Id)" | slog -buffer
				$agentIds.Add($agent.Id)
			}
		}
	} 

	
	if ($agentIds -ne $null) {
		"Creating new job" | slog -buffer
		
		$request = New-Object Replay.Core.Contracts.Backup.BackupJobRequest
		$request.AgentIds = $agentIds
		$request.EndDate = [System.DateTime]::UtcNow
		$request.MaxSegmentSizeMB = -1
		
		$pathUri = New-Object System.Uri($basePath)
		if ($pathUri.IsUnc) {
			$request.Location = New-Object Replay.Core.Contracts.Backup.NetworkBackupLocation
			if ($username) {
				$request.Location.User = $username
				$request.Location.Password = $pass
			}
		}
		else {
			$request.Location = New-Object Replay.Core.Contracts.Backup.BackupLocation
		}
		
		$request.Location.Path = $basePath
		$jobId = $backupManagement.StartBackup($request)
		

		if ($jobId) {
			"StartBackup API successful" | slog -debug -buffer
			return GetArchiveJobInfo($jobId)
		} else {
			"StartBackup API failed." | slog -error -buffer
		}
		
	} else {
		"No agent IDs provided" | slog -error -buffer
	}

	return $null
	}
	catch {
			$_ | slog -error -buffer
	}
}

###########################################
# Write Transfer job information to file  #
###########################################

function GetArchiveJobInfo
{
param(
	$jobId,
	[switch]$report
)
	try {
	LoadAssembly "Core.Contracts.dll"
	
	$backgroundJobManagement = $coreClient.BackgroundJobManagement
	
	$activeJob = $backgroundJobManagement.GetJob($jobId)

	if ($activeJob) {
		$rate = FormatJobRate($activeJob.Rate)
		$done = FormatJobPercentDone $activeJob.Progress $activeJob.TotalWork
		$phase = $activeJob.Phase
		$id = $activeJob.Id
		
		$logStr = "'" + $activeJob.Summary + "'" + " (" + $jobid + ")" + " is " + $activeJob.Status.ToString().ToUpper() + " at '" + $rate + "'" + " | Phase: " + $phase.ToString() + "(" + $done + " done)"
		
		if ($report) {
			"[progress:$($Global:G_staggerOutputEveryMinutes)/min] $logStr" | slog -buffer
		}
	}
	
	return $activeJob
	}
	catch {
			$_ | slog -error -buffer
	}
}

###########################################
# Check whether job is finished or not    #
###########################################

function IsJobFinished([Replay.Core.Contracts.BackgroundJobs.BackgroundJobInfo]$job)
{
	try {
	LoadAssembly "Reporting.Contracts.dll"
	
	return (($job.Status -eq [Replay.Reporting.Contracts.JobStatus]::Canceled.ToString()) -or ($job.Status -eq [Replay.Reporting.Contracts.JobStatus]::Succeeded.ToString()) -or ($job.Status -eq [Replay.Reporting.Contracts.JobStatus]::Failed.ToString()))
	}
	catch {
			$_ | slog -error -buffer
	}
}

#################################################
# Check whether job is in WAITING state or not  #
#################################################

function IsJobWaiting($job)
{
	try {
	LoadAssembly "Reporting.Contracts.dll"
	return ($job.Status -eq [Replay.Reporting.Contracts.JobStatus]::Waiting.ToString())
	}
	catch {
			$_ | slog -error -buffer
	}
}