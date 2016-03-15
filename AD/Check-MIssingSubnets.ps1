# Watches the netlogon log for hosts in subnets that aren't present in AD Sites & Services so we can add them.
# This script assumes to be running on a domain controller. The easiest path to allow it to run as a Scheduled
#    Task is to add the account it will run as to Backup Users or you'll likely get an error about permission
#    to schedule batch jobs.
# If you're not going to schedule this to run daily, update the call to AddDays appropriately.

$sys = (Get-ChildItem Env: | Where { $_.Name -eq "SystemRoot" }).Value
$logName = $sys+"\debug\netlogon.log"
$after = (Get-Date).AddDays(-1)
$hostname = (Get-ChildItem Env: | Where { $_.Name -eq "COMPUTERNAME"}).Value

$hosts = @("")
$log = (Get-Content $logName) -match $after.Month.ToString("00")+"/"+$after.Day.ToString("00") -match "NO_CLIENT_SITE"
foreach($line in $log){
	$parts = $line.Split(' ')
	$hosts += ($parts[4]+" "+$parts[5])
}

if($hosts.Count -lt 2) { return } # First item is the blank entry when we initialized the list.
$body = "The following hosts contacted the domain from a subnet that is not registered in AD Sites & Services:`n"
$body += ($hosts | Select -Unique | Sort-Object | FT -AutoSize | Out-String)
$body += "`nThis report was generated on host `"$hostname`"."

$smtpServer = "relay.domain.tld"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "donotreply@domain.tld"
$msg.ReplyTo = "donotreply@domain.tld"
$msg.To.Add("systems@domain.tld")
$msg.subject = "AD Sites & Services - Unregistered Subnets"
$msg.body = $body
$smtp.Send($msg)