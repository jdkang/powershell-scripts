function Invoke-RemoteExpressionWithPsExec
{
<#
    .SYNOPSIS
    Execute Powershell on a remote computer using PsExec.
    
    .DESCRIPTION
    Wraps powershell commands using Base64 around PsExec (which itself uses the ADMIN$ share).
    
    This can be useful in situations where psremoting is not available OR you need the process token of the executing users -- e.g. windows updates, MSI installations, etc. PsRemoting doesn't grant such tokens, which is often okay.
    
    You won't get the serilization of psremoting and you'll pay the fee for scaffolding up and down a temporary service.
#>
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