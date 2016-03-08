function Invoke-ScriptWithCredentials {
<#
    .SYNOPSIS
    A wrapper for Invoke-Command to run against a list of FQDN with varying credentials per domain.
    
    .DESCRIPTION
    This function wraps Invoke-Command and stores a mapping of credentials mapped to each unique domain name (e.g. foo.contoso.local, contoso.local, etc). The original intent was for scenarios where one has a "keyring" of credentials across AD forests with no trusts.
    
    By deafult, it will PROMPT you for credentials per unique domain. You can also preemptively pass a hashtable with the -creds arg.
    
    By default, it ASSUMES the current user credential has sufficent access to the current machine's joined domain. This is overridable by using the -PromptForCurrentDomain arg.
    
    It uses the DEFAULT authenitication provider, which encounters the double-hop issue in regards to network resources. This is overridable with the -extraArgs hashtable arg.
    
    .INPUTS
    None.
    
    .OUTPUTS
    System.Management.Automation.PSCustomObject[]
    
    .EXAMPLE
    $computers =
    @{
        'srv1.contoso.local',
        'srv2.contoso.local',
        'srv3.adventureworks.local',
        'srv4.wut.local'
    }
    Invoke-ScriptWithCredentials -computers $computers -filepath '.\foo.ps1'
    
    You would be prompted for the credentials for each unique domain.
    You WON'T be prompted for the domain the user running the script is on.
    
    .EXAMPLE
    $computers =
    @{
        'srv1.contoso.local',
        'srv2.contoso.local',
        'srv3.adventureworks.local',
        'srv4.wut.local'
    }
    $credHt = @{
        'adventureworks.local' = get-credential
        'wut.local' = get-credential
    }
    Invoke-ScriptWithCredentials -computers $computers -filepath '.\foo.ps1'
    
    Preemptively specify a hashtable with the credential mappings.
    
    .EXAMPLE
    $computers =
    @{
        'srv1.contoso.local',
        'srv2.contoso.local',
        'srv3.adventureworks.local',
        'srv4.wut.local'
    }
    $extraArgs = @{ authentication = 'credssp' }
    Invoke-ScriptWithCredentials -computers $computers -filepath '.\foo.ps1' -extraArgs $extraArgs
    
    Pass additional arguments to the Invoke-Command.
#>
param(
    # List of FQDN computer naems to execute script against
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $computers,
    # Path to script
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({Test-Path $_})]
    [string]
    $FilePath,
    # Hashtable that maps PSCredentials to domain names
    [Parameter(Mandatory=$false)]
    [hashtable]
    $Creds = @{},
    # Extra arguments to pass to Invoke-Command, e.g. @{ authentication = 'credssp' }
    [Parameter(Mandatory=$false)]
    [hashtable]
    $extraArgs = @{},
    # Don't assume the current user credential is valid for this machine's domain, i.e. prompt for the currently joined domain as well
    [switch]$PromptForCurrentDomain
)
    # Prompt/Populate Credential HT for each domain
    # Skip current computer domain (i.e. use running credentials)
    $credsHt = @{} + $Creds
    $computerSystem = gwmi win32_computersystem
    $currentDomain = $computerSystem.domain
    foreach ($computer in $computers) {
        $split = $computer.split('.')
        $domain = $split[1..($split.length-1)] -join '.' 
        if ( (($domain -ne $currentDomain) -or $PromptForCurrentDomain) -and !$credsHt[$domain] ) {
            write-host "Crednetial for $($domain): " -f yellow
            $credsHt[$domain] = Get-Credential
        }
    }
    # Iterate script over computers
    foreach ($computer in $computers) {
        $split = $computer.split('.')
        $domain = $split[1..($split.length-1)] -join '.'
        $credSplat = @{}
        $credential = $credsHt[$domain]
        if ($credential) {
            $credSplat.Add('credential',$credential)
        }
        Invoke-Command -ComputerName $computer -FilePath $filePath @credSplat @extraArgs
    }
}