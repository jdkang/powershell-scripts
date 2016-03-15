function Get-LoggedOnUser {
<#
    .SYNOPSIS
    Get logged on user
    
    .DESCRIPTION
    Parses WMI calls to get information about the logged on user of a computer
#>
param(
    [ValidateNotNullOrEmpty()]
    [string]$computername='localhost'
)
    $regexa = '.+Domain="(.+)",Name="(.+)"$'
    $regexd = '.+LogonId="(\d+)"$'

    $logontype = @{
        "0"="Local System"
        "2"="Interactive" #(Local logon)
        "3"="Network" # (Remote logon)
        "4"="Batch" # (Scheduled task)
        "5"="Service" # (Service account logon)
        "7"="Unlock" #(Screen saver)
        "8"="NetworkCleartext" # (Cleartext network logon)
        "9"="NewCredentials" #(RunAs using alternate credentials)
        "10"="RemoteInteractive" #(RDP\TS\RemoteAssistance)
        "11"="CachedInteractive" #(Local w\cached credentials)
    }

    $logon_sessions = @(gwmi win32_logonsession -ComputerName $computername)
    $logon_users = @(gwmi win32_loggedonuser -ComputerName $computername)

    $session_user = @{}

    $logon_users | % {
        $_.antecedent -match $regexa > $nul
        $username = $matches[1] + "\" + $matches[2]
        $_.dependent -match $regexd > $nul
        $session = $matches[1]
        $session_user[$session] += $username
    }


    $logon_sessions | % {
        $starttime = [management.managementdatetimeconverter]::todatetime($_.starttime)

        $loggedonuser = New-Object -TypeName psobject
        $loggedonuser | Add-Member -MemberType NoteProperty -Name "Session" -Value $_.logonid
        $loggedonuser | Add-Member -MemberType NoteProperty -Name "User" -Value $session_user[$_.logonid]
        $loggedonuser | Add-Member -MemberType NoteProperty -Name "Type" -Value $logontype[$_.logontype.tostring()]
        $loggedonuser | Add-Member -MemberType NoteProperty -Name "Auth" -Value $_.authenticationpackage
        $loggedonuser | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $starttime

        $loggedonuser
    }
}