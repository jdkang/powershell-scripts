<#
    .SYNOPSIS
    Installs the F5 X-Fowarded-For Native IIS HTTP Module
    
    .DESCRIPTION
    The F5 X-Forwarded-For Native IIS HTTP Module rewrites the X-FORWARDED-FOR header to appear as the origin IP in "one arm SNAT" type load balancing scenarios, i.e. scenarios where the traffic appears to be coming from the LB rather than the clients.
    
    .LINK
    https://devcentral.f5.com/articles/x-forwarded-for-http-module-for-iis7-source-included
#>
param(
    [string]$installPath = 'C:\inetpub\lib'
)

# Set PWD to .ps1 location
$currentScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
Push-Location $currentScriptPath

write-host 'Loading webadministration' -f green
ipmo Webadministration -ea 0
if (!$?) { throw "Cannot load webadministration module" }

write-host "Creating directory $installPath" -f green
md $installPath -force -ea 0

write-host "Copying files from $($pwd.path) to $installPath" -f green
robocopy . $installPath /e /w:10 /r:10
write-host 'Registering F5XFFHttpModule (x64)' -f green
New-WebGlobalModule -Name 'F5XFFHttpModule' -image "$installPath\F5XFFHttpModule.dll" -precondition 'bitness64'
write-host 'Registering F5XFFHttpModule32 (x86)' -f green
New-WebGlobalModule -Name 'F5XFFHttpModule32' -image "$installPath\F5XFFHttpModule32.dll" -precondition 'bitness32'

write-host 'RESTARTING IIS' -f green
iisreset

