##########################################################################
# Script Boilerplate
##########################################################################
$differentLogLocation = ''
# ---- Check Local Admin ----
$isAdmin = (New-Object Security.Principal.WindowsPrincipal -ArgumentList ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
# ---- noninteractive mode check ----
$nonInteractive = [bool]([Environment]::GetCommandLineArgs() -match '-noni')
# ---- PWD ----
$currentScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
Set-Location $currentScriptPath
$scriptName = Split-Path (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path -Leaf
# ---- Transcript ----
$DebugPreferenceOriginal = $DebugPreference
$VerbosePreferenceOriginal = $VerbosePreference
$WarningPreferenceOriginal = $WarningPreference
$DebugPreference = 2
$VerbosePreference = 2
$WarningPreference = 2
$logLocation = ''
if ($differentLogLocation) {
    $logLocation = $differentLogLocation
} else {
    $logLocation = "$currentScriptPath\logs"
}
if (!(Test-Path $logLocation)) { md $logLocation -force | out-null }
Start-Transcript -Path "$($logLocation)\$($scriptName)-$(get-date -f 'yyyyMMdd-HHmmss').log" -Force
function Stop-VerboseTranscript {
    write-verbose "Cleaning up"
    $DebugPreference = $DebugPreferenceOriginal
    $VerbosePreference = $VerbosePreferenceOriginal
    $WarningPreference = $WarningPreferenceOriginal
    try{
      stop-transcript | out-null
    }
    catch [System.InvalidOperationException]{}
}
# ---- Error unfolding ----
function Pretty-Exception {
param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [System.Management.Automation.ErrorRecord[]]
    $exceptions
)
PROCESS {
    $exceptions | % {
        write-host "---------------------------------------" -f 'red'
        write-host "ERROR RECORD" -f 'red'
        write-host "---------------------------------------" -f 'red'
        write-host "MESSAGE: " -n -f 'red'; write-host $_.Exception.Message -f yellow
        write-host "CATEGORY: " -n -f 'red'; write-host $_.CategoryInfo.Category -f yellow
        write-host "TYPE: " -n -f 'red'; write-host $_.Exception.GetType().FullName -f yellow
        write-host "ID: "  -n -f 'red'; write-host $_.FullyQualifiedErrorID -f yellow
        write-host "LINE: " -n -f 'red'; write-host (($_.InvocationInfo.Line).trim()) -f yellow
        write-host "STACK TRACE:" -f 'red'; write-host $_.ScriptStackTrace -f yellow
        write-host "---- EXCEPTION DETAILS ----" -f 'red'
        write-host ($_.Exception | fl -force | out-string).trim() -f yellow
    }
}}
trap {
    if (!$nonInteractive) {
        $_ | Pretty-Exception
    }
    if ($ENV:TEAMCITY_DATA_PATH) {
        write-host "##teamcity[message text='$($_.message)' status='FAILURE']"
        [Environment]::Exit(1)
    }
    break
}


function Ensure-Module {
param(
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$name,
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$path
)
    write-verbose "Ensuring module $module"
    if ((Get-Module).name -contains $name) { return }
    if (!$path -and (Get-Module -ListAvailable).Name -notcontains $name) {
        throw "Cannot find module in ENV:PsModulePath"
    }
    if ($path -and !(Test-Path $path)) { throw "Cannot find module path" }
    if ($path) {
        ipmo $path
    } else {
        ipmo $name
    }
    if (!$?) { throw "Unable to load module $module" }
}