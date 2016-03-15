function Find-InstalledSoftware {
<#
    .SYNOPSIS
    Find installed software.
    
    .DESCRIPTION
    Search the registry for installed software and optionally installed software with uninstall strings.
#>
param(
    [string]$displayNamePart,
    [switch]$withUninstallString
)

    Set-Variable -Name hklmx64 -Value "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -Option Constant
    Set-Variable -Name hklmx86 -Value "HKLM:\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall" -Option Constant
    Set-Variable -Name hkcux64 -Value "HKCU:\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall" -Option Constant
    Set-Variable -Name hkcux86 -Value "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -Option Constant
    $regs = $hklmx64,$hklmx86,$hkcux64,$hkcux86

    foreach ($reg in $regs) {
        $SubKeys = $null
        if (Test-path $reg) { $SubKeys = Get-ItemProperty "$reg\*" }
        foreach ($subkey in $subkeys) {
            if ( $subkey.DisplayName -like "*$displayNamePart*" ) {
                if (!$withUninstallString -or ($subkey.uninstallstring)) {
                    $subkey
                }
            }
        }
    }
}