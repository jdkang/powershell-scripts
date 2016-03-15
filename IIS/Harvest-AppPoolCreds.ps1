<#
    This script will return all the app pool username/passwords from IIS.
    This is an ideal script for PsRemoting
#>
if ((Get-Module -ListAvailable | Select -expand Name) -contains 'webadministration') {
    ipmo webadministration -ea 0
    foreach ($appPool in (gci iis:\appPools)) {
        if ($appPool.processmodel.identityType -eq 'specificUser') {
            new-object psobject -property @{
                appPoolName = $appPool.name
                username = $appPool.processmodel.username
                password = $appPool.processmodel.password
            }
        }
    }
}