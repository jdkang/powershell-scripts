[CmdletBinding()]
param(
<# IIS: Site #>

    [string]$iisSiteListenIp = '10.128.0.1',
    [string]$iisSiteName = "foo.contoso.local",
    [string]$iisSitePhysicalBasePath = 'E:\IIS-SITES',
    [switch]$replaceExistingPhysicalPath,
    [switch]$replaceExistingIisSite,
    
<# IIS: Appplication Pool #>

    [string]$appPoolUserDomain = 'contoso.local',
    [string]$appPoolUsername="appool_foo",
    [string]$appPoolPassword='PASSWORD',
    [string]$appPoolDotNetVersion = "v4.0",
    [switch]$replaceExistingAppPool,
    
<# NTFS Permissions #>

    [string[]]$ntfsAppPoolDesiredPermissions = @("ReadAndExecute","Synchronize"),
    
<# Backups #>
    [switch]$noRollback
)

######################################################
# Init
#####################################################
# Way to propagate -debug and -verbose via SPLAT (if using CmdletBVinding())
# e.g. Get-Item @verbug
$verbug = @{
	verbose = ($PSCmdlet -and [bool]$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) -or ($VerbosePreference -eq 'continue')
	debug = ($PSCmdlet -and [bool]$PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) -or ($DebugPreference -eq 'continue')
}

# Check Local Admin
$isAdmin = (New-Object Security.Principal.WindowsPrincipal -ArgumentList ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (!$isAdmin) { throw "Local Admin required." }

# script dir
$currentScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
Push-Location $currentScriptPath

# Modules
ipmo .\modules\PowerShellAccessControl\PowerShellAccessControl.psm1 # I promise you don't want to deal with native NTFS ACE api
ipmo webadministration -ea 0
start-sleep -seconds 2 # http://stackoverflow.com/questions/14862854/powershell-command-get-childitem-iis-sites-causes-an-error

function Start-Cleanup {
    remove-module -Name PowerShellAccessControl -force
}

######################################################
# Main
######################################################
# Derived IIS Info
$appPoolName = "$($appPoolUsername) - $($appPoolUserDomain)"
$appPool = "IIS:\AppPools\$appPoolName"
$iisSitePhysicalPath = "$iisSitePhysicalBasePath\$iisSiteName"
$site = "IIS:\Sites\$iisSiteName"
# Domain UPN
$appPoolUpn = ""
$appPoolUpn += $appPoolUsername
if ($appPoolUserDomain) { $appPoolUpn += "@$appPoolUserDomain" }

# IIS Backup
$backupName = "$(Get-Date -f 'yyyyMMdd-HHmmss')_$iisSiteName"
write-verbose "Creating IIS backup/rollback $backupName $($env:WINDIR)\System32\inetsrv\$backupName"
$backup =  Backup-WebConfiguration $backupName
if (!$?) { throw "Unable to create IIS backup" }

try {   
    #-----------------------------------
    # Application Pool
    #-----------------------------------
    $appPoolExists = Test-Path $appPool
    if (!$appPoolExists -or $replaceExistingAppPool) {
        write-verbose "APPPOOL: Creating/Replacing AppPool $appPool"
        if ($appPoolExists) { Remove-WebAppPool -Name (gi $appPool).Name -ea 0 }
        $newAppPool = New-WebAppPool -Name $appPoolName
        $newAppPool.managedRuntimeVersion  = $appPoolDotNetVersion
        $newAppPool.processModel.identityType  = 2 #network service
        if ($appPoolUpn) {
            $newAppPool.processmodel.identityType = 3
            $newAppPool.processmodel.username = $appPoolUpn
            $newAppPool.processmodel.password = $appPoolPassword
        }
        $newAppPool | Set-Item
        if ((Get-WebAppPoolState -Name $appPoolName).Value -ne "Started") {
            throw "APPPOOL: New appPool was not started"
        }
        $newAppPool
    } else {
        write-verbose "APPPOOL: AppPool $appPool already exists. Skipping creation."
    }
    
    #-----------------------------------
    # Add to Local Group IIS_IUSRS
    #-----------------------------------
    $iisIusrs = [ADSI]"WinNT://localhost/IIS_IUSRS"
    $members = @($iisIusrs.psbase.Invoke("Members"))
    $memberNames = @()
    $members | ForEach-Object {
        $memberNames += $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null);
    }
    if (!$memberNames.Contains($appPoolUsername)) {
        write-verbose "IIS_IUSRS: Adding appPool user to IIS_IUSRS"
        $iisIusrs.Add("WinNT://$appPoolUserDomain/$appPoolUsername, user")
    } else {
        write-verbose "IIS_IUSRS: User already in IIS_IUSRS"
    }

    #-----------------------------------
    # Physical Path
    #-----------------------------------
    $physicalPathExists = Test-Path $iisSitePhysicalPath
    if (!$physicalPathExists -or $replaceExistingPhysicalPath) {
        write-verbose "PhysicalPath: Crearing/replacing Path $iisSitePhysicalPath"
        if ($physicalPathExists) { rm -path $iisSitePhysicalPath -recurse -force }
        md $iisSitePhysicalPath -force
    } else {
        write-verbose "PhysicalPath: Path $iisSitePhysicalPath already exists. Skipping creation."
    }
        
    # NTFS: Evaluate Rights
    $hasSufficentPermisisons = $true
    $currentPermissions = (Get-EffectiveAccess -Path $iisSitePhysicalPath -Principal $appPoolUpn)
    if ($currentPermissions) {
        $currentPermissionsList = $currentPermissions.EffectiveAccess.split(',').trim()
        foreach ($p in $ntfsAppPoolDesiredPermissions) {
            if (!$currentPermissionsList.Contains($p)) { 
                write-verbose "NTFS: $appPoolUpn missing permission $p on $iisSitePhysicalPath"
                $hasSufficentPermisisons = $false
            }
        }
    } else {
        write-verbose "NTFS: $appPoolUpn has no permissions to $iisSitePhysicalPath"
        $hasSufficentPermisisons = $false
    }
    # NTFS: Grant Rights
    if (!$hasSufficentPermisisons) {
        write-verbose "NTFS: Writing new NTFS ACE for $appPoolUpn for $iisSitePhysicalPath"
        gi $iisSitePhysicalPath |
            Add-AccessControlEntry (New-AccessControlEntry -Principal $appPoolUpn -FolderRights $ntfsAppPoolDesiredPermissions) -Apply -Force
    } else {
        write-verbose "NTFS: Permissions OK"
    }
    gi $iisSitePhysicalPath | Get-AccessControlEntry
    
    #-----------------------------------
    # Site
    #-----------------------------------
    $iisSiteExists = Test-Path $site
    if (!$iisSiteExists -or $replaceExistingIisSite) {
        write-verbose "SITE: Creating/Replacing Site $site"
        if (Test-Path $site) { Remove-Website -Name (gi $site).Name }
        New-Item $site -physicalPath $iisSitePhysicalPath -bindings @{
            protocol = 'http'
            bindingInformation = ($iisSiteListenIp+":80:"+$iisSiteName)
        } -ApplicationPool $appPoolName
    } else {
        write-verbose "SITE: Site $site already exists. Skipping creation."
    }
} catch {
    #-----------------------------------
    # Rollback
    #-----------------------------------
    Start-Cleanup
    write-error "ERROR during process, rolling back"
    start-sleep -seconds 5 # wait for file unlocks
    if (!$noRollback) {
        write-verbose "Rolling back IIS Settings"
        Restore-WebConfiguration -Name $backupName
    } else {
        write-verbose "Rollback disabled"
    }
    throw
}
Start-Cleanup