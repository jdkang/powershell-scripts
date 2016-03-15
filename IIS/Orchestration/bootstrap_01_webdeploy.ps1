# --------------------------------------------------
# configuration
# --------------------------------------------------
# Web Deploy MSI (v3.5)
$wdMsiPath = '\\server.contoso.local\' +
             'webdeploy\WebDeploy_amd64_en-US.msi'
$msiLogDir = "C:\msiLogs"
$msiForceReinstall = $false

# PFX
# Script looks for domain certs here (with private keys) for use as the wmsvc cert
# This script assumes there are wildcard domain certs in the directory provided.
# e.g. if the server is "bar.contoso.local" it would check \\PATH\SPECIFIED\contoso.local.pfx
$domainCertPfxDir = '\\server.contoso.local\orchestration\pki'
$domainCertPfxPassword = ''

# WDeploy Delegate Rules
$delegateWdeployUsers = @('CONTOSO\delegatedUser')
$delegateWDeployDoNotWipeExistingPermissions = $false

$delegateRules = @()
$delegateRules += @{
    providers = 'appPoolConfig, backupManager, contentPath, createApp, iisApp, setAcl'
    actions = '*'
    path = '{userScope}'
    pathType = 'PathPrefix'
    identityType = 'ProcessIdentity' # wmsvc identity
    permittedUsers = $delegateWdeployUsers
    enabled = $true
}

###################################################
# Funcs
###################################################
#--------------------------------------------------
# X509 / Windows Cert Store
#--------------------------------------------------
function Import-Certificate2 {
param(
    # X509 Certificate path
    [Parameter (Mandatory=$true,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
    [ValidateScript({Test-Path $_})]
    [Alias('pspath')]
    [string]$certificate,
    # Store Name - e.g. Root, My, etc
    [Parameter(Mandatory=$true)]
    [string]$storeName,
    # Store Location - e.g. LocalMachine, CurrentUser
    [Parameter(Mandatory=$true)]
    [string]$storeLocation,
    [Parameter(Mandatory=$false)]
    [switch]$pfx,
    [Parameter(Mandatory=$false)]
    [string]$pfxPassword=''
)
BEGIN {
    $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store -ArgumentList  $storeName,$storeLocation
    $certStore.Open('ReadWrite')
    if (!$?) { throw "Unable to access specified certificate store" }
    $pfxImportFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet -bor
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
}
PROCESS {
    foreach ($c in $certificate) {
        $certFullname = (gi $c).FullName
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        if (!$pfx) {
            $certificate.Import($certFullname)
        }
        else {
            $certificate.Import($certFullname, $pfxPassword, $pfxImportFlags)
        }
        if ($?) {
            $certStore.Add($certificate)
            $certificate
        }
    }
}
END {
    $certStore.Close()
}
}

function Set-PfxToIisBinding {
param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
    [ValidateScript({Test-Path $_})]
    [Alias('pspath')]    
    [string]$pfxPath,
    [Parameter(Mandatory=$false)]
    [string]$pfxPassword="",
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$sslBindingPath
)
    # Import into Windows Cert Store
    $importedCert = (gi $pfxPath | Import-Certificate2 -storeName 'my' -storeLocation 'LocalMachine' -pfx -pfxPassword $pfxPassword)
    if (!$?) { throw "Cannot import .pfx into cert store" }
    write-host "Imported Cert thumbprint: $($importedCert.thumbprint)"

    # Swap SSL Binding
    $importedCertPath = 'cert:\localmachine\my\' + $importedCert.thumbprint
    if (!(Test-Path $importedCertPath)) { throw "Cannot reach imported certificate path" }

    $newIisBinding = $null
    if (!(Test-Path $sslBindingPath) -or 
    ((gi $sslBindingPath).thumbprint -ne $importedCert.thumbprint)) {
        if (Test-Path $sslBindingPath) { Remove-Item -Path $sslBindingPath }
        $newIisBinding = Get-Item -Path $importedCertPath | New-Item -Path $sslBindingPath
        if (!$?) { throw "Could not swap SSL binding." }
    } else {
        $newIisBinding = (gi $sslBindingPath)
    }
    return $newIisBinding
}

function Get-CA ([string]$CAName) {
        $domain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
        $domain = "DC=" + $domain -replace '\.', ", DC="
        $CA = [ADSI]"LDAP://CN=Enrollment Services, CN=Public Key Services, CN=Services, CN=Configuration, $domain"
        $CAs = foreach ($child in $CA.psBase.Children) {
            new-object psobject -property @{
                CAName = $child.Name.ToString()
                Computer = $child.DNSHostName.ToString()
            }
        }
        if ($CAName) {
            $CAs = @( $CAs | ? { $_.CAName -eq $CAName } )
        }
        if ($CAs.Count -eq 0) {throw "Sorry, here is no CA that match your search"}
        $CAs
}

function Request-ADCert {
[CmdletBinding()]
param(
    [Parameter(ParameterSetName='forThisComputer')]
    [switch]$forThisComputer,
    [Parameter(ParameterSetName='pipeline',ValueFromPipeLine=$true)]
    [string[]]$FQDN,
    [switch]$exportable,
    [switch]$exportPfx,
    [Parameter(Mandatory=$false)]
    [string]$exportPfxPassword='',
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$exportPfxDir="$($ENV:LOCALAPPDATA)\pfx",
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$template='WebServer-SuiteBCrypto',
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$CA # you can grab this from certutil
)
BEGIN {
    $certAuthority = ''
    if (!$CA) {
        write-verbose "Looking up CA..."
        $caInfo = (Get-CA | Select -First 1)
        if (!$?) { throw "Unable to automatically get CA info" }
        $certAuthority = ($caInfo.Computer + '\' + $caInfo.CAName)
        write-verbose "Using CA '$certAuthority'"
    } else {
        $certAuthority = $CA
    }
    
    $testCA = & certutil.exe -ping -Config $certAuthority 2>&1
    if (!$?) {
        $testCA
        throw "Cannot reach CA"
    }
    write-verbose "CA Test OK: $testCA"

    if ($exportPfx -and !$exportable) {
        throw "-exportPfx requires -exportable"
    }
    
    if ($exportPfx -and (!(Test-Path $exportPfxDir))) {
        mkdir $exportPfxDir -force -ea 0 | out-null
    }
    $executedForThisComputer = $false
}
PROCESS {
    $f = $FQDN
    
    # Handle -forThisComputer (essentially ignoring pipeline)
    if ($forThisComputer) {
        if ($executedForThisComputer) {
            return # ONLY way to 'continue' a PROCESS block
        } else {
            $f = (Get-WmiObject win32_computersystem).DNSHostName +"." + (Get-WmiObject win32_computersystem).Domain
        }
    }
    
    # Format INF
    $friendlyName = ($f -replace '\*','STAR') + ' ' + (get-date -f 'yyyyMMdd_HHmmss')
$requestBody = ('[NewRequest]
Subject="cn=' + $f + '"
Exportable=' + ([string]$exportable).ToUpper() + '
FriendlyName="' + $friendlyName + '"
MachineKeySet=TRUE
SILENT=TRUE
[RequestAttributes]
CertificateTemplate="' + $template + '"')
    
    write-verbose "Friendly Name: $friendlyName"
    write-verbose "INF Request Body: $requestBody"
    
    # Temporary files
    $certRequestTmpDir = [system.io.path]::GetTempPath() + 'certRequest-' + [guid]::NewGuid()
    md $certRequestTmpDir -force | out-null
    if (!$?) { throw "Cannot create temporary file directory" }
    $infTemplate = "$certRequestTmpDir\request.inf"
    $reqRequest = "$certRequestTmpDir\request.req"
    $certResponse = "$certRequestTmpDir\request.cer"
    
    # Request Cert from ADCS
    try {
        # Write INF to file
        $requestBody | out-file -filepath $infTemplate
        if (!$?) { throw "Unable to write cert request INF file" }
    
        # Generate Request
        $o = & certreq.exe -q -new $infTemplate $reqRequest 2>&1
        if (!$?) {
            $o
            throw "Unable to process new cert request"
        }
        
        # Submit Request / Get Response
        $o = & certreq.exe -config $certAuthority -q -submit $reqRequest $certResponse 2>&1
        if (!$?) {
            $o
            throw "Unable to submit signing request"
        }
        
        # Import Response
        $o = & certreq.exe -q -accept $certResponse 2>&1
        if (!$?) {
            $o
            throw "Unable accept enrollment response"
        }
    } catch {
        throw
    } finally {
        rm -path $certRequestTmpDir -recurse -force -ea 0
    }
    
    # Return either a Windows Cert Store object -OR- a .pfx file object
    $cert = gci cert:\LocalMachine\my | ? { $_.FriendlyName -eq $friendlyName }
    if (!$cert) { throw "Cannot find imported cert in store" }
    
    if (!$exportPfx) {
        return $cert
    } else {
        $certPath = "$exportPfxDir\$($friendlyName)_$($cert.thumbprint).pfx"
        $certBytes = $cert.Export('PFX', $exportPfxPassword)
        [system.IO.file]::WriteAllBytes($certPath,$certBytes)
        return (gi $certPath)
    } 
}
}

#--------------------------------------------------
# Windows Services
#--------------------------------------------------
function Set-WindowsServiceFailuresToRetry {
param(
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [alias('name')]
    [string]$serviceName,
    [int]$resetSeconds=86400, #24h
    [int]$retrySeconds=30
)
    if (Get-Service $serviceName) {
        $retryMilliseconds = $retrySeconds * 1000
        & sc.exe failure $serviceName reset= $resetSeconds actions= "restart/$retryMilliseconds" > $null
    } else {
        write-warning "Cannot find service $serviceName"
    }
}

#--------------------------------------------------
# MSI
#--------------------------------------------------
function CopyAndRun-Msi {
param(
    [string]$msiFilePath="",
    [string]$msiLogDir="C:\msiLogs",
    [string]$msiProperties,
    [int[]]$acceptableExitCodes=@(0)
)
    if (!(Test-Path $msiFilePath)) { throw "Cannot access MSI file" }
    $msiSrcFile = gi $msiFilePath
    
    if (!(Test-Path $msiLogDir)) {
        mkdir $msiLogDir -force | out-null
        if (!$?) { throw "Cannot create msi log directory" }
    }
    # Copy locally
    $tmpCopyDir = [system.io.path]::GetTempPath() + $msiSrcFile.basename + '-' + [guid]::NewGuid()
    mkdir $tmpCopyDir -ea 0 | out-null
    $msiSrcFile | cp -destination $tmpCopyDir
    if (!$?) { throw "Cannot copy web deploy msi file" }

    # Execute
    $msiLogPath = $msiLogDir + "\$($msiSrcFile.BaseName)_" + (Get-Date -f 'yyyyMMdd_HHmmss') + '.log'
    $msiParams = "/i $tmpCopyDir\$($msiSrcFile.Name) /quiet /norestart /L*V $msiLogPath"
    if ($msiProperties) { $msiParams += " $msiProperties" }
    $msiexec = Start-Process -FilePath msiexec.exe -ArgumentList $msiParams -Wait -PassThru
    if ($acceptableExitCodes -notcontains $msiexec.exitcode) {
        write-warning "Check log: $msiLogPath"
        write-error "msiexec returned error code $($msiexec.exitcode)"
        throw "msiexec exited with non-zero code"
    }
}

#---------------------------------------------------
# Web Deploy
#---------------------------------------------------
function Enable-WebDeployBackupsServerScope {
    if (!(Test-WebDeployIsInstalled)) { throw "Web Deploy not installed" }

    # Yes this has to be dot-sourced
    # Yes it relies on setting the PWD
    $wdInstallPath = (Get-WebDeployInstallationInformation).InstallPath | Select -First 1
    push-location "$wdInstallPath\scripts"
    . '.\BackupScripts.ps1'   
    TurnOn-Backups -On $true
    Configure-Backups -Enabled $true
    pop-location
}

function Test-WebDeploy {
param(
    [string]$computer="localhost"
)
    if (!(Test-WebDeployIsInstalled)) { throw "Web Deploy not installed" }

    $wdInstallPath = (Get-WebDeployInstallationInformation).InstallPath | Select -First 1

    & ($wdInstallPath + 'msdeploy.exe') "–verb:dump" "–source:appHostConfig,computername=$computer" 2>&1 > $null
    $?
}

function Get-WebDeployInstallationInformation {
    foreach($number in 3..1)
    {
        $keyPath = "HKLM:\Software\Microsoft\IIS Extensions\MSDeploy\" + $number
        if(Test-Path($keypath))
        {
            Get-ItemProperty $keypath
        }
    }
    return $null
}

function Test-WebDeployIsInstalled {
    $wdInfo = Get-WebDeployInstallationInformation
    if (($wdInfo -ne $null) -and $wdInfo.Install -eq 1) { return $true }
    return $false
}

#--------------------------------------------------
# IIS Configuration
#--------------------------------------------------
[System.Reflection.Assembly]::LoadFrom( ${env:windir} + "\system32\inetsrv\Microsoft.Web.Administration.dll" ) > $null
if (!$?) { throw "Unable to laod Microsoft.Web.Administration.dll" }

function Check-IsUsingSharedConfiguraion
{
    $serverManager = (New-Object Microsoft.Web.Administration.ServerManager)
    $section = $serverManager.GetRedirectionConfiguration().GetSection("configurationRedirection")
    return [bool]$section["enabled"]
}

function Check-DelegationHandlerInstalled
{
    try {
        $serverManager = (New-Object Microsoft.Web.Administration.ServerManager)
        $serverManager.GetAdministrationConfiguration().GetSection("system.webServer/management/delegation").GetCollection() > $null
    } catch {
        return $false
    }
    return $true
}

# When WDeploy gets installed it runs:
# $($ENV:programfiles)\iis\Microsoft Web Deploy V3\scripts\AddDelegationRules.ps1
# Which generates local admins w/ appHostConfig r/w as proxy users and encrypts their passwords as a protected section.
# Admins, by default, have access to the RSA key container 'MyKeys' which allows decryption of these sections.
# http://www.iis.net/learn/publish/using-web-deploy/powershell-scripts-for-automating-web-deploy-setup
function Get-WDeployDelegateCredentials {
    $serverManager = (New-Object Microsoft.Web.Administration.ServerManager)
    $delegationRulesCollection = $serverManager.GetAdministrationConfiguration().GetSection("system.webServer/management/delegation").GetCollection()
    
    $creds = @()
    foreach ($item in $delegationRulesCollection) {
        $username = ""
        $password = ""
        $runAsElement = $item.ChildElements['runAs']
        if ($runAsElement) {
            $username = $runAsElement.Attributes['userName'].Value
            $password = $runAsElement.Attributes['password'].Value
        }
        if ($password) {
            if ($creds.username -notcontains $username) {
                $creds += new-object psobject -property @{
                    username = $username
                    password = $password
                }
            }
        }
    }
    $creds
}

function Clear-AllDelegateRulePermissions {
    $serverManager = (New-Object Microsoft.Web.Administration.ServerManager)
    $delegationRulesCollection = $serverManager.GetAdministrationConfiguration().GetSection("system.webServer/management/delegation").GetCollection()

    foreach ($item in $delegationRulesCollection) {
        $permissions = $item.GetCollection('permissions')
        if ($permissions) { $permissions.Clear() }
    }
    $serverManager.CommitChanges()
}

function Remove-DelegateRuleIfExists {
param(
    [Parameter(Mandatory=$true)]
    [string]$providers,
    [Parameter(Mandatory=$true)]
    [string]$path
)
    $serverManager = (New-Object Microsoft.Web.Administration.ServerManager)
    $delegationRulesCollection = $serverManager.GetAdministrationConfiguration().GetSection("system.webServer/management/delegation").GetCollection()

    for($i=0; $i -lt $delegationRulesCollection.Count; $i++)
    {
        $providersValue = $delegationRulesCollection[$i].Attributes["providers"].Value
        $pathValue = $delegationRulesCollection[$i].Attributes["path"].Value
        $enabledValue = $delegationRulesCollection[$i].Attributes["enabled"].Value
        
        if (($providersValue -eq $providers) -and ($pathValue -eq $path)) {
            $delegationRulesCollection[$i].Delete()
        }
    }
    $serverManager.CommitChanges()
}

function Create-DelegateRule {
param(
    [Parameter(Mandatory=$true)]
    [string]$providers,
    [Parameter(Mandatory=$true)]
    [string]$actions,
    [Parameter(Mandatory=$true)]
    [string]$path,
    [Parameter(Mandatory=$true)]
    [string]$pathType,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Specificuser','CurrentUser','ProcessIdentity')]
    [string]$identityType,
    [Parameter(Mandatory=$false)]
    [string]$runAsUsername,
    [Parameter(Mandatory=$false)]
    [string]$runAsPassword,
    [Parameter(Mandatory=$true)]
    [string[]]$permittedUsers,
    [switch]$enabled,
    [switch]$recreateIfExists
)
    if ($recreateIfExists) {
        Remove-DelegateRuleIfExists -providers $providers -path $path
    }
    $serverManager = (New-Object Microsoft.Web.Administration.ServerManager)
    $delegationRulesCollection = $serverManager.GetAdministrationConfiguration().GetSection("system.webServer/management/delegation").GetCollection()
    
    $newRule = $delegationRulesCollection.CreateElement("rule")
    # Rule Attributes
    $newRule.Attributes["providers"].Value = $providers
    $newRule.Attributes["actions"].Value = $action
    $newRule.Attributes["path"].Value = $path
    $newRule.Attributes["pathType"].Value = $pathType
    $newRule.Attributes["enabled"].Value = [string]$enabled
    # RunAs
    $runAs = $newRule.GetChildElement("runAs")
    $runAs.Attributes["identityType"].Value = $identityType
    if ($identityType -eq 'SpecificUser') {
        $runAs.Attributes["userName"].Value = $runAsUsername
        $runAs.Attributes["password"].Value = $runAsPassword
    }
    # Permissions
    $permissions = $newRule.GetCollection("permissions")
    foreach ($user in $permittedUsers) {
        $p = $permissions.CreateElement("user")
        $p.Attributes["name"].Value = $user
        $p.Attributes["accessType"].Value = "Allow"
        $p.Attributes["isRole"].Value = "False"
        $permissions.Add($p) | out-null
    }
    $delegationRulesCollection.Add($newRule) | out-null
    $serverManager.CommitChanges()
}

###################################################
# init
###################################################
write-host "Initializing"
ipmo WebAdministration
if (!$?) { throw "Could not load IIS WebAdministration" }

write-host "Checking redirect.config (config redirection not supported)"
if (Check-IsUsingSharedConfiguraion) { throw "Script not setup for redirect.config farm setups" }

###################################################
# Main
###################################################
#--------------------------------------------------
# Remote Management Service
#--------------------------------------------------
write-host "Stopping Remote Management Service"
$wmsvcService = Get-Service wmsvc
$wmsvcService | Stop-Service
$wmsvcService.WaitForStatus('Stopped','00:01:00')

# Windows Service
write-host "Configuring Remote Management Service Windows Service"
$wmsvcService | Set-Service -StartupType Automatic
$wmsvcService | Set-WindowsServiceFailuresToRetry

# Find .pfx and Import into Windows Cert Store
write-host "Configuring Remote Management Service Certificate"
$sslBindingPath = 'IIS:\SslBindings\0.0.0.0!8172'

$computerDomain = (Get-WmiObject Win32_ComputerSystem).domain
$domainCertPath = "$domainCertPfxDir\$computerDomain" + '.pfx'
write-host "Searching for .pfx $domainCertPath"
if (!(Test-Path $domainCertPath)) { throw "Cannot access domain .pfx" }

# Ensure IIS Binding
# ALTERNATIVELY, you can use Request-ADCert if you have ADCS configured
write-host "Ensuring cert $domainCertPath set to binding $sslBindingPath"
$iisBinding = Set-PfxToIisBinding -pfxPath $domainCertPath -pfxPassword $domainCertPfxPassword -sslBindingPath $sslBindingPath
write-host "IIS Binding Thumbprint: $($iisBinding.thumbprint)"

# Enable remote management
write-host "Enabling Remote Management"
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WebManagement\Server" -Name "EnableRemoteManagement" -Value "1"

# Start Service
write-host "Starting Remote Management Service"
$wmsvcService | Start-Service
$wmsvcService.WaitForStatus('Running','00:01:00')

#--------------------------------------------------
# Web Deploy MSI
#--------------------------------------------------
# Must be installed after wmsvc is enabled in order to 
# get 'Management Service Integration' features
# Obtain apropos INSTALLLEVEL from Orca
#--------------------------------------------------
write-host "Ensuring Web Deploy and Delegation handler are installed"
$validInstallState = $false

if (!(Test-WebDeployIsInstalled)) {
    write-host "Web Deploy not installed"
    $validInstallState = $false
} elseif (!(Check-DelegationHandlerInstalled)) {
    write-host "Delegate Handler not installed"
    $validInstallState = $false
} else {
    write-host "Installs OK"
    $validInstallState = $true
}

if (!$validInstallState -or $msiForceReinstall) {
    if ($msiForceReinstall) { write-host "Forcing reinstall" }
    write-host "Installing WebDeploy from source msi $wdMsiPath"
    CopyAndRun-Msi -msiFilePath $wdMsiPath -msiProperties 'INSTALLLEVEL=11' -msiLogDir $msiLogDir
}

write-host "Enabling Web Deploy/IIS Backups on Server scope"
Enable-WebDeployBackupsServerScope

#--------------------------------------------------
# Web Deploy Agent Service
#--------------------------------------------------
write-host "Stopping Web Deploy Agent Service"
$msdepsvcService = Get-Service msdepsvc
$msdepsvcService | Stop-Service
$msdepsvcService.WaitForStatus('Stopped','00:01:00')
write-host "Configuring Web Deploy Agent Windows Service"
$msdepsvcService | Set-Service -StartupType Automatic
$msdepsvcService | Set-WindowsServiceFailuresToRetry
write-host "Stopping Web Deploy Agent Service"
$msdepsvcService | Start-Service
$msdepsvcService.WaitForStatus('Running','00:01:00')

#--------------------------------------------------
# Smoke test
#--------------------------------------------------
write-host "Smoke Tests"
# WebDeploy
write-host "Testing WebDeploy on localhost"
if (!(Test-WebDeploy)) { throw "Could not reach web deploy on localhost" }

# Delegation handler
write-host "Verifying delegation handler"
if (!(Check-DelegationHandlerInstalled)) { throw "Delegation handler not installed" }

#--------------------------------------------------
# Delegate Rules
#--------------------------------------------------
# Clear rule permissions
if (!$delegateWDeployDoNotWipeExistingPermissions) {
    write-host "Clearing all delegation rule permissions"
    Clear-AllDelegateRulePermissions
}

# Create Rules
foreach ($rule in $delegateRules) {
    write-host "Creating rule:"
    write-host ($rule.GetEnumerator() | ? { $_.Name -ne 'runAsPassword' } | ft -auto | out-string)
    Create-DelegateRule -recreateIfExists @rule
}