<#
    .SYNOPSIS
    Basic IIS initlization
    
    .DESCRIPTION
    Sets up server roles and moves default IIS root.
#>

# --------------------------------------------------------------------
# Define the variables.
# --------------------------------------------------------------------
$InetPubRoot = "E:\Inetpub"
$InetPubLog = "E:\Inetpub\Log"
$InetPubWWWRoot = "E:\Inetpub\WWWRoot"

# --------------------------------------------------------------------
# Windows Features
# --------------------------------------------------------------------
Import-Module ServerManager -ea 0
$winFeaturesDesired = @(
    # ---- IIS ----
        # Web Server (IIS)
        # 'web-server',
            # Common HTTP Features
            # 'Web-Common-Http',
                'Web-Static-Content',
                'Web-Default-Doc',
                'Web-Dir-Browsing',
                'Web-Http-Errors',
                'Web-Http-Redirect',
            # Application Development
            # 'Web-App-Dev',
                'Web-Asp-Net',
                'Web-Asp-Net45',
                'Web-Net-Ext',
                'Web-Net-Ext45',
                'Web-ISAPI-Ext',
                'Web-ISAPI-Filter',
            # Health and Diagnostics
            # 'Web-Health',
                'Web-Http-Logging',
                'Web-Log-Libraries',
                'Web-Request-Monitor',
            # Performance
            'Web-Performance',
            # Security
            # 'Web-Security',
                'Web-Basic-Auth',
                'Web-Filtering',
                'Web-Windows-Auth',
            # Management Tools
            # 'Web-Mgmt-Tools',
                'Web-Mgmt-Console',
                'Web-Mgmt-Service',
    # ---- NON-IIS ----
        # Windows Process Activation Service
        # I don't think we need WAS for IIS8+
        # 'WAS',
        # .NET Framework 4.5
        # 'NET-Framework-45-Features',
            'NET-Framework-45-ASPNET',
        # Remote Server Administration Tools
        # 'RSAT',
            'RSAT-Web-Server'
)
$winFeatures = Get-WindowsFeature $winFeaturesDesired

$winFeaturesNeedInstalled = $winFeatures |
                            ? { $_.Installed -eq $false }
$winFeaturesNotAvailable = $winFeaturesNeedInstalled |
                           ? { $_.InstallState -ne 'Available' }
if ($winFeaturesNotAvailable) {
    foreach ($feature in $winFeaturesNotAvailable) {
        write-warning "Feature $($feature.name) ($($feature.displayname)) is not available for installation"
    }
    throw "1 or more windows features that needs to be installed is not available"
}

Add-WindowsFeature -Name $winFeaturesNeedInstalled -IncludeAllSubFeature

# --------------------------------------------------------------------
# Loading IIS Modules
# --------------------------------------------------------------------
Import-Module WebAdministration

# --------------------------------------------------------------------
# Creating IIS Folder Structure
# --------------------------------------------------------------------
New-Item -Path $InetPubRoot -type directory -Force -ErrorAction SilentlyContinue
New-Item -Path $InetPubLog -type directory -Force -ErrorAction SilentlyContinue
New-Item -Path $InetPubWWWRoot -type directory -Force -ErrorAction SilentlyContinue

# --------------------------------------------------------------------
# Copying old WWW Root data to new folder
# --------------------------------------------------------------------
$InetPubOldLocation = @(get-website)[0].physicalPath.ToString()
$InetPubOldLocation =  $InetPubOldLocation.Replace("%SystemDrive%",$env:SystemDrive)
Copy-Item -Path $InetPubOldLocation -Destination $InetPubRoot -Force -Recurse

# --------------------------------------------------------------------
# Setting directory access
# --------------------------------------------------------------------
$Command = "icacls $InetPubWWWRoot /grant BUILTIN\IIS_IUSRS:(OI)(CI)(RX) BUILTIN\Users:(OI)(CI)(RX)"
cmd.exe /c $Command
$Command = "icacls $InetPubLog /grant ""NT SERVICE\TrustedInstaller"":(OI)(CI)(F)"
cmd.exe /c $Command

# --------------------------------------------------------------------
# Setting IIS Variables
# --------------------------------------------------------------------
#Changing Log Location
$Command = "%windir%\system32\inetsrv\appcmd set config -section:system.applicationHost/sites -siteDefaults.logfile.directory:$InetPubLog"
cmd.exe /c $Command
$Command = "%windir%\system32\inetsrv\appcmd set config -section:system.applicationHost/log -centralBinaryLogFile.directory:$InetPubLog"
cmd.exe /c $Command
$Command = "%windir%\system32\inetsrv\appcmd set config -section:system.applicationHost/log -centralW3CLogFile.directory:$InetPubLog"
cmd.exe /c $Command

#Changing the Default Website location
Set-ItemProperty 'IIS:\Sites\Default Web Site' -name physicalPath -value $InetPubWWWRoot

# --------------------------------------------------------------------
# Checking to prevent common errors
# --------------------------------------------------------------------
If (!(Test-Path "C:\inetpub\temp\apppools")) {
  New-Item -Path "C:\inetpub\temp\apppools" -type directory -Force -ErrorAction SilentlyContinue
}

# --------------------------------------------------------------------
# Deleting Old WWWRoot
# --------------------------------------------------------------------
Remove-Item $InetPubOldLocation -Recurse -Force

# --------------------------------------------------------------------
# Resetting IIS
# --------------------------------------------------------------------
& iisreset