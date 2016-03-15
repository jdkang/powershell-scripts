<#
    .SYNOPSIS
    Search a machine's IIS bindings and drives for .pfx
    
    .DESCRIPTION
    This script will (1) gather data about IIS bindings and try to export the certs as .pfx and (2) try to scour a drive for .pfx files and (using a series of provided guesses) try to gather information about those certs as well as create a no-pasword copy of them.
    
    Gathered info is serialized into Powershell XML and exported .pfx are saved to disk.
    
    The intention is to harvest .pfx files with expirations of private keys to be stored in a safe location (e.g. KeePass, Secret Server, etc)
#>

ipmo webadmin*

$computerSystem = gwmi 'WIN32_ComputerSystem'
$fqdn = $computerSystem.name + '.' + $computerSystem.domain
$outDir = 'C:\certHarvester'
$outFilePath = "$outDir\$($fqdn)_$(get-date -f 'yyyyMMdd_HHmmss').xml"
$pfxOutDir = "$outDir\$($fqdn)_pfx"
mkdir $pfxOutDir -force -ea 0 | out-null

$pfxSearchDrives = @('D','E')
$pfxGuesses = @('','guess1','guess2','guess3')
$x509KeySet = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable

# IIS Binding Certs
$harvestedCerts = @()
$iisSslBindings = @()
write-host "Gathering SSL Bindings" -f green
try {
    $iisSslBindings = gci IIS:\SslBindings
}
catch {
    Get-ChildItem HKLM:\SYSTEM\CurrentControlSet\services\HTTP\Parameters\SslBindingInfo |
    ? { !($_ | Get-ItemProperty -Name 'SslCertStoreName' -ea 0) } |
    % {
        write-host "Fixing SSL Binding Info. Adding 'SslCertStoreName'='MY' to $($_.Name)" -f cyan
        $_ | New-ItemProperty -Name 'SslCertStoreName' -Value "MY"
    }
    $iisSslBindings = gci IIS:\SslBindings
}
if (!$iisSslBindings) { throw "Unable to determine IIS SSL bindings" }
foreach ($binding in $iisSslBindings) {
    write-host "Processing $($binding.pspath)" -f green
    $thumbprint = $binding | select -expand thumbprint
    $certPath = "cert:\localmachine\my\$thumbprint"
    if (Test-path $certPath) {
        $notes = @()
        $cert = gi $certPath
        $note = ""
        switch ($binding.port) {
            "8172" { $notes += "webdeploy" }
            default { $note = "" }
        }
        $resultExportedPfx = $false
        [byte[]]$certExportBytes = $null
        try { $certExportBytes = $cert.Export('pfx','') } catch {}
        if ($certExportBytes) {
            [system.io.file]::WriteAllBytes("$pfxOutDir\$($cert.thumbprint).pfx",$certExportBytes)
            $resultExportedPfx = $true
        }
        $harvestedCerts += new-object psobject -property @{
            fqdn = $fqdn
            ipaddress = $binding.ipaddress
            port = $binding.port
            friendlyname = $cert.friendlyname
            thumbprint = $cert.thumbprint
            subjectname = $cert.subjectname.name
            notafter = $cert.notafter
            notbefore = $cert.notbefore
            note = ($notes -join ',')
            location = 'certStore'
            exportedPfx = $resultExportedPfx
        }
    }
}

# Scour for stray .pfx files
write-host ("Searching drives for .pfx files: " + ($pfxSearchDrives -join ',')) -f green
$pfxFilesImported = @()
foreach ($drive in $pfxSearchDrives) {
    write-host "Searching $($drive):\ ..." -f green
    if (Test-Path "$($drive):\") {
        foreach ($file in (gci -path "$($drive):\" -recurse -filter *.pfx -ea 0)) {
            $notes = @()
            $notes += 'found on disk'
            write-host "file: $($file.fullname)" -f green
            $resultImport = $false
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            write-host "Attempting to import with provided guesses ..." -f yellow
            foreach ($guess in $pfxGuesses) {
                try {
                    $cert.Import($file.fullname, $guess, $x509KeySet)
                } catch{}
                if ($?) {
                    $resultImport = $true
                    break
                }
            }
            $resultExportedPfx = $false
            if ($resultImport) {
                write-host "Import Successful" -f yellow
                [byte[]]$certExportBytes = $null
                try { $certExportBytes = $cert.Export('pfx','') } catch {}
                if ($certExportBytes) {
                    [system.io.file]::WriteAllBytes("$pfxOutDir\$($cert.thumbprint).pfx",$certExportBytes)
                    $resultExportedPfx = $true
                } else {
                    $file | cp -destination $pfxOutDir
                }
            } else {
                write-warning "Import failure"
                $note += ', Unable to import pfx'
            }
            write-host "Exported Pfx?: $resultExportedPfx" -f yellow
            $harvestedCerts += new-object psobject -property @{
                fqdn = $fqdn
                ipaddress = $null
                port = $null
                friendlyname = $cert.friendlyname
                thumbprint = $cert.thumbprint
                subjectname = $cert.subjectname.name
                notafter = $cert.notafter
                notbefore = $cert.notbefore
                note = ($notes -join ',')
                location = $file.fullname
                exportedPfx = $resultExportedPfx
            }
        }
    }
}

# Save as XML
if ($harvestedCerts) {
    write-host "Saving manifest $outFilePath" -f green
    $harvestedCerts | export-clixml -path $outFilePath
    $outFile = gi $outFilePath
    $harvestedCerts | ft -auto
    ii $outFile.psparentpath
}

# Cleanup
if ((gci $pfxOutDir | measure).count -eq 0) {
    rm $pfxOutDir -recurse -force
}