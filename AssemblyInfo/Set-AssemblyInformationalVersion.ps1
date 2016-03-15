#requiers -version 3
param(
    [Parameter(Mandatory=$true,HelpMessage="Root file path (recursive)")]
    [string]
    $path,
    [string]
    $propertyName='AssemblyInformationalVersion'
)
if (!(Test-path $path)) { throw "invalid path" }

$propertyNameRegex = '\[assembly:\s?' + $propertyName + '\('

function Get-FileEncoding
{
    [CmdletBinding()] Param (
     [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
     [string]$Path
    )
 
    [byte[]]$byte = get-content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $Path
 
    if ( $byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf )
    { Write-Output 'UTF8' }
    elseif ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff)
    { Write-Output 'Unicode' }
    elseif ($byte[0] -eq 0 -and $byte[1] -eq 0 -and $byte[2] -eq 0xfe -and $byte[3] -eq 0xff)
    { Write-Output 'UTF32' }
    elseif ($byte[0] -eq 0x2b -and $byte[1] -eq 0x2f -and $byte[2] -eq 0x76)
    { Write-Output 'UTF7'}
    else
    { Write-Output 'ASCII' }
}
write-host "Directory (recursive): " -n -f cyan
write-host $path -f magenta
write-host "Matching against " -n -f cyan
write-host "$propertyNameRegex" -f magenta
foreach ($assemblyInfo in (gci -path $path -recurse -filter assemblyinfo.cs)) {
    $c = [System.IO.File]::ReadAllText($assemblyInfo.FullName)
    $relativePath = $assemblyInfo.fullname.replace($path,'.\')
    write-host "$relativePath ... " -n -f green
    if ($c -notmatch $propertyNameRegex) {
        write-host "Adding $propertyName" -f yellow
        $c = $c.trim() + "`r`n"
        $c += ('[assembly: ' + $propertyName + '("1.0.0.0")]')
        $c | Out-File -Force -FilePath $assemblyInfo.pspath -Encoding (Get-FileEncoding $assemblyInfo.pspath)
    } else {
        write-host "OK" -f yellow
    }
}