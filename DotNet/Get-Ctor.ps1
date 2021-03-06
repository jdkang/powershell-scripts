function Get-Ctor {
<#
    .SYNOPSIS
    Prints the constructors for a .NET type
#>
param(
    [type]$type,
    [switch]$fullName
)
    foreach ($ctor in $type.GetConstructors()) {
        $parametersStr = @()
        foreach ($p in $ctor.GetParameters()) {
            $pType = $null
            
            if (!$fullName) {
                $pType = $p.parametertype.name
            } else {
                $pType = $p.parametertype.fullname
            }
            $parametersStr += "$pType $($p.Name)"
        }
        write-host "$($type.Name) (" -n -f green
        write-host ($parametersStr -join ', ') -n -f yellow
        write-host ')' -f green
    }
}
