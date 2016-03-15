<#
    .SYNOPSIS
    Fix Error 234 when using the IIS:\SslBindings provider
    
    .DESCRIPTION
    A missing 'SslCertStoreName' property in the registry can cause the IIS SSL Binding provider to not correctly work.
#>

try {
    gci iis:\SslBindings
}
catch {
    Get-ChildItem HKLM:\SYSTEM\CurrentControlSet\services\HTTP\Parameters\SslBindingInfo |
    ? { !($_ | Get-ItemProperty -Name 'SslCertStoreName' -ea 0) } |
    % {
        $_ | New-ItemProperty -Name 'SslCertStoreName' -Value "MY"
    }
    gci iis:\SslBindings
}