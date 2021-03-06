# ---------------------------------------------------------------
# Get-EventLog
# - Filters Security/Audit event logs much faster
# ---------------------------------------------------------------
Get-EventLog * | Foreach-Object { Get-EventLog -ea 0 -LogName $_.Log -After (Get-Date).AddDays(-1) }

# ---------------------------------------------------------------
# Get-WinEvent
# - Generally better (faster) except with security/audit logs
# - Vista/2008 require -FilterXml rather than FilterHashTable
# ---------------------------------------------------------------
# To get the [long] keywords you'll have to parse the StandardEventKeywords enum
[enum]::GetValues([type]'System.Diagnostics.Eventing.Reader.StandardEventKeywords') | % { new-object psobject -property @{ name=$_;value=$_.value__ } }

Get-WinEvent -ea 0 -LogName * -FilterHashTable @{
    logname='application'
    providername='.Net Runtime'
    keywords=36028797018963968
    ID=1023
    level='error'
}