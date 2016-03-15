. .\Gather-IISData.ps1 # populates $parsedSiteWebAppData

$matchingApps = $parsedSiteWebAppData | ? { $_.name -match 'contosoApp(/.*)?$' }

$meh =
foreach ($x in $parsedSiteWebAppData) {
    $removedElements = $x.removedElements |
                        ? { $_.flatname -like '*HTTPModuleName*' }
    $tenantHostStatus = "inherited"
    if ($hostRewriteRemoveElements) {
        $tenantHostStatus = "removed"
        if ($hostRewriteRemoveElements.readded) {
            $tenantHostStatus = "overriden"
        }
    }
    new-object psobject -property @{
        name = $x.name
        targetframework = $x.targetframework
        tenantHostStatus = $tenantHostStatus
    }
}
$meh