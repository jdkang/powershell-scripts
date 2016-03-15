ipmo webadministration -ea 0

$sitesAndWebapps = @()
$sitesAndWebapps += (get-website)
$sitesAndWebapps += (get-webapplication)
$parsedSiteWebAppData = @()

foreach ($site in (get-website)) {
    # ---- Site ----
    write-host "SITE: " -n -f green
    write-host $site.name -f yellow
    
    # Site web.config
    $conf = $null
    $siteWebConf = "$($site.physicalpath)\web.config"
    if (!(Test-Path $siteWebConf)) {
        write-warning "Could not find path $siteWebConf"
    } else {
        write-verbose "Parsing $siteWebConf"
        [xml]$conf = (gc $siteWebConf)
    }
    
    # Site AppPool
    $appPoolUser = ""
    $appPool = $null
    $appPoolStr = ""
    $appPoolStr = $site.applicationPool
    if ($appPoolStr) {
        $appPool = (gi "IIS:\appPools\$appPoolStr")
        $appPoolUser = $appPool.ProcessModel.username
    }
    
    # Find <remove> elements
    $removedElements = @()
    $removeNodes = $conf | Select-Xml -Xpath '//remove'
    foreach ($n in $removeNodes) {
        $nodeParentAtt = ""
        $nodeParentAtt = (($n.node.ParentNode.Attributes |
                         %{ $_.LocalName + '="' + $_.'#text' + '"' }) -join ' ')
        if ($nodeParentAtt) { $nodeParentAtt = ' ' + $nodeParentAtt }
        $xmlStr = ("<$($n.node.parentNode.localname)" +
                  $nodeParentAtt +
                  '>' + $n.Node.OuterXml)
        $nodeAttName = ''
        $removedItemName = $n.node.Attributes |
                           ? { $_.name -eq 'name' } |
                           Select -expand "#text"
        if ($removedItemName) {
            $nodeAttName = ('.' + $removedItemName)
        }
        $flatName = ($n.node.parentnode.localname +
                    '.' + $n.node.localname + $nodeAttName)
        
        $readded = $null
        if ($removedItemName) {
            $addNodes = $n.node.ParentNode.ChildNodes |
                       ? { $_.localname -eq 'add' } |
                       Select -expand name
            if ($addNodes -contains $removedItemname)
            {
                $readded = $true
            }
        }
        
        $removedElements += new-object psobject -property @{
            flatname = $flatname
            xml = $xmlStr
            readded = $readded
        }
    }
    
    # Site Info
    $parsedSiteWebAppData += new-object psobject -property @{
        name = $site.name
        type = 'site'
        appPool = $appPoolStr
        appPoolUser = $appPool.ProcessModel.username
        physicalPath = $site.physicalpath
        appSettings = $conf.configuration.appsettings.add
        connectionStrings = $conf.configuration.connectionStrings.add
        targetFramework = $conf.configuration."system.web".compilation.targetFramework
        removedElements = $removedElements
    }
    
    # ---- Site: WebApps ----
    $webapps = get-webapplication -site $site.name
    foreach ($webapp in $webapps) {
        if ($webapp) {
            write-host "-- WEBAPP: " -n -f green
            write-host $webapp.path -f yellow
            
            # WebApp web.config
            $conf = $null
            $appWebConf = "$($webapp.physicalpath)\web.config"
            if (!(Test-Path $appWebConf)) {
                write-warning "Could not find path $appWebConf"
            } else {
                write-verbose "Parsing $appWebConf"
                [xml]$conf = (gc $appWebConf)
            }
            
            # WebApp AppPool
            $appPoolUser = ""
            $appPool = $null
            $appPoolStr = ""
            $appPoolStr = $webapp.applicationPool
            if ($appPoolStr) {
                $appPool = (gi "IIS:\appPools\$appPoolStr")
                $appPoolUser = $appPool.ProcessModel.username
            }
            
            # Find <remove> elements
            $removedElements = @()
            $removeNodes = $conf | Select-Xml -Xpath '//remove'
            foreach ($n in $removeNodes) {
                $nodeParentAtt = ""
                $nodeParentAtt = (($n.node.ParentNode.Attributes |
                                 %{ $_.LocalName + '="' + $_.'#text' + '"' }) -join ' ')
                if ($nodeParentAtt) { $nodeParentAtt = ' ' + $nodeParentAtt }
                $xmlStr = ("<$($n.node.parentNode.localname)" +
                          $nodeParentAtt +
                          '>' + $n.Node.OuterXml)
                $nodeAttName = ''
                $removedItemName = $n.node.Attributes |
                                   ? { $_.name -eq 'name' } |
                                   Select -expand "#text"
                if ($removedItemName) {
                    $nodeAttName = ('.' + $removedItemName)
                }
                $flatName = ($n.node.parentnode.localname +
                            '.' + $n.node.localname + $nodeAttName)
                $readded = $null
                if ($removedItemName) {
                    $addNodes = $n.node.ParentNode.ChildNodes |
                               ? { $_.localname -eq 'add' } |
                               Select -expand name
                    if ($addNodes -contains $removedItemname)
                    {
                        $readded = $true
                    }
                }
                                    
                $removedElements += new-object psobject -property @{
                    flatname = $flatname
                    xml = $xmlStr
                    readded = $readded
                }
            }
            
            # WebApp info
            $parsedSiteWebAppData += new-object psobject -property @{
                name = "$($site.name)$($webapp.path)"
                type = 'webapp'
                appPool = $appPoolStr
                appPoolUser = $appPool.ProcessModel.username
                physicalPath = $webapp.physicalpath
                appSettings = $conf.configuration.appsettings.add
                connectionStrings = $conf.configuration.connectionStrings.add
                targetFramework = $conf.configuration."system.web".compilation.targetFramework
                removedElements = $removedElements
            }
        }
    }
}