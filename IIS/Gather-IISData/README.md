`Gather-IISData.ps1` is designed to be dot-sourced to produce `$parsedSiteWebAppData`

This primarly is useful for parsing uncommon (I hope) situations where you have many webapps nested under an IIS site.

It does some rudimentary parsing of `<remove>` tags which can be useful for trying to parse down HttpModule inheritance (which can be a pain). Example of that in `Audit-HttpModuleInherit.ps1`. Probably better to actually use some XDT transform APIs. 

e.g.
```
SITE 1
    /WEBAPP1
    /WEBAPP2
    /WEBAPP3
SITE 2
    /WEBAPP3
    /WEBAPP4
    /WEBAPP5
```