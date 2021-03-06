function New-Struct {
param(
    [Parameter(Mandatory=$true,Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Name,
    [Parameter(Mandatory=$true,Position=1)]
    [hashtable]
    $properties
)
	switch($Properties.Keys){{$_ -isnot [String]}{throw "Invalid Syntax"}}
	switch($Properties.Values){{$_ -isnot [type]}{throw "Invalid Syntax"}}
	$csharpStructCode = "
    using System;
    public struct $Name {

      $($Properties.Keys | % { "  public {0} {1};`n" -f $Properties[$_],($_.ToUpper()[0] + $_.SubString(1)) })
      public $Name ($( [String]::join(',',($Properties.Keys | % { "{0} {1}" -f $Properties[$_],($_.ToLower()) })) )) {
        $($Properties.Keys | % { "    {0} = {1};`n" -f ($_.ToUpper()[0] + $_.SubString(1)),($_.ToLower()) })
      }
    }"
    Add-Type $csharpStructCode
}
