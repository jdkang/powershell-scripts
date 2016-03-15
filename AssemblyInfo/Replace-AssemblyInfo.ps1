

function Replace-AssemblyInfo {
param(
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
				HelpMessage="assemblyinfo.cs file")]
    [ValidateScript({Test-Path $_})]
    [Alias("PsPath")]
    [string[]]
    $Path,
	[Parameter(	Mandatory=$true,
				Position=1,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="assemblyinfo values")]
    [ValidateNotNullOrEmpty()]
    [hashtable]
    $assemblyInfoValues
)
PROCESS {
    foreach ($p in $Path) {
        write-verbose "Processing $p"
        $str = [io.file]::ReadAllText($p)
        
        foreach ($kv in $assemblyInfoValues.GetEnumerator()) {
            $pattern = ($kv.name + '\(".*?"\)')
            $str = ($str -replace $pattern, ($kv.name + '("' + $kv.value + '")'))
        }
        $str
    }
}
}