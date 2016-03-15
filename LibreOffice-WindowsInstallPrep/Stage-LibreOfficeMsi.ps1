<#
    .SYNOPSIS
    Extracts a LibreOffice installation and applies changes to alter the default Save AS behavior to MS Office formats.
    
    .DESCRIPTION
    Extracts the MSI and then applies changes to the XML configuration to save as MS Office formats (e.g. .docx) by default.
    
    Also drops in a copy of a .BAT with 'ideal' MSI install args.
#>
param(
    # LibreOffice MSI
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSCript({ (Test-Path $_) -and (gi $_).Extension -eq '.msi' })]
	[string]$msi,
    # Copy contents after extraction/modification
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
	[string]$copyTo,
    # Don't apply the "Save As" fix for office documents
	[switch]$skipSaveAsFix
)

$currentScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
Push-Location $currentScriptPath


function Replace-DefaultSaveAs {
param(
	[string]$name,
	[string]$currValue,
	[string]$newValue,
	[string]$regLocation,
	[string]$backupXmlPath
)
    # This should probably be fixed to use xpath

	$xmlFile = "$regLocation\$name.xcd"
	if (Test-Path $xmlFile) {
		write-host "Reading $xmlFile" -foregroundcolor magenta
		$encoding = New-Object System.Text.ASCIIEncoding # try to always declare your encoding with file.readalltext
		$buffer = [system.io.file]::ReadAllText((Resolve-Path $xmlFile).ProviderPath, $encoding)
		$currValueElement = "<prop oor:name=`"ooSetupFactoryDefaultFilter`"><value>$currValue</value>"
		$newValueElement = "<prop oor:name=`"ooSetupFactoryDefaultFilter`"><value>$newValue</value>"
		write-host "Replacing default save as element" -foregroundcolor gray
		
		if ($buffer.Contains($currValueElement)) {
			$newXml = $buffer.Replace($currValueElement,$newValueElement)
			
			$xmlFileName = $xmlFile.Name
			$backupXmlPathTs = "$backupXmlPath\xmlbackups-$(Get-Date -f 'yyyyMMddHHmmss')"
			md $backupXmlPathTs -force | out-null
			
			write-host "Backing up file $xmlFileName to $backupXmlPathTs" -foregroundcolor gray
			
			move-item -path $xmlFile -destination "$backupXmlPathTs\$xmlFileName"
			
			write-host "Saving $xmlFile" -foregroundcolor gray
			$newXml | sc -path $xmlFile
		} else {
			write-warning "Could not find $currValueElement in the file"
		}
	} else {
		write-warning "Could not find file - $xmlFile"
	}
}


# http://listarchives.libreoffice.org/global/users/msg12888.html
if (!(Test-Path $msi)) {
    write-warning "Could not find file $msi"
    exit
}

$msiFullPath = (gi $msi).FullName
$msiFileName = (gi $msi).Name

$extractLocation = "$currentScriptPath\$msiFileName-$(Get-Date -f 'yyyyMMddHHmmss')"
md $extractLocation -force | out-null
write-host "Extracting MSI to $extractLocation" -foregroundcolor green
cmd /c "msiexec /a $msiFullPath /qb TARGETDIR=$extractLocation"
if (!$? -or ($LASTEXITCODE -ne 0)) {
    write-warning "ISSUE WITH MSIEXEC - Last Exit Code = $LASTEXITCODE"
    exit
} else {
    write-host "Extractiong went OK?" -foregroundcolor green
}

$registryLocation = "$extractLocation\share\registry"
write-host "Attempting to change XML configs to set default save as format..." -foregroundcolor green


if (!$skipSaveAsFix) {
	#writer
	Replace-DefaultSaveAs -name 'writer' -currValue 'writer8' -newValue 'MS Word 97' -regLocation $registryLocation -backupXmlPath $extractLocation

	#calc
	Replace-DefaultSaveAs -name 'calc' -currValue 'calc8' -newValue 'MS Excel 97' -regLocation $registryLocation -backupXmlPath $extractLocation

	#impress
	Replace-DefaultSaveAs -name 'impress' -currValue 'impress8' -newValue 'MS PowerPoint 97' -regLocation $registryLocation -backupXmlPath $extractLocation
} else {
	write-host "skipping save as changes." -foregroundcolor green
}

#drop in bat file
$batFileTemplate = "INSTALL-LIBRE.bat"
write-host "Creating $batFileTemplate with proper install parameters" -foregroundcolor green
$batFileBuffer = gc ".\$batFileTemplate"
$newBatFile = $batFileBuffer.Replace("LIBREMSIFILE",$msiFileName)
$newBatFile | sc -path "$extractLocation\$batFileTemplate"

#copy files
if ($copyTo) {
	if (Test-Path $copyTo) {
		$srcFolderName = (gi $extractLocation).Name
		$dest = "$copyTo\$srcFolderName" 
		md $dest -force | out-null
		robocopy $extractLocation $dest /e /mt /w:2 /r:2 /xj
	} else {
		write-warning "$copyTo is not a valid location."
	}
}

<#
<prop oor:name="ooSetupFactoryDefaultFilter"><value>writer8</value>

writer.xcd 
writer8
MS Word 97

calc.xcd
calc8
MS Excel 97

impress.xcd
impress8
MS PowerPoint 97
#>


