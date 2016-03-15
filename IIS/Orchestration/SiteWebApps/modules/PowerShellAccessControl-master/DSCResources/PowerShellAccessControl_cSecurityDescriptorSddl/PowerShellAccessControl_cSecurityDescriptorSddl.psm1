Import-Module $PSScriptRoot\..\..\PowerShellAccessControl.psd1
$InheritedAceRegEx = "\([^;]*;[^;]*ID([^;]*;){4}[^;]*\)"

function Get-TargetResource {
	[CmdletBinding()]
	[OutputType([System.Collections.Hashtable])]
	param (
		[parameter(Mandatory = $true)]
		[System.String]
		$Path,

		[parameter(Mandatory = $true)]
		[ValidateSet("File","Directory","RegistryKey","Service","WmiNamespace")]
		[System.String]
		$ObjectType
	)

    $Params = PrepareParams $PSBoundParameters 
    $GetSdParams = $Params.GetSdParams
    $GetSdParams.Audit = $true

    $SD = Get-SecurityDescriptor @GetSdParams

	$returnValue = @{
		Path = $Path
		ObjectType = $ObjectType
		Sddl = $SD.Sddl
	}

	$returnValue
}

function Set-TargetResource {
	[CmdletBinding()]
	param (
		[parameter(Mandatory = $true)]
		[System.String]
		$Path,

		[parameter(Mandatory = $true)]
		[ValidateSet("File","Directory","RegistryKey","Service","WmiNamespace")]
		[System.String]
		$ObjectType,

		[System.String]
		$Sddl
	)

    $Params = PrepareParams $PSBoundParameters 
    $GetSdParams = $Params.GetSdParams
    $NewSdParams = $Params.NewSdParams

    $SourceSD = New-AdaptedSecurityDescriptor @NewSdParams
    if ($SourceSD.GetAccessControlSections() -band [System.Security.AccessControl.AccessControlSections]::Audit) {
        Write-Verbose "Source SDDL contains SACL information; Get-SecurityDescriptor will be called with -Audit switch"
        $GetSdParams.Audit = $true
    }

    # Remove inherited ACEs and check the Sddl strings. This is set up so that it would be easy to
    # allow the inherited ACEs to be kept (maybe add a parameter for this?)
    Write-Debug "Removing inherited ACEs from SDDL string '$Sddl'"
    $SourceSddl = $SourceSD.Sddl -replace $InheritedAceRegEx
    Write-Debug "New SDDL string '$SourceSddl'"

    Write-Verbose ("Applying '{0}' SDDL to {1} ({2} sections)" -f $SourceSddl, $Path, $SourceSD.GetAccessControlSections())
    $SourceSD | Set-SecurityDescriptor -InputObject (Get-SecurityDescriptor @GetSdParams) -Sections ($SourceSD.GetAccessControlSections()) -Force
}

function Test-TargetResource {
	[CmdletBinding()]
	[OutputType([System.Boolean])]
	param (
		[parameter(Mandatory = $true)]
		[System.String]
		$Path,

		[parameter(Mandatory = $true)]
		[ValidateSet("File","Directory","RegistryKey","Service","WmiNamespace")]
		[System.String]
		$ObjectType,

		[System.String]
		$Sddl
	)

    $Params = PrepareParams $PSBoundParameters 
    $GetSdParams = $Params.GetSdParams
    $NewSdParams = $Params.NewSdParams

    $SourceSD = New-AdaptedSecurityDescriptor @NewSdParams
    if ($SourceSD.GetAccessControlSections() -band [System.Security.AccessControl.AccessControlSections]::Audit) {
        Write-Verbose "Source SDDL contains SACL information; Get-SecurityDescriptor will be called with -Audit switch"
        $GetSdParams.Audit = $true
    }

    $CurrentSD = Get-SecurityDescriptor @GetSdParams
    $CurrentSddl = $CurrentSD.SecurityDescriptor.GetSddlForm($SourceSD.GetAccessControlSections())

    # Remove inherited ACEs and check the Sddl strings. This is set up so that it would be easy to
    # allow the inherited ACEs to be kept (maybe add a parameter for this?)
    Write-Debug "Removing inherited ACEs from SDDL strings"
    $CurrentSddl = $CurrentSddl -replace $InheritedAceRegEx
    $SourceSddl = $SourceSD.Sddl -replace $InheritedAceRegEx

    Write-Debug "Sddl for ${Path}: $CurrentSddl"
    Write-Debug "Source Sddl: $SourceSddl"
    
    Write-Verbose "Comparing SDDL strings"
    $CurrentSddl -eq $SourceSddl
}

function PrepareParams {
    param(
        [hashtable] $Parameters
    )

    $GetSdParams = @{}
    $NewSdParams = @{}

    $NewSdParams.Verbose = $GetSdParams.Verbose = $false
    $GetSdParams.Path = $Parameters.Path

    $NewSdParams.Sddl = $Parameters.Sddl

    # The $Type parameter is handled with a ValidateSet(), and the strings mentioned there don't necessarily correspond to the 
    # System.Security.AccessControl.ResourceType enumeration that Get-SecurityDescriptor uses. Here's where we translate that:
    if ($Parameters.ContainsKey("ObjectType")) {
        switch ($Parameters.ObjectType) {
            
            { "File", "Directory" -contains $_ } {
# This actually works better if we let the module figure out if it's a file or directory
#                $GetSdParams.ObjectType = [System.Security.AccessControl.ResourceType]::FileObject
            }

            Directory {
                $NewSdParams.IsContainer = $true
            }

            RegistryKey {
#                $GetSdParams.ObjectType = [System.Security.AccessControl.ResourceType]::RegistryKey
                $NewSdParams.IsContainer = $true
            }

            Service {
                $GetSdParams.ObjectType = [System.Security.AccessControl.ResourceType]::Service
            }

            WmiNamespace {
                $NewSdParams.IsContainer = $true
                $GetSdParams.Path = "CimInstance: \\.\{0}:__SystemSecurity=@" -f $GetSdParams.Path
            }

            default {
                throw ('Unknown $Type parameter: {0}' -f $Parameters.Type)
            }
        }
    }

    @{
        GetSdParams = $GetSdParams
        NewSdParams = $NewSdParams
    }
}

Export-ModuleMember -Function *-TargetResource


