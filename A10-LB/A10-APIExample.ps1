#requires -version 3
[CmdletBinding()]
param(
	[ValidateScript({Test-Connection $_ -Quiet -Count 3})]
	$lbVip = 'ADC.CONTOSO.LOCAL',
	[ValidateNotNullOrEmpty()]
	$username = 'USERNAME',
	[ValidateNotNullOrEmpty()]
	$password = 'PASSWORD'
)
# Convenience SPLATs
$verbug = @{
	verbose = [bool]$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -or ($VerbosePreference -eq 'continue')
	debug = [bool]$PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent -or ($DebugPreference -eq 'continue')
}

$a10loginSplat = @{
    username = $username
    password = $password
}

#https://github.com/ericchou-python/A10_Networks/blob/master/slb_vipCreate.py

###############################################################\
# init
###############################################################
#[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } # ugh
$protocol = 'https'
$apiPath = '/services/rest/V2/'
$script:a10_uri = "$($protocol)://$($lbVip)$apiPath"
$script:a10_session = $null
$script:a10_sessionId = $null

###############################################################
# func
###############################################################
# ---------------------------------------
# Helpers
# ---------------------------------------
function New-ImmutableObject
{
param(
    [Parameter(Mandatory=$true)]
    $object
)
    write-verbose "new-immutableobject: $($object.gettype().name)"
    $immutable = New-Object PSObject
    $ht = @{}
    if ($object.GetType().Name -eq 'PsCustomObject') {
        write-verbose "new-immutableobject: Converting pscustomobject"
        $object.PsObject.Properties
    } elseif ($object.GetType().name -eq 'hashtable') {
        $ht = $object.Clone()
    }
    
    $ht.Keys | %{ 
        $value = $ht[$_]
        $closure = { $value }.GetNewClosure()
        $immutable | Add-Member -name $_ -memberType ScriptProperty -value $closure
    }
    return $immutable
}

# ---------------------------------------
# Generic A10 axAPI Wrappers
# ---------------------------------------
Add-Type -AssemblyName 'system.web' -ea 0

function New-LBSessionId {
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $username,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $password
)
	write-verbose "Requesting New Session ID with username $username"
	$respTuple = Invoke-A10RestMethod -Method POST -apiMethod 'authenticate' -body @{
        username = $username
        password = $password
    } @verbug
	if ($respTuple.data.session_id) {
        write-verbose "sessionId = $($resp.session_id)"
        return $resp.session_id
	} else {
		write-verbose "No Session ID returned"
		return $null
	}
}

function Invoke-A10RestMethod {
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('post','put','delete')]
    [string]
    $method,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $apiMethod,
    [hashtable]
    $body,
    [hashtable]
    $extraQueryStringArgs,
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]
    $sessionId
)
    $requestBodySplat = @{}
    $queryString = [system.web.httputility]::ParseQueryString('')
    
    # Body
    if ($body -and ($body.Count -gt 0)) {
        $requestBodySplat = @{
            Body = ($body | ConvertTo-Json)
            ContentType = 'application/json'
        }
    }
    #write-verbose ($requestBodySplat.Body | out-string)
    
    # QueryString
    if ($sessionId) {
        $queryString['session_id'] = $sessionId
    }
    $queryString['format'] = 'json'
    $queryString['method'] = $apiMethod
    if ($extraQueryStringArgs -and ($extraQueryStringArgs.Count -gt 0)) {
        foreach ($kv in $extraQueryStringArgs.GetEnumerator()) {
            $queryString[$kv.name] = $kv.value
        }
    }
    
    # Request
    $fullRequestUrl = $script:a10_uri
    if ($queryString.ToString()) {
        $fullRequestUrl += ('?' + $queryString.ToString())
    }
    
    $r = Invoke-WebRequest -Method $method -Uri $fullRequestUrl @requestBodySplat @verbug
    if (!$?) { throw "Error Making A10 Request" }
    
    $respData = @{}
    $respError = @{}
    
    if ($resp.Response.Status -eq 'fail') {
        $respError = @{
            code = [int]$r.Response.Err.Code
            message = $r.Response.Err.Msg
        }
        write-error ($r.Response.Err | fl | out-string)
    } else {
        $respData = $r
    }
    
    return New-ImmutableObject @{
        data = New-ImmutableObject $respData
        error = New-ImmutableObject $respError
    }
}
# ---------------------------------------
# Specific A10 axAPI Implementations
# ---------------------------------------
function Get-A10SeviceGroup {
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $sgName
)
	$resp = Invoke-A10RestMethod -Method POST -apiMethod 'slb.service_group.search' -body @{
        name = $sgName
    } @verbug
    $resp
}

#############################################################
# main
#############################################################
$sId = New-LBSessionId @a10loginSplat @verbug

 