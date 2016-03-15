param(
    [switch]$doNotNotifyChatrooms
)
#--------------------------------------------------------
# conf
#--------------------------------------------------------
# use this if the specific room doesn't have its own api key generated
# label: fogbugz
$hipchatGlobalApiKey = 'aaaaaaaaaaaaaaaaaaaaaaaaa'

# request throttling
$hipchatThrottleRequestsSeconds = 3

# hipchat rooms
$defaultRoom = new-object psobject -property @{
    roomid = 1111111
    apitoken = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
}
$teamAwesomeChatRoom = new-object psobject -property @{
    roomid = 1111112
    apitoken = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
}

# Map Projects to Rooms                         
$projectRoomMapping = @{}
$projectRoomMapping.Add('Project X',$teamAwesomeChatRoom)

$hipchatroomDefault = $defaultRoom

# fb query
$fbQuery = 'status:"active" milestone:"SOME MILESTONE TO MONITOR" assignedto:"unassigned" (priority:"1" OR priority:"2") orderby:"priority" orderby:"Due"'

# fogbugz info    
$fbUser = 'auto@contoso.local'
$fbPassword = 'password'
$fbUri = 'https://fogbugz.contoso.local'

#--------------------------------------------------------
# time
#--------------------------------------------------------
Function Convert-FromUnixdate {
param(
    [string]$UnixDate
)
   [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))
}

function Format-TimeSpan {
param(
    [timespan]$timespan
)
    $retStr = ''
    if ($timespan.totalseconds -ge 0) {
        $retStr += 'in'
    }
    if ($timespan.days -ne 0) {
        $retStr += (' ' + [math]::abs($timespan.days).tostring() + " days")
    }
    if ($timespan.hours -ne 0) {
        $retStr += (' ' + [math]::abs($timespan.hours).tostring() + " hours")
    }
    if ($timespan.minutes -ne 0) {
        $retStr += (' ' + [math]::abs($timespan.minutes).tostring() + " minutes")
    }
    if ($timespan.totalseconds -lt 0) {
        $retStr += ' ago'
    }
    
    $retStr.Trim()
}

function ConvertISO8601To-DateTime {
param(
	[string]$string,
	[switch]$utc
)
    $iso8601DateTimeFormat = "yyyy-MM-ddTHH:mm:ssZ"
	if ($string) {
        write-verbose "Converting datetime $string"
		$culture = [System.Globalization.CultureInfo]::InvariantCulture
		$dtStyle = [System.Globalization.DateTimeStyles]::RoundtripKind
		$dt = [datetimeoffset]::ParseExact($string,$iso8601DateTimeFormat,$culture,$dtStyle)
		if ($dt) {
			if (!$utc) {
				$dt.LocalDateTime
			} else {
				$dt.UtcDateTime
			}
		} else {
			$null
		}
	}
}

function ConvertDateTimeTo-ISO8601Utc {
param(
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="datetime")]
    [ValidateNotNullOrEmpty()]
    [datetime]
    $datetime
)
    # z is for zulu
    $dateTime.ToUniversalTime().ToString('s') + 'Z'
}

#--------------------------------------------------------
# fogbugz scaffolding
#--------------------------------------------------------
$script:fbApiUri = ''
$script:fbApiToken = ''

function Get-ApiUri {
    if (!$script:fbApiUri) {
        write-verbose "FogBugz API Info not set"
        write-verbose "Fetching fogbugz api info"
        $resp = Invoke-RestMethod -Uri "$fbUri/api.xml"
        if ($resp) {
            write-verbose "FogBugz API: $($resp.response.version).$($resp.response.minversion)"
            $script:fbApiUri = "$fbUri/" + $resp.response.url.replace('?','')
            write-verbose "API URI: $($script:fbApiUri)"
        } else {
            throw "Could not fetch fogbugz api uri"
        }
    }
}

function Invoke-FbCmd {
param(
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz cmd")]
    [ValidateNotNullOrEmpty()]
    [string]$cmd='',
	[Parameter(	Mandatory=$false,
				Position=1,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="extra arguments")]
    [hashtable]$extraArgs=@{},
	[Parameter(	Mandatory=$false,
				Position=2,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="use existing api token")]
    [switch]$useToken,
	[Parameter(	Mandatory=$false,
				Position=3,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="xml node name")]
    [string]$xmlNodeName
)
    if ($extraArgs -eq $null) { throw "extraArgs cannot be null" }
    
    Get-ApiUri
    $extraArgs.Add('cmd',$cmd)
    if ($useToken) {
        Ensure-FbToken
        $extraArgs.Add('token',$script:fbApiToken)
    }
    
    $extraARgsKvStr = ''
    $extraArgs.GetEnumerator() | %{ $extraARgsKvStr += "$($_.key)=$($_.value);" }
    write-verbose "Executing Fogbugz cmd $cmd $extraARgsKvStr"
    $resp = Invoke-RestMethod -uri $script:fbApiUri -Method POST -Body $extraArgs
    if (!$?) { throw "Error executing fogbugz cmd" }
    if ($resp.response.error) {
        $props = @{
            errorCode = $resp.response.error.code
            errorText = $resp.response.error."#cdata-section"
        }
        throw (new-object psobject -property $props)
    }
    
    $xmlResp = $resp.response
    if ($xmlNodeName) {
        $nodes = $xmlResp.SelectNodes("//$xmlNodeName")
        if (!$nodes) {
            write-verbose "No nodes like xpath //$xmlNodeName"
        } else {
            $nodes
        }
    } else {
        $xmlResp
    }
}

function Validate-FbToken {
param(
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz token")]
    [ValidateNotNullOrEmpty()]
    [string]$token
)
    $tokenValid = $null
    try {
        write-verbose "Validating fogbugz token $token"
        $resp = Invoke-FbCmd -cmd 'logon' -extraArgs @{token=$token}
        $tokenValid = $true
    }
    catch {
        if ($_.targetobject.errorCode -eq 3) {
            write-verbose "fogbugz says token $token is invalid."
            $tokenValid = $false
        } else {
            throw "Unable to validate token"
        }
    }
    finally {
      $tokenValid
    }
}

function Ensure-FbToken {
param(
    [switch]$newToken
)
    if ($newToken) {
        write-verbose "Requesting new token with username $fbUser"
        $loginInfo = @{
                        email=$fbUser
                        password=$fbPassword
                      }
        $resp = Invoke-FbCmd -cmd 'logon' -extraArgs $loginInfo
        if ($resp.token) {
            $script:fbApiToken = $resp.token."#cdata-section"
            $tokenIsValid = Validate-FbToken -token $script:fbApiToken
            if (!$tokenIsValid) {
                throw "Attempt to get a new valid fogbugz api token failed"
            }
        } else {
            throw "Unable to get new Fb Token"
        }
    } else {
        if (!$script:fbApiToken) {
            write-verbose "No fogbugz token set. Requesting new token."
            Ensure-FbToken -newToken
        } else {
            write-verbose "Validating current fogbugz token $($script:fbApiToken)"
            $tokenIsValid = Validate-FbToken -token $script:fbApiToken
            if (!$tokenIsValid) {
                write-verbose "Requesting no fogbugz api token"
                Ensure-FbToken -newToken
            }
        }
    }
}

function Invoke-FbLogoff {
    if ($script:fbApiToken) {
        Invoke-FbCmd -cmd 'logoff' -useToken
        write-verbose "Goodbye"
    }
}
#--------------------------------------------------------
# fogbugz command wrappers
#--------------------------------------------------------
function ConvertFixForTo-PsObject {
param (
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fixfor")]
    [ValidateNotNullOrEmpty()]
    [System.Xml.XmlElement]
    $xmlresp
)
    $nodes = $xmlresp.SelectNodes('//fixfor')
    foreach ($fixfor in $nodes) {
        new-object psobject -property @{
            ixFixFor = $fixfor.ixfixfor
            sFixFor = $fixfor.sFixFor."#cdata-section"
            fDeleted = $fixfor.fDeleted
            dt = (ConvertISO8601To-DateTime -string $fixfor.dt)
            dtStart = (ConvertISO8601To-DateTime -string $fixfor.dtStart)
            sStartNote = $fixfor.sStartNote
            ixProject = $fixfor.ixProject
            sProject = $fixfor.sProject."#cdata-section"
            setixFixForDependency = $fixfor.setixFixForDependency
            fReallyDeleted = $fixfor.fReallyDeleted
        }
    }
}

function ConvertFbXml-ToPsObject {
param(
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="xmlelement resp")]
    [ValidateNotNullOrEmpty()]
    [System.Xml.XmlElement[]]
    $xmlelements,
	[Parameter(	Mandatory=$false,
				Position=1,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="properties to attempt to convert to datetime")]
    [string[]]
    $colsConvertToDateTime
)
    PROCESS {
    foreach ($xmlelement in $xmlelements) {
        $psProps = @{}
        $properties = $xmlelement | gm -membertype property | select -expand Name
        foreach ($property in $properties) {
            $val = $null
            $val = $xmlelement."$property"
            if ($val.gettype().name -eq 'xmlelement') {
                $subXmlEle = $null
                $subXmlEle = ($val | ConvertFbXml-ToPsObject)
                if (!$subXmlEle) {
                    throw "Cannot parse sub xml element"
                }
                $psProps.Add($property,$subXmlEle)
            } else {
                if ($colsConvertToDateTime -contains $property) {
                    $datetimeobj = $null
                    $datetimeobj = (ConvertISO8601To-DateTime -string $val)
                    if (!$datetimeobj) {
                        throw "Cannot convert value to datetime."
                    }
                    $psProps.Add($property,$datetimeobj)
                } else {
                    if ($property -eq '#cdata-section') {
                        return $val
                    } else {
                        $psProps.Add($property,$val)
                    }
                }
            }
        }
        new-object psobject -property $psProps
    }
    }
}

function Get-Person {
param(
	[Parameter(	Mandatory=$false,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="person id")]
    [Alias('ixPerson')]
    [string]
    $personId
)
    $extraArgs = @{}
    $extraArgs.Add('ixPerson',$personId)
    $xmlresp = Invoke-FbCmd -cmd 'viewPerson' -xmlnodename 'person' -useToken -extraArgs $extraArgs
    $xmlresp | ConvertFbXml-ToPsObject
}

function Get-Milestones {
param(
    [string]$project,
    [Alias('fIncludeDeleted')]
    [switch]$includeInactive,
    [Alias('fIncludeReallyDeleted')]
    [switch]$includeDeleted
)
    $extraArgs = @{}
    if ($includeInactive) {
        $extraArgs.Add('fIncludeDeleted',1)
    }
    if ($includeDeleted) {
        $extraArgs.Add('fIncludeReallyDeleted',1)
    }
    $resp = Invoke-FbCmd -cmd 'listFixFors' -useToken -extraArgs $extraArgs
    ConvertFixForTo-PsObject -xmlresp $resp
}

function New-Search {
param(
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz search string")]
    [Alias("query")]
    [string]$q,
	[Parameter(	Mandatory=$false,
				Position=1,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="cols to return")]
    [Alias("columns")]
    [string[]]$cols
)
    $colStr = ''
    $extraArgs = @{}
    $extraArgs.Add('q',$q)
    if ($cols) {
        $colStr = $cols -join ','
        $extraArgs.Add('cols',$colStr)
    }
    $xmlresp = Invoke-FbCmd -cmd 'search' -xmlnodename 'case' -useToken -extraArgs $extraArgs
    $xmlresp | ConvertFbXml-ToPsObject
}

function New-Milestone {
param(
	[Parameter(	Mandatory=$false,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz project")]
    [Alias('ixProject')]
    [int]
    $project=-1, # -1 = global
	[Parameter(	Mandatory=$true,
				Position=1,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz milestone title")]
    [ValidateNotNullOrEmpty()]
    [string]
    [Alias('sFixFor')]
    $title,
	[Parameter(	Mandatory=$false,
				Position=2,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz milestone dt Release")]
    [datetime]
    $dtRelease,
	[Parameter(	Mandatory=$false,
				Position=3,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz milestone dt Start")]
    [datetime]
    $dtStart,
	[Parameter(	Mandatory=$false,
				Position=4,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz milestone dt Start")]
    [string]
    [Alias('sStartNote')]
    $startNote,
	[Parameter(	Mandatory=$false,
				Position=5,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="fogbugz milestone dt Start")]
    [switch]
    [Alias('fAssignable')]
    $assignable
)
    $extraArgs = @{}
    $extraArgs.Add('ixProject',$project)
    $extraArgs.Add('sFixFor', $title)
    if ($assignable) {
        $extraArgs.Add('fAssignable',1)
    } else {
        $extraArgs.Add('fAssignable',0)
    }
    if ($dtRelease) {
        $dtReleaseString = ConvertDateTimeTo-ISO8601Utc -datetime $dtRelease
        if ($dtReleaseString) { $extraARgs.Add('dtRelease',$dtReleaseString) }
    }
    if ($dtStart) {
        $dtStartString = ConvertDateTimeTo-ISO8601Utc -datetime $dtStart
        if ($dtReleaseString) { $extraARgs.Add('dtStart',$dtStartString) }
    }
    if ($startNote) { $extraArgs.Add('sStartNote',$startNote) }
    
    write-verbose "Creating new milestone"
    
    $resp = Invoke-FbCmd -cmd 'newFixFor' -useToken -extraArgs $extraArgs
    if ($resp) {
        ConvertFixForTo-PsObject -xmlresp $resp
        write-verbose ($resp | fl | out-string)
    }
}

function Get-FbCaseFilters {
    $resp = Invoke-FbCmd -cmd 'listFilters' -useToken
    $resp.filters.filter
}
#--------------------------------------------------------
# hipchat
#--------------------------------------------------------
function Check-HipchatStandoff {
param(
    [int]$xLimitRemaining,
    [int]$xFloodControlReset
)
    if ($xLimitRemaining -le 1) {
        $nextReset = Convert-FromUnixdate $xFloodControlReset
        $waitSeconds = ( new-timespan (get-date) $nextReset ).totalseconds
        if ($waitSeconds -gt 0) {
            start-sleep -seconds ($waitSeconds+1)
        }
    }
}

$script:hipchatNotifyThrottleStopWatch = $null
$script:hipchatNotifySpamThrottleStopWatch = $null
$script:notificationsSent = 0

function Send-HipchatNotification {
param(
	[Parameter(	Mandatory=$true,
				Position=0,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="message")]
    [ValidateNotNullOrEmpty()]
    [string]
    $message,
	[Parameter(	Mandatory=$true,
				Position=1,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="room id")]
    [Alias('APIid')]
    [int]
    $roomid,
	[Parameter(	Mandatory=$false,
				Position=2,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="message color")]
    [ValidateSet("yellow","green","red","purple","gray","random")]
    [string]
    $color='gray',
	[Parameter(	Mandatory=$false,
				Position=3,
				ValueFromPipeline=$false,
				ValueFromPipelineByPropertyName=$false,
				HelpMessage="api key")]
    [string]
    $apitoken,
    [switch]$notify = $false,
    [ValidateSet("text","html")]
    [string]$messageFormat="text"
)
    if (!$apitoken) {
        $apitoken = $hipchatGlobalApiKey
    }

    #$uri = "https://api.hipchat.com/v2/room/$roomid/notification?auth_token=$apitoken"
    $uri = "https://api.hipchat.com/v2/room/$roomid/notification"
    $postBodyObj = new-object psobject @{
        color=$color
        message=$message
        notify=$notify.tostring()
        message_format=$messageFormat
    }
    $postBody = ConvertTo-Json -InputObject $postBodyObj
    #$postStr = [System.Text.Encoding]::UTF8.GetBytes($postBody)
    
    $headers = @{}
    $headers.Add('Authorization',"Bearer $apitoken")
    
    # Internal thorttling msg/s
    if ($script:hipchatNotifyThrottleStopWatch) {
        $elapsedTimeSeconds = $null
        $elapsedTimeSeconds = $script:hipchatNotifyThrottleStopWatch.Elapsed.Seconds
        write-verbose "stopwatch: $elapsedTimeSeconds seconds since last hipchat message sent."
        if ($elapsedTimeSeconds -le $hipchatThrottleRequestsSeconds) {
            $timeToWaitSeconds = $hipchatThrottleRequestsSeconds - $elapsedTimeSeconds
            write-verbose "throttle: Waiting $timeToWaitSeconds seconds until sending next hipchat request."
            start-sleep -seconds $timeToWaitSeconds
            write-verbose "stopwatch: Resetting stopwatch"
            $script:hipchatNotifyThrottleStopWatch = $null
        }
    }
    
    # hipchat throttling 30msg/minute
    if ($script:hipchatNotifySpamThrottleStopWatch) {
        if ($script:notificationsSent -ge 29) {
            $spamSleepSeconds = 61 - $script:hipchatNotifySpamThrottleStopWatch.Elapsed.Seconds
            write-verbose "spam: 29 messages sent. Sleeping $spamSleepSeconds"
            start-sleep -seconds $spamSleepSeconds
        }
                
        # Rollover timer
        if ( $script:hipchatNotifySpamThrottleStopWatch.Elapsed.Seconds -ge 59 ) {
            write-verbose "spam: elapsed time $($script:hipchatNotifySpamThrottleStopWatch.Elapsed.Seconds) seconds. Resetting notifications sent counter and timer."
            $script:notificationsSent = 0
            $script:hipchatNotifySpamThrottleStopWatch = $null
        }
    }
    
    write-verbose "Sending hipchat request $uri"
    $resp = $null
    $resp = try {
        Invoke-WebRequest -Method 'POST' -Uri $uri -Headers $headers -Body $postBody -ContentType 'application/json'
    }
    catch {
        if (( $r.ErrorDetails.message | convertfrom-json ).error.code -ne 9999) {
        
        } else {
            throw
        }
    }
    $script:notificationsSent++
    
    # Check flood warning (api global)
    $xLimitRemainning = $resp.headers['X-RateLimit-Remaining']
    $xFloodControlReset = $resp.headers['X-Ratelimit-Reset']
    Check-HipchatStandoff -xLimitRemaining $xLimitRemainning -xFloodControlReset $xFloodControlReset
    
    # Re-initialize timers
    if (!$script:hipchatNotifyThrottleStopWatch) {
        write-verbose "throttle: Initializing new stopwatch"
        $script:hipchatNotifyThrottleStopWatch = [Diagnostics.Stopwatch]::StartNew()
    }
    if (!$script:hipchatNotifySpamThrottleStopWatch) {
        write-verbose "spam: Initializing new spam stopwatch"
        $script:hipchatNotifySpamThrottleStopWatch = [Diagnostics.Stopwatch]::StartNew()
    }
}

#########################################################
# main
#########################################################
write-host "fogbugz query: $fbQuery" -f cyan
$unassignedUhOhCases = New-Search -q $fbQuery -cols @('sTitle',
                                                      'sOriginalTitle',
                                                      'sLatestTextSummary',
                                                      'sProject',
                                                      'ixPersonOpenedBy',
                                                      'ixPriority',
                                                      'ixPersonAssignedTo',
                                                      'dtOpened')
$caseCount = ($unassignedUhOhCases | Measure).Count
write-host "$caseCount unassigned, important cases" -f yellow

foreach ($case in $unassignedUhOhCases) {
    $now = get-date
    $dtOpened = (ConvertISO8601To-DateTime $case.dtOpened)
    $diffStr = (Format-TimeSpan -timespan (new-timespan $now $dtOpened))
    
    $messageStr = "<em>Unassigned Priority $($case.ixPriority) CS Case!</em>"
    $messageStr += " <a href='http://bugs.cardinal-ip.com/default.asp?$($case.ixBug)'>#$($case.ixBug)</a>"
    $messageStr += " in $($case.sProject):" 
    $messageStr += " <strong>$($case.sTitle)</strong>"
    $personOpenedBy = Get-Person $case.ixPersonOpenedBy
    $messageStr += " opened by <a href='mailto:$($personOpenedBy.sEmail)'>$($personOpenedBy.sFullName)</a>"
    $messageStr += " ($diffStr)"
    write-host $messageStr -f green
    
    $hipchatroominfo = $null
    $hipchatroominfo = $projectRoomMapping[$case.sProject]
    if (!$hipchatroominfo) {
        write-verbose "No designated room for project $($case.sProject). Using default room."
        $hipchatroominfo = $hipchatroomDefault
    }
    
    $hipchatExtraParams = @{
        color = 'purple'
        messageFormat = 'html'
    }
    if ($hipchatroominfo.apitoken) {
        $hipchatExtraParams.Add('apitoken',$hipchatroominfo.apitoken)
    }
    write-host "Sending message to roomid: $($hipchatroominfo.roomid)" -f yellow
    if (!$doNotNotifyChatrooms) {
        $resp = Send-HipchatNotification -notify -message $messageStr -roomid $hipchatroominfo.roomid @hipchatExtraParams
    } else {
        write-host "-doNotNotifyChatrooms : not actually sending" -f yellow
    }
}
         
# end
Invoke-FbLogoff