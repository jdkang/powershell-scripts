param(
    # Mailboxes to check
    [string[]]$mailboxen=@('mailbox@contoso.com'),
    # EWS/Exchange Credentials
    [System.Management.Automation.PSCredential]$creds,
    # Category (pre-made via Outlook) that the script can mark the read messages as
    [string]$markCategoryAs = "AUTOMAGIC",
    # Impersonate the target mailboxes using the provided proxy account.
    [switch]$Impersonate=$true, # REQ EWS impersonation rights
	# Utilize exchange autodiscover for CAS servers/etc. Does not work in limited security context, i.e. scheduled task that does not use the windows cred store.
	[switch]$useAutodiscover=$false,
	# Use the windows token the script was run as. Does not work in limited security context, i.e. scheduled task taht do not use the windows cred store.
	[switch]$UseDefaultCredentials=$false,
	# Skips inital checks on security tokens and CAS servers. REQUIRED FOR limited security context, i.e. scheduled task taht do not use the windows cred store.
	[switch]$skipAuthCheck=$true,
    
    <# Hard-coded stuff -- mostly for limited security context #>
    
	# Hard-coded EWS Proxy account name. REQUIRED FOR limited security context, i.e. scheduled task that do not use the windows cred store.
	[string]$exchangeUser='CONTOSO\wat.usr',
	# Hard-coded EWS Proxy password. REQUIRED FOR limited security context, i.e. scheduled task taht do not use the windows cred store.
	[string]$exchangePlaintextPassword='qqqqqqqqqqqqqqqqq',
	# Hard-coded CAS server list. REQUIRED FOR limited security context, i.e. scheduled task taht do not use the windows cred store.
	[string[]]$casServers = @('cas1.contoso.local','cas2.contoso.local'),
    
    <# msc #>

	# Temporary attachment directory.
	[string]$attachmentSaveDirectory = "E:\tmpEWS",
	# Exchange Web Services (EWS) DLL location.
	[string]$ewsDll = 'C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll',
	# Stand off interval before querying EWS again (when using a loop)
	[int]$waitCheckInboxAgainSec = 15
)
##########################################################################
# func
##########################################################################
function Process-EmailAttachments {
param(
	$attachments
)

	$savedAttachments = @()
	foreach ($attachment in $attachments) {
		write-debug "Loading attachment '$($attachment.Name)' ..."
		$attachment.load() # Load Content
		if (!$debugDontForwardOrMark) {
			if ($attachment.ContentType -eq "message/rfc822") {
				# some files require an additional MIME type load
				write-debug "message/rfc822 MIME type detected"
				write-debug "Loading MIME content."
				$attachment.Load($mimePropertySet)
				$attachmentData = $attachment.Item.MimeContent.Content
				$saveFileName = ($attachmentSaveDirectory + “\” + (Get-Date -Format "yyMMddHHmmss") + "_MSG.txt")
			} else {
				# Regular attachments
				$attachmentData = $attachment.Content
				$saveFileName = ($attachmentSaveDirectory + “\” + (Get-Date -Format "yyMMddHHmmss") + "_" + $attachment.Name.ToString())
			}
			
			# save file to disk
			write-debug "Saving file $saveFileName ..."
			$saveFileSizeMB = $attachmentData.Length / 1024 / 1024
			$attachmentFile = new-object System.IO.FileStream($saveFileName, [System.IO.FileMode]::Create)
			write-debug "writing content buffer [$saveFileSizeMB MB] ..." 
			$attachmentFile.Write($attachmentData, 0, $attachmentData.Length)
			write-debug "Closing file ..."
			$attachmentFile.Close()
		
			$savedAttachments += $saveFileName
		} 
		
		if ($debugDontForwardOrMark) {
			write-debug ($attachment | fl | out-string)
		}
	}
	
	return $savedAttachments
}
function Check-ForDistroGroupMembership {
param(
	[Microsoft.Exchange.WebServices.Data.ExchangeService]$ewsService,
	[string]$potentialDistroGroup,
	[string]$memberAddress
)

	write-debug "Resolving name for $potentialDistroGroup"
	$mailboxItem = $ewsService.ResolveName($potentialDistroGroup)
	if ($mailboxItem) {
		write-debug "Checking if Exch Distro List"
		if ($mailboxItem.Mailbox.MailboxType -eq 'PublicGroup') {
			$distroListAddresses = $ewsService.ExpandGroup($potentialDistroGroup)
			if (!$? -or !$distroListAddresses) {
				write-warning "Could to expand distro list."
				return $write-debug
			}
			foreach ($address in $distroListAddresses) {
				if ($address.address -eq $memberAddress) {
					write-debug "$($address.address) found in $potentialDistroGroup"
					return $true
				}
			}
		} else {
			write-warning "Not a distro list."
			return $write-debug
		}
	} else {
		write-warning "Could not resolve a mailbox item for $potentialDistroGroup"
		return $write-debug
	}	
}
##########################################################################
function Process-MailboxItems {
param(
	[string]$mailbox,
	[Microsoft.Exchange.WebServices.Data.ExchangeService]$ewsService,
	[Microsoft.Exchange.WebServices.Data.PropertySet]$propertySet,
	[Microsoft.Exchange.WebServices.Data.ItemView]$ivItemView,
	[Microsoft.Exchange.WebServices.Data.PropertySet]$mimePropertySet,
	[Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo]$ewsServiceearchFilter,
    [string]$markCategoryAs,
    [switch]$debugShowReturnFindItems
)
	#Mailbox objects, id
	write-debug "Fetching $mailbox folderID for Inbox"
	$mb = new-object Microsoft.Exchange.WebServices.Data.Mailbox($mailbox)	
	$folderInbox = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
	$folderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($folderInbox,$mb)
	write-debug "FolderID = $folderId"

	write-debug "Checking mailbox $mailbox"
	$findResults = $null
	$colour = $null
	# If you just wanted YOUR (cred) mailbox you wouldn't need to fetch the folderID
	#$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
	
	# User the folderID of a mailbox folder the provided creds have access to.
	$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsService,$folderId)

    if ($inbox.UnreadCount -gt 0) {
		write-debug "UNREAD COUNT: $($inbox.UnreadCount)"
		$findResults = $ewsService.FindItems($folderId,$ewsServiceearchFilter,$ivItemView)
	
		# Load extended as a batch (same as GetItem)
		write-debug "Loading extended data."
		$ewsService.LoadPropertiesForItems($findResults,$propertySet) | out-null
		
		# DEBUG
		if ($debugShowReturnFindItems) {
			write-debug "DEBUG: showing finditems objects after LoadPropertiesForItems"
			$findResults | Select *
			$GLOBAL:debugfindresults = $findResults
			write-host 'results saved to $GLOBAL:debugfindresults'
			return
		}
		
		foreach ($item in $findResults.Items) {
			# Load Data
			$messageFrom = ($item | Select -expand From).Address
			$messageTo = ($item | Select -expand ToRecipients | Select -expand Address)
			$messageBody = ($item | Select -expand Body).Text -replace "`r`n`r`n", "`r`n"
			[string[]]$messageCc += ($item | Select -expand CcRecipients | Select -expand Address)
			$messageSubject = ""
			if ($item.Subject) {
				$messageSubject = ($item | Select -expand Subject)
			}
			write-debug "------------------------------------------------------------------------"
			write-debug "SUB: $messageSubject"
			write-debug "FROM: $messageFrom" 
			write-debug "TO: $messageTo"
			write-debug "CC: $messageCc"
			write-debug "------------------------------------------------------------------------"
            $attachments = $null
            $attachmentCount = ($item.Attachments | Measure).Count
            write-debug "Message has $attachmentCount attachments."
            if ($attachmentCount -gt 0) {
                $attachments = Process-EmailAttachments -attachments $item.Attachments
            }
            
           
           # Mark with a category (must exist already)
            if ($markCategoryAs) {
                if (!($item.Categories.Contains($markCategoryAs))) {
                    write-debug "Marking message with Category $markCategoryAs"
                    $item.Categories.Add($markCategoryAs)
                    $madeItemChanges = $true
                }
            }
			
            # Mark as read
            $item.isread = $true
            
            # Update the message
            $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
		}
	} else { # ! (unreadcount gt 0)
		write-debug "Mo new messages."
	} 
	
	# Clear out the attachment directory
	write-debug "Clearing out temporary attachment cache $attachmentSaveDirectory"
	gci $attachmentSaveDirectory | rm -force | out-null
}
##########################################################################
# init
##########################################################################
# pwd
$currentScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
Push-Location $currentScriptPath

# create tmp attachment dir
write-debug "Checking for attachment tmp dir $attachmentSaveDirectory"
if (!(Test-Path $attachmentSaveDirectory)) {
	md $attachmentSaveDirectory -force | out-write-debug
}

# basic cred parameter check
if ( (!$creds) -and (!$exchangeUser) -and (!$UseDefaultCredentials) ) {
	write-warning "No authentication mechanism specified."
	exit
}

if ($exchangeUser -and (!$exchangePlaintextPassword)) {
	write-warning "Password required if specifying a hard-coded exchange user."
	exit
}

# Support for hard coded creds (ick)
# was necessary for scheduled task s4u support
if ($exchangeUser -and $exchangePlaintextPassword) {
	write-debug "Creating credential from supplied plaintext user $exchangeUser"
	$creds = (New-Object System.Management.Automation.PSCredential $exchangeUser,(convertto-securestring $exchangePlaintextPassword -asplaintext -force))
}

if ($UseDefaultCredentials -and $creds) {
	write-warning "Choose either UseDefaultCredentials flag or give ipmersination creds."
	exit
}
write-debug "Loading EWS"
[Reflection.Assembly]::LoadFile($ewsDll) | out-write-debug

# ----- EWS Init ----
# EWS Service (e.g. EX2007)
write-debug "Creating EWS objects"
$ewsService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

# Ace/ADSI identity
<#
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
#>

# Credentials
if ($UseDefaultCredentials) {
	write-debug "Using default credentials."
	$ewsService.UseDefaultCredentials = $true
} else {
    # Prompt if hard-coded creds weren't supplied
	if (!$creds) {
		write-debug "Supply impersination credentials."
		$creds = Get-Credential
	}
    
	if (!$skipAuthCheck) {
		# verify creds
		write-debug "Validating credentials."
		$CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
		$domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$creds.username,$creds.GetNetworkCredential().password)

		if ($domain.name -eq $write-debug)
		{
			write-warning "Authentication failed - please verify your username and password."
			exit
		} else {
			write-debug "Supplied Creds OK"
			$ewsService.Credentials = $creds.GetNetworkCredential() # Convert PsCredential to NetworkCredential
		}
	} else {
		write-debug "Skipping cred validation. Setting EWS Service credentials anyways."
		$ewsService.Credentials = $creds.GetNetworkCredential() # Convert PsCredential to NetworkCredential
	}
	
}

# AutoDiscover
$autodiscoverEmail = $mailboxen[0] # Pick an email to use for autodiscover
if ($useAutodiscover) {
	write-debug "Autodiscovering based on $autodiscoverEmail"
	$ewsService.AutodiscoverUrl($autodiscoverEmail)
	if (!$?) {
		write-warning "Error auto-discovering."
		exit
	}
} else {
	# "Hard coded" logic.
	write-debug "Checking CAS servers from hard coded list."
	$okCasServer = ""
	foreach ($casServer in $casServers) {
		$statusCode = (Invoke-WebRequest -usebasicparsing -uri "http://$casServer").StatusCode
		if ($statusCode -eq 200) {
			write-debug "$casServer CAS returned 200"
			$okCasServer = $casServer
			break
		}
	}
	if ($okCasServer) {
		write-debug "Setting service URL based on online CAS."
		$ewsService.Url = "http://$okCasServer/EWS/Exchange.asmx"
	} else {
		write-debug "Could not find an online CAS server."
	}
}

# Check for URL
if (!$ewsService.Url) {
	write-warning "EWS Service Url is blank"
	exit
}

write-debug "Service URL = $($ewsService.Url)"

write-debug "Setting up views, filters, etc..."
# API: http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data%28v=exchg.80%29.aspx
# e.g. http://stackoverflow.com/questions/1614720/using-ews-api-to-search-through-different-users-mailboxes

# views and filters
# http://gsexdev.blogspot.com/2012/02/ews-managed-api-and-powershell-how-to.html
# http://gsexdev.blogspot.com/2009/04/using-ews-managed-api-with-powershell.html
# http://social.msdn.microsoft.com/Forums/exchange/en-US/a4811684-054f-497c-ab4b-04dc03f10188/ews-searchfilterisequalto-in-powershell?forum=exchangesvrdevelopment
# http://stackoverflow.com/questions/11243911/ews-body-plain-text

$propertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
# Plain text
$propertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
<#
# headers - http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.internetmessageheader%28v=exchg.80%29.aspx
$PR_TRANSPORT_MSG_HEADERS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x007D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$propertySet.Add($PR_TRANSPORT_MSG_HEADERS)
#>

$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(999) 
$ivItemView.PropertySet = $propertySet # Have to do this for the view and the item loads

# For embedded .MSGs
$mimePropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)

#  new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
$ewsServiceearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)

# more complex collection filter
<#
$Sfsub = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $Subject[0])
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfsub)
$sfCollection.add($Sfha)
#>

# EWS Splat
$ewsCommonParameters = @{
		'ewsService' = $ewsService
		'propertySet' = $propertySet
		'ivItemView' = $ivItemView
		'mimePropertySet' = $mimePropertySet
		'ewsServiceearchFilter' = $ewsServiceearchFilter
}

##########################################################################
# main
##########################################################################
while (1) {
    foreach ($mail in $mailboxen) {
        Ticket-MailboxItems -mailbox $mailbox @ewsCommonParameters
    }
    
    start-sleep -seconds $waitCheckInboxAgainSec
}

