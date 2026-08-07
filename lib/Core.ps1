<#
.SYNOPSIS
    Core Exchange content search functions for EXRESearcher.
    Search-Mailbox, ComplianceSearch, mailbox enumeration, bulk operations.
#>

# ═══════════════════════════════════════════════════════════════════════════════
# CONNECTION
# ═══════════════════════════════════════════════════════════════════════════════

function Test-ExchangeManagementShell {
    <#
    .SYNOPSIS
        Check if running inside Exchange Management Shell (EMS).
        Returns $true if Exchange snap-in/module is already loaded.
    #>
    [CmdletBinding()]
    param()

    # Check for Exchange snap-in (on-prem EMS)
    $snap = Get-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.*' -ErrorAction SilentlyContinue
    if ($snap) { return $true }

    # Check if key Exchange cmdlets are already available
    $cmd = Get-Command -Name 'Get-ExchangeServer' -ErrorAction SilentlyContinue
    if ($cmd) { return $true }

    return $false
}

function Find-ExchangeServers {
    <#
    .SYNOPSIS
        Auto-discover Exchange servers in the organization.
        Returns array of server objects with Name, FQDN, Role, Version.
        Requires EMS or an active Exchange remote session.
    #>
    [CmdletBinding()]
    param()

    $servers = @(Get-ExchangeServer -ErrorAction Stop | Where-Object {
        $_.ServerRole -match 'Mailbox'
    } | ForEach-Object {
        [PSCustomObject]@{
            Name    = $_.Name
            FQDN    = $_.Fqdn
            Role    = "$($_.ServerRole)"
            Version = "$($_.AdminDisplayVersion)"
            Site    = "$($_.Site)"
        }
    })
    return $servers
}

function Connect-ExchangeSearch {
    <#
    .SYNOPSIS
        Connect to Exchange via remote PowerShell (Kerberos).
        Returns PSSession object or a marker hashtable if EMS is already loaded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server
    )

    # If running inside EMS, cmdlets are already available - no remote session needed
    if (Test-ExchangeManagementShell) {
        return @{ IsEMS = $true; Server = $Server }
    }

    $uri = "http://$Server/PowerShell/"
    $session = New-PSSession -ConfigurationName 'Microsoft.Exchange' `
                              -ConnectionUri $uri `
                              -Authentication 'Kerberos' `
                              -ErrorAction Stop

    Import-PSSession -Session $session -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
    return $session
}

function Disconnect-ExchangeSearch {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Session)
    # EMS sessions are local - nothing to disconnect
    if ($Session -is [hashtable] -and $Session.IsEMS) { return }
    try { Remove-PSSession -Session $Session -ErrorAction SilentlyContinue } catch {}
}

function Get-ExchangeServerVersion {
    <#
    .SYNOPSIS
        Get Exchange server version info.
    #>
    [CmdletBinding()]
    param([string]$Server)
    try {
        $exServer = Get-ExchangeServer -Identity $Server -ErrorAction Stop
        return [PSCustomObject]@{
            Name           = $exServer.Name
            Edition        = $exServer.Edition
            AdminVersion   = "$($exServer.AdminDisplayVersion)"
            ServerRole     = "$($exServer.ServerRole)"
            Site           = "$($exServer.Site)"
        }
    } catch {
        return [PSCustomObject]@{ Name = $Server; Edition = 'Unknown'; AdminVersion = 'N/A'; ServerRole = 'N/A'; Site = 'N/A' }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# EWS HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

function New-EwsNamespaceManager {
    <# Create XmlNamespaceManager for EWS SOAP responses #>
    param([System.Xml.XmlDocument]$Xml)
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    $nsMgr.AddNamespace('s', 'http://schemas.xmlsoap.org/soap/envelope/')
    $nsMgr.AddNamespace('m', 'http://schemas.microsoft.com/exchange/services/2006/messages')
    $nsMgr.AddNamespace('t', 'http://schemas.microsoft.com/exchange/services/2006/types')
    return $nsMgr
}

function Initialize-EwsCertPolicy {
    <#
    .SYNOPSIS
        Install a compiled certificate validation callback (a PowerShell
        scriptblock callback crashes on non-PS threads with "no Runspace").
        Valid certs always pass; invalid ones pass only while TrustAllCerts
        is enabled (default; set EXRE_STRICT_TLS=1 to enforce validation).
    #>
    if (-not ('EXRESearcher.CertPolicy' -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
namespace EXRESearcher {
    public static class CertPolicy {
        public static bool TrustAllCerts = true;
        public static void Install() {
            ServicePointManager.ServerCertificateValidationCallback = Validate;
        }
        private static bool Validate(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors) {
            if (errors == SslPolicyErrors.None) { return true; }
            return TrustAllCerts;
        }
    }
}
"@
    }
    [EXRESearcher.CertPolicy]::TrustAllCerts = ($env:EXRE_STRICT_TLS -ne '1')
    [EXRESearcher.CertPolicy]::Install()
}

# Well-known folder paths -> EWS distinguished folder IDs
$script:EwsWellKnownFolders = @{
    'Inbox'          = 'inbox'
    'Sent Items'     = 'sentitems'
    'Drafts'         = 'drafts'
    'Deleted Items'  = 'deleteditems'
    'Junk Email'     = 'junkemail'
    'Outbox'         = 'outbox'
    'Notes'          = 'notes'
    'Calendar'       = 'calendar'
    'Contacts'       = 'contacts'
    'Tasks'          = 'tasks'
}

function Invoke-EwsRequest {
    <#
    .SYNOPSIS
        Send EWS SOAP request.
        With -Mailbox: tries a plain request first (the SOAP body must reference
        the target mailbox explicitly - Full Access path), then Exchange
        impersonation. Without -Mailbox: single plain request (caller's own
        mailbox). Never silently retargets another mailbox.
        Returns [xml] response.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EwsUrl,
        [Parameter(Mandatory)][string]$SoapBody,
        [string]$Mailbox
    )

    try { Initialize-EwsCertPolicy } catch {}
    $headers = @{ 'Content-Type' = 'text/xml; charset=utf-8' }

    $escapedMailbox = if ($Mailbox) { [System.Security.SecurityElement]::Escape($Mailbox) } else { '' }

    # Plain request (works for own mailbox and Full Access on target mailbox)
    $soapPlain = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
$SoapBody
  </soap:Body>
</soap:Envelope>
"@

    # Impersonation request (requires ApplicationImpersonation role)
    $soapImpersonate = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>$escapedMailbox</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
$SoapBody
  </soap:Body>
</soap:Envelope>
"@

    $variants = if ($Mailbox) { @($soapPlain, $soapImpersonate) } else { @($soapPlain) }
    $lastError = ''

    foreach ($soap in $variants) {
        try {
            $response = Invoke-WebRequest -Uri $EwsUrl -Method POST -Body $soap -Headers $headers `
                            -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop
            [xml]$xml = $response.Content
            $nsMgr = New-EwsNamespaceManager -Xml $xml

            # Check for SOAP fault
            $fault = $xml.SelectSingleNode('//s:Fault/faultstring', $nsMgr)
            if ($fault) { $lastError = $fault.InnerText; continue }

            # Multi-folder requests return one response message per folder;
            # fail over only when there is no successful response at all.
            $respMsg = $xml.SelectNodes('//*[@ResponseClass]', $nsMgr)
            $successCount = 0
            $errorText = ''
            foreach ($r in $respMsg) {
                if ($r.ResponseClass -eq 'Error') {
                    if (-not $errorText) { $errorText = $r.MessageText }
                } else {
                    $successCount++
                }
            }
            if ($errorText -and $successCount -eq 0) { $lastError = $errorText; continue }

            return $xml
        } catch {
            $lastError = "$_"
        }
    }

    throw "EWS request failed: $lastError"
}

function Get-EwsMailboxFolderIds {
    <#
    .SYNOPSIS
        Enumerate all folder IDs of a mailbox (FindFolder Deep from msgfolderroot).
        FindItem does not support Deep traversal, so whole-mailbox item searches
        must pass an explicit folder list as ParentFolderIds.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$Server,
        [int]$MaxFolders = 1000
    )

    $ewsUrl = "https://$Server/EWS/Exchange.asmx"
    $escapedMailbox = [System.Security.SecurityElement]::Escape($Mailbox)

    $soapBody = @"
    <m:FindFolder Traversal="Deep">
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:IndexedPageFolderView MaxEntriesReturned="$MaxFolders" Offset="0" BasePoint="Beginning" />
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot">
          <t:Mailbox><t:EmailAddress>$escapedMailbox</t:EmailAddress></t:Mailbox>
        </t:DistinguishedFolderId>
      </m:ParentFolderIds>
    </m:FindFolder>
"@

    $xml = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $soapBody -Mailbox $Mailbox
    $ns = New-EwsNamespaceManager -Xml $xml

    $ids = @()
    foreach ($node in $xml.SelectNodes('//t:Folder/t:FolderId', $ns)) {
        $ids += $node.GetAttribute('Id')
    }
    return $ids
}

function Get-EwsParentFolderXml {
    <#
    .SYNOPSIS
        Build ParentFolderIds inner XML covering the whole mailbox.
        Falls back to well-known folders if folder enumeration fails.
    #>
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$Server
    )
    $escapedMailbox = [System.Security.SecurityElement]::Escape($Mailbox)
    $folderIds = @()
    try { $folderIds = @(Get-EwsMailboxFolderIds -Mailbox $Mailbox -Server $Server) } catch {}
    if ($folderIds.Count -gt 0) {
        return ($folderIds | ForEach-Object { "<t:FolderId Id=`"$_`" />" }) -join "`n        "
    }
    return (@('inbox', 'sentitems', 'deleteditems', 'junkemail', 'drafts') | ForEach-Object {
        "<t:DistinguishedFolderId Id=`"$_`"><t:Mailbox><t:EmailAddress>$escapedMailbox</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId>"
    }) -join "`n        "
}

function Get-EwsFolderXml {
    <#
    .SYNOPSIS
        Resolve a folder path to ParentFolderIds inner XML for a single folder.
        Throws if the folder cannot be found.
    #>
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$Server
    )
    $escapedMailbox = [System.Security.SecurityElement]::Escape($Mailbox)
    $folderClean = $FolderPath.TrimStart('/').Trim()
    $distinguishedId = $script:EwsWellKnownFolders[$folderClean]
    if ($distinguishedId) {
        return @"
<t:DistinguishedFolderId Id="$distinguishedId">
          <t:Mailbox><t:EmailAddress>$escapedMailbox</t:EmailAddress></t:Mailbox>
        </t:DistinguishedFolderId>
"@
    }
    $folderId = Find-EwsFolderId -Mailbox $Mailbox -FolderPath $folderClean -Server $Server
    if (-not $folderId) {
        throw "Folder '$folderClean' not found in mailbox $Mailbox"
    }
    return "<t:FolderId Id=`"$folderId`" />"
}

function Get-EwsFolderMessages {
    <#
    .SYNOPSIS
        Paged EWS FindItem over the given parent folder(s).
        Returns @{ Items = <Id/ChangeKey/Subject/From/Received/Size objects>; Total = <int> }.
        RestrictionXml and QueryStringXml are mutually exclusive (EWS limitation).
        Paging works only for a single parent folder; multi-folder requests
        return the first page per folder (EWS pages per-folder).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$ParentFolderXml,
        [string]$RestrictionXml = '',
        [string]$QueryStringXml = '',
        [int]$MaxItems = 50000
    )

    $ewsUrl = "https://$Server/EWS/Exchange.asmx"
    $items = @()
    $total = 0
    $offset = 0

    while ($true) {
        $pageSize = [Math]::Min(1000, $MaxItems - $items.Count)
        if ($pageSize -le 0) { break }

        $soapBody = @"
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:Size" />
          <t:FieldURI FieldURI="message:From" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="$pageSize" Offset="$offset" BasePoint="Beginning" />
      $RestrictionXml
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived" />
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        $ParentFolderXml
      </m:ParentFolderIds>
      $QueryStringXml
    </m:FindItem>
"@

        $xml = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $soapBody -Mailbox $Mailbox
        $ns = New-EwsNamespaceManager -Xml $xml

        $roots = $xml.SelectNodes('//m:RootFolder', $ns)
        if ($roots.Count -eq 0) { break }

        if ($offset -eq 0) {
            foreach ($root in $roots) {
                $t = 0
                if ([int]::TryParse($root.GetAttribute('TotalItemsInView'), [ref]$t)) { $total += $t }
            }
        }

        $pageNodes = $xml.SelectNodes('//t:Items/*', $ns)
        foreach ($node in $pageNodes) {
            $idNode = $node.SelectSingleNode('t:ItemId', $ns)
            if (-not $idNode) { continue }
            $sz = 0
            $sizeNode = $node.SelectSingleNode('t:Size', $ns)
            if ($sizeNode) { [int]::TryParse($sizeNode.InnerText, [ref]$sz) | Out-Null }
            $subjNode = $node.SelectSingleNode('t:Subject', $ns)
            $recvNode = $node.SelectSingleNode('t:DateTimeReceived', $ns)
            $fromNode = $node.SelectSingleNode('t:From/t:Mailbox/t:EmailAddress', $ns)
            if (-not $fromNode) { $fromNode = $node.SelectSingleNode('t:From/t:Mailbox/t:Name', $ns) }
            $items += [PSCustomObject]@{
                Id        = $idNode.GetAttribute('Id')
                ChangeKey = $idNode.GetAttribute('ChangeKey')
                Subject   = if ($subjNode) { $subjNode.InnerText } else { '' }
                From      = if ($fromNode) { $fromNode.InnerText } else { '' }
                Received  = if ($recvNode) { $recvNode.InnerText } else { '' }
                Size      = $sz
            }
        }

        # Stop conditions: all folders exhausted, multi-folder (no shared offset), empty page
        $allLast = $true
        foreach ($root in $roots) {
            if ($root.GetAttribute('IncludesLastItemInRange') -ne 'true') { $allLast = $false }
        }
        if ($allLast -or $roots.Count -gt 1 -or $pageNodes.Count -eq 0) { break }

        $newOffset = 0
        [int]::TryParse($roots[0].GetAttribute('IndexedPagingOffset'), [ref]$newOffset) | Out-Null
        if ($newOffset -le $offset) { break }
        $offset = $newOffset
    }

    return @{ Items = $items; Total = [Math]::Max($total, $items.Count) }
}

# ═══════════════════════════════════════════════════════════════════════════════
# EWS MESSAGE PREVIEW
# ═══════════════════════════════════════════════════════════════════════════════

function Get-MailboxMessagePreview {
    <#
    .SYNOPSIS
        Retrieve individual messages matching a KQL query via EWS FindItem.
        Tries without impersonation first (Full Access), then with impersonation.
        Returns array of message objects with Subject, From, To, Received, Size.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$SearchQuery,
        [string]$Server,
        [int]$MaxResults = 200
    )

    if (-not $Server) {
        $Server = (Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' } | Select-Object -First 1).Fqdn
    }

    $ewsUrl = "https://$Server/EWS/Exchange.asmx"
    $escapedQuery = [System.Security.SecurityElement]::Escape($SearchQuery)

    # FindItem cannot traverse Deep; search all folders explicitly
    # (Shallow on msgfolderroot alone would only see items directly in the root)
    $parentFolderXml = Get-EwsParentFolderXml -Mailbox $Mailbox -Server $Server

    $soapBody = @"
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:Size" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="message:ToRecipients" />
          <t:FieldURI FieldURI="item:HasAttachments" />
          <t:FieldURI FieldURI="item:ItemClass" />
          <t:FieldURI FieldURI="item:Importance" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="$MaxResults" Offset="0" BasePoint="Beginning" />
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived" />
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        $parentFolderXml
      </m:ParentFolderIds>
      <m:QueryString>$escapedQuery</m:QueryString>
    </m:FindItem>
"@

    $xml = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $soapBody -Mailbox $Mailbox
    $ns = New-EwsNamespaceManager -Xml $xml

    $items = $xml.SelectNodes('//t:Message', $ns)
    $results = @()

    foreach ($item in $items) {
        $fromName = $item.SelectSingleNode('t:From/t:Mailbox/t:Name', $ns)
        $fromAddr = $item.SelectSingleNode('t:From/t:Mailbox/t:EmailAddress', $ns)
        $toNodes  = $item.SelectNodes('t:ToRecipients/t:Mailbox', $ns)
        $toList   = @($toNodes | ForEach-Object {
            $n = $_.SelectSingleNode('t:Name', $ns)
            $e = $_.SelectSingleNode('t:EmailAddress', $ns)
            if ($n) { $n.InnerText } elseif ($e) { $e.InnerText } else { '' }
        }) -join '; '

        $sizeBytes = 0
        $sizeNode = $item.SelectSingleNode('t:Size', $ns)
        if ($sizeNode) { [int]::TryParse($sizeNode.InnerText, [ref]$sizeBytes) | Out-Null }
        $sizeKB = [math]::Round($sizeBytes / 1024, 1)

        $subjectNode  = $item.SelectSingleNode('t:Subject', $ns)
        $receivedNode = $item.SelectSingleNode('t:DateTimeReceived', $ns)
        $attachNode   = $item.SelectSingleNode('t:HasAttachments', $ns)
        $importNode   = $item.SelectSingleNode('t:Importance', $ns)
        $classNode    = $item.SelectSingleNode('t:ItemClass', $ns)

        $results += [PSCustomObject]@{
            Subject     = if ($subjectNode)  { $subjectNode.InnerText }  else { '(no subject)' }
            From        = if ($fromName) { $fromName.InnerText } elseif ($fromAddr) { $fromAddr.InnerText } else { '' }
            To          = $toList
            Received    = if ($receivedNode) { $receivedNode.InnerText } else { '' }
            SizeKB      = $sizeKB
            HasAttach   = if ($attachNode)   { $attachNode.InnerText }   else { 'false' }
            Importance  = if ($importNode)   { $importNode.InnerText }   else { 'Normal' }
            ItemClass   = if ($classNode)    { $classNode.InnerText }    else { 'IPM.Note' }
        }
    }

    # MaxEntriesReturned applies per folder; cap the combined result set
    return @($results | Sort-Object Received -Descending | Select-Object -First $MaxResults)
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAILBOX ENUMERATION
# ═══════════════════════════════════════════════════════════════════════════════

function Get-SearchableMailboxes {
    <#
    .SYNOPSIS
        Get list of mailboxes for search scope selection.
    #>
    [CmdletBinding()]
    param(
        [string]$Filter,
        [string]$Database,
        [string]$OrganizationalUnit,
        [ValidateSet('All','UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','DiscoveryMailbox')]
        [string]$RecipientType = 'All',
        [int]$ResultSize = 500
    )

    $params = @{ ResultSize = $ResultSize; ErrorAction = 'Stop' }

    if ($Filter) {
        $params['Filter'] = $Filter
    }
    if ($Database) {
        $params['Database'] = $Database
    }
    if ($OrganizationalUnit) {
        $params['OrganizationalUnit'] = $OrganizationalUnit
    }

    $mailboxes = Get-Mailbox @params

    if ($RecipientType -ne 'All') {
        $mailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq $RecipientType }
    }

    return $mailboxes | ForEach-Object {
        [PSCustomObject]@{
            DisplayName       = $_.DisplayName
            PrimarySmtp       = $_.PrimarySmtpAddress
            Alias             = $_.Alias
            Database          = "$($_.Database)"
            RecipientType     = "$($_.RecipientTypeDetails)"
            OrganizationalUnit = "$($_.OrganizationalUnit)"
            ItemCount         = ''
            TotalSize         = ''
        }
    } | Sort-Object DisplayName
}

function Get-MailboxDatabases {
    <#
    .SYNOPSIS
        Get list of mailbox databases for scope selection.
    #>
    [CmdletBinding()]
    param()
    try {
        return Get-MailboxDatabase -ErrorAction Stop | ForEach-Object {
            [PSCustomObject]@{
                Name           = $_.Name
                Server         = "$($_.Server)"
                MailboxCount   = "$($_.DatabaseSize)"
                EdbPath        = "$($_.EdbFilePath)"
                Mounted        = $_.Mounted
            }
        } | Sort-Object Name
    } catch {
        return @()
    }
}

function Get-DistributionGroupMembers {
    <#
    .SYNOPSIS
        Get members of a distribution group for search scope.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Identity)
    try {
        $members = Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop
        return $members | Where-Object { $_.RecipientType -match 'Mailbox' } | ForEach-Object {
            $_.PrimarySmtpAddress
        }
    } catch {
        return @()
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# EWS MESSAGE-ID SEARCH
# ═══════════════════════════════════════════════════════════════════════════════

function Find-MessageByMessageId {
    <#
    .SYNOPSIS
        Find messages by Internet Message-ID header via EWS FindItem with Restriction.
        Search-Mailbox does not support messageid: keyword — EWS is required.
        Returns search result objects compatible with Invoke-MailboxSearch output.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$MessageId,
        [ValidateSet('Estimate','DeleteContent')]
        [string]$Action = 'Estimate',
        [string]$Server
    )

    if (-not $Server) {
        $Server = (Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' } | Select-Object -First 1).Fqdn
    }

    $ewsUrl = "https://$Server/EWS/Exchange.asmx"

    # Normalize MessageId - ensure angle brackets
    $msgId = $MessageId.Trim()
    if (-not $msgId.StartsWith('<')) { $msgId = "<$msgId" }
    if (-not $msgId.EndsWith('>'))   { $msgId = "$msgId>" }
    $escapedMsgId = [System.Security.SecurityElement]::Escape($msgId)

    # Search every folder: FindItem has no Deep traversal
    $parentFolderXml = Get-EwsParentFolderXml -Mailbox $Mailbox -Server $Server
    $restrictionXml = @"
      <m:Restriction>
        <t:IsEqualTo>
          <t:FieldURI FieldURI="message:InternetMessageId" />
          <t:FieldURIOrConstant>
            <t:Constant Value="$escapedMsgId" />
          </t:FieldURIOrConstant>
        </t:IsEqualTo>
      </m:Restriction>
"@

    $found = Get-EwsFolderMessages -Mailbox $Mailbox -Server $Server `
                -ParentFolderXml $parentFolderXml -RestrictionXml $restrictionXml -MaxItems 100
    $itemIds = @($found.Items)
    $totalSize = ($itemIds | Measure-Object -Property Size -Sum).Sum
    if (-not $totalSize) { $totalSize = 0 }

    $deleted = 0
    if ($Action -eq 'DeleteContent' -and $itemIds.Count -gt 0) {
        foreach ($itemRef in $itemIds) {
            try {
                $deleteBody = @"
    <m:DeleteItem DeleteType="SoftDelete" AffectedTaskOccurrences="AllOccurrences">
      <m:ItemIds>
        <t:ItemId Id="$($itemRef.Id)" ChangeKey="$($itemRef.ChangeKey)" />
      </m:ItemIds>
    </m:DeleteItem>
"@
                $null = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $deleteBody -Mailbox $Mailbox
                $deleted++
            } catch {
                Write-Warning "Failed to delete item: $_"
            }
        }
    }

    $sizeStr = if ($totalSize -gt 1MB) { "$([math]::Round($totalSize/1MB, 2)) MB ($totalSize bytes)" }
               elseif ($totalSize -gt 1KB) { "$([math]::Round($totalSize/1KB, 1)) KB ($totalSize bytes)" }
               else { "$totalSize bytes" }

    return [PSCustomObject]@{
        Mailbox       = $Mailbox
        DisplayName   = ''
        Success       = $true
        ResultItems   = if ($Action -eq 'DeleteContent') { $deleted } else { $itemIds.Count }
        ResultSize    = $sizeStr
        Action        = $Action
        SearchQuery   = "messageid:`"$msgId`""
        TargetMailbox = ''
        TargetFolder  = ''
        Timestamp     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SEARCH-MAILBOX (Exchange 2019 SE native)
# ═══════════════════════════════════════════════════════════════════════════════

function ConvertTo-SearchMailboxQuery {
    <#
    .SYNOPSIS
        Strip KQL keywords not supported by Search-Mailbox (folder:, hasattachment:).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Query)

    $clean = $Query
    $clean = $clean -replace '\s*AND\s*folder:"[^"]*"', ''
    $clean = $clean -replace 'folder:"[^"]*"\s*(AND\s*)?', ''
    $clean = $clean -replace '\s*AND\s*hasattachment:\w+', ''
    $clean = $clean -replace 'hasattachment:\w+\s*(AND\s*)?', ''
    $clean = $clean.Trim()
    if (-not $clean) { return '*' }
    return $clean
}

function Invoke-MailboxSearch {
    <#
    .SYNOPSIS
        Search mailbox content using Search-Mailbox.
        Supports EstimateResultOnly, LogOnly, and DeleteContent.
    .PARAMETER SearchQuery
        KQL query string (e.g. 'subject:"invoice" AND from:"user@domain.com"')
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Mailboxes,
        [Parameter(Mandatory)][string]$SearchQuery,
        [ValidateSet('Estimate','LogOnly','CopyToFolder','DeleteContent')]
        [string]$Action = 'Estimate',
        [string]$TargetMailbox,
        [string]$TargetFolder = 'SearchResults',
        [switch]$Force
    )

    $cleanQuery = ConvertTo-SearchMailboxQuery -Query $SearchQuery

    $results = @()

    foreach ($mbx in $Mailboxes) {
        $params = @{
            Identity    = $mbx
            SearchQuery = $cleanQuery
            ErrorAction = 'Stop'
        }

        switch ($Action) {
            'Estimate' {
                $params['EstimateResultOnly'] = $true
            }
            'LogOnly' {
                if (-not $TargetMailbox) { throw "TargetMailbox required for LogOnly action" }
                $params['TargetMailbox'] = $TargetMailbox
                $params['TargetFolder'] = $TargetFolder
                $params['LogOnly'] = $true
            }
            'CopyToFolder' {
                if (-not $TargetMailbox) { throw "TargetMailbox required for CopyToFolder action" }
                $params['TargetMailbox'] = $TargetMailbox
                $params['TargetFolder'] = $TargetFolder
            }
            'DeleteContent' {
                $params['DeleteContent'] = $true
                if ($Force) {
                    $params['Force'] = $true
                }
            }
        }

        try {
            $searchResult = Search-Mailbox @params
            foreach ($r in $searchResult) {
                $results += [PSCustomObject]@{
                    Mailbox        = "$($r.Identity)"
                    DisplayName    = "$($r.DisplayName)"
                    Success        = $r.Success
                    ResultItems    = $r.ResultItemsCount
                    ResultSize     = "$($r.ResultItemsSize)"
                    Action         = $Action
                    SearchQuery    = $SearchQuery
                    TargetMailbox  = if ($Action -ne 'Estimate') { $TargetMailbox } else { '' }
                    TargetFolder   = if ($Action -ne 'Estimate') { $TargetFolder } else { '' }
                    Timestamp      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
            }
        } catch {
            $results += [PSCustomObject]@{
                Mailbox        = $mbx
                DisplayName    = ''
                Success        = $false
                ResultItems    = 0
                ResultSize     = ''
                Action         = $Action
                SearchQuery    = $SearchQuery
                TargetMailbox  = ''
                TargetFolder   = ''
                Timestamp      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Error          = "$_"
            }
        }
    }

    return $results
}

function Build-SearchQuery {
    <#
    .SYNOPSIS
        Build KQL search query from individual filter parameters.
        Supports folder scoping via folder:"path" KQL syntax.
    #>
    [CmdletBinding()]
    param(
        [string]$Subject,
        [string]$From,
        [string]$To,
        [string]$Keywords,
        [string]$AttachmentName,
        [string]$MessageId,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [ValidateSet('','IPM.Note','IPM.Appointment','IPM.Contact','IPM.Task')]
        [string]$MessageKind,
        [string]$Folder,
        [ValidateSet('','Small','Medium','Large','VeryLarge')]
        [string]$SizeRange,
        [switch]$HasAttachment
    )

    $parts = @()

    # Search-Mailbox supported KQL: subject, from, to, cc, bcc, participants,
    # body, attachment, sent, received, kind, size
    # NOT supported: messageid, folder, hasattachment
    if ($Subject)        { $parts += "subject:`"$Subject`"" }
    if ($From)           { $parts += "from:`"$From`"" }
    if ($To)             { $parts += "to:`"$To`"" }
    if ($Keywords)       { $parts += "$Keywords" }
    if ($AttachmentName) { $parts += "attachment:`"$AttachmentName`"" }
    if ($MessageId)      { $parts += "`"$MessageId`"" }  # free-text search (messageid: not supported)
    if ($MessageKind)    { $parts += "kind:$MessageKind" }
    if ($Folder)         { $parts += "folder:`"$Folder`"" }  # only for EWS, stripped for Search-Mailbox
    if ($HasAttachment)  { $parts += "hasattachment:true" }  # only for EWS, stripped for Search-Mailbox

    if ($SizeRange) {
        switch ($SizeRange) {
            'Small'     { $parts += "size<10KB" }
            'Medium'    { $parts += "size>=10KB AND size<1MB" }
            'Large'     { $parts += "size>=1MB AND size<10MB" }
            'VeryLarge' { $parts += "size>=10MB" }
        }
    }

    if ($StartDate) {
        $parts += "received>=$($StartDate.ToString('yyyy-MM-dd'))"
    }
    if ($EndDate) {
        $parts += "received<=$($EndDate.ToString('yyyy-MM-dd'))"
    }

    if ($parts.Count -eq 0) {
        return '*'
    }

    return ($parts -join ' AND ')
}

# ═══════════════════════════════════════════════════════════════════════════════
# COMPLIANCE SEARCH (Exchange 2019 SE - In-Place eDiscovery)
# ═══════════════════════════════════════════════════════════════════════════════

function New-ContentSearch {
    <#
    .SYNOPSIS
        Create and start a new compliance/content search.
        Uses New-MailboxSearch (Exchange 2019 SE In-Place eDiscovery).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$SearchQuery,
        [string[]]$SourceMailboxes,
        [switch]$AllMailboxes,
        [switch]$EstimateOnly,
        [string]$TargetMailbox,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [string[]]$Senders,
        [string[]]$Recipients
    )

    $params = @{
        Name         = $Name
        SearchQuery  = $SearchQuery
        ErrorAction  = 'Stop'
    }

    if ($AllMailboxes) {
        $params['AllSourceMailboxes'] = $true
    } elseif ($SourceMailboxes) {
        $params['SourceMailboxes'] = $SourceMailboxes
    }

    if ($EstimateOnly) {
        $params['EstimateOnly'] = $true
    }

    if ($TargetMailbox) {
        $params['TargetMailbox'] = $TargetMailbox
    }

    if ($StartDate) { $params['StartDate'] = $StartDate }
    if ($EndDate)   { $params['EndDate'] = $EndDate }
    if ($Senders)   { $params['Senders'] = $Senders }
    if ($Recipients) { $params['Recipients'] = $Recipients }

    $search = New-MailboxSearch @params
    Start-MailboxSearch -Identity $search.Name -ErrorAction Stop

    return [PSCustomObject]@{
        Name         = $search.Name
        SearchQuery  = $SearchQuery
        Status       = 'Started'
        CreatedBy    = $search.CreatedBy
        CreatedTime  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }
}

function Get-ContentSearches {
    <#
    .SYNOPSIS
        Get all existing mailbox searches with their status.
    #>
    [CmdletBinding()]
    param([string]$Name)

    $params = @{ ErrorAction = 'Stop' }
    if ($Name) { $params['Identity'] = $Name }

    try {
        $searches = Get-MailboxSearch @params
        return $searches | ForEach-Object {
            [PSCustomObject]@{
                Name               = $_.Name
                Status             = "$($_.Status)"
                SearchQuery        = "$($_.SearchQuery)"
                SourceMailboxes    = ($_.SourceMailboxes -join '; ')
                AllSourceMailboxes = $_.AllSourceMailboxes
                ResultItemCount    = "$($_.ResultItemCountEstimate)"
                ResultSize         = "$($_.ResultSizeEstimate)"
                TargetMailbox      = "$($_.TargetMailbox)"
                StartDate          = "$($_.StartDate)"
                EndDate            = "$($_.EndDate)"
                CreatedBy          = "$($_.CreatedBy)"
                LastModified       = "$($_.LastModifiedTime)"
                EstimateOnly       = $_.EstimateOnly
            }
        }
    } catch {
        return @()
    }
}

function Get-ContentSearchStatus {
    <#
    .SYNOPSIS
        Get detailed status of a specific search.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    try {
        $search = Get-MailboxSearch -Identity $Name -ErrorAction Stop
        $statusText = "$($search.Status)"

        return [PSCustomObject]@{
            Name             = $search.Name
            Status           = $statusText
            SearchQuery      = "$($search.SearchQuery)"
            ResultItems      = "$($search.ResultItemCountEstimate)"
            ResultSize       = "$($search.ResultSizeEstimate)"
            PercentComplete  = "$($search.PercentComplete)"
            Errors           = ($search.Errors -join '; ')
            SourceMailboxes  = ($search.SourceMailboxes -join '; ')
            AllMailboxes     = $search.AllSourceMailboxes
            LastStartTime    = "$($search.LastStartTime)"
            LastEndTime      = "$($search.LastEndTime)"
        }
    } catch {
        return [PSCustomObject]@{
            Name    = $Name
            Status  = 'Error'
            Errors  = "$_"
        }
    }
}

function Remove-ContentSearch {
    <#
    .SYNOPSIS
        Remove/delete a mailbox search.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    try {
        Stop-MailboxSearch -Identity $Name -ErrorAction SilentlyContinue
        Remove-MailboxSearch -Identity $Name -Confirm:$false -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Stop-ContentSearch {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    try {
        Stop-MailboxSearch -Identity $Name -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# ORGANIZATION-WIDE OPERATIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Remove-MessageFromOrganization {
    <#
    .SYNOPSIS
        Delete a specific message from ALL mailboxes in the organization.
        Typically used for phishing/malware cleanup.
    .PARAMETER SearchQuery
        KQL query to identify the message (e.g. subject + sender + date)
    .PARAMETER WhatIf
        If true, only estimates — does not delete.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SearchQuery,
        [switch]$WhatIf,
        [int]$BatchSize = 50
    )

    $allMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
                    Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' } |
                    ForEach-Object { $_.PrimarySmtpAddress }

    $totalResults = @()
    $batches = [math]::Ceiling($allMailboxes.Count / $BatchSize)
    $action = if ($WhatIf) { 'Estimate' } else { 'DeleteContent' }

    for ($i = 0; $i -lt $batches; $i++) {
        $start = $i * $BatchSize
        $batch = $allMailboxes[$start..([math]::Min($start + $BatchSize - 1, $allMailboxes.Count - 1))]

        $batchResults = Invoke-MailboxSearch -Mailboxes $batch -SearchQuery $SearchQuery -Action $action -Force
        $totalResults += $batchResults
    }

    $summary = [PSCustomObject]@{
        Action           = $action
        SearchQuery      = $SearchQuery
        TotalMailboxes   = $allMailboxes.Count
        AffectedMailboxes = ($totalResults | Where-Object { $_.ResultItems -gt 0 } | Measure-Object).Count
        TotalItems       = ($totalResults | Measure-Object -Property ResultItems -Sum).Sum
        Timestamp        = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    return @{
        Summary = $summary
        Details = $totalResults
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAILBOX STATISTICS
# ═══════════════════════════════════════════════════════════════════════════════

function Get-MailboxQuickStats {
    <#
    .SYNOPSIS
        Get quick statistics for selected mailboxes.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Mailboxes)

    $results = @()
    foreach ($mbx in $Mailboxes) {
        try {
            $stats = Get-MailboxStatistics -Identity $mbx -ErrorAction Stop
            $results += [PSCustomObject]@{
                Mailbox       = $mbx
                DisplayName   = "$($stats.DisplayName)"
                ItemCount     = $stats.ItemCount
                TotalSize     = "$($stats.TotalItemSize)"
                DeletedItems  = $stats.DeletedItemCount
                DeletedSize   = "$($stats.TotalDeletedItemSize)"
                LastLogonTime = "$($stats.LastLogonTime)"
                LastLogoffTime = "$($stats.LastLogoffTime)"
                Database      = "$($stats.DatabaseName)"
            }
        } catch {
            $results += [PSCustomObject]@{
                Mailbox     = $mbx
                DisplayName = ''
                ItemCount   = 0
                Error       = "$_"
            }
        }
    }
    return $results
}

function Get-MailboxFolderStats {
    <#
    .SYNOPSIS
        Get folder-level statistics for a mailbox.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Mailbox)

    try {
        $folders = Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop
        return $folders | ForEach-Object {
            [PSCustomObject]@{
                FolderPath   = $_.FolderPath
                FolderType   = $_.FolderType
                ItemCount    = $_.ItemsInFolder
                FolderSize   = "$($_.FolderSize)"
                SubFolders   = $_.ItemsInFolderAndSubfolders
                OldestItem   = "$($_.OldestItemReceivedDate)"
                NewestItem   = "$($_.NewestItemReceivedDate)"
            }
        }
    } catch {
        return @()
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# FOLDER OPERATIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Get-MailboxFolderList {
    <#
    .SYNOPSIS
        Get flat list of folder paths for a mailbox (for folder picker).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Mailbox)
    try {
        $folders = Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop
        return $folders | ForEach-Object {
            [PSCustomObject]@{
                FolderPath   = $_.FolderPath
                FolderType   = $_.FolderType
                ItemCount    = $_.ItemsInFolder
                FolderSize   = "$($_.FolderSize)"
                OldestItem   = "$($_.OldestItemReceivedDate)"
                NewestItem   = "$($_.NewestItemReceivedDate)"
            }
        }
    } catch {
        return @()
    }
}

function Invoke-FolderCleanup {
    <#
    .SYNOPSIS
        Search and optionally delete messages from a specific folder.
        Uses EWS when folder is specified (Search-Mailbox doesn't support folder: keyword).
        Falls back to Search-Mailbox for all-folders search.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [string]$FolderPath,
        [int]$OlderThanDays = 0,
        [string]$Subject,
        [string]$From,
        [string]$SizeRange,
        [switch]$HasAttachment,
        [ValidateSet('Estimate','DeleteContent')]
        [string]$Action = 'Estimate',
        [string]$Server
    )

    # Build KQL query WITHOUT folder (not supported by Search-Mailbox)
    $queryParams = @{}
    if ($Subject)       { $queryParams['Subject'] = $Subject }
    if ($From)          { $queryParams['From'] = $From }
    if ($SizeRange)     { $queryParams['SizeRange'] = $SizeRange }
    if ($HasAttachment) { $queryParams['HasAttachment'] = $true }
    if ($OlderThanDays -gt 0) {
        $queryParams['EndDate'] = (Get-Date).AddDays(-$OlderThanDays)
    }
    $query = Build-SearchQuery @queryParams

    # If a specific folder is selected, use EWS for folder-scoped operations
    if ($FolderPath) {
        if (-not $Server) {
            $Server = (Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' } | Select-Object -First 1).Fqdn
        }
        return Invoke-FolderCleanupEWS -Mailbox $Mailbox -FolderPath $FolderPath `
            -SearchQuery $query -Action $Action -Server $Server
    }

    # All folders — use Search-Mailbox
    $searchParams = @{
        Mailboxes   = @($Mailbox)
        SearchQuery = $query
        Action      = $Action
    }
    if ($Action -eq 'DeleteContent') {
        $searchParams['Force'] = $true
    }

    return Invoke-MailboxSearch @searchParams
}

function Invoke-FolderCleanupEWS {
    <#
    .SYNOPSIS
        EWS-based folder search and delete. Scopes to a specific folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$SearchQuery = '*',
        [ValidateSet('Estimate','DeleteContent')]
        [string]$Action = 'Estimate',
        [Parameter(Mandatory)][string]$Server
    )

    $ewsUrl = "https://$Server/EWS/Exchange.asmx"
    $folderClean = $FolderPath.TrimStart('/').Trim()

    $parentFolderXml = Get-EwsFolderXml -Mailbox $Mailbox -FolderPath $folderClean -Server $Server

    # Build query string for EWS (skip if wildcard)
    $queryStringXml = ''
    if ($SearchQuery -and $SearchQuery -ne '*') {
        $queryStringXml = "<m:QueryString>$([System.Security.SecurityElement]::Escape($SearchQuery))</m:QueryString>"
    }

    # Paged FindItem: collects ALL matching item ids, not just the first page
    $found = Get-EwsFolderMessages -Mailbox $Mailbox -Server $Server `
                -ParentFolderXml $parentFolderXml -QueryStringXml $queryStringXml
    $itemIds = @($found.Items)
    $totalCount = $found.Total
    $totalSize = ($itemIds | Measure-Object -Property Size -Sum).Sum
    if (-not $totalSize) { $totalSize = 0 }

    $deleted = 0
    if ($Action -eq 'DeleteContent' -and $itemIds.Count -gt 0) {
        # Delete items in batches of 100
        for ($i = 0; $i -lt $itemIds.Count; $i += 100) {
            $batch = @($itemIds[$i..([Math]::Min($i + 99, $itemIds.Count - 1))])
            $itemIdXml = ($batch | ForEach-Object {
                "<t:ItemId Id=`"$($_.Id)`" ChangeKey=`"$($_.ChangeKey)`" />"
            }) -join "`n"

            try {
                $deleteBody = @"
    <m:DeleteItem DeleteType="SoftDelete" AffectedTaskOccurrences="AllOccurrences">
      <m:ItemIds>
        $itemIdXml
      </m:ItemIds>
    </m:DeleteItem>
"@
                $null = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $deleteBody -Mailbox $Mailbox
                $deleted += $batch.Count
            } catch {
                Write-Warning "Batch delete failed: $_"
            }
        }
    }

    $sizeStr = if ($totalSize -gt 1MB) { "$([math]::Round($totalSize/1MB, 2)) MB ($totalSize bytes)" }
               elseif ($totalSize -gt 1KB) { "$([math]::Round($totalSize/1KB, 1)) KB ($totalSize bytes)" }
               else { "$totalSize bytes" }

    return [PSCustomObject]@{
        Mailbox       = $Mailbox
        DisplayName   = ''
        Success       = $true
        ResultItems   = if ($Action -eq 'DeleteContent') { $deleted } else { $totalCount }
        ResultSize    = $sizeStr
        Action        = $Action
        SearchQuery   = if ($SearchQuery -eq '*') { "folder:`"$folderClean`"" } else { "$SearchQuery (folder:`"$folderClean`")" }
        TargetMailbox = ''
        TargetFolder  = ''
        Timestamp     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }
}

function Find-EwsFolderId {
    <#
    .SYNOPSIS
        Find EWS folder ID by folder path (for custom/non-distinguished folders).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$Server
    )

    $ewsUrl = "https://$Server/EWS/Exchange.asmx"

    # Split path and walk from msgfolderroot
    $pathParts = $FolderPath.Split('/\', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($pathParts.Count -eq 0) { return $null }
    $match = $null
    $currentParent = @"
        <t:DistinguishedFolderId Id="msgfolderroot">
          <t:Mailbox><t:EmailAddress>$([System.Security.SecurityElement]::Escape($Mailbox))</t:EmailAddress></t:Mailbox>
        </t:DistinguishedFolderId>
"@

    foreach ($part in $pathParts) {
        $findFolderBody = @"
    <m:FindFolder Traversal="Shallow">
      <m:FolderShape><t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties><t:FieldURI FieldURI="folder:DisplayName" /></t:AdditionalProperties>
      </m:FolderShape>
      <m:ParentFolderIds>$currentParent</m:ParentFolderIds>
    </m:FindFolder>
"@
        $xml = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $findFolderBody -Mailbox $Mailbox
        $ns = New-EwsNamespaceManager -Xml $xml

        $folders = $xml.SelectNodes('//t:Folder', $ns)
        $match = $null
        foreach ($f in $folders) {
            $dn = $f.SelectSingleNode('t:DisplayName', $ns)
            if ($dn -and $dn.InnerText -eq $part) {
                $match = $f.SelectSingleNode('t:FolderId', $ns)
                break
            }
        }
        if (-not $match) { return $null }
        $currentParent = "<t:FolderId Id=`"$($match.GetAttribute('Id'))`" />"
    }

    if ($match) { return $match.GetAttribute('Id') }
    return $null
}

function Invoke-PurgeDeletedItems {
    <#
    .SYNOPSIS
        Purge (hard-delete) items from Recoverable Items / Deletions folder.
        Uses Search-Mailbox -SearchDumpsterOnly -DeleteContent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [string]$SearchQuery = '*',
        [ValidateSet('Estimate','DeleteContent')]
        [string]$Action = 'Estimate'
    )

    $params = @{
        Identity    = $Mailbox
        SearchQuery = $SearchQuery
        SearchDumpsterOnly = $true
        ErrorAction = 'Stop'
    }

    if ($Action -eq 'Estimate') {
        $params['EstimateResultOnly'] = $true
    } else {
        $params['DeleteContent'] = $true
        $params['Force'] = $true
    }

    try {
        $result = Search-Mailbox @params
        return $result | ForEach-Object {
            [PSCustomObject]@{
                Mailbox     = "$($_.Identity)"
                DisplayName = "$($_.DisplayName)"
                Success     = $_.Success
                ResultItems = $_.ResultItemsCount
                ResultSize  = "$($_.ResultItemsSize)"
                Action      = "$Action (Dumpster)"
                SearchQuery = $SearchQuery
                Timestamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        }
    } catch {
        return [PSCustomObject]@{
            Mailbox = $Mailbox; Success = $false; ResultItems = 0
            Action = "$Action (Dumpster)"; Error = "$_"
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# DUPLICATE DETECTION
# ═══════════════════════════════════════════════════════════════════════════════

function Find-MailboxDuplicates {
    <#
    .SYNOPSIS
        Find duplicate messages in a mailbox via EWS.
        A duplicate = same Subject + From + received date within one folder.
        (Search-Mailbox cannot scope to a folder, so EWS is used.)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [string]$FolderPath,
        [int]$DaysBack = 30,
        [int]$MaxItemsPerFolder = 2000,
        [string]$Server
    )

    if (-not $Server) {
        $Server = (Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' } | Select-Object -First 1).Fqdn
    }

    $folders = @(Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop |
                 Where-Object { $_.ItemsInFolder -gt 0 })
    if ($FolderPath) {
        $folders = @($folders | Where-Object { $_.FolderPath -eq $FolderPath })
    }

    $sinceUtc = (Get-Date).AddDays(-$DaysBack).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $restrictionXml = @"
      <m:Restriction>
        <t:IsGreaterThanOrEqualTo>
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURIOrConstant>
            <t:Constant Value="$sinceUtc" />
          </t:FieldURIOrConstant>
        </t:IsGreaterThanOrEqualTo>
      </m:Restriction>
"@

    $results = @()
    foreach ($folder in $folders) {
        if ("$($folder.FolderType)" -match 'Recoverable|Audits|Calendar|Contacts|Tasks|Notes|Journal|Conversation|Sync') { continue }
        $folderName = "$($folder.FolderPath)".TrimStart('/')
        if (-not $folderName) { continue }

        try {
            $folderXml = Get-EwsFolderXml -Mailbox $Mailbox -FolderPath $folderName -Server $Server
            $found = Get-EwsFolderMessages -Mailbox $Mailbox -Server $Server `
                        -ParentFolderXml $folderXml -RestrictionXml $restrictionXml `
                        -MaxItems $MaxItemsPerFolder

            $groups = @($found.Items | Group-Object {
                $day = if ($_.Received.Length -ge 10) { $_.Received.Substring(0, 10) } else { $_.Received }
                "$($_.Subject)|$($_.From)|$day"
            } | Where-Object { $_.Count -gt 1 })

            $dupItems = 0
            foreach ($g in $groups) { $dupItems += ($g.Count - 1) }  # keep one per group

            $results += [PSCustomObject]@{
                FolderPath      = $folder.FolderPath
                FolderType      = "$($folder.FolderType)"
                ItemCount       = $folder.ItemsInFolder
                ItemsScanned    = $found.Items.Count
                DuplicateGroups = $groups.Count
                DuplicateItems  = $dupItems
                FolderSize      = "$($folder.FolderSize)"
                Status          = if ($groups.Count -gt 0) { 'PossibleDupes' } else { 'Normal' }
            }
        } catch {
            $results += [PSCustomObject]@{
                FolderPath      = $folder.FolderPath
                FolderType      = "$($folder.FolderType)"
                ItemCount       = $folder.ItemsInFolder
                ItemsScanned    = 0
                DuplicateGroups = 0
                DuplicateItems  = 0
                FolderSize      = "$($folder.FolderSize)"
                Status          = "Error: $_"
            }
        }
    }

    return $results
}

function Remove-FolderDuplicates {
    <#
    .SYNOPSIS
        Backup a folder's items into a new backup folder in the SAME mailbox
        via EWS, optionally removing them from the source folder.
        BackupOnly      = CopyItem (originals stay in the source folder).
        BackupAndDelete = MoveItem (single atomic backup + remove from source).
    .DESCRIPTION
        Search-Mailbox cannot scope to a folder and EWS cannot copy items
        across mailboxes, so the backup folder is created under the root of
        the same mailbox ("Backup-<folder>-<timestamp>").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$BackupFolderName,
        [ValidateSet('BackupOnly','BackupAndDelete')]
        [string]$Action = 'BackupOnly',
        [string]$Server
    )

    if (-not $Server) {
        $Server = (Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' } | Select-Object -First 1).Fqdn
    }

    $ewsUrl = "https://$Server/EWS/Exchange.asmx"
    $escapedMailbox = [System.Security.SecurityElement]::Escape($Mailbox)
    $folderClean = $FolderPath.TrimStart('/').Trim()

    $sourceXml = Get-EwsFolderXml -Mailbox $Mailbox -FolderPath $folderClean -Server $Server

    if (-not $BackupFolderName) {
        $leaf = ($folderClean -split '[\\/]')[-1]
        $BackupFolderName = "Backup-$leaf-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    }
    $escapedBackupName = [System.Security.SecurityElement]::Escape($BackupFolderName)

    # Create the backup folder under the mailbox root
    $createBody = @"
    <m:CreateFolder>
      <m:ParentFolderId>
        <t:DistinguishedFolderId Id="msgfolderroot">
          <t:Mailbox><t:EmailAddress>$escapedMailbox</t:EmailAddress></t:Mailbox>
        </t:DistinguishedFolderId>
      </m:ParentFolderId>
      <m:Folders>
        <t:Folder>
          <t:DisplayName>$escapedBackupName</t:DisplayName>
        </t:Folder>
      </m:Folders>
    </m:CreateFolder>
"@
    $xml = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $createBody -Mailbox $Mailbox
    $ns = New-EwsNamespaceManager -Xml $xml
    $backupIdNode = $xml.SelectSingleNode('//t:Folder/t:FolderId', $ns)
    if (-not $backupIdNode) {
        throw "Could not create backup folder '$BackupFolderName' in mailbox $Mailbox"
    }
    $backupFolderId = $backupIdNode.GetAttribute('Id')

    $found = Get-EwsFolderMessages -Mailbox $Mailbox -Server $Server -ParentFolderXml $sourceXml
    $itemIds = @($found.Items)

    $op = if ($Action -eq 'BackupAndDelete') { 'MoveItem' } else { 'CopyItem' }
    $processed = 0
    for ($i = 0; $i -lt $itemIds.Count; $i += 100) {
        $batch = @($itemIds[$i..([Math]::Min($i + 99, $itemIds.Count - 1))])
        $itemIdXml = ($batch | ForEach-Object {
            "<t:ItemId Id=`"$($_.Id)`" ChangeKey=`"$($_.ChangeKey)`" />"
        }) -join "`n        "
        $opBody = @"
    <m:$op>
      <m:ToFolderId>
        <t:FolderId Id="$backupFolderId" />
      </m:ToFolderId>
      <m:ItemIds>
        $itemIdXml
      </m:ItemIds>
    </m:$op>
"@
        try {
            $null = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $opBody -Mailbox $Mailbox
            $processed += $batch.Count
        } catch {
            Write-Warning "$op batch failed: $_"
        }
    }

    return [PSCustomObject]@{
        Mailbox        = $Mailbox
        SourceFolder   = "/$folderClean"
        BackupFolder   = $BackupFolderName
        Action         = if ($Action -eq 'BackupAndDelete') { 'Move (backup + remove from source)' } else { 'Copy (backup only)' }
        ItemsFound     = $itemIds.Count
        ItemsProcessed = $processed
        Success        = ($processed -eq $itemIds.Count)
        Timestamp      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SEARCH HISTORY & LOGGING
# ═══════════════════════════════════════════════════════════════════════════════

function Write-SearchLog {
    <#
    .SYNOPSIS
        Log search operations to CSV file for audit trail.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Action,
        [string]$SearchQuery,
        [string]$Scope,
        [string]$Result,
        [string]$Details
    )

    $logDir = Join-Path $env:APPDATA 'EXRESearcher'
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    $logFile = Join-Path $logDir 'search-audit.csv'

    $entry = [PSCustomObject]@{
        Timestamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Operator    = "$env:USERDOMAIN\$env:USERNAME"
        Action      = $Action
        SearchQuery = $SearchQuery
        Scope       = $Scope
        Result      = $Result
        Details     = $Details
        Computer    = $env:COMPUTERNAME
    }

    $entry | Export-Csv -Path $logFile -Append -NoTypeInformation -Encoding UTF8
}

function Get-SearchLog {
    [CmdletBinding()]
    param([int]$Last = 100)

    $logFile = Join-Path $env:APPDATA 'EXRESearcher\search-audit.csv'
    if (Test-Path $logFile) {
        return Import-Csv -Path $logFile -Encoding UTF8 | Select-Object -Last $Last
    }
    return @()
}

# ═══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

function Export-SearchResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('CSV','JSON')][string]$Format = 'CSV'
    )

    switch ($Format) {
        'CSV'  { $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 }
        'JSON' { $Data | ConvertTo-Json -Depth 5 | Set-Content -Path $FilePath -Encoding UTF8 }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# PERMISSIONS & DISCOVERY
# ═══════════════════════════════════════════════════════════════════════════════

function Test-SearchPermissions {
    <#
    .SYNOPSIS
        Check if current user has required roles for Search-Mailbox operations.
        Returns diagnostic info about RBAC roles.
    #>
    [CmdletBinding()]
    param()
    try {
        $roles = Get-ManagementRoleAssignment -RoleAssignee "$env:USERDOMAIN\$env:USERNAME" -ErrorAction SilentlyContinue
        $hasSearch = $false
        $hasImportExport = $false
        $hasDiscovery = $false

        foreach ($r in $roles) {
            $roleName = "$($r.Role)"
            if ($roleName -match 'Mailbox Search')          { $hasSearch = $true }
            if ($roleName -match 'Mailbox Import Export')    { $hasImportExport = $true }
            if ($roleName -match 'Discovery')                { $hasDiscovery = $true }
        }

        return [PSCustomObject]@{
            User              = "$env:USERDOMAIN\$env:USERNAME"
            MailboxSearch     = $hasSearch
            MailboxImportExport = $hasImportExport
            DiscoveryManagement = $hasDiscovery
            TotalRoles        = ($roles | Measure-Object).Count
            Details           = ($roles | ForEach-Object { "$($_.Role)" }) -join '; '
        }
    } catch {
        return [PSCustomObject]@{
            User    = "$env:USERDOMAIN\$env:USERNAME"
            Error   = "$_"
        }
    }
}

function Get-DiscoveryMailbox {
    <#
    .SYNOPSIS
        Find the Discovery Search Mailbox for use as target in Search-Mailbox operations.
    #>
    [CmdletBinding()]
    param()
    try {
        $discovery = Get-Mailbox -Filter "RecipientTypeDetails -eq 'DiscoveryMailbox'" -ResultSize 10 -ErrorAction Stop
        return $discovery | ForEach-Object {
            [PSCustomObject]@{
                DisplayName  = $_.DisplayName
                PrimarySmtp  = $_.PrimarySmtpAddress
                Database     = "$($_.Database)"
            }
        }
    } catch {
        return @()
    }
}

function Get-MailboxPermissions {
    <#
    .SYNOPSIS
        Get full access permissions on a mailbox.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Mailbox)
    try {
        $perms = Get-MailboxPermission -Identity $Mailbox -ErrorAction Stop |
                 Where-Object { -not $_.IsInherited -and $_.User -ne 'NT AUTHORITY\SELF' }
        return $perms | ForEach-Object {
            [PSCustomObject]@{
                Mailbox      = $Mailbox
                User         = "$($_.User)"
                AccessRights = ($_.AccessRights -join ', ')
                Deny         = $_.Deny
            }
        }
    } catch {
        return @()
    }
}
