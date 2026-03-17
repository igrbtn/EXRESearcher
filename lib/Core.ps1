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

function Invoke-EwsRequest {
    <#
    .SYNOPSIS
        Send EWS SOAP request. Tries without impersonation first, then with.
        Returns [xml] response.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EwsUrl,
        [Parameter(Mandatory)][string]$SoapBody,
        [string]$Mailbox
    )

    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
    $headers = @{ 'Content-Type' = 'text/xml; charset=utf-8' }

    $escapedMailbox = if ($Mailbox) { [System.Security.SecurityElement]::Escape($Mailbox) } else { '' }

    # SOAP without impersonation
    $soap1 = @"
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

    # SOAP with impersonation
    $soap2 = @"
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

    $ns = @{
        s = 'http://schemas.xmlsoap.org/soap/envelope/'
        m = 'http://schemas.microsoft.com/exchange/services/2006/messages'
        t = 'http://schemas.microsoft.com/exchange/services/2006/types'
    }
    # Also build a variant without Mailbox element (for own mailbox access)
    $soapBodyNoMbx = $SoapBody -replace '<t:Mailbox>\s*<t:EmailAddress>[^<]*</t:EmailAddress>\s*</t:Mailbox>', ''
    $soap0 = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
$soapBodyNoMbx
  </soap:Body>
</soap:Envelope>
"@

    $lastError = ''

    # Try: 1) own mailbox, 2) with Mailbox element (Full Access), 3) with impersonation
    foreach ($soap in @($soap0, $soap1, $soap2)) {
        try {
            $response = Invoke-WebRequest -Uri $EwsUrl -Method POST -Body $soap -Headers $headers `
                            -UseDefaultCredentials -ErrorAction Stop
            [xml]$xml = $response.Content

            # Check for SOAP fault
            $fault = $xml.SelectSingleNode('//s:Fault/faultstring', $ns)
            if ($fault) { $lastError = $fault.InnerText; continue }

            # Check for error response
            $respMsg = $xml.SelectNodes('//*[@ResponseClass]', $ns)
            $hasError = $false
            foreach ($r in $respMsg) {
                if ($r.ResponseClass -eq 'Error') {
                    $lastError = $r.MessageText
                    $hasError = $true
                    break
                }
            }
            if ($hasError) { continue }

            return $xml
        } catch {
            $lastError = "$_"
        }
    }

    throw "EWS request failed: $lastError"
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
    $escapedMailbox = [System.Security.SecurityElement]::Escape($Mailbox)
    $escapedQuery = [System.Security.SecurityElement]::Escape($SearchQuery)

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
        <t:DistinguishedFolderId Id="msgfolderroot">
          <t:Mailbox><t:EmailAddress>$escapedMailbox</t:EmailAddress></t:Mailbox>
        </t:DistinguishedFolderId>
      </m:ParentFolderIds>
      <m:QueryString>$escapedQuery</m:QueryString>
    </m:FindItem>
"@

    $xml = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $soapBody -Mailbox $Mailbox
    $ns = @{
        m = 'http://schemas.microsoft.com/exchange/services/2006/messages'
        t = 'http://schemas.microsoft.com/exchange/services/2006/types'
    }
    $responseMsg = $xml.SelectSingleNode('//m:FindItemResponseMessage', $ns)

    $items = $xml.SelectNodes('//t:Message', $ns)
    $results = @()

    foreach ($item in $items) {
        $fromName = $item.SelectSingleNode('t:From/t:Mailbox/t:Name', $ns)
        $fromAddr = $item.SelectSingleNode('t:From/t:Mailbox/t:EmailAddress', $ns)
        $toNodes  = $item.SelectNodes('t:ToRecipients/t:Mailbox', $ns)
        $toList   = @($toNodes | ForEach-Object {
            $n = $_.SelectSingleNode('t:Name', $ns)
            if ($n) { $n.InnerText } else { $_.SelectSingleNode('t:EmailAddress', $ns).InnerText }
        }) -join '; '

        $sizeBytes = 0
        $sizeNode = $item.SelectSingleNode('t:Size', $ns)
        if ($sizeNode) { [int]::TryParse($sizeNode.InnerText, [ref]$sizeBytes) | Out-Null }
        $sizeKB = [math]::Round($sizeBytes / 1024, 1)

        $results += [PSCustomObject]@{
            Subject     = $item.SelectSingleNode('t:Subject', $ns).InnerText
            From        = if ($fromName) { $fromName.InnerText } elseif ($fromAddr) { $fromAddr.InnerText } else { '' }
            To          = $toList
            Received    = $item.SelectSingleNode('t:DateTimeReceived', $ns).InnerText
            SizeKB      = $sizeKB
            HasAttach   = $item.SelectSingleNode('t:HasAttachments', $ns).InnerText
            Importance  = $item.SelectSingleNode('t:Importance', $ns).InnerText
            ItemClass   = $item.SelectSingleNode('t:ItemClass', $ns).InnerText
        }
    }

    return $results
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
# SEARCH-MAILBOX (Exchange 2019 SE native)
# ═══════════════════════════════════════════════════════════════════════════════

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

    # Strip KQL keywords not supported by Search-Mailbox
    $cleanQuery = $SearchQuery
    $cleanQuery = $cleanQuery -replace '\s*AND\s*folder:"[^"]*"', ''
    $cleanQuery = $cleanQuery -replace 'folder:"[^"]*"\s*(AND\s*)?', ''
    $cleanQuery = $cleanQuery -replace '\s*AND\s*hasattachment:\w+', ''
    $cleanQuery = $cleanQuery -replace 'hasattachment:\w+\s*(AND\s*)?', ''
    $cleanQuery = $cleanQuery.Trim()
    if (-not $cleanQuery) { $cleanQuery = '*' }

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

function Search-AllMailboxes {
    <#
    .SYNOPSIS
        Search across ALL mailboxes in the organization.
        Returns estimate results per mailbox.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SearchQuery,
        [int]$BatchSize = 50
    )

    $allMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
                    Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' } |
                    ForEach-Object { $_.PrimarySmtpAddress }

    $totalResults = @()
    $batches = [math]::Ceiling($allMailboxes.Count / $BatchSize)

    for ($i = 0; $i -lt $batches; $i++) {
        $start = $i * $BatchSize
        $batch = $allMailboxes[$start..([math]::Min($start + $BatchSize - 1, $allMailboxes.Count - 1))]

        $batchResults = Invoke-MailboxSearch -Mailboxes $batch -SearchQuery $SearchQuery -Action 'Estimate'
        $totalResults += $batchResults | Where-Object { $_.ResultItems -gt 0 }
    }

    return $totalResults | Sort-Object { [int]$_.ResultItems } -Descending
}

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
    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}

    # Map common folder paths to EWS distinguished folder IDs
    $folderClean = $FolderPath.TrimStart('/').Trim()
    $distinguishedMap = @{
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

    $distinguishedId = $distinguishedMap[$folderClean]

    # Build parent folder XML
    if ($distinguishedId) {
        $parentFolderXml = @"
        <t:DistinguishedFolderId Id="$distinguishedId">
          <t:Mailbox><t:EmailAddress>$([System.Security.SecurityElement]::Escape($Mailbox))</t:EmailAddress></t:Mailbox>
        </t:DistinguishedFolderId>
"@
    } else {
        # Custom folder — need to find its ID first
        $folderId = Find-EwsFolderId -Mailbox $Mailbox -FolderPath $folderClean -Server $Server
        if (-not $folderId) {
            throw "Folder '$folderClean' not found in mailbox $Mailbox"
        }
        $parentFolderXml = "<t:FolderId Id=`"$folderId`" />"
    }

    # Build query string for EWS (skip if wildcard)
    $queryStringXml = ''
    if ($SearchQuery -and $SearchQuery -ne '*') {
        $queryStringXml = "<m:QueryString>$([System.Security.SecurityElement]::Escape($SearchQuery))</m:QueryString>"
    }

    # FindItem to get items in the folder
    $soapBody = @"
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:Size" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="item:HasAttachments" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="1000" Offset="0" BasePoint="Beginning" />
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived" />
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        $parentFolderXml
      </m:ParentFolderIds>
      $queryStringXml
    </m:FindItem>
"@

    $xml = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $soapBody -Mailbox $Mailbox
    $ns = @{
        m = 'http://schemas.microsoft.com/exchange/services/2006/messages'
        t = 'http://schemas.microsoft.com/exchange/services/2006/types'
    }

    $totalCount = 0
    $rootFolder = $xml.SelectSingleNode('//m:RootFolder', $ns)
    if ($rootFolder) {
        [int]::TryParse($rootFolder.GetAttribute('TotalItemsInView'), [ref]$totalCount) | Out-Null
    }

    $items = $xml.SelectNodes('//t:Message', $ns)
    $totalSize = 0
    $itemIds = @()

    foreach ($item in $items) {
        $sizeNode = $item.SelectSingleNode('t:Size', $ns)
        if ($sizeNode) {
            $sz = 0
            [int]::TryParse($sizeNode.InnerText, [ref]$sz) | Out-Null
            $totalSize += $sz
        }
        $idNode = $item.SelectSingleNode('t:ItemId', $ns)
        if ($idNode) {
            $itemIds += @{ Id = $idNode.GetAttribute('Id'); ChangeKey = $idNode.GetAttribute('ChangeKey') }
        }
    }

    if ($Action -eq 'DeleteContent' -and $itemIds.Count -gt 0) {
        # Delete items in batches of 100
        for ($i = 0; $i -lt $itemIds.Count; $i += 100) {
            $batch = $itemIds[$i..[Math]::Min($i + 99, $itemIds.Count - 1)]
            $itemIdXml = ($batch | ForEach-Object {
                "<t:ItemId Id=`"$($_.Id)`" ChangeKey=`"$($_.ChangeKey)`" />"
            }) -join "`n"

            $deleteBody = @"
    <m:DeleteItem DeleteType="SoftDelete" AffectedTaskOccurrences="AllOccurrences">
      <m:ItemIds>
        $itemIdXml
      </m:ItemIds>
    </m:DeleteItem>
"@
            $null = Invoke-EwsRequest -EwsUrl $ewsUrl -SoapBody $deleteBody -Mailbox $Mailbox
        }
    }

    $sizeStr = if ($totalSize -gt 1MB) { "$([math]::Round($totalSize/1MB, 2)) MB ($totalSize bytes)" }
               elseif ($totalSize -gt 1KB) { "$([math]::Round($totalSize/1KB, 1)) KB ($totalSize bytes)" }
               else { "$totalSize bytes" }

    return [PSCustomObject]@{
        Mailbox       = $Mailbox
        DisplayName   = ''
        Success       = $true
        ResultItems   = $totalCount
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
    $ns = @{
        m = 'http://schemas.microsoft.com/exchange/services/2006/messages'
        t = 'http://schemas.microsoft.com/exchange/services/2006/types'
    }

    # Split path and walk from msgfolderroot
    $pathParts = $FolderPath.Split('/\', [System.StringSplitOptions]::RemoveEmptyEntries)
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

    return $match.GetAttribute('Id')
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
        Find potential duplicate messages in a mailbox.
        Strategy: search by subject+sender combinations and identify
        messages that appear multiple times (same subject, same sender, same day).
    .DESCRIPTION
        1. Gets folder statistics to identify folders with many items.
        2. For each target folder, uses Search-Mailbox to export to discovery mailbox.
        3. Analyzes results by grouping on subject+from+date.

        Simplified approach: searches for exact subject matches and returns
        folders/counts where duplicates are likely.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [string]$FolderPath,
        [string]$TargetMailbox,
        [int]$DaysBack = 30
    )

    $results = @()

    # Step 1: Get folder stats to identify heavy folders
    $folders = Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop |
               Where-Object { $_.ItemsInFolder -gt 0 }

    if ($FolderPath) {
        $folders = $folders | Where-Object { $_.FolderPath -eq $FolderPath }
    }

    $startDate = (Get-Date).AddDays(-$DaysBack)

    # Step 2: For each folder with items, do estimate search
    # We search the mailbox scoped to each folder and look for subjects
    # that produce high result counts
    foreach ($folder in $folders) {
        $folderName = $folder.FolderPath.TrimStart('/')
        if (-not $folderName) { continue }
        if ($folder.FolderType -in @('RecoverableItems','Audits','Calendar','Contacts','Tasks')) { continue }

        $query = Build-SearchQuery -Folder $folderName -StartDate $startDate

        try {
            $estimate = Search-Mailbox -Identity $Mailbox -SearchQuery $query `
                        -EstimateResultOnly -ErrorAction Stop

            if ($estimate.ResultItemsCount -gt 0) {
                $results += [PSCustomObject]@{
                    FolderPath  = $folder.FolderPath
                    FolderType  = $folder.FolderType
                    ItemCount   = $folder.ItemsInFolder
                    SearchHits  = $estimate.ResultItemsCount
                    FolderSize  = "$($folder.FolderSize)"
                    OldestItem  = "$($folder.OldestItemReceivedDate)"
                    NewestItem  = "$($folder.NewestItemReceivedDate)"
                    Status      = if ($estimate.ResultItemsCount -gt $folder.ItemsInFolder) { 'PossibleDupes' } else { 'Normal' }
                }
            }
        } catch {
            $results += [PSCustomObject]@{
                FolderPath = $folder.FolderPath; FolderType = $folder.FolderType
                ItemCount = $folder.ItemsInFolder; Error = "$_"
            }
        }
    }

    return $results
}

function Remove-FolderDuplicates {
    <#
    .SYNOPSIS
        Remove duplicate messages from a specific folder.
        Strategy: Export folder content to a discovery mailbox (which dedupes),
        then delete original folder content and copy back from discovery.

        Simpler approach: Search for messages with same subject in the folder,
        log to target mailbox (which captures unique), then delete all from source
        folder and rely on the target copy.
    .DESCRIPTION
        This is a two-step process:
        1. Copy unique messages to target/backup mailbox folder
        2. Delete from source folder

        IMPORTANT: Always backup first! Use Estimate to review before deleting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$TargetMailbox,
        [string]$TargetFolder = 'DuplicateBackup',
        [ValidateSet('BackupOnly','BackupAndDelete')]
        [string]$Action = 'BackupOnly'
    )

    $folderName = $FolderPath.TrimStart('/')
    $query = Build-SearchQuery -Folder $folderName
    $results = @()

    # Step 1: Always backup first — copy to target mailbox
    try {
        $copyResult = Search-Mailbox -Identity $Mailbox -SearchQuery $query `
                      -TargetMailbox $TargetMailbox -TargetFolder $TargetFolder `
                      -ErrorAction Stop

        $results += [PSCustomObject]@{
            Step        = '1-Backup'
            Mailbox     = $Mailbox
            Folder      = $FolderPath
            Action      = 'CopyToTarget'
            ItemsCopied = $copyResult.ResultItemsCount
            TargetMailbox = $TargetMailbox
            TargetFolder  = $TargetFolder
            Success     = $copyResult.Success
            Timestamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
    } catch {
        return [PSCustomObject]@{
            Step = '1-Backup'; Action = 'CopyToTarget'; Success = $false; Error = "$_"
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
    }

    # Step 2: Delete from source folder (only if BackupAndDelete)
    if ($Action -eq 'BackupAndDelete') {
        try {
            $deleteResult = Search-Mailbox -Identity $Mailbox -SearchQuery $query `
                           -DeleteContent -Force -ErrorAction Stop

            $results += [PSCustomObject]@{
                Step         = '2-Delete'
                Mailbox      = $Mailbox
                Folder       = $FolderPath
                Action       = 'DeleteContent'
                ItemsDeleted = $deleteResult.ResultItemsCount
                Success      = $deleteResult.Success
                Timestamp    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        } catch {
            $results += [PSCustomObject]@{
                Step = '2-Delete'; Action = 'DeleteContent'; Success = $false; Error = "$_"
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
        }
    }

    return $results
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
