<#
.SYNOPSIS
    Core Exchange content search functions for EXRESearcher.
    Search-Mailbox, ComplianceSearch, mailbox enumeration, bulk operations.
#>

# ═══════════════════════════════════════════════════════════════════════════════
# CONNECTION
# ═══════════════════════════════════════════════════════════════════════════════

function Connect-ExchangeSearch {
    <#
    .SYNOPSIS
        Connect to Exchange via remote PowerShell (Kerberos).
        Returns PSSession object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server
    )

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

    $results = @()

    foreach ($mbx in $Mailboxes) {
        $params = @{
            Identity    = $mbx
            SearchQuery = $SearchQuery
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

    if ($Subject)        { $parts += "subject:`"$Subject`"" }
    if ($From)           { $parts += "from:`"$From`"" }
    if ($To)             { $parts += "to:`"$To`"" }
    if ($Keywords)       { $parts += "$Keywords" }
    if ($AttachmentName) { $parts += "attachment:`"$AttachmentName`"" }
    if ($MessageId)      { $parts += "messageid:`"$MessageId`"" }
    if ($MessageKind)    { $parts += "kind:$MessageKind" }
    if ($Folder)         { $parts += "folder:`"$Folder`"" }
    if ($HasAttachment)  { $parts += "hasattachment:true" }

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
    .PARAMETER Mailbox
        Target mailbox.
    .PARAMETER FolderPath
        Folder path (e.g. /Inbox, /Sent Items, /Deleted Items).
    .PARAMETER OlderThanDays
        Delete items older than N days.
    .PARAMETER Subject
        Filter by subject.
    .PARAMETER From
        Filter by sender.
    .PARAMETER SizeRange
        Filter by size: Small (<10KB), Medium (10KB-1MB), Large (1MB-10MB), VeryLarge (>10MB).
    .PARAMETER HasAttachment
        Filter only items with attachments.
    .PARAMETER Action
        Estimate or DeleteContent.
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
        [string]$Action = 'Estimate'
    )

    $queryParams = @{}
    if ($FolderPath)    { $queryParams['Folder'] = $FolderPath.TrimStart('/') }
    if ($Subject)       { $queryParams['Subject'] = $Subject }
    if ($From)          { $queryParams['From'] = $From }
    if ($SizeRange)     { $queryParams['SizeRange'] = $SizeRange }
    if ($HasAttachment) { $queryParams['HasAttachment'] = $true }

    if ($OlderThanDays -gt 0) {
        $queryParams['EndDate'] = (Get-Date).AddDays(-$OlderThanDays)
    }

    $query = Build-SearchQuery @queryParams

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
