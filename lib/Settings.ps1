<#
.SYNOPSIS
    Settings, cache, and operator logging for EXRESearcher.
#>

# ═══════════════════════════════════════════════════════════════════════════════
# APP DATA / SETTINGS
# ═══════════════════════════════════════════════════════════════════════════════

$script:AppDataPath = Join-Path $env:APPDATA 'EXRESearcher'

function Initialize-AppData {
    if (-not (Test-Path $script:AppDataPath)) {
        New-Item -Path $script:AppDataPath -ItemType Directory -Force | Out-Null
    }
}

function Get-AppSettings {
    $settingsFile = Join-Path $script:AppDataPath 'settings.json'
    if (Test-Path $settingsFile) {
        try {
            return Get-Content -Path $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
            return [PSCustomObject]@{}
        }
    }
    return [PSCustomObject]@{
        LastServer       = ''
        RecentServers    = @()
        WindowWidth      = 1400
        WindowHeight     = 900
        SearchHistory    = @()
        DefaultBatchSize = 50
        MaxResults       = 500
    }
}

function Save-AppSettings {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Settings)
    $settingsFile = Join-Path $script:AppDataPath 'settings.json'
    try {
        $Settings | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsFile -Encoding UTF8
    } catch {}
}

function Update-RecentServers {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Server)
    try {
        $settings = Get-AppSettings
        $recent = @($settings.RecentServers)
        $recent = @($Server) + @($recent | Where-Object { $_ -ne $Server })
        if ($recent.Count -gt 10) { $recent = $recent[0..9] }
        $settings.RecentServers = $recent
        $settings.LastServer = $Server
        Save-AppSettings -Settings $settings
    } catch {}
}

function Add-SearchHistory {
    <#
    .SYNOPSIS
        Save a search query to history for quick re-use.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Query,
        [string]$Scope,
        [string]$Action
    )
    try {
        $settings = Get-AppSettings
        $history = @($settings.SearchHistory)
        $entry = @{
            Query     = $Query
            Scope     = $Scope
            Action    = $Action
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        $history = @($entry) + @($history)
        if ($history.Count -gt 50) { $history = $history[0..49] }
        $settings.SearchHistory = $history
        Save-AppSettings -Settings $settings
    } catch {}
}

# ═══════════════════════════════════════════════════════════════════════════════
# CACHE
# ═══════════════════════════════════════════════════════════════════════════════

$script:Cache = @{}

function Get-CachedData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Key,
        [int]$MaxAgeSec = 300
    )
    if ($script:Cache.ContainsKey($Key)) {
        $entry = $script:Cache[$Key]
        $age = (Get-Date) - $entry.Timestamp
        if ($age.TotalSeconds -lt $MaxAgeSec) {
            return $entry.Data
        }
    }
    return $null
}

function Set-CachedData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Data
    )
    $script:Cache[$Key] = @{
        Data      = $Data
        Timestamp = Get-Date
    }
}

function Clear-Cache {
    $script:Cache.Clear()
}

# ═══════════════════════════════════════════════════════════════════════════════
# OPERATOR AUDIT LOG
# ═══════════════════════════════════════════════════════════════════════════════

function Write-OperatorLog {
    <#
    .SYNOPSIS
        Log operator actions for audit (especially destructive operations).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Action,
        [string]$Target,
        [string]$Details
    )
    $logFile = Join-Path $script:AppDataPath 'operator-log.csv'
    $entry = [PSCustomObject]@{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Operator  = "$env:USERDOMAIN\$env:USERNAME"
        Action    = $Action
        Target    = $Target
        Details   = $Details
        Computer  = $env:COMPUTERNAME
    }
    try {
        $entry | Export-Csv -Path $logFile -Append -NoTypeInformation -Encoding UTF8
    } catch {}
}

function Get-OperatorLog {
    [CmdletBinding()]
    param([int]$Last = 100)
    $logFile = Join-Path $script:AppDataPath 'operator-log.csv'
    if (Test-Path $logFile) {
        try {
            return Import-Csv -Path $logFile -Encoding UTF8 | Select-Object -Last $Last
        } catch { return @() }
    }
    return @()
}
