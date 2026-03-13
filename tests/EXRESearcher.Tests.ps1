<#
.SYNOPSIS
    Pester v5 tests for EXRESearcher.
    Tests core functions that don't require Exchange connectivity.
#>

BeforeAll {
    $scriptRoot = Split-Path -Parent $PSScriptRoot
    . "$scriptRoot/lib/Core.ps1"
    . "$scriptRoot/lib/Settings.ps1"
}

Describe 'Build-SearchQuery' {
    It 'returns wildcard when no filters' {
        $result = Build-SearchQuery
        $result | Should -Be '*'
    }

    It 'builds subject filter' {
        $result = Build-SearchQuery -Subject 'invoice'
        $result | Should -Be 'subject:"invoice"'
    }

    It 'builds from filter' {
        $result = Build-SearchQuery -From 'user@domain.com'
        $result | Should -Be 'from:"user@domain.com"'
    }

    It 'builds combined filters with AND' {
        $result = Build-SearchQuery -Subject 'test' -From 'sender@test.com'
        $result | Should -Match 'subject:"test"'
        $result | Should -Match 'AND'
        $result | Should -Match 'from:"sender@test.com"'
    }

    It 'builds date range filter' {
        $start = [datetime]'2026-01-01'
        $end = [datetime]'2026-01-31'
        $result = Build-SearchQuery -StartDate $start -EndDate $end
        $result | Should -Match 'received>=2026-01-01'
        $result | Should -Match 'received<=2026-01-31'
    }

    It 'builds attachment filter' {
        $result = Build-SearchQuery -AttachmentName 'report.xlsx'
        $result | Should -Be 'attachment:"report.xlsx"'
    }

    It 'builds messageid filter' {
        $result = Build-SearchQuery -MessageId '<abc@domain.com>'
        $result | Should -Match 'messageid:'
    }

    It 'builds keywords' {
        $result = Build-SearchQuery -Keywords 'confidential secret'
        $result | Should -Be 'confidential secret'
    }

    It 'builds complex query' {
        $result = Build-SearchQuery -Subject 'phishing' -From 'bad@evil.com' -StartDate ([datetime]'2026-03-01')
        $result | Should -Match 'subject:"phishing"'
        $result | Should -Match 'from:"bad@evil.com"'
        $result | Should -Match 'received>=2026-03-01'
        ($result -split 'AND').Count | Should -Be 3
    }
}

Describe 'Settings Functions' {
    It 'Get-AppSettings returns object with expected properties' {
        $settings = Get-AppSettings
        $settings | Should -Not -BeNullOrEmpty
    }

    It 'Initialize-AppData creates directory' {
        Initialize-AppData
        $path = Join-Path $env:APPDATA 'EXRESearcher'
        Test-Path $path | Should -Be $true
    }
}

Describe 'Export-SearchResults' {
    BeforeAll {
        $testData = @(
            [PSCustomObject]@{ Mailbox = 'test@test.com'; Items = 5 }
            [PSCustomObject]@{ Mailbox = 'test2@test.com'; Items = 0 }
        )
    }

    It 'exports to CSV' {
        $tempFile = [System.IO.Path]::GetTempFileName() + '.csv'
        Export-SearchResults -Data $testData -FilePath $tempFile -Format CSV
        Test-Path $tempFile | Should -Be $true
        $imported = Import-Csv $tempFile
        $imported.Count | Should -Be 2
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }

    It 'exports to JSON' {
        $tempFile = [System.IO.Path]::GetTempFileName() + '.json'
        Export-SearchResults -Data $testData -FilePath $tempFile -Format JSON
        Test-Path $tempFile | Should -Be $true
        $content = Get-Content $tempFile -Raw
        $content | Should -Match 'test@test.com'
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }
}

Describe 'Write-SearchLog' {
    It 'writes log entry without error' {
        { Write-SearchLog -Action 'Test' -SearchQuery 'test query' -Scope 'test@test.com' -Result '0 items' } | Should -Not -Throw
    }
}

Describe 'GUI Function Existence' {
    BeforeAll {
        # Load main script as dot-source (won't launch GUI due to InvocationName check)
        . "$scriptRoot/EXRESearcher.ps1"
    }

    It 'Show-EXRESearcherGUI function exists' {
        Get-Command Show-EXRESearcherGUI -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }
}
