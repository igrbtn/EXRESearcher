<#
.SYNOPSIS
    EXRESearcher v1.0 — Exchange Content Search & Cleanup GUI.
.DESCRIPTION
    WinForms GUI for searching mailbox content, compliance searches,
    and organization-wide message deletion (phishing/malware cleanup).
    All Exchange operations run asynchronously via runspaces.
.NOTES
    Version: 1.3.0
    Requires: Exchange 2019 SE, Windows PowerShell 5.1
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptRoot = $PSScriptRoot
. "$scriptRoot/lib/Core.ps1"
. "$scriptRoot/lib/Settings.ps1"
. "$scriptRoot/lib/AsyncRunner.ps1"

try { Initialize-AppData } catch {}
$script:Settings = try { Get-AppSettings } catch { @{} }

# ─── Script-scope state ─────────────────────────────────────────────────────
$script:Session               = $null
$script:LastMailboxList        = @()
$script:LastSearchResults      = @()
$script:LastComplianceSearches = @()
$script:LastOrgResults         = @()
$script:LastStatsData          = @()
$script:LastFolderStats        = @()
$script:LastAuditLog           = @()
$script:LastFolderCleanup     = @()
$script:LastDuplicates        = @()

# ═══════════════════════════════════════════════════════════════════════════════
# GUI HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

function New-StyledDGV {
    param([switch]$Multi)
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.ReadOnly = $true
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.SelectionMode = 'FullRowSelect'
    $dgv.AutoSizeColumnsMode = 'Fill'
    $dgv.RowHeadersVisible = $false
    $dgv.BackgroundColor = [System.Drawing.Color]::White
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245,245,250)
    $dgv.BorderStyle = 'None'
    $dgv.Dock = 'Fill'
    if ($Multi) { $dgv.MultiSelect = $true } else { $dgv.MultiSelect = $false }
    return $dgv
}

function New-Btn {
    param([string]$Text, [int]$W=110, [string]$Color='')
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Width = $W
    $btn.Height = 28
    $btn.FlatStyle = 'Flat'
    $btn.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4)
    if ($Color -eq 'Blue') {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(50,100,200)
        $btn.ForeColor = [System.Drawing.Color]::White
    } elseif ($Color -eq 'Red') {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(200,50,50)
        $btn.ForeColor = [System.Drawing.Color]::White
    } elseif ($Color -eq 'Green') {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(50,150,80)
        $btn.ForeColor = [System.Drawing.Color]::White
    } elseif ($Color -eq 'Orange') {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(220,140,20)
        $btn.ForeColor = [System.Drawing.Color]::White
    }
    return $btn
}

function New-ConsoleTextBox {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ScrollBars = 'Both'
    $tb.WordWrap = $false
    $tb.ReadOnly = $true
    $tb.Font = New-Object System.Drawing.Font('Consolas', 10)
    $tb.BackColor = [System.Drawing.Color]::FromArgb(25,25,35)
    $tb.ForeColor = [System.Drawing.Color]::FromArgb(220,220,220)
    $tb.Dock = 'Fill'
    return $tb
}

function New-FlowBar {
    param([int]$H=40)
    $fp = New-Object System.Windows.Forms.FlowLayoutPanel
    $fp.Height = $H
    $fp.Dock = 'Top'
    $fp.FlowDirection = 'LeftToRight'
    $fp.WrapContents = $false
    $fp.Padding = New-Object System.Windows.Forms.Padding(2)
    return $fp
}

function New-BoldLabel {
    param([string]$Text)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding(4,8,4,4)
    return $lbl
}

function New-InlineLabel {
    param([string]$Text, [int]$MarginLeft=0)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding($MarginLeft,8,4,4)
    return $lbl
}

function Show-Export {
    param([object]$Data, [string]$DefaultName='export')
    if (-not $Data -or ($Data | Measure-Object).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No data to export.','Export','OK','Information')
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV Files (*.csv)|*.csv|JSON Files (*.json)|*.json'
    $sfd.FileName = "$DefaultName-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $fmt = if ($sfd.FileName -match '\.json$') { 'JSON' } else { 'CSV' }
            Export-SearchResults -Data $Data -FilePath $sfd.FileName -Format $fmt
            Update-StatusBar "Exported to $($sfd.FileName)"
        } catch {
            Update-StatusBar "Export error: $_"
        }
    }
}

function Update-StatusBar {
    param([string]$Text)
    try {
        if ($script:StatusLabel) { $script:StatusLabel.Text = $Text }
        if ($script:LastActionLabel) {
            $script:LastActionLabel.Text = "Last: $(Get-Date -Format 'HH:mm:ss') $Text"
        }
    } catch {}
}

function Set-DGVData {
    param(
        [System.Windows.Forms.DataGridView]$DGV,
        [array]$Data
    )
    try {
        $DGV.DataSource = $null
        if ($Data -and $Data.Count -gt 0) {
            $dt = New-Object System.Data.DataTable
            $props = $Data[0].PSObject.Properties | ForEach-Object { $_.Name }
            foreach ($p in $props) { [void]$dt.Columns.Add($p) }
            foreach ($item in $Data) {
                $row = $dt.NewRow()
                foreach ($p in $props) { $row[$p] = "$($item.$p)" }
                [void]$dt.Rows.Add($row)
            }
            $DGV.DataSource = $dt
        }
    } catch {}
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN GUI
# ═══════════════════════════════════════════════════════════════════════════════

function Show-EXRESearcherGUI {

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'EXRESearcher v1.3 — Exchange Content Search & Cleanup'
    $form.Size = New-Object System.Drawing.Size(1400, 900)
    $form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.StartPosition = 'CenterScreen'
    $form.KeyPreview = $true

    try {
        if ($script:Settings.WindowWidth -and $script:Settings.WindowHeight) {
            $form.Size = New-Object System.Drawing.Size([int]$script:Settings.WindowWidth, [int]$script:Settings.WindowHeight)
        }
    } catch {}

    # ─── Top Panel ───────────────────────────────────────────────────────────
    $topPanel = New-FlowBar -H 44
    $topPanel.BackColor = [System.Drawing.Color]::FromArgb(240,240,245)

    $lblServer = New-BoldLabel -Text 'Exchange Server:'
    $txtServer = New-Object System.Windows.Forms.TextBox
    $txtServer.Width = 220
    $txtServer.Height = 24
    $txtServer.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $txtServer.AutoCompleteMode = 'SuggestAppend'
    $txtServer.AutoCompleteSource = 'CustomSource'
    $autoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
    try {
        if ($script:Settings.RecentServers) {
            foreach ($s in $script:Settings.RecentServers) { [void]$autoComplete.Add($s) }
        }
    } catch {}
    $txtServer.AutoCompleteCustomSource = $autoComplete
    try { if ($script:Settings.LastServer) { $txtServer.Text = $script:Settings.LastServer } } catch {}

    $btnConnect = New-Btn -Text 'Connect' -W 90 -Color 'Blue'
    $btnDisconnect = New-Btn -Text 'Disconnect' -W 90 -Color 'Red'
    $btnDisconnect.Visible = $false
    $lblConnStatus = New-InlineLabel -Text 'Not connected' -MarginLeft 6
    $lblConnStatus.ForeColor = [System.Drawing.Color]::Gray

    $chkSafeMode = New-Object System.Windows.Forms.CheckBox
    $chkSafeMode.Text = 'Safe Mode (Show commands)'
    $chkSafeMode.AutoSize = $true
    $chkSafeMode.Checked = $true
    $chkSafeMode.Margin = New-Object System.Windows.Forms.Padding(20,8,3,4)
    $chkSafeMode.ForeColor = [System.Drawing.Color]::FromArgb(100,100,100)

    $topPanel.Controls.AddRange(@($lblServer, $txtServer, $btnConnect, $btnDisconnect, $lblConnStatus, $chkSafeMode))
    $form.Controls.Add($topPanel)

    # ─── Status Bar ──────────────────────────────────────────────────────────
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:StatusLabel.Spring = $true
    $script:StatusLabel.TextAlign = 'MiddleLeft'
    $script:StatusLabel.Text = 'Ready'
    $script:LastActionLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:LastActionLabel.Text = ''
    $script:LastActionLabel.Alignment = 'Right'
    [void]$statusStrip.Items.Add($script:StatusLabel)
    [void]$statusStrip.Items.Add($script:LastActionLabel)
    $form.Controls.Add($statusStrip)

    # ─── Job Console Panel ───────────────────────────────────────────────────
    $jobPanel = New-JobConsolePanel -Height 130
    $form.Controls.Add($jobPanel)

    # ─── Tabs ────────────────────────────────────────────────────────────────
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 1: MAILBOX SEARCH (Search-Mailbox)
    # ═══════════════════════════════════════════════════════════════════════════
    $tabSearch = New-Object System.Windows.Forms.TabPage
    $tabSearch.Text = 'Mailbox Search'

    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Dock = 'Fill'

    # --- Search filters row 1 ---
    $searchBar1 = New-FlowBar -H 38
    $lblSubject = New-InlineLabel -Text 'Subject:'
    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Width = 200; $txtSubject.Height = 24
    $txtSubject.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblFrom = New-InlineLabel -Text 'From:'
    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Width = 180; $txtFrom.Height = 24
    $txtFrom.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblTo = New-InlineLabel -Text 'To:'
    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Width = 180; $txtTo.Height = 24
    $txtTo.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblKeywords = New-InlineLabel -Text 'Keywords:'
    $txtKeywords = New-Object System.Windows.Forms.TextBox
    $txtKeywords.Width = 200; $txtKeywords.Height = 24
    $txtKeywords.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $searchBar1.Controls.AddRange(@($lblSubject, $txtSubject, $lblFrom, $txtFrom, $lblTo, $txtTo, $lblKeywords, $txtKeywords))

    # --- Search filters row 2 ---
    $searchBar2 = New-FlowBar -H 38
    $lblDateFrom = New-InlineLabel -Text 'From date:'
    $dtpFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtpFrom.Format = 'Custom'; $dtpFrom.CustomFormat = 'yyyy-MM-dd'
    $dtpFrom.Width = 110; $dtpFrom.Value = (Get-Date).AddDays(-7)
    $dtpFrom.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblDateTo = New-InlineLabel -Text 'To date:'
    $dtpTo = New-Object System.Windows.Forms.DateTimePicker
    $dtpTo.Format = 'Custom'; $dtpTo.CustomFormat = 'yyyy-MM-dd'
    $dtpTo.Width = 110
    $dtpTo.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblAttach = New-InlineLabel -Text 'Attachment:'
    $txtAttach = New-Object System.Windows.Forms.TextBox
    $txtAttach.Width = 140; $txtAttach.Height = 24
    $txtAttach.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblMsgId = New-InlineLabel -Text 'MessageId:'
    $txtMsgId = New-Object System.Windows.Forms.TextBox
    $txtMsgId.Width = 200; $txtMsgId.Height = 24
    $txtMsgId.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $searchBar2.Controls.AddRange(@($lblDateFrom, $dtpFrom, $lblDateTo, $dtpTo, $lblAttach, $txtAttach, $lblMsgId, $txtMsgId))

    # --- Scope & action row ---
    $searchBar3 = New-FlowBar -H 38
    $lblScope = New-InlineLabel -Text 'Scope:'
    $txtScope = New-Object System.Windows.Forms.TextBox
    $txtScope.Width = 300; $txtScope.Height = 24
    $txtScope.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $txtScope.Text = '(enter mailbox or comma-separated list)'
    $txtScope.ForeColor = [System.Drawing.Color]::Gray
    $txtScope.Add_GotFocus({
        if ($txtScope.ForeColor -eq [System.Drawing.Color]::Gray) {
            $txtScope.Text = ''
            $txtScope.ForeColor = [System.Drawing.Color]::Black
        }
    })
    $txtScope.Add_LostFocus({
        if (-not $txtScope.Text) {
            $txtScope.Text = '(enter mailbox or comma-separated list)'
            $txtScope.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $btnEstimate = New-Btn -Text 'Estimate' -W 90 -Color 'Blue'
    $btnSearchLog = New-Btn -Text 'Search + Log' -W 110 -Color 'Green'
    $btnSearchCopy = New-Btn -Text 'Copy to Mailbox' -W 120 -Color 'Orange'
    $btnSearchDelete = New-Btn -Text 'Search + Delete' -W 120 -Color 'Red'
    $searchBar3.Controls.AddRange(@($lblScope, $txtScope, $btnEstimate, $btnSearchLog, $btnSearchCopy, $btnSearchDelete))

    # --- Target & query row ---
    $searchBar4 = New-FlowBar -H 38
    $lblTarget = New-InlineLabel -Text 'Target Mailbox:'
    $txtTarget = New-Object System.Windows.Forms.TextBox
    $txtTarget.Width = 220; $txtTarget.Height = 24
    $txtTarget.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblTargetFolder = New-InlineLabel -Text 'Folder:'
    $txtTargetFolder = New-Object System.Windows.Forms.TextBox
    $txtTargetFolder.Width = 150; $txtTargetFolder.Height = 24; $txtTargetFolder.Text = 'SearchResults'
    $txtTargetFolder.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $btnCheckPerms = New-Btn -Text 'Check Permissions' -W 130
    $lblQueryPreview = New-InlineLabel -Text 'Query: (build filters above)' -MarginLeft 10
    $searchBar4.Controls.AddRange(@($lblTarget, $txtTarget, $lblTargetFolder, $txtTargetFolder, $btnCheckPerms, $lblQueryPreview))

    # --- Results ---
    $dgvSearchResults = New-StyledDGV -Multi

    $dgvSearchResults.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $successCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Success') { $successCol = $c.Index; break } }
            if ($null -ne $successCol) {
                $val = "$($row.Cells[$successCol].Value)"
                if ($val -eq 'True')  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,255,200) }
                elseif ($val -eq 'False') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,200,200) }
            }
            $itemCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'ResultItems') { $itemCol = $c.Index; break } }
            if ($null -ne $itemCol) {
                $val = "$($row.Cells[$itemCol].Value)"
                $num = 0
                if ([int]::TryParse($val, [ref]$num) -and $num -gt 0) {
                    $row.Cells[$itemCol].Style.ForeColor = [System.Drawing.Color]::FromArgb(200,50,50)
                    $row.Cells[$itemCol].Style.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
                }
            }
        } catch {}
    })

    # --- Double-click: preview found messages via EWS ---
    $dgvSearchResults.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $row = $s.Rows[$e.RowIndex]
        $mbx = "$($row.Cells['Mailbox'].Value)"
        $query = "$($row.Cells['SearchQuery'].Value)"
        $count = "$($row.Cells['ResultItems'].Value)"
        if (-not $mbx -or -not $query -or $count -eq '0') { return }

        $serverFqdn = $txtServer.Text.Trim()
        Update-StatusBar "Loading message preview for $mbx..."

        Start-AsyncJob -Name "Preview $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $SearchQuery, $Server)
            return @(Get-MailboxMessagePreview -Mailbox $Mailbox -SearchQuery $SearchQuery -Server $Server -MaxResults 200)
        } -Parameters @{ Mailbox = $mbx; SearchQuery = $query; Server = $serverFqdn } -OnComplete {
            param($result)
            try {
                $messages = @($result)
                if ($messages.Count -eq 0) {
                    Update-StatusBar "No messages returned from EWS preview"
                    return
                }

                # Build preview dialog
                $dlg = New-Object System.Windows.Forms.Form
                $dlg.Text = "Message Preview: $mbx ($($messages.Count) messages)"
                $dlg.Size = New-Object System.Drawing.Size(1100, 600)
                $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
                $dlg.StartPosition = 'CenterParent'
                $dlg.MinimumSize = New-Object System.Drawing.Size(800, 400)

                $grid = New-StyledDGV -Multi
                $grid.Dock = 'Fill'
                Set-DGVData -DGV $grid -Data $messages

                # Resize columns
                foreach ($col in $grid.Columns) {
                    switch ($col.Name) {
                        'Subject'   { $col.Width = 350 }
                        'From'      { $col.Width = 180 }
                        'To'        { $col.Width = 180 }
                        'Received'  { $col.Width = 150 }
                        'SizeKB'    { $col.Width = 70; $col.HeaderText = 'Size (KB)' }
                        'HasAttach' { $col.Width = 50; $col.HeaderText = 'Attach' }
                        'Importance'{ $col.Width = 70 }
                        'ItemClass' { $col.Width = 120; $col.HeaderText = 'Type' }
                    }
                }

                $bottomBar = New-FlowBar -H 38
                $bottomBar.Dock = 'Bottom'
                $btnExportPreview = New-Btn -Text 'Export...' -W 80
                $lblPreviewCount = New-InlineLabel -Text "$($messages.Count) messages" -MarginLeft 10
                $btnExportPreview.Add_Click({ Show-Export -Data $messages -DefaultName 'message-preview' })
                $bottomBar.Controls.AddRange(@($btnExportPreview, $lblPreviewCount))

                $dlg.Controls.Add($grid)
                $dlg.Controls.Add($bottomBar)
                $grid.BringToFront()

                Update-StatusBar "Preview loaded: $($messages.Count) messages in $mbx"
                [void]$dlg.ShowDialog($form)
                $dlg.Dispose()
            } catch {
                Update-StatusBar "Preview display error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Preview error: $err"
            $fixMsg = if ($err -match '(?i)denied|impersonat|403|401|unauthorized') {
                "`n`nFix: Run this in Exchange Management Shell:`nNew-ManagementRoleAssignment -Name `"EWSImpersonation`" -Role `"ApplicationImpersonation`" -User $env:USERNAME"
            } else { '' }
            [System.Windows.Forms.MessageBox]::Show(
                "Could not load message preview:`n$err$fixMsg",
                'Preview Error', 'OK', 'Warning')
        }
    })

    $searchBottom = New-FlowBar -H 38
    $searchBottom.Dock = 'Bottom'
    $btnSearchExport = New-Btn -Text 'Export...' -W 80
    $lblSearchCount = New-InlineLabel -Text '0 results' -MarginLeft 10
    $searchBottom.Controls.AddRange(@($btnSearchExport, $lblSearchCount))

    $searchPanel.Controls.Add($dgvSearchResults)
    $searchPanel.Controls.Add($searchBottom)
    $searchPanel.Controls.Add($searchBar4)
    $searchPanel.Controls.Add($searchBar3)
    $searchPanel.Controls.Add($searchBar2)
    $searchPanel.Controls.Add($searchBar1)
    $tabSearch.Controls.Add($searchPanel)

    # --- Build query helper ---
    $buildCurrentQuery = {
        try {
            $params = @{}
            if ($txtSubject.Text)  { $params['Subject'] = $txtSubject.Text }
            if ($txtFrom.Text)     { $params['From'] = $txtFrom.Text }
            if ($txtTo.Text)       { $params['To'] = $txtTo.Text }
            if ($txtKeywords.Text) { $params['Keywords'] = $txtKeywords.Text }
            if ($txtAttach.Text)   { $params['AttachmentName'] = $txtAttach.Text }
            if ($txtMsgId.Text)    { $params['MessageId'] = $txtMsgId.Text }
            $params['StartDate'] = $dtpFrom.Value
            $params['EndDate']   = $dtpTo.Value
            $query = Build-SearchQuery @params
            $lblQueryPreview.Text = "Query: $query"
            return $query
        } catch {
            return '*'
        }
    }

    # Update query preview on filter change
    foreach ($ctrl in @($txtSubject, $txtFrom, $txtTo, $txtKeywords, $txtAttach, $txtMsgId)) {
        $ctrl.Add_TextChanged({ & $buildCurrentQuery | Out-Null })
    }
    $dtpFrom.Add_ValueChanged({ & $buildCurrentQuery | Out-Null })
    $dtpTo.Add_ValueChanged({ & $buildCurrentQuery | Out-Null })

    $getMailboxScope = {
        $scopeText = $txtScope.Text
        if (-not $scopeText -or $scopeText -match '^\(') { return @() }
        return @($scopeText -split '[,;\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    # --- Connection guard ---
    $requireConnection = {
        if (-not $script:Session) {
            [System.Windows.Forms.MessageBox]::Show(
                'Connect to an Exchange server first.',
                'Not Connected', 'OK', 'Warning')
            return $false
        }
        return $true
    }

    # --- Safe Mode: show command preview ---
    # Returns 'Execute' to proceed, 'Cancel' to abort
    $showCommandPreview = {
        param([string]$CommandText, [string]$Title = 'Command Preview')
        if (-not $chkSafeMode.Checked) { return 'Execute' }

        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = "Safe Mode: $Title"
        $dlg.Size = New-Object System.Drawing.Size(750, 420)
        $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $dlg.StartPosition = 'CenterParent'
        $dlg.MinimumSize = New-Object System.Drawing.Size(500, 300)
        $dlg.FormBorderStyle = 'Sizable'
        $dlg.MaximizeBox = $false

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = 'PowerShell command that will be executed:'
        $lbl.Dock = 'Top'
        $lbl.Height = 24
        $lbl.Padding = New-Object System.Windows.Forms.Padding(6,6,0,0)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Multiline = $true
        $txt.ScrollBars = 'Both'
        $txt.WordWrap = $false
        $txt.Font = New-Object System.Drawing.Font('Consolas', 10)
        $txt.Text = $CommandText
        $txt.Dock = 'Fill'
        $txt.ReadOnly = $false
        $txt.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
        $txt.ForeColor = [System.Drawing.Color]::FromArgb(220,220,180)

        $btnBar = New-Object System.Windows.Forms.FlowLayoutPanel
        $btnBar.Dock = 'Bottom'
        $btnBar.Height = 44
        $btnBar.FlowDirection = 'RightToLeft'
        $btnBar.Padding = New-Object System.Windows.Forms.Padding(6)

        $btnCancel = New-Btn -Text 'Cancel' -W 90
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $btnCopy = New-Btn -Text 'Copy' -W 90 -Color 'Blue'
        $btnExec = New-Btn -Text 'Execute' -W 90 -Color 'Green'
        $btnExec.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $btnCopy.Add_Click({
            [System.Windows.Forms.Clipboard]::SetText($txt.Text)
            $btnCopy.Text = 'Copied!'
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 1500
            $timer.Add_Tick({ $btnCopy.Text = 'Copy'; $timer.Stop(); $timer.Dispose() })
            $timer.Start()
        })

        $btnWhatIf = New-Btn -Text 'WhatIf' -W 90 -Color 'Orange'
        $btnWhatIf.Add_Click({
            try {
                # Take command from textbox, append -WhatIf to each line
                $lines = $txt.Text -split "`r?`n"
                $whatIfCmd = @()
                foreach ($line in $lines) {
                    $l = $line.Trim()
                    if (-not $l -or $l.StartsWith('#')) { continue }
                    if ($l -notmatch '-WhatIf') { $l += ' -WhatIf' }
                    $whatIfCmd += $l
                }
                if ($whatIfCmd.Count -eq 0) { return }
                $script = $whatIfCmd -join "`r`n"

                # Run WhatIf and capture output
                $btnWhatIf.Enabled = $false
                $btnWhatIf.Text = 'Running...'
                $dlg.Refresh()

                $output = try {
                    $sb = [scriptblock]::Create($script)
                    & $sb 2>&1 | Out-String
                } catch { "Error: $_" }

                # Show output in a result dialog
                $resDlg = New-Object System.Windows.Forms.Form
                $resDlg.Text = 'WhatIf Results'
                $resDlg.Size = New-Object System.Drawing.Size(700, 400)
                $resDlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
                $resDlg.StartPosition = 'CenterParent'

                $resTxt = New-Object System.Windows.Forms.TextBox
                $resTxt.Multiline = $true
                $resTxt.ScrollBars = 'Both'
                $resTxt.WordWrap = $false
                $resTxt.Font = New-Object System.Drawing.Font('Consolas', 10)
                $resTxt.Dock = 'Fill'
                $resTxt.ReadOnly = $true
                $resTxt.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
                $resTxt.ForeColor = [System.Drawing.Color]::FromArgb(180,220,180)
                $resTxt.Text = "# Command:`r`n$script`r`n`r`n# Output:`r`n$output"

                $resBar = New-Object System.Windows.Forms.FlowLayoutPanel
                $resBar.Dock = 'Bottom'
                $resBar.Height = 44
                $resBar.FlowDirection = 'RightToLeft'
                $resBar.Padding = New-Object System.Windows.Forms.Padding(6)
                $resClose = New-Btn -Text 'Close' -W 90
                $resClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $resCopy = New-Btn -Text 'Copy' -W 90 -Color 'Blue'
                $resCopy.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($resTxt.Text) })
                $resBar.Controls.AddRange(@($resClose, $resCopy))

                $resDlg.Controls.Add($resTxt)
                $resDlg.Controls.Add($resBar)
                $resTxt.BringToFront()
                [void]$resDlg.ShowDialog($dlg)
                $resDlg.Dispose()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("WhatIf error: $_", 'Error', 'OK', 'Error')
            } finally {
                $btnWhatIf.Enabled = $true
                $btnWhatIf.Text = 'WhatIf'
            }
        })

        $btnBar.Controls.AddRange(@($btnCancel, $btnExec, $btnWhatIf, $btnCopy))
        $dlg.Controls.Add($txt)
        $dlg.Controls.Add($lbl)
        $dlg.Controls.Add($btnBar)
        $txt.BringToFront()
        $dlg.AcceptButton = $btnExec
        $dlg.CancelButton = $btnCancel

        $result = $dlg.ShowDialog($form)
        $dlg.Dispose()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return 'Execute' }
        return 'Cancel'
    }

    # --- Delete confirmation with data loss warning ---
    # Returns 'OK' to proceed, 'ShowScript' to show command, 'Cancel' to abort
    $confirmDelete = {
        param([string]$Message, [string]$CommandText)

        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = 'Confirm Delete'
        $dlg.Size = New-Object System.Drawing.Size(500, 220)
        $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $dlg.StartPosition = 'CenterParent'
        $dlg.FormBorderStyle = 'FixedDialog'
        $dlg.MaximizeBox = $false
        $dlg.MinimizeBox = $false

        $iconBox = New-Object System.Windows.Forms.PictureBox
        $iconBox.Image = [System.Drawing.SystemIcons]::Warning.ToBitmap()
        $iconBox.Size = New-Object System.Drawing.Size(40,40)
        $iconBox.Location = New-Object System.Drawing.Point(15,15)
        $iconBox.SizeMode = 'Zoom'

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $Message
        $lbl.Location = New-Object System.Drawing.Point(65, 15)
        $lbl.Size = New-Object System.Drawing.Size(410, 60)

        $chkConfirm = New-Object System.Windows.Forms.CheckBox
        $chkConfirm.Text = 'I understand that data may be permanently lost'
        $chkConfirm.Location = New-Object System.Drawing.Point(65, 85)
        $chkConfirm.AutoSize = $true
        $chkConfirm.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

        $btnBar = New-Object System.Windows.Forms.FlowLayoutPanel
        $btnBar.Dock = 'Bottom'
        $btnBar.Height = 44
        $btnBar.FlowDirection = 'RightToLeft'
        $btnBar.Padding = New-Object System.Windows.Forms.Padding(6)

        $btnCancel = New-Btn -Text 'Cancel' -W 90
        $btnOK = New-Btn -Text 'Delete' -W 90 -Color 'Red'
        $btnOK.Enabled = $false
        $btnWhatIf2 = New-Btn -Text 'WhatIf' -W 90 -Color 'Orange'
        $btnScript = New-Btn -Text 'Show Script' -W 100 -Color 'Blue'

        $script:deleteDialogResult = 'Cancel'
        $chkConfirm.Add_CheckedChanged({ $btnOK.Enabled = $chkConfirm.Checked })
        $btnOK.Add_Click({ $script:deleteDialogResult = 'OK'; $dlg.Close() })
        $btnCancel.Add_Click({ $script:deleteDialogResult = 'Cancel'; $dlg.Close() })
        $btnScript.Add_Click({
            $script:deleteDialogResult = 'ShowScript'; $dlg.Close()
        })
        $btnWhatIf2.Add_Click({
            try {
                $lines = $CommandText -split "`r?`n"
                $whatIfCmd = @()
                foreach ($line in $lines) {
                    $l = $line.Trim()
                    if (-not $l -or $l.StartsWith('#')) { continue }
                    if ($l -notmatch '-WhatIf') { $l += ' -WhatIf' }
                    $whatIfCmd += $l
                }
                if ($whatIfCmd.Count -eq 0) { return }
                $script2 = $whatIfCmd -join "`r`n"
                $btnWhatIf2.Enabled = $false
                $btnWhatIf2.Text = 'Running...'
                $dlg.Refresh()
                $output = try {
                    $sb = [scriptblock]::Create($script2)
                    & $sb 2>&1 | Out-String
                } catch { "Error: $_" }
                $resDlg = New-Object System.Windows.Forms.Form
                $resDlg.Text = 'WhatIf Results'
                $resDlg.Size = New-Object System.Drawing.Size(700, 400)
                $resDlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
                $resDlg.StartPosition = 'CenterParent'
                $resTxt = New-Object System.Windows.Forms.TextBox
                $resTxt.Multiline = $true; $resTxt.ScrollBars = 'Both'; $resTxt.WordWrap = $false
                $resTxt.Font = New-Object System.Drawing.Font('Consolas', 10)
                $resTxt.Dock = 'Fill'; $resTxt.ReadOnly = $true
                $resTxt.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
                $resTxt.ForeColor = [System.Drawing.Color]::FromArgb(180,220,180)
                $resTxt.Text = "# Command:`r`n$script2`r`n`r`n# Output:`r`n$output"
                $resClose = New-Btn -Text 'Close' -W 90
                $resClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $resBar = New-Object System.Windows.Forms.FlowLayoutPanel
                $resBar.Dock = 'Bottom'; $resBar.Height = 44
                $resBar.FlowDirection = 'RightToLeft'
                $resBar.Padding = New-Object System.Windows.Forms.Padding(6)
                $resBar.Controls.Add($resClose)
                $resDlg.Controls.Add($resTxt); $resDlg.Controls.Add($resBar); $resTxt.BringToFront()
                [void]$resDlg.ShowDialog($dlg); $resDlg.Dispose()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("WhatIf error: $_", 'Error', 'OK', 'Error')
            } finally { $btnWhatIf2.Enabled = $true; $btnWhatIf2.Text = 'WhatIf' }
        })

        $btnBar.Controls.AddRange(@($btnCancel, $btnOK, $btnWhatIf2, $btnScript))
        $dlg.Controls.AddRange(@($iconBox, $lbl, $chkConfirm, $btnBar))

        [void]$dlg.ShowDialog($form)
        $dlg.Dispose()

        if ($script:deleteDialogResult -eq 'ShowScript') {
            # Show the command preview, then return Cancel (user can copy and run manually)
            & $showCommandPreview -CommandText $CommandText -Title 'Delete Command'
            return 'Cancel'
        }
        return $script:deleteDialogResult
    }

    # --- Search actions ---
    $doSearch = {
        param([string]$Action)
        if (-not (& $requireConnection)) { return }
        $query = & $buildCurrentQuery
        $mailboxes = & $getMailboxScope
        if ($mailboxes.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Enter mailbox(es) in Scope field.','Scope Required','OK','Warning')
            return
        }
        if ($Action -in @('LogOnly','CopyToFolder') -and -not $txtTarget.Text) {
            [System.Windows.Forms.MessageBox]::Show('Enter a Target Mailbox for Log/Copy operations.','Target Required','OK','Warning')
            return
        }
        $targetMbx = if ($Action -in @('LogOnly','CopyToFolder')) { $txtTarget.Text } else { '' }
        $targetFld = $txtTargetFolder.Text

        # Build command string for preview
        $mbxList = $mailboxes -join ', '
        $cmdLines = @()
        $queryEscaped = $query -replace '"', '`"'
        foreach ($m in $mailboxes) {
            $cmd = "Search-Mailbox -Identity `"$m`" -SearchQuery `"$queryEscaped`""
            switch ($Action) {
                'Estimate'      { $cmd += " -EstimateResultOnly" }
                'LogOnly'       { $cmd += " -TargetMailbox `"$targetMbx`" -TargetFolder `"$targetFld`" -LogOnly" }
                'CopyToFolder'  { $cmd += " -TargetMailbox `"$targetMbx`" -TargetFolder `"$targetFld`"" }
                'DeleteContent' { $cmd += " -DeleteContent -Force" }
            }
            $cmdLines += $cmd
        }
        $cmdText = $cmdLines -join "`r`n"

        # Delete confirmation with data loss warning
        if ($Action -eq 'DeleteContent') {
            $result = & $confirmDelete -Message "DELETE content matching:`n$query`n`nFrom $($mailboxes.Count) mailbox(es). This action is IRREVERSIBLE!" -CommandText $cmdText
            if ($result -ne 'OK') { return }
        } else {
            # Safe mode preview for non-destructive actions
            if ((& $showCommandPreview -CommandText $cmdText -Title "Search-Mailbox ($Action)") -eq 'Cancel') { return }
        }

        Update-StatusBar "Searching ($Action)..."

        Start-AsyncJob -Name "Search ($Action)" -Form $form -ScriptBlock {
            param($Mailboxes, $SearchQuery, $Action, $TargetMailbox, $TargetFolder)
            $params = @{
                Mailboxes   = $Mailboxes
                SearchQuery = $SearchQuery
                Action      = $Action
            }
            if ($Action -in @('LogOnly','CopyToFolder') -and $TargetMailbox) {
                $params['TargetMailbox'] = $TargetMailbox
                $params['TargetFolder'] = $TargetFolder
            }
            if ($Action -eq 'DeleteContent') {
                $params['Force'] = $true
            }
            return @(Invoke-MailboxSearch @params)
        } -Parameters @{
            Mailboxes     = $mailboxes
            SearchQuery   = $query
            Action        = $Action
            TargetMailbox = $targetMbx
            TargetFolder  = $targetFld
        } -OnComplete {
            param($result)
            try {
                $script:LastSearchResults = @($result)
                Set-DGVData -DGV $dgvSearchResults -Data $script:LastSearchResults
                $totalItems = ($result | Measure-Object -Property ResultItems -Sum).Sum
                $lblSearchCount.Text = "$($result.Count) mailbox(es), $totalItems item(s) found"
                Update-StatusBar "Search complete: $totalItems items in $($result.Count) mailboxes"
                try {
                    Write-SearchLog -Action $Action -SearchQuery $query -Scope ($mailboxes -join ',') -Result "$totalItems items"
                    Add-SearchHistory -Query $query -Scope ($mailboxes -join ',') -Action $Action
                    Write-OperatorLog -Action "Search-$Action" -Target ($mailboxes -join ',') -Details "Query=$query Items=$totalItems"
                } catch {}
            } catch {
                Update-StatusBar "Search UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Search error: $err"
        }
    }

    $btnEstimate.Add_Click({ & $doSearch 'Estimate' })
    $btnSearchLog.Add_Click({ & $doSearch 'LogOnly' })
    $btnSearchCopy.Add_Click({ & $doSearch 'CopyToFolder' })
    $btnSearchDelete.Add_Click({ & $doSearch 'DeleteContent' })
    $btnSearchExport.Add_Click({ Show-Export -Data $script:LastSearchResults -DefaultName 'search-results' })

    $btnCheckPerms.Add_Click({
        if (-not (& $requireConnection)) { return }
        $cmdText = "Get-ManagementRoleAssignment -Role `"Mailbox Search`" -GetEffectiveUsers | Where { `$_.EffectiveUserName -eq `$env:USERNAME }`r`nGet-ManagementRoleAssignment -Role `"Discovery Management`" -GetEffectiveUsers | Where { `$_.EffectiveUserName -eq `$env:USERNAME }`r`nGet-Mailbox -RecipientTypeDetails DiscoveryMailbox"
        if ((& $showCommandPreview -CommandText $cmdText -Title 'Check Permissions') -eq 'Cancel') { return }
        Update-StatusBar 'Checking permissions...'
        Start-AsyncJob -Name 'CheckPermissions' -Form $form -ScriptBlock {
            $perms = Test-SearchPermissions
            $discovery = @(Get-DiscoveryMailbox)
            return @{ Permissions = $perms; Discovery = $discovery }
        } -OnComplete {
            param($result)
            try {
                $p = $result.Permissions
                $d = $result.Discovery
                $sb = [System.Text.StringBuilder]::new()
                [void]$sb.AppendLine("RBAC Permissions for: $($p.User)")
                [void]$sb.AppendLine("=" * 50)
                [void]$sb.AppendLine("Mailbox Search      : $(if ($p.MailboxSearch) { 'YES' } else { 'NO' })")
                [void]$sb.AppendLine("Import/Export        : $(if ($p.MailboxImportExport) { 'YES' } else { 'NO' })")
                [void]$sb.AppendLine("Discovery Management : $(if ($p.DiscoveryManagement) { 'YES' } else { 'NO' })")
                [void]$sb.AppendLine("Total Roles          : $($p.TotalRoles)")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("Discovery Mailboxes:")
                foreach ($dm in $d) {
                    [void]$sb.AppendLine("  $($dm.DisplayName) <$($dm.PrimarySmtp)>")
                }
                if ($d.Count -gt 0) {
                    $txtTarget.Text = "$($d[0].PrimarySmtp)"
                }
                [System.Windows.Forms.MessageBox]::Show($sb.ToString(), 'Permissions Check', 'OK', 'Information')
                Update-StatusBar 'Permissions checked'
            } catch { Update-StatusBar "Permissions UI error: $_" }
        } -OnError {
            param($err)
            Update-StatusBar "Permissions error: $err"
        }
    })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 2: ORG-WIDE SEARCH & DELETE
    # ═══════════════════════════════════════════════════════════════════════════
    $tabOrgWide = New-Object System.Windows.Forms.TabPage
    $tabOrgWide.Text = 'Org-Wide Delete'

    $orgPanel = New-Object System.Windows.Forms.Panel
    $orgPanel.Dock = 'Fill'

    $orgBar1 = New-FlowBar -H 38
    $lblOrgSubject = New-InlineLabel -Text 'Subject:'
    $txtOrgSubject = New-Object System.Windows.Forms.TextBox
    $txtOrgSubject.Width = 250; $txtOrgSubject.Height = 24
    $txtOrgSubject.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblOrgFrom = New-InlineLabel -Text 'From:'
    $txtOrgFrom = New-Object System.Windows.Forms.TextBox
    $txtOrgFrom.Width = 200; $txtOrgFrom.Height = 24
    $txtOrgFrom.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblOrgKeywords = New-InlineLabel -Text 'Keywords:'
    $txtOrgKeywords = New-Object System.Windows.Forms.TextBox
    $txtOrgKeywords.Width = 200; $txtOrgKeywords.Height = 24
    $txtOrgKeywords.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $orgBar1.Controls.AddRange(@($lblOrgSubject, $txtOrgSubject, $lblOrgFrom, $txtOrgFrom, $lblOrgKeywords, $txtOrgKeywords))

    $orgBar2 = New-FlowBar -H 38
    $lblOrgDateFrom = New-InlineLabel -Text 'From:'
    $dtpOrgFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtpOrgFrom.Format = 'Custom'; $dtpOrgFrom.CustomFormat = 'yyyy-MM-dd'
    $dtpOrgFrom.Width = 110; $dtpOrgFrom.Value = (Get-Date).AddDays(-1)
    $dtpOrgFrom.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblOrgDateTo = New-InlineLabel -Text 'To:'
    $dtpOrgTo = New-Object System.Windows.Forms.DateTimePicker
    $dtpOrgTo.Format = 'Custom'; $dtpOrgTo.CustomFormat = 'yyyy-MM-dd'
    $dtpOrgTo.Width = 110
    $dtpOrgTo.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblOrgBatch = New-InlineLabel -Text 'Batch:'
    $nudOrgBatch = New-Object System.Windows.Forms.NumericUpDown
    $nudOrgBatch.Width = 60; $nudOrgBatch.Minimum = 10; $nudOrgBatch.Maximum = 200; $nudOrgBatch.Value = 50
    $nudOrgBatch.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $btnOrgEstimate = New-Btn -Text 'Estimate All' -W 110 -Color 'Orange'
    $btnOrgDelete = New-Btn -Text 'DELETE FROM ALL' -W 140 -Color 'Red'
    $lblOrgQuery = New-InlineLabel -Text '' -MarginLeft 10
    $orgBar2.Controls.AddRange(@($lblOrgDateFrom, $dtpOrgFrom, $lblOrgDateTo, $dtpOrgTo, $lblOrgBatch, $nudOrgBatch, $btnOrgEstimate, $btnOrgDelete, $lblOrgQuery))

    # Warning label
    $lblOrgWarning = New-Object System.Windows.Forms.Label
    $lblOrgWarning.Dock = 'Top'
    $lblOrgWarning.Height = 30
    $lblOrgWarning.Text = '  WARNING: Org-wide delete searches ALL user mailboxes. Use precise filters. Actions are logged.'
    $lblOrgWarning.BackColor = [System.Drawing.Color]::FromArgb(255,240,200)
    $lblOrgWarning.ForeColor = [System.Drawing.Color]::FromArgb(150,80,0)
    $lblOrgWarning.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lblOrgWarning.TextAlign = 'MiddleLeft'

    $orgSplit = New-Object System.Windows.Forms.SplitContainer
    $orgSplit.Dock = 'Fill'
    $orgSplit.Orientation = 'Horizontal'
    $orgSplit.SplitterDistance = 150

    # Summary panel
    $txtOrgSummary = New-ConsoleTextBox
    $orgSplit.Panel1.Controls.Add($txtOrgSummary)

    # Details grid
    $dgvOrgResults = New-StyledDGV
    $dgvOrgResults.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $itemCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'ResultItems') { $itemCol = $c.Index; break } }
            if ($null -ne $itemCol) {
                $val = "$($row.Cells[$itemCol].Value)"
                $num = 0
                if ([int]::TryParse($val, [ref]$num) -and $num -gt 0) {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,220,220)
                }
            }
        } catch {}
    })
    # --- Double-click: preview messages in org-wide results ---
    $dgvOrgResults.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $row = $s.Rows[$e.RowIndex]
        $mbx = "$($row.Cells['Mailbox'].Value)"
        $query = "$($row.Cells['SearchQuery'].Value)"
        $count = "$($row.Cells['ResultItems'].Value)"
        if (-not $mbx -or -not $query -or $count -eq '0') { return }

        $serverFqdn = $txtServer.Text.Trim()
        Update-StatusBar "Loading message preview for $mbx..."

        Start-AsyncJob -Name "Preview $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $SearchQuery, $Server)
            return @(Get-MailboxMessagePreview -Mailbox $Mailbox -SearchQuery $SearchQuery -Server $Server -MaxResults 200)
        } -Parameters @{ Mailbox = $mbx; SearchQuery = $query; Server = $serverFqdn } -OnComplete {
            param($result)
            try {
                $messages = @($result)
                if ($messages.Count -eq 0) {
                    Update-StatusBar "No messages returned from EWS preview"
                    return
                }

                $dlg = New-Object System.Windows.Forms.Form
                $dlg.Text = "Message Preview: $mbx ($($messages.Count) messages)"
                $dlg.Size = New-Object System.Drawing.Size(1100, 600)
                $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
                $dlg.StartPosition = 'CenterParent'
                $dlg.MinimumSize = New-Object System.Drawing.Size(800, 400)

                $grid = New-StyledDGV -Multi
                $grid.Dock = 'Fill'
                Set-DGVData -DGV $grid -Data $messages
                foreach ($col in $grid.Columns) {
                    switch ($col.Name) {
                        'Subject'   { $col.Width = 350 }
                        'From'      { $col.Width = 180 }
                        'To'        { $col.Width = 180 }
                        'Received'  { $col.Width = 150 }
                        'SizeKB'    { $col.Width = 70; $col.HeaderText = 'Size (KB)' }
                        'HasAttach' { $col.Width = 50; $col.HeaderText = 'Attach' }
                        'Importance'{ $col.Width = 70 }
                        'ItemClass' { $col.Width = 120; $col.HeaderText = 'Type' }
                    }
                }

                $bottomBar = New-FlowBar -H 38
                $bottomBar.Dock = 'Bottom'
                $btnExportPreview = New-Btn -Text 'Export...' -W 80
                $lblPreviewCount = New-InlineLabel -Text "$($messages.Count) messages" -MarginLeft 10
                $btnExportPreview.Add_Click({ Show-Export -Data $messages -DefaultName 'message-preview' })
                $bottomBar.Controls.AddRange(@($btnExportPreview, $lblPreviewCount))

                $dlg.Controls.Add($grid)
                $dlg.Controls.Add($bottomBar)
                $grid.BringToFront()

                Update-StatusBar "Preview: $($messages.Count) messages in $mbx"
                [void]$dlg.ShowDialog($form)
                $dlg.Dispose()
            } catch {
                Update-StatusBar "Preview display error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Preview error: $err"
            $fixMsg = if ($err -match '(?i)denied|impersonat|403|401|unauthorized') {
                "`n`nFix: Run this in Exchange Management Shell:`nNew-ManagementRoleAssignment -Name `"EWSImpersonation`" -Role `"ApplicationImpersonation`" -User $env:USERNAME"
            } else { '' }
            [System.Windows.Forms.MessageBox]::Show(
                "Could not load message preview:`n$err$fixMsg",
                'Preview Error', 'OK', 'Warning')
        }
    })
    $orgSplit.Panel2.Controls.Add($dgvOrgResults)

    $orgBottom = New-FlowBar -H 38
    $orgBottom.Dock = 'Bottom'
    $btnOrgExport = New-Btn -Text 'Export...' -W 80
    $orgBottom.Controls.Add($btnOrgExport)

    $orgPanel.Controls.Add($orgSplit)
    $orgPanel.Controls.Add($orgBottom)
    $orgPanel.Controls.Add($lblOrgWarning)
    $orgPanel.Controls.Add($orgBar2)
    $orgPanel.Controls.Add($orgBar1)
    $tabOrgWide.Controls.Add($orgPanel)

    $buildOrgQuery = {
        $params = @{}
        if ($txtOrgSubject.Text)  { $params['Subject'] = $txtOrgSubject.Text }
        if ($txtOrgFrom.Text)     { $params['From'] = $txtOrgFrom.Text }
        if ($txtOrgKeywords.Text) { $params['Keywords'] = $txtOrgKeywords.Text }
        $params['StartDate'] = $dtpOrgFrom.Value
        $params['EndDate']   = $dtpOrgTo.Value
        return Build-SearchQuery @params
    }

    foreach ($ctrl in @($txtOrgSubject, $txtOrgFrom, $txtOrgKeywords)) {
        $ctrl.Add_TextChanged({
            try { $lblOrgQuery.Text = "Query: $(& $buildOrgQuery)" } catch {}
        })
    }

    $doOrgSearch = {
        param([switch]$Delete)
        if (-not (& $requireConnection)) { return }
        $query = & $buildOrgQuery
        if ($query -eq '*') {
            [System.Windows.Forms.MessageBox]::Show('Wildcard search on all mailboxes is not allowed. Add filters.','Safety Check','OK','Error')
            return
        }

        $batchSize = [int]$nudOrgBatch.Value
        $queryEsc = $query -replace '"', '`"'
        $cmdText = "Get-Mailbox -ResultSize Unlimited | Search-Mailbox -SearchQuery `"$queryEsc`""
        if ($Delete) {
            $cmdText += " -DeleteContent -Force"
        } else {
            $cmdText += " -EstimateResultOnly"
        }
        $cmdText += "`r`n# BatchSize: $batchSize mailboxes per batch"

        if ($Delete) {
            $result = & $confirmDelete -Message "DELETE from ALL MAILBOXES matching:`n$query`n`nThis searches every mailbox in the organization. This action is IRREVERSIBLE!" -CommandText $cmdText
            if ($result -ne 'OK') { return }
        } else {
            if ((& $showCommandPreview -CommandText $cmdText -Title 'Org-Wide Estimate') -eq 'Cancel') { return }
        }

        $whatIf = -not $Delete
        Update-StatusBar "Org-wide $(if ($Delete) { 'DELETE' } else { 'estimate' })..."

        Start-AsyncJob -Name "OrgWide $(if ($Delete) { 'DELETE' } else { 'Estimate' })" -Form $form -ScriptBlock {
            param($SearchQuery, $WhatIf, $BatchSize)
            return Remove-MessageFromOrganization -SearchQuery $SearchQuery -WhatIf:$WhatIf -BatchSize $BatchSize
        } -Parameters @{ SearchQuery = $query; WhatIf = $whatIf; BatchSize = $batchSize } -OnComplete {
            param($result)
            try {
                $summary = $result.Summary
                $details = @($result.Details)
                $script:LastOrgResults = $details

                $sb = [System.Text.StringBuilder]::new()
                [void]$sb.AppendLine("=" * 60)
                [void]$sb.AppendLine("ORG-WIDE SEARCH RESULTS")
                [void]$sb.AppendLine("=" * 60)
                [void]$sb.AppendLine("Action          : $($summary.Action)")
                [void]$sb.AppendLine("Query           : $($summary.SearchQuery)")
                [void]$sb.AppendLine("Total Mailboxes : $($summary.TotalMailboxes)")
                [void]$sb.AppendLine("Affected        : $($summary.AffectedMailboxes)")
                [void]$sb.AppendLine("Total Items     : $($summary.TotalItems)")
                [void]$sb.AppendLine("Timestamp       : $($summary.Timestamp)")
                $txtOrgSummary.Text = $sb.ToString()

                $affectedOnly = $details | Where-Object { $_.ResultItems -gt 0 }
                Set-DGVData -DGV $dgvOrgResults -Data @($affectedOnly)

                Update-StatusBar "Org-wide: $($summary.TotalItems) items in $($summary.AffectedMailboxes) mailboxes"
                try {
                    Write-OperatorLog -Action "OrgWide-$($summary.Action)" -Target 'AllMailboxes' `
                        -Details "Query=$query Items=$($summary.TotalItems) Affected=$($summary.AffectedMailboxes)"
                } catch {}
            } catch {
                Update-StatusBar "Org-wide UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Org-wide error: $err"
            $txtOrgSummary.Text = "Error: $err"
        }
    }

    $btnOrgEstimate.Add_Click({ & $doOrgSearch })
    $btnOrgDelete.Add_Click({ & $doOrgSearch -Delete })
    $btnOrgExport.Add_Click({ Show-Export -Data $script:LastOrgResults -DefaultName 'org-wide-results' })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 3: COMPLIANCE SEARCH (In-Place eDiscovery)
    # ═══════════════════════════════════════════════════════════════════════════
    $tabCompliance = New-Object System.Windows.Forms.TabPage
    $tabCompliance.Text = 'eDiscovery'

    $compPanel = New-Object System.Windows.Forms.Panel
    $compPanel.Dock = 'Fill'

    $compBar1 = New-FlowBar -H 38
    $lblCompName = New-InlineLabel -Text 'Search Name:'
    $txtCompName = New-Object System.Windows.Forms.TextBox
    $txtCompName.Width = 200; $txtCompName.Height = 24
    $txtCompName.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblCompQuery = New-InlineLabel -Text 'Query:'
    $txtCompQuery = New-Object System.Windows.Forms.TextBox
    $txtCompQuery.Width = 300; $txtCompQuery.Height = 24
    $txtCompQuery.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $chkCompAll = New-Object System.Windows.Forms.CheckBox
    $chkCompAll.Text = 'All Mailboxes'
    $chkCompAll.AutoSize = $true
    $chkCompAll.Margin = New-Object System.Windows.Forms.Padding(8,8,3,4)
    $compBar1.Controls.AddRange(@($lblCompName, $txtCompName, $lblCompQuery, $txtCompQuery, $chkCompAll))

    $compBar2 = New-FlowBar -H 38
    $lblCompMailboxes = New-InlineLabel -Text 'Mailboxes (comma-sep):'
    $txtCompMailboxes = New-Object System.Windows.Forms.TextBox
    $txtCompMailboxes.Width = 350; $txtCompMailboxes.Height = 24
    $txtCompMailboxes.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $chkCompEstimate = New-Object System.Windows.Forms.CheckBox
    $chkCompEstimate.Text = 'Estimate Only'
    $chkCompEstimate.Checked = $true
    $chkCompEstimate.AutoSize = $true
    $chkCompEstimate.Margin = New-Object System.Windows.Forms.Padding(8,8,8,4)
    $btnCompCreate = New-Btn -Text 'Create Search' -W 110 -Color 'Blue'
    $btnCompRefresh = New-Btn -Text 'Refresh List' -W 100
    $compBar2.Controls.AddRange(@($lblCompMailboxes, $txtCompMailboxes, $chkCompEstimate, $btnCompCreate, $btnCompRefresh))

    $compSplit = New-Object System.Windows.Forms.SplitContainer
    $compSplit.Dock = 'Fill'
    $compSplit.Orientation = 'Horizontal'
    $compSplit.SplitterDistance = 350

    $dgvCompliance = New-StyledDGV
    $dgvCompliance.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $statusCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Status') { $statusCol = $c.Index; break } }
            if ($null -ne $statusCol) {
                $val = "$($row.Cells[$statusCol].Value)"
                if ($val -match 'Succeeded|Completed')     { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,255,200) }
                elseif ($val -match 'InProgress|Started') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,230,255) }
                elseif ($val -match 'Failed|Error')       { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,200,200) }
            }
        } catch {}
    })

    # Context menu for compliance searches
    $ctxComp = New-Object System.Windows.Forms.ContextMenuStrip
    $ctxCompStatus = $ctxComp.Items.Add('Refresh Status')
    $ctxCompStop = $ctxComp.Items.Add('Stop Search')
    $ctxCompRemove = $ctxComp.Items.Add('Remove Search')
    $ctxCompCopyName = $ctxComp.Items.Add('Copy Name')
    $dgvCompliance.ContextMenuStrip = $ctxComp

    $compSplit.Panel1.Controls.Add($dgvCompliance)

    # Status detail panel
    $txtCompDetail = New-ConsoleTextBox
    $compSplit.Panel2.Controls.Add($txtCompDetail)

    $compPanel.Controls.Add($compSplit)
    $compPanel.Controls.Add($compBar2)
    $compPanel.Controls.Add($compBar1)
    $tabCompliance.Controls.Add($compPanel)

    # --- Compliance actions ---
    $refreshCompliance = {
        if (-not (& $requireConnection)) { return }
        if ((& $showCommandPreview -CommandText 'Get-MailboxSearch' -Title 'eDiscovery List') -eq 'Cancel') { return }
        Update-StatusBar 'Loading eDiscovery searches...'
        Start-AsyncJob -Name 'eDiscovery List' -Form $form -ScriptBlock {
            return @(Get-ContentSearches)
        } -OnComplete {
            param($result)
            try {
                $script:LastComplianceSearches = @($result)
                Set-DGVData -DGV $dgvCompliance -Data $script:LastComplianceSearches
                Update-StatusBar "eDiscovery: $($result.Count) searches"
            } catch {
                Update-StatusBar "eDiscovery UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "eDiscovery error: $err"
        }
    }
    $btnCompRefresh.Add_Click($refreshCompliance)

    $btnCompCreate.Add_Click({
        if (-not (& $requireConnection)) { return }
        try {
            $name = $txtCompName.Text.Trim()
            $query = $txtCompQuery.Text.Trim()
            if (-not $name -or -not $query) {
                [System.Windows.Forms.MessageBox]::Show('Name and Query are required.','Create Search','OK','Warning')
                return
            }
            $allMbx = $chkCompAll.Checked
            $mbxList = @()
            if (-not $allMbx -and $txtCompMailboxes.Text) {
                $mbxList = @($txtCompMailboxes.Text -split '[,;\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
            $estimateOnly = $chkCompEstimate.Checked
            $queryEsc = $query -replace '"', '`"'
            $cmdText = "New-MailboxSearch -Name `"$name`" -SearchQuery `"$queryEsc`""
            if ($allMbx) { $cmdText += " -AllMailboxes" }
            elseif ($mbxList.Count -gt 0) { $cmdText += " -SourceMailboxes `"$($mbxList -join '","')`"" }
            if ($estimateOnly) { $cmdText += " -EstimateOnly" }
            $cmdText += "`r`nStart-MailboxSearch -Identity `"$name`""
            if ((& $showCommandPreview -CommandText $cmdText -Title 'Create eDiscovery Search') -eq 'Cancel') { return }
            Update-StatusBar "Creating eDiscovery search '$name'..."

            Start-AsyncJob -Name "Create eDiscovery: $name" -Form $form -ScriptBlock {
                param($Name, $SearchQuery, $AllMailboxes, $SourceMailboxes, $EstimateOnly)
                $params = @{ Name = $Name; SearchQuery = $SearchQuery }
                if ($AllMailboxes) { $params['AllMailboxes'] = $true }
                elseif ($SourceMailboxes.Count -gt 0) { $params['SourceMailboxes'] = $SourceMailboxes }
                if ($EstimateOnly) { $params['EstimateOnly'] = $true }
                return New-ContentSearch @params
            } -Parameters @{
                Name = $name; SearchQuery = $query
                AllMailboxes = $allMbx; SourceMailboxes = $mbxList
                EstimateOnly = $estimateOnly
            } -OnComplete {
                param($result)
                Update-StatusBar "eDiscovery search '$name' created"
                try { Write-OperatorLog -Action 'CreateSearch' -Target $name -Details "Query=$query" } catch {}
                & $refreshCompliance
            } -OnError {
                param($err)
                Update-StatusBar "Create search error: $err"
            }
        } catch { Update-StatusBar "Create search error: $_" }
    })

    # Double-click -> show status detail
    $dgvCompliance.Add_CellDoubleClick({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $searchName = "$($dgvCompliance.Rows[$e.RowIndex].Cells['Name'].Value)"
            if (-not $searchName) { return }
            Update-StatusBar "Getting status for '$searchName'..."
            Start-AsyncJob -Name "Status: $searchName" -Form $form -ScriptBlock {
                param($Name)
                return Get-ContentSearchStatus -Name $Name
            } -Parameters @{ Name = $searchName } -OnComplete {
                param($result)
                try {
                    $sb = [System.Text.StringBuilder]::new()
                    foreach ($prop in $result.PSObject.Properties) {
                        [void]$sb.AppendLine("$($prop.Name): $($prop.Value)")
                    }
                    $txtCompDetail.Text = $sb.ToString()
                    Update-StatusBar "Status loaded for '$searchName'"
                } catch {}
            } -OnError {
                param($err)
                $txtCompDetail.Text = "Error: $err"
            }
        } catch {}
    })

    $ctxCompStatus.Add_Click({
        if ($dgvCompliance.SelectedRows.Count -gt 0) {
            $dgvCompliance.GetType().GetMethod('OnCellDoubleClick', [System.Reflection.BindingFlags]'NonPublic,Instance').Invoke(
                $dgvCompliance, @([System.Windows.Forms.DataGridViewCellEventArgs]::new(0, $dgvCompliance.SelectedRows[0].Index))
            )
        }
    })

    $ctxCompStop.Add_Click({
        try {
            if ($dgvCompliance.SelectedRows.Count -eq 0) { return }
            $name = "$($dgvCompliance.SelectedRows[0].Cells['Name'].Value)"
            Start-AsyncJob -Name "Stop: $name" -Form $form -ScriptBlock {
                param($Name)
                Stop-ContentSearch -Name $Name
            } -Parameters @{ Name = $name } -OnComplete {
                param($r)
                Update-StatusBar "Search '$name' stopped"
                try { Write-OperatorLog -Action 'StopSearch' -Target $name } catch {}
                & $refreshCompliance
            } -OnError { param($err) Update-StatusBar "Stop error: $err" }
        } catch {}
    })

    $ctxCompRemove.Add_Click({
        try {
            if ($dgvCompliance.SelectedRows.Count -eq 0) { return }
            $name = "$($dgvCompliance.SelectedRows[0].Cells['Name'].Value)"
            $confirm = [System.Windows.Forms.MessageBox]::Show("Remove search '$name'?", 'Confirm', 'YesNo', 'Question')
            if ($confirm -eq 'Yes') {
                Start-AsyncJob -Name "Remove: $name" -Form $form -ScriptBlock {
                    param($Name)
                    Remove-ContentSearch -Name $Name
                } -Parameters @{ Name = $name } -OnComplete {
                    param($r)
                    Update-StatusBar "Search '$name' removed"
                    try { Write-OperatorLog -Action 'RemoveSearch' -Target $name } catch {}
                    & $refreshCompliance
                } -OnError { param($err) Update-StatusBar "Remove error: $err" }
            }
        } catch {}
    })

    $ctxCompCopyName.Add_Click({
        try {
            if ($dgvCompliance.SelectedRows.Count -gt 0) {
                $v = "$($dgvCompliance.SelectedRows[0].Cells['Name'].Value)"
                [System.Windows.Forms.Clipboard]::SetText($v)
            }
        } catch {}
    })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 4: MAILBOX BROWSER
    # ═══════════════════════════════════════════════════════════════════════════
    $tabMailboxes = New-Object System.Windows.Forms.TabPage
    $tabMailboxes.Text = 'Mailboxes'

    $mbxPanel = New-Object System.Windows.Forms.Panel
    $mbxPanel.Dock = 'Fill'

    $mbxBar = New-FlowBar -H 38
    $lblMbxFilter = New-InlineLabel -Text 'Filter:'
    $txtMbxFilter = New-Object System.Windows.Forms.TextBox
    $txtMbxFilter.Width = 200; $txtMbxFilter.Height = 24
    $txtMbxFilter.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblMbxType = New-InlineLabel -Text 'Type:'
    $cmbMbxType = New-Object System.Windows.Forms.ComboBox
    $cmbMbxType.DropDownStyle = 'DropDownList'; $cmbMbxType.Width = 130
    $cmbMbxType.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    @('All','UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox') | ForEach-Object { [void]$cmbMbxType.Items.Add($_) }
    $cmbMbxType.SelectedIndex = 0
    $btnMbxLoad = New-Btn -Text 'Load Mailboxes' -W 120 -Color 'Blue'
    $btnMbxStats = New-Btn -Text 'Get Stats' -W 90
    $btnMbxFolders = New-Btn -Text 'Folder Stats' -W 100
    $btnMbxUseInSearch = New-Btn -Text 'Use in Search' -W 110 -Color 'Green'
    $btnMbxExport = New-Btn -Text 'Export...' -W 80
    $mbxBar.Controls.AddRange(@($lblMbxFilter, $txtMbxFilter, $lblMbxType, $cmbMbxType, $btnMbxLoad, $btnMbxStats, $btnMbxFolders, $btnMbxUseInSearch, $btnMbxExport))

    $mbxSplit = New-Object System.Windows.Forms.SplitContainer
    $mbxSplit.Dock = 'Fill'
    $mbxSplit.Orientation = 'Horizontal'
    $mbxSplit.SplitterDistance = 400

    $dgvMailboxes = New-StyledDGV -Multi
    $mbxSplit.Panel1.Controls.Add($dgvMailboxes)

    $dgvMbxDetail = New-StyledDGV
    $mbxSplit.Panel2.Controls.Add($dgvMbxDetail)

    $mbxPanel.Controls.Add($mbxSplit)
    $mbxPanel.Controls.Add($mbxBar)
    $tabMailboxes.Controls.Add($mbxPanel)

    $btnMbxLoad.Add_Click({
        if (-not (& $requireConnection)) { return }
        $filterVal = $txtMbxFilter.Text
        $typeVal = $cmbMbxType.SelectedItem.ToString()
        $cmdText = "Get-Mailbox -ResultSize 500"
        if ($filterVal) { $cmdText += " -Filter `"DisplayName -like '*$filterVal*' -or PrimarySmtpAddress -like '*$filterVal*'`"" }
        if ($typeVal -ne 'All') { $cmdText += " -RecipientTypeDetails $typeVal" }
        if ((& $showCommandPreview -CommandText $cmdText -Title 'Load Mailboxes') -eq 'Cancel') { return }
        Update-StatusBar 'Loading mailboxes...'
        Start-AsyncJob -Name 'LoadMailboxes' -Form $form -ScriptBlock {
            param($Filter, $RecipientType)
            $params = @{ RecipientType = $RecipientType }
            if ($Filter) { $params['Filter'] = "DisplayName -like '*$Filter*' -or PrimarySmtpAddress -like '*$Filter*'" }
            return @(Get-SearchableMailboxes @params)
        } -Parameters @{ Filter = $filterVal; RecipientType = $typeVal } -OnComplete {
            param($result)
            try {
                $script:LastMailboxList = @($result)
                Set-DGVData -DGV $dgvMailboxes -Data $script:LastMailboxList
                Update-StatusBar "Mailboxes: $($result.Count) loaded"
            } catch { Update-StatusBar "Mailbox UI error: $_" }
        } -OnError {
            param($err)
            Update-StatusBar "Mailbox error: $err"
        }
    })

    $btnMbxStats.Add_Click({
        try {
            $selected = @()
            foreach ($row in $dgvMailboxes.SelectedRows) {
                $smtp = "$($row.Cells['PrimarySmtp'].Value)"
                if ($smtp) { $selected += $smtp }
            }
            if ($selected.Count -eq 0) { return }
            $cmdText = ($selected | ForEach-Object { "Get-MailboxStatistics -Identity `"$_`"" }) -join "`r`n"
            if ((& $showCommandPreview -CommandText $cmdText -Title 'Mailbox Statistics') -eq 'Cancel') { return }
            Update-StatusBar "Getting stats for $($selected.Count) mailbox(es)..."
            Start-AsyncJob -Name 'MailboxStats' -Form $form -ScriptBlock {
                param($Mailboxes)
                return @(Get-MailboxQuickStats -Mailboxes $Mailboxes)
            } -Parameters @{ Mailboxes = $selected } -OnComplete {
                param($result)
                try {
                    $script:LastStatsData = @($result)
                    Set-DGVData -DGV $dgvMbxDetail -Data $script:LastStatsData
                    Update-StatusBar "Stats loaded for $($result.Count) mailboxes"
                } catch { Update-StatusBar "Stats UI error: $_" }
            } -OnError { param($err) Update-StatusBar "Stats error: $err" }
        } catch {}
    })

    $btnMbxFolders.Add_Click({
        try {
            if ($dgvMailboxes.SelectedRows.Count -eq 0) { return }
            $smtp = "$($dgvMailboxes.SelectedRows[0].Cells['PrimarySmtp'].Value)"
            if (-not $smtp) { return }
            if ((& $showCommandPreview -CommandText "Get-MailboxFolderStatistics -Identity `"$smtp`"" -Title 'Folder Statistics') -eq 'Cancel') { return }
            Update-StatusBar "Loading folder stats for $smtp..."
            Start-AsyncJob -Name "Folders: $smtp" -Form $form -ScriptBlock {
                param($Mailbox)
                return @(Get-MailboxFolderStats -Mailbox $Mailbox)
            } -Parameters @{ Mailbox = $smtp } -OnComplete {
                param($result)
                try {
                    $script:LastFolderStats = @($result)
                    Set-DGVData -DGV $dgvMbxDetail -Data $script:LastFolderStats
                    Update-StatusBar "Folder stats: $($result.Count) folders"
                } catch { Update-StatusBar "Folder UI error: $_" }
            } -OnError { param($err) Update-StatusBar "Folder error: $err" }
        } catch {}
    })

    $btnMbxUseInSearch.Add_Click({
        try {
            $selected = @()
            foreach ($row in $dgvMailboxes.SelectedRows) {
                $smtp = "$($row.Cells['PrimarySmtp'].Value)"
                if ($smtp) { $selected += $smtp }
            }
            if ($selected.Count -eq 0) { return }
            $txtScope.Text = ($selected -join ', ')
            $txtScope.ForeColor = [System.Drawing.Color]::Black
            $tabs.SelectedTab = $tabSearch
            Update-StatusBar "Set scope to $($selected.Count) mailbox(es)"
        } catch {}
    })

    $btnMbxExport.Add_Click({ Show-Export -Data $script:LastMailboxList -DefaultName 'mailboxes' })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 5: AUDIT LOG
    # ═══════════════════════════════════════════════════════════════════════════
    $tabAudit = New-Object System.Windows.Forms.TabPage
    $tabAudit.Text = 'Audit Log'

    $auditPanel = New-Object System.Windows.Forms.Panel
    $auditPanel.Dock = 'Fill'

    $auditBar = New-FlowBar -H 38
    $btnAuditRefresh = New-Btn -Text 'Refresh' -W 90
    $btnAuditSearchLog = New-Btn -Text 'Search Log' -W 100
    $btnAuditOperLog = New-Btn -Text 'Operator Log' -W 110
    $btnAuditExport = New-Btn -Text 'Export...' -W 80
    $auditBar.Controls.AddRange(@($btnAuditRefresh, $btnAuditSearchLog, $btnAuditOperLog, $btnAuditExport))

    $dgvAudit = New-StyledDGV
    $dgvAudit.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $actionCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Action') { $actionCol = $c.Index; break } }
            if ($null -ne $actionCol) {
                $val = "$($row.Cells[$actionCol].Value)"
                if ($val -match 'Delete')  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,200,200) }
                elseif ($val -match 'Create|Search') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,230,255) }
            }
        } catch {}
    })

    $auditPanel.Controls.Add($dgvAudit)
    $auditPanel.Controls.Add($auditBar)
    $tabAudit.Controls.Add($auditPanel)

    $btnAuditSearchLog.Add_Click({
        try {
            $data = @(Get-SearchLog -Last 200)
            $script:LastAuditLog = $data
            Set-DGVData -DGV $dgvAudit -Data $data
            Update-StatusBar "Search log: $($data.Count) entries"
        } catch { Update-StatusBar "Search log error: $_" }
    })

    $btnAuditOperLog.Add_Click({
        try {
            $data = @(Get-OperatorLog -Last 200)
            $script:LastAuditLog = $data
            Set-DGVData -DGV $dgvAudit -Data $data
            Update-StatusBar "Operator log: $($data.Count) entries"
        } catch { Update-StatusBar "Operator log error: $_" }
    })

    $btnAuditRefresh.Add_Click({
        $btnAuditOperLog.PerformClick()
    })

    $btnAuditExport.Add_Click({ Show-Export -Data $script:LastAuditLog -DefaultName 'audit-log' })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 6: FOLDER CLEANUP & DUPLICATES
    # ═══════════════════════════════════════════════════════════════════════════
    $tabFolder = New-Object System.Windows.Forms.TabPage
    $tabFolder.Text = 'Folder Cleanup'

    $folderPanel = New-Object System.Windows.Forms.Panel
    $folderPanel.Dock = 'Fill'

    # --- Row 1: Mailbox + Folder ---
    $folderBar1 = New-FlowBar -H 38
    $lblFcMailbox = New-InlineLabel -Text 'Mailbox:'
    $txtFcMailbox = New-Object System.Windows.Forms.TextBox
    $txtFcMailbox.Width = 220; $txtFcMailbox.Height = 24
    $txtFcMailbox.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblFcFolder = New-InlineLabel -Text 'Folder:'
    $cmbFcFolder = New-Object System.Windows.Forms.ComboBox
    $cmbFcFolder.Width = 220; $cmbFcFolder.DropDownStyle = 'DropDown'
    $cmbFcFolder.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $btnFcLoadFolders = New-Btn -Text 'Load Folders' -W 100 -Color 'Blue'
    $folderBar1.Controls.AddRange(@($lblFcMailbox, $txtFcMailbox, $lblFcFolder, $cmbFcFolder, $btnFcLoadFolders))

    # --- Row 2: Filters ---
    $folderBar2 = New-FlowBar -H 38
    $lblFcSubject = New-InlineLabel -Text 'Subject:'
    $txtFcSubject = New-Object System.Windows.Forms.TextBox
    $txtFcSubject.Width = 160; $txtFcSubject.Height = 24
    $txtFcSubject.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblFcFrom = New-InlineLabel -Text 'From:'
    $txtFcFrom = New-Object System.Windows.Forms.TextBox
    $txtFcFrom.Width = 160; $txtFcFrom.Height = 24
    $txtFcFrom.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblFcOlder = New-InlineLabel -Text 'Older than (days):'
    $numFcDays = New-Object System.Windows.Forms.NumericUpDown
    $numFcDays.Width = 70; $numFcDays.Minimum = 0; $numFcDays.Maximum = 3650; $numFcDays.Value = 0
    $numFcDays.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $lblFcSize = New-InlineLabel -Text 'Size:'
    $cmbFcSize = New-Object System.Windows.Forms.ComboBox
    $cmbFcSize.Width = 100; $cmbFcSize.DropDownStyle = 'DropDownList'
    $cmbFcSize.Items.AddRange(@('Any','Small','Medium','Large','VeryLarge'))
    $cmbFcSize.SelectedIndex = 0
    $cmbFcSize.Margin = New-Object System.Windows.Forms.Padding(3,6,8,4)
    $chkFcAttach = New-Object System.Windows.Forms.CheckBox
    $chkFcAttach.Text = 'Has Attachment'
    $chkFcAttach.AutoSize = $true
    $chkFcAttach.Margin = New-Object System.Windows.Forms.Padding(8,8,4,4)
    $folderBar2.Controls.AddRange(@($lblFcSubject, $txtFcSubject, $lblFcFrom, $txtFcFrom, $lblFcOlder, $numFcDays, $lblFcSize, $cmbFcSize, $chkFcAttach))

    # --- Row 3: Actions ---
    $folderBar3 = New-FlowBar -H 38
    $btnFcEstimate = New-Btn -Text 'Estimate' -W 100 -Color 'Blue'
    $btnFcDelete = New-Btn -Text 'Delete' -W 100 -Color 'Red'
    $btnFcPurge = New-Btn -Text 'Purge Dumpster' -W 120 -Color 'Orange'
    $lblFcSep = New-InlineLabel -Text '|' -MarginLeft 10
    $btnFcFindDupes = New-Btn -Text 'Find Duplicates' -W 120 -Color 'Green'
    $btnFcBackupDupes = New-Btn -Text 'Backup Folder' -W 110
    $btnFcRemoveDupes = New-Btn -Text 'Backup + Delete' -W 120 -Color 'Red'
    $lblFcTarget = New-InlineLabel -Text 'Target:' -MarginLeft 10
    $txtFcTarget = New-Object System.Windows.Forms.TextBox
    $txtFcTarget.Width = 180; $txtFcTarget.Height = 24
    $txtFcTarget.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $btnFcExport = New-Btn -Text 'Export...' -W 80
    $folderBar3.Controls.AddRange(@($btnFcEstimate, $btnFcDelete, $btnFcPurge, $lblFcSep, $btnFcFindDupes, $btnFcBackupDupes, $btnFcRemoveDupes, $lblFcTarget, $txtFcTarget, $btnFcExport))

    $dgvFolder = New-StyledDGV
    $dgvFolder.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $statusCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Status') { $statusCol = $c.Index; break } }
            if ($null -ne $statusCol) {
                $val = "$($row.Cells[$statusCol].Value)"
                if ($val -eq 'PossibleDupes') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,230,200) }
            }
            $actionCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Action') { $actionCol = $c.Index; break } }
            if ($null -ne $actionCol) {
                $val = "$($row.Cells[$actionCol].Value)"
                if ($val -match 'Delete')  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,200,200) }
            }
        } catch {}
    })

    # --- Grid click -> sync Folder dropdown ---
    $dgvFolder.Add_CellClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        try {
            $fpCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'FolderPath') { $fpCol = $c.Index; break } }
            if ($null -eq $fpCol) { return }
            $folderPath = "$($s.Rows[$e.RowIndex].Cells[$fpCol].Value)"
            if ($folderPath -and $cmbFcFolder.Items.Contains($folderPath)) {
                $cmbFcFolder.SelectedItem = $folderPath
            }
        } catch {}
    })

    # --- Folder dropdown -> highlight grid row ---
    $cmbFcFolder.Add_SelectedIndexChanged({
        try {
            $selected = $cmbFcFolder.SelectedItem
            if (-not $selected -or $selected -eq '(All Folders)') {
                $dgvFolder.ClearSelection()
                return
            }
            foreach ($row in $dgvFolder.Rows) {
                $fpCol = $null
                foreach ($c in $dgvFolder.Columns) { if ($c.Name -eq 'FolderPath') { $fpCol = $c.Index; break } }
                if ($null -ne $fpCol -and "$($row.Cells[$fpCol].Value)" -eq $selected) {
                    $dgvFolder.ClearSelection()
                    $row.Selected = $true
                    $dgvFolder.FirstDisplayedScrollingRowIndex = $row.Index
                    break
                }
            }
        } catch {}
    })

    $folderPanel.Controls.Add($dgvFolder)
    $folderPanel.Controls.Add($folderBar3)
    $folderPanel.Controls.Add($folderBar2)
    $folderPanel.Controls.Add($folderBar1)
    $tabFolder.Controls.Add($folderPanel)

    # --- Folder Cleanup Events ---
    $btnFcLoadFolders.Add_Click({
        if (-not (& $requireConnection)) { return }
        $mbx = $txtFcMailbox.Text.Trim()
        if (-not $mbx) {
            [System.Windows.Forms.MessageBox]::Show('Enter a mailbox.','Folder Cleanup','OK','Warning')
            return
        }
        if ((& $showCommandPreview -CommandText "Get-MailboxFolderStatistics -Identity `"$mbx`" | Select FolderPath, ItemsInFolder, FolderSize" -Title 'Load Folders') -eq 'Cancel') { return }
        Update-StatusBar "Loading folders for $mbx..."
        $btnFcLoadFolders.Enabled = $false
        Start-AsyncJob -Name "Folders $mbx" -Form $form -ScriptBlock {
            param($Mailbox)
            Get-MailboxFolderList -Mailbox $Mailbox
        } -Parameters @{ Mailbox = $mbx } -OnComplete {
            param($result)
            try {
                $cmbFcFolder.Items.Clear()
                $cmbFcFolder.Items.Add('(All Folders)')
                foreach ($f in $result) {
                    $cmbFcFolder.Items.Add($f.FolderPath)
                }
                $cmbFcFolder.SelectedIndex = 0
                Set-DGVData -DGV $dgvFolder -Data $result
                $script:LastFolderCleanup = $result
                Update-StatusBar "Loaded $($result.Count) folders for $mbx"
            } catch { Update-StatusBar "Folder load error: $_" }
            $btnFcLoadFolders.Enabled = $true
        } -OnError {
            param($err)
            $btnFcLoadFolders.Enabled = $true
            Update-StatusBar "Folder load error: $err"
        }
    })

    # Helper: build folder cleanup command text for safe mode
    $buildFcCommand = {
        param($Mbx, $Folder, $Subj, $Frm, $Days, $Sz, $Att, $Action)
        $parts = @()
        if ($Subj) { $parts += "subject:`"$Subj`"" }
        if ($Frm)  { $parts += "from:`"$Frm`"" }
        if ($Days -gt 0) { $parts += "received<=$((Get-Date).AddDays(-$Days).ToString('yyyy-MM-dd'))" }
        if ($Att)  { $parts += "hasattachment:true" }
        $q = if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '*' }

        if ($Folder) {
            $cmd = "# EWS FindItem on folder `"$Folder`" in mailbox `"$Mbx`"`r`n"
            $cmd += "# Query: $q`r`n"
            if ($Sz) { $cmd += "# Size filter: $Sz`r`n" }
            $cmd += "# Equivalent Search-Mailbox (all folders):`r`n"
        } else { $cmd = '' }
        $qEsc = $q -replace '"', '`"'
        $cmd += "Search-Mailbox -Identity `"$Mbx`" -SearchQuery `"$qEsc`""
        if ($Action -eq 'Estimate') { $cmd += " -EstimateResultOnly" }
        else { $cmd += " -DeleteContent -Force" }
        return $cmd
    }

    $btnFcEstimate.Add_Click({
        if (-not (& $requireConnection)) { return }
        $mbx = $txtFcMailbox.Text.Trim()
        if (-not $mbx) {
            [System.Windows.Forms.MessageBox]::Show('Enter a mailbox.','Folder Cleanup','OK','Warning')
            return
        }
        $folder = if ($cmbFcFolder.SelectedItem -and $cmbFcFolder.SelectedItem -ne '(All Folders)') { $cmbFcFolder.SelectedItem } else { '' }
        $subj = $txtFcSubject.Text.Trim()
        $frm = $txtFcFrom.Text.Trim()
        $days = [int]$numFcDays.Value
        $sz = if ($cmbFcSize.SelectedItem -and $cmbFcSize.SelectedItem -ne 'Any') { $cmbFcSize.SelectedItem } else { '' }
        $att = $chkFcAttach.Checked
        $cmdText = & $buildFcCommand -Mbx $mbx -Folder $folder -Subj $subj -Frm $frm -Days $days -Sz $sz -Att $att -Action 'Estimate'
        if ((& $showCommandPreview -CommandText $cmdText -Title 'Folder Estimate') -eq 'Cancel') { return }
        Update-StatusBar "Estimating folder cleanup for $mbx..."
        $btnFcEstimate.Enabled = $false

        $srv = $txtServer.Text.Trim()
        Start-AsyncJob -Name "Folder Estimate $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $FolderPath, $Subject, $From, $OlderThanDays, $SizeRange, $HasAttachment, $Server)
            $p = @{ Mailbox = $Mailbox; Action = 'Estimate' }
            if ($FolderPath)    { $p['FolderPath'] = $FolderPath }
            if ($Subject)       { $p['Subject'] = $Subject }
            if ($From)          { $p['From'] = $From }
            if ($OlderThanDays -gt 0) { $p['OlderThanDays'] = $OlderThanDays }
            if ($SizeRange)     { $p['SizeRange'] = $SizeRange }
            if ($HasAttachment) { $p['HasAttachment'] = $true }
            if ($Server)        { $p['Server'] = $Server }
            Invoke-FolderCleanup @p
        } -Parameters @{
            Mailbox = $mbx; FolderPath = $folder; Subject = $subj; From = $frm
            OlderThanDays = $days; SizeRange = $sz; HasAttachment = $att; Server = $srv
        } -OnComplete {
            param($result)
            try {
                $data = @($result)
                $script:LastFolderCleanup = $data
                Set-DGVData -DGV $dgvFolder -Data $data
                $total = ($data | Measure-Object -Property ResultItems -Sum).Sum
                Update-StatusBar "Folder estimate: $total item(s) match"
                try { Write-SearchLog -Action 'FolderEstimate' -Scope $mbx -Result "$total items" } catch {}
            } catch { Update-StatusBar "Folder estimate error: $_" }
            $btnFcEstimate.Enabled = $true
        } -OnError {
            param($err)
            $btnFcEstimate.Enabled = $true
            Update-StatusBar "Folder estimate error: $err"
            if ($err -match '(?i)denied|impersonat|403|401|unauthorized') {
                [System.Windows.Forms.MessageBox]::Show(
                    "EWS access denied.`n`nFix: Run this in Exchange Management Shell:`nNew-ManagementRoleAssignment -Name `"EWSImpersonation`" -Role `"ApplicationImpersonation`" -User $env:USERNAME",
                    'EWS Impersonation Required', 'OK', 'Warning')
            }
        }
    })

    $btnFcDelete.Add_Click({
        if (-not (& $requireConnection)) { return }
        $mbx = $txtFcMailbox.Text.Trim()
        if (-not $mbx) {
            [System.Windows.Forms.MessageBox]::Show('Enter a mailbox.','Folder Cleanup','OK','Warning')
            return
        }
        $folder = if ($cmbFcFolder.SelectedItem -and $cmbFcFolder.SelectedItem -ne '(All Folders)') { $cmbFcFolder.SelectedItem } else { '' }
        if (-not $folder) {
            [System.Windows.Forms.MessageBox]::Show(
                "Select a specific folder to delete from.`n`nDeleting from (All Folders) is not allowed here.`nUse the Mailbox Search tab for org-wide operations.",
                'Folder Required', 'OK', 'Warning')
            return
        }
        $subj = $txtFcSubject.Text.Trim()
        $frm = $txtFcFrom.Text.Trim()
        $days = [int]$numFcDays.Value
        $sz = if ($cmbFcSize.SelectedItem -and $cmbFcSize.SelectedItem -ne 'Any') { $cmbFcSize.SelectedItem } else { '' }
        $att = $chkFcAttach.Checked
        $cmdText = & $buildFcCommand -Mbx $mbx -Folder $folder -Subj $subj -Frm $frm -Days $days -Sz $sz -Att $att -Action 'Delete'
        $folderInfo = " from folder `"$folder`""
        $result = & $confirmDelete -Message "Delete matching messages from $mbx$folderInfo?`nThis action is permanent!" -CommandText $cmdText
        if ($result -ne 'OK') { return }
        Update-StatusBar "Deleting messages from $mbx..."
        $btnFcDelete.Enabled = $false

        $srv = $txtServer.Text.Trim()
        Start-AsyncJob -Name "Folder Delete $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $FolderPath, $Subject, $From, $OlderThanDays, $SizeRange, $HasAttachment, $Server)
            $p = @{ Mailbox = $Mailbox; Action = 'DeleteContent' }
            if ($FolderPath)    { $p['FolderPath'] = $FolderPath }
            if ($Subject)       { $p['Subject'] = $Subject }
            if ($From)          { $p['From'] = $From }
            if ($OlderThanDays -gt 0) { $p['OlderThanDays'] = $OlderThanDays }
            if ($SizeRange)     { $p['SizeRange'] = $SizeRange }
            if ($HasAttachment) { $p['HasAttachment'] = $true }
            if ($Server)        { $p['Server'] = $Server }
            Invoke-FolderCleanup @p
        } -Parameters @{
            Mailbox = $mbx; FolderPath = $folder; Subject = $subj; From = $frm
            OlderThanDays = $days; SizeRange = $sz; HasAttachment = $att; Server = $srv
        } -OnComplete {
            param($result)
            try {
                $data = @($result)
                $script:LastFolderCleanup = $data
                Set-DGVData -DGV $dgvFolder -Data $data
                $total = ($data | Measure-Object -Property ResultItems -Sum).Sum
                Update-StatusBar "Folder delete complete: $total item(s) deleted"
                try { Write-SearchLog -Action 'FolderDelete' -Scope $mbx -Result "$total items deleted" } catch {}
                try { Write-OperatorLog -Action 'FolderDelete' -Target $mbx -Details "$total items" } catch {}
            } catch { Update-StatusBar "Folder delete error: $_" }
            $btnFcDelete.Enabled = $true
        } -OnError {
            param($err)
            $btnFcDelete.Enabled = $true
            Update-StatusBar "Folder delete error: $err"
            if ($err -match '(?i)denied|impersonat|403|401|unauthorized') {
                [System.Windows.Forms.MessageBox]::Show(
                    "EWS access denied.`n`nFix: Run this in Exchange Management Shell:`nNew-ManagementRoleAssignment -Name `"EWSImpersonation`" -Role `"ApplicationImpersonation`" -User $env:USERNAME",
                    'EWS Impersonation Required', 'OK', 'Warning')
            }
        }
    })

    $btnFcPurge.Add_Click({
        if (-not (& $requireConnection)) { return }
        $mbx = $txtFcMailbox.Text.Trim()
        if (-not $mbx) {
            [System.Windows.Forms.MessageBox]::Show('Enter a mailbox.','Folder Cleanup','OK','Warning')
            return
        }
        $cmdEstimate = "Search-Mailbox -Identity `"$mbx`" -SearchDumpsterOnly -EstimateResultOnly"
        $cmdDelete = "Search-Mailbox -Identity `"$mbx`" -SearchDumpsterOnly -DeleteContent -Force"

        # First ask estimate or delete
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Purge recoverable/deleted items from $mbx?`n`nEstimate first (recommended)?",
            'Dumpster Purge', 'YesNoCancel', 'Question')
        if ($confirm -eq 'Cancel') { return }

        if ($confirm -eq 'No') {
            # Delete — use confirmDelete dialog
            $result = & $confirmDelete -Message "Permanently purge all recoverable items from $mbx?`nThis action is IRREVERSIBLE!" -CommandText $cmdDelete
            if ($result -ne 'OK') { return }
            $action = 'DeleteContent'
        } else {
            # Estimate — use safe mode
            if ((& $showCommandPreview -CommandText $cmdEstimate -Title 'Dumpster Estimate') -eq 'Cancel') { return }
            $action = 'Estimate'
        }

        Update-StatusBar "Purging dumpster for $mbx ($action)..."
        $btnFcPurge.Enabled = $false

        Start-AsyncJob -Name "Dumpster $action $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $Action)
            Invoke-PurgeDeletedItems -Mailbox $Mailbox -Action $Action
        } -Parameters @{ Mailbox = $mbx; Action = $action } -OnComplete {
            param($result)
            try {
                $data = @($result)
                $script:LastFolderCleanup = $data
                Set-DGVData -DGV $dgvFolder -Data $data
                $total = ($data | Measure-Object -Property ResultItems -Sum).Sum
                Update-StatusBar "Dumpster $action`: $total item(s)"
                try { Write-SearchLog -Action "DumpsterPurge-$action" -Scope $mbx -Result "$total items" } catch {}
            } catch { Update-StatusBar "Dumpster error: $_" }
            $btnFcPurge.Enabled = $true
        } -OnError {
            param($err)
            $btnFcPurge.Enabled = $true
            Update-StatusBar "Dumpster error: $err"
        }
    })

    $btnFcFindDupes.Add_Click({
        if (-not (& $requireConnection)) { return }
        $mbx = $txtFcMailbox.Text.Trim()
        if (-not $mbx) {
            [System.Windows.Forms.MessageBox]::Show('Enter a mailbox.','Duplicates','OK','Warning')
            return
        }
        $folder = if ($cmbFcFolder.SelectedItem -and $cmbFcFolder.SelectedItem -ne '(All Folders)') { $cmbFcFolder.SelectedItem } else { '' }
        $cmdText = "# Scan for duplicate messages (by Subject + Sender + Date)`r`nGet-MailboxFolderStatistics -Identity `"$mbx`""
        if ($folder) { $cmdText += " | Where FolderPath -eq `"/$folder`"" }
        $cmdText += "`r`n# Then compare items within each folder for duplicates"
        if ((& $showCommandPreview -CommandText $cmdText -Title 'Find Duplicates') -eq 'Cancel') { return }
        Update-StatusBar "Scanning for duplicates in $mbx..."
        $btnFcFindDupes.Enabled = $false

        Start-AsyncJob -Name "FindDupes $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $FolderPath)
            $p = @{ Mailbox = $Mailbox; DaysBack = 30 }
            if ($FolderPath) { $p['FolderPath'] = $FolderPath }
            Find-MailboxDuplicates @p
        } -Parameters @{ Mailbox = $mbx; FolderPath = $folder } -OnComplete {
            param($result)
            try {
                $data = @($result)
                $script:LastDuplicates = $data
                Set-DGVData -DGV $dgvFolder -Data $data
                $dupeCount = ($data | Where-Object { $_.Status -eq 'PossibleDupes' } | Measure-Object).Count
                Update-StatusBar "Duplicate scan: $($data.Count) folders, $dupeCount with possible duplicates"
            } catch { Update-StatusBar "Duplicate scan error: $_" }
            $btnFcFindDupes.Enabled = $true
        } -OnError {
            param($err)
            $btnFcFindDupes.Enabled = $true
            Update-StatusBar "Duplicate scan error: $err"
        }
    })

    $btnFcBackupDupes.Add_Click({
        if (-not (& $requireConnection)) { return }
        $mbx = $txtFcMailbox.Text.Trim()
        $target = $txtFcTarget.Text.Trim()
        $folder = if ($cmbFcFolder.SelectedItem -and $cmbFcFolder.SelectedItem -ne '(All Folders)') { $cmbFcFolder.SelectedItem } else { '' }
        if (-not $mbx -or -not $folder) {
            [System.Windows.Forms.MessageBox]::Show('Enter mailbox and select a specific folder.','Duplicates','OK','Warning')
            return
        }
        if (-not $target) {
            [System.Windows.Forms.MessageBox]::Show('Enter a target mailbox for backup.','Duplicates','OK','Warning')
            return
        }
        $cmdText = "Search-Mailbox -Identity `"$mbx`" -SearchQuery `"folder:$folder`" -TargetMailbox `"$target`" -TargetFolder `"Backup-$folder`" -LogOnly"
        if ((& $showCommandPreview -CommandText $cmdText -Title 'Backup Folder') -eq 'Cancel') { return }
        Update-StatusBar "Backing up folder $folder from $mbx..."
        $btnFcBackupDupes.Enabled = $false

        Start-AsyncJob -Name "BackupFolder $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $FolderPath, $TargetMailbox)
            Remove-FolderDuplicates -Mailbox $Mailbox -FolderPath $FolderPath -TargetMailbox $TargetMailbox -Action 'BackupOnly'
        } -Parameters @{ Mailbox = $mbx; FolderPath = $folder; TargetMailbox = $target } -OnComplete {
            param($result)
            try {
                $data = @($result)
                $script:LastFolderCleanup = $data
                Set-DGVData -DGV $dgvFolder -Data $data
                Update-StatusBar "Folder backup complete"
                try { Write-OperatorLog -Action 'FolderBackup' -Target $mbx -Details "Folder: $folder -> $target" } catch {}
            } catch { Update-StatusBar "Backup error: $_" }
            $btnFcBackupDupes.Enabled = $true
        } -OnError {
            param($err)
            $btnFcBackupDupes.Enabled = $true
            Update-StatusBar "Backup error: $err"
        }
    })

    $btnFcRemoveDupes.Add_Click({
        if (-not (& $requireConnection)) { return }
        $mbx = $txtFcMailbox.Text.Trim()
        $target = $txtFcTarget.Text.Trim()
        $folder = if ($cmbFcFolder.SelectedItem -and $cmbFcFolder.SelectedItem -ne '(All Folders)') { $cmbFcFolder.SelectedItem } else { '' }
        if (-not $mbx -or -not $folder) {
            [System.Windows.Forms.MessageBox]::Show('Enter mailbox and select a specific folder.','Duplicates','OK','Warning')
            return
        }
        if (-not $target) {
            [System.Windows.Forms.MessageBox]::Show('Enter a target mailbox for backup.','Duplicates','OK','Warning')
            return
        }
        $cmdText = "# Step 1: Backup`r`nSearch-Mailbox -Identity `"$mbx`" -SearchQuery `"folder:$folder`" -TargetMailbox `"$target`" -TargetFolder `"Backup-$folder`"`r`n# Step 2: Delete`r`nSearch-Mailbox -Identity `"$mbx`" -SearchQuery `"folder:$folder`" -DeleteContent -Force"
        $result = & $confirmDelete -Message "BACKUP folder content to $target, then DELETE from $mbx.`n`nFolder: $folder" -CommandText $cmdText
        if ($result -ne 'OK') { return }

        Update-StatusBar "Backup + Delete folder $folder from $mbx..."
        $btnFcRemoveDupes.Enabled = $false

        Start-AsyncJob -Name "RemoveDupes $mbx" -Form $form -ScriptBlock {
            param($Mailbox, $FolderPath, $TargetMailbox)
            Remove-FolderDuplicates -Mailbox $Mailbox -FolderPath $FolderPath -TargetMailbox $TargetMailbox -Action 'BackupAndDelete'
        } -Parameters @{ Mailbox = $mbx; FolderPath = $folder; TargetMailbox = $target } -OnComplete {
            param($result)
            try {
                $data = @($result)
                $script:LastFolderCleanup = $data
                Set-DGVData -DGV $dgvFolder -Data $data
                Update-StatusBar "Backup + Delete complete"
                try { Write-OperatorLog -Action 'FolderBackupDelete' -Target $mbx -Details "Folder: $folder -> $target" } catch {}
            } catch { Update-StatusBar "Backup+Delete error: $_" }
            $btnFcRemoveDupes.Enabled = $true
        } -OnError {
            param($err)
            $btnFcRemoveDupes.Enabled = $true
            Update-StatusBar "Backup+Delete error: $err"
        }
    })

    $btnFcExport.Add_Click({ Show-Export -Data $script:LastFolderCleanup -DefaultName 'folder-cleanup' })

    # ─── Assemble tabs ───────────────────────────────────────────────────────
    $tabs.TabPages.AddRange(@($tabSearch, $tabOrgWide, $tabCompliance, $tabMailboxes, $tabFolder, $tabAudit))
    $form.Controls.Add($tabs)
    $tabs.BringToFront()

    # ─── Connect (async) ─────────────────────────────────────────────────────
    $btnConnect.Add_Click({
        $server = $txtServer.Text.Trim()
        if (-not $server) {
            [System.Windows.Forms.MessageBox]::Show('Enter an Exchange server name.','Connect','OK','Warning')
            return
        }
        $cmdText = "`$session = New-PSSession -ConfigurationName 'Microsoft.Exchange' -ConnectionUri `"http://$server/PowerShell/`" -Authentication Kerberos`r`nImport-PSSession `$session -DisableNameChecking -AllowClobber"
        if ((& $showCommandPreview -CommandText $cmdText -Title 'Connect to Exchange') -eq 'Cancel') { return }
        Update-StatusBar "Connecting to $server..."
        $lblConnStatus.Text = 'Connecting...'
        $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(200,150,0)
        $btnConnect.Enabled = $false

        Start-AsyncJob -Name "Connect $server" -Form $form -ScriptBlock {
            param($Server)
            $session = Connect-ExchangeSearch -Server $Server
            $version = Get-ExchangeServerVersion -Server $Server
            return @{ Session = $session; Version = $version }
        } -Parameters @{ Server = $server } -OnComplete {
            param($result)
            try {
                if ($script:Session) {
                    try { Disconnect-ExchangeSearch -Session $script:Session } catch {}
                }
                $script:Session = $result.Session
                $ver = $result.Version
                $lblConnStatus.Text = "Connected: $server ($($ver.AdminVersion))"
                $lblConnStatus.ForeColor = [System.Drawing.Color]::Green
                $btnConnect.Enabled = $true
                $btnDisconnect.Visible = $true
                try { Update-RecentServers -Server $server } catch {}
                try { Write-OperatorLog -Action 'Connect' -Target $server } catch {}
                Update-StatusBar "Connected to $server — $($ver.Edition) $($ver.AdminVersion)"
            } catch {
                $lblConnStatus.Text = 'Connection setup error'
                $lblConnStatus.ForeColor = [System.Drawing.Color]::Red
                $btnConnect.Enabled = $true
                Update-StatusBar "Connection setup error: $_"
            }
        } -OnError {
            param($err)
            $lblConnStatus.Text = 'Connection failed'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Red
            $btnConnect.Enabled = $true
            Update-StatusBar "Connection failed: $err"
            [System.Windows.Forms.MessageBox]::Show("Failed to connect: $err",'Connection Error','OK','Error')
        }
    })

    $btnDisconnect.Add_Click({
        try {
            if ($script:Session) {
                try { Disconnect-ExchangeSearch -Session $script:Session } catch {}
                $script:Session = $null
            }
            $lblConnStatus.Text = 'Disconnected'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Gray
            $btnDisconnect.Visible = $false
            Update-StatusBar 'Disconnected'
        } catch {}
    })

    # ─── Async Poller ────────────────────────────────────────────────────────
    $asyncPoller = New-AsyncPollerTimer

    # ─── Keyboard shortcuts ──────────────────────────────────────────────────
    $form.Add_KeyDown({
        param($s, $e)
        try {
            if ($e.KeyCode -eq 'F5') {
                $e.Handled = $true
                switch ($tabs.SelectedTab) {
                    $tabCompliance { & $refreshCompliance }
                    $tabMailboxes  { $btnMbxLoad.PerformClick() }
                    $tabAudit      { $btnAuditRefresh.PerformClick() }
                    $tabFolder     { $btnFcLoadFolders.PerformClick() }
                }
            }
            if ($e.Control -and $e.KeyCode -eq 'E') {
                $e.Handled = $true
                switch ($tabs.SelectedTab) {
                    $tabSearch     { Show-Export -Data $script:LastSearchResults -DefaultName 'search' }
                    $tabOrgWide    { Show-Export -Data $script:LastOrgResults -DefaultName 'org-wide' }
                    $tabMailboxes  { Show-Export -Data $script:LastMailboxList -DefaultName 'mailboxes' }
                    $tabFolder     { Show-Export -Data $script:LastFolderCleanup -DefaultName 'folder-cleanup' }
                    $tabAudit      { Show-Export -Data $script:LastAuditLog -DefaultName 'audit' }
                }
            }
        } catch {}
    })

    # ─── Form events ─────────────────────────────────────────────────────────
    $form.Add_FormClosing({
        try {
            $settings = Get-AppSettings
            $settings.WindowWidth = $form.Width
            $settings.WindowHeight = $form.Height
            $settings.LastServer = $txtServer.Text
            Save-AppSettings -Settings $settings
        } catch {}

        try { $asyncPoller.Stop(); $asyncPoller.Dispose() } catch {}

        try {
            foreach ($job in $script:AsyncJobs) {
                if ($job.Status -eq 'Running') {
                    try { $job.PowerShell.Stop() } catch {}
                    try { $job.PowerShell.Dispose() } catch {}
                    try { $job.Runspace.Close(); $job.Runspace.Dispose() } catch {}
                }
            }
        } catch {}

        if ($script:Session) {
            try { Disconnect-ExchangeSearch -Session $script:Session } catch {}
        }
    })

    # ─── Auto-detect EMS and connect ─────────────────────────────────────────
    $form.Add_Shown({
        try {
            if (Test-ExchangeManagementShell) {
                Update-StatusBar 'Exchange Management Shell detected. Discovering servers...'
                $lblConnStatus.Text = 'EMS detected...'
                $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(200,150,0)
                $form.Refresh()

                $servers = @(Find-ExchangeServers)
                if ($servers.Count -gt 0) {
                    # Pick local server if possible, otherwise first one
                    $localName = $env:COMPUTERNAME
                    $local = $servers | Where-Object { $_.Name -eq $localName } | Select-Object -First 1
                    $picked = if ($local) { $local } else { $servers[0] }

                    $txtServer.Text = $picked.FQDN
                    $autoComplete.Clear()
                    foreach ($s in $servers) { [void]$autoComplete.Add($s.FQDN) }

                    # Auto-connect (EMS - no remote session needed)
                    $script:Session = @{ IsEMS = $true; Server = $picked.FQDN }
                    $ver = Get-ExchangeServerVersion -Server $picked.Name
                    $lblConnStatus.Text = "EMS: $($picked.Name) ($($ver.AdminVersion))"
                    $lblConnStatus.ForeColor = [System.Drawing.Color]::Green
                    $btnDisconnect.Visible = $true
                    try { Write-OperatorLog -Action 'AutoConnect-EMS' -Target $picked.Name } catch {}
                    Update-StatusBar "Auto-connected via EMS - $($picked.Name) ($($ver.Edition) $($ver.AdminVersion)) | $($servers.Count) server(s) found"
                } else {
                    Update-StatusBar 'EMS detected but no Mailbox servers found'
                }
            }
        } catch {
            Update-StatusBar "EMS auto-detect: $_"
        }
    })

    [void]$form.ShowDialog()
}

# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if ($MyInvocation.InvocationName -ne '.') {
    Show-EXRESearcherGUI
}
