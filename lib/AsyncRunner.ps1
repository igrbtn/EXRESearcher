<#
.SYNOPSIS
    Async execution framework for EXRESearcher WinForms GUI.
    Uses PowerShell runspaces with form timer polling for thread-safe UI updates.
#>

# ─── Job Tracker ─────────────────────────────────────────────────────────────

$script:AsyncJobs = [System.Collections.Generic.List[hashtable]]::new()
$script:AsyncJobId = 0

function Start-AsyncJob {
    <#
    .SYNOPSIS
        Run a scriptblock asynchronously in a runspace.
        Returns job hashtable with Id, Name, Status, StartTime.
        The polling timer (Start-AsyncPoller) handles completion callbacks.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [hashtable]$Parameters = @{},
        [scriptblock]$OnComplete,
        [scriptblock]$OnError,
        [System.Windows.Forms.Form]$Form
    )

    $script:AsyncJobId++
    $jobId = $script:AsyncJobId

    # Create runspace with imported modules
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    # Import functions from lib files into runspace
    $scriptRoot = $PSScriptRoot
    $initScript = @"
Set-Location '$($PWD.Path)'
. '$scriptRoot/Core.ps1'
. '$scriptRoot/Settings.ps1'
"@

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace

    # First run init script to load functions
    [void]$ps.AddScript($initScript)
    try { [void]$ps.Invoke() } catch {}
    $ps.Commands.Clear()
    $ps.Streams.Error.Clear()

    # Now add the actual work
    [void]$ps.AddScript($ScriptBlock)
    foreach ($key in $Parameters.Keys) {
        [void]$ps.AddParameter($key, $Parameters[$key])
    }

    $asyncResult = $ps.BeginInvoke()

    $job = @{
        Id          = $jobId
        Name        = $Name
        Status      = 'Running'
        StartTime   = Get-Date
        EndTime     = $null
        Duration    = $null
        PowerShell  = $ps
        AsyncResult = $asyncResult
        Runspace    = $runspace
        OnComplete  = $OnComplete
        OnError     = $OnError
        Form        = $Form
        Result      = $null
        Error       = $null
    }

    $script:AsyncJobs.Add($job)

    # Log to job console
    if ($script:JobConsole) {
        $ts = (Get-Date).ToString('HH:mm:ss')
        $script:JobConsole.Invoke([Action]{
            $script:JobConsole.AppendText("[$ts] START  #$jobId $Name`r`n")
            $script:JobConsole.ScrollToCaret()
        })
    }

    return $job
}

function Update-AsyncJobs {
    <#
    .SYNOPSIS
        Called by form timer. Checks all running jobs, invokes callbacks on completion.
        Must be called from UI thread.
    #>
    [CmdletBinding()]
    param()

    for ($i = $script:AsyncJobs.Count - 1; $i -ge 0; $i--) {
        $job = $script:AsyncJobs[$i]
        if ($job.Status -ne 'Running') { continue }

        if ($job.AsyncResult.IsCompleted) {
            try {
                $result = $job.PowerShell.EndInvoke($job.AsyncResult)
                $errors = $job.PowerShell.Streams.Error

                $job.EndTime = Get-Date
                $job.Duration = ($job.EndTime - $job.StartTime)

                if ($errors.Count -gt 0) {
                    $job.Status = 'Failed'
                    $job.Error = ($errors | ForEach-Object { $_.ToString() }) -join '; '

                    # Log
                    $ts = (Get-Date).ToString('HH:mm:ss')
                    $dur = [math]::Round($job.Duration.TotalSeconds, 1)
                    if ($script:JobConsole) {
                        $script:JobConsole.AppendText("[$ts] FAIL   #$($job.Id) $($job.Name) (${dur}s) $($job.Error)`r`n")
                        $script:JobConsole.ScrollToCaret()
                    }

                    if ($job.OnError) {
                        try { & $job.OnError $job.Error } catch {}
                    }
                } else {
                    $job.Status = 'Completed'
                    $job.Result = $result

                    # Log
                    $ts = (Get-Date).ToString('HH:mm:ss')
                    $dur = [math]::Round($job.Duration.TotalSeconds, 1)
                    if ($script:JobConsole) {
                        $script:JobConsole.AppendText("[$ts] DONE   #$($job.Id) $($job.Name) (${dur}s)`r`n")
                        $script:JobConsole.ScrollToCaret()
                    }

                    if ($job.OnComplete) {
                        try { & $job.OnComplete $result } catch {}
                    }
                }
            }
            catch {
                $job.Status = 'Failed'
                $job.Error = $_.ToString()
                $job.EndTime = Get-Date
                $job.Duration = ($job.EndTime - $job.StartTime)

                $ts = (Get-Date).ToString('HH:mm:ss')
                if ($script:JobConsole) {
                    $script:JobConsole.AppendText("[$ts] ERROR  #$($job.Id) $($job.Name): $($job.Error)`r`n")
                    $script:JobConsole.ScrollToCaret()
                }

                if ($job.OnError) {
                    try { & $job.OnError $job.Error } catch {}
                }
            }
            finally {
                # Cleanup runspace
                try {
                    $job.PowerShell.Dispose()
                    $job.Runspace.Close()
                    $job.Runspace.Dispose()
                } catch {}
            }

            # Update progress bar
            Update-AsyncProgress
        }
    }
}

function Update-AsyncProgress {
    <#
    .SYNOPSIS
        Update the progress bar based on running jobs count.
    #>
    $running = ($script:AsyncJobs | Where-Object { $_.Status -eq 'Running' }).Count
    if ($script:ProgressBar) {
        if ($running -gt 0) {
            $script:ProgressBar.Style = 'Marquee'
            $script:ProgressBar.MarqueeAnimationSpeed = 30
            $script:ProgressBar.Visible = $true
        } else {
            $script:ProgressBar.Style = 'Continuous'
            $script:ProgressBar.Value = 0
            $script:ProgressBar.Visible = $false
        }
    }
    if ($script:JobCountLabel) {
        if ($running -gt 0) {
            $script:JobCountLabel.Text = "Running: $running job(s)"
            $script:JobCountLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        } else {
            $script:JobCountLabel.Text = ''
        }
    }
}

function Get-RunningJobCount {
    return ($script:AsyncJobs | Where-Object { $_.Status -eq 'Running' }).Count
}

function Clear-CompletedJobs {
    <#
    .SYNOPSIS
        Remove completed/failed jobs from the tracker (keep last 50).
    #>
    $completed = $script:AsyncJobs | Where-Object { $_.Status -ne 'Running' }
    if ($completed.Count -gt 50) {
        $toRemove = $completed | Select-Object -First ($completed.Count - 50)
        foreach ($j in $toRemove) { $script:AsyncJobs.Remove($j) }
    }
}

# ─── GUI Components for Job Console ──────────────────────────────────────────

function New-JobConsolePanel {
    <#
    .SYNOPSIS
        Creates the bottom panel with progress bar, job count, and collapsible job console.
        Returns the panel. Sets script-scope variables for components.
    #>
    [CmdletBinding()]
    param([int]$Height = 120)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = $Height

    # Top bar: progress + job count + toggle
    $topBar = New-Object System.Windows.Forms.Panel
    $topBar.Dock = 'Top'
    $topBar.Height = 28
    $topBar.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 240)

    $script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $script:ProgressBar.Location = New-Object System.Drawing.Point(5, 4)
    $script:ProgressBar.Size = New-Object System.Drawing.Size(200, 20)
    $script:ProgressBar.Style = 'Continuous'
    $script:ProgressBar.Visible = $false

    $script:JobCountLabel = New-Object System.Windows.Forms.Label
    $script:JobCountLabel.Location = New-Object System.Drawing.Point(215, 6)
    $script:JobCountLabel.AutoSize = $true
    $script:JobCountLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $script:JobCountLabel.Text = ''

    $btnToggle = New-Object System.Windows.Forms.Button
    $btnToggle.Text = 'Jobs'
    $btnToggle.FlatStyle = 'Flat'
    $btnToggle.Size = New-Object System.Drawing.Size(60, 22)
    $btnToggle.Location = New-Object System.Drawing.Point(500, 3)
    $btnToggle.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $btnToggle.Anchor = 'Top, Right'

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = 'Clear'
    $btnClear.FlatStyle = 'Flat'
    $btnClear.Size = New-Object System.Drawing.Size(50, 22)
    $btnClear.Location = New-Object System.Drawing.Point(565, 3)
    $btnClear.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $btnClear.Anchor = 'Top, Right'

    $topBar.Controls.AddRange(@($script:ProgressBar, $script:JobCountLabel, $btnToggle, $btnClear))

    # Job console textbox
    $script:JobConsole = New-Object System.Windows.Forms.TextBox
    $script:JobConsole.Dock = 'Fill'
    $script:JobConsole.Multiline = $true
    $script:JobConsole.ScrollBars = 'Vertical'
    $script:JobConsole.ReadOnly = $true
    $script:JobConsole.WordWrap = $false
    $script:JobConsole.Font = New-Object System.Drawing.Font('Consolas', 8.5)
    $script:JobConsole.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 30)
    $script:JobConsole.ForeColor = [System.Drawing.Color]::FromArgb(180, 200, 180)

    $panel.Controls.Add($script:JobConsole)
    $panel.Controls.Add($topBar)

    # Toggle collapse
    $script:JobConsoleExpanded = $true
    $btnToggle.Add_Click({
        if ($script:JobConsoleExpanded) {
            $panel.Height = 28
            $script:JobConsoleExpanded = $false
            $btnToggle.Text = 'Jobs +'
        } else {
            $panel.Height = $Height
            $script:JobConsoleExpanded = $true
            $btnToggle.Text = 'Jobs'
        }
    })

    $btnClear.Add_Click({
        $script:JobConsole.Text = ''
        Clear-CompletedJobs
    })

    return $panel
}

# ─── Async Poller Timer ──────────────────────────────────────────────────────

function New-AsyncPollerTimer {
    <#
    .SYNOPSIS
        Creates a WinForms timer that polls async jobs every 200ms.
    #>
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 200
    $timer.Add_Tick({ Update-AsyncJobs })
    $timer.Start()
    return $timer
}
