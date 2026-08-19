# Pingoji - compact network monitor for the Windows system tray
# PowerShell version of Pingoji.vbs

[CmdletBinding()]
param(
    [string]$RemoteHost = '8.8.8.8',
    [int]$IntervalMilliseconds = 1000,
    [int]$TimeoutMilliseconds = 1000,
    [switch]$ShowConsole
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $engine = (Get-Process -Id $PID).Path
    $arguments = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath,
        '-RemoteHost', $RemoteHost, '-IntervalMilliseconds', $IntervalMilliseconds,
        '-TimeoutMilliseconds', $TimeoutMilliseconds)
    if ($ShowConsole) { $arguments += '-ShowConsole' }
    Start-Process -FilePath $engine -ArgumentList $arguments
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not $ShowConsole) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class PingojiWindow {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@
    [void][PingojiWindow]::ShowWindow([PingojiWindow]::GetConsoleWindow(), 0)
}

[Windows.Forms.Application]::EnableVisualStyles()

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class PingojiDrag {
    [DllImport("user32.dll")] public static extern bool ReleaseCapture();
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool DestroyIcon(IntPtr handle);
}
'@

function Get-LatencyColor([double]$Milliseconds) {
    if ($Milliseconds -lt 100) {
        return [Drawing.Color]::FromArgb(0, [Math]::Min(255, [int](100 + 1.55 * $Milliseconds)), 0)
    }
    if ($Milliseconds -lt 300) { return [Drawing.Color]::LimeGreen }
    $ratio = [Math]::Min(1.0, ($Milliseconds - 300) / 300)
    return [Drawing.Color]::FromArgb([int](255 * $ratio), [int](255 - 180 * $ratio), 100)
}

function Get-LocalIpAddress {
    $socket = $null
    try {
        # A UDP connect selects the interface Windows would use for this host;
        # it does not need to send any network traffic.
        $socket = [Net.Sockets.Socket]::new(
            [Net.Sockets.AddressFamily]::InterNetwork,
            [Net.Sockets.SocketType]::Dgram,
            [Net.Sockets.ProtocolType]::Udp
        )
        $socket.Connect($RemoteHost, 53)
        return $socket.LocalEndPoint.Address.ToString()
    } catch {
        $candidate = [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
            Where-Object { $_.OperationalStatus -eq [Net.NetworkInformation.OperationalStatus]::Up } |
            ForEach-Object { $_.GetIPProperties().UnicastAddresses } |
            Where-Object { $_.Address.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork -and -not [Net.IPAddress]::IsLoopback($_.Address) } |
            Select-Object -First 1
        if ($null -ne $candidate) { return $candidate.Address.ToString() }
        return $null
    } finally {
        if ($null -ne $socket) { $socket.Dispose() }
    }
}

function New-TrayIcon([Drawing.Color]$Color = [Drawing.Color]::FromArgb(210, 160, 0)) {
    $bitmap = New-Object Drawing.Bitmap 32, 32
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([Drawing.Color]::Transparent)
    $brush = New-Object Drawing.SolidBrush $Color
    $graphics.FillEllipse($brush, 1, 1, 30, 30)
    $pen = New-Object Drawing.Pen ([Drawing.Color]::White), 3
    $graphics.DrawLine($pen, 7, 23, 7, 18)
    $graphics.DrawLine($pen, 13, 23, 13, 14)
    $graphics.DrawLine($pen, 19, 23, 19, 10)
    $graphics.DrawLine($pen, 25, 23, 25, 6)
    $iconHandle = $bitmap.GetHicon()
    $icon = [Drawing.Icon]::FromHandle($iconHandle).Clone()
    [void][PingojiDrag]::DestroyIcon($iconHandle)
    $pen.Dispose(); $brush.Dispose(); $graphics.Dispose(); $bitmap.Dispose()
    return $icon
}

$form = New-Object Windows.Forms.Form
$form.Text = 'Pingoji'
$form.ClientSize = New-Object Drawing.Size 280, 113
$form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::None
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.ShowInTaskbar = $false
$form.TopMost = $true
$form.StartPosition = [Windows.Forms.FormStartPosition]::Manual
$form.BackColor = [Drawing.Color]::WhiteSmoke
$form.Font = New-Object Drawing.Font 'Segoe UI', 9
$form.Padding = New-Object Windows.Forms.Padding 1

$header = New-Object Windows.Forms.Panel
$header.Location = New-Object Drawing.Point 1, 1
$header.Size = New-Object Drawing.Size 278, 18
$header.BackColor = [Drawing.Color]::FromArgb(55, 55, 58)
$form.Controls.Add($header)

$titleLabel = New-Object Windows.Forms.Label
$titleLabel.Location = New-Object Drawing.Point 6, 0
$titleLabel.Size = New-Object Drawing.Size 245, 18
$titleLabel.ForeColor = [Drawing.Color]::WhiteSmoke
$titleLabel.Font = New-Object Drawing.Font 'Segoe UI', 7.5
$titleLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
$titleLabel.Text = 'Pingoji'
$header.Controls.Add($titleLabel)

$closeLabel = New-Object Windows.Forms.Label
$closeLabel.Location = New-Object Drawing.Point 257, 0
$closeLabel.Size = New-Object Drawing.Size 21, 18
$closeLabel.ForeColor = [Drawing.Color]::WhiteSmoke
$closeLabel.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$closeLabel.Text = [char]0x00D7
$header.Controls.Add($closeLabel)
$closeLabel.Add_MouseEnter({ $closeLabel.BackColor = [Drawing.Color]::Firebrick })
$closeLabel.Add_MouseLeave({ $closeLabel.BackColor = [Drawing.Color]::Transparent })

$dragWindow = {
    [void][PingojiDrag]::ReleaseCapture()
    [void][PingojiDrag]::SendMessage($form.Handle, 0xA1, [IntPtr]2, [IntPtr]0)
}
$header.Add_MouseDown($dragWindow)
$titleLabel.Add_MouseDown($dragWindow)

$statusLabel = New-Object Windows.Forms.Label
$statusLabel.Location = New-Object Drawing.Point 7, 20
$statusLabel.Size = New-Object Drawing.Size 266, 19
$statusLabel.Font = New-Object Drawing.Font 'Segoe UI', 8
$statusLabel.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$statusLabel.Text = "Starting monitor for $RemoteHost..."
$form.Controls.Add($statusLabel)

$recentPanel = New-Object Windows.Forms.FlowLayoutPanel
$recentPanel.Location = New-Object Drawing.Point 8, 39
$recentPanel.Size = New-Object Drawing.Size 264, 32
$recentPanel.WrapContents = $false
$recentPanel.Margin = 0
$form.Controls.Add($recentPanel)

$recentBlocks = @()
for ($i = 0; $i -lt 5; $i++) {
    $block = New-Object Windows.Forms.Panel
    $block.Size = New-Object Drawing.Size 48, 28
    $block.Margin = New-Object Windows.Forms.Padding 2
    $block.BackColor = [Drawing.Color]::DarkGray
    $valueLabel = New-Object Windows.Forms.Label
    $valueLabel.Location = New-Object Drawing.Point 1, 1
    $valueLabel.Size = New-Object Drawing.Size 33, 26
    $valueLabel.TextAlign = [Drawing.ContentAlignment]::MiddleRight
    $valueLabel.ForeColor = [Drawing.Color]::White
    $valueLabel.Font = New-Object Drawing.Font 'Segoe UI', 8
    $valueLabel.Text = '-'
    $unitLabel = New-Object Windows.Forms.Label
    $unitLabel.Location = New-Object Drawing.Point 34, 5
    $unitLabel.Size = New-Object Drawing.Size 13, 20
    $unitLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $unitLabel.ForeColor = [Drawing.Color]::WhiteSmoke
    $unitLabel.Font = New-Object Drawing.Font 'Segoe UI', 5.5
    $unitLabel.Text = ''
    $block.Controls.Add($valueLabel); $block.Controls.Add($unitLabel)
    $recentPanel.Controls.Add($block)
    $recentBlocks += [pscustomobject]@{ Panel = $block; Value = $valueLabel; Unit = $unitLabel }
}

$historyPanel = New-Object Windows.Forms.Panel
$historyPanel.Location = New-Object Drawing.Point 8, 73
$historyPanel.Size = New-Object Drawing.Size 264, 10
$historyPanel.BackColor = [Drawing.Color]::Gainsboro
$form.Controls.Add($historyPanel)

$minutePanel = New-Object Windows.Forms.Panel
$minutePanel.Location = New-Object Drawing.Point 8, 87
$minutePanel.Size = New-Object Drawing.Size 264, 7
$minutePanel.BackColor = [Drawing.Color]::Gainsboro
$form.Controls.Add($minutePanel)

$hintLabel = New-Object Windows.Forms.Label
$hintLabel.Location = New-Object Drawing.Point 8, 96
$hintLabel.Size = New-Object Drawing.Size 264, 13
$hintLabel.Font = New-Object Drawing.Font 'Segoe UI', 6.5
$hintLabel.ForeColor = [Drawing.Color]::DimGray
$hintLabel.Text = 'Current samples / one-minute averages'
$form.Controls.Add($hintLabel)

$state = [hashtable]::Synchronized(@{
    Samples = New-Object System.Collections.ArrayList
    Minutes = New-Object System.Collections.ArrayList
    Records = New-Object System.Collections.ArrayList
    Consecutive = 0
    StableSince = [DateTime]::Now
    MinuteSum = 0.0
    MinuteCount = 0
    RecentLatency = New-Object System.Collections.ArrayList
    Paused = $false
    Exiting = $false
})

$historyPanel.Add_Paint({
    param($sender, $e)
    $items = @($state.Samples)
    $width = 4
    $max = [Math]::Floor($sender.ClientSize.Width / $width)
    $start = [Math]::Max(0, $items.Count - $max)
    for ($i = $start; $i -lt $items.Count; $i++) {
        $brush = New-Object Drawing.SolidBrush $items[$i]
        $x = ($i - $start) * $width
        $e.Graphics.FillRectangle($brush, $x, 0, $width - 1, $sender.ClientSize.Height)
        $brush.Dispose()
    }
})

$minutePanel.Add_Paint({
    param($sender, $e)
    $items = @($state.Minutes)
    $width = 5
    $max = [Math]::Floor($sender.ClientSize.Width / $width)
    $start = [Math]::Max(0, $items.Count - $max)
    for ($i = $start; $i -lt $items.Count; $i++) {
        $brush = New-Object Drawing.SolidBrush $items[$i]
        $x = ($i - $start) * $width
        $e.Graphics.FillRectangle($brush, $x, 0, $width - 1, $sender.ClientSize.Height)
        $brush.Dispose()
    }
})

$tray = New-Object Windows.Forms.NotifyIcon
$tray.Icon = New-TrayIcon
$tray.Text = 'Pingoji - starting'
$tray.Visible = $true

$menu = New-Object Windows.Forms.ContextMenuStrip
$showItem = $menu.Items.Add('Show Pingoji')
$pauseItem = $menu.Items.Add('Pause monitoring')
$resetItem = $menu.Items.Add('Reset session')
$exportItem = $menu.Items.Add('Export')
$exportSamplesItem = $exportItem.DropDownItems.Add('Detailed samples (CSV)...')
$exportRangesItem = $exportItem.DropDownItems.Add('Availability ranges (CSV)...')
$startupItem = $menu.Items.Add('Start with Windows')
$startupItem.CheckOnClick = $false
[void]$menu.Items.Add('-')
$exitItem = $menu.Items.Add('Exit')
$tray.ContextMenuStrip = $menu

function Update-TrayIcon {
    if ($state.RecentLatency.Count -lt 5) {
        $iconColor = [Drawing.Color]::FromArgb(210, 160, 0)
    } else {
        $average = ($state.RecentLatency | Measure-Object -Average).Average
        if ($average -lt 300) {
            $iconColor = [Drawing.Color]::FromArgb(35, 160, 70)
        } elseif ($average -lt 600) {
            $iconColor = [Drawing.Color]::FromArgb(230, 165, 0)
        } else {
            $iconColor = [Drawing.Color]::FromArgb(205, 45, 45)
        }
    }
    $oldIcon = $tray.Icon
    $tray.Icon = New-TrayIcon $iconColor
    if ($null -ne $oldIcon) { $oldIcon.Dispose() }
}

function Save-CsvFile([object[]]$Data, [string]$SuggestedName) {
    if ($Data.Count -eq 0) {
        [Windows.Forms.MessageBox]::Show('There is no recorded information to export.', 'Pingoji', 'OK', 'Information') | Out-Null
        return
    }
    $dialog = New-Object Windows.Forms.SaveFileDialog
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dialog.DefaultExt = 'csv'
    $dialog.AddExtension = $true
    $dialog.FileName = $SuggestedName
    try {
        if ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK) {
            $Data | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
            $tray.ShowBalloonTip(1500, 'Pingoji export complete', $dialog.FileName, [Windows.Forms.ToolTipIcon]::Info)
        }
    } catch {
        [Windows.Forms.MessageBox]::Show("Could not export the CSV file:`n$($_.Exception.Message)", 'Pingoji', 'OK', 'Error') | Out-Null
    } finally {
        $dialog.Dispose()
    }
}

function Get-AvailabilityRanges {
    $records = @($state.Records)
    if ($records.Count -eq 0) { return @() }
    $ranges = New-Object System.Collections.ArrayList
    $rangeStart = $records[0].Timestamp
    $rangeAvailable = $records[0].Available
    $rangeRecords = New-Object System.Collections.ArrayList

    foreach ($record in $records) {
        if ($record.Available -ne $rangeAvailable) {
            $end = $rangeRecords[$rangeRecords.Count - 1].Timestamp
            $latencies = @($rangeRecords | Where-Object Available | Select-Object -ExpandProperty LatencyMs)
            [void]$ranges.Add([pscustomobject]@{
                RemoteHost = $RemoteHost; LocalIpAddresses = (($rangeRecords.LocalIpAddress | Where-Object { $_ } | Sort-Object -Unique) -join ';')
                Availability = $(if ($rangeAvailable) { 'Available' } else { 'Unavailable' })
                StartTime = $rangeStart.ToString('yyyy-MM-dd HH:mm:ss.fff'); EndTime = $end.ToString('yyyy-MM-dd HH:mm:ss.fff')
                DurationSeconds = [Math]::Round(($end - $rangeStart).TotalSeconds, 3); SampleCount = $rangeRecords.Count
                AverageLatencyMs = $(if ($latencies.Count) { [Math]::Round(($latencies | Measure-Object -Average).Average, 2) } else { $null })
                MinimumLatencyMs = $(if ($latencies.Count) { ($latencies | Measure-Object -Minimum).Minimum } else { $null })
                MaximumLatencyMs = $(if ($latencies.Count) { ($latencies | Measure-Object -Maximum).Maximum } else { $null })
            })
            $rangeStart = $record.Timestamp
            $rangeAvailable = $record.Available
            $rangeRecords = New-Object System.Collections.ArrayList
        }
        [void]$rangeRecords.Add($record)
    }

    $end = $rangeRecords[$rangeRecords.Count - 1].Timestamp
    $latencies = @($rangeRecords | Where-Object Available | Select-Object -ExpandProperty LatencyMs)
    [void]$ranges.Add([pscustomobject]@{
        RemoteHost = $RemoteHost; LocalIpAddresses = (($rangeRecords.LocalIpAddress | Where-Object { $_ } | Sort-Object -Unique) -join ';')
        Availability = $(if ($rangeAvailable) { 'Available' } else { 'Unavailable' })
        StartTime = $rangeStart.ToString('yyyy-MM-dd HH:mm:ss.fff'); EndTime = $end.ToString('yyyy-MM-dd HH:mm:ss.fff')
        DurationSeconds = [Math]::Round(($end - $rangeStart).TotalSeconds, 3); SampleCount = $rangeRecords.Count
        AverageLatencyMs = $(if ($latencies.Count) { [Math]::Round(($latencies | Measure-Object -Average).Average, 2) } else { $null })
        MinimumLatencyMs = $(if ($latencies.Count) { ($latencies | Measure-Object -Minimum).Minimum } else { $null })
        MaximumLatencyMs = $(if ($latencies.Count) { ($latencies | Measure-Object -Maximum).Maximum } else { $null })
    })
    return @($ranges)
}

function Reset-PingojiSession {
    $state.Samples.Clear(); $state.Minutes.Clear(); $state.Records.Clear(); $state.RecentLatency.Clear()
    $state.Consecutive = 0; $state.StableSince = [DateTime]::Now
    $state.MinuteSum = 0.0; $state.MinuteCount = 0
    foreach ($block in $recentBlocks) {
        $block.Value.Text = '-'; $block.Unit.Text = ''; $block.Panel.BackColor = [Drawing.Color]::DarkGray
    }
    $statusLabel.Text = "Starting monitor for $RemoteHost..."
    $tray.Text = 'Pingoji - starting'
    Update-TrayIcon
    $historyPanel.Invalidate(); $minutePanel.Invalidate()
}

$startupRegistryPath = 'Software\Microsoft\Windows\CurrentVersion\Run'
$startupValueName = 'Pingoji'
$startupKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($startupRegistryPath, $false)
try {
    $startupItem.Checked = $null -ne $startupKey -and $null -ne $startupKey.GetValue($startupValueName, $null)
} finally {
    if ($null -ne $startupKey) { $startupKey.Dispose() }
}

$showWindow = {
    # WorkingArea excludes the taskbar, placing the popup directly above its
    # notification corner without covering system-tray icons.
    $screen = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $form.Location = New-Object Drawing.Point ($screen.Right - $form.Width - 8), ($screen.Bottom - $form.Height - 8)
    $form.Show(); $form.WindowState = 'Normal'; $form.Activate()
    $showItem.Text = 'Hide Pingoji'
}
$hideWindow = { $form.Hide(); $showItem.Text = 'Show Pingoji' }

$closeLabel.Add_Click({ $form.Close() })
$showItem.Add_Click({ if ($form.Visible) { & $hideWindow } else { & $showWindow } })
$tray.Add_DoubleClick({ if ($form.Visible) { & $hideWindow } else { & $showWindow } })
$pauseItem.Add_Click({
    $state.Paused = -not $state.Paused
    $pauseItem.Text = if ($state.Paused) { 'Resume monitoring' } else { 'Pause monitoring' }
    if ($state.Paused) { $statusLabel.Text = 'Monitoring paused'; $tray.Text = 'Pingoji - paused' }
})
$resetItem.Add_Click({
    $answer = [Windows.Forms.MessageBox]::Show(
        'Reset all measurements recorded in this session?',
        'Pingoji - reset session',
        [Windows.Forms.MessageBoxButtons]::YesNo,
        [Windows.Forms.MessageBoxIcon]::Question
    )
    if ($answer -eq [Windows.Forms.DialogResult]::Yes) { Reset-PingojiSession }
})
$exportSamplesItem.Add_Click({
    $rows = @($state.Records | ForEach-Object {
        [pscustomobject]@{
            RemoteHost = $RemoteHost
            LocalIpAddress = $_.LocalIpAddress
            Timestamp = $_.Timestamp.ToString('yyyy-MM-dd HH:mm:ss.fff')
            Available = $_.Available
            LatencyMs = $_.LatencyMs
            Result = $_.Result
        }
    })
    Save-CsvFile $rows ("Pingoji-{0}-samples-{1}.csv" -f ($RemoteHost -replace '[^a-zA-Z0-9.-]', '_'), [DateTime]::Now.ToString('yyyyMMdd-HHmmss'))
})
$exportRangesItem.Add_Click({
    $rows = @(Get-AvailabilityRanges)
    Save-CsvFile $rows ("Pingoji-{0}-availability-{1}.csv" -f ($RemoteHost -replace '[^a-zA-Z0-9.-]', '_'), [DateTime]::Now.ToString('yyyyMMdd-HHmmss'))
})
$startupItem.Add_Click({
    $key = $null
    try {
        $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($startupRegistryPath)
        if ($startupItem.Checked) {
            $key.DeleteValue($startupValueName, $false)
            $startupItem.Checked = $false
            $tray.ShowBalloonTip(1200, 'Pingoji', 'Windows startup disabled.', [Windows.Forms.ToolTipIcon]::Info)
        } else {
            $engine = (Get-Process -Id $PID).Path
            $command = '"{0}" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "{1}" -RemoteHost "{2}" -IntervalMilliseconds {3} -TimeoutMilliseconds {4}' -f $engine, $PSCommandPath, $RemoteHost.Replace('"', ''), $IntervalMilliseconds, $TimeoutMilliseconds
            $key.SetValue($startupValueName, $command, [Microsoft.Win32.RegistryValueKind]::String)
            $startupItem.Checked = $true
            $tray.ShowBalloonTip(1200, 'Pingoji', 'Pingoji will start when you sign in to Windows.', [Windows.Forms.ToolTipIcon]::Info)
        }
    } catch {
        [Windows.Forms.MessageBox]::Show("Could not update Windows startup:`n$($_.Exception.Message)", 'Pingoji', 'OK', 'Error') | Out-Null
    } finally {
        if ($null -ne $key) { $key.Dispose() }
    }
})
$exitItem.Add_Click({ $state.Exiting = $true; $form.Close() })
$form.Add_FormClosing({
    param($sender, $e)
    if (-not $state.Exiting) { $e.Cancel = $true; & $hideWindow }
})

function Update-Display([bool]$Success, [double]$Latency, [string]$Message) {
    $color = if ($Success) { Get-LatencyColor $Latency } else { [Drawing.Color]::Red }
    [void]$state.Records.Add([pscustomobject]@{
        Timestamp = [DateTime]::Now
        LocalIpAddress = Get-LocalIpAddress
        Available = $Success
        LatencyMs = $(if ($Success) { [Math]::Round($Latency, 2) } else { $null })
        Result = $Message
    })
    for ($i = 4; $i -gt 0; $i--) {
        $recentBlocks[$i].Value.Text = $recentBlocks[$i - 1].Value.Text
        $recentBlocks[$i].Unit.Text = $recentBlocks[$i - 1].Unit.Text
        $recentBlocks[$i].Panel.BackColor = $recentBlocks[$i - 1].Panel.BackColor
    }
    $recentBlocks[0].Value.Text = if ($Success) { '{0:0}' -f $Latency } else { 'x' }
    $recentBlocks[0].Unit.Text = if ($Success) { 'ms' } else { '' }
    $recentBlocks[0].Panel.BackColor = if ($Success -and $state.Consecutive -lt 3) { [Drawing.Color]::FromArgb(255,191,0) } else { $color }

    [void]$state.RecentLatency.Insert(0, $(if ($Success) { $Latency } else { 1000.0 }))
    if ($state.RecentLatency.Count -gt 5) { $state.RecentLatency.RemoveAt(5) }
    Update-TrayIcon

    [void]$state.Samples.Add($recentBlocks[0].Panel.BackColor)
    if ($state.Samples.Count -gt 1000) { $state.Samples.RemoveAt(0) }
    $state.MinuteSum += $(if ($Success) { $Latency } else { 1000 })
    $state.MinuteCount++
    if ($state.MinuteCount -ge 60) {
        [void]$state.Minutes.Add((Get-LatencyColor ($state.MinuteSum / $state.MinuteCount)))
        if ($state.Minutes.Count -gt 500) { $state.Minutes.RemoveAt(0) }
        $state.MinuteSum = 0.0; $state.MinuteCount = 0
    }
    $historyPanel.Invalidate(); $minutePanel.Invalidate()

    if ($Success -and $state.Consecutive -ge 3) {
        $seconds = [int]([DateTime]::Now - $state.StableSince).TotalSeconds
        $statusLabel.Text = "Connection to $RemoteHost is stable ($seconds s)."
        $tray.Text = "Pingoji - $RemoteHost`: $([int]$Latency) ms"
    } elseif ($Success) {
        $statusLabel.Text = "Connection to $RemoteHost is unstable."
        $tray.Text = "Pingoji - stabilizing ($([int]$Latency) ms)"
    } else {
        $statusLabel.Text = "Connection to $RemoteHost failed: $Message"
        $tray.Text = "Pingoji - $RemoteHost unavailable"
    }
}

$timer = New-Object Windows.Forms.Timer
$timer.Interval = [Math]::Max(250, $IntervalMilliseconds)
$timer.Add_Tick({
    if ($state.Paused) { return }
    $ping = New-Object Net.NetworkInformation.Ping
    try {
        $reply = $ping.Send($RemoteHost, $TimeoutMilliseconds)
        $success = $reply.Status -eq [Net.NetworkInformation.IPStatus]::Success
        if ($success) { $state.Consecutive++ } else { $state.Consecutive = 0; $state.StableSince = [DateTime]::Now }
        Update-Display $success $reply.RoundtripTime ([string]$reply.Status)
    } catch {
        $state.Consecutive = 0; $state.StableSince = [DateTime]::Now
        Update-Display $false 0 $_.Exception.Message
    } finally {
        $ping.Dispose()
    }
})

$form.Add_Shown({
    $tray.ShowBalloonTip(1500, 'Pingoji is running', "Monitoring $RemoteHost. Close the window to keep it in the tray.", [Windows.Forms.ToolTipIcon]::Info)
})
$form.Add_FormClosed({
    $timer.Stop(); $tray.Visible = $false; $tray.Dispose()
})

& $showWindow
$timer.Start()
[Windows.Forms.Application]::Run($form)
