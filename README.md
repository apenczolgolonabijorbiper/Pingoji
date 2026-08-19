# &#8584; Pingoji — network availability monitor

Pingoji is a lightweight Windows network monitor. It continuously pings a remote host, displays recent latency and connection history in a compact always-on-top popup, and keeps running in the system tray.

The current version is implemented in PowerShell and does not require Internet Explorer, AutoHotkey, installation, or administrator privileges.

## Features

- Runs quietly in the Windows system tray with no console window by default.
- Opens a compact, draggable popup immediately above the taskbar notification area.
- Shows the five most recent ping results in color-coded blocks.
- Displays short-term samples and one-minute average history strips.
- Treats the connection as stable after three consecutive successful replies.
- Uses a dynamic tray icon based on the average of the latest five measurements:
  - **Green:** below 300 ms
  - **Yellow:** 300–599 ms, or fewer than five measurements collected
  - **Red:** 600 ms or higher
- Counts a failed ping as 1000 ms when calculating the tray-icon average.
- Can start automatically when the current user signs in to Windows.
- Records measurements in memory for the duration of the session.
- Exports detailed measurements or aggregated availability ranges to CSV.
- Records the local outbound IPv4 address, including changes caused by switching Wi-Fi, Ethernet, or VPN interfaces.

## Requirements

- Windows
- Windows PowerShell 5.1 or a compatible PowerShell installation

No additional applications or modules are needed.

## Running Pingoji

Double-click `Pingoji.cmd`, or start the script from PowerShell:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Pingoji.ps1
```

Pingoji monitors `8.8.8.8` once per second by default. Command-line parameters can change the target and timing:

```powershell
.\Pingoji.ps1 -RemoteHost 1.1.1.1 -IntervalMilliseconds 1000 -TimeoutMilliseconds 1000
```

Available parameters:

| Parameter | Default | Description |
| --- | ---: | --- |
| `-RemoteHost` | `8.8.8.8` | IP address or hostname to monitor |
| `-IntervalMilliseconds` | `1000` | Delay between measurements; minimum 250 ms |
| `-TimeoutMilliseconds` | `1000` | Maximum time to wait for a ping reply |
| `-ShowConsole` | off | Keep the PowerShell console visible for diagnostics |

## Understanding the popup

The five blocks contain the latest latency values. A failed measurement is shown as `x`.

- Darker/faster green shades represent good response times.
- Amber indicates that the connection has not yet produced enough consecutive successful replies to be considered stable.
- Red indicates a failed ping.

The first thin strip records individual measurements. The second records one-minute averages. The popup stays above other windows and can be dragged using its compact title bar.

Clicking the popup's **×** button hides it without stopping monitoring. Double-click the tray icon or select **Show Pingoji** to display it again.

## System-tray menu

Right-click the Pingoji icon to access:

- **Show/Hide Pingoji** — toggles the popup.
- **Pause/Resume monitoring** — temporarily stops or resumes measurements.
- **Reset session** — after confirmation, clears all recorded data, five result blocks, history strips, counters, and tray status.
- **Export** — exports recorded information in CSV format.
- **Start with Windows** — enables or disables automatic startup for the current Windows user. No administrator rights are required.
- **About Pingoji** — shows the build date, author, and a clickable link to the project on GitHub.
- **Exit** — stops monitoring and closes Pingoji.

Closing the popup does not exit the application; use **Exit** from the tray menu.

## CSV exports

Measurements remain in memory until the session is reset or Pingoji exits.

### Detailed samples

**Export → Detailed samples (CSV)...** writes one row for every ping with:

- `RemoteHost`
- `LocalIpAddress`
- `Timestamp`
- `Available`
- `LatencyMs`
- `Result`

`LocalIpAddress` is the local IPv4 address of the interface Windows selected for the monitored host. It is not the public internet-facing address.

### Availability ranges

**Export → Availability ranges (CSV)...** combines consecutive measurements with the same availability state. Each row contains:

- remote host and local IP address or addresses observed during the range,
- `Available` or `Unavailable` state,
- start and end timestamps,
- observed duration and sample count,
- average, minimum, and maximum latency where available.

## Legacy version

`Pingoji.vbs` and `SetAlwaysOnTop.ahk` are retained as the original legacy implementation. That version uses Internet Explorer, an external `ping.exe` process, and AutoHotkey for always-on-top behavior. New users should run `Pingoji.ps1` or `Pingoji.cmd`.

## License

See [LICENSE](LICENSE).
