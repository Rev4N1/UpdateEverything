# Update-All

One self-contained PowerShell script that fully updates a Windows PC: per-user apps, Microsoft Store apps, machine-wide apps and Windows Update (including drivers). It doesn't need any PowerShell modules.

## What it does

1. **Self-elevates** into Windows PowerShell 5.1. If you start it from PowerShell 7, it relaunches in 5.1.
2. **User and Store apps.** For every logged-on user, a temporary scheduled task runs `winget upgrade --all --scope user` *as that user, without elevation*. Store (MSIX) apps get updated through the `msstore` source. Each user's output streams live into the elevated window.
3. **Machine apps.** Runs `winget upgrade --all --scope machine` elevated.
4. **Store scan.** Triggers the Store's own update scan through the MDM WMI bridge. It catches apps that winget doesn't know about, and they install in the background.
5. **Windows Update.** Uses the Windows Update Agent COM API directly (`Microsoft.Update.Session`): search, download, then a quiet install, with a result for each update.

At the end it prints a summary table. If an update needs a restart, it can reboot for you.

### Why there's no module

`PSWindowsUpdate` can fail on some machines with `ArgumentException: Value does not fall within the expected range`. The COM API underneath works fine, so the script calls it directly.

### When winget is missing for the admin account

If App Installer was never registered for the elevated profile, the script registers it. If that fails too, it runs `winget.exe` straight from the package folder with its VC++ runtime on `PATH`.

## Usage

```powershell
Unblock-File .\Update-All.ps1
.\Update-All.ps1
```

```powershell
# Skip some Windows updates by title, and reboot automatically if needed
.\Update-All.ps1 -Exclude 'Silverlight','Apple, Inc.' -Reboot

# Only Windows Update, temporarily targeting a feature release
.\Update-All.ps1 -TargetRelease 26H2 -SkipUserApps -SkipMachineApps -SkipStore
```

| Parameter | Description |
|---|---|
| `-Exclude` | Windows Update titles to skip. Matches any part of the title (literal, case-insensitive). |
| `-SkipUserApps` | Don't update per-user or Store apps for logged-on users. |
| `-SkipMachineApps` | Skip the machine-wide winget upgrade. |
| `-SkipStore` | Don't trigger the Store update scan. |
| `-SkipWindowsUpdate` | Skip Windows Update. |
| `-SkipDrivers` | Windows Update installs software updates only, no drivers. |
| `-TargetRelease` | Temporarily set `TargetReleaseVersionInfo` (for example `26H2`) so that an optional or pinned feature update is offered. The previous policy values are restored afterwards. |
| `-IncludeUnknown` | Let winget upgrade packages whose installed version it can't determine. |
| `-UserTimeoutMinutes` | Maximum time for each user's app run. Default: 90. |
| `-Reboot` | If a restart is needed, restart after a 60-second countdown (`shutdown /a` cancels it). |
| `-NoPause` | Don't wait for Enter at the end. Use this for unattended runs. |

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Everything succeeded. |
| `1` | Something failed or only partly succeeded, or another run is already in progress. |
| `3010` | Succeeded, but a reboot is pending and `-Reboot` wasn't given. |

## Footprint

- Writes no logs.
- Uses a temporary folder, `%SystemRoot%\Temp\UpdateAll_<random>`, and temporary tasks under `\UpdateAll\` in Task Scheduler. Both are removed at the end, including after errors or Ctrl+C. If an earlier run was killed, its leftovers are cleaned up on the next run.
- Only one run at a time, enforced by a global mutex.

## Requirements

- Windows 10 or 11, Windows PowerShell 5.1, an administrator account.
- App Installer (winget) for the app phases.

## Notes

- A management agent (an RMM tool or Intune) may pin `TargetReleaseVersionInfo` or re-apply its own Windows Update settings. `-TargetRelease` only changes the pin for the duration of the run.
- MSIX apps provisioned machine-wide (for example an old `Microsoft.Teams.Free`) can fail in the machine phase with `0x80070005`. Update them per user, or exclude them with `winget pin add --id <id>`.

## License

[MIT](LICENSE)
