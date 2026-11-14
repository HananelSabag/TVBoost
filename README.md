# TVBoost

**A Windows tool that diagnoses, cleans and speeds up an Android TV over ADB**, written for the Xiaomi Mi TV 4S (MiTV-MSSP3). One PowerShell file with a WPF window, Hebrew UI.

It runs on the PC and talks to the TV over the home network. It installs nothing on the TV, needs no root and deletes nothing: unwanted apps are only *disabled*, and every change goes to an undo log that **Restore all** reverses.

## What it does

- **Find and connect.** Type the TV's IP, or let it scan the local network for a TV with network debugging on.
- **Diagnosis.** Memory, storage, CPU load, the home app, Wi-Fi link (signal, band, speed) and a ping to the router, summed up as issues by severity.
- **Network check.** A LAN speed test: a 16 MB push from the PC to the TV, reported in Mbps.
- **Task manager.** The apps running now with their memory, and stop the ones you pick.
- **Debloat.** Scans installed packages against a built-in catalog of what's safe to disable on this TV (vendor promotions, analytics, demo apps) and what must never be touched (remote, HDMI, picture and sound, Play Store, Chromecast).
- **Tweaks and launchers.** Faster animations (0.5×), no always-on Wi-Fi / Bluetooth scanning, choosing another home app, opening a launcher's store page.
- **Restore all.** Re-enables every package and resets every setting from the undo log.

Long operations run in background runspaces, so the window never freezes.

## Run it

Needs Windows PowerShell 5.1 and `adb` (Android SDK platform-tools, found automatically). On the TV, turn on developer options and network debugging.

```powershell
powershell -ExecutionPolicy Bypass -File .\TVBoost.ps1
powershell -ExecutionPolicy Bypass -File .\TVBoost.ps1 -SelfTest    # parser checks without a TV
```

Undo logs and the last IP are saved in `data\` and reports in `reports\`, next to the script (both stay local).

> The package catalog is tuned for the Mi TV 4S. On other models, read the list before disabling anything.

## License

[MIT](LICENSE) © 2026 Hananel Sabag
