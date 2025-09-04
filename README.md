Mizzle-AV (LabDash Launcher Helper)

Overview
- Windows WPF UI to define apps and generate labdash:// links for the Lab Dash app.
- Saves mappings to `apps.json` and includes a PowerShell protocol handler so links work without WSH.

Quick Start (Recommended)
- Launch: double-click `Launch-Setup-Visible.cmd`.
- Fill rows (Path, Name, optional Args), then `Validate` → `Save JSON`.
- Test a link: select a row → `Test Link`.
- Open saved data: `Open JSON`.
- View handler log: `Open Log`.
- See a summary: `Apps Report`.

Files
- `LabDash-Setup-Host.ps1`: Main PowerShell host. Loads the XAML and wires buttons.
- `LabDash-Setup-Window.xaml`: WPF layout and styles.
- `Launch-Setup-Visible.cmd`: Entry point for the UI (no admin required).
- `apps.json`: Your saved mappings (append/merge on Save).
- `apps.example.json`: Example data.
- `Handle-LabDash.ps1`: PowerShell URL protocol handler for `labdash://`.
- `Register-LabDash-Protocol.ps1`: Registers the protocol (per‑user).
- `mizzle.ico`, `labdash.ico`: Icons used in the UI.
- `optional/Launch-Setup-*.vbs`: Legacy launchers (only needed if WSH is enabled).

Register the Protocol (per‑user)
1) Open PowerShell in this folder.
2) Register hidden (normal use): `./Register-LabDash-Protocol.ps1`
   - Or keep console open for debugging: `./Register-LabDash-Protocol.ps1 -KeepOpen`
3) Test: `Start-Process 'labdash://open/your-key'`

Logging
- Primary log: `%LOCALAPPDATA%\LabDashLauncher\labdash-handler.log`
- UI button `Open Log` opens the log or creates it if missing.

Tips
- App Name is editable; the link key is a slug of Name.
- `Validate` shows the labdash URL and any Args for each valid row.
- `Save JSON` merges entries by key; existing entries persist unless replaced.
- To remove an entry, edit `apps.json` directly.

Troubleshooting
- If UI changes don’t appear, make sure you’re launching `Launch-Setup-Visible.cmd` from the source folder.
- Check the effective protocol command: `reg query HKCU\Software\Classes\labdash\shell\open\command`.
 - If a link doesn't launch, click `Open Log` right after testing and review the last lines.

License
- Copyright (c) Twinflame Partners
