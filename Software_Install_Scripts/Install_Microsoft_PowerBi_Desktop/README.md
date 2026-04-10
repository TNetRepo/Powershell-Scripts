# Install-PowerBIDesktop

A PowerShell script that dynamically resolves and silently installs the latest version of Microsoft Power BI Desktop (x64). Designed for deployment via NinjaOne RMM running scripts as SYSTEM, but compatible with any RMM or manual execution.

---

## What the Script Does

1. **Checks if Power BI Desktop is already installed** by querying the registry. If a current installation is found, the script logs the installed version and exits cleanly with no further action. This prevents unnecessary re-downloads and re-installs on machines that are already provisioned.

2. **Dynamically resolves the current installer URL** by parsing Microsoft's Power BI download page to extract the latest x64 installer link. If page parsing fails, it falls back to the `aka.ms/pbidesktopdownload` redirect. If that also fails, a known static Microsoft CDN URL is used as a final fallback — ensuring the script has three layers of URL resolution before giving up.

3. **Downloads the installer** using `System.Net.WebClient`, which is more reliable than `Invoke-WebRequest` for large binary file downloads under SYSTEM context.

4. **Silently installs Power BI Desktop** with no user interaction required and no forced reboot. Exit codes are evaluated and reported, including distinguishing a clean success from a success-with-reboot-required (exit code 3010).

5. **Verifies the installation** by re-querying the registry and logging the installed display name and version number.

6. **Cleans up** the downloaded installer from `%TEMP%` regardless of whether the installation succeeded or failed.

---

## Requirements

- Windows 10 1809 or later / Windows 11
- PowerShell 5.1 or later
- Direct internet access to Microsoft's download servers
- No additional PowerShell modules required

---

## Deployment via NinjaOne

1. In NinjaOne, create a new script and paste the contents of `Install-PowerBIDesktop.ps1`
2. Set the script to run as **SYSTEM**
3. Push to target devices individually or as a group

No variables need to be configured before deployment — the script is ready to run as-is.

> **Note on runtime:** The Power BI Desktop installer is approximately 660MB. On a typical business internet connection, the download and install process takes between 15-40 minutes on machines where Power BI is not already installed. On machines where it is already installed, the script completes in seconds.

---

## Variables

There are no variables that require configuration before running this script. The installer URL is resolved dynamically at runtime.

If you need to pin a specific version rather than always installing the latest, replace the URL resolution logic in Step 2 with a direct hardcoded URL to the desired installer version from Microsoft's Download Center.

---

## Output / Logging

The script writes progress to stdout at each step, which is captured in NinjaOne's activity log. Example output for a fresh install:

```
Checking if Power BI Desktop is already installed...
Power BI Desktop not found. Proceeding with installation...
Resolving current Power BI Desktop download URL...
Installer URL resolved: https://download.microsoft.com/download/.../PBIDesktopSetup_x64.exe
Downloading Power BI Desktop installer...
Download complete. File size: 660.8 MB
Installing Power BI Desktop silently...
Installation completed successfully.
Verifying installation...
Verification OK: Microsoft Power BI Desktop (x64) version 2.152.1279.0 is installed.
Installer cleaned up.
```

Example output when Power BI Desktop is already installed:

```
Checking if Power BI Desktop is already installed...
Power BI Desktop is already installed: Microsoft Power BI Desktop (x64) version 2.152.1279.0. No action required.
```

---

## Exit Codes

| Code | Meaning |
|---|---|
| 0 | Success — installed successfully or already installed |
| 1 | Failure — script caught an error, details logged to output |
| 3010 | Success — installation complete but a reboot is required |

---

## Execution Policy

If running manually rather than via RMM, run the following first to allow script execution for the current session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```
