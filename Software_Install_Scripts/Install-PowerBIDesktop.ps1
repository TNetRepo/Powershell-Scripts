# ============================================================
# Install-PowerBIDesktop.ps1
# Dynamically resolves and installs the latest Power BI Desktop
# Designed for NinjaOne RMM deployment running as SYSTEM
# ============================================================

$ErrorActionPreference = "Stop"
$installerPath = "$env:TEMP\PBIDesktopSetup_x64.exe"

try {
    # --------------------------------------------------------
    # Step 1: Resolve the current download URL dynamically
    # Microsoft's download page lists the current installer URL
    # --------------------------------------------------------
    Write-Host "Resolving current Power BI Desktop download URL..."

    $downloadPage = Invoke-WebRequest `
        -Uri             "https://www.microsoft.com/en-us/download/details.aspx?id=58494" `
        -UseBasicParsing

    # Extract the direct x64 installer URL from the page content
    $installerUrl = ($downloadPage.Links | Where-Object {
        $_.href -match "PBIDesktopSetup_x64\.exe" -and
        $_.href -match "^https"
    } | Select-Object -First 1).href

    # Fallback: if link extraction fails, use the known redirect
    if (-not $installerUrl) {
        Write-Host "Direct URL extraction failed, attempting redirect resolution..."

        $response = Invoke-WebRequest `
            -Uri             "https://aka.ms/pbidesktopdownload" `
            -UseBasicParsing `
            -MaximumRedirection 0 `
            -ErrorAction     SilentlyContinue

        $installerUrl = $response.Headers.Location

        if (-not $installerUrl -or $installerUrl -notmatch "PBIDesktopSetup_x64") {
            # Final fallback: use the known Microsoft Download Center redirect
            $installerUrl = "https://download.microsoft.com/download/8/8/0/880BCA75-79DD-466A-927D-1ABF1F5454B0/PBIDesktopSetup_x64.exe"
            Write-Host "Using fallback URL."
        }
    }

    Write-Host "Installer URL resolved: $installerUrl"

    # --------------------------------------------------------
    # Step 2: Download the installer
    # --------------------------------------------------------
    Write-Host "Downloading Power BI Desktop installer..."

    # Use WebClient for more reliable large file downloads under SYSTEM context
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($installerUrl, $installerPath)

    if (-not (Test-Path $installerPath)) {
        throw "Installer file not found after download — download may have failed silently."
    }

    $fileSize = (Get-Item $installerPath).Length / 1MB
    Write-Host "Download complete. File size: $([math]::Round($fileSize, 1)) MB"

    # --------------------------------------------------------
    # Step 3: Silent install
    # --------------------------------------------------------
    Write-Host "Installing Power BI Desktop silently..."

    $installArgs = "-quiet -norestart ACCEPT_EULA=1"
    $process = Start-Process `
        -FilePath  $installerPath `
        -ArgumentList $installArgs `
        -Wait `
        -PassThru `
        -NoNewWindow

    # Check exit code
    switch ($process.ExitCode) {
        0       { Write-Host "Installation completed successfully." }
        3010    { Write-Host "Installation completed successfully. A reboot is required to finish." }
        1602    { throw "Installation was cancelled by the user or policy." }
        1603    { throw "Fatal error during installation. Check Windows Event Log for details." }
        default { throw "Installation exited with unexpected code: $($process.ExitCode)" }
    }

    # --------------------------------------------------------
    # Step 4: Verify installation
    # --------------------------------------------------------
    Write-Host "Verifying installation..."

    $pbiInstalled = Get-ItemProperty `
        -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", `
              "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match "Microsoft Power BI Desktop" } |
        Select-Object -First 1

    if ($pbiInstalled) {
        Write-Host "Verification OK: $($pbiInstalled.DisplayName) version $($pbiInstalled.DisplayVersion) is installed."
    } else {
        Write-Host "Warning: Installation appeared to succeed but could not verify via registry. Manual check recommended."
    }
}
catch {
    Write-Error "Power BI Desktop installation failed: $_"
    exit 1
}
finally {
    # Clean up installer regardless of outcome
    if (Test-Path $installerPath) {
        Remove-Item $installerPath -Force
        Write-Host "Installer cleaned up."
    }
}
