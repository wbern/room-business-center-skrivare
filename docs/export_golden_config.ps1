# Capture a verified Windows printer baseline from a machine that prints.
#
# Run this in an Administrator PowerShell on ONE PC that has just printed
# successfully.  It produces two Desktop files:
#
#   * printer-config.dat
#       The safe, deployable global driver configuration (PrinterDriverData +
#       global DevMode). This is what install.ps1 replays on other PCs.
#
#   * olivetti-working-baseline_YYYYMMDD_HHMMSS.zip
#       A shareable support bundle containing the configuration and a redacted
#       report of the queue, driver, port and driver-data fingerprints.
#
# The script deliberately never exports `u` (per-user DevMode). That is where
# initials/PINs are stored. The report includes hashes and lengths for driver
# data values, never their contents, so no PIN belongs in the ZIP.
#
# Run it straight from the web:
#
#   irm https://pages.bernting.se/room-business-center-skrivare/export_golden_config.ps1 | iex

[CmdletBinding()]
param(
    [string]$PrinterName = "ROOM Business Center (Olivetti MF224)",
    # Desktop, not $PSScriptRoot — under `irm | iex` there is no script on disk.
    [string]$OutFile = "$([Environment]::GetFolderPath('Desktop'))\printer-config.dat"
)

$ErrorActionPreference = "Stop"

function Get-Sha256Hex {
    param([byte[]]$Bytes)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace("-", "").ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Get-RedactedValueSummary {
    param($Value)

    if ($null -eq $Value) { return "<null>" }
    if ($Value -is [byte[]]) {
        return "byte[]; bytes=$($Value.Length); sha256=$(Get-Sha256Hex $Value)"
    }

    # This is a fingerprint only. Do not put the setting value in the report.
    $text = [string]$Value
    $bytes = [Text.Encoding]::UTF8.GetBytes($text)
    return "$($Value.GetType().FullName); chars=$($text.Length); sha256=$(Get-Sha256Hex $bytes)"
}

# Not #Requires -RunAsAdministrator: that's ignored when the script is piped
# into iex rather than run from a file, so check at runtime instead.
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    throw "run this from an Administrator PowerShell window."
}

$printer = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
if (-not $printer) { throw "printer '$PrinterName' not found on this machine." }

Write-Host ""
Write-Host "This captures the configuration from THIS working PC." -ForegroundColor Yellow
Write-Host "It does not export initials or PINs." -ForegroundColor Green
Write-Host "Before continuing, print a normal page from '$PrinterName' and confirm it succeeds." -ForegroundColor Yellow
if ((Read-Host "Did that print successfully? (y/n)").Trim().ToLowerInvariant() -ne "y") {
    Write-Host "cancelled — use a machine with a verified successful print." -ForegroundColor Yellow
    exit 0
}

$desktop = [Environment]::GetFolderPath('Desktop')
$outDir = Split-Path -Parent $OutFile
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$workDir = Join-Path $env:TEMP "olivetti-working-baseline_$stamp"
$zipFile = Join-Path $desktop "olivetti-working-baseline_$stamp.zip"
$configInBundle = Join-Path $workDir "printer-config.dat"
$reportFile = Join-Path $workDir "baseline-report.txt"
$bundleReadme = Join-Path $workDir "README.txt"

if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
    # d = PrinterDriverData; g = global DevMode. No `u`: per-user credentials
    # must never be published or copied to another person.
    $p = Start-Process rundll32.exe -Wait -PassThru -ArgumentList `
        "printui.dll,PrintUIEntry", "/Ss", "/n", "`"$PrinterName`"", "/a", "`"$OutFile`"", "d", "g"

    if (-not (Test-Path $OutFile)) {
        throw "export produced no file (rundll32 exit $($p.ExitCode))."
    }
    Copy-Item -LiteralPath $OutFile -Destination $configInBundle -Force

    $driver = Get-PrinterDriver -Name $printer.DriverName -ErrorAction SilentlyContinue
    $port = Get-PrinterPort -Name $printer.PortName -ErrorAction SilentlyContinue
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

    $report = @(
        "ROOM Business Center Olivetti working-PC baseline",
        "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "Verified by operator: successful print immediately before export",
        "",
        "Queue:",
        "  Name: $($printer.Name)",
        "  DriverName: $($printer.DriverName)",
        "  PortName: $($printer.PortName)",
        "  PrinterStatus: $($printer.PrinterStatus)",
        "  Shared: $($printer.Shared)",
        "",
        "Operating system:",
        "  ComputerName: $env:COMPUTERNAME",
        "  Caption: $($os.Caption)",
        "  Version: $($os.Version)",
        "  Build: $($os.BuildNumber)",
        "",
        "Driver details:"
    )

    if ($driver) {
        $report += ($driver | Select-Object Name, Manufacturer, MajorVersion, PrinterEnvironment, InfPath, DriverPath, ConfigFile, DataFile, MonitorName, DefaultDataType, IsPackageAware | Format-List | Out-String).TrimEnd()
    } else {
        $report += "  Unable to read the selected driver."
    }

    $report += ""
    $report += "Port details:"
    if ($port) {
        $report += ($port | Select-Object Name, PrinterHostAddress, PortNumber, Protocol, Queue, SNMPEnabled, SNMPDevIndex | Format-List | Out-String).TrimEnd()
    } else {
        $report += "  Unable to read port '$($printer.PortName)'."
    }

    # The deployable .dat carries the real values. Record only safe fingerprints
    # here so later comparisons can detect a changed or non-applied restore.
    $report += ""
    $report += "PrinterDriverData fingerprints (names, types, lengths and hashes only):"
    $ddPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$PrinterName\PrinterDriverData"
    $providerMembers = @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider", "PSStatus")
    try {
        $driverData = Get-ItemProperty -Path $ddPath -ErrorAction Stop
        $entries = @($driverData.PSObject.Properties | Where-Object { $providerMembers -notcontains $_.Name } | Sort-Object Name)
        if ($entries.Count -eq 0) {
            $report += "  <no driver data values found>"
        } else {
            foreach ($entry in $entries) {
                $report += "  $($entry.Name): $(Get-RedactedValueSummary -Value $entry.Value)"
            }
        }
    } catch {
        $report += "  Unable to read PrinterDriverData: $($_.Exception.Message)"
    }

    $configBytes = [IO.File]::ReadAllBytes($OutFile)
    $report += ""
    $report += "Exported configuration:"
    $report += "  File: printer-config.dat"
    $report += "  Bytes: $($configBytes.Length)"
    $report += "  SHA-256: $(Get-Sha256Hex $configBytes)"
    $report += "  PrintUIEntry exit code: $($p.ExitCode)"
    $report += "  Scope: PrinterDriverData + global DevMode only (d g; never u)"
    $report | Set-Content -LiteralPath $reportFile -Encoding UTF8

    @"
This bundle came from a PC that printed successfully immediately before export.

Contents:
  printer-config.dat  Safe global driver/device configuration for the installer.
  baseline-report.txt Exact driver, port, OS and redacted driver-data fingerprints.

It intentionally excludes each user's initials and PIN. Send this ZIP as-is for
comparison; do not add a raw PJL capture, because that can contain a PIN.
"@ | Set-Content -LiteralPath $bundleReadme -Encoding UTF8

    if (Test-Path $zipFile) { Remove-Item $zipFile -Force }
    Compress-Archive -Path (Join-Path $workDir "*") -DestinationPath $zipFile -Force

    Write-Host ""
    Write-Host "Created the shareable baseline bundle:" -ForegroundColor Green
    Write-Host "  $zipFile" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Send the ZIP file. It contains the safe printer-config.dat plus driver and port evidence, but no initials or PIN." -ForegroundColor White
    Write-Host "Keep the separate Desktop printer-config.dat only if you also need to publish it as the installer default." -ForegroundColor DarkGray
} finally {
    if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue }
}
