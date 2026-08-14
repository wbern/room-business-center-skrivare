# Capture the driver configuration ONCE, so no one else has to do it by hand.
#
# The painful manual step in install.ps1 is Configure tab -> Device Option ->
# User Authentication = On. That setting lives in the queue's PrinterDriverData
# registry blob, it is NOT per-user, and it is identical on every machine that
# talks to this printer. So it can be captured here and replayed everywhere by
# the installer (printui.dll /Sr).
#
# Run this on ONE Windows machine that already prints successfully, then commit
# the resulting file to docs/ so the web installer can fetch it. Run it straight
# from the web in an Administrator PowerShell:
#
#   irm https://pages.bernting.se/room-business-center-skrivare/export_golden_config.ps1 | iex
#
# It drops printer-config.dat on your Desktop; send that file to whoever commits
# it to docs/.
#
# IMPORTANT — do NOT put your initials/PIN in the driver before running this.
# We deliberately export only `d` (PrinterDriverData) and `g` (global DevMode),
# never `u` (per-user DevMode), so no personal credential can end up baked into
# a file we publish on the web. Check the reminder below before you run it.

param(
    [string]$PrinterName = "Room_Business_Center_Olivetti_MF224",
    # Desktop, not $PSScriptRoot — under `irm | iex` there is no script on disk.
    [string]$OutFile     = "$([Environment]::GetFolderPath('Desktop'))\printer-config.dat"
)

$ErrorActionPreference = "Stop"

# Not #Requires -RunAsAdministrator: that's ignored when the script is piped
# into iex rather than run from a file, so check at runtime instead.
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "run this from an Administrator PowerShell window."
}

if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    throw "printer '$PrinterName' not found on this machine."
}

Write-Host ""
Write-Host "Before exporting, on THIS machine make sure:" -ForegroundColor Yellow
Write-Host "  * Printer properties > Configure > Obtain Settings > Auto UNTICKED"
Write-Host "  * Device Option > User Authentication = On"
Write-Host "  * Printing preferences > Basic > Authentication/Account Track:"
Write-Host "      credential fields EMPTY (each user fills in their own)" -ForegroundColor Yellow
Write-Host ""
if ((Read-Host "Ready to export? (y/n)") -ne "y") { Write-Host "cancelled."; exit 0 }

$dir = Split-Path -Parent $OutFile
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

# d = PrinterDriverData (the installable/device options), g = global DevMode.
# No `u` on purpose - see the header.
$p = Start-Process rundll32.exe -Wait -PassThru -ArgumentList `
    "printui.dll,PrintUIEntry", "/Ss", "/n", "`"$PrinterName`"", "/a", "`"$OutFile`"", "d", "g"

if (-not (Test-Path $OutFile)) { throw "export produced no file (rundll32 exit $($p.ExitCode))." }

$size = (Get-Item $OutFile).Length
Write-Host ""
Write-Host "exported $size bytes to $OutFile" -ForegroundColor Green
Write-Host ""
Write-Host "Next: commit that file to docs/ and push, so it publishes to" -ForegroundColor White
Write-Host "  https://pages.bernting.se/room-business-center-skrivare/printer-config.dat" -ForegroundColor Cyan
Write-Host "install.ps1 picks it up automatically and skips the manual Configure step." -ForegroundColor White
