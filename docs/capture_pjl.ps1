# Capture what the Windows driver ACTUALLY sends to the printer.
#
# The Mac side proved exactly what the bizhub C250i accepts (see bin/km9100auth):
#
#     @PJL SET BOXHOLDTYPE = PRIVATE
#     @PJL SET KMUSERNAME = "<initials>"
#     @PJL SET KMUSERKEY2 = "<pin>"
#     @PJL SET KMCERTSERVTYPE = NUMBER
#     @PJL SET KMCERTSERVNUM = 101
#
# ...and that "@PJL SET KMCOETYPE = 2" in the same header makes the printer
# reject the job as a login error EVEN WITH CORRECT CREDENTIALS.
#
# Windows has no equivalent of the CUPS backend, so the driver writes that
# header itself. This script points a throwaway copy of the print queue at a
# local listener instead of the printer, so we can read the header the driver
# emitted and compare it to the known-good block above.
#
# Usage, in an Administrator PowerShell:
#     irm https://pages.bernting.se/room-business-center-skrivare/capture_pjl.ps1 | iex
# Then print anything to the "PJL_CAPTURE" printer it creates (Notepad, Ctrl+P).
# Fill in the auth dialog exactly as you did when it failed.

param(
    [string]$SourcePrinter = "Room_Business_Center_Olivetti_MF224",
    [string]$CaptureName   = "PJL_CAPTURE",
    [int]   $ListenPort    = 9100,
    [int]   $TimeoutSec    = 300
)

$ErrorActionPreference = "Stop"

# Runtime check rather than #Requires, which is ignored under `irm | iex`.
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "run this from an Administrator PowerShell window."
}

$src = Get-Printer -Name $SourcePrinter -ErrorAction SilentlyContinue
if (-not $src) { throw "printer '$SourcePrinter' not found. Run Get-Printer to see the real name." }
Write-Host "source queue driver: $($src.DriverName)" -ForegroundColor Cyan

# A loopback RAW port. 127.0.0.1 so nothing leaves the machine.
$portName = "PJLCAP_127.0.0.1"
# Remove the printer first (it holds the port), then always recreate the port:
# a port left over from a run with a different -ListenPort keeps its old port
# number, so the spooler would connect somewhere our listener isn't.
if (Get-Printer -Name $CaptureName -ErrorAction SilentlyContinue) {
    Get-PrintJob -PrinterName $CaptureName -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
    Remove-Printer -Name $CaptureName -Confirm:$false
}
if (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue) {
    Remove-PrinterPort -Name $portName -Confirm:$false -ErrorAction SilentlyContinue
}
Add-PrinterPort -Name $portName -PrinterHostAddress "127.0.0.1" -PortNumber $ListenPort
Add-Printer -Name $CaptureName -DriverName $src.DriverName -PortName $portName
Write-Host "created capture queue '$CaptureName' on $portName" -ForegroundColor Green
Write-Host ""
Write-Host "Now print a page to '$CaptureName' (Notepad -> Ctrl+P)." -ForegroundColor Yellow
Write-Host "Waiting up to $TimeoutSec s..." -ForegroundColor Yellow

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $ListenPort)
$listener.Start()
try {
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while (-not $listener.Pending()) {
        if ((Get-Date) -gt $deadline) { throw "no job arrived within $TimeoutSec s." }
        Start-Sleep -Milliseconds 200
    }
    $client = $listener.AcceptTcpClient()
    $stream = $client.GetStream()
    $buf = New-Object byte[] 8192
    $stream.ReadTimeout = 10000
    # Loop: one Read can return a short segment that cuts the PJL header in half.
    $total = 0
    while ($total -lt $buf.Length) {
        $n = try { $stream.Read($buf, $total, $buf.Length - $total) } catch { 0 }
        if ($n -le 0) { break }
        $total += $n
    }
    $head = [System.Text.Encoding]::ASCII.GetString($buf, 0, [Math]::Max($total, 0))
    $client.Close()
} finally {
    $listener.Stop()
}

$outFile = Join-Path $env:TEMP "pjl-header.txt"
$head | Out-File -FilePath $outFile -Encoding ascii

Write-Host ""
Write-Host "===== PJL header the driver sent =====" -ForegroundColor Cyan
($head -split "`n" | Select-Object -First 60) | ForEach-Object { Write-Host $_ }
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "full first 8 KB saved to $outFile"
Write-Host ""

# Verdict against the known-good Mac header.
function Check($label, $ok) {
    if ($ok) { Write-Host "  [ok]   $label" -ForegroundColor Green }
    else     { Write-Host "  [BAD]  $label" -ForegroundColor Red }
}
Write-Host "Comparison with the known-good header:" -ForegroundColor White
Check "KMUSERNAME present (driver is sending a user at all)" ($head -match "KMUSERNAME")
Check "KMUSERKEY2 present (driver is sending the PIN)"       ($head -match "KMUSERKEY2")
Check "KMCOETYPE ABSENT (its presence = guaranteed login error)" ($head -notmatch "KMCOETYPE")
if ($head -match 'KMUSERNAME\s*=\s*"([^"]*)"') { Write-Host "  username in job: '$($Matches[1])'" }
if ($head -match "KMACCOUNT|KMDEPT")           { Write-Host "  [BAD]  account-track fields present - driver is in Account Track mode, not User Authentication" -ForegroundColor Red }


# The job never completed its transfer, so the spooler will retry 127.0.0.1
# forever and can wedge Remove-Printer. Clear it out here rather than leaving
# it as an exercise.
Write-Host "cleaning up the capture queue..." -ForegroundColor DarkGray
Get-PrintJob -PrinterName $CaptureName -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
Remove-Printer -Name $CaptureName -Confirm:$false -ErrorAction SilentlyContinue
Remove-PrinterPort -Name $portName -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "done - your real printer was not touched." -ForegroundColor DarkGray
