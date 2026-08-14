# ASCII-ONLY BOOTSTRAPPER - do not put non-ASCII characters in this file.
#
# Why this exists:
# Windows PowerShell 5.1 decodes an HTTP response that carries no charset as
# ISO-8859-1, not UTF-8. GitHub Pages serves .ps1 as application/octet-stream
# with no charset and offers no way to add one. So the obvious
#
#     irm https://.../install.ps1 | iex
#
# turns every Swedish "a-ring/a-diaeresis/o-diaeresis" in install.ps1 into
# mojibake before the script is even parsed. (PowerShell 7.4+ defaults to UTF-8
# and is unaffected, but 5.1 is what ships with Windows and is frozen.)
#
# This file stays pure ASCII, so it survives that misdecode by definition. It
# downloads install.ps1 as raw BYTES and decodes them as UTF-8 itself, then runs
# the result. Language is passed in via the PRINTER_LANG environment variable.
#
#   $env:PRINTER_LANG='sv'; irm https://pages.bernting.se/room-business-center-skrivare/boot.ps1 | iex

# Under `irm | iex` this runs in the CALLER's scope, so setting a preference
# here would leave the user's own console permanently on EAP=Stop and make
# unrelated commands they type later blow up. Save it and put it back.
$bootPrevEAP = $ErrorActionPreference
$ErrorActionPreference = "Stop"

$site = if ($env:PRINTER_SITE) { $env:PRINTER_SITE }
        else { "https://pages.bernting.se/room-business-center-skrivare" }
$url  = if ($env:PRINTER_SCRIPT_URL) { $env:PRINTER_SCRIPT_URL } else { "$site/install.ps1" }

# Windows 10 builds still default to TLS 1.0 for .NET web calls; GitHub requires 1.2+.
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

try {
    # DownloadData returns byte[] - no encoding guesswork anywhere in the path.
    $wc = New-Object System.Net.WebClient
    $bytes = $wc.DownloadData($url)
} catch {
    Write-Host ""
    Write-Host "X couldn't download the installer from $url" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Check the internet connection and try again." -ForegroundColor White
    Read-Host "Press Enter to close"
    $ErrorActionPreference = $bootPrevEAP
    return
}
$ErrorActionPreference = $bootPrevEAP   # install.ps1 sets its own from here on

$src = [Text.Encoding]::UTF8.GetString($bytes)
# A UTF-8 BOM would become a stray character at the start of the script text and
# break parsing, so drop it if the file ever gets saved with one.
if ($src.Length -gt 0 -and $src[0] -eq [char]0xFEFF) { $src = $src.Substring(1) }

Invoke-Expression $src
