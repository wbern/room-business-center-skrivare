# One-shot office-printer installer for Windows.
#
# RUN IT VIA boot.ps1, NOT DIRECTLY:
#
#   $env:PRINTER_LANG='sv'; irm https://pages.bernting.se/room-business-center-skrivare/boot.ps1 | iex
#
# This file contains Swedish text, and Windows PowerShell 5.1 decodes a
# charset-less HTTP response as ISO-8859-1 — so fetching THIS file with `irm`
# directly turns every a-ring/a-diaeresis/o-diaeresis into mojibake. boot.ps1 is
# pure ASCII, downloads this as bytes and decodes it as UTF-8. See boot.ps1.
#
# It installs the Olivetti d-Copia MF224 / Konica Minolta bizhub C250i
# (192.168.9.15) with per-user authentication that actually works.
#
# HOW AUTH ACTUALLY WORKS ON THIS PRINTER
# The queue is a RAW TCP/9100 port — a bare socket with no auth challenge, so
# Windows Credential Manager is never consulted for it. The credentials have to
# travel INSIDE the job, as a PJL header the printer reads before the
# PostScript. The Mac side proves the exact block the bizhub accepts (see
# bin/km9100auth), where a CUPS backend injects it:
#
#     @PJL SET KMUSERNAME = "<initials>"
#     @PJL SET KMUSERKEY2 = "<pin>"
#
# Windows has no CUPS backend. The Olivetti/KM driver writes that header
# itself, but ONLY if two separate things are set, and both are GUI settings:
#
#   a. Printer PROPERTIES -> Configure tab -> Device Option ->
#      User Authentication = On. This is how the driver learns the device wants
#      credentials at all. Because we add the queue with Add-Printer instead of
#      Olivetti's own installer, the driver never interrogates the device, so
#      this defaults to None — and with None the driver emits a job with NO auth
#      info, which the C250i discards by default ("Restrict" is the factory
#      setting for such jobs). Greyed out until you untick Auto under
#      "Obtain Settings...".
#   b. Printing PREFERENCES -> Basic tab -> Authentication/Account Track ->
#      Recipient User + the initials/PIN. This is the credential itself.
#      (Note (a) and (b) live in two DIFFERENT windows.)
#
# Miss (a) and (b) is silently ignored — the job still goes out anonymous and
# the printer deletes it ("Radering av fel").
#
# Konica Minolta exposes no supported way to set either from a script — they
# live in the driver's private, undocumented DEVMODE/PrinterDriverData blobs,
# and the Driver Packaging Utility has no silent switches. But (a) is the same
# on every machine and holds no personal data, so we cheat: it's captured once
# with export_golden_config.ps1 and replayed here in step 11b via
# printui.dll /Sr. That leaves only (b), the user's own PIN, which genuinely
# has to be typed once — step 12 opens the right window for it.

[CmdletBinding()]
param(
    [string]$Username,
    [string]$Password,
    [string]$PrinterIP   = "192.168.9.15",
    [int]   $PrinterPort = 9100,
    [string]$PrinterName = "ROOM Business Center (Olivetti MF224)",
    [ValidateSet("en","sv","")]
    [string]$Lang = "",
    [switch]$NoTest
)

# ---- CONFIG (edit to match where the files are hosted) ----------------------
$Site      = if ($env:PRINTER_SITE)       { $env:PRINTER_SITE }       else { "https://pages.bernting.se/room-business-center-skrivare" }
$BootUrl   = if ($env:PRINTER_BOOT_URL)   { $env:PRINTER_BOOT_URL }   else { "$Site/boot.ps1" }
$DriverUrl = if ($env:PRINTER_DRIVER_URL) { $env:PRINTER_DRIVER_URL } else { "$Site/printer-driver-win-x64.zip" }
$ConfigUrl = if ($env:PRINTER_CONFIG_URL) { $env:PRINTER_CONFIG_URL } else { "$Site/printer-config.dat" }
$DriverInf = "KOAWNAA_.inf"                       # INF at the root of the zip
$DriverName = "Generic Universal PS v3.9.12"      # name the Olivetti PS build registers as
# -----------------------------------------------------------------------------

# Same caveat as boot.ps1: under `irm | iex` this is the caller's scope, so the
# non-elevated branch below restores this before handing the console back. The
# elevated window is disposable, so it doesn't need to.
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = "Stop"

# ---- Language ---------------------------------------------------------------
# Priority: -Lang parameter, then PRINTER_LANG (set by the web page's toggle and
# forwarded across the elevation hop), then the user's own culture, then English.
if (-not $Lang) {
    if ($env:PRINTER_LANG -in "sv","en") {
        $Lang = $env:PRINTER_LANG
    } elseif ((Get-Culture).TwoLetterISOLanguageName -eq "sv") {
        $Lang = "sv"
    } else {
        $Lang = "en"
    }
}

# Strings live inline as a hashtable-of-hashtables: no second HTTP fetch to fail,
# no $PSScriptRoot (there is no script on disk under `irm | iex`), and the whole
# table travels with the script across the elevation hop for free.
$MSG = @{
en = @{
    header          = "  Office Printer Setup  -  Olivetti MF224"
    header_short    = "  Office Printer Setup"
    asking_admin    = "asking for administrator access (Windows will pop up a 'Yes/No' prompt)..."
    creds_again     = "  (the new window will ask for your initials and PIN again)"
    continues       = "Setup continues in the new window that just opened."
    no_window_1     = "If no window appeared, right-click PowerShell, choose"
    no_window_2     = "'Run as administrator', and paste the same command there."
    uac_declined    = "administrator access was declined. The printer can't be installed without it. Re-run and click 'Yes' on the prompt."
    trap_title      = "X something went wrong and setup stopped:"
    trap_line       = "  line {0}: {1}"
    trap_advice_1   = "  You can safely re-run the same command."
    trap_advice_2   = "  If it keeps failing, send the red lines above to whoever set this up."
    press_close     = "Press Enter to close"
    checking_reach  = "checking the printer is reachable ({0}:{1})..."
    reachable       = "printer is reachable"
    not_reachable   = "can't reach the printer at {0}. Connect to the office Wi-Fi/network and run this again. (Nothing has been changed.)"
    ask_login       = "Enter your printer login (the initials + 4-digit PIN registered at the printer):"
    ask_initials    = "  Initials (e.g. abc)"
    ask_pin         = "  PIN (e.g. 1234)"
    need_both       = "initials and PIN are both required."
    reg_fix         = "applying the print-auth registry fix..."
    reg_done        = "registry fix applied (RpcAuthnLevelPrivacyEnabled = 0)"
    storing_login   = "storing your printer login..."
    cmdkey_failed   = "couldn't store the credentials (cmdkey failed)."
    login_stored    = "login stored for {0} (user: {1})"
    spooler         = "restarting the Print Spooler..."
    spooler_done    = "spooler restarted"
    driver_present  = "printer driver already installed ({0}) - skipping the 52 MB download"
    driver_dl       = "downloading the printer driver (one-time, ~52 MB)..."
    disk_need       = "only {0} MB free on {2}, and the driver needs about {1} MB while it installs."
    disk_scan       = "looking for files Windows can safely delete..."
    disk_t_wintemp  = "Windows temporary files"
    disk_t_usertemp = "your temporary files"
    disk_t_update   = "downloaded Windows Update files (Windows re-downloads if needed)"
    disk_t_do       = "update files cached for other PCs on the network"
    disk_total      = "  Deleting these frees about {0} MB. None of it is yours - Windows recreates it as needed."
    disk_ask        = "  Delete them? (y/n)"
    disk_cleaning   = "cleaning up..."
    disk_freed      = "freed {0} MB - {1} MB now free"
    disk_skipped    = "left them alone."
    disk_nothing    = "found nothing that's safe to delete automatically."
    disk_bin_ask    = "  Still short. Empty the Recycle Bin too? This permanently deletes what's in it. (y/n)"
    disk_full       = "not enough free space on {2} - only {0} MB left, and installing the driver needs about {1} MB free while it works. Delete some files (Downloads, Recycle Bin, old videos) or run Disk Cleanup, then run this again. Nothing has been changed."
    disk_full_late  = "Windows ran out of disk space while installing the driver ({0} MB free on {1}). Free up about 1 GB - empty the Recycle Bin, clear Downloads, or run Disk Cleanup - then run this again. Nothing else is wrong with the setup."
    driver_dl_fail  = "couldn't download the driver from {0}. Check your connection and try again."
    driver_corrupt  = "driver download looks corrupt ({0} missing). Try again."
    driver_dl_ok    = "driver downloaded"
    driver_install  = "installing the Universal PS driver..."
    driver_register = "registering the driver with Windows..."
    driver_registered = "driver registered as: {0}"
    driver_tried    = "  Names tried:"
    pnputil_said    = "  What pnputil reported:"
    setupapi_said   = "  What Windows logged about this driver:"
    driver_not_staged = "Windows refused to stage the driver into its driver store - trying other routes."
    driver_fallback = "trying the alternative driver install route..."
    pnputil_warn    = "pnputil returned {0} while adding the driver; continuing."
    driver_none     = "the driver installed but Windows didn't register a Universal PS driver. See the list above."
    drivers_known   = "  Drivers Windows currently knows about:"
    driver_ok       = "driver installed: {0}"
    queue_ok        = "the printer is already set up correctly - keeping it (your saved login stays)"
    removed_old     = "removed an old copy of the printer"
    dupes_found     = "found {0} leftover printer(s) pointing at the same printer ({1}) - removing them:"
    dupes_explain_1 = "  They have no login, so anything sent to them is thrown away by the printer."
    dupes_removed   = "removed {0}"
    dupes_fail      = "couldn't remove {0} - {1}"
    port_making     = "creating the printer port..."
    port_reuse      = "reusing the existing printer port ({0})"
    port_ok         = "port created ({0})"
    adding          = "adding the printer..."
    avail_drivers   = "  Available PostScript drivers:"
    add_failed      = "couldn't add the printer with driver '{0}'. See the list above."
    added           = "printer added: {0}"
    bidi_ok         = "bidirectional / SNMP enabled"
    bidi_fail       = "couldn't enable bidirectional/SNMP (not fatal) - {0}"
    default_ok      = "set as your default printer"
    cfg_looking     = "looking for the shared printer configuration..."
    cfg_timeout     = "the configuration restore didn't finish (it may have shown a dialog)."
    cfg_runfail     = "couldn't run the configuration restore."
    cfg_ok          = "driver configured automatically (no Configure tab needed)"
    cfg_notake      = "the shared configuration didn't take - you'll be walked through it instead."
    cfg_none        = "no shared configuration published yet - you'll be walked through it instead."
    wrong_user_1    = "this window is running as {0}, but you're logged in as {1}."
    wrong_user_2    = "The printer login is saved per Windows user, so it must be entered from"
    wrong_user_3    = "{0}'s own session - not here."
    wrong_user_4    = "  Close this window, then as {0} open:"
    wrong_user_5    = "    Settings > Bluetooth & devices > Printers & scanners > {0}"
    wrong_user_6    = "  Under 'Printing preferences', Basic tab, click Authentication/Account Track,"
    wrong_user_7    = "  select 'Recipient User' and enter the initials and PIN."
    wrong_user_8    = "  First, under 'Printer properties', Configure tab: click Obtain Settings,"
    wrong_user_9    = "  untick Auto, then set Device Option > User Authentication to On."
    last_title      = " ONE LAST STEP - and it's the important one"
    last_why_1      = "The printer only accepts jobs that carry your login inside them,"
    last_why_2      = "and Windows won't let a script type your PIN in for you."
    partA_title     = "PART A - tell the driver the printer wants a login"
    partA_note      = "  (skipping this makes Part B do nothing at all)"
    partA_open      = "  A window titled 'Printer properties' will open."
    partA_1         = "  * Go to the  Configure  tab."
    partA_2         = "  * Click  Obtain Settings...  and UNTICK  Auto , then OK."
    partA_3         = "  * In the  Device Option  list, select  User Authentication"
    partA_4         = "    and set  Setting  to  On . Click  Apply , then  OK ."
    partB_title     = "PART B - put your login in"
    partB_only      = "Only one thing left - your login. Everything else is already set."
    partB_open      = "  A window titled 'Printing preferences' will open."
    partB_1         = "  * Go to the  Basic  tab."
    partB_2         = "  * Click  Authentication/Account Track..."
    partB_3         = "  * Under  User Authentication , select  Recipient User , then enter:"
    partB_user      = "       User Name:  {0}"
    partB_pin       = "       Password:   (your PIN)"
    partB_4         = "  * If there's a  Save Settings  box, tick it. If there isn't, that's fine."
    partB_5         = "  * Click  Verify  if it's there, then  OK  on every window."
    verify_note_1   = "Verify can fail even with the right login (it needs SNMP, which some"
    verify_note_2   = "networks block). The real test is the page that prints in a moment."
    press_open_A    = "Press Enter to open the first window (Printer properties)"
    press_open_B    = "Press Enter to open the login window (Printing preferences)"
    dialog_closed   = "settings window closed"
    dialog_fail     = "couldn't open the window automatically. Open Settings > Printers, click {0}, and follow the steps above."
    press_done      = "Press Enter once you've entered your login and clicked OK"
    test_sending    = "sending a test page..."
    test_sent       = "test page sent"
    test_fail       = "couldn't send the test page automatically - try Ctrl+P in any app."
    done_title      = " Setup finished - now check the printer"
    done_1          = "A test page is on its way. Go and collect it."
    done_2          = "IF A PAGE COMES OUT: you're done. Print with Ctrl+P and choose:"
    done_3          = "IF NOTHING COMES OUT within a minute, the login step didn't stick."
    done_4          = "Re-run this command and take Part B slowly - that's the usual cause."
    test_line_1     = "Printer setup complete - user: {0}"
    test_line_2     = "If you can read this, {0} is working."
}
sv = @{
    header          = "  Installation av kontorsskrivaren  -  Olivetti MF224"
    header_short    = "  Installation av kontorsskrivaren"
    asking_admin    = "begär administratörsbehörighet (Windows visar en Ja/Nej-ruta)..."
    creds_again     = "  (det nya fönstret frågar efter dina initialer och din PIN igen)"
    continues       = "Installationen fortsätter i det nya fönstret som just öppnades."
    no_window_1     = "Om inget fönster dök upp: högerklicka på PowerShell, välj"
    no_window_2     = "'Kör som administratör' och klistra in samma kommando där."
    uac_declined    = "administratörsbehörighet nekades. Skrivaren kan inte installeras utan den. Kör igen och klicka 'Ja'."
    trap_title      = "X något gick fel och installationen avbröts:"
    trap_line       = "  rad {0}: {1}"
    trap_advice_1   = "  Du kan lugnt köra samma kommando igen."
    trap_advice_2   = "  Om det fortsätter strula, skicka de röda raderna ovan till den som satte upp det här."
    press_close     = "Tryck Enter för att stänga"
    checking_reach  = "kontrollerar att skrivaren svarar ({0}:{1})..."
    reachable       = "skrivaren svarar"
    not_reachable   = "når inte skrivaren på {0}. Anslut till kontorets wifi/nätverk och kör igen. (Inget har ändrats.)"
    ask_login       = "Ange din skrivarinloggning (initialerna + den 4-siffriga PIN-koden som är registrerad i skrivaren):"
    ask_initials    = "  Initialer (t.ex. abc)"
    ask_pin         = "  PIN (t.ex. 1234)"
    need_both       = "både initialer och PIN krävs."
    reg_fix         = "gör registerfixen för utskriftsinloggning..."
    reg_done        = "registerfix genomförd (RpcAuthnLevelPrivacyEnabled = 0)"
    storing_login   = "sparar din skrivarinloggning..."
    cmdkey_failed   = "kunde inte spara inloggningen (cmdkey misslyckades)."
    login_stored    = "inloggning sparad för {0} (användare: {1})"
    spooler         = "startar om utskriftshanteraren..."
    spooler_done    = "utskriftshanteraren omstartad"
    driver_present  = "skrivardrivrutinen finns redan ({0}) - hoppar över nedladdningen på 52 MB"
    driver_dl       = "laddar ner skrivardrivrutinen (engångsjobb, ca 52 MB)..."
    disk_need       = "bara {0} MB ledigt på {2}, och drivrutinen behöver ungefär {1} MB under installationen."
    disk_scan       = "letar efter filer som Windows kan ta bort utan risk..."
    disk_t_wintemp  = "tillfälliga Windows-filer"
    disk_t_usertemp = "dina tillfälliga filer"
    disk_t_update   = "nedladdade Windows Update-filer (Windows laddar ner igen vid behov)"
    disk_t_do       = "uppdateringsfiler som cachats åt andra datorer i nätverket"
    disk_total      = "  Att ta bort dem frigör ungefär {0} MB. Inget av det är ditt - Windows skapar det på nytt vid behov."
    disk_ask        = "  Ta bort dem? (j/n)"
    disk_cleaning   = "rensar..."
    disk_freed      = "frigjorde {0} MB - {1} MB ledigt nu"
    disk_skipped    = "lät dem vara."
    disk_nothing    = "hittade inget som kan tas bort automatiskt utan risk."
    disk_bin_ask    = "  Fortfarande för lite. Töm Papperskorgen också? Det raderar innehållet permanent. (j/n)"
    disk_full       = "för lite ledigt utrymme på {2} - bara {0} MB kvar, och drivrutinen behöver ungefär {1} MB ledigt under installationen. Ta bort några filer (Hämtade filer, Papperskorgen, gamla videor) eller kör Diskrensning och kör sedan det här igen. Inget har ändrats."
    disk_full_late  = "Windows fick slut på diskutrymme när drivrutinen installerades ({0} MB ledigt på {1}). Frigör ungefär 1 GB - töm Papperskorgen, rensa Hämtade filer eller kör Diskrensning - och kör sedan det här igen. Inget annat är fel med installationen."
    driver_dl_fail  = "kunde inte ladda ner drivrutinen från {0}. Kontrollera uppkopplingen och försök igen."
    driver_corrupt  = "nedladdningen ser trasig ut ({0} saknas). Försök igen."
    driver_dl_ok    = "drivrutinen nedladdad"
    driver_install  = "installerar Universal PS-drivrutinen..."
    driver_register = "registrerar drivrutinen i Windows..."
    driver_registered = "drivrutinen registrerad som: {0}"
    driver_tried    = "  Namn som testades:"
    pnputil_said    = "  Vad pnputil rapporterade:"
    setupapi_said   = "  Vad Windows loggade om drivrutinen:"
    driver_not_staged = "Windows vägrade lägga drivrutinen i sitt drivrutinsarkiv - provar andra vägar."
    driver_fallback = "provar den alternativa installationsvägen för drivrutinen..."
    pnputil_warn    = "pnputil svarade {0} när drivrutinen lades till; fortsätter."
    driver_none     = "drivrutinen installerades men Windows registrerade ingen Universal PS-drivrutin. Se listan ovan."
    drivers_known   = "  Drivrutiner som Windows känner till just nu:"
    driver_ok       = "drivrutin installerad: {0}"
    queue_ok        = "skrivaren är redan korrekt uppsatt - behåller den (din sparade inloggning ligger kvar)"
    removed_old     = "tog bort en gammal kopia av skrivaren"
    dupes_found     = "hittade {0} kvarglömda skrivare som pekar mot samma skrivare ({1}) - tar bort dem:"
    dupes_explain_1 = "  De saknar inloggning, så allt som skickas till dem slängs av skrivaren."
    dupes_removed   = "tog bort {0}"
    dupes_fail      = "kunde inte ta bort {0} - {1}"
    port_making     = "skapar skrivarporten..."
    port_reuse      = "återanvänder den befintliga skrivarporten ({0})"
    port_ok         = "port skapad ({0})"
    adding          = "lägger till skrivaren..."
    avail_drivers   = "  Tillgängliga PostScript-drivrutiner:"
    add_failed      = "kunde inte lägga till skrivaren med drivrutinen '{0}'. Se listan ovan."
    added           = "skrivare tillagd: {0}"
    bidi_ok         = "dubbelriktad kommunikation / SNMP aktiverat"
    bidi_fail       = "kunde inte aktivera dubbelriktad kommunikation/SNMP (inte kritiskt) - {0}"
    default_ok      = "vald som din standardskrivare"
    cfg_looking     = "letar efter den gemensamma skrivarkonfigurationen..."
    cfg_timeout     = "återställningen av konfigurationen blev inte klar (den kan ha visat en dialogruta)."
    cfg_runfail     = "kunde inte köra återställningen av konfigurationen."
    cfg_ok          = "drivrutinen konfigurerad automatiskt (ingen Configure-flik behövs)"
    cfg_notake      = "den gemensamma konfigurationen fastnade inte - du får göra den för hand i stället."
    cfg_none        = "ingen gemensam konfiguration är publicerad än - du får göra den för hand i stället."
    wrong_user_1    = "det här fönstret körs som {0}, men du är inloggad som {1}."
    wrong_user_2    = "Skrivarinloggningen sparas per Windows-användare, så den måste anges från"
    wrong_user_3    = "{0}s egen session - inte här."
    wrong_user_4    = "  Stäng det här fönstret och öppna, som {0}:"
    wrong_user_5    = "    Inställningar > Bluetooth och enheter > Skrivare och skannrar > {0}"
    wrong_user_6    = "  Under 'Utskriftsinställningar', fliken Basic: klicka Authentication/Account Track,"
    wrong_user_7    = "  välj 'Recipient User' och ange initialerna och PIN-koden."
    wrong_user_8    = "  Först, under 'Egenskaper för skrivare', fliken Configure: klicka Obtain Settings,"
    wrong_user_9    = "  bocka ur Auto och sätt Device Option > User Authentication till On."
    last_title      = " ETT SISTA STEG - och det är det viktiga"
    last_why_1      = "Skrivaren tar bara emot utskrifter som bär med sig din inloggning,"
    last_why_2      = "och Windows låter inte ett skript skriva in din PIN åt dig."
    partA_title     = "DEL A - tala om för drivrutinen att skrivaren kräver inloggning"
    partA_note      = "  (hoppar du över det här gör Del B ingen nytta alls)"
    partA_open      = "  Ett fönster som heter 'Egenskaper för skrivare' öppnas."
    partA_1         = "  * Gå till fliken  Configure ."
    partA_2         = "  * Klicka  Obtain Settings...  och BOCKA UR  Auto , sedan OK."
    partA_3         = "  * I listan  Device Option , välj  User Authentication"
    partA_4         = "    och sätt  Setting  till  On . Klicka  Apply , sedan  OK ."
    partB_title     = "DEL B - fyll i din inloggning"
    partB_only      = "Bara en sak kvar - din inloggning. Allt annat är redan klart."
    partB_open      = "  Ett fönster som heter 'Utskriftsinställningar' öppnas."
    partB_1         = "  * Gå till fliken  Basic ."
    partB_2         = "  * Klicka  Authentication/Account Track..."
    partB_3         = "  * Under  User Authentication , välj  Recipient User  och fyll i:"
    partB_user      = "       User Name:  {0}"
    partB_pin       = "       Password:   (din PIN)"
    partB_4         = "  * Finns rutan  Save Settings , bocka i den. Finns den inte gör det inget."
    partB_5         = "  * Klicka  Verify  om knappen finns, sedan  OK  i alla fönster."
    verify_note_1   = "Verify kan misslyckas även med rätt inloggning (den behöver SNMP, som"
    verify_note_2   = "vissa nätverk blockerar). Det riktiga testet är sidan som skrivs ut strax."
    press_open_A    = "Tryck Enter för att öppna det första fönstret (Egenskaper för skrivare)"
    press_open_B    = "Tryck Enter för att öppna inloggningsfönstret (Utskriftsinställningar)"
    dialog_closed   = "inställningsfönstret stängt"
    dialog_fail     = "kunde inte öppna fönstret automatiskt. Öppna Inställningar > Skrivare, klicka {0} och följ stegen ovan."
    press_done      = "Tryck Enter när du har fyllt i din inloggning och klickat OK"
    test_sending    = "skickar en testsida..."
    test_sent       = "testsida skickad"
    test_fail       = "kunde inte skicka testsidan automatiskt - prova Ctrl+P i valfritt program."
    done_title      = " Installationen klar - gå och kolla skrivaren"
    done_1          = "En testsida är på väg. Gå och hämta den."
    done_2          = "KOMMER DET UT EN SIDA: då är du klar. Skriv ut med Ctrl+P och välj:"
    done_3          = "KOMMER DET INTE UT NÅGOT inom en minut fastnade inte inloggningen."
    done_4          = "Kör kommandot igen och ta Del B långsamt - det är den vanliga orsaken."
    test_line_1     = "Skrivarinstallation klar - användare: {0}"
    test_line_2     = "Kan du läsa det här fungerar {0}."
}
}

function T {
    # NB: the second parameter must NOT be called $Args - that's a PowerShell
    # automatic variable, and naming it so makes the binding silently fail,
    # printing literal "{0}" placeholders in every formatted message.
    param([string]$Key, [object[]]$Values)
    $s = $MSG[$Lang][$Key]
    if (-not $s) { $s = $MSG["en"][$Key] }   # fall back to English on a missing key
    if ($Values) { return ($s -f $Values) }
    return $s
}

# ---- Banner -----------------------------------------------------------------
# Deliberately pure ASCII: it has to look right in the legacy conhost raster
# fonts too, where box-drawing characters turn into rubble. "ROOM" was generated
# with figlet's "standard" font; the sub-line is letter-spaced plain text, which
# stays legible at this size in a way a second figlet line does not.
# Single-quoted here-string so nothing inside is interpolated or escaped.
$BANNER = @'
       ___________________________
      |  _______________________  |
      | |                       | |
      | |_______________________| |
      |___________________________|
      |  [][][]             (o)   |
      |___________________________|
        \_______________________/
'@
$WORDMARK = @'
       ____   ___   ___  __  __
      |  _ \ / _ \ / _ \|  \/  |
      | |_) | | | | | | | |\/| |
      |  _ <| |_| | |_| | |  | |
      |_| \_\\___/ \___/|_|  |_|
'@

function Show-Banner {
    Write-Host ""
    foreach ($line in $BANNER  -split "`n") { Write-Host $line.TrimEnd() -ForegroundColor Cyan }
    Write-Host ""
    foreach ($line in $WORDMARK -split "`n") { Write-Host $line.TrimEnd() -ForegroundColor White }
    Write-Host "      B U S I N E S S   C E N T E R" -ForegroundColor DarkCyan
    Write-Host ""
}

function Info($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function Ok($m)   { Write-Host "  $([char]0x2713) $m" -ForegroundColor Green }
function Warn($m) { Write-Host "  ! $m" -ForegroundColor Yellow }
function Die($m)  { Write-Host "X $m" -ForegroundColor Red; Read-Host "`n$(T 'press_close')"; exit 1 }

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

# --- Self-elevate ------------------------------------------------------------
# The elevated window re-runs the ASCII bootstrapper rather than a re-downloaded
# copy of this file. Downloading to %TEMP% and launching it with -File was
# fragile: any failure (AV quarantine, a stale/locked file, an encoding or
# execution-policy hiccup) killed the new window instantly with nothing on
# screen, which is what it did in the field.
#
# The command string is kept pure ASCII: it becomes a -Command argument, and
# non-ASCII there would have to survive yet another encoding layer. Localized
# output only happens after boot.ps1 has decoded this file properly.
#
# Env vars are NOT inherited usefully through ShellExecute's runas verb, so the
# non-secret settings are re-set explicitly inside the command. Credentials are
# deliberately excluded: they'd sit in the elevated process's command line,
# readable via Win32_Process.CommandLine and captured by 4688 audit logging.
if (-not (Test-Admin)) {
    Show-Banner
    Write-Host (T 'header_short') -ForegroundColor White
    Write-Host ""
    Info (T 'asking_admin')
    if ($Username -or $Password) { Write-Host (T 'creds_again') -ForegroundColor DarkGray }

    $fwd = "`$env:PRINTER_LANG='$Lang'; `$env:PRINTER_SITE='$Site';"
    if ($env:PRINTER_SCRIPT_URL) { $fwd += " `$env:PRINTER_SCRIPT_URL='$($env:PRINTER_SCRIPT_URL)';" }
    if ($env:PRINTER_DRIVER_URL) { $fwd += " `$env:PRINTER_DRIVER_URL='$($env:PRINTER_DRIVER_URL)';" }
    if ($env:PRINTER_CONFIG_URL) { $fwd += " `$env:PRINTER_CONFIG_URL='$($env:PRINTER_CONFIG_URL)';" }
    $inner = "$fwd try { irm '$BootUrl' | iex } catch { Write-Host ''; Write-Host 'X the installer could not start:' -ForegroundColor Red; Write-Host `$_ -ForegroundColor Red; Read-Host 'Press Enter to close' }"

    try {
        Start-Process powershell.exe -Verb RunAs -ArgumentList @(
            "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "`"$inner`""
        )
    } catch {
        # NOT Die: this runs in the user's own console, and `exit` there would
        # close their whole window rather than just this script.
        Write-Host ""
        Write-Host "X $(T 'uac_declined')" -ForegroundColor Red
        $ErrorActionPreference = $prevEAP
        return
    }
    Write-Host ""
    Write-Host (T 'continues') -ForegroundColor White
    Write-Host (T 'no_window_1') -ForegroundColor DarkGray
    Write-Host (T 'no_window_2') -ForegroundColor DarkGray
    $ErrorActionPreference = $prevEAP
    return
}

# ===== From here on we are elevated ==========================================
# $ErrorActionPreference is Stop, so any unhandled error would otherwise end the
# process and close the window before anyone could read it.
trap {
    Write-Host ""
    Write-Host (T 'trap_title') -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    # .Line can be null (e.g. errors from .NET calls); calling .Trim() on it
    # inside the trap would throw a second error on top of the first.
    if ($_.InvocationInfo -and $_.InvocationInfo.Line) {
        Write-Host (T 'trap_line' @($_.InvocationInfo.ScriptLineNumber, $_.InvocationInfo.Line.Trim())) -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host (T 'trap_advice_1') -ForegroundColor White
    Write-Host (T 'trap_advice_2') -ForegroundColor White
    Read-Host "`n$(T 'press_close')"
    exit 1
}

Show-Banner
Write-Host (T 'header') -ForegroundColor White
Write-Host ""

# 1) Reachability — fail early, before any changes, if off-network.
Info (T 'checking_reach' @($PrinterIP, $PrinterPort))
$reachable = $false
try {
    $client = New-Object System.Net.Sockets.TcpClient
    $iar = $client.BeginConnect($PrinterIP, $PrinterPort, $null, $null)
    if ($iar.AsyncWaitHandle.WaitOne(3000, $false) -and $client.Connected) { $reachable = $true }
    $client.Close()
} catch { $reachable = $false }
if (-not $reachable) { Die (T 'not_reachable' @($PrinterIP)) }
Ok (T 'reachable')

# 2) Credentials — from params, else prompt.
if (-not $Username) {
    Write-Host ""
    Write-Host (T 'ask_login') -ForegroundColor White
    $Username = (Read-Host (T 'ask_initials')).Trim()
}
if (-not $Password) {
    $Password = (Read-Host (T 'ask_pin')).Trim()
}
if (-not $Username -or -not $Password) { Die (T 'need_both') }

# 3) Registry fix (undo PrintNightmare strict RPC auth).
Info (T 'reg_fix')
$regPath = "HKLM:\System\CurrentControlSet\Control\Print"
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "RpcAuthnLevelPrivacyEnabled" -Value 0 -Type DWord
Ok (T 'reg_done')

# 4) Store credentials in Windows Credential Manager.
# NOTE: this is NOT what authenticates the print job — a RAW 9100 port never
# consults Credential Manager (see the header). Kept only so any future
# IPP/SMB/web access to the device finds a login. The real auth is step 12.
Info (T 'storing_login')
# Deleting a credential that isn't there is normal on a first run. Redirecting a
# native command's stderr while $ErrorActionPreference is Stop can turn that into
# a terminating NativeCommandError in Windows PowerShell 5.1, so relax it here.
& {
    $ErrorActionPreference = "Continue"
    cmdkey /delete:$PrinterIP 2>&1 | Out-Null
}
$null = cmdkey /add:$PrinterIP /user:$Username /pass:$Password
if ($LASTEXITCODE -ne 0) { Die (T 'cmdkey_failed') }
Ok (T 'login_stored' @($PrinterIP, $Username))

# 5) Restart the Print Spooler so it picks up the changes.
Info (T 'spooler')
Restart-Service -Name Spooler -Force
Start-Sleep -Seconds 2
Ok (T 'spooler_done')

# 6) Find the driver — and only download it if it isn't already here.
# The name it registers under varies by package: the Olivetti build calls itself
# "Generic Universal PS v3.9.12", the Konica Minolta build "KONICA MINOLTA
# Universal PS". Same universal PostScript driver, both emit the KM auth PJL, so
# match any of them rather than hardcoding one and failing at Add-Printer.
function Resolve-KmDriver {
    $installed = @(Get-PrinterDriver -ErrorAction SilentlyContinue)
    foreach ($want in @($DriverName, "KONICA MINOLTA Universal PS", "Generic Universal PS")) {
        $hit = $installed | Where-Object { $_.Name -eq $want } | Select-Object -First 1
        if ($hit) { return $hit.Name }
    }
    $hit = $installed |
        Where-Object { $_.Name -like "*Universal PS*" -or $_.Name -like "*Universal*PostScript*" } |
        Select-Object -First 1
    if ($hit) { return $hit.Name }
    return $null
}

$work = $null      # only set if we actually download; cleanup below checks it
$resolvedDriver = Resolve-KmDriver
if ($resolvedDriver) {
    Ok (T 'driver_present' @($resolvedDriver))
} else {
    # Disk space, checked BEFORE the 52 MB download. Installing this driver
    # briefly needs several times its download size: the zip, our extraction,
    # pnputil's own temp copy of the package, and finally the expanded copy in
    # the driver store. On a full disk the failure surfaces late and cryptically
    # (pnputil exit 112 / ERROR_DISK_FULL, with the real reason buried in
    # setupapi.dev.log), so fail early and say the actual number.
    $needMB = 1024
    $sysDrive = ($env:SystemDrive).TrimEnd(':')
    function Get-FreeMB {
        try { return [math]::Round((Get-PSDrive -Name $sysDrive -ErrorAction Stop).Free / 1MB) }
        catch { return -1 }
    }

    # Reclaimable locations. Everything here is data Windows itself regenerates:
    # nothing the user created, nothing that uninstalls an update or a program.
    #
    # Deliberately NOT included:
    #   * cleanmgr automation - it is GUI-bound, /verylowdisk is unreliable, and
    #     it is widely reported to hang at 100% when run unattended.
    #   * "Windows Update Cleanup" / previous installations (Windows.old) - only
    #     reversible by reinstalling, far too destructive for a printer setup.
    #   * DISM /StartComponentCleanup - slow, and it removes the ability to roll
    #     back updates.
    #   * Downloads, Documents, or anything else the user owns.
    # The Recycle Bin IS the user's data, so it is asked for separately below.
    function Get-FolderMB($path) {
        if (-not (Test-Path $path)) { return 0 }
        try {
            $b = (Get-ChildItem $path -Recurse -Force -File -ErrorAction SilentlyContinue |
                  Measure-Object -Property Length -Sum).Sum
            return [math]::Round(($b) / 1MB)
        } catch { return 0 }
    }

    $freeMB = Get-FreeMB
    if ($freeMB -ge 0 -and $freeMB -lt $needMB) {
        Write-Host ""
        Warn (T 'disk_need' @($freeMB, $needMB, $env:SystemDrive))
        Info (T 'disk_scan')

        $targets = @(
            @{ Label = (T 'disk_t_wintemp'); Path = (Join-Path $env:SystemRoot "Temp");                 Svc = $null },
            @{ Label = (T 'disk_t_usertemp'); Path = $env:TEMP;                                          Svc = $null },
            @{ Label = (T 'disk_t_update');   Path = (Join-Path $env:SystemRoot "SoftwareDistribution\Download"); Svc = "wuauserv" },
            @{ Label = (T 'disk_t_do');       Path = (Join-Path $env:SystemRoot "ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"); Svc = "DoSvc" }
        )
        foreach ($t in $targets) { $t.MB = Get-FolderMB $t.Path }
        $found = @($targets | Where-Object { $_.MB -gt 0 })

        if ($found.Count -eq 0) {
            Warn (T 'disk_nothing')
        } else {
            # Sum by hand: in Windows PowerShell 5.1, Measure-Object -Property
            # can't see hashtable keys (PS 7 can), and this has to run on 5.1.
            $totalMB = 0
            foreach ($t in $found) {
                Write-Host ("    {0,6} MB  {1}" -f $t.MB, $t.Label)
                $totalMB += $t.MB
            }
            Write-Host (T 'disk_total' @($totalMB)) -ForegroundColor White
            $ans = (Read-Host (T 'disk_ask')).Trim().ToLower()
            if ($ans -eq "y" -or $ans -eq "j") {
                Info (T 'disk_cleaning')
                foreach ($t in $found) {
                    # Some caches are owned by a running service; stop it, clear,
                    # start it again. Files still locked are skipped, not forced.
                    if ($t.Svc) { try { Stop-Service $t.Svc -Force -ErrorAction SilentlyContinue } catch { } }
                    try {
                        Get-ChildItem $t.Path -Force -ErrorAction SilentlyContinue |
                            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    } catch { }
                    if ($t.Svc) { try { Start-Service $t.Svc -ErrorAction SilentlyContinue } catch { } }
                }
                $after = Get-FreeMB
                Ok (T 'disk_freed' @(($after - $freeMB), $after))
                $freeMB = $after
            } else {
                Warn (T 'disk_skipped')
            }
        }

        # The Recycle Bin is the user's own deleted files, so it is a separate,
        # explicit question with the consequence spelled out - never bundled in.
        if ($freeMB -lt $needMB) {
            $ans2 = (Read-Host (T 'disk_bin_ask')).Trim().ToLower()
            if ($ans2 -eq "y" -or $ans2 -eq "j") {
                try {
                    Clear-RecycleBin -DriveLetter $sysDrive -Force -ErrorAction SilentlyContinue
                    $after = Get-FreeMB
                    Ok (T 'disk_freed' @(($after - $freeMB), $after))
                    $freeMB = $after
                } catch { }
            }
        }

        if ($freeMB -lt $needMB) { Write-Host ""; Die (T 'disk_full' @($freeMB, $needMB, $env:SystemDrive)) }
    }

    Info (T 'driver_dl')
    # A short, local, machine-scoped path. The old %TEMP%\kmdriver_<32-hex-guid>
    # sat deep under a user profile; long paths and per-user temp are both
    # implicated in staging write faults, and this costs nothing to avoid.
    $work = Join-Path $env:SystemRoot "Temp\kmdrv"
    if (Test-Path $work) { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $zip = Join-Path $work "driver.zip"
    try {
        Invoke-WebRequest -Uri $DriverUrl -OutFile $zip -UseBasicParsing
    } catch {
        Die (T 'driver_dl_fail' @($DriverUrl))
    }
    Expand-Archive -Path $zip -DestinationPath $work -Force
    # Reclaim the 52 MB archive immediately - on a tight disk this is exactly
    # the margin the driver-store copy needs.
    Remove-Item $zip -Force -ErrorAction SilentlyContinue
    $inf = Join-Path $work $DriverInf
    if (-not (Test-Path $inf)) { Die (T 'driver_corrupt' @($DriverInf)) }
    Ok (T 'driver_dl_ok')

    # 7) Install the driver into the Windows driver store.
    Info (T 'driver_install')
    # Same stderr/EAP hazard as cmdkey above - pnputil writes to stderr on
    # several non-fatal paths, and with EAP=Stop that would terminate the run.
    # /subdirs is recommended for printer packages whose INF references files
    # that aren't flat next to it; harmless when they are.
    & {
        $ErrorActionPreference = "Continue"
        $script:pnp = & pnputil.exe /add-driver "$inf" /subdirs /install 2>&1
    }
    # pnputil returns 0 (added), 259 (no new driver / already present) or 3010
    # (added, reboot required) on success-ish paths.
    # 112 is ERROR_DISK_FULL and 29 ERROR_WRITE_FAULT - both mean the copy into
    # the driver store failed for space/write reasons, not anything about the
    # driver. Stop here with a straight answer instead of trying three more
    # install routes that must all fail the same way.
    if ($LASTEXITCODE -in 112, 29) {
        $freeNow = 0
        try { $freeNow = [math]::Round((Get-PSDrive -Name ($env:SystemDrive).TrimEnd(':')).Free / 1MB) } catch { }
        Write-Host ""
        Die (T 'disk_full_late' @($freeNow, $env:SystemDrive))
    }
    if ($LASTEXITCODE -notin 0, 259, 3010) { Warn (T 'pnputil_warn' @($LASTEXITCODE)) }

    # Staging the INF into the driver store is NOT enough. The print spooler
    # keeps its own separate list, and Get-PrinterDriver / Add-Printer only see
    # drivers registered with Add-PrinterDriver. On a machine that has never had
    # the Konica package installed, skipping this is why the driver "installed"
    # but Windows reported no Universal PS driver.
    #
    # The name to register is read from the INF rather than guessed: this INF
    # declares "Generic Universal PS" with no version suffix, while the vendor's
    # own installer registers "Generic Universal PS v3.9.12" and the Konica
    # build "KONICA MINOLTA Universal PS". Parse first, then fall back.
    if (-not (Resolve-KmDriver)) {
        Info (T 'driver_register')
        $infNames = @()
        try {
            $infNames = @(Get-Content $inf -ErrorAction Stop | ForEach-Object {
                if ($_ -match '^\s*"([^"]+)"\s*=') { $Matches[1] }
            } | Select-Object -Unique)
        } catch { }
        $candidates = @($infNames + @("Generic Universal PS", $DriverName, "KONICA MINOLTA Universal PS")) |
            Where-Object { $_ } | Select-Object -Unique

        # Prefer the copy pnputil staged into the driver store over our temp
        # extraction: Add-PrinterDriver resolves names against the store, and the
        # staged path is what it expects. Fall back to the temp INF.
        $infPaths = @()
        try {
            $staged = Get-ChildItem "$env:SystemRoot\System32\DriverStore\FileRepository" -Recurse -Filter $DriverInf `
                        -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($staged) { $infPaths += $staged.FullName }
        } catch { }
        $infPaths += $inf

        # Exit code 29 from pnputil is ERROR_WRITE_FAULT: staging into the driver
        # store failed outright. Say so plainly rather than letting the user
        # think the driver is present, because Add-PrinterDriver resolves names
        # against the store and will usually fail too.
        if (-not $staged) { Warn (T 'driver_not_staged') }

        :register foreach ($path in $infPaths) {
            foreach ($cand in $candidates) {
                try {
                    Add-PrinterDriver -Name $cand -InfPath $path -ErrorAction Stop
                    Ok (T 'driver_registered' @($cand))
                    break register
                } catch { }
            }
        }

        # Last resort: printui's own INF installer. It predates the PrintManagement
        # module and takes a different route into the spooler, so it sometimes
        # succeeds where Add-PrinterDriver won't. It reports failure only through
        # dialogs, hence /q plus a bounded wait, and we judge it by whether the
        # driver actually appears afterwards.
        if (-not (Resolve-KmDriver)) {
            Info (T 'driver_fallback')
            foreach ($cand in $candidates) {
                try {
                    $p = Start-Process rundll32.exe -PassThru -ArgumentList `
                        "printui.dll,PrintUIEntry", "/ia", "/q", "/m", "`"$cand`"", "/f", "`"$inf`""
                    if (-not $p.WaitForExit(90000)) { try { $p.Kill() } catch {} }
                } catch { }
                if (Resolve-KmDriver) { Ok (T 'driver_registered' @($cand)); break }
            }
        }
    }

    $resolvedDriver = Resolve-KmDriver
    if (-not $resolvedDriver) {
        Write-Host (T 'drivers_known') -ForegroundColor Yellow
        Get-PrinterDriver | Select-Object -ExpandProperty Name | ForEach-Object { Write-Host "    $_" }
        if ($candidates) {
            Write-Host (T 'driver_tried') -ForegroundColor Yellow
            $candidates | ForEach-Object { Write-Host "    $_" }
        }
        # pnputil's text output says far more than its numeric exit code.
        if ($pnp) {
            Write-Host (T 'pnputil_said') -ForegroundColor Yellow
            $pnp | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
        }
        # setupapi.dev.log is the authoritative record of why staging failed.
        # Pull the lines about THIS package so the next attempt starts from fact
        # instead of guesswork.
        try {
            $log = Join-Path $env:SystemRoot "inf\setupapi.dev.log"
            if (Test-Path $log) {
                $hits = @(Select-String -Path $log -Pattern "koawnaa" -SimpleMatch -ErrorAction SilentlyContinue |
                          Select-Object -Last 12)
                if ($hits) {
                    Write-Host (T 'setupapi_said') -ForegroundColor Yellow
                    $hits | ForEach-Object { Write-Host "    $($_.Line.Trim())" -ForegroundColor DarkGray }
                }
            }
        } catch { }
        Die (T 'driver_none')
    }
    Ok (T 'driver_ok' @($resolvedDriver))
}

$portName = "IP_$PrinterIP"

# 8) Is a correct queue already here? Removing and re-adding wipes BOTH the
# restored driver config and the per-user credentials from step 12, so a user
# re-running this would silently lose their login. Keep a healthy queue.
$existing = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
$queueHealthy = $false
if ($existing -and $existing.DriverName -eq $resolvedDriver -and $existing.PortName -eq $portName) {
    $queueHealthy = $true
    Ok (T 'queue_ok')
} elseif ($existing) {
    Remove-Printer -Name $PrinterName -Confirm:$false
    Ok (T 'removed_old')
}

# 8b) Offer to clear out OTHER queues pointing at the same printer. Earlier
# attempts (or Olivetti's own installer) leave behind e.g. "KONICA MINOLTA
# C250i" on the same IP: unauthenticated duplicates that show up in every Ctrl+P
# list, get picked by mistake, and have their jobs discarded.
$ourPorts = @(Get-PrinterPort -ErrorAction SilentlyContinue |
    Where-Object { $_.PrinterHostAddress -eq $PrinterIP } |
    Select-Object -ExpandProperty Name)
$dupes = @(Get-Printer -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -ne $PrinterName -and $ourPorts -contains $_.PortName })
# Removed without asking: these point at the SAME device on the same IP, and
# without a login the printer discards everything they send. There is nothing in
# them worth keeping, and a y/n prompt here just stalls an otherwise unattended
# run. What was removed is still printed, so it's never a silent change.
if ($dupes.Count -gt 0) {
    Write-Host ""
    Info (T 'dupes_found' @($dupes.Count, $PrinterIP))
    Write-Host (T 'dupes_explain_1')
    foreach ($d in $dupes) {
        try { Remove-Printer -Name $d.Name -Confirm:$false; Ok (T 'dupes_removed' @($d.Name)) }
        catch { Warn (T 'dupes_fail' @($d.Name, $_.Exception.Message)) }
    }
}

if (-not $queueHealthy) {
    # 9) Create the TCP/IP (RAW 9100) port.
    Info (T 'port_making')
    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    if ($existingPort) { Remove-PrinterPort -Name $portName -Confirm:$false -ErrorAction SilentlyContinue }
    # The remove can fail silently if another queue still holds the port. Re-check
    # rather than blindly adding: Add-PrinterPort would throw "already exists" and
    # abort the run with the old printer already deleted.
    if (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue) {
        Ok (T 'port_reuse' @($portName))
    } else {
        Add-PrinterPort -Name $portName -PrinterHostAddress $PrinterIP
        Ok (T 'port_ok' @($portName))
    }

    # 10) Add the printer queue.
    Info (T 'adding')
    try {
        Add-Printer -Name $PrinterName -DriverName $resolvedDriver -PortName $portName
    } catch {
        Write-Host (T 'avail_drivers') -ForegroundColor Yellow
        Get-PrinterDriver | Where-Object { $_.Name -like "*PS*" -or $_.Name -like "*Generic*" } |
            Select-Object -ExpandProperty Name | ForEach-Object { Write-Host "    $_" }
        Die (T 'add_failed' @($resolvedDriver))
    }
    Ok (T 'added' @($PrinterName))
}

# 11) Enable bidirectional + SNMP (MFP auth feature detection).
try {
    Set-Printer -Name $PrinterName -EnableBidirectional $true
    Set-PrinterPort -Name $portName -SNMP 1 -SNMPCommunity "public" -ErrorAction SilentlyContinue
    Ok (T 'bidi_ok')
} catch {
    Warn (T 'bidi_fail' @($_.Exception.Message))
}

# 11a) Make it the default, so the test print on the web page (and any Ctrl+P)
# doesn't quietly go to "Microsoft Print to PDF" instead. Matches install.sh.
try {
    (New-Object -ComObject WScript.Network).SetDefaultPrinter($PrinterName)
    Ok (T 'default_ok')
} catch { }

# 11b) Replay the golden driver configuration, if we've published one.
# The automated half of the auth setup: "Device Option -> User Authentication =
# On" lives in the queue's PrinterDriverData, is identical on every machine and
# carries no personal data, so it's captured once with export_golden_config.ps1
# and restored here. Flags: d = PrinterDriverData, g = global DevMode,
# r = resolve name conflicts. Deliberately NOT `u` (per-user DevMode) — that's
# where credentials live, and nobody's PIN should ride in a published file.
$configRestored = $false
Info (T 'cfg_looking')
$cfg = Join-Path $env:TEMP "printer-config.dat"
try {
    Invoke-WebRequest -Uri $ConfigUrl -OutFile $cfg -UseBasicParsing
} catch {
    $cfg = $null
}
# A host that serves an HTML 404 page with status 200 would hand us a "config"
# that printui can only choke on; sniff for that rather than reporting a
# confusing restore failure.
if ($cfg -and (Test-Path $cfg)) {
    $head = ""
    # A pipeline can't be an argument directly; read the bytes, then slice.
    try {
        $bytes = [IO.File]::ReadAllBytes($cfg)
        $take  = [Math]::Min(64, $bytes.Length)
        if ($take -gt 0) { $head = [Text.Encoding]::ASCII.GetString($bytes, 0, $take) }
    } catch {}
    if ((Get-Item $cfg).Length -eq 0 -or $head -match '(?i)^\s*(<!doctype|<html)') {
        Remove-Item $cfg -Force -ErrorAction SilentlyContinue
        $cfg = $null
    }
}
if ($cfg) {
    # rundll32 exits 0 whether or not PrintUIEntry succeeded, so its exit code
    # proves nothing. Believing a failed restore worked is the dangerous
    # direction — we'd skip Part A and the user would print anonymous jobs. So
    # compare the queue's PrinterDriverData before and after.
    $ddPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$PrinterName\PrinterDriverData"
    # Exclude the six PowerShell provider members by exact name. A -notlike "PS*"
    # filter would also drop real driver values beginning with "PS" — likely on a
    # PostScript driver.
    $providerMembers = @("PSPath","PSParentPath","PSChildName","PSDrive","PSProvider","PSStatus")
    function Get-DriverDataFingerprint($path) {
        try {
            $item = Get-ItemProperty -Path $path -ErrorAction Stop
            return ($item.PSObject.Properties |
                Where-Object { $providerMembers -notcontains $_.Name } |
                Sort-Object Name |
                ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "|"
        } catch { return "" }
    }
    $before = Get-DriverDataFingerprint $ddPath
    try {
        # No /q: it's a real printui switch but undocumented, and an unrecognised
        # switch makes printui show its usage dialog — the modal hang we're
        # avoiding. Bound the wait and kill it if a dialog does appear.
        $proc = Start-Process rundll32.exe -PassThru -ArgumentList `
            "printui.dll,PrintUIEntry", "/Sr", "/n", "`"$PrinterName`"", "/a", "`"$cfg`"", "d", "g", "r"
        if (-not $proc.WaitForExit(60000)) {
            try { $proc.Kill() } catch {}
            Warn (T 'cfg_timeout')
        }
    } catch {
        Warn (T 'cfg_runfail')
    }
    Start-Sleep -Milliseconds 750     # let the spooler flush to the registry
    if ((Get-DriverDataFingerprint $ddPath) -ne $before) {
        $configRestored = $true
        Ok (T 'cfg_ok')
    } else {
        Warn (T 'cfg_notake')
    }
    Remove-Item $cfg -Force -ErrorAction SilentlyContinue
} else {
    Warn (T 'cfg_none')
}

# 12) THE STEP THAT MAKES AUTH WORK — put the credentials into the driver.
# The PIN can't be scripted (undocumented DEVMODE blob), so we open the windows.
#
# First, though: those credentials are stored PER USER (in HKCU) and we are
# elevated. If UAC was answered with a different account than the one logged in,
# the dialog would save the PIN into the admin's profile and the real user's
# jobs would still go out unauthenticated.
$interactive = $null
try { $interactive = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName } catch {}
# Compare leaf names only: Win32_ComputerSystem reports "AzureAD\First Last" or
# "MicrosoftAccount\someone@example.com" where $env:USERDOMAIN is the machine
# name, so a domain-qualified comparison false-positives on ordinary consumer
# and Entra-joined PCs and would send everyone down the manual path.
$leaf = { param($n) if ($n -and $n.Contains("\")) { $n.Split("\")[-1] } else { $n } }
$interactiveLeaf = & $leaf $interactive
if ($interactiveLeaf -and $interactiveLeaf -ne $env:USERNAME) {
    Write-Host ""
    Warn (T 'wrong_user_1' @("$env:USERDOMAIN\$env:USERNAME", $interactive))
    Warn (T 'wrong_user_2')
    Warn (T 'wrong_user_3' @($interactive))
    Write-Host ""
    Write-Host (T 'wrong_user_4' @($interactive)) -ForegroundColor White
    Write-Host (T 'wrong_user_5' @($PrinterName)) -ForegroundColor Cyan
    if (-not $configRestored) {
        Write-Host (T 'wrong_user_8') -ForegroundColor White
        Write-Host (T 'wrong_user_9') -ForegroundColor White
    }
    Write-Host (T 'wrong_user_6') -ForegroundColor White
    Write-Host (T 'wrong_user_7') -ForegroundColor White
    Write-Host ""
    Read-Host (T 'press_close')
    exit 0
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host (T 'last_title') -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host (T 'last_why_1')
Write-Host (T 'last_why_2')
Write-Host ""

# PART A — Configure tab. This lives in Printer PROPERTIES (/p), NOT in Printing
# Preferences (/e); opening the wrong one leaves the user hunting for a tab that
# isn't there. Only needed when the golden config didn't apply.
if (-not $configRestored) {
    Write-Host (T 'partA_title') -ForegroundColor White
    Write-Host (T 'partA_note')
    Write-Host ""
    Write-Host (T 'partA_open')
    Write-Host (T 'partA_1')
    Write-Host (T 'partA_2')
    Write-Host (T 'partA_3')
    Write-Host (T 'partA_4')
    Write-Host ""
    Read-Host (T 'press_open_A')
    try {
        Start-Process rundll32.exe -ArgumentList "printui.dll,PrintUIEntry", "/p", "/n", "`"$PrinterName`"" -Wait
    } catch {
        Warn (T 'dialog_fail' @($PrinterName))
    }
    Write-Host ""
}

# PART B — the credentials themselves, in Printing PREFERENCES (/e).
if ($configRestored) {
    Write-Host (T 'partB_only') -ForegroundColor White
} else {
    Write-Host (T 'partB_title') -ForegroundColor White
}
Write-Host ""
Write-Host (T 'partB_open')
Write-Host (T 'partB_1')
Write-Host (T 'partB_2')
Write-Host (T 'partB_3')
Write-Host (T 'partB_user' @($Username)) -ForegroundColor Cyan
Write-Host (T 'partB_pin') -ForegroundColor Cyan
Write-Host (T 'partB_4')
Write-Host (T 'partB_5')
Write-Host ""
Write-Host (T 'verify_note_1') -ForegroundColor DarkGray
Write-Host (T 'verify_note_2') -ForegroundColor DarkGray
Write-Host ""
Read-Host (T 'press_open_B')
try {
    Start-Process rundll32.exe -ArgumentList "printui.dll,PrintUIEntry", "/e", "/n", "`"$PrinterName`"" -Wait
    Ok (T 'dialog_closed')
} catch {
    Warn (T 'dialog_fail' @($PrinterName))
}
Write-Host ""
Read-Host (T 'press_done')

# 13) Confirmation print.
if (-not $NoTest) {
    Info (T 'test_sending')
    try {
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        @(
            (T 'test_line_1' @($Username)),
            $stamp,
            "",
            (T 'test_line_2' @($PrinterName))
        ) | Out-Printer -Name $PrinterName
        Ok (T 'test_sent')
    } catch {
        Warn (T 'test_fail')
    }
}

# Cleanup. $work is only set when we downloaded; Remove-Item with a null -Path
# is a parameter-binding error that -ErrorAction cannot suppress, so it would
# terminate an otherwise successful run.
if ($work) { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }

# The script cannot verify that the PIN was actually entered, so it must not
# claim success. Tell the user what success looks like instead.
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host (T 'done_title') -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host (T 'done_1') -ForegroundColor White
Write-Host ""
Write-Host (T 'done_2') -ForegroundColor White
Write-Host "  $PrinterName" -ForegroundColor Cyan
Write-Host ""
Write-Host (T 'done_3') -ForegroundColor Yellow
Write-Host (T 'done_4') -ForegroundColor Yellow
Write-Host ""
Read-Host (T 'press_close')
