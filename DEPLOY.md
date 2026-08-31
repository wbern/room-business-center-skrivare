# Deploying the printer setup site

The `docs/` folder is a complete, static GitHub Pages site. Once it's live, any
Mac **or Windows** user adds the office printer by visiting the page and pasting
one line into Terminal (Mac) or PowerShell (Windows). No app, no Apple Developer
account, no Python — the Mac backend is Perl (every Mac has it) and the Windows
installer is plain PowerShell (every PC has it).

The page auto-detects the OS and shows the right flow; a `macOS | Windows` toggle
in the header lets users switch manually.

## What's in here

All site files live under `docs/`.

| File | Purpose |
|------|---------|
| `index.html` | The page users visit — a 5-step wizard (welcome → login → open Terminal/PowerShell → paste command → done), **OS-aware (Mac/Windows)**, EN/SV, light navy/teal theme. Generates the command each user copies (including `$env:PRINTER_LANG`). |
| **macOS** | |
| `install.sh` | The `curl … \| bash` installer. Dependency-free. |
| `km9100auth` | The Perl CUPS backend the installer deploys (injects PJL auth, strips the `KMCOETYPE` line that breaks GUI prints). |
| `km-c250i-driver.pkg` | The Konica Minolta C250i Mac driver (47 MB, signed by KM). Extracted from `IT6PSMACOS_536AMU.dmg`; installs the `KONICAMINOLTAC250i` PPD. |
| **Windows** | |
| `install.ps1` | The installer proper — reached via `boot.ps1`, never fetched directly (see below). Bilingual EN/SV. Self-elevates (UAC), prompts for initials + PIN, downloads the driver, applies the `RpcAuthnLevelPrivacyEnabled=0` registry fix + Olivetti PS driver, replays `printer-config.dat`, then opens the driver's auth dialog for the PIN. Original reference: `printer-windows-setup/web_install.ps1` (kept on disk only, gitignored). |
| `printer-config.dat` | The golden driver configuration (`PrinterDriverData` + global DevMode), captured from a verified working PC with `export_golden_config.ps1`. Carries the `Device Option → User Authentication = On` flag so users never touch the Configure tab. Contains no credentials. |
| `export_golden_config.ps1` | Admin-only working-PC baseline collector. Produces `printer-config.dat` plus a shareable ZIP with safe driver, port and configuration fingerprints. Linked from the discreet Admin control. |
| `capture_pjl.ps1` | Admin-only diagnostic: dumps the PJL header the driver really sends. Not linked from the page. |
| `boot.ps1` | **What the Windows one-liner actually runs.** Pure ASCII by necessity: Windows PowerShell 5.1 decodes a charset-less HTTP response as ISO-8859-1, and GitHub Pages serves `.ps1` as `application/octet-stream` with no charset — so fetching `install.ps1` directly with `irm` would mojibake every `å ä ö`. `boot.ps1` survives that by being ASCII, then re-downloads `install.ps1` as raw bytes and decodes it as UTF-8. **Never put a non-ASCII character in this file**, and keep `install.ps1` UTF-8 **without** BOM. |
| `printer-driver-win-x64.zip` | Olivetti Universal PS v3.9.12 driver, x64 only (52 MB). Zipped from `printer-windows-setup/GEUPDPSWin_3912040MU/driver/win_x64`; INF `KOAWNAA_.inf` at the zip root. Fetched at runtime by `install.ps1`. |

## One-time setup

1. **Drivers — already bundled.** `docs/km-c250i-driver.pkg` (Mac, 47 MB) and
   `docs/printer-driver-win-x64.zip` (Windows, 52 MB) are both in place, so no
   action needed. They're large binaries committed to the repo; if you'd rather
   not commit them, move them to GitHub *Release* assets and set
   `PRINTER_DRIVER_URL` (env var, read by both `install.sh` and `install.ps1`)
   to those URLs.

2. **Push to GitHub and enable Pages.**
   ```sh
   git init && git add docs && git commit -m "printer setup site"
   git branch -M main && git remote add origin <your-repo> && git push -u origin main
   ```
   Repo → Settings → Pages → Source: *Deploy from a branch* → `main` / `/docs`.

3. **Serve it.** This repo deploys as a project site under the `pages.bernting.se`
   GitHub user domain, so it's reachable at
   `https://pages.bernting.se/room-business-center-skrivare` with no per-project
   `CNAME` file. If you fork it elsewhere, the served URL becomes
   `https://<youruser>.github.io/<repo>`; update it in the three places that
   hardcode the origin: `MAC_INSTALL_URL` / `WIN_INSTALL_URL` in `index.html`,
   `SITE` in `install.sh`, and `$Site` in `install.ps1` (each is also overridable
   at runtime via the `PRINTER_SITE` env var).

4. **Share the link.** Send people to `https://pages.bernting.se/room-business-center-skrivare`. That's it.

## Updating later

- Change the Mac backend? Edit `docs/km9100auth` directly, commit, push. (The
  local `install_printer.sh` keeps its own copy at `bin/km9100auth`; both are
  gitignored, on-disk only.)
- Change the Windows installer? Edit `docs/install.ps1` directly, commit, push.
  The local `printer-windows-setup/web_install.ps1` is the original source for
  reference.
- Installers fetch the latest `install.sh` / `install.ps1` (and the Mac backend)
  at run time, so users get fixes automatically the next time they run it.

## Notes

- **Mac:** the generated command contains the user's PIN when one is entered
  (convenient, but it lands in shell history). Leaving the PIN field blank omits
  `-p` from the command, and `install.sh` prompts for the PIN in Terminal
  instead. `install.sh` also removes the no-auth `_192_168_9_15`-style duplicate
  queue macOS auto-creates and sets `ROOM Business Center (Olivetti MF224)` as the
  default. Those are the two things that broke printing originally.
- **Windows auth — how it actually works.** The queue is a RAW TCP/9100 port: a
  bare socket with no auth challenge, so Windows Credential Manager is never
  consulted for it and the `cmdkey` step authenticates nothing (it's kept only
  for future IPP/SMB/web access). Credentials travel *inside* the job, as the
  PJL header `@PJL SET KMUSERNAME` / `KMUSERKEY2` — the same block the Mac's
  `km9100auth` backend injects. On Windows the KM driver writes that header
  itself, but only when **both** of these are set:

  1. `Configure → Device Option → User Authentication = On`. Without it the
     driver emits a job with no auth info at all, and the C250i discards such
     jobs by default (`Restrict`). Because we add the queue with `Add-Printer`
     rather than Olivetti's installer, the driver never interrogates the device,
     so this defaults to `None` — this was the long-standing missing piece.
  2. `Basic → Authentication/Account Track → Recipient User` + initials/PIN.

  (1) is machine-independent and holds no secrets, so it's automated: capture it
  once, publish it, and `install.ps1` replays it. (2) is per-user and lives in an
  undocumented private DEVMODE blob — KM ships no supported way to script it
  (DPU has no silent switches, and the stored hash is derived from the queue
  name), so the installer opens the dialog and the user types the PIN once.

  **Publishing the golden config (do this once):**
  ```powershell
  # on a Windows machine where printing already works, as Administrator:
  irm https://pages.bernting.se/room-business-center-skrivare/export_golden_config.ps1 | iex
  ```
  It writes `printer-config.dat` and a dated `olivetti-working-baseline_*.zip`
  to that machine's Desktop. Send the ZIP for troubleshooting; copy the `.dat`
  here when you want to publish the verified configuration:
  ```sh
  git add docs/printer-config.dat && git commit -m "Add golden driver config" && git push
  ```
  The export uses `/Ss … d g` and never `u` (per-user DevMode), so it does not
  read or publish the user's initials or PIN.
  Until it exists, `install.ps1` detects the 404 and walks users through the
  Configure tab by hand instead.

- **Script language (EN/SV).** `install.ps1` carries both languages inline as a
  `$MSG` hashtable-of-hashtables with a `T 'key'` helper (`Import-LocalizedData`
  is unusable — it resolves `.psd1` against `$PSScriptRoot`, which doesn't exist
  under `irm | iex`). The page's flag toggle bakes `$env:PRINTER_LANG` into the
  copied command; the script falls back to `Get-Culture`, then English. The
  choice is re-set explicitly inside the elevation command because environment
  variables don't survive ShellExecute's `runas` verb. When adding a string, add
  it to **both** language blocks with matching `{0}` placeholders — a missing key
  silently falls back to English.

- **Windows driver baseline.** The verified working PC uses `Generic Universal
  PS` from `KOAWNAA_.inf`, on a RAW 9100 port with SNMP disabled. The installer
  deliberately installs and uses that exact driver instead of silently reusing
  an existing Konica Universal PS variant; it leaves old driver packages alone
  so unrelated printer queues are not disrupted.

- **Windows diagnostics:** `docs/capture_pjl.ps1` clones the queue onto a
  loopback port and dumps the PJL header the driver actually emits, checking it
  against the known-good Mac block. Use it if jobs still vanish:
  ```powershell
  irm https://pages.bernting.se/room-business-center-skrivare/capture_pjl.ps1 | iex
  ```
- The Windows installer can't be run from this Mac (no Windows). Everything in it
  passes a PowerShell AST parse, but the `printui.dll /Sr` restore, the
  golden-config export and the auth dialog have **not** been exercised on real
  hardware. Verify on one PC before wide rollout — that same run is what produces
  `printer-config.dat` for everyone else.
