# Ricoh Printer SNMP Monitoring Tool

![PowerShell](https://img.shields.io/badge/PowerShell-v5.1+-blue.svg)
![Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)
![Deployment](https://img.shields.io/badge/Schedule-Monthly%201st-brightgreen)
![Device](https://img.shields.io/badge/Device-Ricoh%20Printer-red.svg)
![Protocol](https://img.shields.io/badge/Protocol-SNMP-blue.svg)

A PowerShell script (compilable to EXE) that monitors Ricoh IM-series printers via SNMP and produces a monthly dark-themed HTML report with toner levels, grand totals and a per-section page breakdown (Copier and Printer: Full Color / B&W / Single Color / Two-color). The report can optionally be emailed.

## Features

- Auto-discovers Ricoh printers on the local subnet — no hand-written config required
- Startup identity-verification distinguishes "no SNMP response" from "wrong device at this IP" from "OK"
- Counter collection via the standard RicohPrivateMIB, MAC and device-IP via raw UDP SNMP (handles bytes that OlePrn mangles)
- HTML report is **always** written to disk; email is an additional channel controlled by a `SendEmail` flag
- Clean dark-mode 2-column report with colored toner bars and spec-accurate sentinel handling (`-100 = Almost Empty`, `-2 = Unknown`, `0 = Empty`, `-3 = Some Left`)
- Self-contained: no external PowerShell modules, no external executables, pure .NET + COM
- Diagnostic switches for SNMP (`-TestSnmp`) and SMTP (`-TestSmtp`) so you can isolate problems
- Compilable to a standalone EXE so Task Scheduler doesn't hit execution-policy prompts

## System Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+ (Windows PowerShell or PowerShell 7+)
- Network reachability to the printers on SNMP port 161 (UDP)
- (Optional) SMTP relay reachable for emailed reports

## Quick Start

The fastest happy path:

```powershell
# 1. Copy Ricoh-Monitor.ps1 into a folder, say C:\Ricoh-Monitor
cd C:\Ricoh-Monitor

# 2. Allow the script to run for this session (one-shot, no permanent policy change)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 3. Run it - first run will auto-discover printers, write the config, and produce an HTML report
.\Ricoh-Monitor.ps1
```

On first run, with no `Ricoh-Monitor.json` present, the script:

1. Scans the local NIC's subnet for Ricoh printers (parallel SNMP sweep, ~6–15 s per `/24`).
2. If nothing is found there, prompts for additional subnets in CIDR notation (e.g. `192.168.1.0/24`).
3. Writes everything it finds into `Ricoh-Monitor.json` with `SendEmail: false` and placeholder `Smtp` values.
4. Immediately continues with a verify → collect → report pass against those printers.
5. Writes `report_yyyy-MM-dd_HHmm.html` next to the script. Open it in a browser.

Sample first-run console output:

```text
Starting printer monitoring...
=== First-run setup ===
No config file found. Discovering Ricoh printers on the network...

Local NIC subnet: 192.168.2.0/24
Probing 192.168.2.0/24 (254 addresses)...
Done in 6.2s. Found: 0 Ricoh, 3 other SNMP device(s)

No Ricoh printers found.
Enter another subnet to scan (CIDR e.g. 192.168.1.0/24), or blank to skip: 192.168.1.0/24
Probing 192.168.1.0/24 (254 addresses)...
Done in 7.1s. Found: 4 Ricoh, 0 other SNMP device(s)

Discovered Ricoh printers:
  [OK] 192.168.1.146   - RICOH IM C2010
  [OK] 192.168.1.197   - RICOH IM C2000
  [OK] 192.168.1.217   - RICOH IM C2000
  [OK] 192.168.1.230   - RICOH IM C400

Wrote 4 printer(s) to Ricoh-Monitor.json (SendEmail = false).

=== Verification phase ===
[OK]           192.168.1.146  - RICOH IM C2010
[OK]           192.168.1.197  - RICOH IM C2000
[OK]           192.168.1.217  - RICOH IM C2000
[OK]           192.168.1.230  - RICOH IM C400

=== Counter collection ===
Collecting from 192.168.1.146... OK
Collecting from 192.168.1.197... OK
Collecting from 192.168.1.217... OK
Collecting from 192.168.1.230... OK

Report written to: report_2026-04-20_0930.html
Email sending disabled in config; skipping email.
Done.
```

If the report looks right, move on to enabling email (or just leave it disabled and read the HTML files each month — that's a perfectly valid workflow).

## Enabling Email

1. Open `Ricoh-Monitor.json` and fill in the `Smtp` block with your real server/credentials. Example for Microsoft 365:

   ```json
   "Smtp": {
       "Server":   "smtp.office365.com",
       "Port":     587,
       "Username": "scanner@yourdomain.com",
       "Password": "<mailbox-or-app-password>",
       "From":     "scanner@yourdomain.com",
       "FromName": "RICOH Monitor",
       "To":       ["reports@yourdomain.com", "boss@yourdomain.com"]
   }
   ```

   `Port 587` with STARTTLS is the standard. For Microsoft 365 mailboxes with MFA enabled you either need an app password **or** ask IT to enable Authenticated SMTP on this specific mailbox (admin centre → Users → Mail → Manage email apps → Authenticated SMTP).

2. Test the SMTP path in isolation — does **not** require flipping `SendEmail` first:

   ```powershell
   .\Ricoh-Monitor.ps1 -TestSmtp
   ```

   Sends a tiny test email using the `Smtp` block and prints the result. If this fails, fix credentials/firewall before moving on. Common failure: `5.7.57 SmtpClientAuthentication is disabled for the Tenant` → Authenticated SMTP isn't enabled on the mailbox.

3. Once `-TestSmtp` succeeds, set `"SendEmail": true` in `Ricoh-Monitor.json` and run normally — you should receive the full report by email.

## Scheduling the Monthly Run

### Optional: compile to EXE

Task Scheduler running a `.ps1` can hit execution-policy prompts depending on machine settings. Compiling to an EXE sidesteps that entirely.

```powershell
Install-Module -Name PS2EXE -Scope CurrentUser
Invoke-PS2EXE -InputFile ".\Ricoh-Monitor.ps1" -OutputFile ".\Ricoh-Monitor.exe"
```

Both the `.ps1` and the `.exe` read the same `Ricoh-Monitor.json`.

### Register the scheduled task

Example: first of every month at 09:00. Adjust the time/day to suit.

```powershell
$Action   = New-ScheduledTaskAction -Execute "C:\Ricoh-Monitor\Ricoh-Monitor.exe"
$Trigger  = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "9:00AM"
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd

Register-ScheduledTask `
    -TaskName "Monthly Ricoh Report" `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -RunLevel Highest `
    -Description "Executes monthly Ricoh printer monitoring on the 1st at 9:00 AM"
```

If you skipped the EXE step, use `powershell.exe` with `-ExecutionPolicy Bypass -File <path-to-script>` as the action instead.

## Execution Policy Notes

On a default Windows install, running an unsigned `.ps1` fails with "running scripts is disabled on this system." Three ways to handle it:

- **Per session** (recommended for testing): `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` — resets when you close the shell, no persistent change.
- **Per user** (durable): `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` — allows local unsigned scripts forever, still blocks internet-downloaded ones.
- **Per invocation**: `powershell -ExecutionPolicy Bypass -File C:\Ricoh-Monitor\Ricoh-Monitor.ps1` — useful for Task Scheduler if you didn't compile to EXE.

If the script file was synced from Google Drive / OneDrive / email, Windows also tags it with a Mark-of-the-Web zone identifier that triggers the same error even when the policy is `RemoteSigned`. Strip it with `Unblock-File .\Ricoh-Monitor.ps1`, or move the file to a plain local folder like `C:\Ricoh-Monitor`.

## Configuration File

`Ricoh-Monitor.json` is auto-generated on first run — you usually don't write it by hand. For reference, here's the full shape:

```json
{
    "SendEmail": false,
    "Smtp": {
        "Server":   "smtp.example.com",
        "Port":     587,
        "Username": "user@example.com",
        "Password": "your-mailbox-or-app-password",
        "From":     "user@example.com",
        "FromName": "RICOH Monitor",
        "To":       ["recipient1@example.com", "recipient2@example.com"]
    },
    "Printers": [
        { "IP": "192.168.1.146", "SnmpCommunity": "public" },
        { "IP": "192.168.1.230", "SnmpCommunity": "public" }
    ]
}
```

- `SendEmail` — when `false` the script still produces the HTML report on disk but does not email it. Default for a fresh config. Switch to `true` after `-TestSmtp` confirms SMTP works.
- `Smtp` — server, credentials and recipients used when `SendEmail` is `true`. `To` is an array; list as many recipients as you like. `FromName` (optional) sets the display name shown in the recipient's inbox — e.g. `RICOH Monitor <scanner@yourdomain.com>`. If omitted, only the bare address appears.
- `Printers` — list of `{IP, SnmpCommunity}` entries. Auto-populated by discovery; can be edited manually. On every subsequent run the script also re-sweeps every `/24` these IPs cover, so newly-installed printers get added automatically.

> **Note:** the config contains credentials. It's listed in `.gitignore` and must not be committed.

## Command-line Parameters

- *(no parameter)* — Normal run: (optional first-run discovery) → known-subnet rescan → verify → collect → report. HTML is always written to disk; email is sent only if `SendEmail` is `true`.
- `-Discover` (switch) — Re-run discovery interactively: scans the local NIC subnet, then prompts for additional CIDRs, shows hits and asks y/N before merging into the config. Use this when you want to rescan deliberately, independent of the automatic every-run sweep.
- `-TestSnmp <IP>` (string) — Raw-SNMP diagnostic for one IP. Prints the encoded request bytes, the raw response bytes and the parsed varbind. Use when a MAC or device IP comes back `Unavailable` or a printer's data looks wrong.
- `-TestSmtp` (switch) — Sends a small test email via the current `Smtp` config. Does not require `SendEmail` to be `true`. Use to validate credentials and server reachability in isolation.

Parameters are case-insensitive and accept unambiguous prefixes (e.g. `-disc`).

## Data Collected

### Identity

- Manufacturer — `.1.3.6.1.4.1.367.3.2.1.1.1.7.0`
- Model — `.1.3.6.1.4.1.367.3.2.1.1.1.1.0`
- Serial Number — `.1.3.6.1.4.1.367.3.2.1.2.1.4.0`
- MAC Address — `.1.3.6.1.4.1.367.3.2.1.7.2.1.7.0`
- IP Address (device-reported) — `.1.3.6.1.4.1.367.3.2.1.7.2.1.3.0`

### Page counters

All from `.1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.X`:

- Total Pages — `.9.1`
- Total Pages Printer — `.9.8`
- Total Pages Copier — `.9.2`
- Copier breakdown — Full Color `.9.5`, B&W `.9.3`, Single Color `.9.7`, Two-color `.9.4`
- Printer breakdown — Full Color `.9.11`, B&W `.9.9`, Single Color `.9.41`, Two-color `.9.42`

### Toner levels

CMYK %, under `.1.3.6.1.4.1.367.3.2.1.2.24.1.1.5` (index 1 = K, 2 = C, 3 = M, 4 = Y). The script decodes the spec sentinels: `-100 = Almost Empty`, `0 = Empty`, `-2 = Unknown`, `-3 = Some Left`.

### Status

- Error State — `.1.3.6.1.4.1.367.3.2.1.2.2.13.0` (`0 = OK`, `2 = ADF Jam`, `3 = Hardware Error`, `4 = Service Call`)

## Startup Verification Output

Every run starts by identifying each configured IP via two SNMP gets (Manufacturer + Model):

```text
[OK]           192.168.1.146  - RICOH IM C2010
[WRONG DEVICE] 192.168.1.50   - Manufacturer: HP, Model: LaserJet Pro M404
[NO SNMP]      192.168.1.99   - timeout
```

Counter collection runs only on `[OK]` printers. `[WRONG DEVICE]` and `[NO SNMP]` entries still appear in the report with their status, so issues aren't silently dropped.

## Troubleshooting

### Printer not responding

- Verify SNMP is enabled on the printer (Device Settings → Network → SNMP)
- Check that the `SnmpCommunity` in the config matches the printer's community string
- Confirm network reachability: `Test-NetConnection <PrinterIP> -Port 161`
- Run `.\Ricoh-Monitor.ps1 -TestSnmp <PrinterIP>` to see exactly where the SNMP pipeline fails

### MAC or device IP shows "Unavailable"

These two fields use SNMP types (`OctetString` with high bytes, `IpAddress`) that the standard Windows COM SNMP wrapper mangles. The script falls back to raw UDP SNMP for these specifically, which usually works. If either still shows `Unavailable` even after a successful probe, run `-TestSnmp <IP>` to see whether the UDP response came back at all (firewall blocking inbound UDP 161 is a common cause).

### Email failures

- Run `.\Ricoh-Monitor.ps1 -TestSmtp` to exercise just the SMTP path
- Verify the `Smtp` credentials in `Ricoh-Monitor.json`
- **Microsoft 365**: the mailbox must have Authenticated SMTP enabled (admin centre → Users → Mail → Manage email apps). If tenant-wide SMTP AUTH is disabled, mailbox-level enable overrides it
- The script forces TLS 1.2 before STARTTLS — older servers requiring TLS 1.0 will not work without code changes
- Check Windows Event Logs for additional SMTP errors

### Script won't run — "running scripts is disabled"

See the [Execution Policy Notes](#execution-policy-notes) section. If the file is on Google Drive / OneDrive, also run `Unblock-File .\Ricoh-Monitor.ps1` to remove the Mark-of-the-Web zone identifier.

### First-run discovery finds nothing

- Confirm your PC and the printers are on the same VLAN, or supply the printer subnet at the "Enter another subnet to scan" prompt
- Windows Firewall may block inbound UDP 161 responses; allow them for the running user / PowerShell / EXE
- If you gave up with an empty config, either delete `Ricoh-Monitor.json` to re-trigger first-run discovery, or run `.\Ricoh-Monitor.ps1 -Discover` to rescan manually

## Supported Models

Confirmed on:

- Ricoh IM C2010
- Ricoh IM C2000
- Ricoh IM C400

Other Ricoh IM C models are very likely to work — the OID set used is the standard `RicohPrivateMIB` table branch (`.1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.X`) which Ricoh IM firmware exposes consistently. The startup verification phase will tell you immediately whether a new printer responds correctly.

**Ricoh MP C series** probably works — the OID set here is a superset of what the original MP C version of the script used — but is **untested** in this version. If the per-section breakdown values look wrong, the likely cause is that `.9.9` returns 0 on MP C; the previous version's `.9.26 − .9.10` arithmetic would need to be reintroduced for Printer B&W.

### Feedback

Have you tested this tool with other Ricoh models? Let me know which models work (or don't) for you — feedback helps improve compatibility.
