# Ricoh Printer SNMP Monitoring Tool

![PowerShell](https://img.shields.io/badge/PowerShell-v5.1+-blue.svg)
![Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)
![Deployment](https://img.shields.io/badge/Schedule-Monthly%201st-brightgreen)
![Device](https://img.shields.io/badge/Device-Ricoh%20Printer-red.svg)
![Protocol](https://img.shields.io/badge/Protocol-SNMP-blue.svg)

A PowerShell script (compilable to EXE) that monitors Ricoh IM-series printers via SNMP and sends scheduled email reports with toner levels and per-section page counts.

## Features

- Monthly automated execution (1st day of month)
- Compilable to a standalone EXE for easy deployment
- Collects printer metrics via SNMP (v2c)
- Tracks toner levels (CMYK), grand totals, and per-section breakdown (Copier and Printer: Full Color / B&W / Single Color / Two-color)
- Startup identity-verification phase that distinguishes "no SNMP response" from "wrong device" before collecting counters
- Auto-discovery of new Ricoh printers on the local subnet (first run and every subsequent run re-scans known `/24`s)
- Email sending controlled by a `SendEmail` flag in the config file; the HTML report is always written to disk regardless
- Central configuration via JSON file (auto-generated on first run)
- Automatic retry mechanism (3 attempts) for counter collection

## System Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+ (Windows PowerShell or PowerShell 7+)
- Network access to printers on SNMP port 161 (UDP)
- SMTP server access for email notifications

## Deployment

### 1. Compile the script (optional, Admin PowerShell)

```powershell
Install-Module -Name PS2EXE -Scope CurrentUser
Invoke-PS2EXE -InputFile ".\Ricoh-Monitor.ps1" -OutputFile ".\Ricoh-Monitor.exe"
```

Running as a plain `.ps1` also works, but the EXE avoids execution-policy prompts when scheduled.

### 2. Install

Create a deployment folder (example: `C:\Ricoh-Monitor\`) and place these files in it:

- `Ricoh-Monitor.exe` (or `Ricoh-Monitor.ps1`)
- `Ricoh-Monitor.json` — auto-created on first run if missing

### 3. Configure

Edit `C:\Ricoh-Monitor\Ricoh-Monitor.json`:

```json
{
    "SendEmail": false,
    "Smtp": {
        "Server":   "smtp.example.com",
        "Port":     587,
        "Username": "user@example.com",
        "Password": "your-mailbox-or-app-password",
        "From":     "user@example.com",
        "To":       ["recipient1@example.com", "recipient2@example.com"]
    },
    "Printers": [
        { "IP": "192.168.1.146", "SnmpCommunity": "public" },
        { "IP": "192.168.1.230", "SnmpCommunity": "public" }
    ]
}
```

- `SendEmail` — when `false` the script still runs and writes a timestamped HTML report (`report_yyyy-MM-dd_HHmm.html`) next to the EXE, but does **not** send the email. Switch to `true` once you've validated the report content.
- `Smtp` — server, credentials and recipients used when `SendEmail` is `true`. Port `587` with STARTTLS is the standard. For Microsoft 365 mailboxes with MFA you need an app password (or ask IT to enable Authenticated SMTP for the mailbox).
- `Printers` — list of `{IP, SnmpCommunity}` entries. The script can also auto-discover them on first run or via `-Discover`.

On first run, if `Ricoh-Monitor.json` is missing, the script auto-discovers Ricoh printers on the local subnet, writes them into a fresh config (with `SendEmail` set to `false` and `Smtp` filled with placeholder values), and continues. Edit the `Smtp` section before flipping `SendEmail` to `true`.

> **Note:** the config file contains credentials and is therefore listed in `.gitignore` — only the script and README are tracked in the repo.

### 4. Schedule the task (Admin PowerShell)

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
    -Description "Executes monthly printer monitoring on the 1st at 9:00 AM"
```

## Command-line Parameters

- *(no parameter)* — Normal run: verify -> collect -> report. HTML always written to disk; email sent only if `SendEmail` is `true`.
- `-Discover` (switch) — Interactive discovery: scans the local NIC subnet, then prompts for additional CIDRs. Asks y/N before merging into the config.
- `-TestSnmp <IP>` (string) — Raw-SNMP diagnostic for the given IP. Dumps encoded request, raw response, parsed value. Use when MAC/IP comes back `Unavailable`.
- `-TestSmtp` (switch) — Sends a small test email using the current `Smtp` config. Does not require `SendEmail` to be `true`. Use to validate credentials in isolation.

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

CMYK %, under `.1.3.6.1.4.1.367.3.2.1.2.24.1.1.5`:

- Black `.1`, Cyan `.2`, Magenta `.3`, Yellow `.4`

### Status

- Error State — `.1.3.6.1.4.1.367.3.2.1.2.2.13.0`

## Startup Verification

Before collecting counters, the script probes each configured IP with two SNMP gets (Manufacturer + Model) and prints one of three statuses:

```text
[OK]           192.168.1.146  - RICOH IM C2010
[WRONG DEVICE] 192.168.1.50   - Manufacturer: HP, Model: LaserJet Pro M404
[NO SNMP]      192.168.1.99   - timeout
```

Counter collection runs only on `[OK]` printers. The other two still appear in the report with their status so issues aren't silently dropped.

## Troubleshooting

### Printers not responding

- Verify SNMP is enabled on the printer (Device Settings -> Network -> SNMP)
- Check that the `SnmpCommunity` in the config matches the printer's community string
- Test connectivity with `Test-NetConnection <PrinterIP> -Port 161`
- Run `.\Ricoh-Monitor.ps1 -TestSnmp <PrinterIP>` to see exactly where the SNMP pipeline fails

### Email failures

- Run `.\Ricoh-Monitor.ps1 -TestSmtp` to exercise the SMTP path in isolation
- Verify SMTP credentials in `Ricoh-Monitor.json`
- For Microsoft 365: ensure Authenticated SMTP is enabled on the mailbox (admin -> Users -> Mail -> Manage email apps)
- Check TLS requirements — the script forces TLS 1.2 before STARTTLS
- Review Windows Event Logs for additional SMTP errors

### First-run issues

- Run the EXE or `.ps1` manually once to generate the config file
- Ensure the deployment folder has write permissions for the running account
- If no printers are found, the script writes an empty config — edit it manually and re-run

## Supported Models

Confirmed on:

- Ricoh IM C2010
- Ricoh IM C2000
- Ricoh IM C400

Other Ricoh IM C models are likely to work — the OID set used is the standard `RicohPrivateMIB` table branch (`.1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.X`) and Ricoh IM firmware exposes this branch consistently. The startup verification phase will tell you immediately whether a new printer responds correctly.

**Ricoh MP C series** probably works — the OID set here is a superset of what the original MP C version of the script used — but is **untested** in this version. If the per-section breakdown values look wrong, the most likely cause is that `.9.9` returns 0 on MP C; in that case the previous version's `.9.26 - .9.10` arithmetic would need to be reintroduced for Printer B&W.

### Feedback

Have you tested this tool with other Ricoh models? Let me know which models work (or don't) for you — feedback helps improve compatibility.
