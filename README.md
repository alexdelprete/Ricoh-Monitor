# Ricoh Printer SNMP Monitoring Tool

![PowerShell](https://img.shields.io/badge/PowerShell-v5.1+-blue.svg)  
![Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)   
![Deployment](https://img.shields.io/badge/Schedule-Monthly%201st-brightgreen)   
![Device](https://img.shields.io/badge/Device-Ricoh%20Printer-red.svg)   
![Protocol](https://img.shields.io/badge/Protocol-SNMP-blue.svg)   



A compiled PowerShell executable that monitors Ricoh IM-series printers via SNMP and sends scheduled email reports with toner levels and per-section page counts.

## Features

- Monthly automated execution (1st day of month)
- Compiled EXE for easy deployment
- Collects printer metrics via SNMP (v2c)
- Tracks toner levels (CMYK), grand totals, and per-section breakdown (Copier and Printer: Full Color / B&W / Single Color / Two-color)
- Startup identity-verification phase that distinguishes "no SNMP response" from "wrong device" before collecting counters
- Email sending controlled by a `SendEmail` flag in the config file; when off, the HTML report is written to disk for review
- Central configuration via JSON file (auto-generated on first run)
- Automatic retry mechanism (3 attempts) for counter collection

## System Requirements

- Windows 10/11 or Windows Server 2016+
- .NET Framework 4.7.2 or later
- Network access to printers on SNMP port (161)
- SMTP server access for email notifications

## Deployment

1. **Compile the Script** (Admin PowerShell):
   ```powershell
   Invoke-PS2EXE Invoke-PS2EXE -InputFile "NAME_OF_SCRIPT.ps1" -OutputFile "APP_NAME.exe" -IconFile "ICON.ico" -Title "TITLE" -Company "COMPANY" -Product "PRODUCT" -Description "DESCRIPTION OF APP"


## Installation:

Create folder: C:\Printer Monitor\

Place these files in the folder:

- MPC Monitor.exe
- printers_config.json (auto-created on first run if missing)

### Configuration:

Edit `C:\Printer Monitor\printers_config.json`:

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

- `SendEmail`: when `false` the script still runs and writes a timestamped HTML report (`report_yyyy-MM-dd_HHmm.html`) next to the EXE, but does **not** send the email. Switch to `true` once you've validated the report content.
- `Smtp`: server, credentials and recipients used when `SendEmail` is `true`. Port `587` with STARTTLS is the standard. For Microsoft 365 mailboxes with MFA you need an app password (or ask IT to enable Authenticated SMTP for the mailbox).
- `Printers`: list of `{IP, Community}` entries. The script can also auto-discover them on first run or via `-Discover`.

On first run, if `printers_config.json` is missing, the script auto-discovers Ricoh printers on the local subnet, writes them into a fresh config (with `SendEmail` set to `false` and `Smtp` filled with placeholder values), and continues. Edit the `Smtp` section before flipping `SendEmail` to `true`.

> **Note:** the config file contains credentials and is therefore listed in `.gitignore` — only the script and README are tracked in the repo.


### Schedule the Task (Admin PowerShell):

```powershell
$Action = New-ScheduledTaskAction -Execute "C:\Printer Monitor\APP_NAME.exe"
$Trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "9:00AM"
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd

Register-ScheduledTask `
    -TaskName "Monthly Ricoh Report" `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -RunLevel Highest `
    -Description "Executes monthly printer monitoring on the 1st at 9:00 AM"
```

### Data Collected

Identity:
- Manufacturer — `.1.3.6.1.4.1.367.3.2.1.1.1.7.0`
- Model — `.1.3.6.1.4.1.367.3.2.1.1.1.1.0`
- Serial Number — `.1.3.6.1.4.1.367.3.2.1.2.1.4.0`
- IP Address (device-reported) — `.1.3.6.1.4.1.367.3.2.1.7.2.1.3.0`

Page counters (all from the `.1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.X` table branch):
- Total Pages — `.9.1`
- Total Pages Printer — `.9.8`
- Total Pages Copier — `.9.2`
- Copier breakdown: Full Color `.9.5`, B&W `.9.3`, Single Color `.9.7`, Two-color `.9.4`
- Printer breakdown: Full Color `.9.11`, B&W `.9.9`, Single Color `.9.41`, Two-color `.9.42`

Toner levels (CMYK %):
- Black `.5.1`, Cyan `.5.2`, Magenta `.5.3`, Yellow `.5.4` (under `.1.3.6.1.4.1.367.3.2.1.2.24.1.1`)

Status:
- Error State — `.1.3.6.1.4.1.367.3.2.1.2.2.13.0`

### Startup Verification

Before collecting counters, the script probes each configured IP with two SNMP gets (Manufacturer + Model) and prints one of three statuses:

```
[OK]           192.168.1.146  - RICOH IM C2010
[WRONG DEVICE] 192.168.1.50   - Manufacturer: HP, Model: LaserJet Pro M404
[NO SNMP]      192.168.1.99   - timeout
```

Counter collection runs only on `[OK]` printers. The other two still appear in the report with their status so issues aren't silently dropped.

### Troubleshooting

##### Printers not responding:

Verify SNMP is enabled on printers

Check community strings match

Test connectivity (Test-NetConnection <IP> -Port 161)

#### Email failures:

Verify SMTP credentials

Check TLS requirements for your email provider

Review Windows Event Logs for errors

##### First-run issues:

Run EXE manually to generate config file

Ensure folder has write permissions


### ✅ Supported Models

Confirmed on:
- Ricoh IM C2010
- Ricoh IM C2000
- Ricoh IM C400

Other Ricoh IM C models are likely to work — the OID set used is the standard `RicohPrivateMIB` table branch (`.1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.X`) and Ricoh IM firmware exposes this branch consistently. The startup verification phase will tell you immediately whether a new printer responds correctly.

**Ricoh MP C series**: probably works — the OID set chosen here is a superset of what the original MP C version of the script used — but is **untested** in this version. If you try it on an MP C and the per-section breakdown values look wrong, the most likely cause is that `.9.9` returns 0 on MP C; in that case the previous version's `.9.26 - .9.10` arithmetic would need to be reintroduced for Printer B&W.

#### Have you tested this tool with other Ricoh models?
Let me know which models work (or don't work) for you – feedback helps improve compatibility.





