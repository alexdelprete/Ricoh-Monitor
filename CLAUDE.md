# Ricoh-Monitor — Project Notes for Claude

## What this is

A single-file PowerShell script (`Ricoh-Monitor.ps1`) that polls Ricoh IM-series
MFPs over SNMP v2c, builds a dark-themed HTML report with per-section page
counters and toner bars, and optionally emails it via SMTP (STARTTLS / M365).

One script, one JSON config, one markdown readme. No external modules.

## Fleet / deployment context

- 4 printers on `192.168.1.0/24`: IM C2010 (.146), IM C2000 (.197, .217), IM C400 (.230).
- Operator PC is on `192.168.1.0/24` *or* `192.168.2.0/24` — discovery handles
  cross-subnet by reading known subnets from existing config entries.
- SMTP target: `smtp.office365.com:587` with Authenticated SMTP enabled on the
  mailbox (tenant policy `SmtpClientAuthenticationDisabled` must be off, or per-mailbox override).
- Run cadence: monthly scheduled task (optionally as PS2EXE-compiled `.exe`).

## Files in the repo

- `Ricoh-Monitor.ps1` — the script (only code file).
- `README.md` — user-facing docs: Quick Start, Enabling Email, Scheduling,
  Execution Policy, Config, Parameters, Data Collected, Troubleshooting,
  Supported Models, References.
- `.gitignore` — excludes `Ricoh-Monitor.json`, `report_*.html`, `*.exe`, `*.zip`, editor junk.
- `.markdownlint.json` — `{default:true, MD013:false}`.
- `CLAUDE.md` — this file.

**Not in the repo** (gitignored): `Ricoh-Monitor.json` (contains SMTP creds),
generated `report_*.html` files, PS2EXE build artifacts.

## Script architecture

### Entry points (mutually exclusive)

- `-Discover` → `Invoke-Discovery` (full rescan: local NIC subnet + known subnets + prompt for extras).
- `-TestSnmp <IP>` → `Invoke-RawSnmpDiagnostic` (dumps request/response bytes for MAC + device IP OIDs).
- `-TestSmtp` → `Invoke-SmtpTest` (sends a minimal test email using config credentials).
- No switch → normal pipeline: first-run discovery if no config, else verify → collect → report → optionally email → always write HTML to disk.

### Data flow (normal run)

1. `Get-PrintersConfig` — load or generate `Ricoh-Monitor.json`.
2. `Invoke-NewPrinterScan` — rescan known /24s for new Ricoh hosts.
3. For each printer: `Test-RicohPrinter` (fast SNMP probe on Manufacturer + Model OIDs) → classify `Verified | WrongDevice | NoResponse`.
4. For each `Verified`: `Get-SnmpData` queries the 18-OID `$OIDs` table via `OlePrn.OleSNMP`.
5. Raw UDP SNMP (`Get-SnmpMacAddress`, `Get-SnmpDeviceIp`) for the two OIDs OlePrn corrupts.
6. `Build-HtmlReport` + `Format-PrinterCard` → dark-theme 2-column grid with toner bars and OK/error badges.
7. HTML always written to `report_yyyy-MM-dd_HHmm.html`. Email sent only if `SendEmail:true`.

### OID set (all under `.1.3.6.1.4.1.367.3.2.1`)

Identity (verification phase):
- `.1.1.7.0` ricohSysOemID (Manufacturer)
- `.1.1.1.0` ricohSysName (Model)

Counter phase (`$OIDs`):
- `.2.1.4.0` Serial
- `.2.19.5.1.9.{1,8,2}` Total / Printer / Copier scalar
- `.2.19.5.1.9.{5,3,7,4}` Copier Full / B&W / Single / Two-color
- `.2.19.5.1.9.{11,9,41,42}` Printer Full / B&W / Single / Two-color
- `.2.24.1.1.5.{1..4}` Toner KCMY
- `.2.2.13.0` Error State

Raw SNMP (OlePrn can't decode):
- `.7.2.1.7.0` MAC (OctetString with high bytes)
- `.7.2.1.3.0` IP (IpAddress type)

Toner sentinels: `-100`=Almost Empty, `0`=Empty, `-2`=Unknown, `-3`=Some Left.
Error codes: `0`=noError, `2`=feedError, `3`=hardwareError, `4`=servicemanCall.

### Raw SNMP implementation

Hand-rolled BER in `List[byte]` (never `byte[] + byte[]` — that promotes to
`object[]` and silently mangles packets). Helpers: `ConvertTo-BerLength`,
`ConvertTo-BerInteger`, `ConvertTo-BerOid`, `Add-BerTlv`, `New-SnmpGetRequest`,
`Read-BerHeader`, `Get-SnmpValueFromResponse`, `Invoke-RawSnmpGet`.

Transport: `System.Net.Sockets.UdpClient`, port 161, 2s timeout.

## Key constraints (durable)

- **No external dependencies.** Windows-built-in only (`OlePrn.OleSNMP`,
  `System.Net.Sockets.UdpClient`, `Send-MailMessage`). No SharpSnmpLib, no
  ThreadJob, no PS modules to install.
- **PowerShell 5.1 compatible** (also runs on 7). Italian Windows default
  encoding is Windows-1252; script is ASCII-only — no em-dashes, no BOM.
- **Lint-clean**: PSScriptAnalyzer 0 warnings, markdownlint-cli2 0 errors.
- **No secrets in repo.** `Ricoh-Monitor.json` is gitignored; confirm `git log`
  never exposed SMTP creds before pushing.
- **First-run UX matters.** Script must self-bootstrap: no config → discover,
  prompt, write default, exit with guidance.

## Known quirks

- **OlePrn UTF-8 marshaling** replaces bytes ≥0x80 with U+FFFD in IpAddress /
  high-byte OctetString returns. That's why MAC and device IP go through raw UDP.
- **Exchange GAL display-name override**: internal recipients see the mailbox's
  GAL display name regardless of RFC 5322 `FromName`. External recipients (e.g.
  Gmail) see `FromName` correctly. Not a bug.
- **Mark-of-the-Web**: Google Drive sync adds a Zone.Identifier ADS that blocks
  `RemoteSigned` execution. Fix: `Unblock-File Ricoh-Monitor.ps1` or copy to a
  local folder. README covers this in Execution Policy notes.
- **Send-MailMessage** is marked obsolete by Microsoft but is the only built-in
  SMTP client. Warning suppressed with `-WarningAction SilentlyContinue`.
  Non-terminating errors require `-ErrorAction Stop` to be catchable.

## References used during development

See README §References for full list. Primary sources:

- Ricoh Private MIB Spec Part 4 (v4.050-4 April 2012 — PDF mirrors on Internet
  Archive + iobroker forum; v4.260 is current official via RiDP).
- Per-device Web Image Monitor (`http://<ip>/web/guest/en/websys/webArch/mainFrame.cgi`)
  as ground-truth for IM-series counter index meanings.
- RFC 3805 (Printer MIB v2), RFC 3411, RFC 3412.

## When editing

- Changes to OID selection: cross-check against the PDF spec *and* at least one
  printer's Web Image Monitor counter page before committing.
- Changes to HTML/report: the user prefers dark theme, 2-column grid, modern
  look. Don't regress to the original bright/table layout.
- Changes to discovery: keep parallel probing via runspace pool (no ThreadJob dep).
- Don't reintroduce hardcoded SMTP credentials.
- Don't add modules / NuGet packages / external tools as dependencies.
- Keep README happy-path driven with sample transcripts.
