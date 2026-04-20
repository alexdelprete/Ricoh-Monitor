<#
.NOTES
    Version: 3.3
    Authors: Samuel Jesus, Alessandro Del Prete

    SNMP OID reference (Ricoh Private MIB Specification, Part 4):
        Official (current, registration required) - Ricoh Developer Program:
            https://ricoh-ridp.com/resources/downloads/private-mib-specification-v4260
        Older version used during development (v4.050-4, Apr 2012):
            https://archive.org/details/294-privatemibspecificationv-4-050-4
            https://forum.iobroker.net/assets/uploads/files/294_privatemibspecificationv4_050-4.pdf

    Every OID this script reads is named in the official spec:
      - .1.1.*       sysDescr group       (model, manufacturer, serial)
      - .2.2.*       engStat group        (error codes)
      - .2.19.*      engCounter group     (page counters: aggregates + table)
      - .2.24.*      engToner group       (toner levels K/C/M/Y)
      - .7.2.1.*     ricohNetIp group     (MAC address)

    The per-counter index meanings under ricohEngCounterValue (.2.19.5.1.9.X)
    are defined in per-model-family tables in the spec. The Ricoh IM-series
    mapping isn't in the 2012 PDF; the indices we use (1, 2, 3, 4, 5, 7, 8, 9,
    11, 41, 42) were identified by snmpwalk on an IM C2000 cross-checked
    against the printer's web counter page. Each is annotated below.

.PARAMETER Discover
    Run discovery instead of normal monitoring. Scans the local NIC's subnet
    (auto-detected) for Ricoh printers; if none are found there, prompts for
    additional CIDR ranges to scan. Discovered printers can be optionally
    added to Ricoh-Monitor.json.

.PARAMETER TestSnmp
    Diagnostic: send raw SNMP GetRequest packets to the given IP for MAC
    and device-IP OIDs. Dumps the encoded request, the raw response, and
    the parsed value. Use when the MAC/IP fields come back Unavailable to
    see where the pipeline is failing.

.PARAMETER TestSmtp
    Diagnostic: sends a small test email using the Smtp section from
    Ricoh-Monitor.json. Prints each step (config summary, send result,
    and the real error text if it fails) so SMTP problems can be debugged
    in isolation without running a full monitoring pass.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '',
    Justification = 'Interactive CLI; Write-Host is the correct tool for colored console output and is not captured as pipeline data.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '',
    Justification = 'Password is read from the local Ricoh-Monitor.json config (by design outside source control) and must be converted to a SecureString to build the PSCredential for Send-MailMessage.')]
param(
    [switch]$Discover,
    [string]$TestSnmp,
    [switch]$TestSmtp
)

# SMTP settings live in Ricoh-Monitor.json (the Smtp section), not in the
# script. This keeps credentials out of the repo and makes it possible to ship
# the same script/EXE to multiple sites without recompiling.

# -----------------------------------------------------------------------------
# Identity OIDs - queried during the verification phase to confirm each IP is a
# Ricoh printer before doing any heavy counter collection. Both are documented
# in the RicohPrivateMIB.
# -----------------------------------------------------------------------------
$IdentityOIDs = [ordered]@{
    "Manufacturer" = ".1.3.6.1.4.1.367.3.2.1.1.1.7.0"   # ricohSysOemID  (spec types it 'riochSysOemID' - typo in the official PDF)
    "Model"        = ".1.3.6.1.4.1.367.3.2.1.1.1.1.0"   # ricohSysName   (DisplayString - "MODEL(MDL) NAME")
}

# -----------------------------------------------------------------------------
# Counter / status / network OIDs - queried per verified printer.
# All OIDs below are named in the official Ricoh MIB Spec (Part 4).
# -----------------------------------------------------------------------------
$OIDs = [ordered]@{
    "Serial Number"        = ".1.3.6.1.4.1.367.3.2.1.2.1.4.0"        # ricohEngSerialNumber (OctetString)

    # NOTE: ricohNetIpPhysicalAddress (.7.2.1.7.0) and ricohNetIp (.7.2.1.3.0)
    # are intentionally NOT in this hashtable. Their SNMP types (OctetString
    # with bytes >= 0x80, IpAddress with 4 raw bytes) get mangled by OlePrn's
    # UTF-8 marshalling path. They're fetched separately in Get-SnmpData via
    # Invoke-RawSnmpGet, which uses raw UDP + BER and preserves every byte.

    # Aggregate page totals - scalar leaves under ricohEngCounter (.2.19).
    "Total Pages"          = ".1.3.6.1.4.1.367.3.2.1.2.19.1.0"       # ricohEngCounterTotal   - "Total of all counters for the devices"
    "Total Pages Printer"  = ".1.3.6.1.4.1.367.3.2.1.2.19.2.0"       # ricohEngCounterPrinter - "Counter for printer application"
    "Total Pages Copier"   = ".1.3.6.1.4.1.367.3.2.1.2.19.4.0"       # ricohEngCounterCopier  - "Counter for copy application"
    # (ricohEngCounterFax = .19.3.0 also exists in the spec but isn't reported.)

    # Per-section breakdown - column ricohEngCounterValue (.2.19.5.1.9.X).
    # The column itself is named in the spec; per-index meanings live in
    # per-model-family tables (Table 3.2.1.2.19.5.1.1-<modelCode>). The
    # Ricoh IM-series table isn't in the 2012 PDF, so the index meanings
    # below were verified by walking an IM C2000 + cross-checking the
    # printer's web counter page. Comments give the verified meaning.
    # Copier:
    "Copier Full Color"    = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.5"   # ricohEngCounterValue index 5 -> Copier Full Color
    "Copier B&W"           = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.3"   # ricohEngCounterValue index 3 -> Copier Black & White
    "Copier Single Color"  = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.7"   # ricohEngCounterValue index 7 -> Copier Single Color
    "Copier Two-color"     = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.4"   # ricohEngCounterValue index 4 -> Copier Two-color
    # Printer:
    "Printer Full Color"   = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.11"  # ricohEngCounterValue index 11 -> Printer Full Color
    "Printer B&W"          = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.9"   # ricohEngCounterValue index  9 -> Printer Black & White (clean - does NOT include 2-color)
    "Printer Single Color" = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.41"  # ricohEngCounterValue index 41 -> Printer Single Color
    "Printer Two-color"    = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.42"  # ricohEngCounterValue index 42 -> Printer Two-color

    # Toner levels - column ricohEngTonerLevel (.2.24.1.1.5.X), OctetString.
    # Per spec values: 100..20 in steps of 10 = remaining %; -100 = "near
    # empty" (10-1% remaining); 0 = "empty"; -2 = "unknown / cannot measure";
    # -3 = "some remaining" (Type-M4 mode). Toner index 1=K, 2=C, 3=M, 4=Y
    # per Table 3.2.1.2.24.1 (4-color models).
    "Black Toner %"        = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.1"   # ricohEngTonerLevel idx 1 = Black
    "Cyan Toner %"         = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.2"   # ricohEngTonerLevel idx 2 = Cyan
    "Magenta Toner %"      = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.3"   # ricohEngTonerLevel idx 3 = Magenta
    "Yellow Toner %"       = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.4"   # ricohEngTonerLevel idx 4 = Yellow

    # Error / status - ricohEngScanStatError values per spec:
    #   0 = noError, 2 = feedError (ADF jam), 3 = hardwareError, 4 = servicemanCall.
    "Error State"          = ".1.3.6.1.4.1.367.3.2.1.2.2.13.0"       # ricohEngScanStatError
}

# Note: the per-section field order in the HTML report is hardcoded inside
# Build-HtmlReport (Page Counters, Copier Breakdown, Printer Breakdown, Toner,
# Status). The list of fields the script actually collects is defined by the
# $OIDs hashtable above.

# -----------------------------------------------------------------------------
# Get-PrintersConfig
#
# Loads Ricoh-Monitor.json from the working directory and returns the parsed
# object. The file is expected to exist by the time this is called - main
# triggers Invoke-FirstRunDiscovery on first run to create it.
#
# Expected shape:
#   { "SendEmail": <bool>, "Printers": [ { "IP": "...", "SnmpCommunity": "..." }, ... ] }
# -----------------------------------------------------------------------------
function Get-PrintersConfig {
    param(
        [string]$ConfigPath = "Ricoh-Monitor.json"
    )

    if (-not (Test-Path $ConfigPath)) {
        # Defensive: should never hit this because main runs first-run discovery
        # when the file is absent. Bail out clearly if it ever does.
        Write-Host "Config file '$ConfigPath' not found." -ForegroundColor Red
        exit 1
    }

    try {
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json -ErrorAction Stop
        # Sanity check: bail out early with a clear message if the file is using
        # the old flat-array shape from v1/v2 of the script.
        if ($null -eq $config.Printers) {
            Write-Host "Config file is missing the 'Printers' array. Expected shape: { SendEmail: bool, Printers: [...] }" -ForegroundColor Red
            exit 1
        }
        return $config
    }
    catch {
        Write-Host "Error reading configuration file: $_" -ForegroundColor Red
        exit 1
    }
}

# -----------------------------------------------------------------------------
# Test-RicohPrinter
#
# Lightweight identity probe. Opens an SNMP session with a short timeout and
# fetches only Manufacturer + Model. Used during the verification phase to
# decide whether to do the full counter collection on this IP.
#
# Returns a PSCustomObject with:
#   IP           - the IP that was probed
#   Status       - "Verified" | "WrongDevice" | "NoResponse"
#   Manufacturer - string from SNMP, or $null
#   Model        - string from SNMP, or $null
#   Error        - exception message when Status = NoResponse
#
# Status semantics:
#   Verified     - SNMP responded and Manufacturer string contains "RICOH"
#   WrongDevice  - SNMP responded but it isn't a Ricoh (e.g. another brand
#                  printer or a switch on that IP)
#   NoResponse   - SNMP didn't answer (timeout, wrong community, SNMP off,
#                  host down)
# -----------------------------------------------------------------------------
function Test-RicohPrinter {
    param(
        [string]$IP,
        [string]$SnmpCommunity
    )

    $result = [ordered]@{
        IP           = $IP
        Status       = $null
        Manufacturer = $null
        Model        = $null
        Error        = $null
    }

    try {
        $snmp = New-Object -ComObject "OlePrn.OleSNMP"
        # Open(target, community, retries, timeoutMs) - short timeout / single
        # retry: this is just the fast verification probe, not data collection.
        $snmp.Open($IP, $SnmpCommunity, 1, 1500)

        try {
            # Two-OID identity fetch. If either Get throws we treat the whole
            # probe as NoResponse (the printer is reachable enough to answer
            # one of these OIDs or it isn't a usable target).
            $result.Manufacturer = $snmp.Get($IdentityOIDs.Manufacturer)
            $result.Model        = $snmp.Get($IdentityOIDs.Model)
        }
        catch {
            $snmp.Close()
            $result.Status = "NoResponse"
            $result.Error  = $_.Exception.Message
            return [pscustomobject]$result
        }

        $snmp.Close()

        # Manufacturer must contain "RICOH" (case-insensitive). Anything else
        # - including an empty string - is a wrong device.
        if ($result.Manufacturer -and ("$($result.Manufacturer)" -match "RICOH")) {
            $result.Status = "Verified"
        } else {
            $result.Status = "WrongDevice"
        }
    }
    catch {
        # Outer catch handles failures opening the SNMP COM object itself
        # (e.g. host unreachable before we even sent a Get).
        $result.Status = "NoResponse"
        $result.Error  = $_.Exception.Message
    }

    return [pscustomobject]$result
}

# -----------------------------------------------------------------------------
# Get-SnmpData
#
# Full counter collection for a single, already-verified printer.
# Re-uses Manufacturer/Model from the verification result (no second lookup),
# then iterates over the $OIDs hashtable doing one Get per entry inside a
# single SNMP session. Per-OID Get errors are stored as the literal "Error"
# string so the report still renders the row.
#
# If the entire SNMP session fails to open, retries up to $MaxRetries times
# with a 2-second sleep between attempts. After exhausting retries, every
# OID value is set to "Unavailable" and Status is set to "Failed after N attempts".
# -----------------------------------------------------------------------------
function Get-SnmpData {
    param(
        [object]$Verification,    # PSCustomObject from Test-RicohPrinter
        [string]$SnmpCommunity,
        [int]$MaxRetries = 3
    )

    $IP = $Verification.IP
    $result = [ordered]@{
        "IP"           = $IP
        "Status"       = "Verified"
        "Manufacturer" = $Verification.Manufacturer
        "Model"        = $Verification.Model
    }

    $retryCount = 0
    $success    = $false

    while ($retryCount -lt $MaxRetries -and -not $success) {
        try {
            $snmp = New-Object -ComObject "OlePrn.OleSNMP"
            # Longer timeout (3000 ms) and 2 internal retries here - counter
            # collection is the part we actually care about getting right.
            $snmp.Open($IP, $SnmpCommunity, 2, 3000)

            # Per-OID loop: each Get is wrapped so a single OID failure
            # (e.g. printer doesn't expose one of the breakdown OIDs)
            # doesn't abort the whole collection.
            foreach ($oid in $OIDs.GetEnumerator()) {
                try {
                    $value = $snmp.Get($oid.Value)
                    $result[$oid.Name] = $value
                }
                catch {
                    $result[$oid.Name] = "Error"
                }
            }

            $snmp.Close()

            # Fetch MAC and device-reported IP via raw SNMP - both have SNMP
            # types (OctetString w/ high bytes, IpAddress) that OlePrn mangles
            # through its UTF-8 marshalling path. Raw UDP+BER preserves bytes.
            $mac = Get-SnmpMacAddress -IP $IP -SnmpCommunity $SnmpCommunity
            $result["MAC Address"] = if ($mac) { $mac } else { "Unavailable" }

            $devIp = Get-SnmpDeviceIp -IP $IP -SnmpCommunity $SnmpCommunity
            $result["IP Address (device)"] = if ($devIp) { $devIp } else { "Unavailable" }

            $success = $true
        }
        catch {
            # Whole-session failure - sleep and retry, or give up after MaxRetries.
            $retryCount++
            if ($retryCount -eq $MaxRetries) {
                $result["Status"] = "Failed after $MaxRetries attempts"
                foreach ($oid in $OIDs.GetEnumerator()) {
                    $result[$oid.Name] = "Unavailable"
                }
            }
            Start-Sleep -Seconds 2
        }
    }

    return $result
}

# -----------------------------------------------------------------------------
# Format-Counter
#
# Returns an integer counter formatted with thousands separators (e.g. 14415 ->
# "14,415"). Pass-through for non-numeric values ("Error", "Unavailable") so
# the report still shows a meaningful cell instead of a crash.
# -----------------------------------------------------------------------------
function Format-Counter {
    param($value)
    if ($value -is [int] -or ("$value" -match '^-?\d+$')) {
        return ("{0:N0}" -f [int]$value)
    }
    return $value
}

# =============================================================================
# RAW SNMP HELPERS
#
# Minimal self-contained SNMPv2c GetRequest over UDP, with a hand-rolled BER
# encoder/decoder. Used for OIDs whose SNMP-level types (IpAddress, OctetString
# with bytes >= 0x80) get mangled by OlePrn's UTF-8 marshalling. No external
# library: just System.Net.Sockets.UdpClient, which ships with .NET Framework
# and .NET on every Windows install.
# =============================================================================

# --- BER primitives ---------------------------------------------------------

# BER length octets: <128 = direct, >=128 = 0x80|n followed by n big-endian bytes.
function ConvertTo-BerLength {
    param([int]$Length)
    if ($Length -lt 128) { return ,[byte]$Length }
    $bytes = [System.Collections.Generic.List[byte]]::new()
    $n = $Length
    while ($n -gt 0) {
        $bytes.Insert(0, [byte]($n -band 0xFF))
        $n = $n -shr 8
    }
    $head = [byte](0x80 -bor $bytes.Count)
    return ,(@($head) + $bytes.ToArray())
}

# BER INTEGER encoding: big-endian two's complement, minimum bytes, sign-preserving.
function ConvertTo-BerInteger {
    param([int]$Value)
    if ($Value -eq 0) { return ,[byte]0x00 }
    $bytes = [BitConverter]::GetBytes([int32]$Value)
    [Array]::Reverse($bytes)  # BitConverter is little-endian on x86; BER wants big-endian
    $i = 0
    while ($i -lt $bytes.Length - 1 -and $bytes[$i] -eq 0 -and ($bytes[$i + 1] -band 0x80) -eq 0) {
        $i++
    }
    return ,([byte[]]$bytes[$i..($bytes.Length - 1)])
}

# BER OID encoding: first two sub-ids folded into 40*a+b, rest variable-length 7-bit groups.
function ConvertTo-BerOid {
    param([string]$Oid)
    $parts = $Oid.TrimStart('.').Split('.') | ForEach-Object { [int]$_ }
    $out = [System.Collections.Generic.List[byte]]::new()
    $out.Add([byte](40 * $parts[0] + $parts[1]))
    for ($i = 2; $i -lt $parts.Length; $i++) {
        $n = $parts[$i]
        if ($n -lt 128) {
            $out.Add([byte]$n)
        } else {
            $stack = [System.Collections.Generic.Stack[byte]]::new()
            $stack.Push([byte]($n -band 0x7F))
            $n = $n -shr 7
            while ($n -gt 0) {
                $stack.Push([byte](($n -band 0x7F) -bor 0x80))
                $n = $n -shr 7
            }
            while ($stack.Count -gt 0) { $out.Add($stack.Pop()) }
        }
    }
    return ,$out.ToArray()
}

# --- SNMP message build / parse --------------------------------------------

# Appends a TLV to a target List[byte]. Using a shared helper that writes
# straight into a running buffer avoids PowerShell's array-concatenation
# coercion (byte[] + byte[] -> object[]), which is how an earlier version
# of this code produced malformed SNMP packets that printers silently dropped.
function Add-BerTlv {
    param(
        [System.Collections.Generic.List[byte]]$Target,
        [byte]$Tag,
        [byte[]]$Value
    )
    $Target.Add([byte]$Tag)
    foreach ($b in (ConvertTo-BerLength $Value.Length)) { $Target.Add([byte]$b) }
    foreach ($b in $Value) { $Target.Add([byte]$b) }
}

# Builds a complete SNMPv2c GetRequest packet for one OID. Returns byte[].
function New-SnmpGetRequest {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Pure builder function; returns bytes, changes no external state.')]
    param(
        [string]$SnmpCommunity,
        [string]$Oid,
        [int]$RequestId
    )

    $oidBytes  = ConvertTo-BerOid $Oid
    $commBytes = [System.Text.Encoding]::ASCII.GetBytes($SnmpCommunity)
    $reqBytes  = ConvertTo-BerInteger $RequestId

    # varbind inner = OID TLV + NULL TLV
    $vbInner = [System.Collections.Generic.List[byte]]::new()
    Add-BerTlv -Target $vbInner -Tag 0x06 -Value $oidBytes
    $vbInner.Add([byte]0x05); $vbInner.Add([byte]0x00)    # NULL

    # varbind = SEQUENCE { ... }
    $vbOuter = [System.Collections.Generic.List[byte]]::new()
    Add-BerTlv -Target $vbOuter -Tag 0x30 -Value $vbInner.ToArray()

    # variable-bindings = SEQUENCE OF VarBind
    $vbList = [System.Collections.Generic.List[byte]]::new()
    Add-BerTlv -Target $vbList -Tag 0x30 -Value $vbOuter.ToArray()

    # GetRequest-PDU body = { request-id, error-status, error-index, varbinds }
    $pduBody = [System.Collections.Generic.List[byte]]::new()
    Add-BerTlv -Target $pduBody -Tag 0x02 -Value $reqBytes
    Add-BerTlv -Target $pduBody -Tag 0x02 -Value @([byte]0x00)   # error-status = 0
    Add-BerTlv -Target $pduBody -Tag 0x02 -Value @([byte]0x00)   # error-index  = 0
    $pduBody.AddRange($vbList)

    # GetRequest-PDU with [0] tag
    $pdu = [System.Collections.Generic.List[byte]]::new()
    Add-BerTlv -Target $pdu -Tag 0xA0 -Value $pduBody.ToArray()

    # Outer message body = version(=1 for v2c) + community + pdu
    $msgBody = [System.Collections.Generic.List[byte]]::new()
    Add-BerTlv -Target $msgBody -Tag 0x02 -Value @([byte]0x01)   # version = 1 (v2c)
    Add-BerTlv -Target $msgBody -Tag 0x04 -Value $commBytes      # community
    $msgBody.AddRange($pdu)

    # Message = SEQUENCE { ... }
    $msg = [System.Collections.Generic.List[byte]]::new()
    Add-BerTlv -Target $msg -Tag 0x30 -Value $msgBody.ToArray()

    return ,$msg.ToArray()
}

# Reads a BER {tag, length} header at $Offset and returns where the data starts.
function Read-BerHeader {
    param([byte[]]$Buffer, [int]$Offset)
    $tag     = $Buffer[$Offset]
    $lenByte = $Buffer[$Offset + 1]
    if ($lenByte -lt 128) {
        return [pscustomobject]@{
            Tag = $tag; Length = [int]$lenByte
            HeaderLength = 2; DataOffset = $Offset + 2
        }
    }
    $lenLen = $lenByte -band 0x7F
    $len    = 0
    for ($i = 0; $i -lt $lenLen; $i++) {
        $len = ($len -shl 8) -bor $Buffer[$Offset + 2 + $i]
    }
    return [pscustomobject]@{
        Tag = $tag; Length = $len
        HeaderLength = 2 + $lenLen; DataOffset = $Offset + 2 + $lenLen
    }
}

# Walks a SNMP GetResponse and returns the single varbind's value as raw bytes.
# Return: @{ Type = <byte tag>; Bytes = <byte[]> } or $null on error.
function Get-SnmpValueFromResponse {
    param([byte[]]$Response)
    try {
        $p = 0
        $outer = Read-BerHeader $Response $p
        if ($outer.Tag -ne 0x30) { return $null }
        $p = $outer.DataOffset

        for ($i = 0; $i -lt 2; $i++) {   # skip version, community
            $h = Read-BerHeader $Response $p
            $p += $h.HeaderLength + $h.Length
        }

        $pdu = Read-BerHeader $Response $p
        if ($pdu.Tag -ne 0xA2) { return $null }   # must be GetResponse-PDU
        $p = $pdu.DataOffset

        for ($i = 0; $i -lt 3; $i++) {   # skip request-id, error-status, error-index
            $h = Read-BerHeader $Response $p
            $p += $h.HeaderLength + $h.Length
        }

        $vbs = Read-BerHeader $Response $p
        if ($vbs.Tag -ne 0x30) { return $null }
        $p = $vbs.DataOffset

        $vb = Read-BerHeader $Response $p
        if ($vb.Tag -ne 0x30) { return $null }
        $p = $vb.DataOffset

        $oid = Read-BerHeader $Response $p
        $p += $oid.HeaderLength + $oid.Length

        $val = Read-BerHeader $Response $p
        $bytes = New-Object byte[] $val.Length
        if ($val.Length -gt 0) {
            [Array]::Copy($Response, $val.DataOffset, $bytes, 0, $val.Length)
        }
        return [pscustomobject]@{ Type = $val.Tag; Bytes = $bytes }
    } catch {
        return $null
    }
}

# Top-level: send a GetRequest and return the parsed value, or $null on any
# failure (timeout, unreachable, malformed response, error-status != 0).
function Invoke-RawSnmpGet {
    param(
        [string]$IP,
        [string]$SnmpCommunity = 'public',
        [string]$Oid,
        [int]$TimeoutMs = 2000
    )
    $udp = $null
    try {
        $reqId  = Get-Random -Minimum 1 -Maximum 2147483647
        $packet = New-SnmpGetRequest -SnmpCommunity $SnmpCommunity -Oid $Oid -RequestId $reqId

        $udp = New-Object System.Net.Sockets.UdpClient
        $udp.Client.ReceiveTimeout = $TimeoutMs
        [void]$udp.Send($packet, $packet.Length, $IP, 161)
        $ep       = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any, 0)
        $response = $udp.Receive([ref]$ep)
        return Get-SnmpValueFromResponse $response
    } catch {
        return $null
    } finally {
        if ($udp) { $udp.Close() }
    }
}

# --- Typed wrappers --------------------------------------------------------

# Fetch ricohNetIpPhysicalAddress (MAC) and return colon-hex string or $null.
function Get-SnmpMacAddress {
    param([string]$IP, [string]$SnmpCommunity = 'public')
    $r = Invoke-RawSnmpGet -IP $IP -SnmpCommunity $SnmpCommunity -Oid ".1.3.6.1.4.1.367.3.2.1.7.2.1.7.0"
    if ($null -eq $r) { return $null }
    # OctetString tag = 0x04, expect 6 bytes for a MAC.
    if ($r.Type -eq 0x04 -and $r.Bytes.Length -eq 6) {
        return (($r.Bytes | ForEach-Object { '{0:X2}' -f $_ }) -join ':')
    }
    return $null
}

# Fetch ricohNetIp (device-reported IP) and return dotted-quad string or $null.
function Get-SnmpDeviceIp {
    param([string]$IP, [string]$SnmpCommunity = 'public')
    $r = Invoke-RawSnmpGet -IP $IP -SnmpCommunity $SnmpCommunity -Oid ".1.3.6.1.4.1.367.3.2.1.7.2.1.3.0"
    if ($null -eq $r) { return $null }
    # IpAddress tag = 0x40, exactly 4 bytes.
    if ($r.Type -eq 0x40 -and $r.Bytes.Length -eq 4) {
        return ([System.Net.IPAddress]$r.Bytes).ToString()
    }
    return $null
}

# Diagnostic helper for the -TestSnmp parameter. Walks one MAC and one IP
# probe end-to-end and prints every step (encoded request bytes, raw UDP
# response, parsed varbind) so it's obvious where the pipeline fails.
function Invoke-RawSnmpDiagnostic {
    param([string]$IP, [string]$SnmpCommunity = 'public')

    foreach ($probe in @(
        @{ Label = 'MAC';        Oid = '.1.3.6.1.4.1.367.3.2.1.7.2.1.7.0'; ExpectedTag = 0x04 },
        @{ Label = 'Device IP';  Oid = '.1.3.6.1.4.1.367.3.2.1.7.2.1.3.0'; ExpectedTag = 0x40 }
    )) {
        Write-Host ""
        Write-Host "=== Probe: $($probe.Label)  OID $($probe.Oid) ===" -ForegroundColor Cyan

        try {
            $reqId  = Get-Random -Minimum 1 -Maximum 2147483647
            $packet = New-SnmpGetRequest -SnmpCommunity $SnmpCommunity -Oid $probe.Oid -RequestId $reqId
            Write-Host "Request type: $($packet.GetType().FullName), length $($packet.Length)"
            Write-Host ("Request bytes: " + (($packet | ForEach-Object { '{0:X2}' -f $_ }) -join ' '))
        } catch {
            Write-Host "Build failed: $_" -ForegroundColor Red
            continue
        }

        try {
            $udp = New-Object System.Net.Sockets.UdpClient
            $udp.Client.ReceiveTimeout = 3000
            $sent = $udp.Send($packet, $packet.Length, $IP, 161)
            Write-Host "Sent $sent bytes to ${IP}:161"

            $ep       = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any, 0)
            $response = $udp.Receive([ref]$ep)
            Write-Host "Got $($response.Length) bytes from $ep"
            Write-Host ("Response bytes: " + (($response | ForEach-Object { '{0:X2}' -f $_ }) -join ' '))

            $parsed = Get-SnmpValueFromResponse $response
            if ($null -eq $parsed) {
                Write-Host "Parse returned null" -ForegroundColor Yellow
            } else {
                Write-Host ("Parsed: Tag 0x{0:X2}, Length {1}, Bytes {2}" -f `
                    $parsed.Type, $parsed.Bytes.Length,
                    (($parsed.Bytes | ForEach-Object { '{0:X2}' -f $_ }) -join ' '))
                if ($parsed.Type -ne $probe.ExpectedTag) {
                    Write-Host "Tag mismatch: expected 0x$('{0:X2}' -f $probe.ExpectedTag)" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "Send/receive/parse failed: $_" -ForegroundColor Red
        } finally {
            if ($udp) { $udp.Close() }
        }
    }
    Write-Host ""
}

# -----------------------------------------------------------------------------
# Format-TonerBar
#
# Returns an HTML fragment with a colored fill bar + label.
#
# Per the Ricoh MIB Spec (Table 3.2.1.2.24.1.1.5), ricohEngTonerLevel
# returns one of:
#   100..20 (steps of 10) - remaining percentage
#     0                   - toner empty
#    -2                   - unknown (cannot measure / no sensor)
#    -3                   - "some remaining" (Type-M4 mode, no exact %)
#   -100                  - toner near empty (10-1 %)
#
# Sentinels render as colored badges; numeric percentages render as a fill
# bar + percentage label. Layout uses a 2-column inner table (100px bar +
# 110px label) so all four toner rows start their bars at the same X
# position regardless of label width.
# -----------------------------------------------------------------------------
function Format-TonerBar {
    param($value)

    # Non-numeric: just display whatever we got (e.g. "Error", "Unavailable").
    if (-not ($value -is [int] -or ("$value" -match '^-?\d+$'))) {
        return $value
    }

    $pct = [int]$value
    $emptyTrackColor = "#2a3142"
    $emptyTrack      = "<table cellpadding='0' cellspacing='0' border='0' style='width:100px;border-collapse:collapse;'><tr><td style='width:100px;height:8px;background-color:$emptyTrackColor;font-size:0;line-height:0;'>&nbsp;</td></tr></table>"

    # Helper: small inline pill for sentinel labels.
    function Format-TonerBadge { param($bg, $fg, $text)
        "<span style=`"display:inline-block;background:$bg;color:$fg;padding:3px 9px;border-radius:4px;font-size:11px;font-weight:600;white-space:nowrap;`">$text</span>"
    }

    # --- Resolve label by spec sentinel --------------------------------------
    if ($pct -eq 0) {
        $barCell = $emptyTrack
        $label   = Format-TonerBadge "rgba(239,68,68,0.18)" "#f87171" "Empty"
    }
    elseif ($pct -eq -100) {
        $barCell = $emptyTrack
        $label   = Format-TonerBadge "rgba(239,68,68,0.18)" "#f87171" "Almost Empty"
    }
    elseif ($pct -eq -2) {
        $barCell = $emptyTrack
        $label   = Format-TonerBadge "rgba(154,163,178,0.18)" "#9aa3b2" "Unknown"
    }
    elseif ($pct -eq -3) {
        $barCell = $emptyTrack
        $label   = Format-TonerBadge "rgba(234,179,8,0.18)" "#facc15" "Some Left"
    }
    elseif ($pct -lt 0) {
        # Out-of-spec negative - show raw value, badge it as 'check device'.
        $barCell = $emptyTrack
        $label   = Format-TonerBadge "rgba(154,163,178,0.18)" "#9aa3b2" "Unknown ($pct)"
    }
    else {
        # Numeric percentage - color-graded fill bar.
        $color = if ($pct -ge 50) { "#22c55e" }       # green
                 elseif ($pct -ge 25) { "#eab308" }   # yellow
                 elseif ($pct -ge 10) { "#f97316" }   # orange
                 else                 { "#ef4444" }   # red

        $totalPx = 100
        $fillPx  = [Math]::Min($totalPx, [Math]::Max(0, [int]($pct * $totalPx / 100)))
        $emptyPx = $totalPx - $fillPx

        $barCell = "<table cellpadding='0' cellspacing='0' border='0' style='width:100px;border-collapse:collapse;'><tr>"
        if ($fillPx  -gt 0) { $barCell += "<td style='width:${fillPx}px;height:8px;background-color:$color;font-size:0;line-height:0;'>&nbsp;</td>" }
        if ($emptyPx -gt 0) { $barCell += "<td style='width:${emptyPx}px;height:8px;background-color:$emptyTrackColor;font-size:0;line-height:0;'>&nbsp;</td>" }
        $barCell += "</tr></table>"

        $label = "<span style='font-variant-numeric:tabular-nums;color:#e6e8eb;'>${pct}%</span>"
    }

    # --- Outer 2-column table forces consistent bar X-position across rows ---
    return "<table align='right' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;'><tr><td style='width:100px;padding-right:14px;vertical-align:middle;'>$barCell</td><td style='width:110px;text-align:right;vertical-align:middle;'>$label</td></tr></table>"
}

# -----------------------------------------------------------------------------
# Format-ErrorState
#
# Ricoh error code: 0 means no error. Renders 0 as a green "OK" pill and any
# non-zero value as a red "Code N" pill. Pass-through for non-numeric values.
# -----------------------------------------------------------------------------
function Format-ErrorState {
    param($value)
    if (-not ($value -is [int] -or ("$value" -match '^-?\d+$'))) {
        return $value
    }

    $code = [int]$value

    # Translate per spec (RicohEngScanStatErrorTC):
    #   0 = noError, 2 = feedError (ADF jam), 3 = hardwareError, 4 = servicemanCall
    $label = switch ($code) {
        0       { 'OK' }
        2       { 'ADF Jam' }
        3       { 'Hardware Error' }
        4       { 'Service Call' }
        default { "Code $code" }
    }

    if ($code -eq 0) {
        return "<span style=`"display:inline-block;background:rgba(34,197,94,0.18);color:#4ade80;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:600;`">$label</span>"
    }
    return "<span style=`"display:inline-block;background:rgba(239,68,68,0.18);color:#f87171;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:600;`">$label</span>"
}

# -----------------------------------------------------------------------------
# Format-PrinterCard
#
# Returns the inner HTML for one printer card (dark theme). The Verified path
# emits a full data card with sectioned counters/toner/status; the other paths
# (WrongDevice, NoResponse, Failed-after-N) emit compact status cards so
# problem printers remain visible without empty data rows.
#
# This function returns the *card body* only - it's wrapped in the 2-column
# grid cell by Build-HtmlReport.
# -----------------------------------------------------------------------------
function Format-PrinterCard {
    param([object]$printer)

    $status = $printer['Status']
    $ip     = $printer['IP']

    if ($status -eq 'NoResponse') {
        $errMsg = if ($printer['Error']) { $printer['Error'] } else { 'No SNMP response' }
        return @"
<div class="card card-err">
  <table class="card-hdr"><tr><td class="card-name">Unreachable</td><td class="card-ip">$ip</td></tr></table>
  <p class="status-line"><strong>No SNMP response</strong> &mdash; $errMsg</p>
</div>
"@
    }

    if ($status -eq 'WrongDevice') {
        $mfg   = if ($printer['Manufacturer']) { $printer['Manufacturer'] } else { '<empty>' }
        $model = if ($printer['Model'])        { $printer['Model'] }        else { '<empty>' }
        return @"
<div class="card card-warn">
  <table class="card-hdr"><tr><td class="card-name">Not a Ricoh printer</td><td class="card-ip">$ip</td></tr></table>
  <p class="status-line">SNMP responded with <strong>$mfg / $model</strong> &mdash; check the IP in the config.</p>
</div>
"@
    }

    if ("$status" -like 'Failed after*') {
        return @"
<div class="card card-err">
  <table class="card-hdr"><tr><td class="card-name">$($printer['Model'])</td><td class="card-ip">$ip</td></tr></table>
  <p class="status-line"><strong>Counter collection failed:</strong> $status</p>
</div>
"@
    }

    # Verified - full data card.
    $title = if ($printer['Manufacturer']) { "$($printer['Manufacturer']) $($printer['Model'])" } else { $printer['Model'] }
    return @"
<div class="card">
  <table class="card-hdr">
    <tr>
      <td class="card-name">$title</td>
      <td class="card-ip">$ip</td>
    </tr>
    <tr class="meta">
      <td>S/N: $($printer['Serial Number'])</td>
      <td class="card-mac">MAC: $($printer['MAC Address'])</td>
    </tr>
  </table>
  <table class="grid">
    <tr class="sec"><td colspan="2">Page Counters</td></tr>
    <tr><td class="lbl">Total Pages</td><td class="val">$(Format-Counter $printer['Total Pages'])</td></tr>
    <tr><td class="lbl">Printer Total</td><td class="val">$(Format-Counter $printer['Total Pages Printer'])</td></tr>
    <tr><td class="lbl">Copier Total</td><td class="val">$(Format-Counter $printer['Total Pages Copier'])</td></tr>

    <tr class="sec"><td colspan="2">Copier Breakdown</td></tr>
    <tr><td class="lbl">Full Color</td><td class="val">$(Format-Counter $printer['Copier Full Color'])</td></tr>
    <tr><td class="lbl">B&amp;W</td><td class="val">$(Format-Counter $printer['Copier B&W'])</td></tr>
    <tr><td class="lbl">Single Color</td><td class="val">$(Format-Counter $printer['Copier Single Color'])</td></tr>
    <tr><td class="lbl">Two-color</td><td class="val">$(Format-Counter $printer['Copier Two-color'])</td></tr>

    <tr class="sec"><td colspan="2">Printer Breakdown</td></tr>
    <tr><td class="lbl">Full Color</td><td class="val">$(Format-Counter $printer['Printer Full Color'])</td></tr>
    <tr><td class="lbl">B&amp;W</td><td class="val">$(Format-Counter $printer['Printer B&W'])</td></tr>
    <tr><td class="lbl">Single Color</td><td class="val">$(Format-Counter $printer['Printer Single Color'])</td></tr>
    <tr><td class="lbl">Two-color</td><td class="val">$(Format-Counter $printer['Printer Two-color'])</td></tr>

    <tr class="sec"><td colspan="2">Toner Levels</td></tr>
    <tr><td class="lbl">Black</td><td class="val">$(Format-TonerBar $printer['Black Toner %'])</td></tr>
    <tr><td class="lbl">Cyan</td><td class="val">$(Format-TonerBar $printer['Cyan Toner %'])</td></tr>
    <tr><td class="lbl">Magenta</td><td class="val">$(Format-TonerBar $printer['Magenta Toner %'])</td></tr>
    <tr><td class="lbl">Yellow</td><td class="val">$(Format-TonerBar $printer['Yellow Toner %'])</td></tr>

    <tr class="sec"><td colspan="2">Status</td></tr>
    <tr><td class="lbl">Error State</td><td class="val">$(Format-ErrorState $printer['Error State'])</td></tr>
  </table>
</div>
"@
}

# -----------------------------------------------------------------------------
# Build-HtmlReport
#
# Renders the array of per-printer result hashtables into a single, modern
# email-friendly HTML document with a dark theme and a 2-column responsive
# grid of printer cards.
#
# Layout uses outer/inner tables for cross-email-client compatibility (Gmail,
# Outlook 365, Apple Mail, the Outlook desktop Word renderer, etc.). The
# 2-column grid is implemented as table rows with two width="50%" cells per
# row; printers are paired off and rendered into those cells. An odd number
# of printers leaves one empty cell on the last row.
# -----------------------------------------------------------------------------
function Build-HtmlReport {
    param([array]$PrintersData)

    $date = Get-Date -Format "dd MMM yyyy - HH:mm"

    # Summary counts for the header chip row.
    $okCount   = @($PrintersData | Where-Object { $_['Status'] -eq 'Verified' }).Count
    $warnCount = @($PrintersData | Where-Object { $_['Status'] -eq 'WrongDevice' }).Count
    $errCount  = @($PrintersData | Where-Object { $_['Status'] -eq 'NoResponse' -or "$($_['Status'])" -like 'Failed after*' }).Count

    # ---- Document head + page wrapper (dark palette) ------------------------
    # Palette (GitHub-dark-inspired):
    #   page bg       #0a0e16    surface bg    #131820    section bg #1a2030
    #   border        #2a3142    text primary  #e6e8eb    text 2nd  #9aa3b2
    #   text muted    #6b7280
    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>RICOH Monitor - $date</title>
<style>
  body { margin:0; padding:0; background:#0a0e16; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif; color:#e6e8eb; -webkit-font-smoothing:antialiased; }
  .wrapper { width:100%; background:#0a0e16; padding:24px 0; }
  .container { max-width:920px; margin:0 auto; }
  .hdr { padding:8px 12px 20px; }
  .hdr-title { font-size:22px; font-weight:600; margin:0; color:#e6e8eb; letter-spacing:-0.3px; }
  .hdr-sub { font-size:13px; color:#9aa3b2; margin-top:4px; }
  .summary { padding:0 12px 16px; font-size:13px; color:#9aa3b2; }
  .chip { display:inline-block; padding:4px 11px; border-radius:11px; font-weight:600; font-size:12px; margin-right:6px; }
  .chip-ok   { background:rgba(34,197,94,0.18);  color:#4ade80; }
  .chip-warn { background:rgba(234,179,8,0.18);  color:#facc15; }
  .chip-err  { background:rgba(239,68,68,0.18);  color:#f87171; }
  .grid-cell { padding:8px; vertical-align:top; }
  .card { background:#131820; border:1px solid #2a3142; border-radius:8px; padding:18px 20px; }
  .card-warn { border-color:rgba(234,179,8,0.45); background:#1a1810; }
  .card-err  { border-color:rgba(239,68,68,0.45); background:#1a1015; }
  .card-hdr { width:100%; border-collapse:collapse; margin-bottom:14px; }
  .card-hdr td { padding:0; vertical-align:baseline; }
  .card-name { font-size:16px; font-weight:600; color:#e6e8eb; }
  .card-ip   { text-align:right; font-size:12px; color:#9aa3b2; font-family:Consolas,'SF Mono',Menlo,monospace; }
  .card-hdr tr.meta td { padding-top:3px; font-size:11px; color:#6b7280; text-transform:uppercase; letter-spacing:0.4px; }
  .card-hdr tr.meta td.card-mac { text-align:right; font-family:Consolas,'SF Mono',Menlo,monospace; text-transform:none; letter-spacing:0; }
  .grid { width:100%; border-collapse:collapse; }
  .grid td { padding:7px 10px; font-size:13px; border-bottom:1px solid #1f2532; }
  .grid td.lbl { color:#9aa3b2; }
  .grid td.val { font-weight:500; text-align:right; color:#e6e8eb; font-variant-numeric:tabular-nums; }
  .grid tr.sec td { background:#1a2030; font-weight:600; color:#9aa3b2; text-transform:uppercase; font-size:10px; letter-spacing:1.2px; padding:7px 10px; border-bottom:1px solid #2a3142; border-top:1px solid #2a3142; }
  .grid tr.sec:first-child td { border-top:none; }
  .status-line { font-size:13px; color:#9aa3b2; margin:8px 0 0; }
  .status-line strong { color:#e6e8eb; }
</style>
</head>
<body>
<table class="wrapper" cellpadding="0" cellspacing="0" border="0" width="100%">
<tr><td align="center">
<table class="container" cellpadding="0" cellspacing="0" border="0" width="920">
  <tr><td class="hdr">
    <div class="hdr-title">RICOH Monitor</div>
    <div class="hdr-sub">$date</div>
  </td></tr>
  <tr><td class="summary">
    <span class="chip chip-ok">$okCount OK</span><span class="chip chip-warn">$warnCount Wrong device</span><span class="chip chip-err">$errCount Unreachable</span>
  </td></tr>
  <tr><td>
    <table class="grid-outer" cellpadding="0" cellspacing="0" border="0" width="100%">
"@

    # ---- 2-column grid of printer cards -------------------------------------
    # Iterate in pairs; each row has two width="50%" cells. Odd count leaves
    # the last right cell empty.
    for ($i = 0; $i -lt $PrintersData.Count; $i += 2) {
        $left  = $PrintersData[$i]
        $right = if ($i + 1 -lt $PrintersData.Count) { $PrintersData[$i + 1] } else { $null }

        $leftHtml  = Format-PrinterCard $left
        $rightHtml = if ($null -ne $right) { Format-PrinterCard $right } else { "&nbsp;" }

        $html += @"
      <tr>
        <td class="grid-cell" width="50%">$leftHtml</td>
        <td class="grid-cell" width="50%">$rightHtml</td>
      </tr>
"@
    }

    $html += @"
    </table>
  </td></tr>
</table>
</td></tr>
</table>
</body>
</html>
"@
    return $html
}

# -----------------------------------------------------------------------------
# Send-Report
#
# Dispatches the rendered HTML report.
#   - The HTML is ALWAYS written to report_yyyy-MM-dd_HHmm.html in the working
#     directory, regardless of whether email sending is enabled. The on-disk
#     report is the canonical output; email is an additional channel.
#   - When SendEmail = $true and a valid Smtp section is in the config, the
#     same HTML is also emailed via Send-MailMessage.
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Invoke-SmtpTest
#
# Loads Ricoh-Monitor.json and sends a small test email via Send-MailMessage
# using the Smtp section. Prints every configured field before the send so
# typos are obvious in the output. Used by -TestSmtp.
# -----------------------------------------------------------------------------
function Invoke-SmtpTest {
    $config = Get-PrintersConfig
    $smtp   = $config.Smtp

    if (-not $smtp -or [string]::IsNullOrWhiteSpace("$($smtp.Server)")) {
        Write-Host "Smtp section missing or empty in Ricoh-Monitor.json." -ForegroundColor Red
        Write-Host "Add Smtp.{Server, Port, Username, Password, From, To} and re-run." -ForegroundColor Red
        return
    }

    $date = Get-Date -Format "dd-MM-yyyy HH:mm"
    Write-Host "=== SMTP Test ===" -ForegroundColor Cyan
    Write-Host ("  Server:   {0}:{1}" -f $smtp.Server, $smtp.Port)
    Write-Host ("  Username: {0}" -f $smtp.Username)
    Write-Host ("  From:     {0}" -f $smtp.From)
    Write-Host ("  To:       {0}" -f (@($smtp.To) -join ', '))
    Write-Host ""

    $subject = "RICOH Monitor SMTP test - $date"
    $body    = @"
<html><body style="font-family:Arial,sans-serif;color:#2d3748;">
<h2 style="color:#1a202c;">RICOH Monitor - SMTP test</h2>
<p>This is a test message generated by <code>Ricoh-Monitor.ps1 -TestSmtp</code>.</p>
<p>If you can see this email, the SMTP section of <code>Ricoh-Monitor.json</code> is working correctly.</p>
<p style="color:#718096;font-size:12px;">Sent at $date</p>
</body></html>
"@

    # Force TLS 1.2 - PS 5.1's Send-MailMessage otherwise negotiates TLS 1.0.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $credential = New-Object System.Management.Automation.PSCredential (
        $smtp.Username,
        (ConvertTo-SecureString "$($smtp.Password)" -AsPlainText -Force)
    )

    try {
        Send-MailMessage -From $smtp.From `
                        -To $smtp.To `
                        -Subject $subject `
                        -Body $body `
                        -BodyAsHtml `
                        -SmtpServer $smtp.Server `
                        -Port ([int]$smtp.Port) `
                        -UseSsl `
                        -Credential $credential `
                        -ErrorAction Stop `
                        -WarningAction SilentlyContinue
        Write-Host "Test email sent successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Test email failed: $_" -ForegroundColor Red
    }
}

function Send-Report {
    param(
        [array]$PrintersData,
        [bool]$SendEmail,
        [object]$SmtpConfig
    )

    $date    = Get-Date -Format "dd-MM-yyyy HH:mm"
    $subject = "RICOH Monitor - $date"
    $html    = Build-HtmlReport -PrintersData $PrintersData

    # --- Always write the report to disk ------------------------------------
    $reportPath = "report_" + (Get-Date -Format "yyyy-MM-dd_HHmm") + ".html"
    $html | Out-File -FilePath $reportPath -Encoding utf8
    Write-Host "Report written to: $reportPath" -ForegroundColor Green

    # --- Optional: also email it --------------------------------------------
    if (-not $SendEmail) {
        Write-Host "Email sending disabled in config; skipping email." -ForegroundColor Yellow
        return
    }

    # Validate the SMTP block before attempting a send so a missing/empty
    # config produces a clear message instead of a Send-MailMessage error.
    if (-not $SmtpConfig -or [string]::IsNullOrWhiteSpace("$($SmtpConfig.Server)")) {
        Write-Host "SendEmail is true but the 'Smtp' section in Ricoh-Monitor.json is missing or empty." -ForegroundColor Red
        Write-Host "Add Smtp.{Server, Port, Username, Password, From, To} and re-run." -ForegroundColor Red
        return
    }

    # Force TLS 1.2 before the STARTTLS handshake - PS 5.1's Send-MailMessage
    # otherwise negotiates TLS 1.0, which modern Exchange rejects.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $credential = New-Object System.Management.Automation.PSCredential (
        $SmtpConfig.Username,
        (ConvertTo-SecureString "$($SmtpConfig.Password)" -AsPlainText -Force)
    )

    try {
        # -UseSsl on SMTP means "issue STARTTLS and encrypt the session".
        # Correct flag for Exchange on port 587 despite the misleading name.
        # -ErrorAction Stop turns the cmdlet's non-terminating errors
        # (DNS failure, auth failure, etc.) into terminating ones so the
        # catch block below actually fires instead of letting execution
        # fall through to the success message.
        Send-MailMessage -From $SmtpConfig.From `
                        -To $SmtpConfig.To `
                        -Subject $subject `
                        -Body $html `
                        -BodyAsHtml `
                        -SmtpServer $SmtpConfig.Server `
                        -Port ([int]$SmtpConfig.Port) `
                        -UseSsl `
                        -Credential $credential `
                        -ErrorAction Stop `
                        -WarningAction SilentlyContinue
        Write-Host "Email sent successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to send email: $_" -ForegroundColor Red
    }
}

# =============================================================================
# DISCOVERY HELPERS
# Used only when the script is launched with -Discover. Sweeps a CIDR range,
# classifies each IP via the same Manufacturer/Model SNMP probe used by the
# verification phase, and offers to merge any Ricoh hits into the config file.
# =============================================================================

# -----------------------------------------------------------------------------
# Get-LocalSubnetCidr
#
# Returns the local NIC's subnet in CIDR form (e.g. "192.168.2.0/24"), picked
# from whichever IPv4 interface is "Up" and has a default gateway. Returns
# $null if no such NIC is found (then discovery falls back to a manual prompt).
# -----------------------------------------------------------------------------
function Get-LocalSubnetCidr {
    try {
        $cfg = Get-NetIPConfiguration -ErrorAction Stop |
            Where-Object { $null -ne $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } |
            Select-Object -First 1
    }
    catch {
        return $null
    }
    if (-not $cfg -or -not $cfg.IPv4Address) { return $null }

    $ip     = $cfg.IPv4Address.IPAddress
    $prefix = [int]$cfg.IPv4Address.PrefixLength

    # Mask the host bits off to get the network base address.
    $ipInt    = ConvertTo-IpInt $ip
    $hostBits = 32 - $prefix
    $mask     = if ($hostBits -ge 32) { [UInt32]0 } else { [UInt32]([UInt32]::MaxValue) -shl $hostBits }
    $netInt   = [UInt32]($ipInt -band $mask)

    return "$(ConvertTo-IpString $netInt)/$prefix"
}

# IPv4 string <-> UInt32 helpers (used by Get-LocalSubnetCidr and Expand-Cidr).
function ConvertTo-IpInt {
    param([string]$Ip)
    $bytes = ([System.Net.IPAddress]::Parse($Ip)).GetAddressBytes()
    [Array]::Reverse($bytes)
    return [BitConverter]::ToUInt32($bytes, 0)
}
function ConvertTo-IpString {
    param([UInt32]$IpInt)
    $bytes = [BitConverter]::GetBytes($IpInt)
    [Array]::Reverse($bytes)
    return ([System.Net.IPAddress]::new($bytes)).ToString()
}

# -----------------------------------------------------------------------------
# Expand-Cidr
#
# Returns the list of usable host IPs in a CIDR block, excluding network and
# broadcast addresses. Refuses anything wider than /20 (~4094 hosts) to avoid
# accidentally fanning out a discovery scan across thousands of addresses.
# -----------------------------------------------------------------------------
function Expand-Cidr {
    param([string]$Cidr)

    if ($Cidr -notmatch '^(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/(\d{1,2})$') {
        throw "Invalid CIDR format: '$Cidr'. Expected e.g. 192.168.1.0/24"
    }
    $baseIp = $Matches[1]
    $prefix = [int]$Matches[2]

    if ($prefix -lt 20 -or $prefix -gt 30) {
        throw "Discovery supports /20 to /30 (got /$prefix). Narrow the range."
    }

    $hostBits  = 32 - $prefix
    $hostCount = [int]([Math]::Pow(2, $hostBits)) - 2

    $baseInt = ConvertTo-IpInt $baseIp
    $mask    = if ($hostBits -ge 32) { [UInt32]0 } else { [UInt32]([UInt32]::MaxValue) -shl $hostBits }
    $netInt  = [UInt32]($baseInt -band $mask)

    return 1..$hostCount | ForEach-Object { ConvertTo-IpString ([UInt32]($netInt + $_)) }
}

# -----------------------------------------------------------------------------
# Invoke-ParallelProbe
#
# Probes a list of IPs in parallel using a runspace pool (works on PS 5.1+,
# no module dependency). Each runspace creates its own OlePrn.OleSNMP COM
# object and runs the same Manufacturer/Model probe Test-RicohPrinter uses.
# Returns one result object per IP with Status = Verified | WrongDevice | NoResponse.
# -----------------------------------------------------------------------------
function Invoke-ParallelProbe {
    param(
        [string[]]$Ips,
        [string]$SnmpCommunity  = "public",
        [int]$Concurrency   = 50,
        [int]$TimeoutMs     = 1000
    )

    # Self-contained probe - runs inside each runspace so it can't reference
    # the parent script's variables. Mirrors Test-RicohPrinter's logic.
    $probe = {
        param($IP, $SnmpCommunity, $TimeoutMs)
        $r = [pscustomobject]@{ IP = $IP; Status = 'NoResponse'; Manufacturer = $null; Model = $null }
        try {
            $snmp = New-Object -ComObject "OlePrn.OleSNMP"
            $snmp.Open($IP, $SnmpCommunity, 1, $TimeoutMs)
            try {
                $r.Manufacturer = $snmp.Get(".1.3.6.1.4.1.367.3.2.1.1.1.7.0")
                $r.Model        = $snmp.Get(".1.3.6.1.4.1.367.3.2.1.1.1.1.0")
            } catch {
                $snmp.Close()
                return $r
            }
            $snmp.Close()
            if ($r.Manufacturer -and "$($r.Manufacturer)" -match "RICOH") {
                $r.Status = 'Verified'
            } else {
                $r.Status = 'WrongDevice'
            }
        } catch {
            # Any failure opening the COM object or talking SNMP = NoResponse.
            # The $r object already has Status = 'NoResponse' from its
            # initialization, so we deliberately swallow the exception here.
            Write-Debug "Probe of $IP failed: $_"
        }
        return $r
    }

    $pool = [RunspaceFactory]::CreateRunspacePool(1, $Concurrency)
    $pool.Open()

    # Kick off all probes asynchronously.
    $jobs = @()
    foreach ($ip in $Ips) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($probe).AddArgument($ip).AddArgument($SnmpCommunity).AddArgument($TimeoutMs)
        $jobs += [pscustomobject]@{ PS = $ps; Handle = $ps.BeginInvoke() }
    }

    # Collect results, with a progress bar so the user knows how far along we are.
    $results = New-Object System.Collections.ArrayList
    $total   = $jobs.Count
    $done    = 0
    foreach ($job in $jobs) {
        $r = $job.PS.EndInvoke($job.Handle)
        $job.PS.Dispose()
        [void]$results.Add($r)
        $done++
        Write-Progress -Activity "Probing subnet" -Status "$done / $total" -PercentComplete (($done / $total) * 100)
    }
    Write-Progress -Activity "Probing subnet" -Completed

    $pool.Close(); $pool.Dispose()
    return $results.ToArray()
}

# -----------------------------------------------------------------------------
# Find-RicohPrintersInCidr
#
# Scans a single CIDR range and returns only the Verified Ricoh hits. Prints
# a one-line summary (elapsed time, hits, other SNMP devices) so the user
# knows whether the network was alive at all.
# -----------------------------------------------------------------------------
function Find-RicohPrintersInCidr {
    param(
        [string]$Cidr,
        [string]$SnmpCommunity = "public"
    )

    $ips = Expand-Cidr $Cidr
    Write-Host "Probing $Cidr ($($ips.Count) addresses)..." -ForegroundColor Cyan
    $start   = Get-Date
    $results = Invoke-ParallelProbe -Ips $ips -SnmpCommunity $SnmpCommunity
    $elapsed = ((Get-Date) - $start).TotalSeconds

    $verified = @($results | Where-Object { $_.Status -eq 'Verified' })
    $wrong    = @($results | Where-Object { $_.Status -eq 'WrongDevice' })
    Write-Host ("Done in {0:N1}s. Found: {1} Ricoh, {2} other SNMP device(s)" -f $elapsed, $verified.Count, $wrong.Count) -ForegroundColor Green
    return $verified
}

# -----------------------------------------------------------------------------
# Save-PrintersToConfig
#
# Merges the given printers into Ricoh-Monitor.json, preserving existing
# entries and the SendEmail flag. IPs already in the config are skipped
# (no duplicates). Creates the file if missing. Returns the count of
# printers actually added (excluding duplicates).
#
# Used by both Add-DiscoveredToConfig (interactive, with confirmation) and
# Invoke-FirstRunDiscovery (auto-add, no prompt).
# -----------------------------------------------------------------------------
# Returns a placeholder Smtp block used when generating a fresh config.
# Fields are example values to show the operator what to fill in.
function New-DefaultSmtpConfig {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Pure factory returning a hashtable; changes no external state.')]
    param()
    return [ordered]@{
        Server   = "smtp.example.com"
        Port     = 587
        Username = "user@example.com"
        Password = ""
        From     = "user@example.com"
        To       = @("recipient1@example.com", "recipient2@example.com")
    }
}

function Save-PrintersToConfig {
    param(
        [array]$Printers,
        [string]$ConfigPath = "Ricoh-Monitor.json"
    )

    if (Test-Path $ConfigPath) {
        $existing         = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        $sendEmail        = if ($null -ne $existing.SendEmail) { [bool]$existing.SendEmail } else { $false }
        $smtp             = if ($null -ne $existing.Smtp)      { $existing.Smtp }             else { New-DefaultSmtpConfig }
        $existingPrinters = if ($existing.Printers)            { @($existing.Printers) }      else { @() }
    } else {
        $sendEmail        = $false
        $smtp             = New-DefaultSmtpConfig
        $existingPrinters = @()
    }

    $existingIps = @($existingPrinters | ForEach-Object { $_.IP })
    $merged      = @($existingPrinters)
    $added       = 0
    foreach ($p in $Printers) {
        if ($existingIps -notcontains $p.IP) {
            $merged += [ordered]@{ IP = $p.IP; SnmpCommunity = "public" }
            $added++
        }
    }

    $output = [ordered]@{
        SendEmail = $sendEmail
        Smtp      = $smtp
        Printers  = $merged
    }
    $output | ConvertTo-Json -Depth 4 | Out-File -FilePath $ConfigPath -Encoding utf8
    return $added
}

# -----------------------------------------------------------------------------
# Add-DiscoveredToConfig
#
# Interactive variant: prompts y/N before persisting. Used by the -Discover
# workflow when re-scanning an already-configured deployment. Wraps
# Save-PrintersToConfig.
# -----------------------------------------------------------------------------
function Add-DiscoveredToConfig {
    param(
        [array]$Found,
        [string]$ConfigPath = "Ricoh-Monitor.json"
    )

    $confirm = Read-Host "Add these to $ConfigPath ? (y/N)"
    if ($confirm -ne 'y' -and $confirm -ne 'Y') {
        Write-Host "Skipped. No changes made." -ForegroundColor Yellow
        return
    }

    $added = Save-PrintersToConfig -Printers $Found -ConfigPath $ConfigPath
    Write-Host "Added $added printer(s) to $ConfigPath." -ForegroundColor Green
    if ($added -lt $Found.Count) {
        Write-Host "($($Found.Count - $added) already in the config - skipped.)" -ForegroundColor DarkGray
    }
}

# -----------------------------------------------------------------------------
# Invoke-FirstRunDiscovery
#
# Runs automatically on the very first launch (when Ricoh-Monitor.json
# doesn't exist). Differences from interactive -Discover mode:
#   - Auto-adds discovered printers without confirmation.
#   - Loops on the prompt-for-CIDR step until either printers are found or
#     the user gives up (blank input).
#   - Always writes a config (empty if user gave up), so subsequent runs
#     don't re-trigger this first-run flow.
# -----------------------------------------------------------------------------
function Invoke-FirstRunDiscovery {
    param([string]$ConfigPath = "Ricoh-Monitor.json")

    Write-Host "=== First-run setup ===" -ForegroundColor Cyan
    Write-Host "No config file found. Discovering Ricoh printers on the network..."
    Write-Host ""

    $allFound = @()

    # Step 1: scan the local NIC subnet automatically.
    $localCidr = Get-LocalSubnetCidr
    if ($localCidr) {
        Write-Host "Local NIC subnet: $localCidr"
        $allFound += @(Find-RicohPrintersInCidr -Cidr $localCidr)
    } else {
        Write-Host "Could not auto-detect local subnet from NIC config." -ForegroundColor Yellow
    }

    # Step 2: keep prompting for additional subnets until something is found
    # or the user gives up by pressing Enter.
    while ($allFound.Count -eq 0) {
        Write-Host ""
        Write-Host "No Ricoh printers found." -ForegroundColor Yellow
        $extra = Read-Host "Enter another subnet to scan (CIDR e.g. 192.168.1.0/24), or blank to skip"
        if ([string]::IsNullOrWhiteSpace($extra)) { break }
        try {
            $allFound += @(Find-RicohPrintersInCidr -Cidr $extra)
        } catch {
            Write-Host "Error: $_" -ForegroundColor Red
        }
    }

    # Step 3: persist results - either the discovered printers or an empty
    # config so we don't re-run discovery on the next launch.
    Write-Host ""
    if ($allFound.Count -gt 0) {
        Write-Host "Discovered Ricoh printers:" -ForegroundColor Green
        foreach ($p in $allFound) {
            Write-Host ("  [OK] {0,-15} - {1} {2}" -f $p.IP, $p.Manufacturer, $p.Model) -ForegroundColor Green
        }
        $added = Save-PrintersToConfig -Printers $allFound -ConfigPath $ConfigPath
        Write-Host ""
        Write-Host "Wrote $added printer(s) to $ConfigPath (SendEmail = false)." -ForegroundColor Green
    } else {
        $emptyConfig = [ordered]@{
            SendEmail = $false
            Smtp      = New-DefaultSmtpConfig
            Printers  = @()
        }
        $emptyConfig | ConvertTo-Json -Depth 4 | Out-File -FilePath $ConfigPath -Encoding utf8
        Write-Host "Wrote empty config to $ConfigPath. Edit it manually to add printer IPs, then re-run." -ForegroundColor Yellow
    }
    Write-Host ""
}

# -----------------------------------------------------------------------------
# Get-KnownSubnet
#
# Derives the unique /24 subnets covered by the printers currently in the
# config. Used by Invoke-NewPrinterScan to know what to sweep. Grouping is
# done on the first three octets of each IP - cheap, no SNMP, no mask math,
# and correct for every "/24 printer VLAN" layout small offices use.
# -----------------------------------------------------------------------------
function Get-KnownSubnet {
    param([array]$Printers)
    $subnets = @()
    foreach ($p in $Printers) {
        if ("$($p.IP)" -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.\d{1,3}$') {
            $cidr = "$($Matches[1]).0/24"
            if ($subnets -notcontains $cidr) { $subnets += $cidr }
        }
    }
    return $subnets
}

# -----------------------------------------------------------------------------
# Invoke-NewPrinterScan
#
# Runs as a normal-flow step between config-load and verification. Sweeps
# every /24 that the existing printers occupy, auto-adds any freshly
# discovered Ricoh printers to the config (skipping already-known IPs),
# and returns the list of new ones so main can refresh its printer list.
# Prints a short summary either way.
# -----------------------------------------------------------------------------
function Invoke-NewPrinterScan {
    param(
        [array]$Printers,
        [string]$ConfigPath = "Ricoh-Monitor.json"
    )

    $subnets = Get-KnownSubnet -Printers $Printers
    if ($subnets.Count -eq 0) {
        return @()
    }

    Write-Host "=== Scanning known subnets for new printers ===" -ForegroundColor Cyan
    $knownIps   = @($Printers | ForEach-Object { $_.IP })
    $newlyFound = @()

    foreach ($cidr in $subnets) {
        try {
            $hits = Find-RicohPrintersInCidr -Cidr $cidr
            foreach ($h in $hits) {
                if ($knownIps -notcontains $h.IP) {
                    $newlyFound += $h
                }
            }
        } catch {
            Write-Host "Error scanning $cidr : $_" -ForegroundColor Red
        }
    }

    if ($newlyFound.Count -eq 0) {
        Write-Host "No new printers found." -ForegroundColor DarkGray
        Write-Host ""
        return @()
    }

    Write-Host ""
    Write-Host "Found $($newlyFound.Count) new printer(s):" -ForegroundColor Green
    foreach ($p in $newlyFound) {
        Write-Host ("  [NEW] {0,-15} - {1} {2}" -f $p.IP, $p.Manufacturer, $p.Model) -ForegroundColor Green
    }
    $added = Save-PrintersToConfig -Printers $newlyFound -ConfigPath $ConfigPath
    Write-Host "Added $added new printer(s) to $ConfigPath." -ForegroundColor Green
    Write-Host ""
    return $newlyFound
}

# -----------------------------------------------------------------------------
# Invoke-Discovery
#
# Interactive discovery (triggered by -Discover). Differs from first-run:
#   1. Detect the local NIC's subnet and scan it.
#   2. Always prompt for additional CIDRs (even if hits found) so multi-VLAN
#      setups can be scanned in one pass.
#   3. Show every discovered Ricoh printer and ask y/N before merging into
#      the config.
# -----------------------------------------------------------------------------
function Invoke-Discovery {
    Write-Host "=== Ricoh Printer Discovery ===" -ForegroundColor Cyan
    Write-Host ""

    $allFound = @()

    # --- Step 1: try the local NIC's subnet ----------------------------------
    $localCidr = Get-LocalSubnetCidr
    if ($localCidr) {
        Write-Host "Local NIC subnet: $localCidr"
        $hits = Find-RicohPrintersInCidr -Cidr $localCidr
        $allFound += $hits
    } else {
        Write-Host "Could not auto-detect local subnet from NIC config." -ForegroundColor Yellow
    }

    # --- Step 2: prompt for additional subnets if needed ---------------------
    # Loop until user enters blank. They can keep adding subnets even after
    # finding hits, in case they have multiple printer VLANs.
    if ($allFound.Count -eq 0) {
        Write-Host ""
        Write-Host "No Ricoh printers found on the local subnet." -ForegroundColor Yellow
    }

    while ($true) {
        Write-Host ""
        $extra = Read-Host "Scan another subnet? Enter CIDR (e.g. 192.168.1.0/24) or blank to stop"
        if ([string]::IsNullOrWhiteSpace($extra)) { break }
        try {
            $hits = Find-RicohPrintersInCidr -Cidr $extra
            $allFound += $hits
        } catch {
            Write-Host "Error: $_" -ForegroundColor Red
        }
    }

    # --- Step 3: present results and offer to merge into config -------------
    Write-Host ""
    if ($allFound.Count -eq 0) {
        Write-Host "Discovery finished. No Ricoh printers found." -ForegroundColor Yellow
        return
    }

    Write-Host "Discovered Ricoh printers:" -ForegroundColor Green
    foreach ($p in $allFound) {
        Write-Host ("  [OK] {0,-15} - {1} {2}" -f $p.IP, $p.Manufacturer, $p.Model) -ForegroundColor Green
    }
    Write-Host ""

    Add-DiscoveredToConfig -Found $allFound
}

# =============================================================================
# Main Execution
#
# Two modes:
#   - With -Discover  : run subnet discovery, optionally merge results into config.
#   - Without         : normal monitoring run (verify -> collect -> report).
#
# Normal-mode pipeline:
#   1. Load (or generate) the config file.
#   2. Verification phase - probe every IP, classify each, print a summary line.
#   3. Counter collection - full SNMP sweep on Verified printers; non-verified
#      entries become stub records so they still appear in the final report.
#   4. Build the HTML report and either email it or write it to disk.
# =============================================================================
try {
    if ($Discover) {
        Invoke-Discovery
        return
    }

    if ($TestSnmp) {
        Invoke-RawSnmpDiagnostic -IP $TestSnmp
        return
    }

    if ($TestSmtp) {
        Invoke-SmtpTest
        return
    }

    Write-Host "Starting printer monitoring..." -ForegroundColor Cyan

    # --- Step 0: first-run bootstrap -----------------------------------------
    # If no config exists, auto-discover printers on the local subnet and
    # write them to Ricoh-Monitor.json before we proceed. The function
    # always leaves a config file on disk (empty if nothing was found and
    # the user gave up), so subsequent runs go straight to Step 1.
    if (-not (Test-Path "Ricoh-Monitor.json")) {
        Invoke-FirstRunDiscovery
    }

    # --- Step 1: load config -------------------------------------------------
    $config    = Get-PrintersConfig
    $sendEmail = [bool]$config.SendEmail
    $printers  = @($config.Printers)
    Write-Host "Loaded configuration for $($printers.Count) printers (SendEmail: $sendEmail)"
    Write-Host ""

    # --- Step 1.5: scan known subnets for newly-installed printers ----------
    # Sweeps every /24 covered by the existing config. New Ricoh hits are
    # auto-added to Ricoh-Monitor.json so the next verify+collect pass
    # includes them in this month's report.
    $newPrinters = Invoke-NewPrinterScan -Printers $printers
    if ($newPrinters.Count -gt 0) {
        # Reload config so the newly-added printers are in $printers.
        $config   = Get-PrintersConfig
        $printers = @($config.Printers)
    }

    # --- Step 2: verification phase ------------------------------------------
    # Fast identity probe on each IP. Output one summary line per printer so
    # the operator can immediately see which targets are good, wrong, or down.
    Write-Host "=== Verification phase ===" -ForegroundColor Cyan

    $verifications = @()
    foreach ($printer in $printers) {
        $v = Test-RicohPrinter -IP $printer.IP -SnmpCommunity $printer.SnmpCommunity
        switch ($v.Status) {
            "Verified" {
                Write-Host ("[OK]           {0,-15} - {1} {2}" -f $v.IP, $v.Manufacturer, $v.Model) -ForegroundColor Green
            }
            "WrongDevice" {
                $mfg   = if ($v.Manufacturer) { $v.Manufacturer } else { "<empty>" }
                $model = if ($v.Model)        { $v.Model }        else { "<empty>" }
                Write-Host ("[WRONG DEVICE] {0,-15} - Manufacturer: {1}, Model: {2}" -f $v.IP, $mfg, $model) -ForegroundColor Yellow
            }
            "NoResponse" {
                Write-Host ("[NO SNMP]      {0,-15} - {1}" -f $v.IP, $v.Error) -ForegroundColor Red
            }
        }
        # Keep the original printer-config entry alongside the verification
        # result so step 3 has both the connection details and the identity.
        $verifications += [pscustomobject]@{ Printer = $printer; Verification = $v }
    }

    Write-Host ""

    # --- Step 3: counter collection ------------------------------------------
    # Full SNMP sweep only on Verified printers. Non-verified printers still
    # become entries in $allPrintersData (with Status = WrongDevice/NoResponse)
    # so they show up in the final report and aren't silently dropped.
    Write-Host "=== Counter collection ===" -ForegroundColor Cyan

    $allPrintersData = @()
    foreach ($entry in $verifications) {
        if ($entry.Verification.Status -eq "Verified") {
            Write-Host "Collecting from $($entry.Printer.IP)..."
            $data = Get-SnmpData -Verification $entry.Verification -SnmpCommunity $entry.Printer.SnmpCommunity
            if ($data["Status"] -like "Failed after*") {
                Write-Host "  $($data['Status'])" -ForegroundColor Yellow
            } else {
                Write-Host "  OK" -ForegroundColor Green
            }
            $allPrintersData += $data
        } else {
            # Stub record for WrongDevice / NoResponse - Build-HtmlReport
            # branches on the Status field to render these differently.
            $stub = [ordered]@{
                "IP"           = $entry.Printer.IP
                "Status"       = $entry.Verification.Status
                "Error"        = $entry.Verification.Error
                "Manufacturer" = $entry.Verification.Manufacturer
                "Model"        = $entry.Verification.Model
            }
            $allPrintersData += $stub
        }
    }

    Write-Host ""

    # --- Step 4: dispatch report ---------------------------------------------
    Send-Report -PrintersData $allPrintersData -SendEmail $sendEmail -SmtpConfig $config.Smtp
    Write-Host "Done." -ForegroundColor Green
}
catch {
    Write-Host "Error in main execution: $_" -ForegroundColor Red
}
