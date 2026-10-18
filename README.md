# CPE TR-069 / TR-181 / TR-143 Automated Validation Testbed

## Overview and Purpose

This project provides an **automated laboratory testbed and validation suite** designed to test, audit, and certify **Customer Premises Equipment (CPE)** devices—specifically 5G/LTE Fixed Wireless Access (FWA) and broadband routers (such as the **TCL HH516L 5G CPE**)—against operator **Test Object Lists (TOL)** and Broadband Forum standards:

- **TR-069 (CWMP)**: CPE WAN Management Protocol session handling, `Inform` / `InformResponse` lifecycle, `PeriodicInformInterval` timing compliance, `GetParameterValues` (GPV), `SetParameterValues` (SPV), and `GetParameterAttributes` (GPA).
- **TR-181 Issue 2 (`Device:2`)**: Full data model discovery, structural inspection, read/write capability verification, interface statistics consistency, and security/redaction auditing.
- **TR-143 & Network Diagnostics**: Active throughput and reachability diagnostics (`DownloadDiagnostics` and `IPPing`) triggered via CWMP against a controlled local test server.
- **Operator TOL Compliance**: Automated execution and reporting of the 78 test cases defined in the reference operator matrix (`TOL _Huawei_H173-383.xlsx`), producing ready-to-deliver **Excel workbooks**, **JSON dumps**, **packet-level PCAP evidence**, and an **executive PowerPoint audit report**.

---

## Technology Stack & Architecture

The testbed combines a containerized **Kathará / Docker** network environment with a single-file Python orchestration and validation agent (`tr069_validator_agent.py`):

1. **Kathará & Docker Emulated Infrastructure (`lab.conf`)**:
   - **ACS Node (`acs` — `Dockerfile.genieacs`, `acs.startup`)**: Runs **GenieACS v1.2.9** on Node.js 20 (`local/genieacs:latest`), dual-homed on the internal management network (`mgmt_net`, `192.168.100.1/24`) and bridged to the physical CPE network interface (`cpe_net` bridged to `enp3s0`, `192.168.111.254/24`). Exposes:
     - `7557/tcp`: CWMP endpoint for CPE communication (`genieacs-cwmp`).
     - `7547/tcp`: Northbound REST API (`genieacs-nbi`) used by the validator agent.
     - `3000/tcp`: GenieACS Web UI (`genieacs-ui`).
   - **Database Node (`mongo` — `Dockerfile.mongo`, `mongo.startup`)**: Runs **MongoDB 5.0** (`local/mongo:5.0`) on `192.168.100.2:27017`, storing device trees, CWMP tasks, presets, provisions, and faults.
2. **Packet Capture & Protocol Dissection (`tcpdump` + `tshark`)**:
   - Uses an ephemeral Docker helper container (instantiated from the already-installed Kathará lab image on `--net=host`) to run `tcpdump` on the CWMP port without disrupting the running lab.
   - Uses **Wireshark CLI (`tshark`)** on the host to dissect HTTP/SOAP XML payloads, isolating `2 PERIODIC` `Inform` requests and matching HTTP `200 OK` `InformResponse` packets.
3. **TR-143 HTTP Payload Server (`tr143_server/`)**:
   - Serves a deterministic 50 MiB binary payload (`tr143_server/test_50MB.bin`) over HTTP (`http://192.168.1.219:8080/test_50MB.bin` by default) to validate `Device.IP.Diagnostics.DownloadDiagnostics`.
4. **Python Validator Agent (`tr069_validator_agent.py`)**:
   - Self-contained orchestrator running in the local virtual environment (`venv/`).
   - Uses Python standard library modules (`urllib.request`, `subprocess`, `xml.etree.ElementTree`, `fcntl`, `hashlib`, `unittest`) alongside `openpyxl` (`3.1.5`) for Excel generation and `python-pptx` (`1.0.2`) for automated PowerPoint report compilation.

---

## Technical Workflow: What the Validator Does

A standard execution of `tr069_validator_agent.py` performs an end-to-end validation pipeline divided into six technical phases:

### 1. Concurrency Control, State Protection & Safe Rollback
- Acquires an exclusive non-blocking file lock (`logs/tol_tcl.lock` via `fcntl.flock`) to prevent overlapping test runs against the same CPE.
- Registers `SIGINT` and `SIGTERM` signal handlers and tracks every modified parameter's original value and XSD type in `ACS.protection`.
- Guarantees in `finally` blocks that any temporary parameter change is restored via SPV + GPV verification, stale ACS tasks/faults created during the run are cleaned up, and temporary capture containers are removed. If a parameter restoration ever fails (`restore_failed`), all subsequent active writes are immediately halted (`BLOCCATO`).

### 2. GenieACS NBI & Container Orchestration (`ACS` Class)
- Queries the GenieACS NBI REST API (`/devices`, `/tasks`, `/faults`, `/presets`, `/provisions`) and inspects `/var/log/cwmp.log` inside the ACS container.
- Triggers a full `refreshObject` on the `Device` root tree so the cached data model reflects live CPE state.
- Executes targeted `getParameterValues` (GPV), `setParameterValues` (SPV), and provision-driven `getParameterAttributes` (GPA) tasks, validating that each parameter's `_timestamp` (or `_attributesTimestamp`) is strictly newer than the task dispatch time and that no CWMP fault codes (`CWMPFault`) were raised.
- Inspects active GenieACS presets and provisions (`interval_policy`) to detect whether the ACS enforces a fixed `PeriodicInformInterval` (e.g., 30 seconds) for the CPE's serial number.

### 3. TOL Test Suite Execution (`TOLSuite` Class — 78 Test Cases)
Evaluates all 78 rows of the source operator matrix (`TOL _Huawei_H173-383.xlsx`):
- **Read-Only Audits (Default Mode)**:
  - **Device Identity & Versions**: Validates `Manufacturer`, `OUI`, `ProductClass`, `SerialNumber`, `HardwareVersion`, `SoftwareVersion`, `ModemFirmwareVersion`, and `UpTime`.
  - **IMEI Validation**: Reads both `Device.DeviceInfo.IMEI` and `Device.Cellular.Interface.1.IMEI`, checking 15-digit numeric format compliance and cross-node consistency.
  - **Cellular, APN & WAN Mapping**: Audits `Device.Cellular.Interface.*`, USIM fields (with PIN redaction), and `Device.Cellular.AccessPoint.*` (Data/VoIP profiles), correlating `Device.IP.Interface.*` via `LowerLayers`.
  - **LAN, DHCPv4 & Wi-Fi**: Reads `Device.Ethernet.*`, `Device.DHCPv4.Server.*`, and `Device.WiFi.*` (radios, SSIDs, channels, bandwidth, utilization, associated devices, and security modes).
  - **Interface Traffic Counters**: Performs two distinct GPV samples across `Device.Cellular.Interface.*` and `Device.Ethernet.Interface.*` (`BytesSent`, `BytesReceived`, `PacketsSent`, `PacketsReceived`, `ErrorsSent`, `ErrorsReceived`), parsing 64-bit `unsignedLong` values without precision loss to verify counter increments and detect unexpected resets/rollovers.
  - **Parameter Attributes (GPA)**: Triggers a live provision refresh of notification attributes on `SoftwareVersion`, `ProvisioningCode`, `ConnectionRequestURL`, and the active cellular IPv4 address, verifying `Notification=2` (Active Notification).
  - **Out-of-Scope / Protected Cases**: NTP remains strictly read-only; IMS (`CPE_0046`) remains `NON ESEGUITO`; wildcard queries without explicit `InstanceWildcardsSupported` capability and unconfigured upload/UDP echo endpoints are marked `BLOCCATO` with detailed technical rationale.
- **Reversible Active Tests (`--tol-active` Mode)**:
  - **Wi-Fi SSID Advertisement Toggle (Row 57)**: Toggles writable `SSIDAdvertisementEnabled` booleans via SPV, verifies the change via GPV, and restores the original state (executed only when wired management is confirmed and no Wi-Fi clients are associated).
  - **Wi-Fi WPA2/WPA3 Security Mode Switching (Row 59)**: Switches `Security.ModeEnabled` among supported modes (`WPA3-Personal`, `WPA2-WPA3-Personal`, `WPA2-Personal`), verifies via GPV, and restores the initial mode without modifying passphrases.
  - **Periodic Inform Interval Write Test (Row 9)**: Tests changing `PeriodicInformInterval` to 60s and back to 30s, or safely blocks the 60s write if an active GenieACS preset enforces 30s.
  - **TR-143 Download Diagnostics (Row 20)**: Verifies via HTTP `HEAD` that the local 50 MiB payload matches `tr143_server/test_50MB.bin`, triggers `Device.IP.Diagnostics.DownloadDiagnostics`, polls until `DiagnosticsState == Completed`, and validates `TestBytesReceived`, `BOMTime`, and `EOMTime`.
  - **IP Ping Diagnostics (Row 30)**: Configures `Device.IP.Diagnostics.IPPing` (3 repetitions, 2000 ms timeout), triggers execution, and evaluates `SuccessCount`, `FailureCount`, and min/avg/max response times.
  - **Safety Guarantee**: Never performs factory resets, reboots, firmware upgrades, APN modifications, or password/credential changes.
- **Preservation of Manual User Validations (`VALIDAZIONE UTENTE`)**:
  - Previously validated manual outcomes in `TOL_TCL.xlsx` (such as `CPE_0056_tr069` on row 68) are automatically preserved across runs under the `VALIDAZIONE UTENTE` scope—provided the CPE identity (`_id`), source TOL SHA-256 hash, and test ID match—while still logging the latest automated check in the notes column.

### 4. Packet-Level Periodic Inform Capture & Verification (`PeriodicCapture`)
- Reads the active CWMP port from `Device.ManagementServer.URL` and `PeriodicInformInterval` from the CPE.
- Spawns a temporary container on the host network to capture TCP traffic on the CWMP port over two full inform cycles, then filters and saves only the relevant TCP streams to `proof.pcap` (with `0600` permissions and SHA-256 integrity hashing).
- Dissects the PCAP via `tshark` (`decode_pcap_messages` and `pcap_periodic_rows`) to:
  1. Match the exact CPE `SerialNumber` inside the SOAP `Inform` XML body.
  2. Extract three consecutive `2 PERIODIC` `Inform` requests and verify that each receives an HTTP `200 OK` SOAP `InformResponse` on the same TCP stream.
  3. Verify `RetryCount == 0` on all three informs.
  4. Measure the two inter-arrival intervals and verify compliance with `PeriodicInformInterval` within a `±5s` tolerance.
  5. Cross-correlate packet timestamps and `cpeRequestId` values with GenieACS `/var/log/cwmp.log`.

### 5. Secret Redaction & Multi-Format Data Model Export
- Automatically identifies sensitive parameters (`password`, `passphrase`, `secret`, `token`, `privatekey`, `presharedkey`, `wepkey`, `username`, `credential`, `pin`, `puk`, `psk`) via regex (`SENSITIVE_PATH`).
- Masks sensitive values as `[oscurato]` in tables and recursively scrubs any occurrence of those secret strings from device IDs, metadata, notes, and log excerpts (`redact_secrets`, `sanitise_text`).
- Explicitly distinguishes between **empty string** (`valore vuoto`), **`null`**, and **unread node** (`non letto`).
- Exports the full sanitized tree to `cpe_datamodel_full.json` and `cpe_datamodel_flat.xlsx`, and updates the baseline audit table in `test.xlsx`.

### 6. Automated PowerPoint & Excel Report Generation
- **Excel Output (`TOL_TCL.xlsx`)**: Copies `TOL _Huawei_H173-383.xlsx`, writes test outcomes and parameter values into Column `J` (with conditional color formatting for `PASS`, `FAIL`, `BLOCCATO`, `PARZIALE`) and technical notes into Column `K`, and creates the `Valori_TCL` sheet containing every read parameter, value, status, scope, and UTC timestamp. Before completion, it cryptographically verifies that source columns `A:I`, cell styles, and merged ranges of the original template were not altered.
- **PowerPoint Output (`report_audit_tr069.pptx`)**: Built from the 13-slide editorial template (`report_template_tr069.pptx`):
  - **Slides 1–8**: Dynamically updated with current audit metadata, TOL summary counts across all 78 test cases, TR-181 data model coverage metrics, and packet-level `2 PERIODIC` PCAP timing evidence (displaying `WARN` instead of `PASS` if PCAP evidence is missing, stale `>1h`, or mismatched).
  - **Slides 9–13**: Preserved verbatim from the editorial template (including testbed architecture diagrams, proposed Hyperway IDU/ODU power-up matrices, and network impairment test proposals).
  - **Slides 14+ (Appendix)**: Automatically generated multi-slide editable tables listing every parameter in the CPE data model. No OLE objects or Excel files are embedded inside the presentation.

---

## Usage & Command-Line Interface

Run the validator using the preconfigured Python virtual environment (`venv`):

```bash
# Default run: refreshes Device data model, runs read-only TOL suite, captures
# 3 Periodic Informs (2 intervals) via PCAP, and generates Excel, JSON, and PPTX reports
venv/bin/python tr069_validator_agent.py
```

### CLI Options

```bash
# Run full TOL suite including reversible active tests (SPV + read-back + rollback) and TR-143 diagnostics
venv/bin/python tr069_validator_agent.py --tol-active

# Capture a new Periodic Inform PCAP proof and regenerate the PPTX report without rewriting TOL_TCL.xlsx
venv/bin/python tr069_validator_agent.py --periodic-only

# Regenerate data model exports, test.xlsx, and PPTX report from cached ACS data without running live CPE tasks
venv/bin/python tr069_validator_agent.py --report-only

# Validate Periodic Inform intervals from an existing PCAP file instead of running a live capture
venv/bin/python tr069_validator_agent.py --pcap /path/to/capture.pcap

# Specify a custom 13-slide editorial PowerPoint template
venv/bin/python tr069_validator_agent.py --report-only --template report_template_tr069.pptx

# Specify a custom HTTP/HTTPS URL for the 50 MiB TR-143 DownloadDiagnostics payload
venv/bin/python tr069_validator_agent.py --tol-active --download-url http://192.168.1.219:8080/test_50MB.bin

# Run the 22 isolated unit and pipeline self-tests (no ACS or CPE connection required)
venv/bin/python tr069_validator_agent.py --self-test
```

### Exit Codes

| Exit Code | Meaning |
| :--- | :--- |
| `0` | Execution completed successfully and Periodic Inform PCAP evidence is valid and compliant. |
| `2` | Execution completed, but measured Periodic Inform intervals or retry counts failed compliance checks. |
| `3` | Execution blocked due to unavailable prerequisites, unreachable ACS/CPE, or missing/expired PCAP evidence. |
| `130` | Execution interrupted by user/signal (`SIGINT` / `SIGTERM`); temporary changes restored and helper resources cleaned up. |

*(Individual test case outcomes—`PASS`, `FAIL`, `PARZIALE`, `BLOCCATO`, `NON ESEGUITO`—are recorded inside `TOL_TCL.xlsx` and `logs/tol_tcl_latest.json`.)*

---

## Repository Structure & Key Files

- `tr069_validator_agent.py`: Single-entry-point Python validator for TOL execution, TR-069/TR-181 auditing, PCAP capture/dissection, and Excel/PowerPoint report generation.
- `lab.conf`, `acs.startup`, `mongo.startup`, `Dockerfile.genieacs`, `Dockerfile.mongo`: Kathará and Docker configuration defining the GenieACS + MongoDB testbed topology.
- `TOL _Huawei_H173-383.xlsx`: Original read-only operator TOL template (columns `A:I` are strictly preserved).
- `TOL_TCL.xlsx`: Generated TOL workbook containing TCL CPE test results (`J`), technical notes (`K`), and the full parameter sheet (`Valori_TCL`).
- `test.xlsx`: Baseline synthetic audit workbook (automatically recreated from built-in baseline expectations if missing; expects `PeriodicInformInterval = 30`).
- `report_template_tr069.pptx`: 13-slide editorial PowerPoint template used as the input layout for report generation (slides 9–13 contain static editorial content and architecture/test proposals).
- `report_audit_tr069.pptx`: Generated final PowerPoint presentation (slides 1–8 dynamically updated, slides 9–13 preserved, slides 14+ populated with the full TR-181 data model appendix).
- `cpe_datamodel_full.json` & `cpe_datamodel_flat.xlsx`: Complete exported CPE data model with sensitive credentials redacted and explicit state tagging (`valore vuoto`, `null`, `non letto`).
- `periodic_inform_audit.log`: Human-readable summary of verified CWMP Periodic Inform frames, HTTP 200 responses, and measured intervals.
- `tr143_server/test_50MB.bin`: 50 MiB payload served by the local HTTP server for TR-143 `DownloadDiagnostics` testing.
- `venv/` & `requirements.txt`: Python virtual environment and dependencies (`openpyxl==3.1.5`, `python-pptx==1.0.2`).
- `logs/`: Timestamped execution directories containing structured event logs (`events.jsonl`), TOL results (`results.json`), private filtered PCAP captures (`proof.pcap`, mode `0600`), and workspace consolidation records (`logs/consolidation_20261017/cleanup_manifest.json`).

---

## Editorial Template & Future Extensions

- **Editing Slides**: To customize editorial text or graphics, edit `report_template_tr069.pptx` (not `report_audit_tr069.pptx`, which is overwritten on each run). Keep all 13 template slides and the shape IDs updated by the validator on slides 1–8.
- **Proposed Extended Scenarios (Slides 10–13)**: The **Hyperway IDU/ODU** startup sequence matrices (two startup orders × two delay intervals of 5s and 300s) and the **network interruption/degradation scenarios** documented in slides 10–13 are architectural test proposals, not automated routines executed by this script. Running them requires dedicated power/link control hardware, a controllable WAN/cellular topology, and agreed pass/fail timeout criteria.
