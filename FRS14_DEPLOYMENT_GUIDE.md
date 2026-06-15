# FRS14 Bridge — Deployment Guide

## What Is This Project?

This project converts **PDF documents → SGML** for Thomson Reuters securities outsourcing.  
The pipeline is:

```
PDF  →  [FRS14 bridge]  →  DOCX  →  [batch_runner_deploy.py]  →  _TR.sgm (SGML)
```

The UI runs on **Ubuntu/Plexus** (Linux).  
The PDF-to-DOCX conversion uses **ABBYY FineReader Server 14** which is **Windows-only**.

---

## The Core Problem

ABBYY FineReader Server 14 (FRS14) **cannot run on Linux**.  
Our Plexus deployment is Ubuntu. These two facts create a conflict:

```
Ubuntu (Plexus)           ✗ Cannot run ABBYY FRS14
Windows machine           ✓ Has ABBYY FRS14 installed and licensed
```

If the Streamlit app simply tries to call ABBYY directly from Ubuntu, it fails immediately.

---

## The Solution: FRS14 HTTP Bridge

We solved this by splitting the conversion into two parts across two machines:

### Part 1 — Windows side (`frs14_bridge.py`)

A lightweight **Flask HTTP server** that runs permanently on the Windows FRS14 machine.  
It listens on port `7090` and does the following when it receives a PDF:

1. Writes an XML ticket to the FRS14 hot folder input directory
2. Drops the PDF into the same hot folder
3. ABBYY FineReader Server picks it up and converts it automatically
4. The bridge polls the FRS14 output folder every 5 seconds (up to 600 seconds)
5. Once the DOCX appears, the bridge reads it and sends the bytes back over HTTP

### Part 2 — Ubuntu side (`streamlit_app.py`)

The Streamlit app on Ubuntu simply calls:
```
POST http://<windows-ip>:7090/convert
```
with the PDF as a file upload. It gets DOCX bytes back. No ABBYY, no Windows COM, no file shares needed.

### Architecture Diagram

```
┌──────────────────────────────────┐         ┌──────────────────────────────────────┐
│   UBUNTU / PLEXUS                │         │   WINDOWS MACHINE (FRS14)            │
│                                  │         │                                      │
│   streamlit_app.py               │         │   frs14_bridge.py  (port 7090)       │
│                                  │  HTTP   │                                      │
│   User uploads PDF  ─────────────┼────────▶│   Receives PDF                       │
│                                  │  POST   │   ↓                                  │
│                                  │         │   Writes XML ticket to hot folder    │
│                                  │         │   ↓                                  │
│                                  │         │   Drops PDF into hot folder          │
│                                  │         │   ↓                                  │
│                                  │         │   ABBYY FRS14 converts PDF → DOCX   │
│                                  │         │   ↓                                  │
│   Receives DOCX bytes  ◀─────────┼─────────│   Reads DOCX, returns bytes         │
│   ↓                              │  HTTP   │                                      │
│   batch_runner_deploy.py         │  200 OK │                                      │
│   ↓                              │         │                                      │
│   Output: _TR.sgm (SGML)         │         │                                      │
└──────────────────────────────────┘         └──────────────────────────────────────┘
```

---

## Hot Folder Method (Why Not COM API?)

FRS14 has two ways to submit jobs:

| Method | Status | Notes |
|---|---|---|
| COM API (`ProcessFileAsync`) | ❌ Broken | Requires Windows COM on the calling machine. Clogged with stale jobs. |
| Hot folder | ✅ Working | Write ticket + drop PDF → ABBYY picks it up automatically. Tested end-to-end. |

The bridge uses **hot folder exclusively**. This is the only proven working method.

**Hot folder paths on the Windows machine:**

| Purpose | Path |
|---|---|
| Input (ticket + PDF dropped here) | `C:\Users\Public\ABBYY\ABBYY FineReader Server 14.0\Default Workflow\Input Folder` |
| Output (converted DOCX appears here) | `C:\FRS_Output` |
| Staging (temp storage) | `C:\FRS_Input` |

All paths are overridable via environment variables on the Windows side.

---

## Deployment: Step-by-Step

### Prerequisites

| Machine | Requirement |
|---|---|
| Windows | Python 3.x installed, `flask` and `waitress` installed (`pip install flask waitress`) |
| Windows | ABBYY FineReader Server 14 installed, licensed, and the Default Workflow active |
| Ubuntu | Python 3.x, all packages from `requirements.txt` installed |
| Network | Ubuntu can reach Windows on port 7090 (firewall rule added) |

---

### Step 1 — Find the Windows IP from Ubuntu

Run this inside Ubuntu (WSL or remote server):

```bash
# If using WSL on the same Windows machine:
cat /etc/resolv.conf | grep nameserver | awk '{print $2}'
# Typically returns: 172.25.16.1

# If Ubuntu is a remote server on the same LAN, run this on Windows:
# Get-NetIPAddress -AddressFamily IPv4 | Where-Object InterfaceAlias -match "Wi-Fi"
# Typically returns: 192.168.1.x
```

Note down this IP — it is your `<WINDOWS_IP>`.

---

### Step 2 — Start the Bridge on Windows

Open a **PowerShell terminal on the Windows machine** and run:

```powershell
cd "C:\Users\C303180\OneDrive - Thomson Reuters Incorporated\Desktop\TR\sgml-pipeline-deployment_EXE\SGMLConverter_v2.3\SGMLConverter\_internal\pipeline"

python frs14_bridge.py --host 0.0.0.0 --port 7090
```

Expected output:
```
============================================================
FRS14 HTTP bridge  (hot folder)
  Listening : 0.0.0.0:7090
  WF input  : C:\Users\Public\ABBYY\ABBYY FineReader Server 14.0\Default Workflow\Input Folder
  Output    : C:\FRS_Output
  Staging   : C:\FRS_Input
  Timeout   : 600s
============================================================
Using waitress (multi-threaded)
```

> **Important:** Keep this terminal open. The bridge must stay running for the entire session.  
> If Windows reboots, restart this command. No other steps are needed after reboot —  
> ABBYY FRS14 services start automatically.

---

### Step 3 — Configure Ubuntu to Point at Windows

On Ubuntu, set the environment variable before starting the app:

```bash
export FRS14_SERVER_URL=http://<WINDOWS_IP>:7090
```

Or add it to your `.env` file in the deployment directory:

```bash
# .env
FRS14_SERVER_URL=http://172.25.16.1:7090   # WSL on same machine
# OR
FRS14_SERVER_URL=http://192.168.1.3:7090   # Remote Ubuntu on LAN
```

> **Why localhost fails:**  
> `http://localhost:7090` on Ubuntu points to Ubuntu itself — not the Windows machine.  
> The bridge is not running on Ubuntu, so you get `Connection refused`.  
> You **must** use the Windows machine's IP address.

---

### Step 4 — Install Dependencies on Ubuntu

```bash
cd /path/to/your/deployed/repo
pip install -r requirements.txt
```

Key packages for the bridge connection:

| Package | Version | Purpose |
|---|---|---|
| `requests` | 2.32.3 | Ubuntu calls the bridge via HTTP POST |
| `flask` | 3.1.3 | Bridge HTTP server (Windows side) |
| `waitress` | 3.0.2 | Production WSGI server for bridge (Windows side) |
| `httpx` | 0.28.1 | Used by Anthropic SDK and streamlit_app |

---

### Step 5 — Verify Connectivity from Ubuntu

Before starting the app, confirm Ubuntu can reach the bridge:

```bash
curl http://<WINDOWS_IP>:7090/health
```

Expected response (HTTP 200):
```json
{
  "status": "ok",
  "service": "FRS14 bridge (hot folder)",
  "frs14_service": "running",
  "wf_input_folder": "ok",
  "output_folder": "ok",
  "staging_folder": "ok"
}
```

**If you get `Connection refused`:**
- The bridge is not running on Windows → go back to Step 2

**If you get `No route to host`:**
- Firewall is blocking port 7090 on Windows
- Run this on Windows as Administrator:
  ```powershell
  netsh advfirewall firewall add rule name="FRS14 Bridge" dir=in action=allow protocol=TCP localport=7090
  ```

**If status is `degraded`:**
- ABBYY FRS14 service is not running on Windows
- Open Windows Services (`services.msc`) and start `ABBYY FineReader Server 14.0`

---

### Step 6 — Start the Streamlit App on Ubuntu

```bash
cd /path/to/your/deployed/repo
FRS14_SERVER_URL=http://<WINDOWS_IP>:7090 streamlit run streamlit_app.py
```

Or if `.env` is configured:

```bash
streamlit run streamlit_app.py
```

The sidebar will show the FRS14 Bridge Config panel with a green "✅ FRS14 bridge reachable" message.

---

## What Happens When a PDF Is Uploaded

```
1. User uploads PDF in Streamlit UI (Ubuntu)
2. streamlit_app._pdf_to_docx() sends POST to http://<WINDOWS_IP>:7090/convert
3. frs14_bridge.py (Windows) receives PDF
4. Bridge writes XML ticket → drops PDF into ABBYY hot folder
5. ABBYY FineReader Server 14 processes the PDF (OCR + layout analysis)
6. ABBYY writes DOCX to C:\FRS_Output
7. Bridge detects DOCX, reads bytes, returns them as HTTP 200 response
8. Ubuntu receives DOCX bytes (typically in 60–180 seconds)
9. batch_runner_deploy.py converts DOCX → SGML (_TR.sgm)
10. User downloads the SGML file
```

---

## Environment Variables Reference

### Ubuntu / Plexus side

| Variable | Required | Default | Description |
|---|---|---|---|
| `FRS14_SERVER_URL` | **Yes** | `http://localhost:7090` | URL of the Windows FRS14 bridge |
| `ANTHROPIC_MODEL` | Yes | — | Claude model for LLM calls |
| `TR_AUTH_URL` | Yes | — | Thomson Reuters auth endpoint |
| `WORKSPACE_ID` | Yes | — | TR workspace ID |
| `USE_LLM` | No | `true` | Enable/disable LLM in pipeline |

### Windows bridge side

| Variable | Default | Description |
|---|---|---|
| `FRS14_WF_INPUT` | `C:\Users\Public\ABBYY\...\Input Folder` | Hot folder input path |
| `FRS14_OUTPUT` | `C:\FRS_Output` | ABBYY output folder path |
| `FRS14_STAGING` | `C:\FRS_Input` | Staging area for PDFs |
| `FRS14_POLL_TIMEOUT` | `600` | Seconds to wait for DOCX output |
| `FRS14_POLL_INTERVAL` | `5` | Seconds between output checks |

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `FRS14 bridge not reachable at http://localhost:7090` | `FRS14_SERVER_URL` not set on Ubuntu | Set `FRS14_SERVER_URL=http://<WINDOWS_IP>:7090` |
| `Connection refused` on curl | Bridge not running on Windows | Start `python frs14_bridge.py` on Windows |
| `No route to host` | Windows Firewall blocking 7090 | Add inbound rule for TCP 7090 on Windows |
| Bridge health returns `degraded` | ABBYY FRS14 service down | Start the FRS14 service on Windows (`services.msc`) |
| Conversion times out after 600s | ABBYY workflow not processing | Check ABBYY Station is connected and workflow is active |
| DOCX is 0 bytes or very small | ABBYY failed OCR | Check FRS14 logs; reboot Windows if stale jobs queued |

---

## Key Files

| File | Location | Purpose |
|---|---|---|
| `frs14_bridge.py` | `_internal/pipeline/` | **Windows** — Flask HTTP bridge server |
| `streamlit_app.py` | repo root | **Ubuntu** — Streamlit UI, calls bridge |
| `batch_runner_deploy.py` | `pipeline/` | **Ubuntu** — DOCX → SGML converter |
| `requirements.txt` | repo root | All Python dependencies for Ubuntu |
| `.env.example` | repo root | Template for environment variables |
