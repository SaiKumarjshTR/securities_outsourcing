# ABBYY FineReader 12 — Plexus Deployment Guide & Knowledge Base

**Project:** SGML Pipeline (Securities Outsourcing) — PDF→DOCX/SGML conversion
**Repo:** `/home/securities_outsourcing`
**Component:** ABBYY FineReader Engine 12 (Linux) inside a Docker container on Plexus (AWS EKS)
**Current image version:** `v0.0.22`
**Last updated:** 2026-07-13

---

## Table of Contents
1. [The Problem](#1-the-problem)
2. [Root Cause](#2-root-cause)
3. [Why It Works Locally But Fails in Plexus](#3-why-it-works-locally-but-fails-in-plexus)
4. [The Three Deployment Options](#4-the-three-deployment-options)
5. [NEW: /dev/shm Size — Second Independent Failure Cause](#5-new-devshm-size--second-independent-failure-cause)
6. [The v0.0.22 Fixes](#6-the-v0022-fixes)
7. [Key Files & Paths](#7-key-files--paths)
8. [Build & Push Procedure](#8-build--push-procedure)
9. [Local Verification (Proof It Works)](#9-local-verification-proof-it-works)
10. [Plexus Deployment Steps](#10-plexus-deployment-steps)
11. [EC2 License Server Setup (Option B)](#11-ec2-license-server-setup-option-b)
12. [NEW: Plexus Two-Container LS Setup (Option B2 — No EC2)](#12-new-plexus-two-container-ls-setup-option-b2--no-ec2)
13. [The Dockerfile Explained](#13-the-dockerfile-explained)
14. [Temp Folder / Session Management Clarification](#14-temp-folder--session-management-clarification)
15. ["Baked Into the Image" Explained](#15-baked-into-the-image-explained)
16. [Email / Communication Templates](#16-email--communication-templates)
17. [Technical Reference](#17-technical-reference)

---

## 1. The Problem

When deployed to Plexus, ABBYY conversion fails:

```
❌ ABBYY conversion failed (rc=1): ABBYY FineReader Engine 12 Sample (c) 2022 ABBYY Development Inc.
Error: Failed to communicate with the online licensing service.
```

The container starts fine, Streamlit UI comes up, but **every PDF→DOCX conversion fails**.

---

## 2. Root Cause

ABBYY FineReader Engine 12 uses an **Online Protection License**:
- **Serial:** `SWAT-1241-1007-4471-4083-6499`
- **Customer Project ID:** `3AA4tXXQRiDuRh9Hh7P9`
- **License type:** Online Protection

On **every conversion**, the ABBYY engine's `LicensingService` must reach ABBYY's cloud at `account.abbyy.com:443` (HTTPS) to lease a license slot. If it cannot reach that endpoint, it waits ~2 minutes then fails with *"Failed to communicate with the online licensing service."*

```
┌───────────────────────┐      HTTPS:443       ┌──────────────────────┐
│  Plexus Container     │ ──────────────────→  │  account.abbyy.com   │
│  (FREngine12 CLI)     │   "Can I use the     │  (ABBYY Cloud)       │
│                       │    license?"         │                      │
│                       │  ←────────────────── │  "Approved for 60s"  │
└───────────────────────┘                      └──────────────────────┘
```

**Plexus containers have no outbound internet access** (TR network policy blocks egress), so this call never completes → conversion fails.

> **This is a network/licensing problem, NOT a code or image problem.** No code change can bypass ABBYY's license check.

---

## 3. Why It Works Locally But Fails in Plexus

| Environment | Internet to `account.abbyy.com`? | Result |
|---|---|---|
| Local (WSL dev machine) | ✅ Yes — via TR Zscaler corporate proxy | Conversion works |
| Plexus (AWS pod) | ❌ No — outbound internet blocked | "Failed to communicate…" |

The container already bundles TR/Zscaler CA certs (`abbyy_bundle/tr_certs/`) so SSL verification succeeds — but that only matters if the container can reach the internet in the first place.

---

## 4. The Three Deployment Options

| | Option A: Allowlist | Option B: EC2 License Server | Option B2: Plexus LS Container | Option C: Local License |
|---|---|---|---|---|
| **What** | Allowlist outbound `account.abbyy.com:443` from Plexus pods | Run ABBYY `LicensingService` on an EC2 in a **public** subnet | Run `LicensingService` as a **separate Plexus container** with its own internet policy | Ask ABBYY to convert to an offline Local License |
| **Code changes** | None | Done (v0.0.22, `ABBYY_LS_HOST`) | Done (`Dockerfile.licensing`, `ABBYY_LS_HOST`) | Dockerfile changes |
| **New infra** | None | 1× t3.nano EC2 (~$4/mo) | 1× Plexus deployment | None |
| **Runtime internet** | Yes (container) | Yes (EC2 only) | **Yes (LS container only)** | **None** |
| **Who does the work** | TR network/infra team | Infra (EC2) + us (code done) | **Us + Plexus internet policy for LS pod** | ABBYY + us |
| **Can we do it alone?** | ❌ Needs network team | ❌ Needs infra to launch EC2 | ⚡ Nearly — needs network policy for LS pod only | ❌ Needs ABBYY approval |
| **Risk / timeline** | Low / 1–2 days | Low / infra 1–2h, then us ~2h | Low / 2–4h | Medium / 2–5 days (ABBYY) |

> **Option B2 is the new recommended path**: it doesn't need EC2, uses our existing code, and only the small LS container needs outbound internet (much easier to approve than opening the entire app pod).

### Public subnet is critical (Option B — EC2)
- **Private subnet** EC2s route internet traffic through **Transit Gateway** `tgw-0d3aebf5bb40f0224` → same TR policy that blocks Plexus. They CANNOT reach `account.abbyy.com`.
- **Public subnet** (`subnet-0af01744c5615bbec`) routes `0.0.0.0/0` directly to **Internet Gateway** `igw-051d5e32b554b95e5` → unrestricted. An EC2 here CAN reach `account.abbyy.com`.

---

## 5. NEW: /dev/shm Size — Second Independent Failure Cause

> **This was discovered from ABBYY's official Docker documentation and is separate from the licensing issue.**

### What the ABBYY docs say
> "`shm_size: 1g` is mandatory. FRE relies on POSIX shared memory and the default 64MB `/dev/shm` allocation is not enough."

### The problem in Kubernetes (Plexus)
| Environment | `/dev/shm` size | ABBYY OCR |
|---|---|---|
| `docker run` (default) | 64 MB | ❌ Fails |
| Kubernetes pod (default) | 64 MB | ❌ Fails |
| Dev machine (this WSL host) | 7.8 GB | ✅ Works |
| Fixed Kubernetes pod | 1 GB+ | ✅ Works |

### How to fix (Kubernetes pod spec)
```yaml
volumes:
  - name: dshm
    emptyDir:
      medium: Memory
      sizeLimit: 1Gi
volumeMounts:
  - mountPath: /dev/shm
    name: dshm
```

In Plexus, request the infra/Plexus team to add this to the pod spec for `sgml-pipeline-prod`.

### How to verify
```bash
# Inside the container:
df -h /dev/shm
# Should show: Filesystem Size ≥ 1G

# Or via diagnostics:
python3 /app/abbyy_convert.py --diag   # now includes SHM Size check
```

The `docker_start.sh` also logs a warning at startup if `/dev/shm` is under 1GB.

---

## 6. The v0.0.22 Fixes

### Changes from v0.0.21 → v0.0.22
| Fix | Detail |
|---|---|
| Missing system libs | Added `libx11-6`, `libfreetype6`, `libice6`, `locales` — required by ABBYY official Docker docs |
| Removed dev library | `libProtection.Developer.so` removed from runtime image (dev-only per ABBYY docs) |
| shm check | `docker_start.sh` logs a warning if `/dev/shm` < 1GB at container start |
| New diagnostic | `abbyy_convert.py --diag` now includes check #5: SHM Size |
| Dockerfile.licensing | New lightweight LS-only container for two-container Plexus deployment |

### The v0.0.21 fix (ABBYY_LS_HOST) — still present in v0.0.22
`scripts/docker_start.sh` patches **all three** copies of `LicensingSettings.xml` when `ABBYY_LS_HOST` is set:
```
/opt/ABBYY/FREngine12/Bin/LicensingSettings.xml
/opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml
/usr/local/lib/ABBYY/SDK/12/Licensing/LicensingSettings.xml
```

### How the container behaves
- **`ABBYY_LS_HOST` NOT set** → starts a LOCAL `LicensingService` (tries `account.abbyy.com` directly → fails in Plexus).
- **`ABBYY_LS_HOST=<ip-or-hostname>` set** → patches all 3 settings files to point to the external/sidecar license server on port 3023, skips local service.

---
The container has **THREE** copies of `LicensingSettings.xml`. The FREngine CLI reads its copy from the **FREngine12 tree** (`Bin/` and `CommonBin/Licensing/`), NOT from `/usr/local/lib`. The earlier version only patched the `/usr/local/lib` copy, so the engine kept using `ServerAddress="127.0.0.1"` and Option B would silently fail.

The three files:
```
/opt/ABBYY/FREngine12/Bin/LicensingSettings.xml
/opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml
/usr/local/lib/ABBYY/SDK/12/Licensing/LicensingSettings.xml
```

*(This detail is preserved for reference — the v0.0.22 `docker_start.sh` patches all three correctly.)*

---

## 7. Key Files & Paths

### Application / build files
| File | Purpose |
|---|---|
| `Dockerfile` | Image definition, `LABEL version="0.0.22"` |
| `Dockerfile.licensing` | **NEW** — LicensingService-only container for two-container Plexus deployment (Option B2) |
| `scripts/docker_start.sh` | Container entrypoint. Contains `ABBYY_LS_HOST` patch logic + shm check. NO `set -e`. |
| `scripts/bundle_abbyy.sh` | Bundles ABBYY from host into `abbyy_bundle/` |
| `scripts/setup_ec2_license_server.sh` | Installs LicensingService on an EC2 (Option B) |
| `scripts/build_push_v017.sh` | **NEW** — Builds + pushes `v0.0.22` |
| `scripts/build_push_licensing_v001.sh` | **NEW** — Builds + pushes `sgml-abbyy-ls:1.0.0` (LS sidecar image) |
| `abbyy_convert.py` | Standalone diagnostic + conversion script (`--diag` flag, now includes SHM check) |
| `app/session_manager.py` | Session temp folder logic (line 88 warning) |
| `app/app_config.py` | `TEMP_DIR` config (line 43) |

### ABBYY install paths inside the container (baked into image)
| What | Path |
|---|---|
| Engine + CLI + libraries | `/opt/ABBYY/FREngine12` |
| CLI converter | `/opt/ABBYY/FREngine12/Samples/CommandLineInterface/CommandLineInterface` |
| Engine libs (`LD_LIBRARY_PATH`) | `/opt/ABBYY/FREngine12/Bin` |
| LicensingService daemon | `/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService` |
| Licensing libraries | `/usr/local/lib/ABBYY/SDK/12/Licensing` |
| License files / tokens | `/var/lib/ABBYY/SDK/12/Licenses/` |
| License activation token | `/var/lib/ABBYY/SDK/12/Licenses/SWAT12411007447140836499.ABBYY.ActivationToken` |
| Protection state | `/var/lib/ABBYY/SDK/12/Licenses/Protection.ccf` |

### abbyy_bundle/ structure (build context)
```
abbyy_bundle/
├── FREngine12/                      # full engine SDK
│   ├── Bin/LicensingSettings.xml
│   └── CommonBin/Licensing/LicensingSettings.xml
├── LicensingService                 # daemon binary (ELF x86-64)
├── codemeter.deb                    # CodeMeter license manager
├── tr_certs/                        # 8 TR/Zscaler CA certs
├── usr_local_lib_abbyy/
│   └── SDK/12/Licensing/            # LicensingService libs + LicensingSettings.xml
└── var_lib_abbyy/
    └── ABBYY/SDK/12/Licenses/       # activation tokens + Protection.ccf
```

---

## 8. Build & Push Procedure

### AWS login (required first)
```bash
source /root/cloud-tool-env/bin/activate && cloud-tool --region us-east-1 login
# User: MGMT\mC303180
# Token expires ~4h; re-run if expired.
```

### Build & push v0.0.22 (current)
```bash
bash scripts/build_push_v017.sh C303180 <jfrog_token>
```

JFrog token: use your personal token from JFrog Artifactory (see TR1 JFrog portal).

### Also build & push the LS sidecar image (for Option B2)
```bash
bash scripts/build_push_licensing_v001.sh
```

### Result (v0.0.22)
- **Image tag:** `sgml-pipeline-prod-0.0.22`
- **ECR URI:** `127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1:sgml-pipeline-prod-0.0.22`
- Build takes ~50 min (slow proxy pip downloads)

### Previous v0.0.21 image
- **Digest:** `sha256:00631bb0bc69950136b4786a5cca554176e4bf16b1288823605a180aa373a979`

### Manual push (if needed)
```bash
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"
aws --profile tr-aiml-hackathon-prod ecr get-login-password --region us-east-1 | \
    docker login --username AWS --password-stdin "127288631409.dkr.ecr.us-east-1.amazonaws.com"
docker tag sgml-pipeline-prod:0.0.22 "$ECR:sgml-pipeline-prod-0.0.22"
docker push "$ECR:sgml-pipeline-prod-0.0.22"
```

---

## 9. Local Verification (Proof It Works)

Run the image locally with `--shm-size=1g` (required by ABBYY) and verify:

```bash
# Start container with shm fix (default/local LS mode)
docker run -d --name abbyytest --shm-size=1g -p 18501:8501 sgml-pipeline-prod:0.0.22

# Run ABBYY diagnostics (all 5 checks including SHM)
docker exec abbyytest python3 /app/abbyy_convert.py --diag

# Real single-page conversion test
docker exec abbyytest bash -c 'python3 /app/abbyy_convert.py /opt/ABBYY/FREngine12/Help/FREngine12AdminGuide.pdf /tmp/verify.docx'
```

### Verified results (2026-07-10)
| Check | Result |
|---|---|
| Container starts, Streamlit up on 8501 | ✅ |
| `/dev/shm` = 1GB (with `--shm-size=1g`) | ✅ |
| CodeMeterLin + LicensingService running, port 3023 listening | ✅ |
| ABBYY diagnostics (all 5 checks: binaries/services/network/license/shm) | ✅ all passed |
| Real PDF → DOCX conversion | ✅ valid DOCX (`PK` header) in 1.2s |
| `ABBYY_LS_HOST` patches all 3 LicensingSettings.xml | ✅ all three rewritten |

### Verify the ABBYY_LS_HOST patch
```bash
docker run -d --name abbyyls --shm-size=1g -e ABBYY_LS_HOST=10.156.60.42 sgml-pipeline-prod:0.0.22
docker logs abbyyls 2>&1 | grep -iE "ABBYY_LS_HOST|patch|ServerAddress"
docker exec abbyyls bash -c 'for f in \
  /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml \
  /opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml \
  /usr/local/lib/ABBYY/SDK/12/Licensing/LicensingSettings.xml; do \
  echo "[$f]"; grep -o "ServerAddress=\"[^\"]*\"" "$f"; done'
```

### Cleanup
```bash
docker rm -f abbyytest abbyyls
```

> **Important:** Local success proves the image + engine are healthy. The Plexus error is only resolved once EITHER: (a) shm is fixed in the pod spec, AND (b) a license server is reachable (Option A/B/B2) OR a Local License is used (Option C).

---

## 10. Plexus Deployment Steps

1. **Model Registry** → `sgml-pipeline-prod` → Add Version `0.0.22`
   - Image: `127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1:sgml-pipeline-prod-0.0.21`
   - Health: `/_stcore/health`  Port: `8501`
2. **Deployment → Environment tab** → add env var:
   - Key: `ABBYY_LS_HOST`
   - Value: `<ec2-private-ip>` — **must be the real EC2 IP** (e.g. `10.156.60.42`), digits only, no `<>`, no `http://`, no port.
3. **Activate** v0.0.21

> **Warning:** Setting `ABBYY_LS_HOST` alone is not enough — an EC2 license server must actually be running at that IP on port 3023. Without it, you get a connection error instead. The EC2 must exist FIRST.

### Post-deploy verification (run inside the pod)
```bash
# Confirm env var set + all 3 files patched
env | grep ABBYY_LS_HOST
for f in /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml \
         /opt/ABBYY/FREngine12/CommonBin/Licensing/LicensingSettings.xml \
         /usr/local/lib/ABBYY/SDK/12/Licensing/LicensingSettings.xml; do
  grep -o 'ServerAddress="[^"]*"' "$f"; done

# Confirm container can reach the EC2 on 3023
timeout 5 bash -c "cat < /dev/tcp/${ABBYY_LS_HOST}/3023" \
  && echo "✅ REACHABLE" || echo "❌ cannot reach ${ABBYY_LS_HOST}:3023"

# Full end-to-end
python3 /app/abbyy_convert.py --diag
python3 /app/abbyy_convert.py /opt/ABBYY/FREngine12/Help/FREngine12AdminGuide.pdf /tmp/verify.docx
```

Look for `✓ External LicensingService REACHABLE` in the Plexus startup logs (docker_start.sh tests this automatically).

---

## 11. EC2 License Server Setup (Option B)

### Infra request specs
- Instance: `t3.nano` (x86-64 — required; ABBYY binary is x86-64 ELF)
- AMI: Amazon Linux 2023 (x86-64)
- **Subnet:** `subnet-0af01744c5615bbec` (tr-vpc-1 **public**, us-east-1a) — must route via IGW `igw-051d5e32b554b95e5`
- Auto-assign public IP: yes
- Instance profile: `AmazonSSMManagedInstanceCore`
- **Security Group inbound:** TCP `3023` from Plexus CIDR `10.156.60.0/22`
- **Security Group outbound:** TCP `443` to `0.0.0.0/0` (default)

### AWS environment reference (account 058264507063)
| Item | Value |
|---|---|
| VPC | `vpc-0ae693f600bdcc28e` (tr-vpc-1) |
| Public subnets | `subnet-0af01744c5615bbec` (1a), `subnet-0cb2f0078c2bafecb` (1c), `subnet-08410db7f832f14d5` (1b) |
| Internet Gateway | `igw-051d5e32b554b95e5` |
| Transit Gateway (blocks internet) | `tgw-0d3aebf5bb40f0224` |
| Plexus pod CIDR | `10.156.60.0/22` |

### Install (after EC2 exists)
```bash
bash scripts/setup_ec2_license_server.sh ec2-user@<ec2-private-ip>
```

### What setup_ec2_license_server.sh does
1. Copies `LicensingService` binary → `/usr/local/bin/ABBYY/SDK/12/Licensing/`
2. Copies shared libs → `/usr/local/lib/ABBYY/`
3. Copies license files → `/var/lib/ABBYY/SDK/12/Licenses/`
4. Creates systemd service `abbyy-licensing.service`:
   ```ini
   [Service]
   ExecStart=/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService /standalone
   Environment=LD_LIBRARY_PATH=/usr/local/lib/ABBYY/SDK/12/Licensing
   Restart=always
   ```
5. Enables + starts the service, verifies port 3023 LISTENING

### Architecture
```
[Plexus Pod] --TCP:3023--> [EC2 License Server] --HTTPS:443--> account.abbyy.com
 (private subnet)           (public subnet, IGW)
```

### Blockers (current DataScientist role)
- `ec2:RunInstances` **DENIED** — cannot launch EC2
- `ec2:CreateSecurityGroup` **DENIED**
- `ssm:StartSession` **DENIED**
- No SSH key pairs
→ Infra/cloud team must launch the EC2.

---

## 12. NEW: Plexus Two-Container LS Setup (Option B2 — No EC2)

This is the **recommended production path** from ABBYY's official Docker documentation. Instead of an EC2, the LicensingService runs as a **separate, lightweight Plexus service** (`sgml-abbyy-ls`). Only that service needs outbound internet access — much easier to get approved than opening the entire app pod.

### Architecture
```
[App Pod — sgml-pipeline-prod]  --TCP:3023-->  [LS Pod — sgml-abbyy-ls]  --HTTPS:443-->  account.abbyy.com
  - Streamlit UI (8501)                          - LicensingService only
  - ABBYY CLI engine                             - ~50MB image
  - NO internet needed                           - Needs 443 outbound to *.abbyy.com
```

### Step 1: Build and push the LS image
```bash
source /root/cloud-tool-env/bin/activate && cloud-tool --region us-east-1 login
bash scripts/build_push_licensing_v001.sh
```
Image: `sgml-abbyy-ls:1.0.0` → ECR tag: `...sgml-abbyy-ls-1.0.0`

### Step 2: Deploy `sgml-abbyy-ls` in Plexus
- Model Registry → Add model `sgml-abbyy-ls`
- Image: `127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1:sgml-abbyy-ls-1.0.0`
- Port: `3023`
- Health check: TCP (not HTTP)
- **Network policy: allow outbound TCP 443 to `*.abbyy.com` from this pod only**
- Replicas: 1 (single Online License supports one LS at a time)

### Step 3: Get the cluster service hostname
After deploying, the LS is accessible within the cluster at a service DNS name like:
```
sgml-abbyy-ls.default.svc.cluster.local
```
or just `sgml-abbyy-ls` within the same namespace. Check via Plexus UI or `kubectl get svc`.

### Step 4: Configure the app pod
In `sgml-pipeline-prod` deployment, set:
```
ABBYY_LS_HOST=sgml-abbyy-ls.default.svc.cluster.local
```
(or `sgml-abbyy-ls` if same namespace)

### Step 5: Fix /dev/shm in the app pod (CRITICAL)
In Plexus pod spec for `sgml-pipeline-prod`:
```yaml
volumes:
  - name: dshm
    emptyDir:
      medium: Memory
      sizeLimit: 1Gi
volumeMounts:
  - mountPath: /dev/shm
    name: dshm
```

### Verify
After both pods are running:
```bash
# In the app pod:
env | grep ABBYY_LS_HOST
timeout 5 bash -c "cat < /dev/tcp/${ABBYY_LS_HOST}/3023" && echo "✅ LS REACHABLE"
df -h /dev/shm                          # must show ≥ 1G
python3 /app/abbyy_convert.py --diag    # all 5 checks should pass
```

---

## 13. Installer-Based Deployment (Ubuntu exe / activatefre.sh)

This is ABBYY's **official recommended method**: run the `FRE12.sh` self-extracting installer (or our `FRE_bundle.sh` equivalent) during `docker build`, which runs `activatefre.sh` to configure `LicensingSettings.xml` at **build time**.

### Key advantage over the bundle-copy approach
| | Bundle COPY (Dockerfile) | Installer approach (Dockerfile.installer) |
|---|---|---|
| LicensingSettings.xml config | Patched at **runtime** by `docker_start.sh` | Configured at **build time** by `activatefre.sh` |
| Requires `ABBYY_LS_HOST` env var | Yes (Mode B) | Optional — can be baked in |
| Official ABBYY method | No | **Yes** |
| Requires FRE*.sh installer | No | Yes (or use `FRE_bundle.sh`) |

### Two ways to get the FRE*.sh installer

**Option A — Official installer from ABBYY portal (recommended):**
```
1. Go to: https://support.abbyy.com/hc/en-us/p/FineReader-Engine-Trial
2. Log in with your ABBYY account
3. Download "FineReader Engine 12 for Linux" → FRE12.sh
4. Place FRE12.sh in /home/securities_outsourcing/ (build context root)
5. In Dockerfile.installer: uncomment "Option 1" block, comment "Option 2" block
```

**Option B — Create installer from existing bundle (no download needed):**
```bash
bash scripts/create_fre_bundle_installer.sh
# Creates: FRE_bundle.sh (~3.3GB self-extracting installer)
# Same CLI as official FRE12.sh — works identically
```

### Build modes

**Mode A — Hard-code LS address at build time (best for Plexus):**
```bash
bash scripts/build_push_installer.sh C303180 <jfrog_token> \
  sgml-abbyy-ls.default.svc.cluster.local
```
`activatefre.sh` writes `ServerAddress="sgml-abbyy-ls.default.svc.cluster.local"` directly into `LicensingSettings.xml` at build time. No env var needed at runtime.

**Mode B — Runtime-configurable (same flexibility as v0.0.22):**
```bash
bash scripts/build_push_installer.sh C303180 <jfrog_token>
# Then in Plexus: set ABBYY_LS_HOST=sgml-abbyy-ls.default.svc.cluster.local
```

### How `activatefre.sh` silent mode works
`activatefre.sh` is ABBYY's official setup script. In Docker builds it runs with:
```bash
export RootScriptLaunched=1   # tells script we are root (Docker always is)

activatefre.sh -- \
  --install-dir /opt/ABBYY/FREngine12 \
  --service-address sgml-abbyy-ls \    # writes ServerAddress to LicensingSettings.xml
  --skip-local-service-installation \  # no systemd (correct for Docker)
  --skip-local-license-activation      # don't try account.abbyy.com at build time
```

Key parameters:
| Parameter | Effect |
|---|---|
| `--service-address HOST` | Sets `ServerAddress="HOST"` in both LicensingSettings.xml files |
| `--license-path FILE` | Copies token file to `/var/lib/ABBYY/SDK/12/Licenses/` |
| `--license-password PWD` | Used with `--license-path` for online activation |
| `--skip-local-service-installation` | Does NOT install LicensingService as systemd service |
| `--skip-local-license-activation` | Does NOT call `account.abbyy.com` during build |

### Verify after build
```bash
docker run --rm sgml-pipeline-prod:0.0.22-installer \
  bash -c 'cat /opt/ABBYY/FREngine12/Bin/LicensingSettings.xml'
# Should show ServerAddress="sgml-abbyy-ls..." (Mode A) or "127.0.0.1" (Mode B)
```

---

## 14. The Dockerfile Explained

### Big picture
Packages 3 things into one image: (1) Streamlit web app (port 8501), (2) Python deps, (3) ABBYY FineReader Engine 12 (Linux).

### Step-by-step
1. **`FROM python:3.12-slim`** — minimal Debian + Python 3.12 base.
2. **System libraries** (`apt-get install`) — shared libs ABBYY/CodeMeter need (`libglib2.0-0`, `libgomp1`, `libusb-1.0-0`); `procps` (pgrep/ps), `iproute2` (ss).
3. **Python deps** — `pip install -r requirements.txt` from TR JFrog then public PyPI (the slow build step).
4. **App code** — copies `app/`, `pipeline/`, `streamlit_app.py`, `data/` into `/app`.
5. **ABBYY (4 pieces):**
   - **(a) TR/Zscaler certs** → trust TR's SSL re-signing for the license HTTPS call.
   - **(b) CodeMeter** (`codemeter.deb`, `--force-depends`) — headless license daemon.
   - **(c) Engine** → `/opt/ABBYY/FREngine12` (permanent), made executable.
   - **(d) LicensingService + license files** → `/usr/local/bin`, `/usr/local/lib`, `/var/lib/ABBYY`.
6. **Startup script** → copies `docker_start.sh`.
7. **ENV vars** — `ABBYY_CLI`, `ABBYY_LIB`, `LD_LIBRARY_PATH` (so binaries find their libs).
8. **HEALTHCHECK** — `/_stcore/health`; **EXPOSE 8501**.
9. **`CMD ["/app/docker_start.sh"]`** — entrypoint.

### docker_start.sh boot sequence
```
1. Start CodeMeterLin        (license daemon)
2. Start LicensingService    → OR if ABBYY_LS_HOST set, patch all 3 LicensingSettings.xml + skip local
3. Wait ~10s
4. exec Streamlit on port 8501
```
- **No `set -e`** — if licensing fails, Streamlit must still start so the UI is reachable.

### Conversion runtime flow
```
User uploads PDF (Streamlit 8501)
  → Python calls ABBYY CLI (/opt/ABBYY/.../CommandLineInterface)
  → Engine asks LicensingService "am I licensed?"
  → LicensingService validates (local→internet, OR external EC2 if ABBYY_LS_HOST set)
  → Engine converts PDF→DOCX → returned to user
```

---

## 15. Temp Folder / Session Management Clarification

### The harmless warning
```
WARNING:root:Cannot write to project tmp directory: [Errno 30] Read-only file system: '/app/tmp', using fallback
```
Source: `app/session_manager.py:88`.

### Session temp folder fallback order
1. `SESSIONS_TEMP_DIR` env var → (not set)
2. `/app/tmp/sessions` → **read-only in Plexus → WARNING → skip**
3. System temp `/tmp/pending_legislation_sessions` → **writable → USED ✅**

### Where sessions actually live
| Environment | `/app` writable? | Session location |
|---|---|---|
| Docker (local) | Yes | `/app/tmp/sessions` |
| Plexus | No (read-only) | `/tmp/pending_legislation_sessions` |

### Key facts
- The temp warning is **cosmetic** and **unrelated** to the ABBYY licensing failure.
- ABBYY (`/opt/ABBYY`) and session temp (`/app/tmp` or `/tmp`) are **completely separate** and do not overlap.
- Sessions are **ephemeral** (inside container, wiped on restart).
- **Optional cleanup** to remove the warning: set env var `SESSIONS_TEMP_DIR=/tmp/sessions` in Plexus.

### Rule of thumb
| Type | Where | Example |
|---|---|---|
| Program (permanent) | `/opt`, `/usr/local` | ABBYY engine |
| Temp/runtime data (throwaway) | `/tmp` | session files, output |

> **Do NOT** put ABBYY in a temp folder — `/tmp` is wiped on restart, which would break every conversion. ABBYY correctly lives in `/opt/ABBYY` (permanent).

---

## 16. "Baked Into the Image" Explained

"Baked in" = ABBYY is built into the container permanently at **build time**, so it's inside every copy of the image. Nothing is downloaded/installed at runtime.

The `COPY` lines that bake it in (run once during `docker build`):
```dockerfile
COPY abbyy_bundle/FREngine12 /opt/ABBYY/FREngine12
COPY abbyy_bundle/LicensingService /usr/local/bin/ABBYY/...
COPY abbyy_bundle/usr_local_lib_abbyy/ /usr/local/lib/ABBYY/
COPY abbyy_bundle/var_lib_abbyy/ABBYY/ /var/lib/ABBYY/
COPY abbyy_bundle/codemeter.deb ...
```

### The key nuance
- ✅ **ABBYY software** = baked in, always present, survives restarts, works offline.
- ❌ **ABBYY license validation** = a live action on every conversion that needs internet to `account.abbyy.com` → this is what fails in Plexus.

So the engine is baked in and fine; only its **internet license check** is blocked in Plexus.

---

## 17. Email / Communication Templates

### Reply to team lead (explaining the actual error)
> Hi [Lead],
>
> You've understood correctly — ABBYY is baked into the image (`/opt/ABBYY`), and the `/app/tmp` message is just a harmless warning (the app falls back to `/tmp`). That is **not** the problem.
>
> **The actual error is a licensing/network issue:**
> `ABBYY conversion failed (rc=1): Error: Failed to communicate with the online licensing service.`
>
> **Root cause:** Our ABBYY license is an online license. On every conversion, the engine must reach `account.abbyy.com:443` to validate. Plexus has no outbound internet, so it fails. Works locally (internet via Zscaler), fails in Plexus.
>
> It is purely a network reachability problem — not a code/image problem (verified locally end-to-end).
>
> **To fix (all need infra/network help):**
> 1. Allowlist outbound `account.abbyy.com:443` from Plexus pods (simplest).
> 2. Small EC2 license server in a public subnet — image already supports via `ABBYY_LS_HOST`.
> 3. Offline license from ABBYY (permanent, depends on ABBYY).
>
> We're on the correct path — the blocker is the network/infra piece.

### Meeting request to cloud expert (Sai Kiran)
> **Subject:** 30 min Monday? — need cloud help with a Plexus issue
>
> Hi Sai Kiran,
>
> [Lead] suggested I reach out. Could we meet 30 minutes on Monday?
>
> **The issue:** We run ABBYY (PDF→DOCX) in a Plexus container. ABBYY needs internet (`account.abbyy.com`) to check its license on every conversion. Plexus blocks outbound internet → *"Failed to communicate with the online licensing service."* Works locally (has internet).
>
> **Help needed:** advice on the best cloud fix — either allowlist this one outbound connection, or run a small EC2 as a license server the container talks to. Our role can't create EC2s, so I need your guidance.
>
> Thanks, [Your name]

### Network team allowlist request (Option A)
> **Destination:** `account.abbyy.com` (IP `52.211.188.155` / verify with nslookup)
> **Port/Protocol:** 443 (HTTPS), outbound only
> **Source:** Plexus pod CIDR `10.156.60.0/22`
> **Purpose:** ABBYY FREngine12 online license validation. No document data transmitted — license sync only, few KB per conversion.

---

## 18. Technical Reference

| Item | Value |
|---|---|
| ABBYY Serial | `SWAT-1241-1007-4471-4083-6499` |
| Customer Project ID | `3AA4tXXQRiDuRh9Hh7P9` |
| License type | Online Protection |
| ABBYY license endpoint | `account.abbyy.com:443` (HTTPS) |
| LicensingService port | `3023` (TCP) |
| Current image | `sgml-pipeline-prod-0.0.21` |
| Image digest | `sha256:00631bb0bc69950136b4786a5cca554176e4bf16b1288823605a180aa373a979` |
| ECR registry | `127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1` |
| AWS account (compute) | `058264507063` (tr-aiml-hackathon-prod) |
| VPC | `vpc-0ae693f600bdcc28e` (tr-vpc-1) |
| Public subnet (for EC2) | `subnet-0af01744c5615bbec` (us-east-1a) |
| Internet Gateway | `igw-051d5e32b554b95e5` |
| Transit Gateway (blocks internet) | `tgw-0d3aebf5bb40f0224` |
| Plexus pod CIDR | `10.156.60.0/22` |
| UI port | `8501` (Streamlit) |
| Health check | `/_stcore/health` |
| AWS login user | `MGMT\mC303180` |
| AWS login cmd | `source /root/cloud-tool-env/bin/activate && cloud-tool --region us-east-1 login` |
| ABBYY support | support@abbyy.com / larysa.lototska@abbyy.com / dennis.lappert@abbyy.com |

### Version history
| Version | Issue | Fix |
|---|---|---|
| v0.0.13 | Access to `/var/lib/ABBYY/.../FineReader Engine` denied | Bundled full `/var/lib/ABBYY/` (Protection.ccf) |
| v0.0.14 | Files at wrong path (extra dir level) | Fixed Dockerfile COPY path |
| v0.0.15 | Container wouldn't start | Removed `set -e` from docker_start.sh; reduced wait to 10s |
| v0.0.16 | SSL cert rejection | Added 8 TR/Zscaler CA certs + `update-ca-certificates` |
| v0.0.18–20 | Online license fails in Plexus (no internet) | Added `ABBYY_LS_HOST` external LS support |
| **v0.0.21** | Option B failed — only 1 of 3 LicensingSettings.xml patched | **Patch ALL 3 copies in docker_start.sh** |

---

## 19. Business User Distribution — Standalone Exe (Ubuntu + Windows)

**Key insight:** Business user machines **have internet access**. The ABBYY Online License
connects to `account.abbyy.com:443` and WORKS perfectly. The Plexus networking problem does
not exist for local installs.

**Constraint:** ABBYY FREngine12 is a **Linux x86-64 ELF binary** — it cannot run natively
on Windows. Windows users must use **WSL2** (Windows Subsystem for Linux 2), which runs a
real Ubuntu kernel inside Windows. The same binary works identically.

---

### 19.1 Deployment Paths

| Platform | Path | Internet Required | Admin Required |
|---|---|---|---|
| **Ubuntu / Debian** | `setup_local_ubuntu.sh` (direct) | Yes (license + pip packages on install) | Yes (apt install) |
| **Ubuntu** | `sgml-pipeline-ubuntu.run` (single file) | Yes | Yes |
| **Windows 10/11** | `setup_windows_wsl.ps1` + WSL2 | Yes (WSL2 Ubuntu download + pip) | **Yes (Admin)** |
| **Windows** (after setup) | `launch_sgml_pipeline.bat` | Yes (license validation per conversion) | No |

---

### 19.2 Creating the Distributable Bundle

**On the Ubuntu build machine (this repo):**

```bash
# Creates dist/sgml-pipeline-bundle.tar.gz  AND  dist/sgml-pipeline-ubuntu.run
bash scripts/create_portable_bundle.sh
```

This packages:
- App code (`app/`, `pipeline/`, `validator/`, `streamlit_app.py`, etc.)
- **Full ABBYY FREngine12 + LicensingService** from `abbyy_bundle/` (~3.3 GB)
- All setup scripts

**Share with business users:**
- **Ubuntu users** → `dist/sgml-pipeline-bundle-ubuntu.run` (single self-extracting file)
- **Windows users** → `dist/sgml-pipeline-bundle.tar.gz` + `scripts/setup_windows_wsl.ps1`

---

### 19.3 Ubuntu Install — User Steps

1. Copy `sgml-pipeline-bundle-ubuntu.run` to the Ubuntu machine
2. Open Terminal and run:
   ```bash
   bash ~/sgml-pipeline-bundle-ubuntu.run
   ```
3. Installation completes (~5 min). Browser opens to `http://localhost:8501`.

After install, the `sgml-pipeline` command is available system-wide:
```bash
sgml-pipeline start     # Start app + open browser
sgml-pipeline stop      # Stop
sgml-pipeline status    # Check services
sgml-pipeline diag      # Run ABBYY diagnostics
```

---

### 19.4 Windows Install — User Steps (WSL2)

**Requirements:** Windows 10 version 2004+ (Build 19041+) or Windows 11, Admin account

1. Copy these two files to the Windows machine:
   - `sgml-pipeline-bundle.tar.gz`
   - `scripts/setup_windows_wsl.ps1`

2. Right-click `setup_windows_wsl.ps1` → **Run with PowerShell (as Administrator)**

3. The script:
   - Checks Windows version compatibility
   - Installs WSL2 + Ubuntu 22.04 (if not already present)
   - Extracts and installs the SGML Pipeline inside Ubuntu WSL
   - Creates **Desktop shortcut** "SGML Pipeline"
   - Creates Start Menu entry
   - Launches the app immediately

4. On subsequent uses: **Double-click "SGML Pipeline" on Desktop**

**Note:** If WSL2 is not yet installed, the script enables the Windows feature and **requires a restart**. Run the script again after restarting.

---

### 19.5 Windows Daily Use (after setup)

| Action | Method |
|---|---|
| Start app | Double-click "SGML Pipeline" Desktop shortcut |
| Start app (CMD) | `scripts\launch_sgml_pipeline.bat` |
| Stop app | `scripts\stop_sgml_pipeline.bat` |
| WSL command | `wsl -- sgml-pipeline start` |
| Status | `wsl -- sgml-pipeline status` |
| Diagnostics | `wsl -- sgml-pipeline diag` |
| App URL | `http://localhost:8501` |

---

### 19.6 Files Created

| File | Purpose |
|---|---|
| `scripts/setup_local_ubuntu.sh` | Ubuntu installer — installs everything, creates systemd service + CLI |
| `scripts/setup_windows_wsl.ps1` | Windows installer — sets up WSL2 + Ubuntu + Desktop shortcut |
| `scripts/launch_sgml_pipeline.bat` | Windows launcher — starts app in WSL2, opens browser |
| `scripts/stop_sgml_pipeline.bat` | Windows stop script |
| `scripts/create_portable_bundle.sh` | Creates distributable `.tar.gz` + `.run` files |

---

### 19.7 Why This Works (vs Plexus)

| | Plexus | Business User Machine |
|---|---|---|
| Internet to `account.abbyy.com` | ❌ Blocked by TR network policy | ✅ Open (no firewall block) |
| ABBYY Online License | ❌ Fails on every conversion | ✅ Works every time |
| `/dev/shm` size | ❌ 64 MB (K8s default) | ✅ Typically 8 GB (desktop default) |
| Install complexity | ❌ K8s + ECR + IAM + VPC | ✅ One script, one restart |

Business user deployment **bypasses all Plexus infrastructure issues entirely**.

---

### 19.8 Limitations

- **Windows requires Admin** for WSL2 install (one-time only; subsequent use is non-admin)
- **ABBYY is Linux only** — no native Windows ABBYY without a separate Windows license
- **Bundle is large (~3.5 GB)** — requires broadband to transfer; use a shared drive or S3
- **Online License per seat** — each business user machine connects to `account.abbyy.com` independently; verify seat count with ABBYY support if deploying to many machines

---

## Summary / Current Status (2026-07-13)

- ✅ Image `v0.0.21` built, pushed to ECR, verified locally (engine + fix both work).
- ✅ Code fully supports the EC2 license server path (`ABBYY_LS_HOST` patches all 3 settings files).
- ⚠️ **Plexus still shows the licensing error** because it's deployed WITHOUT a working license server — need Option A (allowlist) or Option B (EC2 + set `ABBYY_LS_HOST`).
- 🚧 **Blocker:** infrastructure — we lack `ec2:RunInstances`. Cloud team (Sai Kiran) engaged to help set up Option A or B.
- 📌 The temp folder warning is unrelated and harmless. ABBYY is correctly baked into `/opt/ABBYY`.
