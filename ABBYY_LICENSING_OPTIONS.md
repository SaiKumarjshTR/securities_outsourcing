# ABBYY FREngine12 Licensing — Problem & Solutions for Lead Review

**Date**: 2026-07-02  
**Context**: SGML Pipeline (Plexus / AWS Deployment)  
**Current Version**: v0.0.18 — LicensingService starts (sidebar green ✅) but PDF→DOCX conversion fails after 2-minute timeout

---

## The Problem — Why Conversion Is Failing

ABBYY FREngine12 uses an **Online Protection License** (Serial: SWAT-1241-1007-4471-4083-6499). This type of license works like a time-limited "lease" model:

1. Every time someone uploads a PDF and clicks **Convert**, the ABBYY engine contacts ABBYY's cloud server at `account.abbyy.com` to **check out a license slot**.
2. After the conversion is done, it **checks the slot back in**.
3. If the engine cannot reach `account.abbyy.com`, it simply refuses to process and returns the error: **"Failed to communicate with the online licensing service"**.

```
What ABBYY needs to do every conversion:
┌───────────────────────┐      HTTPS:443       ┌──────────────────────┐
│  Plexus Container     │ ──────────────────→  │  account.abbyy.com   │
│  (FREngine12 CLI)     │   "Can I use the     │  (ABBYY Cloud)       │
│                       │    license?"         │                      │
│                       │  ←────────────────── │  "Yes, here's a      │
│                       │   "Approved for 60s" │   token"             │
└───────────────────────┘                      └──────────────────────┘
```

**What happens in Plexus**: The container is inside an AWS VPC (Virtual Private Cloud). Thomson Reuters has locked down outbound internet from Plexus containers — the container cannot reach `account.abbyy.com`. So the engine waits 2 minutes for a response that never comes, then fails.

**Why it works locally (WSL)**: On our WSL development machine, traffic goes out through TR's corporate Zscaler proxy which DOES allow `account.abbyy.com`. We already bundled the Zscaler SSL certificates into the container — but that only matters if the container can reach the internet in the first place.

---

## Three Solutions

---

## Option A — Network Allowlist (Plexus/Infra Team)

### What It Is
Simply ask the Thomson Reuters Plexus infrastructure team to **open a hole in the firewall** so our Plexus containers can reach `account.abbyy.com` on port 443 (HTTPS).

### How It Works

```
Before (broken):
┌─────────────────┐         ✗ BLOCKED         ┌──────────────────────┐
│ Plexus Container│  ──────────────────────→  │  account.abbyy.com   │
└─────────────────┘    AWS Network Policy      └──────────────────────┘

After (fixed):
┌─────────────────┐         ✓ ALLOWED         ┌──────────────────────┐
│ Plexus Container│  ──────────────────────→  │  account.abbyy.com   │
└─────────────────┘   (allowlist added)        └──────────────────────┘
```

### What Needs to Change
- **Code changes**: ZERO. Container v0.0.18 is already correct.
- **Rebuild needed**: No.
- **Who does the work**: Plexus infra / cloud networking team.

### The Ticket Request
> "Please allow outbound TCP 443 (HTTPS) from Plexus container pods to:
> - Hostname: `account.abbyy.com`
> - IP: `52.211.188.155` (verify with `nslookup account.abbyy.com`)
> - Purpose: ABBYY FREngine12 license validation (Online License type)
> - Serial: SWAT-1241-1007-4471-4083-6499"

### Pros
| ✅ | Detail |
|---|---|
| Fastest | No code changes, no rebuilds, no new infrastructure |
| Cleanest | The "correct" fix — containers are supposed to have managed internet access |
| Zero maintenance | Nothing new to keep running |

### Cons
| ❌ | Detail |
|---|---|
| Dependent on IT | Requires a ticket and approval from another team |
| Unknown timeline | Could be hours or days depending on process |
| May be rejected | Some environments have strict "no direct internet" policies |

### Estimated Timeline
**If approved**: 1–2 days for infra team to update VPC/SG rules.  
**Overall risk**: Low — this is a standard "egress allowlist" request.

---

## Option B — EC2 License Server (Separate AWS Machine)

### What It Is
Run ABBYY's `LicensingService` on a **separate, small EC2 instance** in the same AWS account. That EC2 HAS internet access (via Internet Gateway in a public subnet). Our Plexus container connects to the EC2 on port 3023 instead of running its own local LicensingService.

This is ABBYY's own **recommended production architecture** (documented in FREngine12AdminGuide.pdf, pages 104–108, "Two-Container Deployment").

### How It Works

```
Current (broken) — container tries to reach internet directly:
┌─────────────────────────────────┐      BLOCKED      ┌──────────────────┐
│ Plexus Container (private subnet│  ─────────────→   │ account.abbyy.com│
│ 10.156.60.x → TGW → TR policy) │                    └──────────────────┘
└─────────────────────────────────┘
  
Proposed (working) — container uses EC2 as license proxy:
┌─────────────────────────────┐  TCP:3023   ┌──────────────────────────────┐   HTTPS:443   ┌──────────────────┐
│ Plexus Container            │ ──────────→ │ EC2 in PUBLIC SUBNET         │ ────────────→               │ account.abbyy.com│
│ (private subnet 10.156.60.x)│             │ subnet-0af01744c5615bbec     │               └──────────────────┘
│ No local LicensingService   │             │ → Internet Gateway (IGW)     │
└─────────────────────────────┘             │ → Direct internet (no TR TGW)│
  [Internal VPC traffic - OK]               └──────────────────────────────┘
```

### Why Public Subnet Matters (Critical)
We investigated the VPC routing. **Private subnet** EC2s (where existing instances `ec2-Ncib`, `ec2-cKIR` live) route all internet traffic through **Transit Gateway (TGW) `tgw-0d3aebf5bb40f0224`** — the same TR corporate network policy that blocks Plexus containers. They cannot reach `account.abbyy.com` either.

**Public subnet** (`subnet-0af01744c5615bbec`) routes `0.0.0.0/0` directly to **Internet Gateway** — unrestricted direct internet. An EC2 here CAN reach `account.abbyy.com`.

### What We Investigated (AWS Account `058264507063`)

| Item | Finding |
|---|---|
| VPC | `vpc-0ae693f600bdcc28e` (tr-vpc-1) |
| Public subnets with IGW | `subnet-0af01744c5615bbec` (us-east-1a), `subnet-0cb2f0078c2bafecb` (us-east-1c), `subnet-08410db7f832f14d5` (us-east-1b) |
| Existing running x86-64 instances | `ec2-Ncib` (t3.nano), `ec2-cKIR` (t3.small) — BUT both in private subnets → TGW |
| Existing ARM64 instance | `TrBastionASG` — SSM Online but ARM64, LicensingService won't run |
| Key pairs | None in account |
| NAT Gateways | None |
| Internet Gateways | `igw-051d5e32b554b95e5` (attached to public subnets only) |

### ❌ Blockers — Cannot Do This Ourselves (Current DataScientist Role)

| Blocker | Detail |
|---|---|
| `ec2:RunInstances` **DENIED** | Cannot launch new EC2 instances |
| `ec2:CreateSecurityGroup` **DENIED** | Cannot create security group for port 3023 |
| `ssm:StartSession` **DENIED** | Cannot shell into any existing instance |
| No SSH key pairs | Cannot SSH to any instance even if open |
| Existing EC2s in wrong subnet | ec2-Ncib/ec2-cKIR are in private subnets → TGW → same block as Plexus |

### ✅ What Infra Team Needs to Do (Option B)

Give this exact specification to your cloud admin / Plexus infra team:

> **Request**: Launch a t3.nano EC2 in the tr-vpc-1 **public subnet** and install ABBYY LicensingService
>
> **EC2 specs**:
> - Type: `t3.nano` (~$4/month)
> - AMI: Amazon Linux 2023 (x86-64)
> - Subnet: `subnet-0af01744c5615bbec` (tr-vpc-1.public.us-east-1a) — needs IGW for internet
> - Security Group inbound: TCP port 3023 from `10.156.60.0/22` (Plexus pod CIDR)
> - Security Group outbound: TCP 443 to `0.0.0.0/0` (already default)
> - Instance Profile: with `AmazonSSMManagedInstanceCore` (for SSM access without SSH key)
>
> **After launch**: We have a setup script ready (`scripts/setup_ec2_license_server.sh`) to install LicensingService

### What We Have Ready (Code is Complete)
- **v0.0.19** `docker_start.sh` — supports `ABBYY_LS_HOST=<ec2-ip>` env var
- **`scripts/setup_ec2_license_server.sh`** — installs LicensingService, libs, license files, creates systemd service
- **`scripts/build_push_v014.sh`** — builds and pushes v0.0.19

After infra team sets up EC2, the only steps remaining:
```bash
# Step 1 - Install on EC2 (takes ~5 min)
bash scripts/setup_ec2_license_server.sh ec2-user@<ec2-private-ip>

# Step 2 - Build and push v0.0.19
bash scripts/build_push_v014.sh C303180 <jfrog_token>

# Step 3 - Redeploy in Plexus with env var:
# ABBYY_LS_HOST = <ec2-private-ip>
```

### Pros
| ✅ | Detail |
|---|---|
| No ABBYY involvement | We control the fix entirely (once EC2 exists) |
| Works immediately | Once EC2 is set up, just redeploy |
| ABBYY-recommended | Their documented production pattern |
| Reliable | EC2 with systemd auto-restarts LicensingService |

### Cons
| ❌ | Detail |
|---|---|
| Requires infra team | Need cloud admin to launch EC2 (we lack `ec2:RunInstances`) |
| New infrastructure | Small EC2 to manage, patch, monitor |
| Small cost | ~$4/month for t3.nano |
| Single point of failure | If EC2 goes down, all conversions fail |

### Estimated Timeline
**Infra team EC2 launch**: 1–2 hours of their time.
**Our setup + deploy after that**: ~2 hours.
**Overall risk**: Low — architecture tested locally, code ready.

---

## Option C — ABBYY Local License (Best for Long-Term Production)

### What It Is
Ask ABBYY to convert our license from **Online Protection** type to **Local License** type. A Local License is a cryptographic file that is activated **once at Docker build time** and then works **forever offline** — no internet access needed at runtime, ever.

### How It Works

```
Online License (current):
  docker build  →  Container starts  →  Every conversion calls account.abbyy.com  ← PROBLEM

Local License (requested):
  docker build  →  [Activate once, internet needed here only]  →  Container starts
                                                                         ↓
                                                         Every conversion reads LOCAL file only
                                                         NO internet needed at runtime  ← FIXED
```

The license file (`.ABBYY.LocalLicense`) is tied to specific machine/container characteristics and expires on a date set by ABBYY. Once activated, it's stored in `/var/lib/ABBYY/SDK/12/Licenses/` inside the image — fully self-contained.

### What Needs to Change

1. **Email ABBYY Support** (template in `ABBYY_LOCAL_LICENSE_REQUEST.md`):
   - Serial: `SWAT-1241-1007-4471-4083-6499`
   - Customer Project ID: `3AA4tXXQRiDuRh9Hh7P9`
   - Request: Convert to Local License, or provide a Local License file

2. **Update Dockerfile** to activate license at build time (ABBYY provides exact command with Local License).

3. **Simplify `docker_start.sh`** — remove LicensingService startup entirely (local licenses may not need it, or need a simplified mode).

4. **Rebuild and push** the container.

### Pros
| ✅ | Detail |
|---|---|
| Permanent fix | No internet dependency at runtime, forever |
| No extra infra | No EC2, no network rules, no IT tickets |
| Simpler container | No LicensingService daemon needed inside container |
| Production-grade | Best practice for air-gapped / restricted environments |

### Cons
| ❌ | Detail |
|---|---|
| ABBYY-dependent | Need ABBYY's approval — they control license types |
| Unknown timeline | Could be days or weeks depending on ABBYY's process |
| May not be possible | Depends on our contract/license agreement with ABBYY |
| Expiry management | Local licenses expire; need process to renew and rebuild |

### Estimated Timeline
**ABBYY response**: 2–5 business days.  
**Overall risk**: Medium — depends entirely on ABBYY's willingness to convert.

---

## Summary Table

| | Option A: Allowlist | Option B: EC2 Server | Option C: Local License |
|---|---|---|---|
| **Who does the work** | Plexus infra team | Infra team (EC2) + us (code) | ABBYY + us |
| **Code changes** | None | v0.0.19 (done ✅) | Yes (Dockerfile) |
| **New infrastructure** | None | 1 × t3.nano EC2 in public subnet | None |
| **Can we do it alone?** | ❌ Needs infra team | ❌ Needs infra team for EC2 | ❌ Needs ABBYY approval |
| **Timeline** | 1–2 days | Infra: 1–2h, then us: 2h | 2–5 days (ABBYY) |
| **Runtime internet needed** | Yes (container) | Yes (EC2 only) | **No** |
| **Risk** | Low | Low | Medium |
| **Recommended for** | Quick unblock | Immediate workaround | Long-term production |

---

## Recommendation

**Do A and C in parallel immediately. Use B as the backup if A is delayed.**

1. **Today**: Raise the Plexus infra allowlist ticket (Option A). Zero effort, highest ROI if approved quickly.
2. **Today**: Email ABBYY to start the Local License conversion process (Option C). This is the right long-term fix.
3. **If A is not approved within 1 day**: Spin up EC2 and deploy v0.0.19 (Option B). This unblocks the demo within hours.

---

## Technical Reference

| Item | Value |
|---|---|
| ABBYY Serial | SWAT-1241-1007-4471-4083-6499 |
| Customer Project ID | 3AA4tXXQRiDuRh9Hh7P9 |
| License Type | Online Protection (current) |
| ABBYY License Server | account.abbyy.com (port 443 HTTPS) |
| LicensingService internal port | 3023 (TCP) |
| ECR Registry | 127288631409.dkr.ecr.us-east-1.amazonaws.com/... |
| Current deployed version | v0.0.18 (LicensingService starts ✅, conversion fails ❌) |
| Ready-to-deploy version | v0.0.19 (supports `ABBYY_LS_HOST` env var for Option B) |
| ABBYY support email | support@abbyy.com |
| ABBYY support portal | https://support.abbyy.com |
