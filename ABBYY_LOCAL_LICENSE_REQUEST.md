# Email to ABBYY Support — Request Local License for Air-Gapped Docker Deployment

**To**: dennis.lappert@abbyy.com; larysa.lototska@abbyy.com  
**Subject**: FREngine12 Linux — Request Local License (.ABBYY.LocalLicense) for air-gapped AWS deployment

---

Hi Dennis, Hi Larysa,

We are deploying **ABBYY FineReader Engine 12 for Linux** in a Docker container on AWS (Plexus/EKS) and need help switching to a Local License that works without internet access at runtime.

**License details**
- Serial: `SWAT-1241-1007-4471-4083-6499`
- Customer Project ID: `3AA4tXXQRiDuRh9Hh7P9`
- Current license type: **Online Protection** (`.ABBYY.ActivationToken`)
- Platform: Linux x86-64 (Docker container, AWS EKS / Plexus)

**The issue**  
Our AWS container environment does not allow outbound connections to `account.abbyy.com:443` due to corporate network policy. Every conversion fails with:
> *"Failed to communicate with the online licensing service"*

**What we need**  
Per the ABBYY official Docker documentation ([docs.abbyy.com](https://docs.abbyy.com/fine-reader/engine/distribution/distribution-linux/running-abbyy-finereader-engine-12-inside-a-docker-container-2)), a **Local License** (`.ABBYY.LocalLicense`) is activated at Docker build time and requires no internet at runtime — which is perfect for our air-gapped environment.

**Our request**  
Could you please provide a **Local License** (`.ABBYY.LocalLicense` file + password) for our serial `SWAT-1241-1007-4471-4083-6499`? We have the ABBYY FRE12 installer and can run the activation at Docker build time following your documentation.

If the Online Protection trial cannot be converted to a Local License, could you advise on:
1. Whether a new Local License can be issued for our trial, or
2. An alternative mechanism to support air-gapped Docker deployment?

Thank you for your help.

Best regards,  
Saikumar Bhukya  
Thomson Reuters — Securities Outsourcing

---

# Email to ABBYY Support — Original (Online License issue)

---

Hi Dennis, Hi Larysa,

Thank you for sharing **ABBYY FineReader Engine 12 for Linux** — we are using this in a Docker container on AWS (Plexus) and need your help with a licensing issue.

**License details**
- Serial: `SWAT-1241-1007-4471-4083-6499`
- Customer Project ID: `3AA4tXXQRiDuRh9Hh7P9`
- License type: Online Protection
- Platform: Linux x86-64 (Debian Docker container)

**The issue**  
Every conversion attempt fails with:
> *"Failed to communicate with the online licensing service"*

Our AWS container (Plexus) environment does not allow outbound connections to `account.abbyy.com:443` due to network policy. The LicensingService starts correctly and the setup works on our local machine — the only problem is the blocked outbound internet in production (Plexus).

**What we need**  
We need a licensing solution that works in a container environment without outbound internet access. Could you please suggest the best approach for our case and advise on next steps?

Thank you.

Best regards,  
Saikumar B

---

# Email to TR Network/Security Team

**To**: network-security@thomsonreuters.com *(replace with actual team alias)*  
**Subject**: Plexus — Outbound allowlist request for ABBYY license validation (`account.abbyy.com:443`)

---

Hi Team,

Hope you're doing well. We have a request for a narrow outbound network exception for our Plexus container deployment and would appreciate your guidance on feasibility.

**What we need allowed**
- **Destination**: `account.abbyy.com` (IP: `52.211.188.155`)
- **Port / Protocol**: `443` (HTTPS)
- **Source**: Plexus container pods (namespace: *(add your Plexus namespace)*)
- **Direction**: Outbound only

**Why**  
We are running ABBYY FineReader Engine 12 inside a Plexus container for PDF-to-DOCX conversion. ABBYY uses an Online License model — on every conversion, the engine contacts `account.abbyy.com:443` to validate the license lease. Without this connection, every conversion fails with:
> *"Failed to communicate with the online licensing service"*

We have confirmed with ABBYY support (Larysa Lototska) that:
- This is the **only** supported licensing mechanism for FREngine12 on Linux/Docker/Kubernetes
- The connection carries **no document data** — it is exclusively for license sync (page counter update and license file refresh)
- Traffic volume is **minimal** (a few KB per conversion)

This is a single, specific HTTPS endpoint — no broad internet access is needed.

**Asking**  
Could you please advise if it is feasible to allowlist `account.abbyy.com:443` for Plexus container outbound traffic? If there is a standard process or form to raise this formally, we are happy to follow it.

Thank you for your time and support.

Best regards,  
Saikumar B

---

# Reply to ABBYY (Larysa)

**To**: larysa.lototska@abbyy.com  
**Subject**: Re: FREngine12 Linux — Online License failing in AWS container — Need licensing guidance

---

Hi Larysa,

Thank you for the clear explanation — this is very helpful.

We understand that the Online License requires an active connection to `*.abbyy.com:443` and that this is the only supported mechanism for FREngine12 on Linux/Docker. We are actively working with our TR network team to get the outbound allowlist exception approved.

Your clarification that **no document data is transmitted** — only license sync — will help us make the case to our security team. We will keep you posted on the progress.

Thank you again for your support.

Best regards,  
Saikumar B

---

# Follow-up to Surabhi Chaudhary (TR Network Team)

**To**: surabhi.chaudhary@thomsonreuters.com  
**Subject**: Re: ABBYY FREngine12 — Network allowlist needed for Plexus deployment

---

Hi Surabhi,

Thank you for connecting with us. Let me walk you through the issue in detail so you have everything needed to assess the request.

**What we are building**  
We are deploying an AI-based document processing solution on **Plexus** that uses **ABBYY FineReader Engine 12 for Linux** to convert PDFs to DOCX. The application runs as a Docker container on Plexus (AWS).

**The problem**  
ABBYY FineReader Engine 12 uses an **Online License** model. On every conversion, the engine's internal `LicensingService` contacts ABBYY's licensing server at `account.abbyy.com:443` to validate the license lease. Without this connection, the conversion fails immediately with:
> *"Failed to communicate with the online licensing service"*

Our Plexus container environment does not allow outbound internet connections, so the engine cannot reach `account.abbyy.com`. The setup is fully working on our local development machine where this endpoint is reachable.

**Confirmation from ABBYY**  
We have verified this with ABBYY support (Larysa Lototska). They confirmed:
- Online licensing via `account.abbyy.com:443` is the **only supported mechanism** for FREngine12 on Linux/Docker/Kubernetes — there is no offline or air-gapped alternative
- **No document data is transmitted** through this connection — it carries only license sync data (page counter and license file updates)
- Traffic is **minimal** — a few KB per conversion, purely for licensing purposes

**What we are requesting**  
A narrow outbound allowlist exception for the Plexus container pods:

| Detail | Value |
|---|---|
| Destination hostname | `account.abbyy.com` |
| Destination IP | `52.211.188.155` |
| Port / Protocol | `443` (HTTPS) |
| Source | Plexus container pods *(we can provide the exact namespace/CIDR if needed)* |
| Direction | Outbound only |

This is a single, specific endpoint — not a broad internet access request.

Could you please advise if this exception is feasible, and let us know if there is a formal process or form we need to follow to raise it?

Happy to jump on a call if that would be easier.

Thank you, Surabhi.

Best regards,  
Saikumar B
