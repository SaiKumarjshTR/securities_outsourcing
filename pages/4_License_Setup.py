"""
ABBYY FREngine 12 — License Setup
===================================
Allows business users to configure their own ABBYY FREngine 12 license:

  Option A : Upload a pre-existing .ActivationToken file supplied by ABBYY
  Option B : Online activation using Serial Number + Customer Project ID
             (connects to account.abbyy.com:443 — internet required)

After any license change the LicensingService daemon is restarted automatically.

ABBYY license credentials the user will need:
  - Serial Number       (e.g. SWAT12411007447140836499)
  - Customer Project ID (e.g. TR_Securities_SGML)
  - (Option A only) .ActivationToken file, or
  - (Option B only) Online License Path + Password if prompted by ABBYY
"""

import os
import re
import subprocess
import time
from pathlib import Path

import streamlit as st

# ── Runtime paths (override via env vars in Docker / local install) ────────────
LICENSES_DIR = Path(os.getenv("ABBYY_LICENSES_DIR", "/var/lib/ABBYY/SDK/12/Licenses"))
LS_BIN       = Path("/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService")
ACTIVATE_SH  = Path(os.getenv("ABBYY_ACTIVATE_SH",
                               "/opt/ABBYY/FREngine12/activatefre.sh"))
ABBYY_CLI    = Path(os.getenv("ABBYY_CLI",
                               "/opt/ABBYY/FREngine12/Samples/CommandLineInterface"
                               "/CommandLineInterface"))

# ── Page layout ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="License Setup — SGML Pipeline",
    page_icon="🔑",
    layout="wide",
)

st.title("🔑 ABBYY FREngine 12 — License Setup")
st.markdown(
    "Configure the ABBYY FREngine 12 license so PDF → DOCX conversion works "
    "on this machine.  Each business unit / server needs its own valid license."
)
st.markdown("---")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_tokens() -> list[str]:
    """Return sorted list of .ActivationToken filenames in the licenses dir."""
    if not LICENSES_DIR.exists():
        return []
    return sorted(f.name for f in LICENSES_DIR.glob("*.ActivationToken"))


def _ls_running() -> bool:
    r = subprocess.run(["pgrep", "-f", "LicensingService"], capture_output=True)
    return r.returncode == 0


def _restart_ls() -> tuple[bool, str]:
    """Kill + restart the local LicensingService daemon.  Returns (ok, message)."""
    # Stop any running instance
    subprocess.run(["pkill", "-f", "LicensingService"], capture_output=True)
    time.sleep(2)

    if not LS_BIN.exists():
        return False, f"LicensingService binary not found at `{LS_BIN}`"

    env = {**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8"}
    subprocess.Popen(
        [str(LS_BIN), "/start"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
        env=env,
    )
    time.sleep(3)

    chk = subprocess.run(["pgrep", "-f", "LicensingService"], capture_output=True)
    if chk.returncode == 0:
        return True, "LicensingService restarted successfully ✅"
    return False, "LicensingService could not be started — check system logs."


def _safe_str(value: str, allow: str = r"^[A-Za-z0-9_.\-]+$") -> bool:
    """Return True if value matches the safe character pattern."""
    return bool(value and re.match(allow, value.strip()))


# ── 1. Current Status ─────────────────────────────────────────────────────────
st.subheader("📊 Current License Status")

tokens  = _get_tokens()
running = _ls_running()
cli_ok  = ABBYY_CLI.exists()

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("**License Token Files**")
    if tokens:
        st.success(f"✅ {len(tokens)} token file(s) installed")
        for t in tokens:
            st.caption(f"• `{t}`")
        st.caption(f"Location: `{LICENSES_DIR}`")
    else:
        st.error("❌ No token files found")
        st.caption(f"Expected location: `{LICENSES_DIR}`")

with c2:
    st.markdown("**LicensingService**")
    if running:
        st.success("✅ Running (port 3023)")
    else:
        st.error("❌ Not running")
        st.caption("Use **Service Control** below to restart it.")

with c3:
    st.markdown("**FREngine12 CLI**")
    if cli_ok:
        st.success("✅ Found")
        st.caption(f"`{ABBYY_CLI.name}`")
    else:
        st.error("❌ CLI not found")
        st.caption(f"Expected: `{ABBYY_CLI}`")

if tokens and running and cli_ok:
    st.success("🎉 **ABBYY is fully licensed and ready.** PDF → DOCX conversion is available.")
elif tokens and not running:
    st.warning("⚠️ License token is installed but LicensingService is not running. "
               "Use **Service Control** below to restart.")
else:
    st.info("👇 Choose one of the options below to install your ABBYY license.")

st.markdown("---")

# ── 2. Option A — Upload .ActivationToken file ────────────────────────────────
st.subheader("Option A — Upload License Token File")
st.markdown(
    "Use this option if **ABBYY has provided you a `.ActivationToken` file** "
    "(typically delivered via email or download after your purchase). "
    "No internet connection required — just upload the file."
)

with st.expander("📂 Upload .ActivationToken", expanded=not bool(tokens)):
    uploaded = st.file_uploader(
        "Select your `.ActivationToken` file",
        key="token_uploader",
        help="The file name should end in `.ActivationToken` — provided by ABBYY.",
    )

    if uploaded is not None:
        fname = uploaded.name
        fbytes = uploaded.read()

        # Basic validation — file name must end with .ActivationToken
        if not fname.endswith(".ActivationToken"):
            st.error("❌ File must end with `.ActivationToken`.  "
                     "Please upload the correct file from ABBYY.")
        elif len(fbytes) < 50:
            st.error("❌ File looks too small — it may be corrupted or empty.")
        else:
            st.info(f"Ready to install: **`{fname}`** ({len(fbytes):,} bytes)")

            if st.button("💾 Install Token File & Restart Service",
                         type="primary", key="btn_install_token"):
                try:
                    LICENSES_DIR.mkdir(parents=True, exist_ok=True)
                    target = LICENSES_DIR / fname
                    target.write_bytes(fbytes)
                    st.success(f"✅ Token saved → `{target}`")
                except OSError as e:
                    st.error(f"❌ Could not write file: {e}")
                    st.stop()

                with st.spinner("Restarting LicensingService..."):
                    ok, msg = _restart_ls()
                if ok:
                    st.success(msg)
                else:
                    st.warning(f"⚠️ {msg}")
                st.rerun()

st.markdown("---")

# ── 3. Option B — Online activation ──────────────────────────────────────────
st.subheader("Option B — Online Activation via Serial Number")
st.markdown(
    "Use this if you have an ABBYY **Serial Number** and **Customer Project ID** "
    "but no `.ActivationToken` file yet.  "
    "The activation wizard will connect to `account.abbyy.com:443` to generate "
    "your token automatically.  **Internet access is required.**"
)

if not ACTIVATE_SH.exists():
    st.error(f"❌ Activation script not found at `{ACTIVATE_SH}`.  "
             "Ensure the ABBYY FREngine12 bundle is installed correctly.")
else:
    with st.expander("🔑 Enter Activation Credentials", expanded=not bool(tokens)):
        with st.form("online_activation_form", clear_on_submit=False):
            st.markdown("**Required credentials** (from your ABBYY purchase order / support email):")

            col_s, col_p = st.columns(2)
            with col_s:
                serial = st.text_input(
                    "Serial Number  *",
                    placeholder="e.g.  SWAT12411007447140836499",
                    help="Uppercase letters and digits — provided by ABBYY.",
                )
            with col_p:
                project_id = st.text_input(
                    "Customer Project ID  *",
                    placeholder="e.g.  TR_Securities_SGML",
                    help="Your ABBYY Customer Project ID.",
                )

            st.markdown("**Optional** (only required if your license uses a custom path / password):")

            col_lp, col_pw = st.columns(2)
            with col_lp:
                lic_path = st.text_input(
                    "Online License Path",
                    placeholder="leave blank for standard activation",
                    help="Only needed for custom license configurations.",
                )
            with col_pw:
                lic_password = st.text_input(
                    "License Password",
                    type="password",
                    placeholder="leave blank if not required",
                    help="Password for the online license (usually not required).",
                )

            st.markdown("---")
            submitted = st.form_submit_button("🚀 Activate Now", type="primary")

        if submitted:
            serial     = serial.strip()
            project_id = project_id.strip()
            lic_path   = lic_path.strip()

            # ── Input validation (prevents shell injection) ────────────────
            errors = []
            if not serial:
                errors.append("Serial Number is required.")
            elif not _safe_str(serial):
                errors.append("Serial Number must contain only letters, digits, hyphens, underscores, or dots.")
            if not project_id:
                errors.append("Customer Project ID is required.")
            elif not _safe_str(project_id):
                errors.append("Customer Project ID must contain only letters, digits, hyphens, underscores, or dots.")
            if lic_path and not _safe_str(lic_path, r"^[A-Za-z0-9_./:@\-]+$"):
                errors.append("License Path contains invalid characters.")

            if errors:
                for err in errors:
                    st.error(f"❌ {err}")
            else:
                # Build the command — all arguments are validated strings
                cmd = [
                    "bash", str(ACTIVATE_SH),
                    "--serial-number",   serial,
                    "--project-id",      project_id,
                    "--skip-local-service-installation",
                ]
                if lic_path:
                    cmd += ["--license-path", lic_path]
                if lic_password.strip():
                    cmd += ["--license-password", lic_password.strip()]

                st.info("Connecting to `account.abbyy.com` for activation…")
                st.code(" ".join(
                    f"--license-password ****" if p == lic_password.strip() and lic_password.strip()
                    else p
                    for p in cmd
                ), language="bash")

                with st.spinner("Activating license (may take 30–60 s)…"):
                    try:
                        env = {**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8"}
                        result = subprocess.run(
                            cmd,
                            capture_output=True,
                            text=True,
                            timeout=120,
                            env=env,
                        )
                        output = (result.stdout + result.stderr).strip()

                        if result.returncode == 0:
                            st.success("✅ Activation successful!")
                            st.code(output)
                            with st.spinner("Restarting LicensingService…"):
                                ok2, msg2 = _restart_ls()
                            if ok2:
                                st.success(msg2)
                            else:
                                st.warning(f"⚠️ {msg2}")
                            st.rerun()
                        else:
                            st.error(f"❌ Activation failed (exit code {result.returncode})")
                            st.code(output)
                            st.info(
                                "Possible causes:\n"
                                "• Serial Number or Project ID incorrect\n"
                                "• No internet access to `account.abbyy.com:443`\n"
                                "• License already activated on another machine\n"
                                "Contact ABBYY support if the problem persists."
                            )
                    except subprocess.TimeoutExpired:
                        st.error(
                            "❌ Activation timed out (120 s).  "
                            "Check that this machine can reach `account.abbyy.com:443`."
                        )
                    except Exception as exc:
                        st.error(f"❌ Unexpected error: {exc}")

st.markdown("---")

# ── 4. Service Control ────────────────────────────────────────────────────────
st.subheader("⚙️ Service Control")
st.markdown("Restart the ABBYY LicensingService daemon after any license change.")

col_btn, col_note = st.columns([1, 3])
with col_btn:
    if st.button("🔄 Restart LicensingService", use_container_width=True,
                 key="btn_restart_ls"):
        with st.spinner("Restarting LicensingService…"):
            ok3, msg3 = _restart_ls()
        if ok3:
            st.success(msg3)
        else:
            st.error(f"❌ {msg3}")
        st.rerun()

with col_note:
    st.caption(
        "The LicensingService daemon must be running for ABBYY to accept conversion jobs. "
        "If PDF → DOCX conversions are failing, try restarting the service here first."
    )

st.markdown("---")

# ── 5. Help / FAQ ─────────────────────────────────────────────────────────────
with st.expander("❓ Help & FAQ", expanded=False):
    st.markdown("""
**Where do I get the Serial Number and Project ID?**
ABBYY sends these in your purchase confirmation email or license delivery email.
Contact your ABBYY account manager or TR procurement if you don't have them.

**What is the `.ActivationToken` file?**
It is an XML file generated by ABBYY that encodes your license grant.
ABBYY may provide it directly if your environment has no internet access (offline activation).

**Can multiple users share one license?**
An ABBYY FREngine 12 seat license typically allows one concurrent conversion process.
For high-volume / concurrent usage, contact ABBYY for a multi-seat license.

**My activation failed — what do I do?**
1. Verify Serial Number and Project ID (copy-paste from the ABBYY email — no leading/trailing spaces).
2. Check internet access: `curl -I https://account.abbyy.com` should return `200 OK`.
3. If your network uses a corporate proxy / Zscaler, TR CA certificates must be installed
   (Dockerfile already bundles them — they are in `abbyy_bundle/tr_certs/`).
4. Contact ABBYY support at `support@abbyy.com` with your Serial Number.

**Where is the license stored on disk?**
All token files live in `/var/lib/ABBYY/SDK/12/Licenses/`.
They are named `<something>.ActivationToken`.

**Can I back up my license?**
Yes — copy the `.ActivationToken` file to a safe location.
If you reinstall, use **Option A** to restore it.
""")
