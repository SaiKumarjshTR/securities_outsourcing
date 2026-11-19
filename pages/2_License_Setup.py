"""
ABBYY FREngine 12 — License Setup
===================================
Allows business users to configure their own ABBYY FREngine 12 license:

  Option A : Upload a pre-existing .ActivationToken file supplied by ABBYY
  Option B : Online activation using Serial Number + Customer Project ID
             (connects to account.abbyy.com:443 — internet required)

After any license change the LicensingService daemon is restarted automatically.

ABBYY license credentials the user will need:
  - Serial Number       (from the ABBYY license email)
  - Customer Project ID (from the ABBYY license email)
  - (Option A only) .ActivationToken file, or
  - (Option B only) Online License Path + Password if prompted by ABBYY

Default license: when ABBYY_SERIAL_NUMBER and ABBYY_CUSTOMER_PROJECT_ID are set in the
deployment, the server activates that production license by itself at start-up and this
page only tells users that nothing needs to be done (see abbyy_license.py). Options A and
B then act as an optional override that lasts until the server restarts.
"""

import os
import re
import subprocess
import sys
import time
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import abbyy_license  # noqa: E402

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
    page_title="ABBYY FREngine12 License",
    page_icon="🔑",
    layout="wide",
)

st.title("🔑 ABBYY FREngine 12 — License Setup")

_view = abbyy_license.describe()
# True when a default (production) license is configured for this server in the deployment.
_default_mode = _view.default_configured and _view.kind != "external"

if st.session_state.pop("_reset_confirm", False):
    st.session_state["confirm_override"] = False

_flash = st.session_state.pop("_license_flash", None)
if _flash:
    getattr(st, _flash[0])(_flash[1])

_BANNER_ICON = {"success": "✅", "info": "ℹ️", "warning": "⚠️", "error": "🚫"}


def _render_banner(view) -> bool:
    msg = abbyy_license.ui_message(view)
    if not msg:
        return False
    level, title, body = msg
    icon = "⏳" if view.kind in ("activating", "retrying") else _BANNER_ICON[level]
    getattr(st, level)(f"**{title}**\n\n{body}", icon=icon)
    return True


def _flash_result(outcome, ok_text: str, fail_text: str) -> None:
    if outcome.ok:
        st.session_state["_license_flash"] = ("success", ok_text)
    else:
        st.session_state["_license_flash"] = ("error", f"{fail_text} {abbyy_license.reason_text(outcome.reason)}")


def _banner_actions(view) -> None:
    if not view.default_configured:
        return
    if view.kind == "default_failed":
        if st.button("🔁 Try the default license again", key="btn_retry_default"):
            with st.spinner("Contacting ABBYY (this can take up to a minute)…"):
                outcome = abbyy_license.activate_default()
            _flash_result(outcome, "The default ABBYY license is now active.",
                          "The default license still could not be set up.")
            st.rerun()
    elif view.kind == "custom_active":
        if st.session_state.get("_confirm_switch"):
            st.warning("Switching back takes about 10 seconds. PDF conversions that start meanwhile wait for it. "
                       "If a conversion is running right now, you will be asked to try again in a minute.")
            col_yes, col_no, _ = st.columns([1, 1, 3])
            if col_yes.button("Yes, switch back", type="primary", key="btn_switch_yes"):
                st.session_state["_confirm_switch"] = False
                with st.spinner("Switching back to the default license…"):
                    outcome = abbyy_license.switch_back_to_default()
                _flash_result(outcome, "Switched back: the default ABBYY license is active again.",
                              "Could not switch back; the license you set up is still in use.")
                st.rerun()
            if col_no.button("Cancel", key="btn_switch_no"):
                st.session_state["_confirm_switch"] = False
                st.rerun()
        elif st.button("↩️ Switch back to the default license", key="btn_switch"):
            st.session_state["_confirm_switch"] = True
            st.rerun()


if _view.kind in ("activating", "retrying"):
    @st.fragment(run_every="3s")
    def _live_banner():
        current = abbyy_license.describe()
        if current.kind not in ("activating", "retrying"):
            st.rerun()  # finished: redraw the whole page
        _render_banner(current)

    _live_banner()
    _has_banner = True
else:
    _has_banner = _render_banner(_view)
    _banner_actions(_view)

if _has_banner:
    st.markdown("Use this page only if you need to work with a different ABBYY license.")
else:
    st.markdown(
        "Configure the ABBYY FREngine 12 license so PDF → DOCX conversion works "
        "on this machine.  Each business unit / server needs its own valid license."
    )
_uptime = abbyy_license.uptime_text()
if _uptime:
    st.caption(_uptime)
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

    env = {**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8", "LD_LIBRARY_PATH": "/usr/local/lib/ABBYY/SDK/12/Licensing:" + os.environ.get("LD_LIBRARY_PATH", "")}
    subprocess.Popen(
        [str(LS_BIN), "/standalone"],
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


def _finish_custom_activation(res, wrong_input_hint: str, staged_token=None) -> None:
    """Show the result of a user-initiated activation (success reloads the page)."""
    if res.ok:
        note = ("ABBYY confirmed that the license works." if res.verified
                else "Convert a PDF on the main page to confirm that it works.")
        st.session_state["_license_flash"] = ("success", f"✅ Activation successful. {note}")
        st.session_state["_reset_confirm"] = True
        st.rerun()
    if staged_token is not None:
        Path(staged_token).unlink(missing_ok=True)  # do not leave a rejected token behind
    if res.reason in ("converting", "busy"):
        st.warning(f"⏳ {abbyy_license.reason_text(res.reason)} Nothing was changed.")
        return
    st.error(f"❌ Activation failed. {abbyy_license.reason_text(res.reason)}")
    if res.output:
        st.code(res.output)
    if res.reason == "refused":
        st.info(
            "ABBYY's server was reached and answered, but it did not grant the activation. Most often:\n"
            "• The number of activations allowed for this serial number is used up (ABBYY counts every new container "
            "or server as a new computer)\n"
            "• The serial number cannot be activated this way (for example, it is meant for Online licensing)\n"
            "Ask ABBYY to check the activations of the serial number."
        )
    elif res.reason in ("invalid_serial", "network", "limit", "unknown", "config"):
        st.info(
            "Possible causes:\n"
            f"• {wrong_input_hint}\n"
            "• No internet access to ABBYY's servers (see **Test connection to ABBYY** further down)\n"
            "• License already activated on another machine\n"
            "Contact ABBYY support if the problem persists."
        )
    if res.reason == "timeout":
        st.info("Check that this machine can reach `account.abbyy.com:443`.")


# ── 1. Current Status ─────────────────────────────────────────────────────────
st.subheader("📊 Current License Status")

tokens  = _get_tokens()
running = _ls_running()
cli_ok  = ABBYY_CLI.exists()

c1, c2, c3 = st.columns(3)

_LICENSE_KINDS_OK = {"default_active": "✅ Default production license is active",
                     "custom_active": "✅ A license set up from this page is active"}

with c1:
    if _default_mode or _view.kind in _LICENSE_KINDS_OK:
        # Never list file names here: they can contain the license identifiers.
        st.markdown("**ABBYY License**")
        if _view.kind in _LICENSE_KINDS_OK:
            st.success(_LICENSE_KINDS_OK[_view.kind])
        elif _view.kind in ("activating", "retrying"):
            st.info("⏳ Being set up automatically…")
        else:
            st.error("❌ Not active")
    else:
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

if _has_banner:
    if _view.kind in ("default_failed", "misconfigured"):
        st.info("👇 You can use your own ABBYY license with one of the options below.")
elif tokens and running and cli_ok:
    st.success("🎉 **ABBYY is fully licensed and ready.** PDF → DOCX conversion is available.")
elif tokens and not running:
    st.warning("⚠️ License token is installed but LicensingService is not running. "
               "Use **Service Control** below to restart.")
else:
    st.info("👇 Choose one of the options below to install your ABBYY license.")

st.markdown("---")

# ── Use a different license: confirmation (only when a license is already in place) ─
_needs_confirm = _default_mode or _view.kind == "custom_active"
_override_ok = True
if _needs_confirm:
    st.subheader("Use a different ABBYY license")
    if _view.kind in ("activating", "retrying"):
        _override_ok = False
        st.info("Please wait until the default license has been set up before changing it.")
    else:
        if _default_mode:
            st.markdown(
                "Only do this if you have your own ABBYY license. It replaces the default license "
                "**for everyone using this server**, and the default comes back when the server restarts."
            )
            _label = ("I understand that this changes the license for everyone using this server "
                      "until it restarts.")
        else:
            _label = "I understand that this replaces the license currently in use on this server."
        _override_ok = st.checkbox(_label, key="confirm_override")
        if not _override_ok:
            st.caption("Tick the box to enable the activation buttons below.")
    st.markdown("---")

_expand_forms = ((not _has_banner) and (not tokens)) or _view.kind in ("default_failed", "misconfigured")

# ── 2. Option A — Upload .ActivationToken file ────────────────────────────────
st.subheader("Option A — Upload License Token File")
st.markdown(
    "Use this option if **ABBYY has provided you a `.ActivationToken` file** "
    "(typically delivered via email or download after your purchase). "
    "No internet connection required — just upload the file."
)

with st.expander("📂 Upload .ActivationToken", expanded=_expand_forms):
    st.markdown(
        "For an **Online License Key**, ABBYY's licensing service requires the token to be "
        "registered via a real activation call (not just copied to disk) — this needs your "
        "**Customer Project ID** and **License Password** from the same ABBYY email."
    )
    uploaded = st.file_uploader(
        "Select your `.ActivationToken` file",
        key="token_uploader",
        help="The file name should end in `.ActivationToken` — provided by ABBYY.",
    )

    col_pid, col_pw = st.columns(2)
    with col_pid:
        token_project_id = st.text_input(
            "Customer Project ID  *",
            placeholder="your Customer Project ID (from the ABBYY email)",
            key="token_project_id",
        )
    with col_pw:
        token_password = st.text_input(
            "License Password  *",
            type="password",
            placeholder="from the ABBYY activation email",
            key="token_password",
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

            if st.button("💾 Install & Activate Token",
                         type="primary", key="btn_install_token", disabled=not _override_ok):
                project_id = token_project_id.strip()
                password = token_password.strip()

                errors = []
                if not project_id:
                    errors.append("Customer Project ID is required.")
                elif not _safe_str(project_id):
                    errors.append("Customer Project ID must contain only letters, digits, hyphens, underscores, or dots.")
                if not password:
                    errors.append("License Password is required.")

                if errors:
                    for err in errors:
                        st.error(f"❌ {err}")
                    st.stop()

                try:
                    LICENSES_DIR.mkdir(parents=True, exist_ok=True)
                    target = LICENSES_DIR / fname
                    target.write_bytes(fbytes)
                    st.success(f"✅ Token saved → `{target}`")
                except OSError as e:
                    st.error(f"❌ Could not write file: {e}")
                    st.stop()

                cmd = [
                    "bash", str(ACTIVATE_SH),
                    "--project-id",      project_id,
                    "--license-path",    str(target),
                    "--license-password", password,
                    "--skip-local-service-installation",
                ]
                _note, _cmd_box = st.empty(), st.empty()
                _note.info("Registering token with ABBYY activation service…")
                _cmd_box.code(" ".join("--license-password ****" if p == password else p for p in cmd),
                              language="bash")

                with st.spinner("Activating license (may take 30–60 s)…"):
                    try:
                        env = {**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8", "LD_LIBRARY_PATH": "/usr/local/lib/ABBYY/SDK/12/Licensing:" + os.environ.get("LD_LIBRARY_PATH", "")}
                        res = abbyy_license.run_custom_activation(cmd, env, timeout=120)
                    except Exception as exc:
                        st.error(f"❌ Unexpected error: {exc}")
                        st.stop()
                if res.reason in ("converting", "busy"):  # ABBYY was never contacted
                    _note.empty()
                    _cmd_box.empty()
                _finish_custom_activation(
                    res, "Project ID or License Password incorrect", staged_token=target)

st.markdown("---")

# ── 3. Option B — Online activation ──────────────────────────────────────────
st.subheader("Option B — Online Activation via Serial Number")
st.markdown(
    "Use this if you have an ABBYY **Serial Number** and **Customer Project ID** "
    "but no `.ActivationToken` file yet.  "
    "The activation tool contacts ABBYY's activation server (`registration2.abbyy.com`, port 443) to activate "
    "the serial number.  **Internet access is required.**"
)

if not ACTIVATE_SH.exists():
    st.error(f"❌ Activation script not found at `{ACTIVATE_SH}`.  "
             "Ensure the ABBYY FREngine12 bundle is installed correctly.")
else:
    with st.expander("🔑 Enter Activation Credentials", expanded=_expand_forms):
        with st.form("online_activation_form", clear_on_submit=False):
            st.markdown("**Required credentials** (from your ABBYY purchase order / support email):")

            col_s, col_p = st.columns(2)
            with col_s:
                serial = st.text_input(
                    "Serial Number  *",
                    placeholder="your ABBYY Serial Number",
                    help="Uppercase letters and digits — provided by ABBYY.",
                )
            with col_p:
                project_id = st.text_input(
                    "Customer Project ID  *",
                    placeholder="your Customer Project ID",
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
            submitted = st.form_submit_button("🚀 Activate Now", type="primary", disabled=not _override_ok)

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

                _note, _cmd_box = st.empty(), st.empty()
                _note.info("Connecting to ABBYY's activation server…")
                _cmd_box.code(" ".join(
                    f"--license-password ****" if p == lic_password.strip() and lic_password.strip()
                    else p
                    for p in cmd
                ), language="bash")

                with st.spinner("Activating license (may take 30–60 s)…"):
                    try:
                        env = {**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8", "LD_LIBRARY_PATH": "/usr/local/lib/ABBYY/SDK/12/Licensing:" + os.environ.get("LD_LIBRARY_PATH", "")}
                        res = abbyy_license.run_custom_activation(cmd, env, timeout=120)
                    except Exception as exc:
                        st.error(f"❌ Unexpected error: {exc}")
                        st.stop()
                if res.reason in ("converting", "busy"):  # ABBYY was never contacted
                    _note.empty()
                    _cmd_box.empty()
                _finish_custom_activation(res, "Serial Number or Project ID incorrect")

st.markdown("---")

# ── 4. Service Control ────────────────────────────────────────────────────────
st.subheader("⚙️ Service Control")
st.markdown("Restart the ABBYY LicensingService daemon after any license change.")

col_btn, col_check, col_note = st.columns([1, 1, 3])
with col_btn:
    if st.button("🔄 Restart LicensingService", use_container_width=True,
                 key="btn_restart_ls"):
        with abbyy_license.conversions_idle() as _idle:
            if _idle:
                with st.spinner("Restarting LicensingService…"):
                    ok3, msg3 = _restart_ls()
                if ok3:
                    st.success(msg3)
                else:
                    st.error(f"❌ {msg3}")
                st.rerun()
            else:
                st.warning(f"⏳ {abbyy_license.reason_text('converting')} The service was not restarted.")

with col_check:
    if st.button("🩺 Check license", use_container_width=True, key="btn_check_license",
                 help="Asks ABBYY whether the license currently in use works. No pages are converted."):
        with st.spinner("Asking ABBYY…"):
            verdict = abbyy_license.recheck()
        if verdict is True:
            st.session_state["_license_flash"] = ("success", "✅ ABBYY confirms that the license works. "
                                                  "(Checked without converting any pages.)")
        elif verdict is False:
            st.session_state["_license_flash"] = ("error", "❌ ABBYY reports that there is no usable license "
                                                  "for the current settings.")
        else:
            st.session_state["_license_flash"] = ("info", "The check could not be completed right now "
                                                  "(another license change or a PDF conversion may be running). "
                                                  "Please try again.")
        st.rerun()

with col_note:
    st.caption(
        "The LicensingService daemon must be running for ABBYY to accept conversion jobs. "
        "If PDF → DOCX conversions are failing, try restarting the service here first. "
        "**Check license** asks ABBYY itself whether the current license works, without converting a page."
    )

st.markdown("---")

# ── 5. Connection check ───────────────────────────────────────────────────────
st.subheader("🌐 Connection check")
st.markdown(
    "Tests from **this server** whether ABBYY's activation and license servers can be reached. "
    "Use it when an activation fails with *Could not activate license via internet*: ABBYY's tool shows that same "
    "line when it cannot connect and when ABBYY's server answers \"no\". "
    "It asks each server for its home page and sends one empty test message to ABBYY's activation service; "
    "no license data is sent."
)


def _render_connection_check(checks) -> None:
    def cell(value) -> str:
        return str(value).replace("|", "/").replace("\n", " ")

    level, sentence = abbyy_license.summarize_connection(checks)
    getattr(st, level)(sentence)
    rows = ["| Server | What it is for | Name lookup | Connect | Secure connection | Web reply | Result |",
            "|---|---|---|---|---|---|---|"]
    verdicts = {"ok": "✅ ok", "expected": "ℹ️ expected", "problem": "❌ problem"}
    for c in checks:
        rows.append(f"| `{cell(c.host)}` | {cell(c.purpose)} | {cell(c.dns)} | {cell(c.tcp)} | {cell(c.tls)} | "
                    f"{cell(c.http)} | {verdicts[c.verdict]} ({c.seconds:.1f} s) |")
    st.markdown("\n".join(rows))
    proxy = abbyy_license.proxy_summary()
    st.caption(
        "Proxy settings on this server: "
        + (", ".join(f"`{n}`" for n in proxy["env"]) if proxy["env"] else "none in the environment")
        + (" · proxy entry in ABBYY's LicensingSettings.xml: yes" if proxy["licensing_settings"]
           else " · no proxy entry in ABBYY's LicensingSettings.xml")
        + " (values are never shown)."
    )
    logs = abbyy_license.abbyy_log_tail()
    if any(logs.values()):
        with st.expander("ABBYY's own license log (technical, identifiers hidden)"):
            for kind, lines in logs.items():
                if lines:
                    st.markdown(f"**{kind}**")
                    st.code("\n".join(lines), language="text")


if st.button("🌐 Test connection to ABBYY", key="btn_conn_check",
             help="Checks name lookup, connection, certificate and web reply for each ABBYY server."):
    with st.spinner("Testing the connection (up to 25 seconds)…"):
        st.session_state["_conn_checks"] = abbyy_license.diagnose_connection()
if st.session_state.get("_conn_checks"):
    _render_connection_check(st.session_state["_conn_checks"])
    st.caption(
        "Which step fails tells what to fix: **Name lookup** failed = DNS is blocked. **Connect** failed = a firewall "
        "blocks the connection. **Secure connection** failed = the certificate is not trusted here (corporate TLS "
        "inspection or a missing certificate). **Web reply** failed = a proxy or firewall answers instead of ABBYY. "
        "The old `.ru` address is expected to fail."
    )

st.markdown("---")

# ── 6. Help / FAQ ─────────────────────────────────────────────────────────────
with st.expander("❓ Help & FAQ", expanded=False):
    st.markdown("""
**Why does this page say my license is already set up?**
Your administrator configured the organization's production ABBYY license on the server. It is
activated automatically every time the server starts, so you never need to enter anything.

**When would I use Option A or Option B?**
Only if you have a different ABBYY license of your own and want this server to use it instead.
The change applies to **everyone using this server**, and the default license comes back by itself
when the server restarts. You can also bring it back immediately with **Switch back to the default license**.

**I set up my own license. How long does it last?**
Until the server restarts (this happens after updates or when the platform moves the app).
After that the default license is used again.

**What does "Check license" do?**
It asks ABBYY whether the license currently in use works. It takes a few seconds and does not
convert or count any pages.

**What if someone is converting a document while I change the license?**
A license change is refused while a PDF conversion is running on this server, and nothing is changed.
Try again in a minute. Conversions that start while a change is in progress simply wait for it (usually 10 to 15 seconds).

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
2. Check internet access with **Test connection to ABBYY** (above): every ABBYY server in the table must be
   reachable over HTTPS (port 443). Serial activation uses `registration2.abbyy.com`.
3. If your network uses a corporate proxy / Zscaler, TR CA certificates must be installed
   (Dockerfile already bundles them — they are in `abbyy_bundle/tr_certs/`).
4. Contact ABBYY support at `support@abbyy.com` with your Serial Number.

**The page says ABBYY "did not grant the activation" (code `refused`). What does that mean?**
The connection works: ABBYY's license server was reached and answered, but said no. Most often the number of
activations allowed for the serial number is used up. ABBYY counts every new container or server as a new computer,
so each restart of a server that has to activate again uses one up. Ask ABBYY to check the activations of the serial
number, or to provide an Online (token) license, which is made for containers.

**Where is the license stored on disk?**
All token files live in `/var/lib/ABBYY/SDK/12/Licenses/`.
They are named `<something>.ActivationToken`.

**Can I back up my license?**
Yes — copy the `.ActivationToken` file to a safe location.
If you reinstall, use **Option A** to restore it.
""")
