#!/usr/bin/env python3
"""ABBYY FREngine 12 license lifecycle for the container.

One engine shared by docker_entrypoint.sh (background activation of the default
production license configured in Plexus) and the Streamlit pages (banner,
override, switch back). License identifiers never leave this module: nothing it
logs, stores or returns contains the serial number, Customer Project ID or any
raw activation output.

Plexus configuration (Deployment Job wizard: Environment / Secrets steps; read as environment variables):
    ABBYY_SERIAL_NUMBER         the production serial number
    ABBYY_CUSTOMER_PROJECT_ID   the production Customer Project ID

CLI (used by docker_entrypoint.sh):
    python3 abbyy_license.py autoactivate    # background job at container start
    python3 abbyy_license.py status          # JSON state without identifiers
"""
from __future__ import annotations

import fcntl
import json
import os
import re
import shutil
import signal
import socket
import ssl
import subprocess
import sys
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional

_SAFE = re.compile(r"^[A-Za-z0-9_.\-]+$")
_DEFAULT_RETRY_DELAYS = (10, 30, 60, 120, 240, 480)
_STALE_AFTER = 45 * 60  # seconds; longer than the whole retry schedule
_NOT_STARTED_AFTER = 5 * 60

# Friendly, non-technical explanations for the reason codes stored in the state file.
REASON_TEXT = {
    "invalid_serial": "ABBYY did not accept the license details configured for this server.",
    "network": "This server could not reach ABBYY's license server (this is usually temporary).",
    "timeout": "ABBYY's license server took too long to answer (this is usually temporary).",
    "service": "The ABBYY license service on this server is not ready.",
    "limit": "ABBYY reports that this license cannot be activated on another server right now.",
    "refused": "ABBYY's license server was reached, but it did not grant the activation.",
    "config": "The ABBYY settings on this server could not be updated.",
    "busy": "Another license change is in progress. Please try again in a moment.",
    "converting": "A PDF conversion is running on this server right now. Please try again in a minute, when it has finished.",
    "not_started": "The automatic license setup did not start on this server.",
    "not_licensed": "ABBYY accepted the request, but the license cannot be used with the Customer Project ID provided. "
                    "The serial number and Customer Project ID probably do not belong together.",
    "unknown": "An unexpected problem occurred while activating the license.",
}
# Once the automatic retries are used up, "usually temporary" would be misleading.
FINAL_REASON_TEXT = {
    "network": "This server could not reach ABBYY's license server, even after several tries over about 15 minutes. "
               "Something between this server and ABBYY, for example a firewall, is probably blocking the connection.",
    "timeout": "ABBYY's license server did not answer in time, even after several tries. "
               "Something between this server and ABBYY, for example a firewall, is probably blocking the connection.",
    "refused": "ABBYY's license server was reached from this server, but it refused to activate the serial number. "
               "The connection itself works, so this is not a firewall problem, and trying again only repeats the answer "
               "until something changes at ABBYY. Usually the number of activations the serial number allows is used up "
               "(ABBYY counts every new container as a new computer), or the serial number cannot be activated this way.",
}


# ── Settings (read at call time so tests and the container can override) ──────
@dataclass(frozen=True)
class Settings:
    fre_dir: Path
    activate_sh: Path
    license_manager: Path
    ls_bin: Path
    ls_lib: str
    ls_port: int
    ls_host: str
    state_dir: Path
    skip_ls_restart: bool
    retry_delays: tuple
    licenses_dir: Path
    staging_dir: Path


def settings() -> Settings:
    fre = Path(os.environ.get("ABBYY_FRE_DIR", "/opt/ABBYY/FREngine12"))
    delays = os.environ.get("ABBYY_AUTOACTIVATE_DELAYS", "")
    parsed = tuple(int(x) for x in delays.split(",") if x.strip().isdigit()) if delays else _DEFAULT_RETRY_DELAYS
    return Settings(
        fre_dir=fre,
        activate_sh=Path(os.environ.get("ABBYY_ACTIVATE_SH", str(fre / "activatefre.sh"))),
        license_manager=Path(os.environ.get("ABBYY_LICENSE_MANAGER", str(fre / "Bin" / "LicenseManager.Console"))),
        ls_bin=Path(os.environ.get("ABBYY_LS_BIN", "/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService")),
        ls_lib=os.environ.get("ABBYY_LS_LIB", "/usr/local/lib/ABBYY/SDK/12/Licensing"),
        ls_port=int(os.environ.get("ABBYY_LS_PORT", "3023")),
        ls_host=os.environ.get("ABBYY_LS_HOST", "").strip(),
        state_dir=Path(os.environ.get("ABBYY_STATE_DIR", "/tmp/abbyy_runtime")),
        skip_ls_restart=os.environ.get("ABBYY_SKIP_LS_RESTART", "") == "1",
        retry_delays=parsed,
        licenses_dir=Path(os.environ.get("ABBYY_REAL_LICENSES_DIR", "/var/lib/ABBYY/SDK/12/Licenses")),
        staging_dir=Path(os.environ.get("ABBYY_LICENSES_DIR", "/tmp/abbyy_licenses")),
    )


# ── Default license (from Plexus environment) ─────────────────────────────────
@dataclass(frozen=True)
class DefaultLicense:
    serial: str
    project_id: str


def default_license() -> Optional[DefaultLicense]:
    serial = os.environ.get("ABBYY_SERIAL_NUMBER", "").strip()
    project = os.environ.get("ABBYY_CUSTOMER_PROJECT_ID", "").strip()
    if serial and project and _SAFE.match(serial) and _SAFE.match(project):
        return DefaultLicense(serial, project)
    return None


def default_configured() -> bool:
    return default_license() is not None


def default_misconfigured() -> bool:
    """Some default-license variable is set, but not both or not valid."""
    serial = os.environ.get("ABBYY_SERIAL_NUMBER", "").strip()
    project = os.environ.get("ABBYY_CUSTOMER_PROJECT_ID", "").strip()
    return bool(serial or project) and default_license() is None


# ── Redaction and logging ─────────────────────────────────────────────────────
def _variants(secret: str) -> set:
    s = secret.strip()
    compact = re.sub(r"[-\s]", "", s)
    out = {s, compact, s.lower(), s.upper(), compact.lower(), compact.upper()}
    if len(compact) == 24:  # serials are written XXXX-XXXX-XXXX-XXXX-XXXX-XXXX or compact
        out.add("-".join(compact[i:i + 4] for i in range(0, 24, 4)))
    return {v for v in out if len(v) >= 4}


def redact(text: str, extra: tuple = ()) -> str:
    """Hide the default license identifiers (and any extra secrets) in text."""
    if not text:
        return text
    secrets = list(extra)
    d = default_license()
    if d:
        secrets += [d.serial, d.project_id]
    # Longest first so a serial is hidden before any shorter value that it contains.
    for v in sorted({x for s in secrets if s for x in _variants(s)}, key=len, reverse=True):
        text = text.replace(v, "[hidden]")
    return text


def log(message: str) -> None:
    print(f"[abbyy-license] {redact(message)}", flush=True)


def _mask_ids(text: str) -> str:
    """Hide the default identifiers and anything that looks like another identifier or token."""
    return re.sub(r"[A-Za-z0-9_\-]{16,}", "<id>", redact(text))


def tool_output_tail(output: str, max_lines: int = 8) -> str:
    """The last meaningful lines of an ABBYY tool's output, safe to log or show."""
    lines = [_mask_ids(raw.strip())[:160] for raw in (output or "").splitlines() if raw.strip()]
    lines = [ln for ln in lines if not ln.startswith("Copyright") and "activation script" not in ln]
    return "\n".join(lines[-max_lines:])


# ── Connection check: which step fails when the activation cannot reach ABBYY? ─
# ABBYY's tool prints one and the same line, "Could not activate license via internet", whether its web request
# failed or ABBYY's server answered "no" (measured 2026-11-16 with a local forwarder: ABBYY answered "Failed, limit
# exceed" within 2 s and the tool printed that line). This tells, from THIS server, whether DNS, the connection,
# the TLS certificate or the web reply fails, and whether ABBYY's activation service answers at all.
# Host names: found inside ABBYY's licensing library and named in ABBYY's firewall guidance for activation.
# Only an ordinary GET of each home page and one EMPTY test message to the activation address are sent; no license data.
CHECK_HOSTS = (
    ("registration2.abbyy.com", "serial activation (address found in ABBYY's tool)"),
    ("activation.abbyy.com", "activation (named in ABBYY's firewall guidance)"),
    ("registration.abbyy.com", "registration (named in ABBYY's firewall guidance)"),
    ("license-cloud.abbyy.com", "online licenses (token files)"),
    ("account.abbyy.com", "ABBYY account"),
    ("activation1.abbyy.ru", "old address inside ABBYY's tool (no longer exists, failing is normal)"),
    ("example.com", "control: any public web site"),
)
_PROXY_VARS = ("HTTPS_PROXY", "HTTP_PROXY", "ALL_PROXY", "NO_PROXY")
_CONNECTION_REASONS = ("network", "timeout", "unknown")
EXPECTED_TO_FAIL = frozenset({"activation1.abbyy.ru"})  # the name no longer exists in DNS
CONTROL_HOST = "example.com"
ACTIVATION_HOST = "registration2.abbyy.com"
ACTIVATION_PATH = "/ProductWebCall/BusinessVega.asmx"  # where ABBYY's tool posts its activation request (measured)
ACTIVATION_PROBE = "activation service (an empty test message, no license data)"
_SOAP_PING = (b'<?xml version="1.0" encoding="utf-8"?><soap12:Envelope '
              b'xmlns:soap12="http://www.w3.org/2003/05/soap-envelope"><soap12:Body/></soap12:Envelope>')
_ROUTINE_LOG_LINE = re.compile(r"Application context|Initializing|Deinitializing|Server address|"
                               r"Connection with .* lost|Starting |initialized|Protocol type", re.I)


@dataclass
class HostCheck:
    host: str
    purpose: str
    dns: str = "-"
    tcp: str = "-"
    tls: str = "-"
    http: str = "-"
    seconds: float = 0.0
    kind: str = "page"  # page (a server's home page) | activation (empty test message to ABBYY's activation service)

    @property
    def ok(self) -> bool:
        return all(step.startswith("ok") for step in (self.dns, self.tcp, self.tls, self.http))

    @property
    def verdict(self) -> str:
        """ok | expected (a failure that is normal for this host) | problem"""
        if self.ok:
            return "ok"
        return "expected" if self.host in EXPECTED_TO_FAIL else "problem"

    def failing_step(self) -> str:
        for attr, label in (("dns", "name lookup"), ("tcp", "connection"), ("tls", "secure connection (certificate)"),
                            ("http", "web reply")):
            if not getattr(self, attr).startswith("ok"):
                return label
        return ""


def _short(exc: BaseException) -> str:
    return (str(exc).strip().replace("\n", " ") or type(exc).__name__)[:140]


def _issuer_name(cert: Optional[dict]) -> str:
    try:
        parts = {k: v for rdn in (cert or {}).get("issuer", ()) for k, v in rdn}
    except (TypeError, ValueError):
        return ""
    return parts.get("commonName") or parts.get("organizationName") or ""


def _describe_reply(data: bytes, soap_probe: bool = False) -> str:
    head, _, body = data.partition(b"\r\n\r\n")
    first = head.split(b"\r\n", 1)[0].decode("latin-1", "replace")
    m = re.match(r"HTTP/\d(?:\.\d)?\s+(\d{3})", first)
    if not m:
        return "FAILED: no web reply (" + (first[:60] or "empty answer") + ")"
    code = int(m.group(1))
    # 403/407/451 and 5xx are typical of a firewall or proxy block page. A SOAP service answers an empty test message
    # with 500 (a fault): for the probe that proves the service itself answered.
    if code in (403, 407, 451) or (code >= 500 and not (soap_probe and code == 500)):
        snippet = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body.decode("latin-1", "replace"))).strip()[:100]
        return f"FAILED: HTTP {code}" + (f" ({snippet})" if snippet else "")
    return f"ok HTTP {code}"


def _check_host(host: str, purpose: str, timeout: float = 6.0, port: int = 443, probe: Optional[str] = None) -> HostCheck:
    """probe: path of ABBYY's activation service; an empty test message is POSTed there instead of GET /."""
    res = HostCheck(host, purpose, kind="activation" if probe else "page")
    started = time.monotonic()
    sock = tls = None
    try:
        try:
            infos = socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
        except OSError as exc:
            res.dns = f"FAILED: {_short(exc)}"
            return res
        addrs = list(dict.fromkeys(info[4][0] for info in infos))
        res.dns = "ok " + ", ".join(addrs[:2])
        last: Optional[BaseException] = None
        for addr in addrs:
            try:
                sock = socket.create_connection((addr, port), timeout)
                break
            except OSError as exc:
                last = exc
        if sock is None:
            res.tcp = f"FAILED: {_short(last) if last else 'no address'}"
            return res
        res.tcp = "ok"
        try:
            tls = ssl.create_default_context().wrap_socket(sock, server_hostname=host)
        except ssl.SSLCertVerificationError as exc:
            res.tls = f"FAILED: certificate not trusted ({exc.verify_message})"
            return res
        except (ssl.SSLError, OSError) as exc:
            res.tls = f"FAILED: {_short(exc)}"
            return res
        issuer = _issuer_name(tls.getpeercert())
        res.tls = f"ok {tls.version()}" + (f", signed by {issuer}" if issuer else "")
        try:
            tls.settimeout(timeout)
            if probe:
                tls.sendall((f"POST {probe} HTTP/1.1\r\nHost: {host}\r\nUser-Agent: abbyy-connection-check\r\n"
                             "Content-Type: application/soap+xml; charset=utf-8\r\n"
                             f"Content-Length: {len(_SOAP_PING)}\r\nConnection: close\r\n\r\n").encode() + _SOAP_PING)
            else:
                tls.sendall((f"GET / HTTP/1.1\r\nHost: {host}\r\nUser-Agent: abbyy-connection-check\r\n"
                             "Accept: */*\r\nConnection: close\r\n\r\n").encode())
            data = b""
            while len(data) < 4096:
                chunk = tls.recv(2048)
                if not chunk:
                    break
                data += chunk
            res.http = _describe_reply(data, soap_probe=bool(probe))
        except (ssl.SSLError, OSError) as exc:
            res.http = f"FAILED: {_short(exc)}"
        return res
    except Exception as exc:  # a diagnostic must never raise
        res.http = f"FAILED: {type(exc).__name__}"
        return res
    finally:
        for s in (tls, sock):
            try:
                if s is not None:
                    s.close()
            except OSError:
                pass
        res.seconds = time.monotonic() - started


def diagnose_connection(hosts=None, timeout: float = 6.0, port: int = 443) -> list:
    """Check every host in parallel; never raises; the answers keep the order of `hosts`.

    The default list gets one more check at the end: an empty test message to ABBYY's activation address, to see that
    the activation service itself answers (what the activation tool needs), not only the web server in front of it."""
    jobs = [(h, p, None) for h, p in (hosts or CHECK_HOSTS)]
    if hosts is None:
        jobs.append((ACTIVATION_HOST, ACTIVATION_PROBE, ACTIVATION_PATH))
    results: list = [None] * len(jobs)

    def work(i: int, host: str, purpose: str, probe: Optional[str]) -> None:
        try:
            results[i] = (_check_host(host, purpose, timeout, port) if probe is None
                          else _check_host(host, purpose, timeout, port, probe))
        except Exception as exc:
            results[i] = HostCheck(host, purpose, dns=f"FAILED: {type(exc).__name__}",
                                   kind="activation" if probe else "page")

    threads = [threading.Thread(target=work, args=(i, *job), daemon=True) for i, job in enumerate(jobs)]
    for t in threads:
        t.start()
    deadline = time.monotonic() + timeout * 4  # a resolver that never answers must not hang the page
    for t in threads:
        t.join(max(0.0, deadline - time.monotonic()))
    return [r or HostCheck(h, p, dns="FAILED: no answer in time", kind="activation" if probe else "page")
            for r, (h, p, probe) in zip(results, jobs)]


def connection_log_lines(checks) -> list:
    return [f"{c.host}{' [activation request]' if c.kind == 'activation' else ''}: dns={c.dns}; tcp={c.tcp}; "
            f"tls={c.tls}; web={c.http} ({c.seconds:.1f} s)" for c in checks]


def summarize_connection(checks) -> tuple:
    """(level, plain sentence) for the result of diagnose_connection()."""
    if not checks:
        return ("info", "The connection check returned no results.")
    problems = [c for c in checks if c.verdict == "problem"]
    if not problems:
        return ("success", "This server can reach every ABBYY server that is needed. If an activation still fails, "
                "ABBYY's server was reached and refused it. Most often the number of activations allowed for the serial "
                "number is used up: ask ABBYY to check the activations of the serial number.")
    def label(c) -> str:
        return f"`{c.host}`" + (" (activation request)" if c.kind == "activation" else "")

    where = "; ".join(f"{label(c)} fails at the {c.failing_step()}" for c in problems)
    control = next((c for c in checks if c.host == CONTROL_HOST and c.verdict == "problem"), None)
    if control is None:
        return ("error", f"Not every ABBYY server can be reached from this server: {where}. "
                "Ask the network team to allow outbound HTTPS (port 443) to the servers that fail.")
    step = control.failing_step()
    if step in ("name lookup", "connection"):
        lead = f"This server seems to have no general internet access: even an ordinary web site fails at the {step}."
    elif step.startswith("secure connection"):
        lead = ("This server does not trust the certificates it is shown, even for an ordinary web site. Most likely the "
                "certificate of a corporate TLS-inspection proxy is missing here.")
    else:
        lead = "Even an ordinary web site gets no proper web reply, so a proxy or firewall is answering instead of the sites."
    return ("error", f"{lead} Details: {where}.")


def proxy_summary() -> dict:
    """Which proxy settings exist (never their values, they can hold credentials)."""
    present = [n for n in _PROXY_VARS if os.environ.get(n) or os.environ.get(n.lower())]
    configured = False
    for name in ("Bin_LicensingSettings.xml", "CommonBin_LicensingSettings.xml"):
        try:
            text = _config_files()[name].read_text(errors="ignore")
        except OSError:
            continue
        configured = configured or bool(re.search(r"<ProxyServer\b|ProxyServer\s*=", text))
    return {"env": present, "licensing_settings": configured}


def abbyy_log_tail(max_lines: int = 25, max_bytes: int = 65536) -> dict:
    """Last non-routine lines of ABBYY's own licensing logs, identifiers masked."""
    out = {}
    for kind in ("ProductProtection", "NetworkServer"):
        lines: list = []
        try:
            files = sorted(settings().licenses_dir.glob(f"{kind}_*.log"), key=lambda p: p.stat().st_mtime)
            if files:
                with open(files[-1], "rb") as fh:
                    fh.seek(0, os.SEEK_END)
                    fh.seek(max(0, fh.tell() - max_bytes))
                    raw = fh.read()
                text = (raw.decode("utf-16", "replace") if raw.count(b"\x00") > len(raw) // 4
                        else raw.decode("utf-8", "replace"))
                text = text.replace("\ufeff", "").replace("\ufffd", "")
                lines = [_mask_ids(ln)[:200] for ln in text.splitlines()
                         if ln.strip() and not _ROUTINE_LOG_LINE.search(ln)]
        except OSError:
            lines = []
        out[kind] = lines[-max_lines:]
    return out


def activation_service_reached(checks) -> bool:
    """True when ABBYY's activation service answered the empty test message (the service the serial activation uses).

    Deliberately not "every server in the table is fine": one failed connection to some other address, or a hiccup on
    the home page of the same host, must not hide that the activation service itself answered."""
    return any(c.kind == "activation" and c.ok for c in checks)


def _checks_with_second_look() -> tuple:
    """(checks, looked_twice). One failed connection can be a hiccup (seen: "Network is unreachable" for one of eight
    parallel checks while the others to the same address worked), so if ABBYY's activation service did not answer,
    look once more before calling it unreachable."""
    checks = diagnose_connection()
    if activation_service_reached(checks):
        return checks, False
    time.sleep(2)
    return diagnose_connection(), True


def _log_connection_check() -> bool:
    """Run the connection check and log it. True when ABBYY's activation service answered (False: it did not, or unknown)."""
    try:
        checks, again = _checks_with_second_look()
        if again:
            log("connection check: ABBYY's activation service did not answer the first time, so it was checked again.")
        for line in connection_log_lines(checks):
            log(f"connection check: {line}")
        if checks:
            log(f"connection check result: {summarize_connection(checks)[1]}")
        return activation_service_reached(checks)
    except Exception as exc:  # diagnostics must never stop the activation job
        log(f"connection check could not run: {type(exc).__name__}")
        return False


def refine_reason(reason: str) -> str:
    """Tell "cannot connect" from "ABBYY said no". The ABBYY tool prints the same line for both (for example when the
    activations of the serial number are used up, ABBYY answers "Failed, limit exceed"). Only `network` is refined:
    if ABBYY's activation service answers from here, the reason is `refused`, not a connection problem."""
    if reason == "network" and _log_connection_check():
        return "refused"
    return reason


# ── State file (never holds identifiers) ──────────────────────────────────────
_STATE_KEYS = ("mode", "status", "reason", "attempt", "next_retry_at", "verified", "checked_at", "ts")


def _state_path() -> Path:
    return settings().state_dir / "license_state.json"


def read_state() -> dict:
    try:
        raw = json.loads(_state_path().read_text())
        return {k: raw[k] for k in _STATE_KEYS if k in raw}
    except (OSError, ValueError):
        return {}


def write_state(**fields) -> dict:
    state = read_state()
    state.update({k: v for k, v in fields.items() if k in _STATE_KEYS})
    state["ts"] = time.time()
    path = _state_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + f".{os.getpid()}.{threading.get_ident()}.tmp")
    tmp.write_text(json.dumps(state))
    os.replace(tmp, path)
    return state


@contextmanager
def _lock(blocking: bool = False) -> Iterator[bool]:
    """Exclusive lock shared by the entrypoint job and every Streamlit session."""
    d = settings().state_dir
    d.mkdir(parents=True, exist_ok=True)
    fd = os.open(d / "license.lock", os.O_CREAT | os.O_RDWR, 0o666)
    got = False
    try:
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | (0 if blocking else fcntl.LOCK_NB))
            got = True
        except BlockingIOError:
            got = False
        yield got
    finally:
        if got:
            fcntl.flock(fd, fcntl.LOCK_UN)
        os.close(fd)


# ── PDF conversions versus license changes ────────────────────────────────────
# Replacing the working license restarts the licensing service, so it must never overlap a PDF
# conversion. A conversion holds a shared flock while the ABBYY tool runs; a change needs the
# exclusive flock and is refused when it cannot get it; conversions that start meanwhile wait.
# The kernel drops a flock when its process dies, so the lock can never get stuck.
def _conversion_lock_fd() -> Optional[int]:
    try:
        d = settings().state_dir
        d.mkdir(parents=True, exist_ok=True)
        return os.open(d / "conversion.lock", os.O_CREAT | os.O_RDWR, 0o666)
    except OSError:
        return None  # a lock problem must never stop a conversion


@contextmanager
def conversion_slot(wait: float = 90.0, poll: float = 0.5, sleep=None) -> Iterator[bool]:
    """Hold this while the ABBYY tool converts a PDF; waits while a license change is running.

    Yields False when the change did not finish within `wait` seconds (do not convert then).
    """
    sleep = sleep or time.sleep
    fd = _conversion_lock_fd()
    if fd is None:
        yield True
        return
    got = False
    try:
        end = time.time() + wait
        while True:
            try:
                fcntl.flock(fd, fcntl.LOCK_SH | fcntl.LOCK_NB)
                got = True
                break
            except BlockingIOError:
                if time.time() >= end:
                    break
                sleep(poll)
            except OSError:  # locking is not supported here
                yield True
                return
        yield got
    finally:
        if got:
            fcntl.flock(fd, fcntl.LOCK_UN)
        os.close(fd)


@contextmanager
def conversions_idle() -> Iterator[bool]:
    """True only while no conversion is running; new conversions wait until this is released."""
    fd = _conversion_lock_fd()
    if fd is None:
        yield True
        return
    got = False
    try:
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
            got = True
        except BlockingIOError:
            pass
        except OSError:  # locking is not supported here
            yield True
            return
        yield got
    finally:
        if got:
            fcntl.flock(fd, fcntl.LOCK_UN)
        os.close(fd)


@contextmanager
def _license_change() -> Iterator[str]:
    """Both locks for a change that replaces the working license.

    Yields "" when the change may go ahead, otherwise why not: "busy" (another license change)
    or "converting" (a PDF conversion is running).
    """
    with _lock() as got:
        if not got:
            yield "busy"
            return
        with conversions_idle() as idle:
            yield "" if idle else "converting"


# ── Licensing service control ─────────────────────────────────────────────────
def _ls_pids() -> list:
    """PIDs whose executable path is exactly the LicensingService binary.

    Matching argv[0] (not `pkill -f`) means a wrapper shell or this process can
    never match itself, and zombies (empty cmdline) are ignored.
    """
    target = str(settings().ls_bin)
    pids = []
    for entry in os.listdir("/proc"):
        if entry.isdigit():
            try:
                argv = Path(f"/proc/{entry}/cmdline").read_bytes().split(b"\0")
            except OSError:
                continue
            if argv and argv[0].decode(errors="ignore") == target:
                pids.append(int(entry))
    return pids


def ls_running() -> bool:
    s = settings()
    if s.ls_host:
        return _port_open(s.ls_host, s.ls_port)
    return bool(_ls_pids())


def _port_open(host: str, port: int, timeout: float = 1.0) -> bool:
    try:
        socket.create_connection((host, port), timeout).close()
        return True
    except OSError:
        return False


def wait_for_ls(timeout: float = 60.0) -> bool:
    s = settings()
    if s.skip_ls_restart:
        return True
    end = time.time() + timeout
    while time.time() < end:
        if _port_open("127.0.0.1", s.ls_port):
            return True
        time.sleep(1)
    return False


def restart_ls(timeout: float = 25.0) -> bool:
    """Stop and start the local LicensingService; True once it listens again."""
    s = settings()
    if s.skip_ls_restart:
        return True
    if not s.ls_bin.exists():
        return False
    for pid in _ls_pids():
        try:
            os.kill(pid, signal.SIGTERM)
        except OSError:
            pass
    end = time.time() + 10
    while _ls_pids() and time.time() < end:
        time.sleep(0.2)
    for pid in _ls_pids():
        try:
            os.kill(pid, signal.SIGKILL)
        except OSError:
            pass
    env = {**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8",
           "LD_LIBRARY_PATH": s.ls_lib + ":" + os.environ.get("LD_LIBRARY_PATH", "")}
    subprocess.Popen([str(s.ls_bin), "/standalone"], stdout=subprocess.DEVNULL,
                     stderr=subprocess.DEVNULL, start_new_session=True, env=env)
    end = time.time() + timeout
    while time.time() < end:
        if _port_open("127.0.0.1", s.ls_port):
            return True
        time.sleep(0.3)
    return False


# ── Config snapshot / restore ─────────────────────────────────────────────────
# activatefre.sh rewrites SamplesConfig.h BEFORE it tries to activate, so a failed
# attempt would otherwise leave the working license pointing at wrong values.
def _config_files() -> dict:
    fre = settings().fre_dir
    return {
        "SamplesConfig.h": fre / "Samples" / "SamplesConfig.h",
        "SamplesConfig.java": fre / "Samples" / "Java" / "SamplesConfig.java",
        "Bin_LicensingSettings.xml": fre / "Bin" / "LicensingSettings.xml",
        "CommonBin_LicensingSettings.xml": fre / "CommonBin" / "Licensing" / "LicensingSettings.xml",
    }


def snapshot() -> dict:
    snap = {}
    for name, path in _config_files().items():
        try:
            snap[name] = path.read_bytes()
        except OSError:
            snap[name] = None
    return snap


def restore(snap: dict) -> None:
    for name, path in _config_files().items():
        data = snap.get(name)
        if data is None:
            continue
        try:
            path.write_bytes(data)  # writes through the /tmp symlink, keeps its mode
        except OSError as exc:
            log(f"could not restore {name}: {type(exc).__name__}")


# ── Token files (Option A licenses) ───────────────────────────────────────────
# activatefre.sh copies an uploaded token into the real Licenses folder before anything is
# verified, so a rejected token would stay there, and a custom token would outlive a switch back.
def _token_names(directory: Path) -> set:
    try:
        return {p.name for p in directory.glob("*.ActivationToken")}
    except OSError:
        return set()


def _remove_tokens(names: set) -> int:
    removed = 0
    for directory in (settings().licenses_dir, settings().staging_dir):
        for name in names:
            try:
                (directory / name).unlink()
                removed += 1
            except OSError:
                pass
    return removed


def _release_online_licenses(timeout: int = 60) -> None:
    """Best effort: tell ABBYY this server no longer uses an online license."""
    lm = settings().license_manager
    if not lm.exists():
        return
    try:
        subprocess.run([str(lm), "--releaseOnlineLicenses"], capture_output=True, text=True,
                       timeout=timeout, env=_script_env(), cwd=str(lm.parent))
    except (OSError, subprocess.TimeoutExpired):
        log("could not release the online license (ignored).")


def _default_snapshot_dir() -> Path:
    return settings().state_dir / "default_config"


def save_default_snapshot() -> None:
    d = _default_snapshot_dir()
    d.mkdir(parents=True, exist_ok=True)
    for name, data in snapshot().items():
        if data is not None:
            (d / name).write_bytes(data)


def restore_default_snapshot() -> bool:
    d = _default_snapshot_dir()
    snap = {name: (d / name).read_bytes() if (d / name).exists() else None for name in _config_files()}
    if snap.get("SamplesConfig.h") is None:
        return False
    restore(snap)
    return True


# ── Activation ────────────────────────────────────────────────────────────────
@dataclass
class Outcome:
    ok: bool
    reason: str = ""
    transient: bool = False
    detail: str = ""  # redacted last lines of the ABBYY tool's own output (failures only)
    seconds: float = 0.0  # how long the ABBYY tool ran


_INVALID = re.compile(r"incorrect serial|invalid serial|wrong serial", re.I)
_LIMIT = re.compile(r"activation limit|limit (has been )?(reached|exceeded)|maximum (number of )?activations|"
                    r"too many activations|no (more )?activations|already (been )?activated on", re.I)
_NETWORK = re.compile(r"could not connect|cannot connect|connection (refused|reset|failed|error)|network|"
                      r"unreachable|resolve|dns|ssl|certificate|proxy|internet|no route", re.I)
_CONFIG = re.compile(r"configuring samples - failed|read-only file system|permission denied|cannot (create|register)", re.I)


def classify(returncode: int, output: str, timed_out: bool = False) -> Outcome:
    if timed_out:
        return Outcome(False, "timeout", True)
    if returncode == 0:
        return Outcome(True, "ok")
    if _INVALID.search(output):
        return Outcome(False, "invalid_serial", False)
    if _LIMIT.search(output):
        return Outcome(False, "limit", False)
    if _NETWORK.search(output):
        return Outcome(False, "network", True)
    if _CONFIG.search(output):
        return Outcome(False, "config", False)
    return Outcome(False, "unknown", True)


def _script_env() -> dict:
    s = settings()
    return {**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8",
            "LD_LIBRARY_PATH": s.ls_lib + ":" + os.environ.get("LD_LIBRARY_PATH", "")}


# ── Asking ABBYY itself whether the license works (costs no pages) ─────────────────────────────────
CONFIG_FIELDS = (("ABBYY_PROJECT_ID", "GetFreProjectId"), ("ABBYY_LICENSE_PATH", "GetLicensePath"),
                 ("ABBYY_LICENSE_PASSWORD", "GetLicensePassword"))
_NOT_LICENSED = re.compile(r"not licensed|no licenses? (on|found|available)|license.{0,40}(invalid|expired|not valid)", re.I)
_ERROR_LINE = re.compile(r"^\s*Error:", re.M)


def config_env() -> dict:
    """The ABBYY_* variables a conversion passes to the CLI (same source as streamlit_app)."""
    try:
        text = _config_files()["SamplesConfig.h"].read_text()
    except OSError:
        return {}
    out = {}
    for var, fn in CONFIG_FIELDS:
        m = re.search(fn + r'\(\)\s*\{\s*static wchar_t \w+\[\]\s*=\s*L"([^"]*)"', text)
        if m:
            out[var] = m.group(1)
    return out


def probe_license(timeout: int = 60) -> Optional[bool]:
    """True: ABBYY accepts the license now in the config. False: definitely not licensed.
    None: cannot tell (tool missing, timeout, some other error), never blamed on the license.

    Runs the conversion tool with no input: it checks the license first, so this
    takes about two seconds and consumes no pages.
    """
    fre = settings().fre_dir
    cli = Path(os.environ.get("ABBYY_CLI", str(fre / "Samples" / "CommandLineInterface" / "CommandLineInterface")))
    if not cli.exists():
        return None
    lib = os.environ.get("ABBYY_LIB", str(fre / "Bin"))
    env = {**_script_env(), "LD_LIBRARY_PATH": f"{lib}:{_script_env()['LD_LIBRARY_PATH']}", **config_env()}
    try:
        r = subprocess.run([str(cli)], capture_output=True, text=True, timeout=timeout, env=env)
    except (OSError, subprocess.TimeoutExpired):
        return None
    out = (r.stdout or "") + (r.stderr or "")
    if _NOT_LICENSED.search(out):
        return False
    if _ERROR_LINE.search(out) or r.returncode not in (0, 1):
        return None
    return True


def probe_with_retries(attempts: int = 3, pause: float = 3.0, sleep=None) -> Optional[bool]:
    """A freshly restarted licensing service may need a moment before it answers."""
    sleep = sleep or time.sleep
    verdict = None
    for i in range(attempts):
        verdict = probe_license()
        if verdict is not False:
            return verdict
        if i < attempts - 1:
            sleep(pause)
    return False


def _finalize(snap: dict) -> tuple:
    """After the activation script succeeded: restart the service and let ABBYY confirm.

    Returns (Outcome, verdict). On a failure the configuration is rolled back to `snap`.
    """
    if not restart_ls():
        restore(snap)
        return Outcome(False, "service", True), None
    verdict = probe_with_retries()
    if verdict is False:
        restore(snap)
        return Outcome(False, "not_licensed", False), False
    return Outcome(True, "ok"), verdict


def _activate_serial(serial: str, project_id: str, timeout: int = 120) -> Outcome:
    """Run activatefre.sh for a serial license; restores the config on failure."""
    s = settings()
    snap = snapshot()
    cmd = ["bash", str(s.activate_sh), "--serial-number", serial, "--project-id", project_id,
           "--skip-local-service-installation"]
    started = time.monotonic()
    detail = ""
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=_script_env())
        output = (r.stdout or "") + (r.stderr or "")
        outcome = classify(r.returncode, output)
        detail = tool_output_tail(output)
    except subprocess.TimeoutExpired:
        outcome = classify(1, "", timed_out=True)
    except OSError:
        outcome = Outcome(False, "config", False)
    outcome.seconds = time.monotonic() - started
    if not outcome.ok:
        outcome.detail = detail
        restore(snap)
    return outcome


def activate_default(respect_custom: bool = False, retrying: bool = False) -> Outcome:
    """Activate the default license from the environment (one attempt).

    respect_custom: the background job passes True so it never overwrites a
    license that someone set up from the page in the meantime.
    retrying: the background job passes True. It classifies a failure itself (with its own connection check) and
    writes the retry or final state, so here a failure is neither refined nor written down as final.
    """
    d = default_license()
    if d is None:
        return Outcome(False, "config", False)
    with _lock() as got:
        if not got:
            return Outcome(False, "busy", True)
        if respect_custom and read_state().get("mode") == "custom":
            return Outcome(True, "custom")
        snap = snapshot()
        write_state(mode="default", status="activating", reason="")
        outcome = _activate_serial(d.serial, d.project_id)
        if not retrying:
            outcome.reason = refine_reason(outcome.reason)
        verdict = None
        if outcome.ok:
            outcome, verdict = _finalize(snap)
        if outcome.ok:
            save_default_snapshot()
            write_state(mode="default", status="active", reason="ok", attempt=0, verified=verdict, checked_at=time.time())
        elif not retrying:
            write_state(mode="default", status="failed", reason=outcome.reason)
        return outcome


def _drop_custom_tokens(names: set) -> None:
    """After a switch back: the default is a serial license, so tokens belong to the custom one."""
    if names:
        _remove_tokens(names)
        _release_online_licenses()


def switch_back_to_default() -> Outcome:
    """Leave a custom license and use the default again (offline when possible)."""
    outcome = _switch_back_to_default()
    log("switched back to the default license." if outcome.ok else f"switching back failed ({outcome.reason}).")
    return outcome


def _switch_back_to_default() -> Outcome:
    d = default_license()
    if d is None:
        return Outcome(False, "config", False)
    with _license_change() as blocked:
        if blocked:
            return Outcome(False, blocked, True)
        previous = read_state()
        snap = snapshot()
        tokens_before = _token_names(settings().licenses_dir) | _token_names(settings().staging_dir)
        if restore_default_snapshot():
            outcome, verdict = _finalize(snap)
            if outcome.ok:
                write_state(mode="default", status="active", reason="ok", verified=verdict, checked_at=time.time())
                _drop_custom_tokens(tokens_before)
            return outcome  # on failure _finalize already restored the custom configuration
        write_state(mode="default", status="activating", reason="")
        outcome = _activate_serial(d.serial, d.project_id)
        outcome.reason = refine_reason(outcome.reason)
        verdict = None
        if outcome.ok:
            outcome, verdict = _finalize(snap)
        if outcome.ok:
            save_default_snapshot()
            write_state(mode="default", status="active", reason="ok", verified=verdict, checked_at=time.time())
            _drop_custom_tokens(tokens_before)
            return outcome
        if previous.get("mode") == "custom":
            write_state(**previous)  # the custom license keeps working; the caller shows the error
        else:
            write_state(mode="default", status="failed", reason=outcome.reason)
        return outcome


@dataclass
class CustomResult:
    ok: bool
    returncode: int
    output: str  # redacted stdout + stderr of the activation script
    reason: str  # ok | busy | timeout | invalid_serial | network | refused | limit | config | not_licensed | service | unknown
    verified: Optional[bool] = None


def run_custom_activation(cmd: list, env: dict, timeout: int = 120) -> CustomResult:
    """Run a user-initiated activatefre.sh command (License Setup page) safely.

    The service is restarted and ABBYY is asked whether the license really works.
    Whenever anything fails the previous configuration is restored, so a typo can
    never break the license that was working. Output is redacted so the default
    identifiers can never be displayed. Refused while a PDF conversion is running.
    """
    res = _run_custom_activation(cmd, env, timeout)
    if not res.ok and "--serial-number" in cmd:
        # Only a serial activation: it talks to the activation service the check looks at. A token (Online license)
        # uses other servers, so "the activation service answered" would say nothing about it.
        res.reason = refine_reason(res.reason)  # after the license lock was released: only talks to ABBYY's servers
    log("a custom license was activated from the License Setup page." if res.ok
        else f"custom license change failed ({res.reason}).")
    return res


def _run_custom_activation(cmd: list, env: dict, timeout: int) -> CustomResult:
    with _license_change() as blocked:
        if blocked:
            return CustomResult(False, 1, REASON_TEXT[blocked], blocked)
        snap = snapshot()
        tokens_before = _token_names(settings().licenses_dir)
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=env)
        except subprocess.TimeoutExpired:
            restore(snap)
            _remove_tokens(_token_names(settings().licenses_dir) - tokens_before)
            return CustomResult(False, 124, REASON_TEXT["timeout"], "timeout")
        output = redact(((r.stdout or "") + (r.stderr or "")).strip())
        outcome = classify(r.returncode, output)
        if outcome.ok:
            outcome, verdict = _finalize(snap)
        else:
            restore(snap)
            verdict = None
        if not outcome.ok:
            _remove_tokens(_token_names(settings().licenses_dir) - tokens_before)
            return CustomResult(False, r.returncode, output, outcome.reason, verdict)
        write_state(mode="custom", status="active", reason="ok", verified=verdict, checked_at=time.time())
        return CustomResult(True, 0, output, "ok", verdict)


def recheck() -> Optional[bool]:
    """Ask ABBYY now whether the current license works (the page's "Check now" button).

    None when it cannot be answered right now: another change or a PDF conversion is running
    (a busy license could otherwise be mistaken for a broken one).
    """
    with _license_change() as blocked:
        if blocked:
            return None
        verdict = probe_with_retries(attempts=2)
        st = read_state()
        if verdict is False and st.get("status") == "active":
            write_state(status="failed", reason="not_licensed", verified=False, checked_at=time.time())
        elif verdict is not None and st.get("mode") in ("default", "custom"):
            write_state(verified=verdict, checked_at=time.time())
        return verdict


# ── Container start-up job ────────────────────────────────────────────────────
_REFUSALS_BEFORE_GIVING_UP = 2  # ABBYY answered "no" this many attempts in a row: asking again only repeats the answer


def _license_survived_restart(previous: dict) -> bool:
    """True when an earlier container of this pod activated the default license and ABBYY still accepts it.

    The pod's /tmp (activation data, configuration and the state file) outlives a restarted container (crash, memory
    limit, failed health probe) but not a replaced pod. Asking ABBYY first costs a second and no pages, and avoids
    activating a serial number again, which may only be allowed a few activations. A new pod has an empty /tmp, so
    there is no earlier state and it activates as before. Only a definite yes counts: a probe that cannot tell
    (tool missing, timeout) leads to the normal activation.
    """
    if previous.get("mode") != "default" or previous.get("status") != "active":
        return False
    return probe_with_retries() is True


def autoactivate(sleep=None) -> int:
    """Activate the default license in the background; never raises."""
    sleep = sleep or time.sleep
    s = settings()
    try:
        if s.ls_host:
            write_state(mode="none", status="pending", reason="external")
            log("an external licensing service is configured; default license activation skipped.")
            return 0
        if default_license() is None:
            if default_misconfigured():
                write_state(mode="none", status="failed", reason="config")
                log("default license variables are incomplete or invalid; nothing activated.")
            else:
                log("no default license configured; nothing to activate.")
            return 0
        previous = read_state()  # what an earlier container of this pod left behind ({} in a new pod)
        write_state(mode="default", status="pending", reason="", attempt=0)
        if not wait_for_ls():
            log("licensing service did not start in time; will still try to activate.")
        if _license_survived_restart(previous):
            write_state(mode="default", status="active", reason="ok", attempt=0, verified=True, checked_at=time.time())
            log("the default license activated before this restart is still accepted by ABBYY; nothing was activated again.")
            return 0
        delays = (0,) + tuple(s.retry_delays)
        outcome = Outcome(False, "unknown", True)
        reached = False  # did the latest connection check find ABBYY's activation service reachable?
        refusals = 0  # failed attempts in a row where ABBYY's activation service answered
        for attempt, delay in enumerate(delays, start=1):
            if delay:
                write_state(mode="default", status="retrying", reason=outcome.reason, attempt=attempt - 1,
                            next_retry_at=time.time() + delay)
                sleep(delay)
            if read_state().get("mode") == "custom":
                log("a custom license was set up from the page; stopping automatic activation.")
                return 0
            log(f"activating default license (attempt {attempt}/{len(delays)})...")
            outcome = activate_default(respect_custom=True, retrying=True)
            if outcome.ok:
                log("default license is active." if outcome.reason == "ok" else "a custom license is in use; default left alone.")
                return 0
            log(f"attempt {attempt} failed: {outcome.reason}" + (f" (the tool ran {outcome.seconds:.0f} s)" if outcome.seconds else ""))
            for line in outcome.detail.splitlines():
                log(f"  ABBYY tool: {line}")
            if outcome.reason in _CONNECTION_REASONS and (attempt in (1, len(delays)) or refusals):
                # first failure (is the network up yet?), the last one, and every one while a refusal is suspected
                reached = _log_connection_check()
                if reached and outcome.reason == "network":
                    log("ABBYY's activation service answered, so this is not a connection problem: ABBYY refused the activation.")
            if outcome.reason == "network" and reached:
                outcome.reason = "refused"
            refusals = refusals + 1 if outcome.reason == "refused" else 0
            if refusals >= _REFUSALS_BEFORE_GIVING_UP:
                outcome.transient = False
                log(f"ABBYY refused the activation {refusals} times in a row; not asking again automatically. "
                    "Use the button on the License Setup page after ABBYY has changed something.")
            if not outcome.transient:
                break
        write_state(mode="default", status="failed", reason=outcome.reason)
        log(f"default license could not be activated ({outcome.reason}).")
        return 1
    except Exception as exc:  # the container must keep serving the UI no matter what
        log(f"unexpected error during automatic activation: {type(exc).__name__}")
        try:
            write_state(mode="default", status="failed", reason="unknown")
        except Exception:
            pass
        return 1


def _boot_ts() -> Optional[float]:
    try:
        return float((settings().state_dir / "boot_ts").read_text().strip())
    except (OSError, ValueError):
        return None


def _container_starts() -> Optional[int]:
    """How many times the application started in this pod (docker_entrypoint.sh counts them in /tmp)."""
    try:
        return int((settings().state_dir / "start_count").read_text().strip())
    except (OSError, ValueError):
        return None


def uptime_text(now: Optional[float] = None) -> Optional[str]:
    """One plain sentence on how long this server has been running, for the License Setup page (None when unknown).

    Shows whether the license is being set up again because the server just restarted. Holds no license details.
    """
    started = _boot_ts()
    if not started:
        return None
    seconds = max(0, int((now if now is not None else time.time()) - started))
    days, rest = divmod(seconds, 86400)
    hours, rest = divmod(rest, 3600)
    minutes = rest // 60
    if days:
        span = f"{days} d {hours} h"
    elif hours:
        span = f"{hours} h {minutes} min"
    else:
        span = f"{minutes} min" if minutes else "less than a minute"
    text = f"This server has been running for {span} (started {time.strftime('%Y-%m-%d %H:%M', time.gmtime(started))} UTC)."
    starts = _container_starts()
    if starts and starts > 1:
        text += (f" The application has restarted {starts - 1} time{'s' if starts > 2 else ''} on this server "
                 "since the server itself was created.")
    return text


# ── View model for the UI ─────────────────────────────────────────────────────
@dataclass
class View:
    kind: str  # default_active | activating | retrying | default_failed | custom_active | misconfigured | external | none
    default_configured: bool
    reason: str = ""
    ls_running: bool = False
    next_retry_in: int = 0
    verified: Optional[bool] = None


def describe() -> View:
    st = read_state()
    configured = default_configured()
    mode, status, reason = st.get("mode", ""), st.get("status", ""), st.get("reason", "")
    if not st and configured and not settings().ls_host:
        # Nothing written yet: normal for a few seconds after start, a failure if it never starts.
        boot = _boot_ts()
        mode = "default"
        status, reason = ("pending", "") if boot and time.time() - boot < _NOT_STARTED_AFTER else ("failed", "not_started")
    if status in ("pending", "activating", "retrying") and time.time() - st.get("ts", time.time()) > _STALE_AFTER:
        status, reason = "failed", "unknown"  # the background job died; do not show "setting up" forever
    running = ls_running()
    if mode == "custom" and status == "active":
        kind = "custom_active"
    elif settings().ls_host:
        kind = "external"
    elif configured:
        kind = {"active": "default_active", "failed": "default_failed", "retrying": "retrying"}.get(status, "activating")
        if mode != "default":
            kind = "activating"
    elif default_misconfigured():
        kind = "misconfigured"
    else:
        kind = "none"
    nxt = max(0, int(st.get("next_retry_at", 0) - time.time())) if kind == "retrying" else 0
    return View(kind, configured, reason, running, nxt, st.get("verified"))


def reason_text(reason: str, final: bool = False) -> str:
    if final and reason in FINAL_REASON_TEXT:
        return FINAL_REASON_TEXT[reason]
    return REASON_TEXT.get(reason, REASON_TEXT["unknown"])


def ui_message(view: View) -> Optional[tuple]:
    """(level, title, body) for the banner, or None when the page should look as it always did."""
    k = view.kind
    if k in ("default_active", "custom_active") and not view.ls_running and not settings().skip_ls_restart:
        return ("warning", "The ABBYY license is set up, but its license service is not running.",
                "PDF conversion will not work until it is started again. Use the **Restart LicensingService** button "
                "further down this page.")
    if k == "default_active":
        return ("success", "ABBYY license is already set up. There is nothing for you to do.",
                "Your organization's production ABBYY license is configured on this server and ready to use. "
                "Just upload your PDF on the main page and convert it. "
                "The options further down this page are only for special cases, for example a team that has its own ABBYY license.")
    if k in ("activating", "retrying"):
        extra = ""
        if k == "retrying":
            extra = (f" The first attempt did not go through. {reason_text(view.reason)} "
                     "The server is trying again automatically.")
        return ("info", "Setting up the ABBYY license...",
                "This happens automatically when the server starts and usually takes about a minute. "
                "You do not need to enter anything. PDF conversion will be available as soon as it finishes "
                "(uploading a Word/DOCX file works right away)." + extra)
    if k == "default_failed":
        pointer = ""
        if view.reason in ("network", "timeout"):
            pointer = " The support team can use the **Connection check** at the bottom of this page to see which step fails."
        elif view.reason == "refused":
            pointer = (" The support team should ask ABBYY to check how many activations are left on the serial number, "
                       "or to provide an Online (token) license, which is made for containers.")
        return ("error", "The default ABBYY license could not be set up automatically.",
                f"{reason_text(view.reason, final=True)} This is not something you did wrong. Please contact the application "
                f"support team and mention the code **{view.reason or 'unknown'}**.{pointer} "
                "Until it is fixed you can use your own license with the options below.")
    if k == "custom_active":
        if view.default_configured:
            return ("warning", "This server is using a different ABBYY license (set up from this page).",
                    "Everyone using this server is currently using that license. The company's default license "
                    "comes back automatically when the server restarts, or right now with the button below.")
        return ("info", "An ABBYY license was set up from this page.",
                "It will be lost when the server restarts because no default license is configured for this server.")
    if k == "misconfigured":
        return ("warning", "The default ABBYY license settings on this server are incomplete.",
                "Please contact the application support team. Meanwhile you can use your own license with the options below.")
    if k == "external":
        return ("info", "ABBYY licensing is provided by a central license server.",
                "There is nothing to set up on this page.")
    return None


def conversion_hint(tool_output: str = "") -> Optional[str]:
    """Plain-language line for under a failed PDF conversion when the license is the likely cause."""
    view = describe()
    if view.kind == "default_failed":
        return (f"The default ABBYY license could not be set up on this server (code: {view.reason or 'unknown'}). "
                "Please contact the application support team and mention this code. "
                "If you have your own ABBYY license you can use it on the **License Setup** page.")
    if view.kind == "misconfigured":
        return "The default ABBYY license settings on this server are incomplete. Please contact the application support team."
    if _NOT_LICENSED.search(tool_output or ""):
        return ("ABBYY reports that no usable license is set up on this server. "
                "Please open the **License Setup** page or contact the application support team.")
    return None


# ── CLI ───────────────────────────────────────────────────────────────────────
def main(argv: list) -> int:
    cmd = argv[1] if len(argv) > 1 else "status"
    if cmd == "autoactivate":
        return autoactivate()
    if cmd == "status":
        v = describe()
        print(json.dumps({"kind": v.kind, "default_configured": v.default_configured,
                          "reason": v.reason, "ls_running": v.ls_running}))
        return 0
    print("usage: abbyy_license.py autoactivate|status", file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
