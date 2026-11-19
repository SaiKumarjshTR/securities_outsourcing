"""Unit tests for abbyy_license.py.

They run the real module against a stub activatefre.sh that behaves like ABBYY's
script (verified against the real one): it rewrites SamplesConfig.h BEFORE it
tries to activate, and it never echoes the serial number. All identifiers here
are made up.
"""
import json
import os
import shutil
import socket
import ssl
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import abbyy_license as L  # noqa: E402

REAL_DIAGNOSE = L.diagnose_connection  # the fixture below replaces it so no test touches the network

SERIAL = "TESTSERIAL0000000000TEST"
PROJECT = "TestDefaultProject01"
CUSTOM_SERIAL = "CUSTOMSERIAL000000000CUST"
CUSTOM_PROJECT = "TestCustomProject02"

CONFIG_TMPL = """// auto-generated
static wchar_t* GetFreProjectId() {{
        static wchar_t FreProjectId[] = L"{project}";
        return FreProjectId;
}}
static wchar_t* GetLicensePath() {{
        static wchar_t LicensePath[] = L"{path}";
        return LicensePath;
}}
static wchar_t* GetLicensePassword() {{
        static wchar_t LicensePassword[] = L"{pw}";
        return LicensePassword;
}}
"""

STUB = r"""#!/bin/bash
echo "ABBYY FineReader Engine 12 activation script"
S=""; P=""
while [ $# -gt 0 ]; do
  case "$1" in --serial-number) S="$2"; shift;; --project-id) P="$2"; shift;; esac
  shift
done
echo x >> "$STUB_COUNT_FILE"
n=$(wc -l < "$STUB_COUNT_FILE")
cat > "$STUB_FRE_DIR/Samples/SamplesConfig.h" <<EOF
static wchar_t* GetFreProjectId() {
        static wchar_t FreProjectId[] = L"$P";
        return FreProjectId;
}
static wchar_t* GetLicensePath() {
        static wchar_t LicensePath[] = L"";
        return LicensePath;
}
static wchar_t* GetLicensePassword() {
        static wchar_t LicensePassword[] = L"";
        return LicensePassword;
}
EOF
echo "Configuring samples - ok"
echo "Activating license..."
mode="$STUB_MODE"
if [ -n "$STUB_FAIL_FIRST" ] && [ "$n" -le "$STUB_FAIL_FIRST" ]; then mode="network"; fi
case "$mode" in
  ok) echo "Activation completed."; exit 0;;
  invalid) echo "Incorrect serial number"; echo "Activation has failed."; exit 1;;
  network) echo "Could not connect to the activation server"; echo "Activation has failed."; exit 1;;
  leak) echo "serial=$S project=$P"; exit 1;;
  hang) sleep 30;;
esac
"""


# Stand-in for ABBYY's command-line tool. It reads the same ABBYY_* variables the app
# exports and prints what the REAL tool printed when measured: "Error: ABBYY FineReader
# Engine is not licensed." for an unknown project id and "Error: There are no licenses
# on 127.0.0.1" for an empty one. Ids listed in FAKE_CLI_OK_IDS count as licensed.
FAKE_CLI = r"""#!/bin/bash
echo "ABBYY FineReader Engine 12 Sample"
echo "(c) 2022 ABBYY Development Inc."
echo "x" >> "$FAKE_CLI_COUNT_FILE"
case "$FAKE_CLI_MODE" in
  hang) sleep 30;;
  other_error) echo "Error: Cannot open the input file."; exit 1;;
  crash) exit 139;;
  flaky) n=$(wc -l < "$FAKE_CLI_COUNT_FILE"); if [ "$n" -le 2 ]; then echo "Error: There are no licenses on 127.0.0.1"; exit 1; fi;;
esac
if [ -z "$ABBYY_PROJECT_ID" ]; then echo "Error: There are no licenses on 127.0.0.1"; exit 1; fi
case " $FAKE_CLI_OK_IDS " in
  *" $ABBYY_PROJECT_ID "*) echo "No input file specified."; exit 1;;
esac
echo "Error: ABBYY FineReader Engine is not licensed."; exit 1
"""


def cfg_values(fre: Path) -> dict:
    import re
    t = (fre / "Samples" / "SamplesConfig.h").read_text()
    out = {}
    for n in ("GetFreProjectId", "GetLicensePath", "GetLicensePassword"):
        out[n] = re.search(n + r'\(\)\s*\{\s*static wchar_t \w+\[\]\s*=\s*L"([^"]*)"', t).group(1)
    return out


@pytest.fixture
def box(tmp_path, monkeypatch):
    fre = tmp_path / "fre"
    for d in ("Samples/Java", "Bin", "CommonBin/Licensing"):
        (fre / d).mkdir(parents=True)
    (fre / "Samples" / "SamplesConfig.h").write_text(CONFIG_TMPL.format(project="OriginalProject", path="", pw=""))
    (fre / "Samples" / "Java" / "SamplesConfig.java").write_text("original-java")
    (fre / "Bin" / "LicensingSettings.xml").write_text("<original-bin/>")
    (fre / "CommonBin" / "Licensing" / "LicensingSettings.xml").write_text("<original-common/>")
    stub = tmp_path / "activatefre.sh"
    stub.write_text(STUB)
    stub.chmod(0o755)
    count = tmp_path / "count.txt"
    count.write_text("")
    cli = tmp_path / "CommandLineInterface"
    cli.write_text(FAKE_CLI)
    cli.chmod(0o755)
    cli_count = tmp_path / "cli_count.txt"
    cli_count.write_text("")
    state = tmp_path / "state"
    env = {
        "ABBYY_FRE_DIR": str(fre), "ABBYY_ACTIVATE_SH": str(stub), "ABBYY_STATE_DIR": str(state),
        "ABBYY_SKIP_LS_RESTART": "1", "ABBYY_LS_BIN": str(tmp_path / "no-such-LicensingService"),
        "STUB_MODE": "ok", "STUB_COUNT_FILE": str(count), "STUB_FRE_DIR": str(fre),
        "ABBYY_CLI": str(cli), "ABBYY_LIB": str(tmp_path), "FAKE_CLI_COUNT_FILE": str(cli_count),
        "FAKE_CLI_OK_IDS": f"{PROJECT} {CUSTOM_PROJECT}", "FAKE_CLI_MODE": "",
        "ABBYY_SERIAL_NUMBER": SERIAL, "ABBYY_CUSTOMER_PROJECT_ID": PROJECT,
        "ABBYY_AUTOACTIVATE_DELAYS": "1,1",
    }
    for k, v in env.items():
        monkeypatch.setenv(k, v)
    monkeypatch.delenv("ABBYY_LS_HOST", raising=False)

    class Box:
        pass
    b = Box()
    b.fre, b.state, b.count, b.tmp = fre, state, count, tmp_path
    b.attempts = lambda: len(count.read_text().splitlines())
    b.probes = lambda: len(cli_count.read_text().splitlines())
    monkeypatch.setattr(L.time, "sleep", lambda s: None)  # the module's real waits are not what is under test
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: [])  # never touch the real network
    b.mode = lambda m: monkeypatch.setenv("STUB_MODE", m)
    b.env = monkeypatch.setenv
    b.unset = monkeypatch.delenv
    return b


# ── default license from the environment ──────────────────────────────────────
def test_default_license_requires_both_values(box):
    assert L.default_license() == L.DefaultLicense(SERIAL, PROJECT)
    box.unset("ABBYY_CUSTOMER_PROJECT_ID")
    assert L.default_license() is None and L.default_configured() is False
    assert L.default_misconfigured() is True


def test_unsafe_values_are_rejected_not_passed_to_a_shell(box):
    box.env("ABBYY_SERIAL_NUMBER", "ABC; touch /tmp/pwned")
    assert L.default_license() is None
    assert L.default_misconfigured() is True


def test_nothing_configured_is_not_a_misconfiguration(box):
    box.unset("ABBYY_SERIAL_NUMBER")
    box.unset("ABBYY_CUSTOMER_PROJECT_ID")
    assert L.default_license() is None and L.default_misconfigured() is False


# ── redaction ────────────────────────────────────────────────────────────────
def test_redact_hides_serial_project_and_common_variants(box):
    dashed = "-".join(SERIAL[i:i + 4] for i in range(0, 24, 4))
    text = f"a {SERIAL} b {SERIAL.lower()} c {dashed} d {PROJECT} e {PROJECT.upper()} f"
    out = L.redact(text)
    for secret in (SERIAL, SERIAL.lower(), dashed, PROJECT, PROJECT.upper()):
        assert secret not in out
    assert out.count("[hidden]") == 5 and out.startswith("a ") and out.endswith(" f")


def test_redact_leaves_other_text_alone_and_accepts_extra_secrets(box):
    assert L.redact("Incorrect serial number") == "Incorrect serial number"
    assert "EXTRASECRET1" not in L.redact("x EXTRASECRET1 y", extra=("EXTRASECRET1",))


def test_log_never_prints_identifiers(box, capsys):
    L.log(f"serial {SERIAL} project {PROJECT}")
    out = capsys.readouterr().out
    assert SERIAL not in out and PROJECT not in out and "[abbyy-license]" in out


# ── classification (messages measured on the real tool) ───────────────────────
@pytest.mark.parametrize("rc,text,timed,reason,transient", [
    (0, "Activation completed.", False, "ok", False),
    (1, "Incorrect serial number\nActivation has failed.", False, "invalid_serial", False),
    (1, "Could not connect to the activation server", False, "network", True),
    (1, "SSL certificate problem: unable to get local issuer certificate", False, "network", True),
    (1, "Activation limit has been reached for this serial", False, "limit", False),
    (1, "Configuring samples - failed: Read-only file system", False, "config", False),
    (1, "something nobody has seen before", False, "unknown", True),
    (1, "", True, "timeout", True),
])
def test_classify(rc, text, timed, reason, transient):
    o = L.classify(rc, text, timed_out=timed)
    assert (o.ok, o.reason, o.transient) == (reason == "ok", reason, transient if reason != "ok" else False)


# ── config safety: the failure mode observed on the real script ───────────────
def test_failed_activation_restores_the_configuration(box):
    before = cfg_values(box.fre)
    box.mode("invalid")
    o = L._activate_serial("BADSERIAL", "BadProject")
    assert (o.ok, o.reason) == (False, "invalid_serial")
    assert cfg_values(box.fre) == before  # the stub rewrote it first, like the real script does
    assert (box.fre / "Bin" / "LicensingSettings.xml").read_text() == "<original-bin/>"


def test_successful_activation_keeps_the_new_configuration(box):
    o = L._activate_serial(SERIAL, PROJECT)
    assert o.ok and cfg_values(box.fre)["GetFreProjectId"] == PROJECT


def test_timeout_is_transient_and_restores_the_configuration(box):
    box.mode("hang")
    before = cfg_values(box.fre)
    o = L._activate_serial(SERIAL, PROJECT, timeout=1)
    assert (o.ok, o.reason, o.transient) == (False, "timeout", True)
    assert cfg_values(box.fre) == before


# ── default activation ───────────────────────────────────────────────────────
def test_activate_default_success_writes_state_without_identifiers(box):
    o = L.activate_default()
    assert o.ok
    st = L.read_state()
    assert (st["mode"], st["status"]) == ("default", "active")
    raw = (box.state / "license_state.json").read_text()
    assert SERIAL not in raw and PROJECT not in raw
    assert L.describe().kind == "default_active"


def test_activate_default_failure_marks_failed_and_keeps_old_config(box):
    box.mode("invalid")
    before = cfg_values(box.fre)
    o = L.activate_default()
    assert (o.ok, o.reason) == (False, "invalid_serial")
    st = L.read_state()
    assert (st["status"], st["reason"]) == ("failed", "invalid_serial")
    assert cfg_values(box.fre) == before


def test_second_change_while_one_is_running_is_refused(box):
    with L._lock() as got:
        assert got
        o = L.activate_default()
    assert (o.ok, o.reason, o.transient) == (False, "busy", True)
    assert box.attempts() == 0


# ── start-up job ──────────────────────────────────────────────────────────────
def test_autoactivate_retries_transient_errors_then_succeeds(box):
    box.env("STUB_FAIL_FIRST", "2")
    slept = []
    assert L.autoactivate(sleep=slept.append) == 0
    assert box.attempts() == 3 and slept == [1, 1]
    assert L.read_state()["status"] == "active"


def test_autoactivate_does_not_retry_a_wrong_serial(box):
    box.mode("invalid")
    assert L.autoactivate(sleep=lambda s: None) == 1
    assert box.attempts() == 1  # never hammer ABBYY with a serial it rejected
    st = L.read_state()
    assert (st["status"], st["reason"]) == ("failed", "invalid_serial")


def test_autoactivate_gives_up_after_the_schedule(box):
    box.mode("network")
    assert L.autoactivate(sleep=lambda s: None) == 1
    assert box.attempts() == 3  # 1 try + 2 scheduled retries
    assert L.read_state()["reason"] == "network"


def test_autoactivate_never_overwrites_a_custom_license(box):
    box.env("STUB_FAIL_FIRST", "1")
    box.env("ABBYY_AUTOACTIVATE_DELAYS", "1,1,1")

    def user_sets_custom_license(_seconds):
        L.write_state(mode="custom", status="active", reason="ok")

    assert L.autoactivate(sleep=user_sets_custom_license) == 0
    assert box.attempts() == 1 and L.read_state()["mode"] == "custom"


def test_background_attempt_that_loses_the_race_to_a_custom_license_does_nothing(box):
    # Custom license is set up AFTER the job's own check but BEFORE it takes the lock.
    L.write_state(mode="custom", status="active", reason="ok")
    before = cfg_values(box.fre)
    o = L.activate_default(respect_custom=True)
    assert o.ok and o.reason == "custom"
    assert box.attempts() == 0 and cfg_values(box.fre) == before
    assert L.read_state()["mode"] == "custom"
    # a deliberate, user-initiated call (no respect_custom) is allowed to take over
    assert L.activate_default().ok and L.read_state()["mode"] == "default"


def test_autoactivate_without_configuration_does_nothing(box, capsys):
    box.unset("ABBYY_SERIAL_NUMBER")
    box.unset("ABBYY_CUSTOMER_PROJECT_ID")
    assert L.autoactivate() == 0
    assert box.attempts() == 0 and L.read_state() == {}
    assert "no default license configured" in capsys.readouterr().out


def test_autoactivate_with_incomplete_configuration_reports_it(box):
    box.unset("ABBYY_CUSTOMER_PROJECT_ID")
    assert L.autoactivate() == 0
    assert L.read_state()["reason"] == "config" and box.attempts() == 0
    assert L.describe().kind == "misconfigured"


def test_autoactivate_skips_when_an_external_license_server_is_used(box):
    box.env("ABBYY_LS_HOST", "licenses.example.invalid")
    assert L.autoactivate() == 0 and box.attempts() == 0
    assert L.describe().kind == "external"


def test_autoactivate_output_never_contains_identifiers_even_if_the_script_leaks(box, capsys):
    box.mode("leak")  # this stub prints the serial and project id
    L.autoactivate(sleep=lambda s: None)
    out = capsys.readouterr().out
    assert SERIAL not in out and PROJECT not in out


# ── a container that restarts inside the same pod keeps its activation ────────
# The pod's /tmp (state file, configuration, activation data) survives a restarted container, so the start-up job
# asks ABBYY first and only activates when ABBYY no longer accepts what is there. A new pod has an empty /tmp.
def _first_container(box):
    assert L.autoactivate(sleep=lambda s: None) == 0
    assert box.attempts() == 1 and L.read_state()["status"] == "active"


def test_a_restarted_container_keeps_the_activated_license_and_asks_abbyy_instead_of_activating(box, capsys):
    _first_container(box)
    probes = box.probes()
    box.mode("network")  # an activation would fail now: the point is that none is tried
    assert L.autoactivate(sleep=lambda s: None) == 0
    assert box.attempts() == 1  # the serial number was not activated a second time
    assert box.probes() > probes  # ABBYY itself was asked whether the license works
    st = L.read_state()
    assert (st["mode"], st["status"], st["reason"], st["verified"]) == ("default", "active", "ok", True)
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT
    out = capsys.readouterr().out
    assert "still accepted by ABBYY; nothing was activated again" in out
    assert SERIAL not in out and PROJECT not in out


def test_every_further_restart_of_the_same_pod_also_keeps_it(box):
    _first_container(box)
    for _ in range(3):
        assert L.autoactivate(sleep=lambda s: None) == 0
    assert box.attempts() == 1 and L.read_state()["status"] == "active"


def test_a_restarted_container_activates_again_when_abbyy_no_longer_accepts_the_license(box):
    _first_container(box)
    (box.fre / "Samples" / "SamplesConfig.h").write_text(CONFIG_TMPL.format(project="OriginalProject", path="", pw=""))
    assert L.autoactivate(sleep=lambda s: None) == 0
    assert box.attempts() == 2  # the earlier activation is gone: this is the normal activation
    assert (L.read_state()["mode"], L.read_state()["status"]) == ("default", "active")
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT


@pytest.mark.parametrize("cli_mode", ["other_error", "crash"])
def test_a_probe_that_cannot_tell_does_not_skip_the_activation(box, cli_mode):
    _first_container(box)
    box.env("FAKE_CLI_MODE", cli_mode)  # the tool answers with something that is neither "licensed" nor "not licensed"
    assert L.autoactivate(sleep=lambda s: None) == 0
    assert box.attempts() == 2  # only a definite yes from ABBYY counts


@pytest.mark.parametrize("earlier", [
    {},  # a new pod: empty /tmp, no state
    {"mode": "default", "status": "failed", "reason": "refused"},
    {"mode": "default", "status": "retrying", "reason": "network", "attempt": 2},
    {"mode": "default", "status": "activating", "reason": ""},  # the container died while it was activating
    {"mode": "custom", "status": "active", "reason": "ok"},  # an override does not survive: the default comes back
])
def test_without_an_active_default_license_from_before_the_start_activates_as_before(box, earlier):
    # A license ABBYY would accept sits in the configuration, yet it must not be reused here: it is not the default one.
    (box.fre / "Samples" / "SamplesConfig.h").write_text(CONFIG_TMPL.format(project=CUSTOM_PROJECT, path="", pw=""))
    if earlier:
        L.write_state(**earlier)
    assert L.autoactivate(sleep=lambda s: None) == 0
    assert box.attempts() == 1
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT and L.read_state()["status"] == "active"


def test_a_restart_keeps_the_license_state_file_free_of_identifiers(box):
    _first_container(box)
    L.autoactivate(sleep=lambda s: None)
    raw = (box.state / "license_state.json").read_text()
    assert SERIAL not in raw and PROJECT not in raw


# ── ABBYY says "no": how often the start-up job asks ─────────────────────────
def test_the_start_up_job_stops_after_abbyy_refused_twice_in_a_row(box, monkeypatch, capsys):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached())
    box.env("ABBYY_AUTOACTIVATE_DELAYS", "1,1,1,1,1,1")  # the real schedule has six retries
    box.mode("network")  # the tool prints its generic line: the check tells that ABBYY answered
    writes, seen = [], []
    real_write = L.write_state
    monkeypatch.setattr(L, "write_state", lambda **kw: (writes.append((kw.get("status"), kw.get("reason"))), real_write(**kw))[1])
    assert L.autoactivate(sleep=lambda s: seen.append(dict(L.read_state()))) == 1
    assert box.attempts() == 2  # a third request would only repeat ABBYY's answer
    assert [(s["status"], s["reason"]) for s in seen] == [("retrying", "refused")]  # shown while the second try is awaited
    assert [w for w in writes if w[0] == "failed"] == [("failed", "refused")]  # red only at the end, never in between
    assert (L.read_state()["status"], L.read_state()["reason"]) == ("failed", "refused")
    out = capsys.readouterr().out
    assert "attempt 1 failed: network" in out  # the tool's own classification is logged as it was
    assert out.count("ABBYY's activation service answered, so this is not a connection problem") == 2  # a fresh look each time
    assert "refused the activation 2 times in a row" in out
    assert "default license could not be activated (refused)" in out
    for secret in (SERIAL, PROJECT):
        assert secret not in out


def test_one_refusal_is_not_final_abbyy_may_say_yes_the_next_time(box, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached())
    box.env("STUB_FAIL_FIRST", "1")  # the first request fails with the generic line, the second one works
    assert L.autoactivate(sleep=lambda s: None) == 0
    assert box.attempts() == 2 and L.read_state()["status"] == "active"


def test_a_refusal_followed_by_a_dead_network_does_not_stop_the_retries(box, monkeypatch):
    def answers():
        yield _reached()  # attempt 1: ABBYY answered
        while True:
            yield _blocked()  # from then on nothing gets through

    gen = answers()
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: next(gen))
    box.env("ABBYY_AUTOACTIVATE_DELAYS", "1,1,1")
    box.mode("network")
    seen = []
    assert L.autoactivate(sleep=lambda s: seen.append(L.read_state()["reason"])) == 1
    assert box.attempts() == 4  # the whole schedule: refused once is not refused twice in a row
    assert seen == ["refused", "network", "network"] and L.read_state()["reason"] == "network"


def test_two_refusals_that_are_not_in_a_row_do_not_stop_the_job(box, monkeypatch):
    def answers():
        yield _reached()  # attempt 1: ABBYY answered
        yield _blocked()  # attempt 2: nothing gets through (the check looks twice)
        yield _blocked()
        while True:
            yield _reached()  # the last attempt: ABBYY answers again

    gen = answers()
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: next(gen))
    box.env("ABBYY_AUTOACTIVATE_DELAYS", "1,1,1")
    box.mode("network")
    assert L.autoactivate(sleep=lambda s: None) == 1
    assert box.attempts() == 4  # the network problem in between broke the row, so the whole schedule was used
    assert L.read_state()["reason"] == "refused"


def test_other_transient_errors_keep_the_whole_schedule(box):
    box.mode("network")  # no connection check result: this stays a plain network problem
    box.env("ABBYY_AUTOACTIVATE_DELAYS", "1,1,1,1")
    assert L.autoactivate(sleep=lambda s: None) == 1
    assert box.attempts() == 5 and L.read_state()["reason"] == "network"


# ── how long the server has been running (shown on the License Setup page) ────
def _boot(box, after=0, starts=None):
    box.state.mkdir(parents=True, exist_ok=True)
    (box.state / "boot_ts").write_text("1000000000")  # 2001-09-09 01:46:40 UTC
    if starts is not None:
        (box.state / "start_count").write_text(str(starts))
    return 1000000000 + after


def test_uptime_text_says_how_long_the_server_runs_and_when_it_started(box):
    assert L.uptime_text(now=_boot(box, 2 * 3600 + 14 * 60 + 5)) == \
        "This server has been running for 2 h 14 min (started 2001-09-09 01:46 UTC)."
    assert L.uptime_text(now=_boot(box, 7 * 60 + 59)) == "This server has been running for 7 min (started 2001-09-09 01:46 UTC)."
    assert "less than a minute" in L.uptime_text(now=_boot(box, 30))
    assert "3 d 5 h" in L.uptime_text(now=_boot(box, 3 * 86400 + 5 * 3600 + 59 * 60))


def test_uptime_text_counts_restarts_inside_the_same_server_only_when_there_were_some(box):
    now = _boot(box, 600, starts=1)
    assert "restarted" not in L.uptime_text(now=now)
    assert L.uptime_text(now=_boot(box, 600, starts=2)).endswith(
        " The application has restarted 1 time on this server since the server itself was created.")
    assert "restarted 4 times on this server" in L.uptime_text(now=_boot(box, 600, starts=5))


def test_uptime_text_is_silent_when_the_start_time_is_unknown_and_never_raises(box):
    assert L.uptime_text() is None  # no file yet
    box.state.mkdir(parents=True, exist_ok=True)
    (box.state / "boot_ts").write_text("not a number")
    assert L.uptime_text() is None
    _boot(box)
    (box.state / "start_count").write_text("garbage")
    assert "restarted" not in L.uptime_text(now=1000000100)
    assert L.uptime_text(now=1)  # a clock that is behind the start time must not give a negative duration


def test_uptime_text_holds_no_license_details(box):
    now = _boot(box, 600, starts=3)
    text = L.uptime_text(now=now)
    assert SERIAL not in text and PROJECT not in text and "token" not in text.lower()


# ── override and switch back ──────────────────────────────────────────────────
def _custom_cmd(box, serial=CUSTOM_SERIAL, project=CUSTOM_PROJECT):
    return ["bash", os.environ["ABBYY_ACTIVATE_SH"], "--serial-number", serial,
            "--project-id", project, "--skip-local-service-installation"]


def test_failed_override_does_not_break_the_working_default(box):
    assert L.activate_default().ok
    box.mode("invalid")
    r = L.run_custom_activation(_custom_cmd(box), dict(os.environ))
    assert not r.ok and r.reason == "invalid_serial"
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT  # restored
    assert (L.read_state()["mode"], L.read_state()["status"]) == ("default", "active")


def test_successful_override_then_switch_back_needs_no_new_activation(box):
    assert L.activate_default().ok
    r = L.run_custom_activation(_custom_cmd(box), dict(os.environ))
    assert r.ok and r.returncode == 0
    assert cfg_values(box.fre)["GetFreProjectId"] == CUSTOM_PROJECT
    assert L.describe().kind == "custom_active"
    attempts = box.attempts()
    o = L.switch_back_to_default()
    assert o.ok and cfg_values(box.fre)["GetFreProjectId"] == PROJECT
    assert box.attempts() == attempts  # restored from the saved default, ABBYY not contacted
    assert L.describe().kind == "default_active"


def test_switch_back_activates_again_when_no_saved_default_exists(box):
    L.write_state(mode="custom", status="active", reason="ok")
    o = L.switch_back_to_default()
    assert o.ok and box.attempts() == 1
    assert L.describe().kind == "default_active"


def test_failed_switch_back_keeps_the_custom_license_working(box):
    assert L.activate_default().ok
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok
    shutil.rmtree(box.state / "default_config")
    box.mode("invalid")
    o = L.switch_back_to_default()
    assert (o.ok, o.reason) == (False, "invalid_serial")
    assert cfg_values(box.fre)["GetFreProjectId"] == CUSTOM_PROJECT
    assert L.describe().kind == "custom_active"


def test_custom_activation_output_is_redacted(box):
    box.mode("leak")
    r = L.run_custom_activation(_custom_cmd(box, serial=SERIAL, project=PROJECT), dict(os.environ))
    assert not r.ok and r.returncode != 0
    assert SERIAL not in r.output and PROJECT not in r.output and "[hidden]" in r.output


# ── asking ABBYY itself (the probe) ──────────────────────────────────────────
def test_probe_says_yes_for_a_working_project_id(box):
    L.activate_default()
    assert L.probe_license() is True
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT


def test_probe_says_no_for_an_unknown_project_id_and_for_an_empty_one(box):
    box.env("FAKE_CLI_OK_IDS", "SomethingElse")
    assert L.probe_license() is False
    box.env("FAKE_CLI_OK_IDS", PROJECT)
    (box.fre / "Samples" / "SamplesConfig.h").write_text(CONFIG_TMPL.format(project="", path="", pw=""))
    assert L.probe_license() is False


def test_probe_cannot_tell_when_the_tool_is_missing_hangs_or_fails_for_other_reasons(box):
    for mode in ("other_error", "crash"):
        box.env("FAKE_CLI_MODE", mode)
        assert L.probe_license() is None, mode
    box.env("FAKE_CLI_MODE", "hang")
    assert L.probe_license(timeout=1) is None
    box.env("ABBYY_CLI", str(box.tmp / "missing"))
    assert L.probe_license() is None


def test_probe_waits_for_a_service_that_is_still_starting(box):
    L._activate_serial(SERIAL, PROJECT)  # puts a licensed project id into the configuration
    box.env("FAKE_CLI_MODE", "flaky")
    slept = []
    assert L.probe_with_retries(attempts=3, pause=2, sleep=slept.append) is True
    assert box.probes() == 3 and slept == [2, 2]


def test_probe_gives_up_after_its_retries(box):
    box.env("FAKE_CLI_OK_IDS", "SomethingElse")
    assert L.probe_with_retries(attempts=2, pause=0, sleep=lambda s: None) is False
    assert box.probes() == 2


def test_valid_serial_with_wrong_project_id_is_caught_and_rolled_back(box):
    # The script accepts it (exit 0) but ABBYY would refuse every conversion.
    box.env("FAKE_CLI_OK_IDS", "SomeOtherProject")
    before = cfg_values(box.fre)
    o = L.activate_default()
    assert (o.ok, o.reason, o.transient) == (False, "not_licensed", False)
    assert cfg_values(box.fre) == before
    st = L.read_state()
    assert (st["status"], st["reason"]) == ("failed", "not_licensed")


def test_custom_license_that_abbyy_does_not_accept_is_rolled_back(box):
    assert L.activate_default().ok
    box.env("FAKE_CLI_OK_IDS", PROJECT)  # custom project id is not licensed
    r = L.run_custom_activation(_custom_cmd(box), dict(os.environ))
    assert (r.ok, r.reason, r.verified) == (False, "not_licensed", False)
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT
    assert L.describe().kind == "default_active"


def test_service_that_does_not_come_back_rolls_the_activation_back(box, monkeypatch):
    before = cfg_values(box.fre)
    monkeypatch.setattr(L, "restart_ls", lambda *a, **k: False)
    o = L.activate_default()
    assert (o.ok, o.reason, o.transient) == (False, "service", True)
    assert cfg_values(box.fre) == before
    assert (L.read_state()["status"], L.read_state()["reason"]) == ("failed", "service")
    assert box.probes() == 0  # never asks ABBYY while the service is down


def test_service_is_restarted_after_every_successful_activation(box, monkeypatch):
    calls = []
    monkeypatch.setattr(L, "restart_ls", lambda *a, **k: calls.append(1) or True)
    assert L.activate_default().ok and len(calls) == 1
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok and len(calls) == 2
    assert L.switch_back_to_default().ok and len(calls) == 3  # even the offline restore needs it


def test_service_that_does_not_come_back_after_a_custom_activation_keeps_the_old_license(box, monkeypatch):
    assert L.activate_default().ok
    monkeypatch.setattr(L, "restart_ls", lambda *a, **k: False)
    r = L.run_custom_activation(_custom_cmd(box), dict(os.environ))
    assert (r.ok, r.reason) == (False, "service")
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT
    assert L.describe().kind == "default_active"


def test_inconclusive_probe_does_not_block_a_successful_activation(box):
    box.env("FAKE_CLI_MODE", "other_error")
    assert L.activate_default().ok
    assert L.read_state()["status"] == "active" and L.read_state()["verified"] is None


def test_switch_back_is_verified_too(box):
    assert L.activate_default().ok
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok
    box.env("FAKE_CLI_OK_IDS", CUSTOM_PROJECT)  # the default stopped being valid in the meantime
    o = L.switch_back_to_default()
    assert (o.ok, o.reason) == (False, "not_licensed")
    assert cfg_values(box.fre)["GetFreProjectId"] == CUSTOM_PROJECT  # still on the custom license
    assert L.describe().kind == "custom_active"


def test_recheck_flags_a_license_that_stopped_working(box):
    assert L.activate_default().ok
    assert L.recheck() is True and L.read_state()["verified"] is True
    box.env("FAKE_CLI_OK_IDS", "SomethingElse")
    assert L.recheck() is False
    assert (L.read_state()["status"], L.read_state()["reason"]) == ("failed", "not_licensed")


def test_recheck_while_a_change_is_running_does_not_probe(box):
    with L._lock() as got:
        assert got and L.recheck() is None
    assert box.probes() == 0


def test_banner_warns_when_the_license_service_is_down(box):
    assert L.activate_default().ok
    box.unset("ABBYY_SKIP_LS_RESTART")
    box.env("ABBYY_LS_BIN", str(box.tmp / "not-running"))
    level, title, _ = L.ui_message(L.describe())
    assert level == "warning" and "license service is not running" in title


# ── token files (Option A) ────────────────────────────────────────────────────
def _token_cmd(box, project=CUSTOM_PROJECT):
    """Like the page's Option A: the script copies the token into the Licenses folder first."""
    return ["bash", "-c",
            f'cp "{box.stage}/MY.ABBYY.ActivationToken" "{box.licenses}/"; exec bash "$0" --project-id {project} '
            "--skip-local-service-installation", os.environ["ABBYY_ACTIVATE_SH"]]


@pytest.fixture
def tokens(box):
    box.licenses = box.tmp / "licenses"
    box.stage = box.tmp / "stage"
    for d in (box.licenses, box.stage):
        d.mkdir()
    (box.stage / "MY.ABBYY.ActivationToken").write_text("t" * 80)
    box.env("ABBYY_REAL_LICENSES_DIR", str(box.licenses))
    box.env("ABBYY_LICENSES_DIR", str(box.stage))
    return box


def test_rejected_token_does_not_stay_behind(tokens):
    assert L.activate_default().ok
    tokens.env("FAKE_CLI_OK_IDS", PROJECT)  # the custom project id is not licensed
    r = L.run_custom_activation(_token_cmd(tokens), dict(os.environ))
    assert (r.ok, r.reason) == (False, "not_licensed")
    assert not list(tokens.licenses.glob("*.ActivationToken"))
    assert cfg_values(tokens.fre)["GetFreProjectId"] == PROJECT


def test_script_failure_does_not_leave_a_token_behind_either(tokens):
    tokens.mode("invalid")
    r = L.run_custom_activation(_token_cmd(tokens), dict(os.environ))
    assert (r.ok, r.reason) == (False, "invalid_serial")
    assert not list(tokens.licenses.glob("*.ActivationToken"))


def test_accepted_token_stays_while_the_custom_license_is_in_use(tokens):
    assert L.activate_default().ok
    assert L.run_custom_activation(_token_cmd(tokens), dict(os.environ)).ok
    assert (tokens.licenses / "MY.ABBYY.ActivationToken").exists()


def test_switch_back_removes_the_custom_token_and_releases_the_online_license(tokens, monkeypatch):
    assert L.activate_default().ok
    assert L.run_custom_activation(_token_cmd(tokens), dict(os.environ)).ok
    released = []
    monkeypatch.setattr(L, "_release_online_licenses", lambda *a, **k: released.append(1))
    assert L.switch_back_to_default().ok
    assert not list(tokens.licenses.glob("*.ActivationToken"))
    assert not list(tokens.stage.glob("*.ActivationToken"))
    assert released == [1]


def test_failed_switch_back_keeps_the_custom_token(tokens, monkeypatch):
    assert L.activate_default().ok
    assert L.run_custom_activation(_token_cmd(tokens), dict(os.environ)).ok
    tokens.env("FAKE_CLI_OK_IDS", CUSTOM_PROJECT)  # the default no longer verifies
    monkeypatch.setattr(L, "_release_online_licenses", lambda *a, **k: pytest.fail("must not release"))
    assert not L.switch_back_to_default().ok
    assert (tokens.licenses / "MY.ABBYY.ActivationToken").exists()


def test_switch_back_without_any_token_does_not_touch_the_online_license(box, monkeypatch):
    assert L.activate_default().ok
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok
    monkeypatch.setattr(L, "_release_online_licenses", lambda *a, **k: pytest.fail("no token, nothing to release"))
    assert L.switch_back_to_default().ok


# ── what the page shows ───────────────────────────────────────────────────────
def test_banner_for_a_working_default_license(box):
    assert L.activate_default().ok
    level, title, body = L.ui_message(L.describe())
    assert level == "success" and "already set up" in title and "nothing for you to do" in title
    assert "PDF" in body


def test_banner_while_activating_and_after_start_without_state(box):
    (box.state).mkdir(parents=True, exist_ok=True)
    (box.state / "boot_ts").write_text(str(time.time()))
    v = L.describe()
    assert v.kind == "activating" and L.ui_message(v)[0] == "info"
    (box.state / "boot_ts").write_text(str(time.time() - 3600))
    v = L.describe()
    assert (v.kind, v.reason) == ("default_failed", "not_started")


def test_a_dead_background_job_does_not_show_setting_up_forever(box):
    L.write_state(mode="default", status="activating", reason="")
    raw = json.loads((box.state / "license_state.json").read_text())
    raw["ts"] = time.time() - 2 * 3600
    (box.state / "license_state.json").write_text(json.dumps(raw))
    assert L.describe().kind == "default_failed"


def test_banner_for_failure_names_a_code_and_blames_nobody(box):
    L.write_state(mode="default", status="failed", reason="invalid_serial")
    level, title, body = L.ui_message(L.describe())
    assert level == "error" and "invalid_serial" in body and "not something you did wrong" in body


def test_banner_while_retrying_explains_the_reason_in_plain_sentences(box):
    L.write_state(mode="default", status="retrying", reason="network", attempt=1, next_retry_at=time.time() + 60)
    level, title, body = L.ui_message(L.describe())
    assert level == "info" and "Setting up" in title
    assert L.REASON_TEXT["network"] in body and "trying again automatically" in body
    assert "((" not in body and "(This" not in body and ".)" not in body


def test_conversion_hint_names_the_code_when_the_default_license_failed(box):
    L.write_state(mode="default", status="failed", reason="invalid_serial")
    hint = L.conversion_hint("some output nobody can read")
    assert "invalid_serial" in hint and "License Setup" in hint and "support team" in hint
    assert SERIAL not in hint and PROJECT not in hint


def test_conversion_hint_for_incomplete_default_settings(box):
    box.unset("ABBYY_CUSTOMER_PROJECT_ID")
    assert "incomplete" in L.conversion_hint("")


def test_conversion_hint_when_abbyy_itself_says_there_is_no_license(box):
    assert L.activate_default().ok  # the recorded state is healthy, so only the tool's own words count
    for output in ("Error: ABBYY FineReader Engine is not licensed.", "Error: There are no licenses on 127.0.0.1"):
        assert "no usable license" in L.conversion_hint(output)


def test_conversion_hint_stays_silent_for_unrelated_failures(box):
    assert L.activate_default().ok
    assert L.conversion_hint("Error: Cannot open the input file.") is None and L.conversion_hint("") is None


def test_custom_banner_depends_on_whether_a_default_exists(box):
    L.write_state(mode="custom", status="active", reason="ok")
    assert "different ABBYY license" in L.ui_message(L.describe())[1]
    box.unset("ABBYY_SERIAL_NUMBER")
    box.unset("ABBYY_CUSTOMER_PROJECT_ID")
    assert "lost when the server restarts" in L.ui_message(L.describe())[2]


def test_page_looks_as_before_when_nothing_is_configured(box):
    box.unset("ABBYY_SERIAL_NUMBER")
    box.unset("ABBYY_CUSTOMER_PROJECT_ID")
    assert L.describe().kind == "none" and L.ui_message(L.describe()) is None


# ── licensing service helpers ─────────────────────────────────────────────────
def test_ls_pid_lookup_matches_the_exact_binary_only(box):
    fake = box.tmp / "LicensingService"
    shutil.copy(shutil.which("sleep"), fake)
    box.env("ABBYY_LS_BIN", str(fake))
    p = subprocess.Popen([str(fake), "30"])
    try:
        end = time.monotonic() + 5  # /proc shows the new command line only once the child has exec'd
        while p.pid not in L._ls_pids() and time.monotonic() < end:
            threading.Event().wait(0.02)
        assert p.pid in L._ls_pids()
        # a wrapper whose command line merely CONTAINS the name must not match
        w = subprocess.Popen(["sh", "-c", f"echo {fake} ; sleep 30"])
        try:
            assert w.pid not in L._ls_pids()
        finally:
            w.kill(); w.wait()
    finally:
        p.kill(); p.wait()
    assert p.pid not in L._ls_pids()


def test_restart_reports_a_missing_binary(box):
    box.unset("ABBYY_SKIP_LS_RESTART")
    assert L.restart_ls() is False


def test_state_writes_from_many_threads_never_corrupt_the_file(box):
    def worker(i):
        for _ in range(30):
            L.write_state(mode="default", status="retrying", attempt=i)
    ts = [threading.Thread(target=worker, args=(i,)) for i in range(6)]
    [t.start() for t in ts]
    [t.join() for t in ts]
    assert L.read_state()["status"] == "retrying"


# ── PDF conversions versus license changes ────────────────────────────────────
REAL_SLEEP = time.sleep  # the box fixture silences time.sleep for the whole process
ROOT = Path(__file__).resolve().parent.parent


def _start_conversion():
    """A conversion in ANOTHER process: holds the conversion lock until its stdin is closed."""
    code = ("import sys\nsys.path.insert(0, %r)\nimport abbyy_license as L\n"
            "with L.conversion_slot(wait=5) as ok:\n"
            "    print('held' if ok else 'denied', flush=True)\n"
            "    sys.stdin.read()\n" % str(ROOT))
    p = subprocess.Popen([sys.executable, "-c", code], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                         text=True, env=dict(os.environ))
    assert p.stdout.readline().strip() == "held"
    return p


def _end_conversion(p):
    p.stdin.close()
    p.wait(timeout=10)


def test_a_running_conversion_blocks_an_override_and_changes_nothing(box):
    assert L.activate_default().ok
    attempts, probes, before = box.attempts(), box.probes(), cfg_values(box.fre)
    conversion = _start_conversion()
    try:
        r = L.run_custom_activation(_custom_cmd(box), dict(os.environ))
    finally:
        _end_conversion(conversion)
    assert (r.ok, r.reason) == (False, "converting") and r.output == L.REASON_TEXT["converting"]
    assert box.attempts() == attempts and box.probes() == probes  # neither the script nor ABBYY was touched
    assert cfg_values(box.fre) == before and L.describe().kind == "default_active"
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok  # the same request works afterwards


def test_a_running_conversion_blocks_switching_back(box):
    assert L.activate_default().ok
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok
    probes = box.probes()
    conversion = _start_conversion()
    try:
        o = L.switch_back_to_default()
    finally:
        _end_conversion(conversion)
    assert (o.ok, o.reason) == (False, "converting")
    assert cfg_values(box.fre)["GetFreProjectId"] == CUSTOM_PROJECT and L.describe().kind == "custom_active"
    assert box.probes() == probes
    assert L.switch_back_to_default().ok


def test_the_license_check_is_not_run_while_a_conversion_is_running(box):
    assert L.activate_default().ok
    probes = box.probes()
    conversion = _start_conversion()
    try:
        assert L.recheck() is None  # a busy license must not be mistaken for a broken one
    finally:
        _end_conversion(conversion)
    assert box.probes() == probes and L.describe().kind == "default_active"


def test_setting_up_the_default_license_does_not_wait_for_conversions(box):
    conversion = _start_conversion()
    try:
        assert L.activate_default().ok  # start-up job and "Try again": nothing can convert without it
    finally:
        _end_conversion(conversion)


def test_a_conversion_waits_for_a_license_change_and_then_runs(box):
    changing = threading.Event()

    def license_change():
        with L.conversions_idle() as idle:
            assert idle
            changing.set()
            REAL_SLEEP(0.6)

    t = threading.Thread(target=license_change)
    t.start()
    assert changing.wait(5)
    t0 = time.time()
    with L.conversion_slot(wait=10, poll=0.05, sleep=REAL_SLEEP) as ok:
        waited = time.time() - t0
        assert ok
    t.join()
    assert waited >= 0.4


def test_a_conversion_gives_up_when_the_license_change_takes_too_long(box):
    with L.conversions_idle() as idle:
        assert idle
        t0 = time.time()
        with L.conversion_slot(wait=0.3, poll=0.05, sleep=REAL_SLEEP) as ok:
            assert ok is False
        assert 0.25 <= time.time() - t0 < 5


def test_conversions_do_not_block_each_other(box):
    both_in = threading.Barrier(2, timeout=5)
    results = []

    def convert():
        with L.conversion_slot(wait=5, sleep=REAL_SLEEP) as ok:
            try:
                both_in.wait()  # passes only when both hold the slot at the same time
                results.append(ok)
            except threading.BrokenBarrierError:
                results.append("blocked")

    ts = [threading.Thread(target=convert) for _ in range(2)]
    [t.start() for t in ts]
    [t.join() for t in ts]
    assert results == [True, True]


def test_a_killed_conversion_does_not_keep_the_license_locked(box):
    conversion = _start_conversion()
    with L.conversions_idle() as idle:
        assert idle is False
    conversion.kill()
    conversion.wait(timeout=10)
    with L.conversions_idle() as idle:
        assert idle is True


def test_a_lock_problem_never_stops_a_conversion(box):
    blocker = box.tmp / "afile"
    blocker.write_text("x")
    box.env("ABBYY_STATE_DIR", str(blocker / "state"))  # cannot be created: its parent is a file
    with L.conversion_slot() as ok:
        assert ok is True
    with L.conversions_idle() as idle:
        assert idle is True


def test_a_filesystem_without_locking_never_stops_a_conversion(box, monkeypatch):
    def no_locks(fd, op):
        raise OSError(37, "No locks available")  # ENOLCK

    monkeypatch.setattr(L.fcntl, "flock", no_locks)
    with L.conversion_slot() as ok:
        assert ok is True
    with L.conversions_idle() as idle:
        assert idle is True


def test_license_changes_are_logged_without_any_identifier(box, capsys):
    assert L.activate_default().ok
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok
    assert L.switch_back_to_default().ok
    box.mode("invalid")
    assert not L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok
    out = capsys.readouterr().out
    assert "a custom license was activated" in out and "switched back to the default license" in out
    assert "custom license change failed (invalid_serial)" in out
    for secret in (SERIAL, PROJECT, CUSTOM_SERIAL, CUSTOM_PROJECT):
        assert secret not in out


def test_a_refused_change_is_logged_with_its_reason(box, capsys):
    assert L.activate_default().ok
    conversion = _start_conversion()
    try:
        L.run_custom_activation(_custom_cmd(box), dict(os.environ))
    finally:
        _end_conversion(conversion)
    assert "custom license change failed (converting)" in capsys.readouterr().out


# ── connection check (which step fails when ABBYY cannot be reached?) ─────────
class TlsServer:
    """A real local HTTPS server (self-signed certificate for 'localhost') for the connection checks."""

    def __init__(self, tmp_path, reply=b"HTTP/1.1 200 OK\r\nContent-Length: 2\r\nConnection: close\r\n\r\nok"):
        key, self.crt = tmp_path / "key.pem", tmp_path / "cert.pem"
        subprocess.run(["openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-keyout", str(key),
                        "-out", str(self.crt), "-days", "2", "-subj", "/CN=localhost",
                        "-addext", "subjectAltName=DNS:localhost"], check=True, capture_output=True)
        self.ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        self.ctx.load_cert_chain(str(self.crt), str(key))
        self.reply, self.requests = reply, []
        self.sock = socket.socket()
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(8)
        self.port = self.sock.getsockname()[1]
        threading.Thread(target=self._serve, daemon=True).start()

    def _serve(self):
        while True:
            try:
                conn, _ = self.sock.accept()
            except OSError:
                return
            try:
                with self.ctx.wrap_socket(conn, server_side=True) as s:
                    self.requests.append(s.recv(4096))
                    if self.reply:
                        s.sendall(self.reply)
            except (ssl.SSLError, OSError):
                pass

    def close(self):
        self.sock.close()


@pytest.fixture
def tls_server(tmp_path, monkeypatch):
    servers = []

    def make(**kw):
        server = TlsServer(tmp_path, **kw)
        servers.append(server)
        return server

    monkeypatch.delenv("SSL_CERT_DIR", raising=False)
    monkeypatch.delenv("SSL_CERT_FILE", raising=False)
    yield make
    for server in servers:
        server.close()


def _trust(monkeypatch, server):
    monkeypatch.setenv("SSL_CERT_FILE", str(server.crt))


def test_connection_check_passes_every_step_against_a_reachable_server(box, tls_server, monkeypatch):
    server = tls_server()
    _trust(monkeypatch, server)
    res = L._check_host("localhost", "test server", 5.0, port=server.port)
    assert res.ok and res.dns.startswith("ok") and res.tcp == "ok"
    assert res.tls.startswith("ok TLSv1") and "signed by localhost" in res.tls
    assert res.http == "ok HTTP 200"
    request = server.requests[0].decode()
    assert request.startswith("GET / HTTP/1.1") and "Host: localhost" in request
    assert SERIAL not in request and PROJECT not in request  # only an ordinary page request leaves the server


def test_connection_check_reports_a_failed_name_lookup_and_stops_there(box):
    res = L._check_host("no-such-host.invalid", "x", 3.0)
    assert res.dns.startswith("FAILED") and (res.tcp, res.tls, res.http) == ("-", "-", "-") and not res.ok


def test_connection_check_reports_a_blocked_connection(box):
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        free_port = s.getsockname()[1]
    res = L._check_host("localhost", "x", 3.0, port=free_port)
    assert res.dns.startswith("ok") and res.tcp.startswith("FAILED") and (res.tls, res.http) == ("-", "-")


def test_connection_check_reports_an_untrusted_certificate(box, tls_server):
    server = tls_server()  # nobody trusts its self-signed certificate (like a corporate TLS proxy without its CA)
    res = L._check_host("localhost", "x", 5.0, port=server.port)
    assert res.tcp == "ok" and res.tls.startswith("FAILED: certificate not trusted") and res.http == "-"
    assert not res.ok


def test_connection_check_shows_a_firewall_block_page_without_its_markup(box, tls_server, monkeypatch):
    server = tls_server(reply=b"HTTP/1.1 403 Forbidden\r\nConnection: close\r\n\r\n"
                              b"<html><body><h1>Access denied</h1> by corporate policy</body></html>")
    _trust(monkeypatch, server)
    res = L._check_host("localhost", "x", 5.0, port=server.port)
    assert res.tls.startswith("ok") and res.http.startswith("FAILED: HTTP 403")
    assert "Access denied by corporate policy" in res.http and "<" not in res.http and not res.ok


def test_connection_check_accepts_a_404_home_page(box, tls_server, monkeypatch):
    server = tls_server(reply=b"HTTP/1.1 404 Not Found\r\nConnection: close\r\n\r\nnope")
    _trust(monkeypatch, server)
    assert L._check_host("localhost", "x", 5.0, port=server.port).http == "ok HTTP 404"  # the site answered


def test_connection_check_notices_a_server_that_hangs_up_without_answering(box, tls_server, monkeypatch):
    server = tls_server(reply=b"")
    _trust(monkeypatch, server)
    res = L._check_host("localhost", "x", 5.0, port=server.port)
    assert res.http.startswith("FAILED: no web reply")


def test_diagnose_checks_all_hosts_at_the_same_time_and_keeps_their_order(box, monkeypatch):
    barrier = threading.Barrier(3)  # only passes when three checks run concurrently

    def fake(host, purpose, timeout, port):
        barrier.wait(5)
        return L.HostCheck(host, purpose, "ok 1.1.1.1", "ok", "ok TLSv1.3", "ok HTTP 200")

    monkeypatch.setattr(L, "_check_host", fake)
    out = REAL_DIAGNOSE(hosts=(("a.example", "A"), ("b.example", "B"), ("c.example", "C")), timeout=5.0)
    assert [c.host for c in out] == ["a.example", "b.example", "c.example"] and all(c.ok for c in out)


def test_diagnose_survives_a_check_that_crashes(box, monkeypatch):
    def fake(host, purpose, timeout, port):
        if host == "bad.example":
            raise RuntimeError("boom")
        return L.HostCheck(host, purpose, "ok", "ok", "ok", "ok HTTP 200")

    monkeypatch.setattr(L, "_check_host", fake)
    out = REAL_DIAGNOSE(hosts=(("good.example", "G"), ("bad.example", "B")), timeout=2.0)
    assert out[0].ok and out[1].dns == "FAILED: RuntimeError"


def test_diagnose_does_not_hang_when_a_resolver_never_answers(box, monkeypatch):
    release = threading.Event()
    monkeypatch.setattr(L, "_check_host", lambda *a, **k: release.wait(30))
    started = time.monotonic()
    out = REAL_DIAGNOSE(hosts=(("stuck.example", "S"),), timeout=0.2)
    release.set()
    assert time.monotonic() - started < 5 and out[0].dns == "FAILED: no answer in time"


def test_connection_summary_lines_are_one_per_host_with_every_step(box):
    c = L.HostCheck("registration2.abbyy.com", "x", "ok 1.2.3.4", "ok", "FAILED: certificate not trusted (x)", "-", 0.34)
    assert L.connection_log_lines([c]) == [
        "registration2.abbyy.com: dns=ok 1.2.3.4; tcp=ok; tls=FAILED: certificate not trusted (x); web=- (0.3 s)"]


def _hc(host, dns="ok 1.2.3.4", tcp="ok", tls="ok TLSv1.3", http="ok HTTP 200"):
    return L.HostCheck(host, "x", dns, tcp, tls, http, 0.2)


def test_a_host_check_names_its_verdict_and_the_first_step_that_failed(box):
    assert _hc("a.example").verdict == "ok" and _hc("a.example").failing_step() == ""
    blocked = _hc("registration2.abbyy.com", tcp="FAILED: timed out", tls="-", http="-")
    assert blocked.verdict == "problem" and blocked.failing_step() == "connection"
    assert _hc("a.example", tls="FAILED: certificate not trusted (x)", http="-").failing_step() == "secure connection (certificate)"
    assert _hc("a.example", dns="FAILED: x", tcp="-", tls="-", http="-").failing_step() == "name lookup"
    assert _hc("a.example", http="FAILED: HTTP 403 (denied)").failing_step() == "web reply"
    # the old .ru address no longer exists: failing there is normal and must not look like a problem
    assert _hc("activation1.abbyy.ru", dns="FAILED: Name or service not known", tcp="-", tls="-", http="-").verdict == "expected"
    assert _hc("activation1.abbyy.ru").verdict == "ok"  # if it does answer, that is simply fine


def test_summary_when_everything_that_matters_is_reachable(box):
    checks = [_hc("registration2.abbyy.com"), _hc("activation1.abbyy.ru", dns="FAILED: x", tcp="-", tls="-", http="-"),
              _hc("example.com")]
    level, text = L.summarize_connection(checks)
    assert level == "success" and "can reach every ABBYY server" in text


def test_summary_names_the_blocked_server_and_the_step_and_what_to_ask_for(box):
    checks = [_hc("registration2.abbyy.com", tcp="FAILED: timed out", tls="-", http="-"), _hc("example.com")]
    level, text = L.summarize_connection(checks)
    assert level == "error" and "`registration2.abbyy.com` fails at the connection" in text
    assert "network team" in text and "443" in text and "no general internet" not in text


def test_summary_tells_no_internet_from_missing_certificates_from_a_proxy_page(box):
    host = _hc("registration2.abbyy.com", dns="FAILED: x", tcp="-", tls="-", http="-")
    none = L.summarize_connection([host, _hc("example.com", dns="FAILED: x", tcp="-", tls="-", http="-")])[1]
    assert "no general internet access" in none and "name lookup" in none
    certs = L.summarize_connection([host, _hc("example.com", tls="FAILED: certificate not trusted (x)", http="-")])[1]
    assert "does not trust the certificates" in certs and "TLS-inspection" in certs
    proxy = L.summarize_connection([host, _hc("example.com", http="FAILED: HTTP 407 (login)")])[1]
    assert "proxy or firewall is answering" in proxy


def test_summary_of_nothing_says_so(box):
    assert L.summarize_connection([])[0] == "info"


def test_default_host_list_names_the_servers_abbyys_tool_and_documentation_use(box):
    hosts = [h for h, _ in L.CHECK_HOSTS]
    for must in ("registration2.abbyy.com", "activation.abbyy.com", "license-cloud.abbyy.com", "activation1.abbyy.ru"):
        assert must in hosts
    assert hosts[-1] == "example.com"  # the control that tells "no internet at all" from "ABBYY blocked"


def test_proxy_summary_names_the_settings_but_never_shows_their_values(box):
    fake_login = "user:secretpw"  # made up; built at run time so scanners do not take it for a credential
    box.env("HTTPS_PROXY", f"http://{fake_login}@proxy.example:8080")
    box.env("no_proxy", "internal.example")
    (box.fre / "Bin" / "LicensingSettings.xml").write_text('<ProxyServer Server="http://p:1" Login="u" Password="pw2"/>')
    summary = L.proxy_summary()
    assert summary["env"] == ["HTTPS_PROXY", "NO_PROXY"] and summary["licensing_settings"] is True
    assert "secretpw" not in repr(summary) and "pw2" not in repr(summary) and "proxy.example" not in repr(summary)


def test_proxy_summary_when_nothing_is_set(box, monkeypatch):
    for name in L._PROXY_VARS:
        monkeypatch.delenv(name, raising=False)
        monkeypatch.delenv(name.lower(), raising=False)
    assert L.proxy_summary() == {"env": [], "licensing_settings": False}


def _license_dir(box):
    d = box.tmp / "lic"
    d.mkdir(exist_ok=True)
    box.env("ABBYY_REAL_LICENSES_DIR", str(d))
    return d


def test_abbyy_log_tail_uses_the_newest_file_drops_routine_lines_and_masks_identifiers(box):
    d = _license_dir(box)
    old = d / "ProductProtection_2026-01-01 10h00m.log"
    old.write_text("01/01/26 10:00:00 (1) OLD ENTRY that must not appear\n")
    os.utime(old, (1, 1))
    new = d / "ProductProtection_2026-11-16 10h00m.log"
    new.write_text(
        "16/11/26 10:00:01 (7) Application context [0x1] created. Identity is user root\n"
        "16/11/26 10:00:02 (7) Initializing network subsystem\n"
        f"16/11/26 10:00:03 (7) Activation request to registration2.abbyy.com failed for {SERIAL} token ABCDEFGHIJKLMNOPQRSTUVWX\n"
        "16/11/26 10:00:04 (7) Deinitializing product protection\n")
    tail = L.abbyy_log_tail()
    assert len(tail["ProductProtection"]) == 1 and tail["NetworkServer"] == []
    line = tail["ProductProtection"][0]
    assert "registration2.abbyy.com" in line and "failed" in line
    assert SERIAL not in line and "ABCDEFGHIJKLMNOPQRSTUVWX" not in line and "OLD ENTRY" not in line


def test_abbyy_log_tail_reads_utf16_logs_and_survives_missing_ones(box):
    d = _license_dir(box)
    (d / "NetworkServer_2026-11-16 10h00m.log").write_bytes(
        b"\xff\xfe" + "16/11/26 10:00:00 (3) Online licensing server error: timed out\n".encode("utf-16-le"))
    assert L.abbyy_log_tail()["NetworkServer"] == ["16/11/26 10:00:00 (3) Online licensing server error: timed out"]
    box.env("ABBYY_REAL_LICENSES_DIR", str(box.tmp / "does-not-exist"))
    assert L.abbyy_log_tail() == {"ProductProtection": [], "NetworkServer": []}


def test_tool_output_tail_keeps_the_last_lines_and_hides_identifiers(box):
    out = ("ABBYY FineReader Engine 12 activation script\nCopyright (c) 2019 ABBYY\n"
           f"Activating license {SERIAL}\nCould not activate license via internet\n"
           "Activation has failed.\nsession ABCDEFGHIJKLMNOPQRSTUVWX\n")
    tail = L.tool_output_tail(out)
    assert tail.splitlines()[-2:] == ["Activation has failed.", "session <id>"]
    assert "Could not activate license via internet" in tail
    assert SERIAL not in tail and "Copyright" not in tail and "activation script" not in tail
    assert len(L.tool_output_tail("\n".join(f"line {i}" for i in range(50))).splitlines()) == 8
    assert L.tool_output_tail("") == "" and L.tool_output_tail(None) == ""


def test_the_tool_output_of_a_failed_activation_is_kept_redacted_with_its_duration(box):
    box.mode("network")
    o = L._activate_serial(SERIAL, PROJECT)
    assert not o.ok and o.reason == "network" and o.transient and o.seconds >= 0
    assert "Could not connect to the activation server" in o.detail
    box.mode("leak")  # this stub prints the serial and project id
    o = L._activate_serial(SERIAL, PROJECT)
    assert SERIAL not in o.detail and PROJECT not in o.detail
    box.mode("ok")
    assert L._activate_serial(SERIAL, PROJECT).detail == ""  # successes carry no failure detail


def test_the_start_up_job_logs_what_the_tool_said_and_checks_the_connection_twice(box, monkeypatch, capsys):
    canned = [L.HostCheck("registration2.abbyy.com", "x", "ok 9.9.9.9", "FAILED: timed out", "-", "-", 6.0)]
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: canned)
    box.mode("network")
    assert L.autoactivate(sleep=lambda s: None) == 1
    out = capsys.readouterr().out
    assert "attempt 1 failed: network" in out and "attempt 3 failed: network" in out
    assert "ABBYY tool: Could not connect to the activation server" in out
    assert out.count("connection check: registration2.abbyy.com: dns=ok 9.9.9.9; tcp=FAILED: timed out") == 2
    assert out.count("connection check result: Not every ABBYY server can be reached") == 2
    assert "attempt 2 failed" in out  # the middle attempt is logged but not checked again
    for secret in (SERIAL, PROJECT):
        assert secret not in out


def test_a_rejected_serial_is_not_followed_by_a_connection_check(box, monkeypatch, capsys):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: pytest.fail("the network is not the problem here"))
    box.mode("invalid")
    assert L.autoactivate(sleep=lambda s: None) == 1
    out = capsys.readouterr().out
    assert "attempt 1 failed: invalid_serial" in out and "connection check" not in out


def test_a_crashing_connection_check_never_stops_the_start_up_job(box, monkeypatch, capsys):
    def boom(*a, **k):
        raise RuntimeError("diagnostics broke")

    monkeypatch.setattr(L, "diagnose_connection", boom)
    box.mode("network")
    assert L.autoactivate(sleep=lambda s: None) == 1
    assert "connection check could not run: RuntimeError" in capsys.readouterr().out
    assert L.read_state()["reason"] == "network"


def test_banner_after_the_retries_says_blocked_not_temporary_and_names_the_connection_check(box):
    L.write_state(mode="default", status="failed", reason="network")
    level, title, body = L.ui_message(L.describe())
    assert level == "error" and "even after several tries" in body and "firewall" in body
    assert "usually temporary" not in body and "**Connection check**" in body and "**network**" in body
    L.write_state(mode="default", status="failed", reason="invalid_serial")
    assert "Connection check" not in L.ui_message(L.describe())[2]  # wrong serial is not a network problem
    L.write_state(mode="default", status="retrying", reason="network", attempt=1, next_retry_at=time.time() + 60)
    assert "usually temporary" in L.ui_message(L.describe())[2]  # while retrying it still is
    assert L.reason_text("network") == L.REASON_TEXT["network"] and "several tries" in L.reason_text("network", final=True)
    assert L.reason_text("limit", final=True) == L.REASON_TEXT["limit"]


# ── "ABBYY said no" is not "ABBYY could not be reached" ───────────────────────
# Measured on 2026-11-16 with a local forwarder: ABBYY answered the activation request within 2 s with
# "Failed, limit exceed" and the tool printed only its generic line "Could not activate license via internet".
def _reached():
    """A canned connection check: nothing blocked, and ABBYY's activation service answered the test message."""
    return [L.HostCheck("registration2.abbyy.com", "x", "ok 1.2.3.4", "ok", "ok TLSv1.3", "ok HTTP 404", 0.2),
            L.HostCheck("example.com", "control", "ok 5.6.7.8", "ok", "ok TLSv1.3", "ok HTTP 200", 0.2),
            L.HostCheck("registration2.abbyy.com", "p", "ok 1.2.3.4", "ok", "ok TLSv1.3", "ok HTTP 500", 0.3,
                        kind="activation")]


def _blocked():
    return [L.HostCheck("registration2.abbyy.com", "x", "ok 1.2.3.4", "FAILED: timed out", "-", "-", 6.0),
            L.HostCheck("example.com", "control", "ok 5.6.7.8", "ok", "ok TLSv1.3", "ok HTTP 200", 0.2),
            L.HostCheck("registration2.abbyy.com", "p", "ok 1.2.3.4", "FAILED: timed out", "-", "-", 6.0,
                        kind="activation")]


def test_the_activation_test_message_is_empty_and_a_soap_fault_counts_as_an_answer(box, tls_server, monkeypatch):
    server = tls_server(reply=b"HTTP/1.1 500 Internal Server Error\r\nContent-Length: 0\r\nConnection: close\r\n\r\n")
    _trust(monkeypatch, server)
    res = L._check_host("localhost", "probe", 5.0, port=server.port, probe="/ProductWebCall/BusinessVega.asmx")
    assert res.ok and res.kind == "activation" and res.http == "ok HTTP 500"
    request = server.requests[0].decode()
    assert request.startswith("POST /ProductWebCall/BusinessVega.asmx HTTP/1.1") and "<soap12:Body/>" in request
    assert "Content-Type: application/soap+xml" in request
    assert SERIAL not in request and PROJECT not in request and "SerialNumbers" not in request  # no license data


@pytest.mark.parametrize("status,answered", [(200, True), (404, True), (500, True), (403, False), (407, False),
                                             (451, False), (502, False), (503, False)])
def test_the_activation_test_message_is_judged_like_a_firewall_check_but_accepts_a_fault(
        box, tls_server, monkeypatch, status, answered):
    server = tls_server(reply=f"HTTP/1.1 {status} X\r\nContent-Length: 0\r\nConnection: close\r\n\r\n".encode())
    _trust(monkeypatch, server)
    res = L._check_host("localhost", "probe", 5.0, port=server.port, probe="/x.asmx")
    assert res.ok is answered, res.http


def test_a_home_page_check_still_treats_a_500_as_a_failure(box, tls_server, monkeypatch):
    server = tls_server(reply=b"HTTP/1.1 500 Oops\r\nContent-Length: 0\r\nConnection: close\r\n\r\n")
    _trust(monkeypatch, server)
    assert not L._check_host("localhost", "x", 5.0, port=server.port).ok  # only the test message accepts a fault


def test_the_default_check_ends_with_the_activation_test_message_but_custom_lists_do_not(box, monkeypatch):
    calls = []

    def fake(host, purpose, timeout, port, probe=None):
        calls.append((host, probe))
        return L.HostCheck(host, purpose, "ok", "ok", "ok", "ok HTTP 200", kind="activation" if probe else "page")

    monkeypatch.setattr(L, "_check_host", fake)
    out = REAL_DIAGNOSE()
    assert len(out) == len(L.CHECK_HOSTS) + 1 and out[-1].kind == "activation" and out[-1].host == L.ACTIVATION_HOST
    assert [c.kind for c in out].count("activation") == 1 and (L.ACTIVATION_HOST, L.ACTIVATION_PATH) in calls
    calls.clear()
    assert [c.kind for c in REAL_DIAGNOSE(hosts=(("a.example", "A"),))] == ["page"] and calls == [("a.example", None)]


def test_a_crashing_activation_test_still_gives_an_activation_row_that_is_not_reached(box, monkeypatch):
    def fake(host, purpose, timeout, port, probe=None):
        if probe:
            raise RuntimeError("boom")
        return L.HostCheck(host, purpose, "ok", "ok", "ok", "ok HTTP 200")

    monkeypatch.setattr(L, "_check_host", fake)
    out = REAL_DIAGNOSE(timeout=2.0)
    assert out[-1].kind == "activation" and out[-1].dns == "FAILED: RuntimeError"
    assert L.activation_service_reached(out) is False


def test_the_server_counts_as_reached_when_the_activation_service_answered_whatever_else_fails(box):
    assert L.activation_service_reached(_reached()) is True
    assert L.activation_service_reached(_blocked()) is False
    assert L.activation_service_reached([]) is False
    assert L.activation_service_reached([c for c in _reached() if c.kind != "activation"]) is False  # home page only
    # another address fails, or a hiccup hit the home page of the same host: the service itself still answered
    other = _reached() + [L.HostCheck("license-cloud.abbyy.com", "x", "ok 9.9.9.9", "FAILED: timed out", "-", "-")]
    assert L.activation_service_reached(other) is True
    hiccup = [L.HostCheck("registration2.abbyy.com", "x", "ok 1.2.3.4", "FAILED: [Errno 101] Network is unreachable", "-", "-"),
              _reached()[2]]
    assert L.activation_service_reached(hiccup) is True


def test_the_log_marks_the_line_of_the_activation_test_message(box):
    c = L.HostCheck("registration2.abbyy.com", "p", "ok 1.2.3.4", "ok", "ok TLSv1.3", "ok HTTP 500", 0.5, kind="activation")
    assert L.connection_log_lines([c])[0].startswith("registration2.abbyy.com [activation request]: dns=ok 1.2.3.4")


def test_the_all_clear_summary_says_what_a_failed_activation_then_means(box):
    level, text = L.summarize_connection([_hc("registration2.abbyy.com"), _hc("example.com")])
    assert level == "success" and "refused it" in text and "activations" in text


def test_the_summary_names_the_activation_request_separately_from_the_home_page(box):
    level, text = L.summarize_connection(_blocked())
    assert level == "error"
    assert "`registration2.abbyy.com` fails at the connection" in text
    assert "`registration2.abbyy.com` (activation request) fails at the connection" in text


def test_only_a_network_failure_is_refined_and_only_when_abbyy_answers(box, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached())
    assert L.refine_reason("network") == "refused"
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _blocked())
    assert L.refine_reason("network") == "network"
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: pytest.fail("no check for any other reason"))
    for reason in ("ok", "invalid_serial", "timeout", "unknown", "limit", "service", "busy", "converting"):
        assert L.refine_reason(reason) == reason


def test_a_check_that_crashes_leaves_the_reason_alone(box, monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("diagnostics broke")

    monkeypatch.setattr(L, "diagnose_connection", boom)
    assert L.refine_reason("network") == "network"


def test_try_again_reports_refused_when_abbyy_answers_and_network_when_it_cannot_be_reached(box, monkeypatch):
    box.mode("network")
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached())
    o = L.activate_default()
    assert (o.ok, o.reason) == (False, "refused")
    assert (L.read_state()["status"], L.read_state()["reason"]) == ("failed", "refused")
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _blocked())
    o = L.activate_default()
    assert (o.reason, L.read_state()["reason"]) == ("network", "network")


def test_a_start_up_job_attempt_leaves_the_classification_and_the_final_state_to_the_job(box, monkeypatch):
    box.mode("network")
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: pytest.fail("the job runs its own check"))
    o = L.activate_default(respect_custom=True, retrying=True)
    assert (o.ok, o.reason) == (False, "network")
    assert L.read_state()["status"] == "activating"  # not "failed": the job decides what comes next


def test_the_last_check_decides_when_the_network_went_away_in_between(box, monkeypatch):
    def answers():
        yield _reached()  # attempt 1: ABBYY answered
        while True:
            yield _blocked()  # then the network went away: the check repeats while a refusal is suspected

    gen = answers()
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: next(gen))
    box.mode("network")
    assert L.autoactivate(sleep=lambda s: None) == 1
    assert box.attempts() == 3 and L.read_state()["reason"] == "network"


def test_a_failed_check_at_the_first_attempt_but_a_good_one_at_the_last_gives_refused(box, monkeypatch):
    answers = iter([_blocked(), _blocked(), _reached()])  # attempt 1: blocked twice; attempt 3: reached at once
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: next(answers))
    box.mode("network")
    seen = []
    assert L.autoactivate(sleep=lambda s: seen.append(L.read_state()["reason"])) == 1
    assert seen == ["network", "network"] and L.read_state()["reason"] == "refused"


def test_a_single_failed_look_is_repeated_once_before_the_activation_service_is_called_unreachable(box, monkeypatch, capsys):
    calls = []

    def looks(*a, **k):
        calls.append(1)
        return _blocked() if len(calls) == 1 else _reached()  # a hiccup, then the service answers

    monkeypatch.setattr(L, "diagnose_connection", looks)
    assert L.refine_reason("network") == "refused" and len(calls) == 2
    assert "did not answer the first time, so it was checked again" in capsys.readouterr().out
    calls.clear()
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: calls.append(1) or _blocked())
    assert L.refine_reason("network") == "network" and len(calls) == 2  # blocked both times: now it is a network problem
    calls.clear()
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: calls.append(1) or _reached())
    assert L.refine_reason("network") == "refused" and len(calls) == 1  # no second look when the first one is fine


def test_the_second_look_waits_a_moment_so_a_hiccup_can_pass(box, monkeypatch):
    pauses = []
    monkeypatch.setattr(L.time, "sleep", pauses.append)
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _blocked())
    L.refine_reason("network")
    assert pauses == [2]


def test_a_token_activation_is_not_refined_because_it_uses_other_servers(box, monkeypatch):
    assert L.activate_default().ok
    box.mode("network")
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: pytest.fail("nothing to check for a token"))
    token_cmd = ["bash", os.environ["ABBYY_ACTIVATE_SH"], "--project-id", CUSTOM_PROJECT, "--license-path", "/x/t.ActivationToken",
                 "--license-password", "pw", "--skip-local-service-installation"]
    r = L.run_custom_activation(token_cmd, dict(os.environ))
    assert (r.ok, r.reason) == (False, "network")


def test_a_rejected_serial_stays_a_rejected_serial_even_when_abbyy_is_reachable(box, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached())
    box.mode("invalid")
    assert L.autoactivate(sleep=lambda s: None) == 1
    assert L.read_state()["reason"] == "invalid_serial"


def test_a_custom_license_that_abbyy_declines_is_reported_as_refused_not_as_a_network_problem(box, monkeypatch):
    assert L.activate_default().ok
    box.mode("network")
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached())
    r = L.run_custom_activation(_custom_cmd(box), dict(os.environ))
    assert (r.ok, r.reason) == (False, "refused")
    assert cfg_values(box.fre)["GetFreProjectId"] == PROJECT and L.describe().kind == "default_active"
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _blocked())
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).reason == "network"


def test_switch_back_reports_refused_when_abbyy_declines_the_new_activation(box, monkeypatch):
    assert L.activate_default().ok
    assert L.run_custom_activation(_custom_cmd(box), dict(os.environ)).ok
    shutil.rmtree(box.state / "default_config")  # nothing saved to restore: a real activation is needed
    box.mode("network")
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached())
    o = L.switch_back_to_default()
    assert (o.ok, o.reason) == (False, "refused")
    assert L.describe().kind == "custom_active"  # the custom license keeps working


def test_banner_for_a_refused_activation_points_to_abbys_answer_not_to_the_network(box):
    L.write_state(mode="default", status="failed", reason="refused")
    level, title, body = L.ui_message(L.describe())
    assert level == "error" and "**refused**" in body and "not something you did wrong" in body
    assert "not a firewall problem" in body and "activations" in body and "Online (token) license" in body
    assert "Connection check" not in body and "usually temporary" not in body and "is probably blocking" not in body
    L.write_state(mode="default", status="retrying", reason="refused", attempt=1, next_retry_at=time.time() + 60)
    retry = L.ui_message(L.describe())[2]
    assert "was reached, but it did not grant the activation" in retry and "trying again automatically" in retry
    assert L.reason_text("refused", final=True) == L.FINAL_REASON_TEXT["refused"] != L.REASON_TEXT["refused"]
    L.write_state(mode="default", status="failed", reason="refused")
    assert "code: refused" in L.conversion_hint()  # the line under a failed PDF conversion names it too
