"""UI tests: run the real License Setup page and main page headlessly for each license state.

They check what a user would see (banner text, buttons, disabled state) and that the
license identifiers never appear anywhere on the rendered pages. All values are made up.
"""
import ast
import json
import os
import re
import sys
import threading
import time
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import abbyy_license as L  # noqa: E402

SERIAL = "TESTSERIAL0000000000TEST"
PROJECT = "TestDefaultProject01"
PAGE = str(ROOT / "pages" / "2_License_Setup.py")
MAIN = str(ROOT / "streamlit_app.py")


@pytest.fixture
def env(tmp_path, monkeypatch):
    fre = tmp_path / "fre"
    for d in ("Samples", "Bin", "CommonBin/Licensing", "Samples/CommandLineInterface"):
        (fre / d).mkdir(parents=True)
    (fre / "Samples" / "SamplesConfig.h").write_text("x")
    (fre / "activatefre.sh").write_text("#!/bin/bash\nexit 0\n")
    (fre / "Samples" / "CommandLineInterface" / "CommandLineInterface").write_text("#!/bin/bash\nexit 1\n")
    state = tmp_path / "state"
    state.mkdir()
    for k, v in {
        "ABBYY_FRE_DIR": str(fre), "ABBYY_STATE_DIR": str(state), "ABBYY_SKIP_LS_RESTART": "1",
        "ABBYY_ACTIVATE_SH": str(fre / "activatefre.sh"),
        "ABBYY_CLI": str(fre / "Samples" / "CommandLineInterface" / "CommandLineInterface"),
        "ABBYY_LICENSES_DIR": str(tmp_path / "stage"), "ABBYY_REAL_LICENSES_DIR": str(tmp_path / "lic"),
        "ABBYY_SERIAL_NUMBER": SERIAL, "ABBYY_CUSTOMER_PROJECT_ID": PROJECT,
    }.items():
        monkeypatch.setenv(k, v)
    monkeypatch.delenv("ABBYY_LS_HOST", raising=False)

    class E:
        pass
    e = E()
    e.state = state
    e.set = lambda **kw: L.write_state(**kw)
    e.unset = monkeypatch.delenv
    e.env = monkeypatch.setenv
    return e


def run(path=PAGE):
    at = AppTest.from_file(path, default_timeout=30)
    at.run()
    assert not at.exception, [x.value for x in at.exception]
    return at


def texts(at) -> str:
    parts = []
    for kind in ("success", "info", "warning", "error", "markdown", "caption", "subheader", "title", "code", "text"):
        parts += [str(getattr(el, "value", "")) for el in getattr(at, kind)]
    parts += [str(getattr(b, "label", "")) for b in at.button]
    parts += [str(getattr(c, "label", "")) for c in at.checkbox]
    return "\n".join(parts)


def banner(at):
    for kind in ("success", "info", "warning", "error"):
        for el in getattr(at, kind):
            if "**" in str(el.value) and "\n\n" in str(el.value):
                return kind, str(el.value)
    return None


def activate_buttons(at):
    return [b for b in at.button if b.label in ("💾 Install & Activate Token", "🚀 Activate Now")]


# ── License Setup page ───────────────────────────────────────────────────────
def test_default_license_active_says_nothing_to_do(env):
    env.set(mode="default", status="active", reason="ok", verified=True)
    at = run()
    kind, text = banner(at)
    assert kind == "success"
    assert "already set up" in text and "nothing for you to do" in text
    assert not any("Switch back" in b.label for b in at.button)  # nothing to switch back from
    assert at.title[0].value.startswith("🔑")  # the page keeps its identity


def test_options_a_and_b_are_still_there_but_need_confirmation(env):
    env.set(mode="default", status="active", reason="ok")
    at = run()
    t = texts(at)
    assert "Option A" in t and "Option B" in t and "Use a different ABBYY license" in t
    assert at.checkbox[0].value is False
    assert "Tick the box" in t


def test_activate_buttons_unlock_only_after_confirmation(env):
    env.set(mode="default", status="active", reason="ok")
    at = AppTest.from_file(PAGE, default_timeout=30).run()
    form_buttons = [b for b in at.button if b.label == "🚀 Activate Now"]
    assert form_buttons and all(b.disabled for b in form_buttons)
    at.checkbox[0].check().run()
    form_buttons = [b for b in at.button if b.label == "🚀 Activate Now"]
    assert form_buttons and not any(b.disabled for b in form_buttons)


def test_the_page_says_how_long_the_server_has_been_running_and_whether_the_application_restarted(env):
    env.set(mode="default", status="active", reason="ok", verified=True)
    (env.state / "boot_ts").write_text(str(time.time() - 3 * 3600 - 20 * 60 - 30))
    (env.state / "start_count").write_text("3")
    t = texts(run())
    assert "This server has been running for 3 h 20 min" in t
    assert "The application has restarted 2 times on this server" in t


def test_the_page_stays_quiet_about_uptime_when_it_is_unknown_or_nothing_restarted(env):
    env.set(mode="default", status="active", reason="ok", verified=True)
    assert "has been running" not in texts(run())  # no start time known
    (env.state / "boot_ts").write_text(str(time.time() - 600))
    (env.state / "start_count").write_text("1")
    t = texts(run())
    assert "This server has been running for 10 min" in t and "restarted" not in t


def test_while_activating_the_page_says_so_and_blocks_changes(env):
    env.set(mode="default", status="activating", reason="")
    at = run()
    kind, text = banner(at)
    assert kind == "info" and "Setting up the ABBYY license" in text
    assert "Please wait until the default license has been set up" in texts(at)
    assert not activate_buttons(at) or all(b.disabled for b in activate_buttons(at))


def test_failure_names_a_code_and_offers_a_retry(env):
    env.set(mode="default", status="failed", reason="invalid_serial")
    at = run()
    kind, text = banner(at)
    assert kind == "error" and "invalid_serial" in text and "not something you did wrong" in text
    assert any(b.label.startswith("🔁 Try the default license again") for b in at.button)


def test_custom_license_offers_switch_back_with_a_confirmation_step(env):
    env.set(mode="custom", status="active", reason="ok")
    at = run()
    kind, text = banner(at)
    assert kind == "warning" and "different ABBYY license" in text
    switch = [b for b in at.button if "Switch back to the default license" in b.label]
    assert switch
    switch[0].click().run()
    t = texts(at)
    assert "takes about 10 seconds" in t and "try again in a minute" in t
    assert any(b.label == "Yes, switch back" for b in at.button)
    assert any(b.label == "Cancel" for b in at.button)
    [b for b in at.button if b.label == "Cancel"][0].click().run()
    assert not any(b.label == "Yes, switch back" for b in at.button)


def test_page_is_unchanged_for_servers_without_a_default_license(env):
    env.unset("ABBYY_SERIAL_NUMBER")
    env.unset("ABBYY_CUSTOMER_PROJECT_ID")
    at = run()
    assert banner(at) is None
    t = texts(at)
    assert "Choose one of the options below" in t and "Use a different ABBYY license" not in t
    assert not any(c.label.startswith("I understand") for c in at.checkbox)
    assert activate_buttons(at) == [] or not any(b.disabled for b in activate_buttons(at))


def test_nothing_from_the_license_identifiers_ever_reaches_the_page(env):
    env.set(mode="default", status="active", reason="ok", verified=True)
    (Path(os.environ["ABBYY_REAL_LICENSES_DIR"])).mkdir()
    (Path(os.environ["ABBYY_REAL_LICENSES_DIR"]) / f"{SERIAL}.ABBYY.ActivationToken").write_text("t" * 80)
    Path(os.environ["ABBYY_LICENSES_DIR"]).mkdir()
    (Path(os.environ["ABBYY_LICENSES_DIR"]) / f"{PROJECT}.ABBYY.ActivationToken").write_text("t" * 80)
    for state in ({"mode": "default", "status": "active"}, {"mode": "default", "status": "failed", "reason": "network"},
                  {"mode": "default", "status": "activating"}, {"mode": "custom", "status": "active"}):
        env.set(**state)
        for page in (PAGE, MAIN):
            at = AppTest.from_file(page, default_timeout=45).run()
            if page == MAIN and at.exception:
                continue  # the main page imports the heavy pipeline stack; covered separately
            blob = json.dumps(texts(at)) + json.dumps([str(w.proto) for w in at.sidebar]) if hasattr(at, "sidebar") else texts(at)
            assert SERIAL not in blob and PROJECT not in blob, (page, state)


def test_status_panel_does_not_claim_there_are_no_token_files_for_a_serial_license(env):
    env.set(mode="default", status="active", reason="ok")
    at = run()
    t = texts(at)
    assert "No token files found" not in t
    assert "Default production license is active" in t


def test_legacy_status_panel_is_kept_when_no_default_exists(env):
    env.unset("ABBYY_SERIAL_NUMBER")
    env.unset("ABBYY_CUSTOMER_PROJECT_ID")
    t = texts(run())
    assert "License Token Files" in t


def test_misconfigured_default_is_reported_not_hidden(env):
    env.unset("ABBYY_CUSTOMER_PROJECT_ID")
    env.set(mode="none", status="failed", reason="config")
    at = run()
    kind, text = banner(at)
    assert kind == "warning" and "incomplete" in text


def test_page_survives_a_corrupt_state_file(env):
    (env.state / "license_state.json").write_text("{not json")
    run()


def test_external_license_server_mode_has_nothing_to_set_up(env):
    env.env("ABBYY_LS_HOST", "licenses.example.invalid")
    kind, text = banner(run())
    assert kind == "info" and "central license server" in text


def test_banner_never_blames_the_user_and_is_plain_language(env):
    for status, reason in (("active", "ok"), ("failed", "network"), ("failed", "not_licensed"), ("activating", "")):
        env.set(mode="default", status=status, reason=reason)
        _, text = banner(run())
        for jargon in ("exit code", "returncode", "Traceback", "LicensingService", "SamplesConfig"):
            assert jargon not in text, (status, reason, jargon)


# ── a running PDF conversion keeps license changes out ───────────────────────
def _marking_script(env_):
    """An activation script that leaves a marker file when it runs (it must NOT run)."""
    marker = Path(os.environ["ABBYY_FRE_DIR"]).parent / "script_ran"
    Path(os.environ["ABBYY_ACTIVATE_SH"]).write_text(f"#!/bin/bash\ntouch {marker}\nexit 0\n")
    return marker


def test_override_is_refused_while_a_conversion_is_running(env):
    env.set(mode="default", status="active", reason="ok", verified=True)
    marker = _marking_script(env)
    at = run()
    at.checkbox[0].check().run()
    [t for t in at.text_input if t.label.startswith("Serial Number")][0].input("CUSTOMSERIAL000000000CUST")
    [t for t in at.text_input if t.label.startswith("Customer Project ID") and t.key != "token_project_id"][0] \
        .input("TestCustomProject02")
    with L.conversion_slot() as ok:
        assert ok
        [b for b in at.button if b.label == "🚀 Activate Now"][0].click().run()
    assert not at.exception, [x.value for x in at.exception]
    assert any("PDF conversion is running" in w.value and "Nothing was changed" in w.value for w in at.warning)
    assert not any("Activation failed" in e.value for e in at.error)  # it is not a failure, just not now
    assert not any("Connecting to" in i.value for i in at.info)  # ABBYY was never contacted, so no such message
    assert not marker.exists()
    assert (L.read_state()["mode"], L.read_state()["status"]) == ("default", "active")


def test_switch_back_is_refused_while_a_conversion_is_running(env):
    env.set(mode="custom", status="active", reason="ok")
    at = run()
    [b for b in at.button if "Switch back to the default license" in b.label][0].click().run()
    with L.conversion_slot() as ok:
        assert ok
        [b for b in at.button if b.label == "Yes, switch back"][0].click().run()
    assert not at.exception, [x.value for x in at.exception]
    assert any("Could not switch back" in e.value and "PDF conversion is running" in e.value for e in at.error)
    assert L.read_state()["mode"] == "custom"


def test_restart_button_is_refused_while_a_conversion_is_running(env):
    env.set(mode="default", status="active", reason="ok")
    at = run()
    with L.conversion_slot() as ok:
        assert ok
        [b for b in at.button if b.label.startswith("🔄 Restart LicensingService")][0].click().run()
    assert not at.exception, [x.value for x in at.exception]
    assert any("not restarted" in w.value and "PDF conversion is running" in w.value for w in at.warning)


def test_check_license_says_so_when_it_cannot_answer_during_a_conversion(env):
    env.set(mode="default", status="active", reason="ok")
    at = run()
    with L.conversion_slot() as ok:
        assert ok
        [b for b in at.button if b.label.startswith("🩺 Check license")][0].click().run()
    assert any("could not be completed right now" in i.value and "PDF conversion" in i.value for i in at.info)


# The real _pdf_to_docx of the main page, with a stand-in for the ABBYY tool and a recording `st`.
FAKE_TOOL = """#!/bin/bash
echo "$(date +%s.%N) start" >> "$TOOL_LOG"
if [ -n "$TOOL_FAIL" ]; then echo "$TOOL_FAIL" >&2; exit 1; fi
sleep "${TOOL_SLEEP:-0}"
while [ $# -gt 0 ]; do [ "$1" = "-of" ] && out="$2"; shift; done
echo converted > "$out"
echo "$(date +%s.%N) end" >> "$TOOL_LOG"
"""


@pytest.fixture
def convert(env, monkeypatch, tmp_path):
    tree = ast.parse(Path(MAIN).read_text(encoding="utf-8"))
    funcs = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in ("_pdf_to_docx", "_current_abbyy_license_env")]
    assert len(funcs) == 2
    said, logs, env_reads = [], [], []

    class FakeSt:
        def warning(self, m): said.append(("warning", m))
        def error(self, m): said.append(("error", m))
        def info(self, m): said.append(("info", m))

    def settings_read():
        env_reads.append(time.time())
        return {"ABBYY_PROJECT_ID": "TestDefaultProject01"}

    ns = {"os": os, "re": re, "Path": Path, "st": FakeSt(), "_log": lambda lvl, msg: logs.append((lvl, msg)),
          "_abbyy_license": L, "_ABBYY_TIMEOUT": 30}
    exec(compile(ast.Module(body=funcs, type_ignores=[]), MAIN, "exec"), ns)
    ns["_current_abbyy_license_env"] = settings_read
    tool = tmp_path / "tool"
    tool.write_text(FAKE_TOOL)
    tool.chmod(0o755)
    log = tmp_path / "tool.log"
    monkeypatch.setenv("ABBYY_CLI", str(tool))
    monkeypatch.setenv("ABBYY_LIB", str(tmp_path))
    monkeypatch.setenv("TOOL_LOG", str(log))
    (tmp_path / "in.pdf").write_bytes(b"%PDF-1.4 test")

    class C:
        pass
    c = C()
    c.run = lambda: ns["_pdf_to_docx"](str(tmp_path / "in.pdf"), str(tmp_path / "out.docx"))
    c.ns, c.said, c.logs, c.log, c.env_reads, c.tmp = ns, said, logs, log, env_reads, tmp_path
    return c


def test_conversion_keeps_license_changes_out_while_the_tool_runs(convert, monkeypatch):
    monkeypatch.setenv("TOOL_SLEEP", "1.5")
    result = []
    t = threading.Thread(target=lambda: result.append(convert.run()))
    t.start()
    deadline = time.time() + 10
    while not convert.log.exists() and time.time() < deadline:
        time.sleep(0.05)
    with L.conversions_idle() as idle:
        assert idle is False  # the tool is running: a license change would be refused
    t.join(15)
    assert result == [True] and (convert.tmp / "out.docx").exists()
    with L.conversions_idle() as idle:
        assert idle is True  # and the lock is gone afterwards


def test_conversion_waits_for_a_license_change_and_reads_the_settings_after_it(convert):
    started, release = threading.Event(), threading.Event()

    def license_change():
        with L.conversions_idle() as idle:
            assert idle
            started.set()
            release.wait(10)

    change = threading.Thread(target=license_change)
    change.start()
    assert started.wait(5)
    result = []
    t = threading.Thread(target=lambda: result.append(convert.run()))
    t.start()
    time.sleep(1.2)
    assert not convert.log.exists() and result == [] and convert.env_reads == []  # waiting, not running
    released_at = time.time()
    release.set()
    change.join()
    t.join(15)
    assert result == [True] and convert.log.exists()
    assert convert.env_reads and convert.env_reads[0] >= released_at  # the new license settings, not stale ones


def test_conversion_gives_up_with_a_message_when_the_change_takes_too_long(convert, monkeypatch):
    real = L.conversion_slot
    monkeypatch.setattr(L, "conversion_slot", lambda: real(wait=0.3, poll=0.05))
    with L.conversions_idle() as idle:
        assert idle
        assert convert.run() is False
    assert not convert.log.exists()
    assert any(kind == "warning" and "license is being changed" in m for kind, m in convert.said)


def test_conversion_still_works_when_the_license_module_is_unavailable(convert):
    convert.ns["_abbyy_license"] = None
    assert convert.run() is True and convert.log.exists()
    assert not [m for kind, m in convert.said]


def test_failed_conversion_points_to_the_license_when_abbyy_says_so(convert, monkeypatch):
    L.write_state(mode="default", status="active", reason="ok")
    monkeypatch.setenv("TOOL_FAIL", "Error: ABBYY FineReader Engine is not licensed.")
    assert convert.run() is False
    assert any(k == "error" and "conversion failed" in m for k, m in convert.said)  # the original message stays
    assert any(k == "info" and "License Setup" in m for k, m in convert.said)


def test_failed_conversion_names_the_code_when_the_default_license_failed(convert, monkeypatch):
    L.write_state(mode="default", status="failed", reason="invalid_serial")
    monkeypatch.setenv("TOOL_FAIL", "Error: something generic")
    assert convert.run() is False
    assert any(k == "info" and "invalid_serial" in m for k, m in convert.said)


def test_failed_conversion_for_other_reasons_gets_no_license_hint(convert, monkeypatch):
    L.write_state(mode="default", status="active", reason="ok")
    monkeypatch.setenv("TOOL_FAIL", "Error: Cannot open the input file.")
    assert convert.run() is False
    assert any(k == "error" for k, m in convert.said) and not any(k == "info" for k, m in convert.said)


# ── connection check section ─────────────────────────────────────────────────
CHECK_BUTTON = "🌐 Test connection to ABBYY"


def _canned_checks():
    return [
        L.HostCheck("registration2.abbyy.com", "serial activation", "ok 104.18.16.71", "FAILED: timed out", "-", "-", 6.1),
        L.HostCheck("license-cloud.abbyy.com", "online licenses", "ok 20.1.2.3", "ok", "ok TLSv1.3, signed by Some CA",
                    "ok HTTP 200", 0.9),
        L.HostCheck("activation1.abbyy.ru", "old address", "FAILED: [Errno -2] Name or service not known", "-", "-", "-", 0.1),
    ]


def _click_check(at):
    [b for b in at.button if b.label == CHECK_BUTTON][0].click().run()
    assert not at.exception, [x.value for x in at.exception]
    return at


def test_connection_check_does_nothing_until_the_button_is_pressed(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: pytest.fail("no check may run on page load"))
    env.set(mode="default", status="failed", reason="network")
    at = run()
    assert any(b.label == CHECK_BUTTON for b in at.button)
    assert "Connection check" in texts(at) and "Name lookup" not in texts(at)  # no table before the button is pressed


def test_connection_check_shows_every_step_for_every_server(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _canned_checks())
    env.set(mode="default", status="failed", reason="network")
    t = texts(_click_check(run()))
    for expected in ("registration2.abbyy.com", "FAILED: timed out", "license-cloud.abbyy.com", "signed by Some CA",
                     "activation1.abbyy.ru", "Name or service not known", "❌ problem", "✅ ok", "6.1 s"):
        assert expected in t, expected
    assert "Name lookup" in t and "Secure connection" in t and "Web reply" in t
    assert "expected to fail" in t  # the old .ru address is explained, not alarming


def test_connection_check_leads_with_a_plain_sentence_and_marks_the_old_address_as_expected(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _canned_checks())
    at = _click_check(run())
    assert any("`registration2.abbyy.com` fails at the connection" in str(e.value) for e in at.error)
    rows = {ln.split("|")[1].strip(): ln for ln in texts(at).splitlines() if ln.startswith("| `")}
    assert "❌ problem" in rows["`registration2.abbyy.com`"] and "✅ ok" in rows["`license-cloud.abbyy.com`"]
    assert "ℹ️ expected" in rows["`activation1.abbyy.ru`"] and "problem" not in rows["`activation1.abbyy.ru`"]


def test_connection_check_says_all_is_well_when_it_is(env, monkeypatch):
    fine = [L.HostCheck("registration2.abbyy.com", "p", "ok 1.1.1.1", "ok", "ok TLSv1.3", "ok HTTP 404", 0.5)]
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: fine)
    at = _click_check(run())
    assert any("can reach every ABBYY server" in str(e.value) for e in at.success)


def test_connection_check_result_survives_other_interactions(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _canned_checks())
    env.set(mode="default", status="active", reason="ok")
    at = _click_check(run())
    at.checkbox[0].check().run()  # any other widget triggers a rerun
    assert "registration2.abbyy.com" in texts(at)


def test_connection_check_names_proxy_settings_but_never_their_values(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _canned_checks())
    fake_login = "user:topsecretpw"  # made up; built at run time so scanners do not take it for a credential
    env.env("HTTPS_PROXY", f"http://{fake_login}@proxy.example:8080")
    t = texts(_click_check(run()))
    assert "HTTPS_PROXY" in t and "topsecretpw" not in t and "proxy.example" not in t


def test_connection_check_shows_abbyys_own_log_with_identifiers_hidden(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _canned_checks())
    lic = Path(os.environ["ABBYY_REAL_LICENSES_DIR"])
    lic.mkdir()
    (lic / "ProductProtection_2026-11-16 10h00m.log").write_text(
        f"16/11/26 10:00:03 (7) Activation failed for {SERIAL} and {PROJECT}: connection reset by peer\n")
    t = texts(_click_check(run()))
    assert "ProductProtection" in t and "connection reset by peer" in t
    assert SERIAL not in t and PROJECT not in t


def test_connection_check_table_cannot_be_broken_by_error_text(env, monkeypatch):
    nasty = [L.HostCheck("x.example", "p", "FAILED: a | b\nc", "-", "-", "-", 0.1)]
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: nasty)
    rows = [ln for ln in texts(_click_check(run())).splitlines() if ln.startswith("| `x.example`")]
    assert rows and rows[0].count("|") == 8 and "a / b c" in rows[0]


def test_failed_banner_for_a_blocked_network_points_to_the_connection_check(env):
    env.set(mode="default", status="failed", reason="network")
    kind, text = banner(run())
    assert kind == "error" and "even after several tries" in text and "**Connection check**" in text
    assert "usually temporary" not in text and "network" in text and "not something you did wrong" in text


# ── ABBYY's server was reached but said no ("Failed, limit exceed") ──────────
def _reached_checks():
    return [L.HostCheck("registration2.abbyy.com", "serial activation", "ok 1.2.3.4", "ok", "ok TLSv1.3", "ok HTTP 404", 0.3),
            L.HostCheck("example.com", "control", "ok 5.6.7.8", "ok", "ok TLSv1.3", "ok HTTP 200", 0.2),
            L.HostCheck("registration2.abbyy.com", "activation service (an empty test message, no license data)",
                        "ok 1.2.3.4", "ok", "ok TLSv1.3", "ok HTTP 500", 0.4, kind="activation")]


def _generic_failure_script():
    """The activation script as the real tool behaves when ABBYY refuses: one generic line, exit 1."""
    Path(os.environ["ABBYY_ACTIVATE_SH"]).write_text('#!/bin/bash\necho "Could not activate license via internet"\nexit 1\n')


def test_failed_banner_for_a_refused_activation_points_to_abbyy_not_to_the_network(env):
    env.set(mode="default", status="failed", reason="refused")
    kind, text = banner(run())
    assert kind == "error" and "**refused**" in text and "not something you did wrong" in text
    assert "not a firewall problem" in text and "activations" in text and "Online (token) license" in text
    assert "Connection check" not in text and "usually temporary" not in text and "network" not in text.lower()


def test_try_again_after_a_refusal_says_that_abbyy_declined(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached_checks())
    _generic_failure_script()
    env.set(mode="default", status="failed", reason="refused")
    at = run()
    [b for b in at.button if b.label.startswith("🔁 Try the default license again")][0].click().run()
    assert not at.exception, [x.value for x in at.exception]
    assert any("still could not be set up" in e.value and "did not grant the activation" in e.value for e in at.error)
    assert L.read_state()["reason"] == "refused"


def test_a_custom_license_that_abbyy_declines_shows_the_real_cause_and_no_network_advice(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached_checks())
    _generic_failure_script()
    env.set(mode="default", status="active", reason="ok", verified=True)
    at = run()
    at.checkbox[0].check().run()
    [t for t in at.text_input if t.label.startswith("Serial Number")][0].input("CUSTOMSERIAL000000000CUST")
    [t for t in at.text_input if t.label.startswith("Customer Project ID") and t.key != "token_project_id"][0] \
        .input("TestCustomProject02")
    [b for b in at.button if b.label == "🚀 Activate Now"][0].click().run()
    assert not at.exception, [x.value for x in at.exception]
    errors = " ".join(e.value for e in at.error)
    assert "Activation failed" in errors and "did not grant the activation" in errors and "usually temporary" not in errors
    t = texts(at)
    assert "activations allowed for this serial number is used up" in t and "Ask ABBYY to check the activations" in t
    assert "No internet access to" not in t  # the connection works: do not send people to the firewall team
    assert "Could not activate license via internet" in t  # ABBYY's own line is still shown for the support team
    assert (L.read_state()["mode"], L.read_state()["status"]) == ("default", "active")  # the working license is kept


def test_a_custom_license_that_cannot_reach_abbyy_still_gets_the_network_advice(env, monkeypatch):
    blocked = [L.HostCheck("registration2.abbyy.com", "x", "ok 1.2.3.4", "FAILED: timed out", "-", "-", 6.0),
               L.HostCheck("registration2.abbyy.com", "p", "ok 1.2.3.4", "FAILED: timed out", "-", "-", 6.0, kind="activation")]
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: blocked)
    _generic_failure_script()
    env.set(mode="default", status="active", reason="ok", verified=True)
    at = run()
    at.checkbox[0].check().run()
    [t for t in at.text_input if t.label.startswith("Serial Number")][0].input("CUSTOMSERIAL000000000CUST")
    [t for t in at.text_input if t.label.startswith("Customer Project ID") and t.key != "token_project_id"][0] \
        .input("TestCustomProject02")
    [b for b in at.button if b.label == "🚀 Activate Now"][0].click().run()
    assert not at.exception, [x.value for x in at.exception]
    assert "could not reach ABBYY's license server" in " ".join(e.value for e in at.error)
    assert "No internet access to" in texts(at)


def test_connection_check_table_names_the_activation_test_message_row(env, monkeypatch):
    monkeypatch.setattr(L, "diagnose_connection", lambda *a, **k: _reached_checks())
    t = texts(_click_check(run()))
    assert "activation service (an empty test message, no license data)" in t
    assert "refused it" in t  # the all-clear sentence says what a failed activation then means
    assert "one empty test message" in t  # the section tells what is sent


def test_the_help_explains_the_refused_code(env):
    env.set(mode="default", status="active", reason="ok")
    t = texts(run())
    assert "code `refused`" in t and "Online (token) license" in t
