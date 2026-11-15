import json
import subprocess
import sys

from executar import supervise
from goener.pipeline import RunLock


def test_supervisor_records_abrupt_exit_and_waits_for_desktop(tmp_path, monkeypatch):
    prompts = []
    monkeypatch.setattr("builtins.input", lambda message: prompts.append(message))
    code = supervise([sys.executable, "-c", "import os;os._exit(37)"], tmp_path, desktop=True)
    assert code == 37 and prompts
    logs = list((tmp_path / "resultados/controle/diagnosticos").glob("*.log"))
    assert "codigo=37" in logs[0].read_text(encoding="utf-8")


def test_success_does_not_pause(tmp_path, monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: (_ for _ in ()).throw(AssertionError("pause")))
    assert supervise([sys.executable, "-c", "print('ok')"], tmp_path, desktop=True) == 0


def test_lock_recovers_after_process_is_killed(tmp_path):
    path = tmp_path / "execution.lock"
    command = (
        "from pathlib import Path;from goener.pipeline import RunLock;import time;lock=RunLock(Path("
        + repr(str(path))
        + "));lock.__enter__();print('ready',flush=True);time.sleep(30)"
    )
    proc = subprocess.Popen([sys.executable, "-c", command], stdout=subprocess.PIPE, text=True)
    try:
        assert proc.stdout.readline().strip() == "ready"
        proc.kill()
        proc.wait(timeout=10)
        assert path.exists()
        with RunLock(path) as lock:
            assert lock.recovered["pid"] == proc.pid
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait()
        proc.stdout.close()
    assert not path.exists()


def test_legacy_live_process_is_not_removed(tmp_path):
    path = tmp_path / "execution.lock"
    path.write_text(json.dumps({"pid": __import__("os").getpid()}))
    import pytest

    with pytest.raises(RuntimeError, match="bloqueada"):
        with RunLock(path):
            pass
    assert path.exists()


def test_supervisor_keeps_diagnostic_in_selected_session(tmp_path):
    folder = tmp_path / "resultados/Fechamento teste"
    folder.mkdir(parents=True)
    script = (
        "import os;from pathlib import Path;Path(os.environ['GOENER_SESSION_TARGET']).write_text("
        + repr(str(folder))
        + ",encoding='utf-8');print('ok')"
    )
    assert supervise([sys.executable, "-c", script], tmp_path) == 0
    logs = list((folder / "runtime/diagnosticos").glob("*.log"))
    assert len(logs) == 1 and "codigo=0" in logs[0].read_text(encoding="utf-8")


def test_windows_kills_worker_and_descendant_when_supervisor_is_terminated(tmp_path):
    import os
    import time

    import pytest

    if os.name != "nt":
        pytest.skip("Windows Job Object")
    ready = tmp_path / "processos.json"
    worker = (
        "import subprocess,sys,os,time,json;from pathlib import Path;"
        'child=subprocess.Popen([sys.executable,"-c","import time;time.sleep(120)"]);'
        f"Path({str(ready)!r}).write_text(json.dumps([os.getpid(),child.pid]));"
        "time.sleep(120)"
    )
    script = (
        "from executar import supervise;from pathlib import Path;import sys;"
        f'supervise([sys.executable,"-c",{worker!r}],Path({str(tmp_path)!r}))'
    )
    parent = subprocess.Popen(
        [sys.executable, "-c", script], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )
    pids = []
    try:
        deadline = time.monotonic() + 15
        while not ready.exists() and time.monotonic() < deadline:
            assert parent.poll() is None, "Supervisor encerrou antes de iniciar os filhos"
            time.sleep(0.05)
        assert ready.exists()
        pids = json.loads(ready.read_text())
        assert all(RunLock.process_alive(pid) for pid in pids)
        parent.kill()  # Sem finally/KeyboardInterrupt: simula encerramento abrupto da janela.
        parent.wait(timeout=10)
        deadline = time.monotonic() + 10
        while any(RunLock.process_alive(pid) for pid in pids) and time.monotonic() < deadline:
            time.sleep(0.05)
        assert all(not RunLock.process_alive(pid) for pid in pids)
    finally:
        if parent.poll() is None:
            parent.kill()
            parent.wait(timeout=10)
        for pid in pids:
            if RunLock.process_alive(pid):
                os.kill(pid, 9)
