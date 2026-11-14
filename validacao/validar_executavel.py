"""Verifica o EXE com os arquivos do case, sem depender do pacote fonte no subprocesso."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sqlite3
import subprocess
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def check_artifacts(folder):
    summary = json.loads((folder / "resumo.json").read_text(encoding="utf-8"))
    assert summary["files"] == 17
    assert summary["statuses"] == {"ok": 12, "revisar": 4, "falhou": 1}
    assert summary["approved_balance"] == "4114372.29"
    manifest = json.loads((folder / "manifesto.json").read_text(encoding="utf-8"))
    for name, digest in manifest["artifacts"].items():
        assert hashlib.sha256((folder / name).read_bytes()).hexdigest() == digest
    for name, digest in manifest["inputs"].items():
        assert hashlib.sha256((ROOT / "extratos" / name).read_bytes()).hexdigest() == digest
    return summary


def main():
    executable = ROOT / ".local/bin/Goener.exe"
    target = ROOT / ".local/validacao_exe"
    target.mkdir(parents=True, exist_ok=True)
    shutil.copy2(executable, target / "Goener.exe")
    for name in ("config.toml", "modelo.xlsx", ".env"):
        if (ROOT / name).is_file():
            shutil.copy2(ROOT / name, target / name)
    shutil.copytree(ROOT / "extratos", target / "extratos", dirs_exist_ok=True)
    outside = ROOT / ".local/cwd_fora"
    outside.mkdir(exist_ok=True)
    env = os.environ.copy()
    for key in ("PYTHONPATH", "PYTHONHOME", "GEMINI_API_KEY", "GOOGLE_API_KEY", "OPENAI_API_KEY"):
        env.pop(key, None)
    env.update(
        PATH=str(Path(os.environ["WINDIR"]) / "System32"),
        PYTHONNOUSERSITE="1",
        HTTP_PROXY="http://127.0.0.1:9",
        HTTPS_PROXY="http://127.0.0.1:9",
        NO_PROXY="",
    )
    report = {
        "executable_sha256": hashlib.sha256(executable.read_bytes()).hexdigest(),
        "executable_bytes": executable.stat().st_size,
        "conditions": {
            "python_installed_on_host": True,
            "python_on_subprocess_path": False,
            "source_package_in_test_folder": False,
            "outside_working_directory": True,
            "network_blocked_by_proxy": True,
            "clean_windows_vm": False,
        },
        "runs": [],
    }

    def execute(label, args):
        started = time.perf_counter()
        process = subprocess.run(
            [str(target / "Goener.exe"), *args],
            cwd=outside,
            env=env,
            capture_output=True,
            text=True,
            timeout=300,
        )
        assert process.returncode == 0, process.stderr[-1500:]
        latest = json.loads((target / ".local/outputs/latest.json").read_text(encoding="utf-8"))
        folder = Path(latest["directory"])
        summary = check_artifacts(folder)
        destination = ROOT / "validacao/execucoes" / label
        destination.mkdir(parents=True, exist_ok=True)
        for artifact in folder.iterdir():
            if artifact.is_file():
                shutil.copy2(artifact, destination / artifact.name)
        (destination / "console.log").write_text(process.stdout + process.stderr, encoding="utf-8")
        report["runs"].append(
            {
                "label": label,
                "args": args,
                "exit_code": process.returncode,
                "wall_seconds": round(time.perf_counter() - started, 3),
                "run_id": summary["run_id"],
                "processing_seconds": summary["duration_seconds"],
                "git_commit": summary["git_commit"],
                "statuses": summary["statuses"],
                "cache_hits": summary["cache_hits"],
                "ai_summary": summary["ai_summary"],
            }
        )
        return summary

    first = execute("executavel_sem_ia", ["--ai-mode", "off", "--no-cache"])
    assert first["cache_hits"] == 0 and first["ai_summary"]["http_attempts"] == 0
    # Cache não é distribuído. Copiamos as análises reais locais apenas para validar
    # o reaproveitamento no EXE sem fazer novas chamadas externas neste ensaio.
    runtime = target / ".local/runtime"
    for name in ("extractions.sqlite3", "gemini.sqlite3"):
        source_path = ROOT / ".local/runtime" / name
        if not source_path.exists():
            continue
        source = sqlite3.connect(source_path)
        destination = sqlite3.connect(runtime / name)
        try:
            source.backup(destination)
        finally:
            destination.close()
            source.close()
    if (ROOT / ".local/runtime/gemini.sqlite3").exists() and (target / ".env").exists():
        second = execute("executavel_ia_cache", [])
        assert second["ai_summary"]["cached_decisions"] == 5
        assert second["ai_summary"]["http_attempts"] == 0
    (ROOT / "validacao/executavel.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
