import json
import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from goener.models import Result, Statement, money
from goener.pipeline import RunLock, file_hash, run
from goener.reporting import HEADERS, write_workbook

ROOT = Path(__file__).resolve().parents[1]


def source(config, name="pau_ferro_out26.xlsx", destination=None):
    p = config.input_dir / (destination or name)
    shutil.copyfile(ROOT / "extratos" / name, p)
    return p


def test_batch_preserves_duplicates_and_failures_and_has_evidence(config):
    source(config)
    source(config, destination="pau_ferro_out26 (1).xlsx")
    (config.input_dir / "quebrado.pdf").write_bytes(b"not a pdf")
    source(config, "aviso_agencia_horizonte.pdf")
    before = {p.name: file_hash(p) for p in config.input_dir.iterdir()}
    out, summary, code = run(config)
    assert code == 0
    assert summary["files"] == 4 and summary["statuses"] == {"revisar": 1, "ok": 1, "falhou": 2}
    assert summary["duplicates"] == 1 and summary["approved_balance"] == "225934.57"
    assert {p.name: file_hash(p) for p in config.input_dir.iterdir()} == before
    rows = load_workbook(out / "consolidado.xlsx")["Consolidado"]
    assert rows.max_row == 5
    assert [c.value for c in rows[1]] == HEADERS
    manifest = json.loads((out / "manifesto.json").read_text(encoding="utf8"))
    assert all(file_hash(out / n) == sha for n, sha in manifest["artifacts"].items())
    events = [
        json.loads(line)
        for line in (out / "execucao.jsonl").read_text(encoding="utf8").splitlines()
    ]
    assert sum(e["event"] == "FILE_FINISHED" for e in events) == 4


def test_resume_revalidates_period_and_produces_same_rows(config):
    source(config)
    out1, _, _ = run(config)
    out2, summary, _ = run(config)
    assert summary["cache_hits"] == 1
    a = list(load_workbook(out1 / "consolidado.xlsx")["Consolidado"].values)
    b = list(load_workbook(out2 / "consolidado.xlsx")["Consolidado"].values)
    assert a == b
    config.period = "2026-09"
    _, summary, _ = run(config)
    assert summary["cache_hits"] == 1 and summary["statuses"] == {"revisar": 1}
    assert summary["approved_balance"] == "0.00"


def test_strict_delivers_artifacts_but_returns_two(config):
    source(config, "aviso_agencia_horizonte.pdf")
    config.strict = True
    out, summary, code = run(config)
    assert code == 2 and summary["exit_code"] == 2
    assert (out / "consolidado.xlsx").is_file()


def test_stop_registers_every_unprocessed_file(config):
    source(config)
    config.runtime_dir.mkdir()
    config.stop_file.touch()
    out, summary, code = run(config)
    assert code == 3 and summary["statuses"] == {"falhou": 1}
    assert "CANCELADO_STOP" in (out / "evidencias.json").read_text(encoding="utf8")


def test_lock_prevents_overlapping_execution(config):
    source(config)
    with RunLock(config.runtime_dir / "execution.lock"):
        with pytest.raises(RuntimeError, match="bloqueada"):
            run(config)


def test_failed_publication_does_not_advance_latest(config, monkeypatch):
    source(config)
    out, _, _ = run(config)
    latest = (config.output_dir / "latest.json").read_bytes()

    def denied(*args):
        raise PermissionError("Destino indisponível")

    monkeypatch.setattr("goener.pipeline.write_workbook", denied)
    with pytest.raises(PermissionError):
        run(config)
    assert (config.output_dir / "latest.json").read_bytes() == latest
    assert out.is_dir()
    assert not (config.runtime_dir / "execution.lock").exists()


def test_excel_output_does_not_turn_text_into_formulas(config):
    s = Statement(bank="=WEBSERVICE(1)", account="0001-0", period="2026-10", closing=money(0))
    r = Result("=SUM(1,1).pdf", "sha", "PDF", s)
    out = config.output_dir
    out.mkdir()
    write_workbook(config.template, out / "result.xlsx", [r])
    wb = load_workbook(out / "result.xlsx", data_only=False)
    assert wb["Consolidado"]["A2"].data_type == "s"
    assert wb["Consolidado"]["B2"].data_type == "s"
    assert wb["Consolidado"]["E2"].value == 0
    assert wb.sheetnames == load_workbook(config.template).sheetnames


def test_empty_input_is_not_reported_as_success(config):
    with pytest.raises(ValueError, match="vazia"):
        run(config)
