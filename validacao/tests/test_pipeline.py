import json
import shutil
import subprocess
from pathlib import Path

import pytest
from openpyxl import load_workbook

from goener.core import Result, Statement, money
from goener.pipeline import RunLock, file_hash, publish_shortcuts, run
from goener.reporting import HEADERS, write_workbook

ROOT = Path(__file__).resolve().parents[2]


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
    assert (
        summary["git_commit"]
        == subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()
    )
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
    latest = (config.output_dir / "historico/latest.json").read_bytes()

    def denied(*args):
        raise PermissionError("Destino indisponível")

    monkeypatch.setattr("goener.pipeline.write_workbook", denied)
    with pytest.raises(PermissionError):
        run(config)
    assert (config.output_dir / "historico/latest.json").read_bytes() == latest
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
    assert wb.sheetnames == ["Consolidado", "Resumo", "Revisões", "Instruções"]


def test_empty_input_is_not_reported_as_success(config):
    with pytest.raises(ValueError, match="vazia"):
        run(config)


def test_html_keeps_ai_advice_separate_and_opens_source_page(config):
    from goener.reporting import write_html

    s = Statement("Banco", "1-0", "2026-10", closing=money(10))
    r = Result("extrato com espaço.pdf", "sha", "PDF", s)
    r.ai = {
        "decision": "revisar",
        "message": "Conferir R$ 1.080,00 <script>alert(1)</script>",
        "source_disagreements": ["saldo final"],
        "source": {"pages": 2},
        "source_check": {"location": "Página 2, saldo final", "quote": "Saldo 99,00"},
    }
    summary = {
        "run_id": "test",
        "period": "2026-10",
        "duration_seconds": 0.1,
        "approved_balance": "10.00",
        "statuses": {"ok": 1},
        "ai_mode": "consultivo",
        "configuration": {"input_dir": str(config.input_dir), "ai_scope": "todos"},
    }
    config.output_dir.mkdir()
    path = config.output_dir / "relatorio.html"
    write_html(path, summary, [r])
    text = path.read_text(encoding="utf-8")
    assert 'data-attention="true"' in text and 'data-ai="analisada"' in text
    assert "extrato%20com%20espa%C3%A7o.pdf#page=2" in text
    assert "&lt;script&gt;" in text and "<script>alert(1)</script>" not in text
    assert "R$ 1.080,00" in text
    assert "Onde conferir" in text and "indicação da IA" in text
    assert r.status == "ok" and r.included("consultivo", "2026-10")


def test_workbook_retains_model_with_human_review_reference(config):
    from goener.core import Transaction

    s = Statement("Banco", "1-0", "2026-10", closing=money(10), evidence={"closing": {"page": 3}})
    s.transactions = [
        Transaction(
            "2026-10-26", "PIX repetido", money(10), money(10), location="página 2, linha 10"
        ),
        Transaction(
            "2026-10-26", "PIX repetido", money(10), money(20), location="página 2, linha 11"
        ),
    ]
    s.issue("LANCAMENTO_REPETIDO", "Possível lançamento repetido")
    r = Result("extrato.pdf", "sha", "PDF", s)
    assert r.review_reference["page"] == 2
    config.output_dir.mkdir()
    path = config.output_dir / "consolidado.xlsx"
    write_workbook(config.template, path, [r])
    wb = load_workbook(path)
    ws = wb["Consolidado"]
    assert ws.max_column == 8 and [c.value for c in ws[1]] == HEADERS
    assert ws["A1"].fill.fgColor.rgb.endswith("003836")
    assert (
        "Onde conferir:" in ws["H2"].value
        and "linha 10" in ws["H2"].value
        and "linha 11" in ws["H2"].value
    )
    assert not ws.sheet_view.showGridLines and ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == "A1:H2"
    wb.close()


def test_summary_uses_editable_table_and_formulas_without_changing_source(config):
    r = Result("extrato.pdf", "sha", "PDF", Statement("Banco", "1-0", "2026-10", closing=money(10)))
    config.output_dir.mkdir()
    path = config.output_dir / "formulas.xlsx"
    write_workbook(config.template, path, [r], "2026-11")
    wb = load_workbook(path, data_only=False)
    assert [c.value for c in wb["Consolidado"][1]] == HEADERS
    assert wb["Consolidado"].tables["Extratos"].ref == "A1:H2"
    assert wb["Consolidado"]["E2"].value == 10
    assert wb["Resumo"]["B3"].value == "2026-11"
    assert (
        wb["Resumo"]["B6"].value
        == '=SUMIFS(Extratos[Saldo final],Extratos[Status],"ok",Extratos[Período],$B$3)'
    )
    assert all(wb["Resumo"][f"B{i}"].data_type == "f" for i in [6, 7, 9, 10, 11, 12, 13, 14])
    assert wb.calculation.calcMode == "auto" and wb.calculation.fullCalcOnLoad
    assert len(wb["Consolidado"].data_validations.dataValidation) == 1
    wb.close()


def test_package_excludes_env_by_default_and_includes_only_when_requested(
    config, tmp_path, monkeypatch
):
    import zipfile
    from types import SimpleNamespace

    from goener import entrega

    source(config)
    run_dir, _, _ = run(config)
    secret = "fake-private-demo-key"
    (tmp_path / ".env").write_text("GEMINI_API_KEY=" + secret, encoding="utf-8")
    (tmp_path / ".env.example").write_text("GEMINI_API_KEY=", encoding="utf-8")
    monkeypatch.setattr(entrega, "ROOT", tmp_path)
    monkeypatch.setattr(
        entrega.subprocess, "run", lambda *a, **kw: SimpleNamespace(stdout=b".env\0.env.example\0")
    )
    default = entrega.build(run_dir, tmp_path / "public")
    with zipfile.ZipFile(default) as z:
        assert "goener-case/.env" not in z.namelist()
        assert not any(secret.encode() in z.read(n) for n in z.namelist())
    private = entrega.build(run_dir, tmp_path / "private", include_env=True)
    with zipfile.ZipFile(private) as z:
        assert secret in z.read("goener-case/.env").decode()


def test_time_estimate_separates_full_statements_from_triage_and_annualizes():
    from goener.reporting import estimate_effort

    rows = [
        Result(
            str(i),
            str(i),
            "PDF",
            Statement(bank="Banco", account="1", period="2026-10", closing=money(10)),
        )
        for i in range(14)
    ]
    rows += [
        Result("copy", "0", "PDF", rows[0].statement, duplicate_of="0"),
        Result(
            "other",
            "other",
            "PDF",
            Statement(bank="Banco", account="1", period="2026-09", closing=money(10)),
        ),
        Result("notice", "notice", "PDF", Statement()),
    ]
    report = estimate_effort(
        rows,
        {
            "period": "2026-10",
            "duration_seconds": 60,
            "configuration": {
                "manual_minutes_per_file": 8,
                "manual_check_minutes_per_file": 2,
                "manual_sample_files": 1,
            },
        },
    )
    assert report["workload_counts"] == {
        "current_unique_statements": 14,
        "duplicates": 1,
        "other_period": 1,
        "insufficient_or_nonstatement": 1,
    }
    assert report["manual_sample_minutes"] == 10
    assert report["projected_manual_seconds"] == (14 * 10 + 3 * 1) * 60
    assert report["projected_difference_seconds"] == 8520
    assert report["scenarios"]["central"]["annual_hours"] == 28.4
    assert report["uncertainty"]["confidence_level"] is None
    assert report["uncertainty"]["empirical_confidence_interval"] is None
    assert "assisted_human_minutes" not in report


def test_time_estimate_uses_actual_volume_and_frequency_without_fixed_file_count():
    from goener.reporting import estimate_effort

    rows = [
        Result(
            str(i),
            str(i),
            "Excel",
            Statement(bank="Banco", account="1", period="2026-10", closing=money(0)),
        )
        for i in range(3)
    ]
    report = estimate_effort(
        rows, {"period": "2026-10", "duration_seconds": 60, "configuration": {"monthly_batches": 2}}
    )
    assert report["workload_counts"]["current_unique_statements"] == 3
    assert report["triage_files"] == 0
    assert report["scenarios"]["central"]["annual_hours"] == 11.6
    assert report["current_statements_by_format"] == {"Excel": 3}
    assert (
        report["scenarios"]["lower"]["monthly_hours"]
        < report["scenarios"]["central"]["monthly_hours"]
        < report["scenarios"]["upper"]["monthly_hours"]
    )


def test_time_estimate_does_not_hide_negative_gain_or_invent_confidence():
    from goener.reporting import estimate_effort

    report = estimate_effort(
        [Result("notice", "a", "PDF", Statement())], {"period": "2026-10", "duration_seconds": 200}
    )
    assert report["projected_difference_seconds"] == -140
    assert report["uncertainty"]["type"] == "scenario_sensitivity_not_confidence_interval"


def test_html_does_not_include_time_estimate_or_manual_projection(config):
    from goener.reporting import write_html

    r = Result("extrato.pdf", "a", "PDF", Statement("Banco", "1", "2026-10", closing=money(10)))
    summary = {
        "period": "2026-10",
        "duration_seconds": 1,
        "approved_balance": "10.00",
        "run_id": "test",
        "statuses": {"ok": 1},
        "time_estimate": {"projected_manual_seconds": 600, "manual_source": "NAO_EXIBIR_NO_HTML"},
    }
    config.output_dir.mkdir()
    path = config.output_dir / "report.html"
    write_html(path, summary, [r])
    text = path.read_text(encoding="utf-8")
    assert "NAO_EXIBIR_NO_HTML" not in text and "projected_manual_seconds" not in text
    assert "capacidade liberada" not in text.lower() and "21 minutos" not in text.lower()
    assert "Composição do saldo" in text and 'data-tab="files-panel"' in text
    assert "__BANKS__" not in text and "__ROWS__" not in text


def test_latest_shortcuts_match_published_run(config):
    source(config)
    directory, _, _ = run(config)
    assert (config.output_dir / "consolidado.xlsx").read_bytes() == (
        directory / "consolidado.xlsx"
    ).read_bytes()
    assert f"historico/{directory.name}/relatorio.html" in (
        config.output_dir / "relatorio.html"
    ).read_text(encoding="utf-8")


def test_locked_shortcut_does_not_discard_published_run(tmp_path, monkeypatch):
    directory = tmp_path / "run"
    directory.mkdir()
    (directory / "consolidado.xlsx").write_bytes(b"published")
    events = []

    class Log:
        def event(self, name, **fields):
            events.append((name, fields))

    original = __import__("os").replace

    def locked(source, target):
        if Path(target).name == "consolidado.xlsx":
            raise PermissionError("Planilha aberta")
        return original(source, target)

    monkeypatch.setattr("goener.pipeline.os.replace", locked)
    publish_shortcuts(directory, Log())
    assert (directory / "consolidado.xlsx").read_bytes() == b"published"
    assert (tmp_path / "relatorio.html").is_file()
    assert events[0][0] == "SHORTCUT_UNAVAILABLE"
