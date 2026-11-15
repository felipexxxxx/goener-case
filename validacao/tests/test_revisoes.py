import json
import shutil
from pathlib import Path

import pdfplumber
import pytest
from openpyxl import load_workbook

from goener.pipeline import file_hash, run

ROOT = Path(__file__).resolve().parents[2]
NAME = "extrato_sol_do_sertao_I_out2026.pdf"


def setup_case(config):
    shutil.copyfile(ROOT / "extratos" / NAME, config.input_dir / NAME)
    return run(config)[0]


def edit_review(config, decision="Corrigir e liberar", value=456099.12, reviewer="Avaliador"):
    path = config.output_dir / "consolidado.xlsx"
    wb = load_workbook(path)
    ws = wb["Revisões"]
    ws["F5"], ws["G5"] = decision, value
    ws["H5"], ws["I5"], ws["J5"] = "Saldo confirmado após conferência", reviewer, "14/11/2026"
    wb.save(path)
    wb.close()


def test_review_updates_excel_html_preserves_original_and_reuses_without_ai(config, monkeypatch):
    original = file_hash(config.input_dir / NAME) if (config.input_dir / NAME).exists() else None
    before = setup_case(config)
    original = original or file_hash(config.input_dir / NAME)
    edit_review(config)

    def forbidden(*a, **kw):
        pytest.fail("Atualizar revisões não deve extrair nem chamar IA")

    monkeypatch.setattr("goener.pipeline.process_file", forbidden)
    monkeypatch.setattr("goener.pipeline.review", forbidden)
    after, summary, code = run(config, reviews_only=True)
    assert code == 0
    assert summary["approved_balance"] == "456099.12"
    assert summary["human_reviews_applied"] == 1
    assert summary["ai_summary"]["http_attempts"] == 0
    wb = load_workbook(after / "consolidado.xlsx")
    assert wb["Consolidado"]["E8"].value == 456099.12
    assert wb["Consolidado"]["G8"].value == "Incluído"
    assert wb["Revisões"]["D5"].value == 457179.12
    assert wb["Consolidado"]["E2"].data_type == "f"
    wb.close()
    assert "Revisão humana registrada" in (after / "relatorio.html").read_text(encoding="utf-8")
    assert file_hash(config.input_dir / NAME) == original
    assert before.exists()
    _, repeated, _ = run(config, reviews_only=True)
    assert repeated["approved_balance"] == "456099.12"


@pytest.mark.parametrize(
    "decision,value,reviewer,applied,total",
    [
        ("Corrigir e liberar", 0, "Avaliador", 1, "0.00"),
        ("Liberar", None, "Avaliador", 1, "457179.12"),
        ("Excluir", None, "Avaliador", 1, "0.00"),
        ("Corrigir e liberar", 123, "", 0, "0.00"),
    ],
)
def test_review_decisions_and_required_fields(config, decision, value, reviewer, applied, total):
    setup_case(config)
    edit_review(config, decision, value, reviewer)
    _, summary, _ = run(config, reviews_only=True)
    assert summary["human_reviews_applied"] == applied
    assert summary["approved_balance"] == total


def test_modified_source_rejects_quick_refresh_and_invalidates_old_review(config):
    setup_case(config)
    edit_review(config)
    run(config, reviews_only=True)
    source = config.input_dir / NAME
    source.write_bytes(source.read_bytes() + b"\n%updated\n")
    latest = (config.output_dir / "runtime/historico/latest.json").read_bytes()
    with pytest.raises(ValueError, match="documento mudou"):
        run(config, reviews_only=True)
    assert (config.output_dir / "runtime/historico/latest.json").read_bytes() == latest
    _, summary, _ = run(config)
    assert summary["human_reviews_invalid"] == 1
    assert summary["approved_balance"] == "0.00"
    _, summary, _ = run(config)
    assert summary["human_reviews_invalid"] == 1


def test_highlight_contains_both_references_and_manifest_checks_nested_files(config):
    directory = setup_case(config)
    evidence = json.loads((directory / "evidencias.json").read_text(encoding="utf-8"))[0]
    assert evidence["highlight"]["marks"] == 2
    with pdfplumber.open(directory / evidence["highlight"]["path"]) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        assert "457.179,12" in text and "456.099,12" in text
    manifest = json.loads((directory / "manifesto.json").read_text(encoding="utf-8"))
    assert all(file_hash(directory / name) == sha for name, sha in manifest["artifacts"].items())


def test_workbook_changed_during_run_is_not_overwritten(config, monkeypatch):
    setup_case(config)
    from goener.revisoes import write_highlights

    def concurrent_edit(*args):
        edit_review(config)
        write_highlights(*args)

    monkeypatch.setattr("goener.pipeline.write_highlights", concurrent_edit)
    with pytest.raises(RuntimeError, match="planilha operacional mudou"):
        run(config)
    wb = load_workbook(config.output_dir / "consolidado.xlsx")
    assert wb["Revisões"]["F5"].value == "Corrigir e liberar"
    wb.close()


def test_repeated_movements_are_both_highlighted_with_short_date(config):
    name = "EXT_102026_LagoaAzul.pdf"
    shutil.copyfile(ROOT / "extratos" / name, config.input_dir / name)
    directory, _, _ = run(config)
    evidence = json.loads((directory / "evidencias.json").read_text(encoding="utf-8"))[0]
    assert evidence["highlight"]["marks"] == 3


def test_excel_original_release_label_and_visible_dropdown(config):
    setup_case(config)
    wb = load_workbook(config.output_dir / "consolidado.xlsx")
    ws = wb["Revisões"]
    choice = next(d for d in ws.data_validations.dataValidation if "F5" in d.sqref)
    assert choice.showDropDown is False and choice.showInputMessage
    assert "Liberar saldo original" in choice.formula1
    assert "Zero" in ws["G4"].comment.text
    wb.close()
    edit_review(config, decision="Liberar saldo original", value=None)
    _, summary, code = run(config, reviews_only=True)
    assert code == 0 and summary["approved_balance"] == "457179.12"
