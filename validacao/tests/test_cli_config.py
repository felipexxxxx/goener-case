from pathlib import Path

import pytest

from goener.core import load_config, previous_month
from goener.pipeline import main

ROOT = Path(__file__).resolve().parents[2]


def test_cli_reports_configuration_error(tmp_path, capsys):
    assert main(["--config", str(tmp_path / "missing.toml")]) == 1
    assert "ERRO" in capsys.readouterr().err


def test_cli_forwards_operational_switches(monkeypatch, tmp_path, capsys):
    def fake_run(config, *, reviews_only=False):
        assert not reviews_only
        assert config.period == "2026-09"
        assert not config.cache_enabled and not config.ocr_enabled and config.strict
        return (
            tmp_path,
            {
                "files": 1,
                "statuses": {"falhou": 1},
                "duplicates": 0,
                "cache_hits": 0,
                "duration_seconds": 0.1,
                "exit_code": 2,
            },
            2,
        )

    monkeypatch.setattr("goener.pipeline.run", fake_run)
    assert (
        main(
            [
                "--config",
                str(ROOT / "config.toml"),
                "--period",
                "2026-09",
                "--no-cache",
                "--no-ocr",
                "--strict",
            ]
        )
        == 2
    )
    assert '"exit_code": 2' in capsys.readouterr().out


def test_period_and_path_validation(tmp_path):
    config = ROOT / "config.toml"
    with pytest.raises(ValueError, match="Competência"):
        load_config(config, period="2026-13")
    with pytest.raises(ValueError, match="fora da pasta"):
        load_config(config, output_dir="extratos/output")
    with pytest.raises(ValueError, match="confiança"):
        load_config(config, ocr_min_confidence=1.5)
    with pytest.raises(ValueError, match="positivos"):
        load_config(config, max_pdf_pages=0)
    # A resolução dos caminhos usa o config, não o cwd.
    assert load_config(config).input_dir == ROOT / "extratos"


def test_previous_month_uses_sao_paulo():
    assert len(previous_month()) == 7


def test_dotenv_is_loaded_next_to_config_with_priority_over_environment(tmp_path, monkeypatch):
    import os
    import shutil

    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOENER_DOTENV_TEST", raising=False)
    monkeypatch.setenv("GOENER_DOTENV_LITERAL", "existing")
    cfg = tmp_path / "config.toml"
    shutil.copyfile(ROOT / "config.toml", cfg)
    (tmp_path / ".env").write_text(
        "GEMINI_API_KEY=fake-test-key\nGOENER_DOTENV_TEST=${UNEXPANDED}\nGOENER_DOTENV_LITERAL=replaced\n",
        encoding="utf-8",
    )
    load_config(cfg)
    assert os.environ["GEMINI_API_KEY"] == "fake-test-key"
    assert os.environ["GOENER_DOTENV_TEST"] == "${UNEXPANDED}"
    assert os.environ["GOENER_DOTENV_LITERAL"] == "replaced"
    monkeypatch.delenv("GEMINI_API_KEY")
    monkeypatch.delenv("GOENER_DOTENV_TEST")
