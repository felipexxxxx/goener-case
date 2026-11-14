from pathlib import Path

import pytest
from openpyxl import Workbook

from goener.core import Statement, load_config, money, validate
from goener.readers import LocalOCR, extract, metadata

ROOT = Path(__file__).resolve().parents[1]
EXPECTED = {
    "cajueiro_solar_extrato_10-2026.xlsx": ("18844-6", "175138.64", 19),
    "extrato_mororo_iii_10-2026.xlsx": ("62091-6", "435596.81", 19),
    "JANDAIRA_extrato_outubro.xlsx": ("43291-5", "276394.02", 19),
    "JATOBA_II_extrato_outubro.xlsx": ("12213-0", "285950.28", 18),
    "pau_ferro_out26.xlsx": ("83812-1", "225934.57", 18),
    "CARNAUBA_I_outubro.pdf": ("69294-8", "298409.32", 21),
    "EXT_CHAPADA_VERDE_102026.pdf": ("72359-9", "397981.91", 21),
    "extrato_barauna_ii_out2026.pdf": ("21884-8", "386502.31", 19),
    "extrato_sol_do_sertao_I_out2026.pdf": ("31827-4", "457179.12", 20),
    "EXT_102026_LagoaAzul.pdf": ("70431-9", "308924.49", 22),
}


@pytest.mark.integration
@pytest.mark.parametrize(("filename", "expected"), EXPECTED.items())
def test_original_documents(filename, expected):
    config = load_config(ROOT / "config.toml")
    s = extract(ROOT / "extratos" / filename, config, LocalOCR())
    assert (s.account, s.closing, len(s.transactions)) == (
        expected[0],
        money(expected[1]),
        expected[2],
    )
    assert s.period == "2026-10"


def test_compact_ocr_account():
    s = Statement()
    metadata("Banco Litoral\nAg0001·CC9913355-2\nPeriodo:01/09/2026a30/09/2026", s)
    assert s.account == "9913355-2"
    assert s.period == "2026-09"


def test_notice_is_not_a_bank_statement(config):
    with pytest.raises(ValueError, match="comunicado"):
        extract(ROOT / "extratos/aviso_agencia_horizonte.pdf", config, LocalOCR())


def test_excel_formula_is_not_evaluated(config):
    p = config.input_dir / "formula.xlsx"
    w = Workbook()
    w.active["A1"] = '=HYPERLINK("https://example.invalid","Executar")'
    w.save(p)
    w.close()
    with pytest.raises(ValueError, match="Fórmula"):
        extract(p, config, LocalOCR())


def test_ocr_disabled_fails_explicitly(config):
    config.ocr_enabled = False
    with pytest.raises(ValueError, match="OCR desligado"):
        extract(ROOT / "extratos/app_ipe_amarelo_1011.png", config, LocalOCR())


@pytest.mark.ocr
def test_images_read_and_reconcile_without_manual_values(config):
    expected = {
        "app_ipe_amarelo_1011.png": ("34930-3", "271350.74", 19),
        "app_riacho_fundo_1011.png": ("87458-3", "387257.55", 19),
        "app_umburana_1011.png": ("90683-2", "497889.23", 20),
        "captura_facheiro_out.png": ("96828-1", "475966.91", 21),
        "print_serra_dourada_II.png": ("9913355-2", "243032.29", 19),
    }
    ocr = LocalOCR()
    for filename, (account, closing, count) in expected.items():
        s = extract(ROOT / "extratos" / filename, config, ocr)
        validate(s, config.period)
        assert s.account == account, filename
        assert s.closing == money(closing), filename
        assert len(s.transactions) == count, filename
        assert not {"DIVERGENCIA_MOVIMENTOS", "CONCILIACAO_NAO_FECHA"} & {
            i.code for i in s.issues
        }, filename
        if filename.startswith("print_serra"):
            assert s.period == "2026-09"
            assert any(i.code == "FORA_COMPETENCIA" for i in s.issues)
        else:
            assert not s.issues, filename
