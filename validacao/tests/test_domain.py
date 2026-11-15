from decimal import Decimal

import pytest

from goener.core import (
    Result,
    Statement,
    Transaction,
    amounts,
    classify_duplicates,
    money,
    validate,
)


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("1.234,56", "1234.56"),
        ("R$ 1.234,56 D", "-1234.56"),
        ("1.234,56 C", "1234.56"),
        ("-0,50", "-0.50"),
        ("(50,00)", "-50.00"),
        (0, "0.00"),
        (-32.4, "-32.40"),
        ("1234.56", "1234.56"),
    ],
)
def test_money(source, expected):
    assert money(source) == Decimal(expected)


@pytest.mark.parametrize(
    "value", ["387.27,55", "1.2.345,00", "nan", float("inf"), True, None, "custo"]
)
def test_invalid_money_is_not_repaired_silently(value):
    with pytest.raises(ValueError):
        money(value)


def test_malformed_amount_is_not_partially_matched():
    assert amounts("R$ 387.27,55") == []


def statement(closing="110.00"):
    return Statement(
        "Banco Teste",
        "001-2",
        "2026-10",
        money(100),
        money(closing),
        [Transaction("2026-10-01", "Recebimento", money(10), money(110))],
    )


def test_reconciliation_does_not_correct_declared_balance():
    s = statement("120.00")
    validate(s, "2026-10")
    assert s.closing == money(120)
    assert "DIVERGENCIA_SALDO_FINAL" in {i.code for i in s.issues}
    assert s.evidence["reconciliation"]["difference"] == "10.00"


def test_zero_and_negative_balances_are_valid():
    for opening, amount, closing in [(10, -10, 0), (0, -10, -10)]:
        s = Statement(
            "Banco Teste",
            "001-2",
            "2026-10",
            money(opening),
            money(closing),
            [Transaction("2026-10-01", "Débito", money(amount), money(closing))],
        )
        validate(s, "2026-10")
        assert not s.issues


def test_running_balances_detect_error_even_if_monthly_sum_closes():
    s = Statement(
        "Banco Teste",
        "001-2",
        "2026-10",
        money(100),
        money(100),
        [
            Transaction("2026-10-01", "A", money(10), money(115)),
            Transaction("2026-10-02", "B", money(-10), money(100)),
        ],
    )
    validate(s, "2026-10")
    assert any(i.code == "DIVERGENCIA_MOVIMENTOS" for i in s.issues)


def test_duplicate_retains_row_and_chooses_original_name():
    a, b = (
        Result("conta (1).pdf", "same", "PDF", statement()),
        Result("conta.pdf", "same", "PDF", statement()),
    )
    classify_duplicates([a, b])
    assert a.duplicate_of == "conta.pdf"
    assert a.status == "revisar" and b.status == "ok"
    assert a.row()[4] == b.row()[4] == 110


def test_same_account_period_with_different_content_requires_review():
    a, b = Result("a.pdf", "a", "PDF", statement()), Result("b.pdf", "b", "PDF", statement("120"))
    classify_duplicates([a, b])
    assert a.status == b.status == "revisar"


def test_cache_serialization_preserves_negative_and_zero():
    s = statement("0")
    s.opening = money(-1)
    assert Statement.from_dict(s.to_dict()).to_dict() == s.to_dict()


@pytest.mark.parametrize(
    "mode,confirmed,expected",
    [("automatico", True, "competencia"), ("automatico", False, None), ("consultivo", True, None)],
)
def test_other_month_is_conclusive_only_with_matching_readable_ai_evidence(
    mode, confirmed, expected
):
    from goener.reporting import needs_attention

    r = Result(
        "outro_mes.png", "hash", "Imagem", Statement("Banco", "1", "2026-09", closing=money(10))
    )
    r.statement.issue("FORA_COMPETENCIA", "Esperado 2026-10")
    r.ai = {
        "mode": mode,
        "decision": "revisar",
        "source_check": {
            "period": "2026-09" if confirmed else "2026-10",
            "readable": confirmed,
            "location": "topo",
            "quote": "Período setembro",
        },
    }
    assert r.automatic_exclusion == expected
    assert needs_attention(r) == (expected is None)
    assert not r.included(mode, "2026-10")


def test_nonstatement_exclusion_needs_explicit_source_evidence():
    from goener.reporting import needs_attention

    r = Result("comunicado.pdf", "hash", "PDF", Statement())
    r.statement.issue("CAMPO_AUSENTE", "Saldo não encontrado", "falhou")
    r.ai = {
        "mode": "automatico",
        "decision": "nao_considerar",
        "source_check": {
            "account": None,
            "period": None,
            "closing": None,
            "readable": True,
            "location": "página 1",
            "quote": "Este comunicado não é um extrato bancário.",
        },
    }
    assert r.automatic_exclusion == "nao_extrato" and not needs_attention(r)
    assert not r.included("automatico", "2026-10")
    r.ai["source_check"]["quote"] = "Texto ilegível"
    assert r.automatic_exclusion is None and needs_attention(r)
