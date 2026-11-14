from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from .domain import Result, Statement, normalized


def validate(s: Statement, expected_period: str) -> None:
    for key, label in (
        ("bank", "Banco"),
        ("account", "Conta"),
        ("period", "Período"),
        ("closing", "Saldo final"),
    ):
        if getattr(s, key) is None or getattr(s, key) == "":
            s.issue("CAMPO_AUSENTE", f"{label} não encontrado no documento.", "falhou")
    if s.period and s.period != expected_period:
        s.issue(
            "FORA_COMPETENCIA",
            f"Período {s.period}; esperado {expected_period}. Valor retido para revisão.",
        )
    if not s.transactions:
        s.issue("SEM_MOVIMENTACOES", "Movimentações não disponíveis para conciliação.")
        return
    if s.opening is None:
        s.issue(
            "SALDO_INICIAL_AUSENTE",
            "Não foi possível verificar a conciliação desde o início do período.",
        )
    previous = s.opening
    differences = []
    seen = {}
    for tx in s.transactions:
        if previous is not None and tx.balance is not None:
            diff = tx.balance - (previous + tx.amount)
            if diff != Decimal("0.00"):
                differences.append(f"{tx.location}: diferença R$ {diff}")
        previous = tx.balance
        # Igualdade indica suspeita, nunca autorização para apagar um lançamento.
        if tx.date and len(tx.date) == 10:
            key = (tx.date, normalized(tx.description), tx.amount, tx.reference)
            if key in seen:
                s.issue(
                    "LANCAMENTO_REPETIDO",
                    f"Possível lançamento repetido em {tx.date}: R$ {tx.amount}. Conferir documento de origem.",
                )
            seen[key] = tx
    if differences:
        s.issue(
            "DIVERGENCIA_MOVIMENTOS",
            "; ".join(differences[:3])
            + (f"; {len(differences)} divergências no total" if len(differences) > 3 else ""),
        )
    if s.closing is not None and previous is not None and s.closing != previous:
        s.issue(
            "DIVERGENCIA_SALDO_FINAL",
            f"Declarado R$ {s.closing}; último saldo R$ {previous}; diferença R$ {s.closing - previous}. Mantido o valor declarado.",
        )
    if s.opening is not None and s.closing is not None:
        computed = s.opening + sum((t.amount for t in s.transactions), Decimal("0.00"))
        s.evidence["reconciliation"] = {
            "opening": str(s.opening),
            "movements_sum": str(computed - s.opening),
            "computed_closing": str(computed),
            "declared_closing": str(s.closing),
            "difference": str(s.closing - computed),
            "transactions": len(s.transactions),
        }
        if s.closing != computed:
            s.issue(
                "CONCILIACAO_NAO_FECHA",
                f"Saldo inicial + movimentos = R$ {computed}; declarado R$ {s.closing}; diferença R$ {s.closing - computed}.",
            )


def classify_duplicates(results: list[Result]) -> None:
    by_hash: dict[str, list[Result]] = defaultdict(list)
    for result in results:
        if result.sha256:
            by_hash[result.sha256].append(result)
    for group in by_hash.values():
        if len(group) > 1:
            # Prefere o nome sem sufixo de cópia; resultado independe da ordem do filesystem.
            canonical = min(
                group, key=lambda r: ("(" in r.filename, len(r.filename), r.filename.casefold())
            )
            for r in group:
                if r is not canonical:
                    r.duplicate_of = canonical.filename
                    r.statement.issue(
                        "ARQUIVO_DUPLICADO",
                        f"Conteúdo idêntico a {canonical.filename} (SHA-256). Linha preservada; excluir esta cópia de somas.",
                    )
    by_account: dict[tuple, list[Result]] = defaultdict(list)
    for r in results:
        s = r.statement
        if not r.duplicate_of and s.bank and s.account and s.period:
            by_account[(normalized(s.bank), s.account, s.period)].append(r)
    for group in by_account.values():
        if len(group) > 1:
            for r in group:
                r.statement.issue(
                    "CONTA_PERIODO_REPETIDO",
                    "Mais de um conteúdo para banco, conta e período; não somar sem conferência.",
                )
