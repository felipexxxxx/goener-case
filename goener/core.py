from __future__ import annotations

import re
import tomllib
import unicodedata
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

CENT = Decimal("0.01")


MONEY = r"(?<![\d.,])[-+]?\s*(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}(?:\s*[CD](?!\w))?(?![\d.,])"


def normalized(value: Any) -> str:
    return " ".join(
        "".join(
            c
            for c in unicodedata.normalize("NFKD", str(value or ""))
            if not unicodedata.combining(c)
        )
        .upper()
        .split()
    )


def money(value: Any) -> Decimal:
    if isinstance(value, bool) or value is None:
        raise ValueError("Valor monetário ausente ou inválido")
    if isinstance(value, (Decimal, int, float)):
        result = Decimal(str(value))
    else:
        text = str(value).upper().replace("R$", "").replace("\xa0", "").strip()
        debit = text.endswith("D") or (text.startswith("(") and text.endswith(")"))
        text = re.sub(r"[CD]$", "", text).strip().strip("() ")
        if "," in text:
            if not re.fullmatch(r"[-+]?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}", text):
                raise ValueError(f"Moeda brasileira malformada: {value!r}")
            text = text.replace(".", "").replace(",", ".")
        elif not re.fullmatch(r"[-+]?\d+(?:\.\d{1,2})?", text):
            raise ValueError(f"Moeda inválida: {value!r}")
        try:
            result = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError("Moeda inválida") from exc
        if debit:
            result = -abs(result)
    if not result.is_finite():
        raise ValueError("Moeda não finita")
    return result.quantize(CENT, rounding=ROUND_HALF_UP)


def amounts(text: str) -> list[Decimal]:
    return [money(m.group().replace(" ", "")) for m in re.finditer(MONEY, text)]


@dataclass
class Issue:
    code: str
    message: str
    severity: str = "revisar"


@dataclass
class Transaction:
    date: str
    description: str
    amount: Decimal
    balance: Decimal | None
    reference: str = ""
    location: str = ""


@dataclass
class Statement:
    bank: str = ""
    account: str = ""
    period: str = ""
    opening: Decimal | None = None
    closing: Decimal | None = None
    transactions: list[Transaction] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    evidence: dict = field(default_factory=dict)

    def issue(self, code: str, message: str, severity: str = "revisar") -> None:
        if not any(i.code == code and i.message == message for i in self.issues):
            self.issues.append(Issue(code, message, severity))

    def to_dict(self) -> dict:
        import json

        return json.loads(json.dumps(asdict(self), default=str, ensure_ascii=False))

    @classmethod
    def from_dict(cls, data: dict) -> Statement:
        data = dict(data)
        for key in ("opening", "closing"):
            data[key] = money(data[key]) if data.get(key) is not None else None
        data["issues"] = [Issue(**i) for i in data.get("issues", [])]
        data["transactions"] = [
            Transaction(
                **{
                    **t,
                    "amount": money(t["amount"]),
                    "balance": money(t["balance"]) if t["balance"] is not None else None,
                }
            )
            for t in data.get("transactions", [])
        ]
        return cls(**data)


@dataclass
class Result:
    filename: str
    sha256: str
    format: str
    statement: Statement
    duration_ms: float = 0
    cached: bool = False
    duplicate_of: str | None = None
    ai: dict = field(default_factory=dict)

    @property
    def status(self) -> str:
        levels = {i.severity for i in self.statement.issues}
        return "falhou" if "falhou" in levels else "revisar" if "revisar" in levels else "ok"

    @property
    def observation(self) -> str:
        notes = [i.message for i in self.statement.issues]
        if self.ai:
            notes.append(self.ai.get("message", ""))
        return " | ".join(n for n in notes if n)

    def included(self, mode: str, period: str) -> bool:
        eligible = self.status == "ok" and not self.duplicate_of and self.statement.period == period
        return eligible and (mode != "automatico" or self.ai.get("decision") == "considerar")

    def row(self) -> list:
        s = self.statement
        return [
            self.filename,
            s.bank,
            s.account,
            s.period,
            float(s.closing) if s.closing is not None else None,
            self.format,
            self.status,
            self.observation,
        ]

    def to_dict(self) -> dict:
        return {
            "file": self.filename,
            "sha256": self.sha256,
            "format": self.format,
            "status": self.status,
            "duration_ms": self.duration_ms,
            "cached": self.cached,
            "duplicate_of": self.duplicate_of,
            "ai": self.ai,
            "statement": self.statement.to_dict(),
        }


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


@dataclass
class Config:
    input_dir: Path
    output_dir: Path
    runtime_dir: Path
    template: Path
    stop_file: Path
    period: str
    ocr_enabled: bool = True
    ocr_min_confidence: float = 0.90
    cache_enabled: bool = True
    strict: bool = False
    max_file_mb: int = 25
    max_pdf_pages: int = 20
    max_image_pixels: int = 30_000_000
    retries: int = 2
    ai_mode: str = "off"
    ai_provider: str = "openai"
    ai_model: str = "gpt-4.1-mini"
    ai_document_mode: str = "completo"
    ai_thinking_level: str = "high"
    ai_timeout_seconds: int = 60
    ai_max_attempts: int = 3
    ai_batch_timeout_seconds: int = 300
    manual_minutes_per_file: float = 8.0
    review_minutes_per_file: float = 3.0

    def to_dict(self) -> dict:
        return {k: str(v) if isinstance(v, Path) else v for k, v in asdict(self).items()}


def previous_month() -> str:
    today = datetime.now(ZoneInfo("America/Sao_Paulo"))
    return (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")


def load_config(path: Path, **overrides) -> Config:
    path = path.resolve()
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    data.update({k: v for k, v in overrides.items() if v is not None})
    if not data.get("period"):
        data["period"] = previous_month()
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", data["period"]):
        raise ValueError("Competência inválida; use AAAA-MM")
    for key in ("input_dir", "output_dir", "runtime_dir", "template", "stop_file"):
        data[key] = (path.parent / data[key]).resolve()
    result = Config(**data)
    if result.ai_mode not in ("off", "consultivo", "automatico"):
        raise ValueError("Modo de IA inválido")
    if result.ai_provider not in ("openai", "gemini"):
        raise ValueError("Provedor de IA inválido")
    if result.ai_document_mode not in ("dados", "completo"):
        raise ValueError("Modo de documento da IA inválido")
    if result.ai_thinking_level not in ("low", "medium", "high"):
        raise ValueError("Nível de raciocínio da IA inválido")
    if not (
        1 <= result.ai_timeout_seconds <= 120
        and 1 <= result.ai_max_attempts <= 3
        and 1 <= result.ai_batch_timeout_seconds <= 900
    ):
        raise ValueError("Limites de execução da IA inválidos")
    if result.manual_minutes_per_file <= 0 or result.review_minutes_per_file < 0:
        raise ValueError("Premissas de tempo inválidas")
    for p in (result.output_dir, result.runtime_dir):
        if p == result.input_dir or p.is_relative_to(result.input_dir):
            raise ValueError("Saída e runtime devem ficar fora da pasta de entrada")
    if not 0 < result.ocr_min_confidence <= 1 or not 0 <= result.retries <= 5:
        raise ValueError("Limite de confiança ou tentativas inválido")
    if min(result.max_file_mb, result.max_pdf_pages, result.max_image_pixels) <= 0:
        raise ValueError("Limites de arquivo, páginas e pixels devem ser positivos")
    return result
