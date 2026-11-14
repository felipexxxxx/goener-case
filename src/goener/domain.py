from __future__ import annotations

import re
import unicodedata
from dataclasses import asdict, dataclass, field
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any

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

    @property
    def status(self) -> str:
        levels = {i.severity for i in self.statement.issues}
        return "falhou" if "falhou" in levels else "revisar" if "revisar" in levels else "ok"

    @property
    def observation(self) -> str:
        return " | ".join(f"[{i.code}] {i.message}" for i in self.statement.issues)

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
            "statement": self.statement.to_dict(),
        }
