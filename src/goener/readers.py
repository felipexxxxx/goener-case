from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook
from PIL import Image

from .config import Config
from .domain import MONEY, Statement, Transaction, amounts, money, normalized


@dataclass
class Line:
    text: str
    words: list[dict]
    top: float
    page: int = 1


def group_words(words: list[dict], page: int = 1) -> list[Line]:
    groups: list[list[dict]] = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        height = word["bottom"] - word["top"]
        if groups and abs(word["top"] - groups[-1][0]["top"]) <= max(3, height * 0.45):
            groups[-1].append(word)
        else:
            groups.append([word])
    return [
        Line(
            " ".join(w["text"] for w in sorted(g, key=lambda w: w["x0"])),
            sorted(g, key=lambda w: w["x0"]),
            min(w["top"] for w in g),
            page,
        )
        for g in groups
    ]


def metadata(text: str, statement: Statement) -> None:
    text = normalized(text)
    banks = re.findall(r"BANCO\s+([A-Z]+(?:\s+[A-Z]+)?)", text)
    if banks:
        statement.bank = "Banco " + banks[0].split(" EXTRATO")[0].split(" COMUNICADO")[0].title()
    accounts = re.findall(r"(?:\bCONTA\b|\bCC)\s*[:.]?\s*([\d.]+-\d+)\b", text)
    if accounts:
        statement.account = accounts[0]
        if len(set(accounts)) > 1:
            statement.issue(
                "MULTIPLAS_CONTAS",
                "Mais de uma conta encontrada; separar os extratos antes de processar.",
                "falhou",
            )
    periods = re.findall(r"(\d{2})/(\d{2})/(\d{4})\s*A\s*(\d{2})/(\d{2})/(\d{4})", text)
    if periods:
        d1, m1, y1, d2, m2, y2 = periods[0]
        start, end = date(int(y1), int(m1), int(d1)), date(int(y2), int(m2), int(d2))
        statement.period = start.strftime("%Y-%m")
        if start > end or (y1, m1) != (y2, m2):
            statement.issue(
                "INTERVALO_IRREGULAR", "Intervalo não corresponde a um único mês ordenado."
            )


def iso_date(value, period: str = "") -> str:
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    s = str(value)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return date.fromisoformat(s).isoformat()
    m = re.fullmatch(r"(\d{2})/(\d{2})(?:/(\d{4}))?", s)
    if m:
        day, month, year = m.groups()
        if not year and not period:
            raise ValueError("Ano ausente na data e no período")
        return date(int(year or period[:4]), int(month), int(day)).isoformat()
    raise ValueError(f"Data não reconhecida: {value!r}")


def guard_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        if sum(i.file_size for i in z.infolist()) > 100_000_000:
            raise ValueError("Planilha excede o limite descompactado de 100 MB")
        for n in z.namelist():
            low = n.lower()
            if any(
                k in low
                for k in (
                    "vbaproject",
                    "macrosheet",
                    "externallink",
                    "connections.xml",
                    "embeddings/",
                    "activex/",
                )
            ):
                raise ValueError(
                    "Planilha contém macro, conexão externa ou objeto ativo não suportado"
                )
            if low.endswith((".xml", ".rels")):
                content = z.read(n)
                if b"<!DOCTYPE" in content or b"<!ENTITY" in content:
                    raise ValueError("XML com declaração de entidade não permitido")
                if low.endswith(".rels") and b'TargetMode="External"' in content:
                    raise ValueError("Planilha com relacionamento externo não suportado")


def read_excel(path: Path, config: Config) -> Statement:
    guard_xlsx(path)
    s = Statement(evidence={"engine": "openpyxl", "sheets": []})
    wb = load_workbook(path, read_only=False, data_only=False, keep_links=False)
    try:
        for ws in wb:
            if ws.max_row > 100_000 or ws.max_column > 100:
                raise ValueError("Dimensões da planilha excedem o limite do leitor")
            if ws.sheet_state != "visible":
                s.issue("ABA_OCULTA", f"Aba oculta inspecionada: {ws.title}.")
            for row in ws:
                for cell in row:
                    if cell.data_type == "f":
                        raise ValueError(f"Fórmula não é avaliada: {ws.title}!{cell.coordinate}")
            rows = list(ws.values)
            s.evidence["sheets"].append(ws.title)
            meta = {
                normalized(r[0]): r[1]
                for r in rows
                if len(r) > 1 and r[0] is not None and r[1] is not None
            }
            s.bank = str(meta.get("BANCO", s.bank))
            if rows and normalized(rows[0][0]).startswith("BANCO "):
                s.bank = str(rows[0][0])
            s.account = str(meta.get("CONTA", s.account))
            period_text = str(meta.get("PERIODO", ""))
            metadata(period_text, s)
            if "SALDO_INICIAL" in meta:
                s.opening = money(meta["SALDO_INICIAL"])
            if "SALDO_FINAL" in meta:
                s.closing = money(meta["SALDO_FINAL"])
                s.evidence["closing"] = {
                    "sheet": ws.title,
                    "label": "saldo_final",
                    "value": str(s.closing),
                }
            header_idx = next(
                (
                    i
                    for i, r in enumerate(rows)
                    if "DATA" in [normalized(v) for v in r]
                    and any(normalized(v) in ("DESCRICAO", "HISTORICO") for v in r)
                ),
                None,
            )
            if header_idx is None:
                continue
            headers = {normalized(v): i for i, v in enumerate(rows[header_idx]) if v is not None}

            def cell(row, *names, default=None, headers=headers):
                idx = next((headers[n] for n in names if n in headers), None)
                return row[idx] if idx is not None and idx < len(row) else default

            for number, row in enumerate(rows[header_idx + 1 :], header_idx + 2):
                if all(v is None for v in row):
                    continue
                desc = str(cell(row, "DESCRICAO", "HISTORICO", default=""))
                balance_raw = cell(row, "SALDO", "SALDO (R$)")
                balance = money(balance_raw) if balance_raw is not None else None
                location = f"{ws.title}!linha {number}"
                if normalized(desc) == "SALDO ANTERIOR":
                    s.opening = balance
                    continue
                if normalized(desc) == "SALDO FINAL":
                    s.closing = balance
                    s.evidence["closing"] = {
                        "sheet": ws.title,
                        "row": number,
                        "value": str(balance),
                    }
                    continue
                value = money(cell(row, "VALOR", "VALOR (R$)"))
                kind = normalized(cell(row, "TIPO", default=""))
                if kind:
                    if kind not in ("CREDITO", "DEBITO"):
                        raise ValueError(f"Tipo de movimento desconhecido em {location}")
                    if value < 0:
                        raise ValueError(f"Sinal incompatível com coluna tipo em {location}")
                    value = -value if kind == "DEBITO" else value
                tx_date = iso_date(cell(row, "DATA"), s.period)
                s.transactions.append(
                    Transaction(
                        tx_date,
                        desc,
                        value,
                        balance,
                        str(cell(row, "DOCUMENTO", default="") or ""),
                        location,
                    )
                )
        if not s.period and s.transactions:
            months = {t.date[:7] for t in s.transactions}
            if len(months) == 1:
                s.period = months.pop()
            else:
                s.issue(
                    "MULTIPLOS_PERIODOS",
                    "Lançamentos de mais de um mês sem período explícito.",
                    "falhou",
                )
        return s
    finally:
        wb.close()


def parse_pdf_lines(lines: list[Line]) -> Statement:
    s = Statement(evidence={"engine": "pdfplumber", "text": "\n".join(line.text for line in lines)})
    if "NAO E UM EXTRATO" in normalized(s.evidence["text"]):
        raise ValueError("Documento é um comunicado; não é um extrato bancário")
    metadata(s.evidence["text"], s)
    for line in lines:
        label = normalized(line.text)
        values = amounts(line.text)
        if not values:
            continue
        if "SALDO ANTERIOR" in label:
            s.opening = values[-1]
        elif "SALDO FINAL" in label or (
            "SALDO EM" in label and s.period and f"/{s.period[5:]}/{s.period[:4]}" in label
        ):
            s.closing = values[-1]
            s.evidence["closing"] = {
                "page": line.page,
                "top": line.top,
                "text": line.text,
                "value": str(s.closing),
            }
        elif "SALDO EM" in label:
            s.opening = values[-1]
    debit_x = None
    for line in lines:
        if "CREDITO" in normalized(line.text) and "DEBITO" in normalized(line.text):
            debit_x = next(
                (w["x0"] for w in line.words if normalized(w["text"]).startswith("DEBITO")), None
            )
    for line in lines:
        m = re.match(r"^(\d{2}/\d{2}(?:/\d{4})?)\s+", line.text)
        if (
            not m
            or "SALDO ANTERIOR" in normalized(line.text)
            or "SALDO FINAL" in normalized(line.text)
        ):
            continue
        values = amounts(line.text)
        if len(values) != 2:
            s.issue(
                "LINHA_NAO_LIDA",
                f"Linha financeira incompleta na página {line.page}, y={line.top:.0f}.",
            )
            continue
        value, balance = values
        first_money = next((w for w in line.words if re.search(MONEY, w["text"])), None)
        if debit_x is not None and first_money is not None:
            value = -abs(value) if first_money["x0"] >= debit_x - 12 else abs(value)
        first_match = re.search(MONEY, line.text)
        description = line.text[m.end() : first_match.start()].strip()
        ref = re.search(r"\b[A-Z]{2,4}\d{5,}\b", description)
        reference = ref.group() if ref else ""
        if reference:
            description = description.replace(reference, "").strip()
        s.transactions.append(
            Transaction(
                iso_date(m.group(1), s.period),
                description,
                value,
                balance,
                reference,
                f"página {line.page}, y={line.top:.0f}",
            )
        )
    return s


class LocalOCR:
    def __init__(self):
        self._engine = None

    def engine(self):
        if self._engine is None:
            from rapidocr import ModelType, OCRVersion, RapidOCR

            self._engine = RapidOCR(
                params={
                    "Global.log_level": "warning",
                    "Det.ocr_version": OCRVersion.PPOCRV4,
                    "Det.model_type": ModelType.MOBILE,
                    "Rec.ocr_version": OCRVersion.PPOCRV4,
                    "Rec.model_type": ModelType.MOBILE,
                    "EngineConfig.onnxruntime.intra_op_num_threads": 2,
                    "EngineConfig.onnxruntime.inter_op_num_threads": 2,
                }
            )
        return self._engine

    def words(self, image: Image.Image, page: int = 1) -> list[dict]:
        # A leitura integral mantém o contexto dos cartões. Nada é enviado à rede.
        words: list[dict] = []
        result = self.engine()(image)
        if result.txts is not None:
            for text, score, box in zip(result.txts, result.scores, result.boxes, strict=True):
                x0, x1 = float(box[:, 0].min()), float(box[:, 0].max())
                y0, y1 = float(box[:, 1].min()), float(box[:, 1].max())
                words.append(
                    {
                        "text": str(text),
                        "score": float(score),
                        "x0": x0,
                        "x1": x1,
                        "top": y0,
                        "bottom": y1,
                        "page": page,
                    }
                )
        # Um hífen curto pode sumir na detecção do print inteiro. Releia só essa região
        # em maior resolução antes de considerar o sinal desconhecido.
        for word in words:
            if (
                word["top"] > 400
                and word["x0"] > image.width * 0.5
                and word["text"].strip().startswith("R$")
                and amounts(word["text"])
            ):
                crop = image.crop(
                    (
                        max(0, int(word["x0"]) - 55),
                        max(0, int(word["top"]) - 12),
                        min(image.width, int(word["x1"]) + 20),
                        min(image.height, int(word["bottom"]) + 12),
                    )
                )
                retry = self.engine()(crop.resize((crop.width * 2, crop.height * 2)))
                for retry_text, score in zip(retry.txts or (), retry.scores or (), strict=True):
                    if (
                        re.match(r"^[-+]\s*R\$", retry_text)
                        and amounts(retry_text)
                        and abs(amounts(retry_text)[0]) == abs(amounts(word["text"])[0])
                    ):
                        word["reread"] = {
                            "original": word["text"],
                            "text": retry_text,
                            "score": float(score),
                            "scale": 2,
                        }
                        word["text"] = retry_text
                        word["score"] = min(word["score"], float(score))
                        break
        return words


def parse_image_words(words: list[dict], width: int, min_confidence: float) -> Statement:
    lines = group_words(words)
    s = Statement(evidence={"engine": "rapidocr-ppocrv4-local", "tokens": words})
    text = "\n".join(line.text for line in lines)
    metadata(text, s)
    final_label = next(
        (w for w in words if normalized(w["text"]).replace(" ", "") == "SALDOFINAL"), None
    )
    for w in words:
        compact = normalized(w["text"]).replace(" ", "")
        if "SALDOINICIAL" in compact and amounts(w["text"]):
            s.opening = amounts(w["text"])[-1]
    if final_label:
        candidates = [
            w
            for w in words
            if w["x0"] < width * 0.50
            and final_label["top"] <= w["top"] <= final_label["bottom"] + 90
            and amounts(w["text"])
        ]
        if candidates:
            candidate = min(candidates, key=lambda w: w["top"])
            s.closing = amounts(candidate["text"])[0]
            s.evidence["closing"] = {**candidate, "value": str(s.closing)}
    # Cada cartão tem valor à direita, seguido de seu saldo. Não depende de a data OCR ser perfeita.
    movement_words = [
        w
        for w in words
        if w["top"] > (final_label["bottom"] + 100 if final_label else 350)
        and w["x0"] > width * 0.5
        and re.match(r"^[-+]?\s*R\$", w["text"])
        and amounts(w["text"])
    ]
    balances = [
        w
        for w in words
        if normalized(w["text"]).replace(" ", "").startswith("SALDOR$") and amounts(w["text"])
    ]
    for w in sorted(movement_words, key=lambda x: x["top"]):
        candidates = [b for b in balances if 0 < b["top"] - w["top"] < 85]
        if not candidates:
            s.issue("MOVIMENTO_INCOMPLETO", f"Movimento sem saldo em y={w['top']:.0f}.")
            continue
        b = min(candidates, key=lambda t: t["top"])
        left = [t for t in words if t["x0"] < width * 0.55 and abs(t["top"] - b["top"]) < 20]
        description = " ".join(t["text"] for t in left)
        date_tokens = [t for t in words if t["x0"] < width * 0.45 and abs(t["top"] - w["top"]) < 18]
        date_text = " ".join(t["text"] for t in date_tokens)
        tx_date = date_text
        day_match = re.match(
            r"^(\d{1,2})\s*(?:OUT|SET|JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|NOV|DEZ)\b",
            normalized(date_text),
        )
        if day_match and s.period:
            tx_date = f"{s.period}-{int(day_match.group(1)):02d}"
        value = amounts(w["text"])[0]
        if w["text"].strip().startswith("-"):
            value = -abs(value)
        elif not w["text"].strip().startswith("+"):
            previous = s.transactions[-1].balance if s.transactions else s.opening
            current = amounts(b["text"])[-1]
            if previous is not None and abs(current - previous) == abs(value):
                value = current - previous
                s.issue(
                    "SINAL_OCR_INFERIDO",
                    f"Sinal ilegível em y={w['top']:.0f}; inferido pela variação entre saldos. Conferir origem.",
                )
            else:
                s.issue(
                    "SINAL_OCR_AUSENTE", f"Sinal financeiro não reconhecido em y={w['top']:.0f}."
                )
        s.transactions.append(
            Transaction(
                tx_date,
                description,
                value,
                amounts(b["text"])[-1],
                location=f"imagem y={w['top']:.0f}",
            )
        )
    if s.closing is None and s.transactions:
        s.closing = s.transactions[-1].balance
        s.evidence["closing"] = {
            "source": "último saldo nas movimentações",
            "value": str(s.closing),
        }
        s.issue(
            "SALDO_DESTACADO_ILEGIVEL",
            "Saldo destacado não reconhecido; informado o último saldo legível das movimentações.",
        )
    critical = [
        w
        for w in words
        if (
            s.evidence.get("closing", {}).get("top") == w["top"]
            or "CC" in normalized(w["text"])
            or normalized(w["text"]).startswith("BANCO")
            or "PERIODO" in normalized(w["text"])
        )
    ]
    low = [w for w in critical if w["score"] < min_confidence]
    if low:
        s.issue(
            "OCR_BAIXA_CONFIANCA",
            f"{len(low)} campo(s) de identificação/saldo abaixo de {min_confidence:.0%} de confiança.",
        )
    money_low = [w for w in movement_words + balances if w["score"] < min_confidence]
    if money_low:
        s.issue(
            "OCR_MOVIMENTO_INCERTO",
            f"{len(money_low)} leitura(s) financeira(s) de baixa confiança.",
        )
    s.evidence["critical_confidence_min"] = min((w["score"] for w in critical), default=0)
    s.evidence["text"] = text
    return s


def read_image(path: Path, config: Config, ocr: LocalOCR) -> Statement:
    if not config.ocr_enabled:
        raise ValueError("OCR desligado por configuração")
    with Image.open(path) as im:
        if im.width * im.height > config.max_image_pixels:
            raise ValueError("Imagem excede o limite de pixels")
        im = im.convert("RGB")
        return parse_image_words(ocr.words(im), im.width, config.ocr_min_confidence)


def read_pdf(path: Path, config: Config, ocr: LocalOCR) -> Statement:
    lines: list[Line] = []
    with pdfplumber.open(path) as pdf:
        if len(pdf.pages) > config.max_pdf_pages:
            raise ValueError("PDF excede o limite de páginas")
        for i, page in enumerate(pdf.pages, 1):
            words = page.extract_words(x_tolerance=2)
            if not words:
                if not config.ocr_enabled:
                    raise ValueError("PDF digitalizado exige OCR, que está desligado")
                rendered = page.to_image(resolution=150).original
                if rendered.width * rendered.height > config.max_image_pixels:
                    raise ValueError("Página excede o limite de pixels")
                words = ocr.words(rendered, i)
            lines.extend(group_words(words, i))
    return parse_pdf_lines(lines)


def extract(path: Path, config: Config, ocr: LocalOCR) -> Statement:
    if path.is_symlink():
        raise ValueError("Links simbólicos não são processados")
    if path.stat().st_size > config.max_file_mb * 1024 * 1024:
        raise ValueError("Arquivo excede o limite de tamanho configurado")
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return read_excel(path, config)
    if suffix == ".pdf":
        return read_pdf(path, config, ocr)
    if suffix in (".png", ".jpg", ".jpeg"):
        return read_image(path, config, ocr)
    raise ValueError(f"Formato não suportado: {suffix or 'sem extensão'}")
