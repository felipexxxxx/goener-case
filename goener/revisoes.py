"""Conferência visual e decisões humanas, sem alterar documentos de entrada."""

from __future__ import annotations

import hashlib
import re
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

import pdfplumber
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

from .core import human_text, money, normalized
from .readers import group_words

HEADERS = [
    "Arquivo",
    "SHA-256",
    "Fechamento",
    "Saldo original",
    "Problema identificado",
    "Decisão",
    "Saldo confirmado",
    "Justificativa",
    "Responsável",
    "Data da revisão",
    "Resultado da revisão",
]
DECISIONS = ("Pendente", "Liberar", "Corrigir e liberar", "Excluir")


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_reviews(path: Path):
    if not path.exists():
        return [], None
    # No Windows, não iniciar uma publicação que substituiria um Excel aberto.
    import os

    if os.name == "nt":
        import ctypes
        from ctypes import wintypes

        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.CreateFileW.argtypes = [
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
            ctypes.c_void_p,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.HANDLE,
        ]
        kernel.CreateFileW.restype = wintypes.HANDLE
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        handle = kernel.CreateFileW(str(path.resolve()), 0x80000000, 0, None, 3, 128, None)
        if handle == wintypes.HANDLE(-1).value:
            raise PermissionError(
                "Feche resultados/consolidado.xlsx e execute novamente. As revisões salvas serão preservadas."
            )
        kernel.CloseHandle(handle)
    if path.stat().st_size > 25_000_000:
        raise ValueError("Planilha operacional excede 25 MB.")
    before = digest(path)
    wb = load_workbook(path, data_only=False, keep_links=False, read_only=True)
    try:
        if "Revisões" not in wb.sheetnames:
            return [], before
        ws = wb["Revisões"]
        if ws.max_row > 10000 or ws.max_column > len(HEADERS):
            raise ValueError("Aba Revisões excede o tamanho permitido.")
        if [c.value for c in ws[4]][: len(HEADERS)] != HEADERS:
            raise ValueError("Cabeçalho da aba Revisões foi alterado. Restaure as colunas.")
        rows = []
        for cells in ws.iter_rows(min_row=5, max_col=len(HEADERS)):
            if not any(c.value is not None for c in cells):
                continue
            if any(c.data_type == "f" for c in cells):
                raise ValueError("Use valores preenchidos, sem fórmulas, na aba Revisões.")
            rows.append([c.value for c in cells])
        if digest(path) != before:
            raise ValueError("A planilha mudou durante a leitura. Salve, feche e tente novamente.")
        return rows, before
    finally:
        wb.close()


def apply_reviews(results, rows, period, log):
    """Uma decisão é válida apenas para nome, conteúdo e fechamento exatos."""
    keys = [(row[0], row[1], row[2]) for row in rows]
    for r in results:
        key = (r.filename, r.sha256, period)
        matched = [row for row in rows if tuple(row[:3]) == key]
        old = [row for row in rows if row[0] == r.filename and row[5] not in (None, "", "Pendente")]
        if not matched:
            if old:
                r.human_review = {
                    "state": "invalid",
                    "message": "Documento ou fechamento mudou; confira novamente na aba Revisões.",
                }
            continue
        row = matched[0]
        decision = row[5] or "Pendente"
        if decision == "Pendente":
            if any(tuple(previous[:3]) != key for previous in old):
                r.human_review = {
                    "state": "invalid",
                    "message": "Documento ou fechamento mudou; registre uma nova decisão para esta versão.",
                }
            continue
        try:
            if keys.count(key) != 1:
                raise ValueError(
                    "Há mais de uma revisão para o mesmo documento; mantenha uma linha por versão."
                )
            if decision not in DECISIONS:
                raise ValueError("Escolha uma decisão da lista.")
            if not all(isinstance(v, str) and v.strip() for v in row[7:9]) or not row[9]:
                raise ValueError("Preencha justificativa, responsável e data.")
            day = row[9]
            if not isinstance(day, (date, datetime)):
                try:
                    day = datetime.fromisoformat(str(day))
                except ValueError:
                    day = datetime.strptime(str(day), "%d/%m/%Y")
            value = None
            if decision != "Excluir":
                s = r.statement
                if (
                    r.duplicate_of
                    or s.period != period
                    or not all((s.bank, s.account, s.period))
                    or s.closing is None
                    or r.rules_status == "falhou"
                ):
                    raise ValueError(
                        "Cópia, outro mês ou documento sem dados válidos não pode ser liberado; substitua o original."
                    )
                if decision == "Corrigir e liberar":
                    if row[6] is None or isinstance(row[6], bool):
                        raise ValueError("Informe o saldo confirmado, inclusive se for zero.")
                    value = money(row[6])
                    if not value.is_finite() or abs(value) > money("999999999999.99"):
                        raise ValueError("Saldo confirmado fora do intervalo permitido.")
                elif row[6] is not None:
                    raise ValueError("Para alterar saldo, escolha Corrigir e liberar.")
            r.human_review = {
                "state": "applied",
                "decision": decision,
                "balance": str(value) if value is not None else None,
                "reason": row[7].strip(),
                "reviewer": row[8].strip(),
                "date": day.isoformat(),
                "sha256": r.sha256,
                "period": period,
                "original_balance": str(r.statement.closing),
                "message": "Revisão humana aplicada.",
            }
        except (ValueError, ArithmeticError) as exc:
            r.human_review = {"state": "invalid", "message": "Revisão não aplicada: " + str(exc)}
        log.event("HUMAN_REVIEW", file=r.filename, **r.human_review)


def write_reviews(wb, results, previous, period):
    from .reporting import needs_attention, workbook_note

    if "Revisões" in wb.sheetnames:
        del wb["Revisões"]
    ws = wb.create_sheet("Revisões", 1)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.sheet_properties.tabColor = "AFCA0B"
    ws.merge_cells("A1:K1")
    ws["A1"] = "Revisões do fechamento"
    ws["A1"].font = Font(name="Aptos", size=20, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="003836")
    ws.row_dimensions[1].height = 40
    ws.merge_cells("A2:K2")
    ws["A2"] = (
        "Preencha as células amarelas. Depois salve, feche e reexecute o robô para atualizar Excel e HTML."
    )
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 34
    ws["A3"] = "Abrir relatório com detalhes"
    ws["A3"].hyperlink = "relatorio.html"
    ws["A3"].font = Font(name="Aptos", color="007F78", underline="single")
    for col, label in enumerate(HEADERS, 1):
        c = ws.cell(4, col, label)
        c.font = Font(name="Aptos", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="003836")
    current = {(r.filename, r.sha256, period): r for r in results}
    rows = []
    for row in previous:
        r = current.get(tuple(row[:3]))
        untouched = row[5] in (None, "", "Pendente") and all(
            value is None or value == "" for value in row[6:10]
        )
        if r and not needs_attention(r) and untouched and not r.human_review:
            continue
        rows.append(list(row))
    keys = {tuple(row[:3]) for row in rows}
    for key, r in current.items():
        if key not in keys and (needs_attention(r) or r.human_review):
            rows.append(
                [
                    *key,
                    float(r.statement.closing) if r.statement.closing is not None else None,
                    "",
                    "Pendente",
                    None,
                    "",
                    "",
                    None,
                    "",
                ]
            )
    for i, row in enumerate(rows, 5):
        r = current.get(tuple(row[:3]))
        if r:
            row[3] = float(r.statement.closing) if r.statement.closing is not None else None
            row[4] = workbook_note(r)
            row[10] = r.human_review.get("message", "Sem decisão registrada.")
        else:
            row[10] = "Histórico: versão ausente ou outro fechamento. Não aplicada."
        if isinstance(row[9], str):
            try:
                row[9] = datetime.fromisoformat(row[9])
            except ValueError:
                try:
                    row[9] = datetime.strptime(row[9], "%d/%m/%Y")
                except ValueError:
                    pass
        for j, value in enumerate(row, 1):
            c = ws.cell(i, j, value)
            if isinstance(value, str):
                c.data_type = "s"
            c.fill = PatternFill("solid", fgColor="FFF2CD" if 6 <= j <= 10 else "F0F7F6")
            c.font = Font(name="Aptos", size=11, color="1565C0" if 6 <= j <= 10 else "003836")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(i, 4).number_format = ws.cell(i, 7).number_format = "#,##0.00;[Red]-#,##0.00"
        ws.cell(i, 10).number_format = "dd/mm/yyyy"
        needed = max(
            len(str(row[0] or "")) / 33,
            len(str(row[4] or "")) / 42,
            len(str(row[7] or "")) / 28,
            len(str(row[10] or "")) / 30,
        )
        ws.row_dimensions[i].height = min(240, max(48, int(needed + 1) * 15))
    for col, width in zip("ABCDEFGHIJK", (33, 16, 14, 18, 42, 23, 19, 28, 21, 17, 30), strict=True):
        ws.column_dimensions[col].width = width
    ws.column_dimensions["B"].hidden = True
    ws.column_dimensions["C"].hidden = True
    ws.column_dimensions["K"].hidden = True
    ws.freeze_panes = "F5"
    ws.auto_filter.ref = f"A4:K{max(5, len(rows) + 4)}"
    dv = DataValidation(type="list", formula1='"' + ",".join(DECISIONS) + '"', allow_blank=True)
    dv.showErrorMessage = True
    dv.error = "Escolha uma opção da lista."
    dv.errorStyle = "stop"
    ws.add_data_validation(dv)
    dv.add(f"F5:F{max(5, len(rows) + 4)}")


def _targets(r):
    """Referências determinísticas; nunca usa coordenadas inventadas pela IA."""
    s = r.statement
    targets = []
    closing = s.evidence.get("closing", {})
    if closing.get("text"):
        targets.append((closing["text"], "Saldo final declarado no documento."))
    codes = {i.code for i in s.issues}
    if codes & {"DIVERGENCIA_SALDO_FINAL", "CONCILIACAO_NAO_FECHA"} and s.transactions:
        t = s.transactions[-1]
        if t.balance is not None:
            targets.append(
                (
                    f"{t.balance:,.2f}".replace(",", "_").replace(".", ",").replace("_", "."),
                    "Saldo após o último lançamento. Compare com o saldo final declarado; a diferença não identifica sozinha qual valor deve ser corrigido.",
                )
            )
    if "FORA_COMPETENCIA" in codes:
        targets += [
            (line, "Período do documento diferente do fechamento solicitado.")
            for line in s.evidence.get("text", "").splitlines()
            if "PERIODO" in normalized(line)
        ]
    if "LANCAMENTO_REPETIDO" in codes:
        seen = set()
        for t in s.transactions:
            key = (t.date, t.description, t.amount, t.reference)
            if key in seen:
                amount = (
                    f"{abs(t.amount):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
                )
                day = date.fromisoformat(t.date).strftime("%d/%m/%Y")
                matches = [
                    line
                    for line in s.evidence.get("text", "").splitlines()
                    if re.match(re.escape(day[:5]) + r"(?:/\d{4})?\s", line)
                    and amount in line
                    and normalized(t.description) in normalized(line)
                    and (not t.reference or t.reference in line)
                ]
                for line in dict.fromkeys(matches):
                    targets.append(
                        (
                            line,
                            "Possível lançamento repetido. Confira as duas ocorrências; nenhuma movimentação foi apagada.",
                        )
                    )
            seen.add(key)
    return targets


def _annotated_pdf(source, destination, r, ocr):
    """Cópia visual rasterizada, com marcações e quadro lateral sempre visível."""
    from PIL import Image

    pdf = pdfplumber.open(source) if r.format == "PDF" else None
    image = Image.open(source).convert("RGB") if pdf is None else None
    c = canvas.Canvas(str(destination))
    style = ParagraphStyle(
        "notes", fontName="Helvetica", fontSize=10, leading=14, textColor="#003836"
    )
    marked = 0
    try:
        pages = pdf.pages if pdf else [None]
        for number, page in enumerate(pages, 1):
            if page:
                width, height = float(page.width), float(page.height)
                picture = page.to_image(resolution=110).original
                words = page.extract_words(x_tolerance=2)
                native_text = bool(words)
                if not words:
                    words = ocr.words(picture, number)
                    sx, sy = width / picture.width, height / picture.height
                    words = [
                        {
                            **w,
                            "x0": w["x0"] * sx,
                            "x1": w["x1"] * sx,
                            "top": w["top"] * sy,
                            "bottom": w["bottom"] * sy,
                        }
                        for w in words
                        if w.get("score", 0) >= 0.90
                    ]
            else:
                picture = image
                scale = min(1, 760 / image.height)
                width, height = image.width * scale, image.height * scale
                words = [
                    {
                        **w,
                        "x0": w["x0"] * scale,
                        "x1": w["x1"] * scale,
                        "top": w["top"] * scale,
                        "bottom": w["bottom"] * scale,
                    }
                    for w in r.statement.evidence.get("tokens", [])
                    if w.get("score", 0) >= 0.90
                ]
            c.setPageSize((width + 245, max(height, 560)))
            ph = max(height, 560)
            c.drawImage(ImageReader(picture), 0, ph - height, width=width, height=height)
            c.setFillColorRGB(0.94, 0.97, 0.96)
            c.rect(width, 0, 245, ph, fill=1, stroke=0)
            notes = []
            for target, explanation in _targets(r):
                matches = [
                    line
                    for line in group_words(words, number)
                    if normalized(target) in normalized(line.text)
                ]
                closing = r.statement.evidence.get("closing", {})
                if (
                    explanation == "Saldo final declarado no documento."
                    and closing.get("top") is not None
                ):
                    if page is None:
                        matches = [
                            line for line in matches if abs(line.top - closing["top"] * scale) < 5
                        ]
                    elif native_text:
                        matches = [
                            line
                            for line in matches
                            if number == closing.get("page") and abs(line.top - closing["top"]) < 4
                        ]
                for line in matches:
                    marked += 1
                    index = len(notes) + 1
                    left, right = min(w["x0"] for w in line.words), max(w["x1"] for w in line.words)
                    top, bottom = line.top, max(w["bottom"] for w in line.words)
                    c.saveState()
                    c.setFillColorRGB(1, 0.79, 0.12)
                    c.setFillAlpha(0.35)
                    c.rect(
                        left - 2,
                        ph - bottom - 2,
                        right - left + 4,
                        bottom - top + 4,
                        fill=1,
                        stroke=0,
                    )
                    c.restoreState()
                    c.setFillColorRGB(0, 0.22, 0.21)
                    c.setFont("Helvetica-Bold", 9)
                    c.drawString(max(1, left - 12), ph - top - 7, str(index))
                    notes.append(f"<b>{index}. {escape(explanation)}</b><br/>{escape(line.text)}")
            text = [
                "<b>Conferência do documento</b>",
                f"Página {number}. Cópia visual; original preservado.",
            ]
            text += notes or [
                "Sem trecho localizado com segurança nesta página. Nenhuma região foi marcada por aproximação."
            ]
            if r.duplicate_of:
                text.append(
                    "Cópia idêntica a "
                    + escape(r.duplicate_of)
                    + "; não conta novamente nos totais."
                )
            for issue in r.statement.issues:
                if issue.code != "IA_RETIDO":
                    text.append(escape(human_text(issue.message)))
            y = ph - 24
            for paragraph in text:
                p = Paragraph(paragraph, style)
                _, h = p.wrap(213, ph)
                if y - h < 20:
                    c.showPage()
                    c.setPageSize((width + 245, ph))
                    y = ph - 24
                p.drawOn(c, width + 16, y - h)
                y -= h + 14
            c.showPage()
        c.save()
    finally:
        if pdf:
            pdf.close()
        if image:
            image.close()
    return marked


def write_highlights(stage, results, config, ocr, log):
    for r in results:
        if not r.statement.issues:
            continue
        source = config.input_dir / r.filename
        try:
            if (
                source.is_symlink()
                or source.resolve().parent != config.input_dir.resolve()
                or digest(source) != r.sha256
            ):
                raise ValueError("Documento mudou ou caminho inválido; destaque não gerado.")
            folder = stage / "destaques"
            folder.mkdir(exist_ok=True)
            name = re.sub(r"[^a-zA-Z0-9_-]", "_", source.stem)[:65] + "_" + r.sha256[:10]
            if r.format in ("PDF", "Imagem"):
                path = folder / (name + ".pdf")
                marks = _annotated_pdf(source, path, r, ocr)
            else:
                path = folder / (name + ".xlsx")
                wb = load_workbook(source, keep_links=False)
                marks = 0
                try:
                    closing = r.statement.evidence.get("closing", {})
                    if closing.get("sheet") and closing.get("cell"):
                        cell = wb[closing["sheet"]][closing["cell"]]
                        cell.fill = PatternFill("solid", fgColor="FFF2CD")
                        cell.comment = Comment(
                            "Saldo final identificado. " + r.observation, "Goener"
                        )
                        marks = 1
                    wb.save(path)
                finally:
                    wb.close()
            if digest(source) != r.sha256:
                path.unlink(missing_ok=True)
                raise ValueError("Documento mudou durante a geração do destaque.")
            r.highlight = {
                "path": "destaques/" + path.name,
                "marks": marks,
                "message": f"{marks} trecho(s) destacado(s)."
                if marks
                else "Sem trecho localizado com segurança; veja as observações na cópia.",
            }
            log.event("HIGHLIGHT_CREATED", file=r.filename, **r.highlight)
        except Exception as exc:
            r.highlight = {"message": "Destaque indisponível: " + str(exc)}
            log.event("HIGHLIGHT_UNAVAILABLE", file=r.filename, message=str(exc))
