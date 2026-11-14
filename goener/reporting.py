from __future__ import annotations

import html
import math
from collections import Counter
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .core import Result

HEADERS = [
    "Arquivo",
    "Banco",
    "Conta",
    "Período",
    "Saldo final",
    "Formato de origem",
    "Status",
    "Observação",
]


def write_workbook(template: Path, destination: Path, results: list[Result]) -> None:
    wb = load_workbook(template, keep_links=False)
    try:
        ws = wb["Consolidado"]
        if [c.value for c in ws[1]][:8] != HEADERS:
            raise ValueError("Cabeçalho do modelo diverge do contrato esperado")
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for row_number, result in enumerate(results, 2):
            for col_number, value in enumerate(result.row(), 1):
                cell = ws.cell(row_number, col_number, value)
                if isinstance(value, str):
                    # Literal inclusive quando um nome/observação começa com =, +, - ou @.
                    cell.data_type = "s"
                cell.alignment = Alignment(vertical="center", wrap_text=col_number in (1, 8))
                if col_number == 5:
                    cell.number_format = "#,##0.00;[Red](#,##0.00)"
                if col_number == 7:
                    colors = {
                        "ok": ("E7F3EC", "176038"),
                        "revisar": ("FFF2CD", "805A00"),
                        "falhou": ("FCE5E5", "A52222"),
                    }
                    bg, fg = colors[result.status]
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.font = Font(name="Calibri", size=11, bold=True, color=fg)
            lines = max(1, math.ceil(len(result.observation) / 95))
            ws.row_dimensions[row_number].height = max(31, lines * 16 + 8)
        for col, width in zip("ABCDEFGH", (48, 22, 18, 14, 20, 22, 14, 110), strict=True):
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:H{len(results) + 1}"
        ws.row_dimensions[1].height = 24
        for cell in ws[1]:
            cell.font = copy(cell.font)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        wb.save(destination)
    finally:
        wb.close()
    # A gravação só conta como concluída depois de reabrir e conferir o resultado.
    check = load_workbook(destination, read_only=True, data_only=False)
    try:
        ws = check["Consolidado"]
        if ws.max_row != len(results) + 1:
            raise ValueError("Contagem de linhas gravadas não confere com a entrada")
        for row, result in zip(ws.iter_rows(min_row=2), results, strict=True):
            expected = result.row()
            for cell, value in zip(row[:8], expected, strict=True):
                actual = cell.value
                if (actual if actual is not None else "") != (value if value is not None else ""):
                    raise ValueError(f"Valor gravado diverge em {cell.coordinate}")
                if cell.data_type == "f":
                    raise ValueError("Fórmula inesperada no consolidado")
    finally:
        check.close()


def write_html(path: Path, summary: dict, results: list[Result]) -> None:
    def br(value):
        return f"{float(value):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")

    def plain(text):
        import re

        text = re.sub(r"\b(?:ValueError|PermissionError|OSError):\s*", "", text)
        return re.sub(r"R\$ (-?\d+\.\d{2})", lambda m: "R$ " + br(m.group(1)), text)

    labels = {
        "ok": "Sem pendências",
        "revisar": "Confira este arquivo",
        "falhou": "Não foi possível ler",
    }
    actions = {
        "ARQUIVO_DUPLICADO": "Use o arquivo original. Esta cópia permanece registrada, mas não entra na soma.",
        "LANCAMENTO_REPETIDO": "Confira os dois lançamentos com o responsável pela conta antes de considerar este saldo.",
        "FORA_COMPETENCIA": "Solicite o extrato do mês do fechamento.",
        "DIVERGENCIA_SALDO_FINAL": "Peça um extrato atualizado ou confirme o saldo com o banco.",
        "ERRO_EXTRACAO": "Confirme se o documento é um extrato completo e envie uma nova cópia.",
    }
    rows = []
    for r in sorted(results, key=lambda r: (r.status == "ok", r.filename.casefold())):
        s = r.statement
        messages = [plain(i.message) for i in s.issues]
        action = next(
            (actions[i.code] for i in s.issues if i.code in actions),
            "Você pode considerar este arquivo conforme a política do fechamento."
            if r.status == "ok"
            else "Confira os dados no documento de origem.",
        )
        issues = "".join("<li>" + html.escape(m) + "</li>" for m in messages)
        ai_text = (
            "<h3>Análise assistida</h3><p>" + html.escape(r.ai.get("message", "")) + "</p>"
            if r.ai
            else ""
        )
        meta = html.escape(
            " · ".join(
                v for v in (s.bank, "Conta " + s.account if s.account else "", s.period) if v
            )
        )
        amount = "R$ " + br(s.closing) if s.closing is not None else "Saldo não identificado"
        rows.append(
            f'<details class="item" data-status="{r.status}"><summary><span class="pill {r.status}">{labels[r.status]}</span><div><span class="filename">{html.escape(r.filename)}</span><span class="meta">{meta}</span></div><span class="amount">{amount}</span></summary><div class="body"><h3>O que encontramos</h3>{"<ul>" + issues + "</ul>" if issues else "<p>Os campos foram identificados e os saldos conferem com as movimentações lidas.</p>"}<div class="explain"><b>Próximo passo</b><br>{html.escape(action)}</div>{ai_text}</div></details>'
        )
    cfg = summary.get("configuration", {})
    minutes = cfg.get("manual_minutes_per_file", 8)
    review_minutes = cfg.get("review_minutes_per_file", 3)
    pending = sum(r.status != "ok" for r in results)
    manual = len(results) * minutes / 60
    assisted = summary["duration_seconds"] / 3600 + pending * review_minutes / 60
    saved = manual - assisted
    months = [
        "janeiro",
        "fevereiro",
        "março",
        "abril",
        "maio",
        "junho",
        "julho",
        "agosto",
        "setembro",
        "outubro",
        "novembro",
        "dezembro",
    ]
    period = summary["period"]
    values = {
        "PERIOD": months[int(period[5:]) - 1].capitalize() + " de " + period[:4],
        "COUNT": len(results),
        "BALANCE": "R$ " + br(summary["approved_balance"]),
        "INCLUDED": summary.get("included_files", summary["statuses"].get("ok", 0)),
        "OK": summary["statuses"].get("ok", 0),
        "ATTENTION": pending,
        "REVIEW": summary["statuses"].get("revisar", 0),
        "FAILED": summary["statuses"].get("falhou", 0),
        "SAVED": br(saved),
        "MANUAL": br(manual),
        "ASSISTED": br(assisted),
        "MINUTES": minutes,
        "REVIEW_MINUTES": review_minutes,
        "DURATION": br(summary["duration_seconds"]),
        "RUN": summary["run_id"],
        "METHOD": "A análise por IA está desligada."
        if summary.get("ai_mode", "off") == "off"
        else "Modo de análise assistida: " + summary["ai_mode"] + ".",
    }
    document = Path(__file__).with_name("report.html").read_text(encoding="utf-8")
    for key, value in values.items():
        document = document.replace("__" + key + "__", html.escape(str(value)))
    document = document.replace("__ROWS__", "".join(rows))
    path.write_text(document, encoding="utf-8")


def status_counts(results: list[Result]) -> dict:
    return dict(Counter(r.status for r in results))
