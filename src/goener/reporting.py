from __future__ import annotations

import html
import json
import math
from collections import Counter
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .domain import Result

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
    cards = "".join(
        f'<div class="card"><span>{html.escape(label)}</span><strong>{html.escape(str(value))}</strong></div>'
        for label, value in (
            ("Arquivos", summary["files"]),
            ("Aprovados", summary["statuses"].get("ok", 0)),
            ("Para revisar", summary["statuses"].get("revisar", 0)),
            ("Falharam", summary["statuses"].get("falhou", 0)),
        )
    )
    rows = []
    for r in results:
        s = r.statement
        evidence = json.dumps(
            {
                "sha256": r.sha256,
                "closing": s.evidence.get("closing"),
                "reconciliation": s.evidence.get("reconciliation"),
                "issues": [i.__dict__ for i in s.issues],
                "cache": r.cached,
                "duration_ms": r.duration_ms,
            },
            ensure_ascii=False,
            indent=2,
        )
        cells = "".join(
            f"<td>{html.escape(str(v if v is not None else '—'))}</td>" for v in r.row()[:7]
        )
        rows.append(
            f'<tr data-status="{r.status}">{cells}<td>{html.escape(r.observation) or "Conciliação sem divergências."}<details><summary>Ver evidência</summary><pre>{html.escape(evidence)}</pre></details></td></tr>'
        )
    document = f"""<!doctype html><html lang="pt-BR"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Goener | execução {html.escape(summary["run_id"])}</title>
<style>body{{margin:0;background:#f4f6f5;color:#16342e;font:15px system-ui}}main{{max-width:1500px;margin:auto;padding:36px}}header p{{color:#53635e}}h1{{font-size:30px;margin:8px 0}}.cards{{display:flex;gap:14px;flex-wrap:wrap;margin:28px 0}}.card{{background:white;border:1px solid #dbe2de;border-radius:12px;padding:20px;min-width:140px}}.card span{{display:block;color:#65746e}}.card strong{{font-size:32px}}.scroll{{overflow:auto;background:white;border:1px solid #dbe2de;border-radius:12px}}table{{border-collapse:collapse;width:100%;font-size:13px}}th,td{{padding:13px;text-align:left;border-bottom:1px solid #e4e9e6;vertical-align:top}}th{{white-space:nowrap;background:#e9f0ed}}td:first-child{{min-width:230px}}td:last-child{{min-width:290px}}tr[data-status=revisar] td:nth-child(7){{color:#875c00;font-weight:700}}tr[data-status=falhou] td:nth-child(7){{color:#ad2222;font-weight:700}}tr[data-status=ok] td:nth-child(7){{color:#176038;font-weight:700}}pre{{white-space:pre-wrap;overflow-wrap:anywhere;font-size:11px}}details{{margin-top:9px}}select{{padding:9px;margin:10px 0 20px;border-radius:6px}}.note{{padding:15px;background:#fff2cd;border-radius:8px;line-height:1.5}}a{{color:#065d4c}}</style>
<main><header><small>GOENER · CONSOLIDAÇÃO BANCÁRIA</small><h1>Fechamento de {html.escape(summary["period"])}</h1><p>Execução {html.escape(summary["run_id"])} · {summary["duration_seconds"]:.2f}s · versão {html.escape(summary["version"])}</p></header><div class="cards">{cards}</div>
<p class="note">Saldo aprovado: <strong>R$ {summary["approved_balance"]}</strong>. Considera somente linhas <b>ok</b> da competência, sem cópias. Valores com ressalvas permanecem no Excel e na evidência. Nenhum valor foi corrigido para fechar a conciliação.</p>
<label for="status">Mostrar</label> <select id="status"><option value="all">Todos os arquivos</option><option value="ok">Aprovados</option><option value="revisar">Para revisar</option><option value="falhou">Falhas</option></select>
<p><a href="consolidado.xlsx">Abrir planilha</a> · <a href="execucao.log">Log legível</a> · <a href="evidencias.json">Evidências completas</a> · <a href="manifesto.json">Manifesto SHA-256</a></p>
<div class="scroll"><table><thead><tr>{"".join("<th>" + html.escape(x) + "</th>" for x in HEADERS)}</tr></thead><tbody>{"".join(rows)}</tbody></table></div><p>Processamento local. Este painel não consulta serviços externos.</p></main>
<script>document.getElementById('status').addEventListener('change',e=>document.querySelectorAll('tbody tr').forEach(r=>r.hidden=e.target.value!=='all'&&r.dataset.status!==e.target.value));</script></html>"""
    path.write_text(document, encoding="utf-8")


def status_counts(results: list[Result]) -> dict:
    return dict(Counter(r.status for r in results))
