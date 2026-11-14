from __future__ import annotations

import html
import math
import os
from collections import Counter
from pathlib import Path
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from .core import Result, human_text

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


def needs_attention(result):
    return (
        result.status != "ok"
        or bool(result.ai.get("source_disagreements"))
        or result.ai.get("decision") in ("revisar", "nao_considerar")
    )


def source_root(run_dir, input_dir):
    try:
        return quote(os.path.relpath(input_dir, run_dir).replace("\\", "/"), safe="/:")
    except ValueError:
        return Path(input_dir).resolve().as_uri()


def estimate_effort(results, summary):
    """Referência manual declarada e duração medida; dados de apoio ao BI, fora do HTML."""
    cfg = summary.get("configuration", {})
    minutes = cfg.get("manual_minutes_per_file", 8.0)
    check_minutes = cfg.get("manual_check_minutes_per_file", 2.0)
    manual_seconds = len(results) * (minutes + check_minutes) * 60
    robot_seconds = summary["duration_seconds"]
    return {
        "version": "2",
        "manual_entry_minutes": minutes,
        "manual_check_minutes": check_minutes,
        "manual_sample_minutes": minutes + check_minutes,
        "manual_sample_files": cfg.get("manual_sample_files", 1),
        "manual_sample_basis": cfg.get(
            "manual_sample_basis", "Referência manual informada pelo responsável."
        ),
        "manual_source": "Declaração do responsável; a automação não auditou a cronometragem manual.",
        "files": len(results),
        "projected_manual_seconds": round(manual_seconds, 3),
        "measured_robot_seconds": robot_seconds,
        "projected_difference_seconds": round(manual_seconds - robot_seconds, 3),
        "uses_extraction_cache": bool(summary.get("cache_hits", 0)),
        "uses_ai_cache": bool(summary.get("ai_summary", {}).get("cached_decisions", 0)),
        "scope": "Somente consolidação dos arquivos recebidos. Não estima revisão humana, espera por terceiros, implantação ou recebimento. Extrapolação de uma amostra não é média medida de todos os formatos. Reexecução com cache não representa primeira execução.",
    }


def write_summary(wb, period):
    if "Resumo" in wb.sheetnames:
        del wb["Resumo"]
    ws = wb.create_sheet("Resumo", 1)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "AFCA0B"
    ws.column_dimensions["A"].width = 49
    ws.column_dimensions["B"].width = 27
    ws.column_dimensions["C"].width = 59
    ws.merge_cells("A1:C1")
    ws["A1"] = "GOENER | Resumo do fechamento"
    ws["A1"].font = Font(name="Aptos", size=20, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="003836")
    ws.row_dimensions[1].height = 44
    ws["A3"] = "Mês do fechamento (editável)"
    ws["B3"] = period
    ws["B3"].data_type = "s"
    ws["B3"].fill = PatternFill("solid", fgColor="FFF2CD")
    ws["B3"].font = Font(name="Aptos", size=12, bold=True, color="1565C0")
    ws["C3"] = "Use AAAA-MM. O resumo considera somente esse mês."
    eligible = 'Extratos[Status],"ok",Extratos[Período],$B$3'
    rows = {
        6: (
            "Saldo considerado no fechamento",
            f"=SUMIFS(Extratos[Saldo final],{eligible})",
            "Status ok e competência igual ao mês acima.",
        ),
        7: (
            "Arquivos considerados",
            f"=COUNTIFS({eligible})",
            "Cópias e pendências permanecem registradas no Consolidado.",
        ),
        9: (
            "Arquivos registrados",
            "=COUNTA(Extratos[Arquivo])",
            "Cada documento ocupa uma linha, inclusive cópias e falhas.",
        ),
        10: (
            "Com status ok",
            '=COUNTIF(Extratos[Status],"ok")',
            "Contagem de todos os meses presentes na tabela.",
        ),
        11: (
            "Com status revisar",
            '=COUNTIF(Extratos[Status],"revisar")',
            "Não entram no saldo considerado.",
        ),
        12: (
            "Com status falhou",
            '=COUNTIF(Extratos[Status],"falhou")',
            "Não entram no saldo considerado.",
        ),
        13: (
            "Fora do mês selecionado",
            '=COUNTIFS(Extratos[Período],"<>"&$B$3,Extratos[Período],"<>")',
            "Períodos preenchidos diferentes do mês selecionado.",
        ),
        14: (
            "Registros sem status válido",
            "=B9-SUM(B10:B12)",
            "Deve ser zero. Confira o status se não for.",
        ),
    }
    for row, (label, formula, explanation) in rows.items():
        ws.cell(row, 1, label)
        ws.cell(row, 2, formula)
        ws.cell(row, 3, explanation)
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor="F0F7F6")
            cell.font = Font(name="Aptos", size=11, color="003836", bold=cell.column == 2)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 39
    ws["B6"].number_format = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
    ws["B6"].font = Font(name="Aptos", size=20, bold=True, color="003836")
    notes = [
        "Como editar: altere os dados na tabela da aba Consolidado. As fórmulas desta aba recalculam no Excel.",
        "Para incluir ou excluir registros, use Inserir/Excluir linhas da tabela. Não apague apenas o nome do arquivo.",
        "Ao marcar ok, você assume a liberação do registro. Confira cópias, competência e pendências antes de alterar o status.",
        "As fórmulas não refazem OCR, hash, conciliação nem análise de IA. As observações descrevem a execução original.",
        "Salve uma cópia revisada. Alterações aqui não atualizam o HTML e não são reaproveitadas pelo robô na próxima execução.",
        "Origem: documentos do case na aba Consolidado. Cálculo: SOMASES e CONT.SES sobre a tabela Extratos.",
    ]
    for row, note in enumerate(notes, 17):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell = ws.cell(row, 1, note)
        cell.font = Font(name="Aptos", size=11, color="526762")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[row].height = 32
    ws.freeze_panes = "A5"
    ws.print_area = "A1:C22"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


def write_workbook(
    template: Path, destination: Path, results: list[Result], period: str | None = None
) -> None:
    wb = load_workbook(template, keep_links=False)
    try:
        ws = wb["Consolidado"]
        if [c.value for c in ws[1]][:8] != HEADERS:
            raise ValueError("Cabeçalho do modelo diverge do contrato esperado")
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 80
        ws.sheet_properties.tabColor = "00A19A"
        bottom = Side(style="hair", color="DCE8E8")
        for row_number, result in enumerate(results, 2):
            for col_number, value in enumerate(result.row(), 1):
                cell = ws.cell(row_number, col_number, value)
                if isinstance(value, str):
                    # Literal inclusive quando um nome/observação começa com =, +, - ou @.
                    cell.data_type = "s"
                cell.font = Font(name="Aptos", size=10, color="003836")
                cell.fill = PatternFill(
                    "solid", fgColor="F4F9F8" if row_number % 2 == 0 else "FFFFFF"
                )
                cell.border = Border(bottom=bottom)
                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=col_number in (1, 2, 8),
                    horizontal="center"
                    if col_number in (3, 4, 6, 7)
                    else "right"
                    if col_number == 5
                    else "left",
                )
                if col_number == 1:
                    cell.font = Font(name="Aptos", size=10, bold=True, color="003836")
                if col_number == 8 and needs_attention(result):
                    cell.fill = PatternFill("solid", fgColor="FFF8E8")
                if col_number == 5:
                    cell.number_format = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
                if col_number == 7:
                    colors = {
                        "ok": ("E7F3EC", "176038"),
                        "revisar": ("FFF2CD", "805A00"),
                        "falhou": ("FCE5E5", "A52222"),
                    }
                    bg, fg = colors[result.status]
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.font = Font(name="Calibri", size=11, bold=True, color=fg)
            lines = sum(
                max(1, math.ceil(len(line) / 78)) for line in result.observation.splitlines()
            )
            ws.row_dimensions[row_number].height = min(409, max(34, lines * 14 + 14))
        for col, width in zip("ABCDEFGH", (40, 21, 17, 13, 21, 17, 14, 86), strict=True):
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:H{len(results) + 1}"
        ws.row_dimensions[1].height = 32
        for cell in ws[1]:
            cell.font = Font(name="Aptos", size=11, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="003836")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.print_options.horizontalCentered = True
        ws.print_area = f"A1:H{len(results) + 1}"
        ws.print_title_rows = "1:1"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.tables.clear()
        if results:
            table = Table(displayName="Extratos", ref=f"A1:H{len(results) + 1}")
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(table)
            status_validation = DataValidation(type="list", formula1='"ok,revisar,falhou"')
            status_validation.errorTitle = "Escolha um status válido"
            status_validation.error = "Use ok, revisar ou falhou."
            status_validation.showErrorMessage = True
            status_validation.errorStyle = "stop"
            ws.add_data_validation(status_validation)
            status_validation.add("G2:G1048576")
            for status, bg, fg in [
                ("ok", "E7F3EC", "176038"),
                ("revisar", "FFF2CD", "805A00"),
                ("falhou", "FCE5E5", "A52222"),
            ]:
                ws.conditional_formatting.add(
                    "G2:G1048576",
                    CellIsRule(
                        operator="equal",
                        formula=[f'"{status}"'],
                        fill=PatternFill("solid", fgColor=bg),
                        font=Font(color=fg, bold=True),
                    ),
                )
            write_summary(
                wb, period or next((r.statement.period for r in results if r.statement.period), "")
            )
        wb.calculation = CalcProperties(
            calcId=0, fullCalcOnLoad=True, forceFullCalc=True, calcMode="auto"
        )
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
        return human_text(text)

    labels = {
        "ok": "Sem pendências",
        "revisar": "Saldo retido"
        if summary.get("ai_mode") == "automatico"
        else "Confira este arquivo",
        "falhou": "Sem dados válidos",
    }
    actions = {
        "ARQUIVO_DUPLICADO": "Use o arquivo original. Esta cópia permanece registrada, mas não entra na soma.",
        "LANCAMENTO_REPETIDO": "Confira os dois lançamentos com o responsável pela conta antes de considerar este saldo.",
        "FORA_COMPETENCIA": "Solicite o extrato do mês do fechamento.",
        "DIVERGENCIA_SALDO_FINAL": "Peça um extrato atualizado ou confirme o saldo com o banco.",
        "ERRO_EXTRACAO": "Confirme se o documento é um extrato completo e envie uma nova cópia.",
    }
    rows, bank_totals, bank_counts, attention_links = [], {}, Counter(), []
    cfg = summary.get("configuration", {})
    input_dir = Path(cfg.get("input_dir", path.parent))
    source_base = source_root(path.parent, input_dir)
    decisions = {
        "considerar": "IA sugere considerar",
        "revisar": "IA sugere revisão",
        "nao_considerar": "IA sugere não considerar",
    }

    def priority(r):
        codes = {i.code for i in r.statement.issues}
        level = (
            4
            if r.duplicate_of
            else 0
            if codes
            & {"CONCILIACAO_NAO_FECHA", "DIVERGENCIA_SALDO_FINAL", "DIVERGENCIA_MOVIMENTOS"}
            else 1
            if "FORA_COMPETENCIA" in codes
            else 3
            if r.statement.closing is None
            else 2
        )
        return (not needs_attention(r), level, r.filename.casefold())

    for r in sorted(results, key=priority):
        s = r.statement
        messages = [
            plain(i.message)
            for i in s.issues
            if i.code not in {"IA_RETIDO", "IA_DIVERGENCIA_ORIGINAL"}
        ]
        action = (
            actions["ARQUIVO_DUPLICADO"]
            if r.duplicate_of
            else next(
                (actions[i.code] for i in s.issues if i.code in actions),
                "Confira a diferença apontada pela IA no original."
                if r.ai.get("source_disagreements")
                else "Nenhuma providência adicional indicada para este arquivo."
                if r.status == "ok"
                else "Confira os dados no documento de origem.",
            )
        )
        ref = r.review_reference
        href = source_base + "/" + quote(r.filename, safe="")
        if r.format == "PDF" and isinstance(ref.get("page"), int) and ref["page"] > 0:
            href += "#page=" + str(ref["page"])
        decision = r.ai.get("decision")
        if decision in decisions:
            ai_state, ai_label = "analisada", decisions[decision]
        elif decision == "indisponivel":
            ai_state, ai_label = "indisponivel", "IA indisponível"
        elif r.ai.get("routing") == "rules_passed":
            ai_state, ai_label = "dispensada", "Só regras · IA dispensada"
        else:
            ai_state, ai_label = "desligada", "IA desligada nesta execução"
        issues = "".join("<li>" + html.escape(m) + "</li>" for m in messages)
        issues = (
            "<ul>" + issues + "</ul>"
            if issues
            else "<p>Os campos e a conciliação passaram nas verificações programadas.</p>"
        )
        ai_message = plain(r.ai.get("message", "Não foi solicitada análise por IA nesta execução."))
        if r.ai.get("selected_model"):
            ai_message += (
                " Modelo utilizado: "
                + r.ai["selected_model"]
                + (" (alternativo ao principal)." if r.ai.get("failover_used") else ".")
            )
        if r.ai.get("policy_block") and r.ai["policy_block"] not in ai_message:
            ai_message += " Bloqueio obrigatório: " + plain(r.ai["policy_block"])
        if r.ai.get("source_disagreements"):
            ai_message += (
                " O modo consultivo preserva os totais das regras; valide esta diferença antes do fechamento."
                if summary.get("ai_mode") == "consultivo"
                else " A diferença foi registrada e o arquivo permanece fora do total."
            )
        meta = html.escape(
            " · ".join(
                v for v in (s.bank, "Conta " + s.account if s.account else "", s.period) if v
            )
        )
        amount = "R$ " + br(s.closing) if s.closing is not None else "Saldo não identificado"
        finding = (
            f"Cópia idêntica a {r.duplicate_of}. Excluída da soma."
            if r.duplicate_of
            else messages[0]
            if messages
            else (
                "A IA apontou uma diferença na leitura do original."
                if needs_attention(r)
                else "Conferido automaticamente pelas regras."
            )
        )
        quote_block = (
            "<blockquote>" + html.escape(ref["quote"]) + "</blockquote>"
            if ref["quote"]
            else '<p class="caption">Não há trecho localizado disponível. Confira o documento completo.</p>'
        )
        if r.included(summary.get("ai_mode", "off"), summary["period"]):
            bank = s.bank or "Banco não identificado"
            bank_totals[bank] = bank_totals.get(bank, 0) + s.closing
            bank_counts[bank] += 1
        if needs_attention(r):
            attention_links.append(
                f'<button class="attention-link" data-open-file="{html.escape(r.filename, quote=True)}"><span class="attention-dot {r.status}"></span><span><b>{html.escape(s.bank or r.filename)}</b><small>{html.escape(finding)}</small></span><span class="arrow" aria-hidden="true">↗</span></button>'
            )
        rows.append(
            f'<details class="item" data-status="{r.status}" data-attention="{str(needs_attention(r)).lower()}" data-ai="{ai_state}" data-file="{html.escape(r.filename, quote=True)}">'
            f'<summary><span class="pill {r.status}">{labels[r.status]}</span><div><span class="filename">{html.escape(r.filename)}</span>'
            f'<span class="meta">{meta}</span><span class="finding">{html.escape(finding)}</span><span class="ai-badge {ai_state}">{html.escape(ai_label)}</span></div>'
            f'<span class="amount">{amount}</span></summary><div class="body"><div class="reviewgrid"><section><h3>1. O que as regras encontraram</h3>{issues}</section>'
            f"<section><h3>2. Recomendação da IA</h3><p>{html.escape(ai_message)}</p></section></div>"
            f'<div class="reference"><h3>3. Onde conferir</h3><p><b>{html.escape(ref["where"])}</b></p><span class="caption">{html.escape(ref["origin"])}</span>{quote_block}'
            f'<a class="btn secondary" href="{html.escape(href, quote=True)}" target="_blank" rel="noopener">Abrir documento de origem ↗</a></div>'
            f'<div class="explain"><b>4. Medida aplicada automaticamente</b><br>{html.escape(r.applied_action)}'
            f'<p class="caption">Ressalva para acompanhamento: {html.escape(action)}</p></div></div></details>'
        )
    pending = sum(needs_attention(r) for r in results)
    largest_bank = max((abs(value) for value in bank_totals.values()), default=1) or 1
    bank_rows = (
        "".join(
            f'<div class="bank-row"><div><b>{html.escape(bank)}</b><small>{bank_counts[bank]} arquivo(s) incluído(s)</small></div><strong>R$ {br(amount)}</strong><div class="bank-track"><span style="width:{float(abs(amount) / largest_bank) * 100:.3f}%"></span></div></div>'
            for bank, amount in sorted(bank_totals.items(), key=lambda item: item[1], reverse=True)
        )
        or '<p class="muted">Nenhum saldo elegível nesta execução.</p>'
    )
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
        "OK": sum(not needs_attention(r) for r in results),
        "BALANCE_LABEL": "Saldo incluído após as validações"
        if summary.get("ai_mode") == "automatico"
        else "Saldo validado pelas regras",
        "ROUTING": "IA desligada"
        if summary.get("ai_mode", "off") == "off"
        else "IA nas exceções"
        if cfg.get("ai_scope") == "excecoes"
        else "IA em todo o lote",
        "ATTENTION": pending,
        "REVIEW": pending - summary["statuses"].get("falhou", 0),
        "FAILED": summary["statuses"].get("falhou", 0),
        "RECEIVED_PERCENT": round(summary["statuses"].get("ok", 0) / max(len(results), 1) * 100, 1),
        "STATUS_LABEL": "Gerado com ressalvas" if pending else "Sem ressalvas identificadas",
        "CACHE_LABEL": "Reexecução com cache"
        if summary.get("cache_hits", 0) or summary.get("ai_summary", {}).get("cached_decisions", 0)
        else "Execução sem cache",
        "MODE_LABEL": "Automático"
        if summary.get("ai_mode") == "automatico"
        else "Consultivo"
        if summary.get("ai_mode") == "consultivo"
        else "Regras locais",
        "DURATION": br(summary["duration_seconds"]),
        "RUN": summary["run_id"],
        "AI_ANALYZED": sum(r.ai.get("decision") in decisions for r in results),
        "AI_UNAVAILABLE": sum(r.ai.get("decision") == "indisponivel" for r in results),
        "AI_EXEMPT": sum(r.ai.get("routing") == "rules_passed" for r in results),
        "SCOPE": "somente exceções" if cfg.get("ai_scope") == "excecoes" else "todos os arquivos",
        "METHOD": "A análise por IA está desligada."
        if summary.get("ai_mode", "off") == "off"
        else "Consultivo: a IA orienta a revisão; os totais seguem as regras."
        if summary["ai_mode"] == "consultivo"
        else "Automático: a decisão é aplicada no escopo configurado, respeitando as regras obrigatórias.",
    }
    document = Path(__file__).with_name("report.html").read_text(encoding="utf-8")
    for key, value in values.items():
        document = document.replace("__" + key + "__", html.escape(str(value)))
    document = document.replace("__ROWS__", "".join(rows))
    document = document.replace("__BANKS__", bank_rows).replace(
        "__ATTENTION_LINKS__",
        "".join(attention_links) or '<p class="muted">Não há ressalvas nesta execução.</p>',
    )
    path.write_text(document, encoding="utf-8")


def status_counts(results: list[Result]) -> dict:
    return dict(Counter(r.status for r in results))
