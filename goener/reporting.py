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
    if result.human_review.get("state") == "applied":
        return False
    return (
        result.status != "ok"
        or bool(result.ai.get("source_disagreements"))
        or result.ai.get("decision") in ("revisar", "nao_considerar")
    )


def latest_pointer(output_dir):
    current = output_dir / "historico/latest.json"
    return current if current.exists() else output_dir / "latest.json"


def publication_directory(output_dir):
    """Resolve o índice relativo e aceita índices antigos ao mover o projeto."""
    import json
    import re

    output_dir = Path(output_dir).resolve()
    pointer = latest_pointer(output_dir)
    data = json.loads(pointer.read_text(encoding="utf-8"))
    run_id = data["run_id"]
    if not isinstance(run_id, str) or not re.fullmatch(r"[0-9]{8}T[0-9]{6}Z_[0-9a-f]{8}", run_id):
        raise ValueError("Identificador de publicação inválido.")
    raw = Path(data["directory"])
    allowed = (output_dir, output_dir / "historico")
    directory = (raw if raw.is_absolute() else pointer.parent / raw).resolve()
    if directory.parent not in allowed or directory.name != run_id:
        # Índice legado pode ter o usuário/unidade anterior. Usar somente o ID local validado.
        directory = (output_dir / "historico" / run_id).resolve()
        if not directory.is_dir():
            directory = (output_dir / run_id).resolve()
    if directory.parent not in allowed or not directory.is_dir():
        raise ValueError("Publicação não encontrada nesta pasta. Execute a automação.")
    return directory


def source_root(run_dir, input_dir):
    try:
        return quote(os.path.relpath(input_dir, run_dir).replace("\\", "/"), safe="/:")
    except ValueError:
        return Path(input_dir).resolve().as_uri()


def estimate_effort(results, summary):
    """Estimativa por atividade; cenários de planejamento não são intervalos de confiança."""
    cfg = summary.get("configuration", {})
    entry = cfg.get("manual_minutes_per_file", 8.0)
    checking = cfg.get("manual_check_minutes_per_file", 2.0)
    triage = cfg.get("manual_triage_minutes", 1.0)
    batches = cfg.get("monthly_batches", 1.0)
    period = summary.get("period", cfg.get("period", ""))
    counts = {
        "current_unique_statements": 0,
        "duplicates": 0,
        "other_period": 0,
        "insufficient_or_nonstatement": 0,
    }
    formats = Counter()
    seen = set()
    for result in results:
        statement = result.statement
        if result.duplicate_of or (result.sha256 and result.sha256 in seen):
            counts["duplicates"] += 1
            continue
        if result.sha256:
            seen.add(result.sha256)
        if (
            not statement.bank
            or not statement.account
            or not statement.period
            or statement.closing is None
            or result.rules_status == "falhou"
        ):
            counts["insufficient_or_nonstatement"] += 1
        elif statement.period != period:
            counts["other_period"] += 1
        else:
            counts["current_unique_statements"] += 1
            formats[result.format] += 1
    full = counts["current_unique_statements"]
    triaged = len(results) - full
    robot_seconds = summary["duration_seconds"]
    scenarios = {}
    # Faixas deliberadas de sensibilidade, sem probabilidade ou cobertura estatística.
    for name, full_factor, triage_factor in (
        ("lower", 0.8, 0.5),
        ("central", 1, 1),
        ("upper", 1.2, 2),
    ):
        manual_seconds = (
            full * (entry + checking) * full_factor + triaged * triage * triage_factor
        ) * 60
        difference = manual_seconds - robot_seconds
        scenarios[name] = {
            "statement_minutes": round((entry + checking) * full_factor, 3),
            "triage_minutes": round(triage * triage_factor, 3),
            "manual_seconds_per_batch": round(manual_seconds, 3),
            "difference_seconds_per_batch": round(difference, 3),
            "monthly_hours": round(difference * batches / 3600, 6),
            "annual_hours": round(difference * batches * 12 / 3600, 6),
        }
    central = scenarios["central"]
    return {
        "version": "3",
        "manual_entry_minutes": entry,
        "manual_check_minutes": checking,
        "manual_sample_minutes": entry + checking,
        "manual_sample_files": cfg.get("manual_sample_files", 1),
        "manual_sample_basis": cfg.get(
            "manual_sample_basis", "Referência informada pelo responsável."
        ),
        "manual_source": "Referência declarada; sem amostra independente de tempos por formato.",
        "files": len(results),
        "workload_counts": counts,
        "current_statements_by_format": dict(formats),
        "triage_files": triaged,
        "triage_minutes_assumption": triage,
        "monthly_batches_assumption": batches,
        "annual_months_assumption": 12,
        "formula": "N_extratos_unicos_do_mes * t_consolidacao + N_triagem * t_triagem - tempo_do_robo",
        "projected_manual_seconds": central["manual_seconds_per_batch"],
        "measured_robot_seconds": robot_seconds,
        "projected_difference_seconds": central["difference_seconds_per_batch"],
        "scenarios": scenarios,
        "uncertainty": {
            "type": "scenario_sensitivity_not_confidence_interval",
            "confidence_level": None,
            "empirical_confidence_interval": None,
            "reason": "Uma amostra manual não permite estimar o desvio-padrão; não há IC de 95% defensável. Triagem e faixas são hipóteses de planejamento.",
            "sensitivity": "Consolidação a 80%/100%/120% da referência; triagem a 50%/100%/200%. Não são quantis medidos nem limites garantidos.",
            "annualization": "12 meses com o mesmo volume e composição. Incerteza sistemática não diminui por multiplicar por 12.",
            "sampling_plan": "Coletar piloto por PDF, Excel e imagem, com triagem separada; medir média e desvio-padrão por grupo, ponderar pelos volumes e dimensionar amostra pelo erro desejado antes de calcular IC.",
        },
        "uses_extraction_cache": bool(summary.get("cache_hits", 0)),
        "uses_ai_cache": bool(summary.get("ai_summary", {}).get("cached_decisions", 0)),
        "scope": "Ganho potencial da consolidação e triagem inicial; não é economia líquida de todo o fechamento. Exclui revisão posterior de exceções, espera pelo banco, recebimento e implantação. Cache não representa primeira execução.",
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
        "Para sincronizar Excel e HTML: preencha a aba Revisões em resultados/consolidado.xlsx, salve, feche e reexecute o robô.",
        "Para incluir ou excluir registros, use Inserir/Excluir linhas da tabela. Não apague apenas o nome do arquivo.",
        "Ao marcar ok, você assume a liberação do registro. Confira cópias, competência e pendências antes de alterar o status.",
        "As fórmulas não refazem OCR, hash, conciliação nem análise de IA. As observações descrevem a execução original.",
        "Edite as colunas amarelas da aba Revisões. Edições diretas no Consolidado atualizam só este Excel e serão substituídas na reexecução.",
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
    template: Path,
    destination: Path,
    results: list[Result],
    period: str | None = None,
    review_rows: list | None = None,
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
        from .revisoes import write_reviews

        write_reviews(wb, results, review_rows or [], period or "")
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
        amount = (
            "R$ " + br(r.effective_balance)
            if r.effective_balance is not None
            else "Saldo não identificado"
        )
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
        human_block = ""
        if r.human_review:
            h = r.human_review
            content = h["message"]
            if h.get("state") == "applied":
                content += f" {h['decision']} por {h['reviewer']}, em {h['date']}. Motivo: {h['reason']}. Valor original: {h['original_balance']}; saldo utilizado: {r.effective_balance}."
                finding = "Revisão humana registrada: " + h["decision"]
                action = "Conferência concluída pelo responsável; os achados originais continuam registrados."
            human_block = "<p><b>Revisão humana:</b> " + html.escape(content) + "</p>"
        highlighted_link = ""
        if r.highlight.get("path"):
            highlighted_link = (
                '<a class="btn secondary" href="'
                + html.escape(quote(r.highlight["path"], safe="/"), quote=True)
                + '" target="_blank" rel="noopener">Abrir cópia destacada ↗</a><p class="caption">'
                + html.escape(r.highlight["message"])
                + "</p>"
            )
        elif r.highlight:
            highlighted_link = '<p class="caption">' + html.escape(r.highlight["message"]) + "</p>"
        quote_block = (
            "<blockquote>" + html.escape(ref["quote"]) + "</blockquote>"
            if ref["quote"]
            else '<p class="caption">Não há trecho localizado disponível. Confira o documento completo.</p>'
        )
        if r.included(summary.get("ai_mode", "off"), summary["period"]):
            bank = s.bank or "Banco não identificado"
            bank_totals[bank] = bank_totals.get(bank, 0) + r.effective_balance
            bank_counts[bank] += 1
        if needs_attention(r):
            attention_links.append(
                f'<button class="attention-link" data-open-file="{html.escape(r.filename, quote=True)}"><span class="attention-dot {r.status}"></span><span><b>{html.escape(s.bank or r.filename)}</b><small>{html.escape(finding)}</small></span><span class="arrow" aria-hidden="true">↗</span></button>'
            )
        review_button = (
            '<button class="btn review-open" hidden type="button" data-review-file="'
            + html.escape(r.filename, quote=True)
            + '" data-review-sha="'
            + r.sha256
            + '" data-run="'
            + summary["run_id"]
            + '" data-review-decision="'
            + html.escape(r.human_review.get("decision", "Pendente"), quote=True)
            + '" data-review-balance="'
            + html.escape(r.human_review.get("balance") or "", quote=True)
            + '" data-review-reason="'
            + html.escape(r.human_review.get("reason", ""), quote=True)
            + '" data-review-reviewer="'
            + html.escape(r.human_review.get("reviewer", ""), quote=True)
            + '" data-review-original="'
            + html.escape(str(s.closing), quote=True)
            + '">Revisar este arquivo</button>'
        )
        rows.append(
            f'<details class="item" data-status="{r.status}" data-attention="{str(needs_attention(r)).lower()}" data-ai="{ai_state}" data-file="{html.escape(r.filename, quote=True)}">'
            f'<summary><span class="pill {r.status}">{labels[r.status]}</span><div><span class="filename">{html.escape(r.filename)}</span>'
            f'<span class="meta">{meta}</span><span class="finding">{html.escape(finding)}</span><span class="ai-badge {ai_state}">{html.escape(ai_label)}</span></div>'
            f'<span class="amount">{amount}</span></summary><div class="body"><div class="reviewgrid"><section><h3>1. O que as regras encontraram</h3>{issues}</section>'
            f"<section><h3>2. Recomendação da IA</h3><p>{html.escape(ai_message)}</p></section></div>"
            f'<div class="reference"><h3>3. Onde conferir</h3><p><b>{html.escape(ref["where"])}</b></p><span class="caption">{html.escape(ref["origin"])}</span>{quote_block}'
            f'<a class="btn secondary" href="{html.escape(href, quote=True)}" target="_blank" rel="noopener">Abrir documento de origem ↗</a>{highlighted_link}{review_button}</div>'
            f'<div class="explain"><b>4. Medida aplicada automaticamente</b><br>{html.escape(r.applied_action)}{human_block}'
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
        "STATUS_LABEL": "Gerado com ressalvas"
        if pending
        else "Pendências tratadas por revisão"
        if summary.get("human_reviews_applied")
        else "Sem ressalvas identificadas",
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
    from .revisoes import DECISIONS

    document = Path(__file__).with_name("report.html").read_text(encoding="utf-8")
    document = document.replace(
        "__REVIEW_OPTIONS__", "".join("<option>" + html.escape(d) + "</option>" for d in DECISIONS)
    )
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
