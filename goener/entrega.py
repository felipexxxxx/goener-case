"""Prepara os três entregáveis do case a partir de uma execução já publicada."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import zipfile
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

from .core import load_config
from .reporting import publication_directory, source_root

ROOT = Path(__file__).resolve().parents[1]


def one_page(path: Path, summary: dict) -> None:
    width, height = 210 * mm, 297 * mm
    c = canvas.Canvas(str(path), pagesize=(width, height))
    c.setTitle("Case Goener - consolidação de saldos bancários")
    c.setAuthor("Felipe de Paula Lins e Silva")
    green = colors.HexColor("#163F38")
    c.setFillColor(green)
    c.rect(0, height - 34 * mm, width, 34 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 19)
    c.drawString(17 * mm, height - 17 * mm, "Consolidação de saldos bancários")
    c.setFont("Helvetica", 10)
    c.drawString(
        17 * mm,
        height - 25 * mm,
        "Case Goener  |  Felipe de Paula Lins e Silva  |  versão " + summary["version"],
    )
    y = height - 45 * mm
    body_style = ParagraphStyle(
        "body",
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#293D37"),
        alignment=TA_LEFT,
    )
    title_style = ParagraphStyle(
        "title",
        parent=body_style,
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=15,
        textColor=green,
    )

    def paragraph(text, title=False):
        nonlocal y
        p = Paragraph(text, title_style if title else body_style)
        _, h = p.wrap(width - 34 * mm, height)
        p.drawOn(c, 17 * mm, y - h)
        y -= h + (3 if title else 4.5) * mm

    paragraph("O que foi automatizado", True)
    paragraph(
        "Python 3.12 lê PDF, Excel e imagem da pasta de entrada e gera planilha, relatório, logs e evidências, com status por arquivo. Publica mesmo com pendências. O painel registra revisões e atualiza Excel e HTML juntos; cópias destacadas direcionam a conferência."
    )
    paragraph("Fluxo e controles", True)
    gap, box_height = 9, 31
    box_width = (width - 34 * mm - 4 * gap) / 5
    labels = [
        ("Pasta de", "entrada"),
        ("Leitura", "PDF / OCR / Excel"),
        ("Regras de", "conferência"),
        ("IA nas", "exceções*"),
        ("Excel, HTML", "e log"),
    ]
    for i, lines in enumerate(labels):
        x = 17 * mm + i * (box_width + gap)
        c.setFillColor(colors.HexColor("#EAF4F1"))
        c.setStrokeColor(colors.HexColor("#BFD4CF"))
        c.roundRect(x, y - box_height, box_width, box_height, 4, fill=1, stroke=1)
        c.setFillColor(green)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x + box_width / 2, y - 12, lines[0])
        c.drawCentredString(x + box_width / 2, y - 23, lines[1])
        if i < 4:
            right = x + box_width
            c.line(right + 1, y - box_height / 2, right + gap - 1, y - box_height / 2)
            c.line(right + gap - 4, y - box_height / 2 + 2, right + gap - 1, y - box_height / 2)
            c.line(right + gap - 4, y - box_height / 2 - 2, right + gap - 1, y - box_height / 2)
    y -= box_height + 9
    paragraph(
        "* Regulares dispensam IA. O Python compara saldo inicial + movimentos com o final declarado e confere os saldos entre lançamentos. Hash identifica cópias; competência, campos e OCR são validados. O Gemini confere original/contexto e pode concordar ou discordar do alerta. Seu JSON é validado; parecer não equivale a autorização para corrigir saldo. Falha da API gera ressalva, sem impedir a publicação. Há cache, retries, cotas, trava e STOP."
    )
    paragraph("Resultado do lote fornecido", True)
    counts = summary["statuses"]
    benchmark_path = ROOT / "validacao/tempo_pipeline.json"
    benchmark = (
        json.loads(benchmark_path.read_text(encoding="utf-8")) if benchmark_path.exists() else None
    )
    reference = benchmark.get("reference", benchmark) if benchmark else None
    seconds = reference["mean_seconds"] if reference else summary["duration_seconds"]
    scenario = (
        " Com 3.8 pulado por cota, sem tentativas nesse modelo; inclui retry no 3.5."
        if benchmark and "reference" in benchmark
        else ""
    )
    timing = (
        f"Pipeline completo: média de <b>{seconds:.2f} segundos</b> em {reference['n']} execuções sem cache local, incluindo IA HTTP high agrupada e publicação.{scenario}"
        if benchmark
        else "O tempo desta atualização não é usado como benchmark do pipeline completo."
    )
    if benchmark:
        rounds = [
            r
            for r in benchmark["rounds"]
            if "run_ids" not in reference or r["run_id"] in reference["run_ids"]
        ]
        valid = sum(r["ai"]["valid_decisions"] for r in rounds)
        expected = sum(r["ai"]["valid_decisions"] + r["ai"]["unavailable_files"] for r in rounds)
        timing += f" IA: {valid}/{expected} análises válidas; esperas e falhas incluídas."
    paragraph(
        f"<b>{summary['files']} arquivos</b>: {counts.get('ok', 0)} com status ok, {counts.get('revisar', 0)} para revisar e {counts.get('falhou', 0)} com falha. {timing} Ressalvas: cópia idêntica, lançamento suspeito de repetição, diferença de R$ 1.080,00, outro mês e comunicado sem dados."
    )
    paragraph("Tempo estimado e qualidade da evidência", True)
    estimate = summary.get("time_estimate", {})
    if estimate.get("version") != "3":
        raise ValueError("Atualize a execução com a versão atual antes de gerar a entrega.")
    # A documentação desconta a média externa do fluxo completo, não uma atualização rápida.
    estimate = json.loads(json.dumps(estimate))
    batches = estimate.get("monthly_batches_assumption", 1)
    for scenario in estimate["scenarios"].values():
        difference = scenario["manual_seconds_per_batch"] - seconds
        scenario["monthly_hours"] = difference * batches / 3600
        scenario["annual_hours"] = difference * batches * 12 / 3600
    counts_time = estimate["workload_counts"]
    reference = estimate["manual_sample_minutes"]
    paragraph(
        f"Base: <b>{counts_time['current_unique_statements']} extratos únicos do mês + {estimate['triage_files']} triagens</b> ({counts_time['duplicates']} cópia, {counts_time['other_period']} outro mês, {counts_time['insufficient_or_nonstatement']} sem dados). Central: <b>{counts_time['current_unique_statements']} × {reference:g} + {estimate['triage_files']} × {estimate['triage_minutes_assumption']:g} = {estimate['projected_manual_seconds'] / 60:g} min/lote</b>; referência de consolidação declarada e triagem hipotética. Ganhos descontam {seconds:.2f} s do pipeline completo e supõem {estimate['monthly_batches_assumption']:g} lote(s)/mês com volume e composição iguais durante 12 meses."
    )
    c.setFont("Helvetica-Bold", 9)
    for x, text in ((17 * mm, "Cenário"), (105 * mm, "Horas/mês"), (151 * mm, "Horas/ano")):
        c.drawString(x, y - 10, text)
    y -= 19
    for name, label in (("lower", "Menor ganho"), ("central", "Central"), ("upper", "Maior ganho")):
        scenario = estimate["scenarios"][name]
        c.setFont("Helvetica", 9)
        c.drawString(17 * mm, y - 10, label)
        c.drawString(105 * mm, y - 10, f"{scenario['monthly_hours']:.2f}".replace(".", ","))
        c.drawString(151 * mm, y - 10, f"{scenario['annual_hours']:.2f}".replace(".", ","))
        y -= 16
    y -= 5
    paragraph(
        "<b>Faixa de planejamento, não IC de 95%.</b> Consolidação a 80%/100%/120% da referência e triagem a 50%/100%/200%; hipóteses, não quantis. Uma amostra manual não estima dispersão. Para inferência, medir um piloto por formato. Excluem-se revisão posterior, espera pelo banco e implantação; sem economia em reais."
    )
    paragraph("Limites e próximos passos", True)
    paragraph(
        "Validações e OCR não garantem detectar todo erro nem autenticam o extrato. Saldos com ressalva exigem conferência. Evoluções propostas:<br/><b>1. P1 — Definir quais extratos são esperados e quando o fechamento está completo.</b> Combinar bancos, contas, mês e prazo; conta sem extrato válido mantém o fechamento parcial.<br/><b>2. P2 — Agendar o processamento e atualizar o fechamento quando houver mudanças na pasta.</b> Verificar periodicamente e atualizar o mês sem duplicar linhas, preservando histórico e revisões; agendamento ainda não implementado.<br/><b>3. P3 — Criar BI financeiro e operacional mensal.</b> Duas páginas: saldos por banco/conta; tempo, pendências, falhas e revisões do robô. Conferir amostra humana para medir qualidade. BI ainda não implementado."
    )
    if y < 19 * mm:
        raise ValueError("Documento ultrapassou uma página; revisar o layout")
    c.setStrokeColor(colors.HexColor("#CFDCD6"))
    c.line(17 * mm, 17 * mm, width - 17 * mm, 17 * mm)
    c.setFont("Helvetica", 8)
    c.setFillColor(colors.HexColor("#63756C"))
    c.drawString(
        17 * mm,
        12 * mm,
        "Execução: " + summary["run_id"] + "  |  Documentos fictícios do case técnico",
    )
    c.showPage()
    c.save()


def build(run_dir: Path, destination: Path, include_env: bool = False) -> Path:
    if include_env and not (ROOT / ".env").is_file():
        raise ValueError("Arquivo .env não encontrado para inclusão no pacote privado")
    from .sessao import write_launchers

    write_launchers(run_dir)
    summary = json.loads((run_dir / "resumo.json").read_text(encoding="utf-8"))
    destination.mkdir(parents=True, exist_ok=True)
    shutil.copy2(run_dir / "consolidado.xlsx", destination / "consolidado.xlsx")
    one_page(destination / "documentacao_1_pagina.pdf", summary)
    names = (
        subprocess.run(["git", "ls-files", "-z"], cwd=ROOT, capture_output=True, check=True)
        .stdout.decode()
        .split("\0")
    )
    archive = destination / "goener_automacao.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:
            if (
                name
                and not name.startswith("entrega/")
                and Path(name).name != ".env"
                and (ROOT / name).is_file()
            ):
                z.write(ROOT / name, "goener-case/" + name)
        if include_env:
            z.write(ROOT / ".env", "goener-case/.env")
        payloads = {
            p.relative_to(run_dir).as_posix(): p.read_bytes()
            for p in sorted(run_dir.rglob("*"))
            if p.is_file() and not p.name.startswith("~$")
        }
        if "relatorio.html" in payloads:
            input_dir = summary.get("configuration", {}).get("input_dir", ROOT / "extratos")
            prefix = source_root(run_dir, input_dir) + "/"
            payloads["relatorio.html"] = (
                payloads["relatorio.html"]
                .decode("utf-8")
                .replace('href="' + prefix, 'href="../extratos/')
                .encode("utf-8")
            )
            manifest = json.loads(payloads["manifesto.json"])
            manifest["artifacts"]["relatorio.html"] = hashlib.sha256(
                payloads["relatorio.html"]
            ).hexdigest()
            payloads["manifesto.json"] = json.dumps(manifest, ensure_ascii=False, indent=2).encode(
                "utf-8"
            )
        for name, data in payloads.items():
            z.writestr("goener-case/resultado/" + name, data)
        z.write(destination / "documentacao_1_pagina.pdf", "goener-case/documentacao_1_pagina.pdf")
    return archive


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--run", type=Path, help="Pasta publicada; padrão: latest.json na saída configurada"
    )
    parser.add_argument("--destination", type=Path, default=ROOT / "entrega")
    parser.add_argument(
        "--include-env",
        action="store_true",
        help="Inclui a credencial .env no ZIP privado, por solicitação explícita",
    )
    args = parser.parse_args()
    from .sessao import select_session

    selected = args.run or publication_directory(
        select_session(load_config(ROOT / "config.toml")).output_dir
    )
    print(build(selected, args.destination, include_env=args.include_env))
