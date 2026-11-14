"""Prepara os três entregáveis do case a partir de uma execução já publicada."""

from __future__ import annotations

import argparse
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
        y -= h + (3 if title else 5) * mm

    paragraph("O que foi automatizado", True)
    paragraph(
        "Um único comando lê a pasta de entrada, extrai banco, conta, competência e saldo final de PDF, Excel e imagem e preenche o modelo fornecido. A execução gera planilha, log, evidências por arquivo e um relatório local filtrável. Não há leitura manual pelo chat nem ajuste manual da base."
    )
    paragraph("Fluxo e controles", True)
    paragraph(
        "Entrada &rarr; hash e limites &rarr; leitura direta ou OCR local &rarr; conciliação das movimentações &rarr; duplicidades &rarr; planilha e log &rarr; reabertura, conferência e publicação. Há checkpoint por conteúdo, bloqueio de execução concorrente, retries de leitura, interruptor STOP e modo estrito. O original permanece intacto."
    )
    paragraph("Resultado do lote fornecido", True)
    counts = summary["statuses"]
    timing = "Reexecução com cache" if summary.get("cache_hits", 0) else "Execução sem cache"
    paragraph(
        f"<b>{summary['files']} arquivos</b>: {counts.get('ok', 0)} com status ok, {counts.get('revisar', 0)} para revisar e {counts.get('falhou', 0)} com falha. {timing}: <b>{summary['duration_seconds']:.2f} segundos</b> nesta máquina. A cópia de Lagoa Azul foi preservada e marcada por SHA-256. Também foram sinalizados um possível lançamento repetido, uma diferença de R$ 1.080,00 no saldo declarado de Sol do Sertão, um extrato de setembro e um comunicado sem dados de extrato."
    )
    paragraph("Premissas e ganho estimado", True)
    cfg = summary.get("configuration", {})
    manual_minutes = cfg.get("manual_minutes_per_file", 8)
    review_minutes = cfg.get("review_minutes_per_file", 3)
    exceptions = counts.get("revisar", 0) + counts.get("falhou", 0)
    manual = summary["files"] * manual_minutes / 60
    assisted = summary["duration_seconds"] / 3600 + exceptions * review_minutes / 60
    saved = manual - assisted
    paragraph(
        f"Hipótese por arquivo: 2 min para abrir/localizar, 3 min para copiar os campos, 2 min para conferir e 1 min para registrar e organizar: {manual_minutes} min no total. Com {summary['files']} arquivos, são <b>{manual:.2f} h mensais</b>. Automação e revisão de {exceptions} pendências ({review_minutes} min por pendência): <b>{assisted:.2f} h</b>. Capacidade estimada liberada: <b>{saved:.2f} h/mês</b>. O tempo do robô é medido; os tempos humanos são premissas editáveis, a validar com a equipe. Não há conversão em dinheiro. Instalação e espera pelo emissor ficam fora da conta."
    )
    paragraph("Limites e próximos passos", True)
    paragraph(
        "Os layouts do case são cobertos; novos layouts exigem validação. Confiança de OCR e conciliação não autenticam o extrato. Valores com ressalva não entram no saldo aprovado e não são corrigidos automaticamente.<br/><b>1. P1:</b> cadastrar contas esperadas para identificar extratos ausentes.<br/><b>2. P2:</b> ampliar leitores e testes com novos layouts autorizados.<br/><b>3. P3:</b> validar a análise assistida com casos revisados pela gestão; Power BI fica para uma etapa posterior."
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


def build(run_dir: Path, destination: Path) -> Path:
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
            if name and not name.startswith("entrega/"):
                z.write(ROOT / name, "goener-case/" + name)
        for p in sorted(run_dir.iterdir()):
            if p.is_file() and not p.name.startswith("~$"):
                z.write(p, "goener-case/resultado/" + p.name)
        z.write(destination / "documentacao_1_pagina.pdf", "goener-case/documentacao_1_pagina.pdf")
    return archive


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--run", type=Path, help="Pasta publicada; padrão: .local/outputs/latest.json"
    )
    parser.add_argument("--destination", type=Path, default=ROOT / "entrega")
    args = parser.parse_args()
    selected = args.run or Path(
        json.loads((ROOT / ".local/outputs/latest.json").read_text(encoding="utf-8"))["directory"]
    )
    print(build(selected, args.destination))
