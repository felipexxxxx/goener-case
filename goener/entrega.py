"""Prepara os três entregáveis do case a partir de uma execução já publicada."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

from .core import load_config
from .reporting import latest_pointer, source_root

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
        "Um executável Windows sem instalação de Python, ou um comando, lê a pasta de entrada, extrai banco, conta, competência e saldo final de PDF, Excel e imagem e preenche o modelo fornecido. A execução gera planilha, log, evidências por arquivo e relatório filtrável. Aplica retenções e exclusões nos totais, preserva uma linha por arquivo e entrega as ressalvas sem aguardar intervenção humana. Os extratos originais são preservados. O painel local permite registrar revisões e atualizar Excel e HTML juntos; cópias destacadas facilitam localizar divergências."
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
        "* Regulares dispensam IA. O Python compara saldo inicial + movimentos com o final declarado e confere os saldos entre lançamentos. Hash identifica cópias; competência, campos e OCR são validados. O Gemini recebe original/contexto e devolve JSON conferido pelo robô. Falha da API gera ressalva, sem impedir a publicação. Cache, retries, cotas, lock e STOP apoiam a operação."
    )
    paragraph("Resultado do lote fornecido", True)
    counts = summary["statuses"]
    timing = "Reexecução com cache" if summary.get("cache_hits", 0) else "Execução sem cache"
    paragraph(
        f"<b>{summary['files']} arquivos</b>: {counts.get('ok', 0)} com status ok, {counts.get('revisar', 0)} para revisar e {counts.get('falhou', 0)} com falha. {timing}: <b>{summary['duration_seconds']:.2f} segundos</b> nesta máquina. A cópia de Lagoa Azul foi preservada e marcada por SHA-256. Também foram sinalizados um possível lançamento repetido, uma diferença de R$ 1.080,00 no saldo declarado de Sol do Sertão, um extrato de setembro e um comunicado sem dados de extrato."
    )
    paragraph("Tempo estimado e qualidade da evidência", True)
    estimate = summary.get("time_estimate", {})
    if estimate.get("version") != "3":
        raise ValueError("Atualize a execução com a versão atual antes de gerar a entrega.")
    counts_time = estimate["workload_counts"]
    reference = estimate["manual_sample_minutes"]
    paragraph(
        f"Base: <b>{counts_time['current_unique_statements']} extratos únicos do mês + {estimate['triage_files']} triagens</b> ({counts_time['duplicates']} cópia, {counts_time['other_period']} outro mês, {counts_time['insufficient_or_nonstatement']} sem dados). Central: <b>{counts_time['current_unique_statements']} × {reference:g} + {estimate['triage_files']} × {estimate['triage_minutes_assumption']:g} = {estimate['projected_manual_seconds'] / 60:g} min/lote</b>; referência de consolidação declarada e triagem hipotética. Ganhos descontam {summary['duration_seconds']:.2f} s do robô e supõem {estimate['monthly_batches_assumption']:g} lote(s)/mês com volume e composição iguais durante 12 meses."
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
        "<b>Faixa de planejamento, não IC de 95%.</b> Consolidação a 80%/100%/120% da referência e triagem a 50%/100%/200%; hipóteses, não quantis medidos. Uma amostra manual não estima desvio-padrão. Para inferência: piloto por PDF/Excel/imagem, média e dispersão por grupo e ponderação pelo volume. Excluem-se revisão posterior de exceções, espera pelo banco e implantação. Cache não representa primeira execução; sem economia em reais."
    )
    paragraph("Limites e próximos passos", True)
    paragraph(
        "Os layouts do case são cobertos; novos layouts exigem validação. Confiança de OCR e conciliação não autenticam o extrato. Valores com ressalva não entram no saldo aprovado e não são corrigidos automaticamente.<br/><b>1. P1:</b> definir origem, banco/conta/competência esperados, prazo e critério de fechamento completo.<br/><b>2. P2:</b> agendar verificação da pasta por mudanças e reconstruir o Excel por competência, sem duplicar linhas; evolução ainda não implementada.<br/><b>3. P3:</b> homologar correções, retomadas, layouts e IA; ampliar a operação após validar os critérios de aprovação."
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


def build_executable() -> Path:
    """Empacota interpretador, bibliotecas e modelos OCR; mantém configurações externas."""
    from .pipeline import git_revision

    workspace = ROOT / ".local/build"
    workspace.mkdir(parents=True, exist_ok=True)
    metadata = workspace / "build_info.json"
    metadata.write_text(json.dumps({"git_commit": git_revision()}), encoding="utf-8")
    args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--noupx",
        "--name",
        "Goener",
        "--distpath",
        str(ROOT / ".local/bin"),
        "--workpath",
        str(workspace / "work"),
        "--specpath",
        str(workspace),
        "--collect-all",
        "rapidocr",
        "--collect-all",
        "onnxruntime",
    ]
    for name in ("openpyxl", "pdfplumber", "rapidocr", "onnxruntime", "Pillow"):
        args += ["--copy-metadata", name]
    for source, destination in [
        (ROOT / "goener/report.html", "goener"),
        (ROOT / "goener/core.py", "goener"),
        (ROOT / "goener/readers.py", "goener"),
        (metadata, "goener"),
    ]:
        args += ["--add-data", f"{source}{os.pathsep}{destination}"]
    args += [str(ROOT / "executar.py")]
    subprocess.run(args, cwd=ROOT, check=True)
    return ROOT / ".local/bin/Goener.exe"


def build(
    run_dir: Path, destination: Path, include_env: bool = False, executable: Path | None = None
) -> Path:
    if include_env and not (ROOT / ".env").is_file():
        raise ValueError("Arquivo .env não encontrado para inclusão no pacote privado")
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
            if name and not name.startswith("entrega/") and Path(name).name != ".env":
                z.write(ROOT / name, "goener-case/" + name)
        if include_env:
            z.write(ROOT / ".env", "goener-case/.env")
        if executable:
            z.write(executable, "goener-case/Goener.exe")
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
    parser.add_argument(
        "--build-exe",
        action="store_true",
        help="Gera executável Windows com PyInstaller e inclui no ZIP",
    )
    parser.add_argument("--exe", type=Path, help="Inclui um executável já construído")
    args = parser.parse_args()
    selected = args.run or Path(
        json.loads(
            latest_pointer(load_config(ROOT / "config.toml").output_dir).read_text(encoding="utf-8")
        )["directory"]
    )
    executable = build_executable() if args.build_exe else args.exe
    print(build(selected, args.destination, include_env=args.include_env, executable=executable))
