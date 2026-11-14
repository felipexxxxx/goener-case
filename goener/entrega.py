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

from .reporting import source_root

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
        "Um executável Windows sem instalação de Python, ou um comando, lê a pasta de entrada, extrai banco, conta, competência e saldo final de PDF, Excel e imagem e preenche o modelo fornecido. A execução gera planilha, log, evidências por arquivo e relatório filtrável. Aplica retenções e exclusões nos totais, preserva uma linha por arquivo e entrega as ressalvas sem aguardar intervenção humana. Os extratos originais são preservados."
    )
    paragraph("Fluxo e controles", True)
    paragraph(
        "Entrada &rarr; hash e limites &rarr; leitura direta ou OCR &rarr; conciliação e competência &rarr; duplicidades &rarr; IA nas exceções &rarr; medidas no consolidado &rarr; validação e publicação. Há cache de extração e análises, controle de cotas, bloqueio de execução concorrente, retries, interruptor STOP e modo estrito. Cópias não entram na soma e inconsistências ficam retidas com ressalva."
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
        "Os layouts do case são cobertos; novos layouts exigem validação. Confiança de OCR e conciliação não autenticam o extrato. Valores com ressalva não entram no saldo aprovado e não são corrigidos automaticamente.<br/><b>1. P1:</b> definir origem, banco/conta/competência esperados, prazo e critério de fechamento completo.<br/><b>2. P2:</b> agendar verificação da pasta por mudanças e reconstruir o Excel por competência, sem duplicar linhas; evolução ainda não implementada.<br/><b>3. P3:</b> homologar correções, retomadas, layouts e IA; avaliar Power BI financeiro e de tempo como evolução futura."
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
            p.name: p.read_bytes()
            for p in sorted(run_dir.iterdir())
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
        "--run", type=Path, help="Pasta publicada; padrão: .local/outputs/latest.json"
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
        json.loads((ROOT / ".local/outputs/latest.json").read_text(encoding="utf-8"))["directory"]
    )
    executable = build_executable() if args.build_exe else args.exe
    print(build(selected, args.destination, include_env=args.include_env, executable=executable))
