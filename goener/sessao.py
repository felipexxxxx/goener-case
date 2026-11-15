"""Pastas de fechamento independentes e iniciadores portáteis."""

import json
import os
import uuid
from dataclasses import replace
from datetime import datetime
from pathlib import Path


def select_session(config, requested=None, *, create=False):
    root = config.output_dir
    pointer = config.runtime_dir / "sessao_atual.json"
    if requested:
        folder = Path(requested).resolve()
        project = Path(__file__).resolve().parents[1]
        if (
            not (folder.is_relative_to(root.resolve()) or folder == project / "resultado")
            or folder.is_symlink()
        ):
            raise ValueError("Escolha uma execução dentro da pasta do projeto.")
        if not folder.is_dir():
            raise ValueError("Pasta de execução não encontrada.")
    elif not create and pointer.exists():
        folder = (root / json.loads(pointer.read_text(encoding="utf-8"))["folder"]).resolve()
        if not folder.is_relative_to(root.resolve()) or not folder.is_dir():
            pointer.unlink(missing_ok=True)
            return select_session(config, create=False)
        if (
            not (folder / "historico/latest.json").exists()
            and not (folder / "resumo.json").exists()
        ):
            candidates = [
                p for p in root.iterdir() if p.is_dir() and (p / "historico/latest.json").exists()
            ]
            if not candidates:
                raise ValueError(
                    "Ainda não há fechamento concluído. Execute o processamento completo."
                )
            folder = max(candidates, key=lambda p: (p / "historico/latest.json").stat().st_mtime)

    elif not create and (root / "historico/latest.json").exists():
        # Compatibilidade com saídas anteriores; a migração local preserva o histórico.
        folder = root
    else:
        name = (
            "Fechamento "
            + config.period
            + " -- "
            + datetime.now().strftime("%d-%m-%Y %Hh%Mm%Ss")
            + " -- "
            + uuid.uuid4().hex[:4]
        )
        folder = root / name
        folder.mkdir(parents=True)
    selected = replace(config, output_dir=folder)
    from .reporting import publication_directory

    try:
        summary = json.loads(
            (publication_directory(folder) / "resumo.json").read_text(encoding="utf-8")
        )
    except FileNotFoundError:
        summary = None
    if requested and summary:
        selected = replace(selected, period=summary["period"], ai_mode=summary["ai_mode"])
    if not requested and summary:
        remember_session(selected, root)
    write_launchers(folder)
    return selected


def write_launchers(folder):
    folder = Path(folder)
    for name, flags in (
        ("Abrir painel.cmd", "--panel"),
        ("Aplicar revisoes.cmd", "--update-reviews --panel"),
    ):
        text = r"""@echo off
setlocal
set "GOENER_SESSION=%~dp0"
set "GOENER_PROJECT=%~dp0"
:procurar
if exist "%GOENER_PROJECT%executar.py" goto abrir
for %%I in ("%GOENER_PROJECT%..") do set "GOENER_PARENT=%%~fI\"
if "%GOENER_PARENT%"=="%GOENER_PROJECT%" goto erro
set "GOENER_PROJECT=%GOENER_PARENT%"
goto procurar
:abrir
cd /d "%GOENER_PROJECT%"
call "Executar automacao.cmd" --session "%GOENER_SESSION%." FLAGS
exit /b %ERRORLEVEL%
:erro
echo Mantenha esta pasta dentro do projeto Goener completo.
pause
exit /b 1
""".replace("FLAGS", flags)
        (folder / name).write_text(text, encoding="ascii", newline="\r\n")


def remember_session(config, root):
    if (
        not (config.output_dir / "historico/latest.json").is_file()
        and not (config.output_dir / "resumo.json").is_file()
    ):
        return
    pointer = config.runtime_dir / "sessao_atual.json"
    pointer.parent.mkdir(parents=True, exist_ok=True)
    temporary = pointer.with_suffix(".tmp")
    temporary.write_text(
        json.dumps({"folder": os.path.relpath(config.output_dir, root)}), encoding="utf-8"
    )
    temporary.replace(pointer)
