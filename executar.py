"""Entrada Python acionada pelo iniciador CMD."""

import faulthandler
import multiprocessing
import os
import subprocess
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path


def supervise(command, root: Path, *, desktop=False):
    """O processo externo conserva o diagnóstico mesmo se uma biblioteca nativa cair."""
    print("Iniciando Goener. Preparando o Python e carregando as dependências...", flush=True)
    folder = root / "resultados/controle/diagnosticos"
    try:
        import json

        pointer = root / "resultados/controle/sessao_atual.json"
        selected = (
            root / "resultados" / json.loads(pointer.read_text(encoding="utf-8"))["folder"]
        ).resolve()
        if selected.is_dir() and selected.is_relative_to((root / "resultados").resolve()):
            folder = selected / "runtime/diagnosticos"
    except (OSError, ValueError, KeyError):
        pass
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError:
        folder = Path(tempfile.gettempdir()) / "goener-diagnosticos"
        folder.mkdir(parents=True, exist_ok=True)
    name = datetime.now().strftime("abertura_%Y%m%d_%H%M%S_%f")
    path = folder / (name + ".log")
    native = folder / (name + "_nativo.log")
    session_target = folder / (name + ".session")
    environment = {
        **os.environ,
        "GOENER_NATIVE_LOG": str(native),
        "GOENER_SESSION_TARGET": str(session_target),
        "PYTHONIOENCODING": "utf-8",
        "GOENER_INTERACTIVE": "1" if desktop else "0",
    }
    code = 1
    with path.open("w", encoding="utf-8") as log:
        log.write(f"INICIO {datetime.now().isoformat()}\n")
        log.flush()
        process = None
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
            )
            for line in process.stdout:
                log.write(line)
                log.flush()
                print(line, end="", flush=True)
            code = process.wait()
        except KeyboardInterrupt:
            if process and process.poll() is None:
                process.terminate()
                process.wait()
            code = 130
            log.write("Interrupção solicitada no console.\n")
        except Exception:
            traceback.print_exc(file=log)
            print("Não foi possível iniciar o processamento.", file=sys.stderr)
        log.write(f"FIM {datetime.now().isoformat()} codigo={code} hex={code & 0xFFFFFFFF:#x}\n")
        log.flush()
    try:
        selected = Path(session_target.read_text(encoding="utf-8")).resolve()
        if (
            selected.is_relative_to((root / "resultados").resolve())
            or selected == (root / "resultado").resolve()
        ):
            destination = selected / "runtime/diagnosticos"
            destination.mkdir(parents=True, exist_ok=True)
            if path.parent != destination:
                new_path = destination / path.name
                path.replace(new_path)
                path = new_path
                if native.exists():
                    new_native = destination / native.name
                    native.replace(new_native)
                    native = new_native
    except OSError:
        pass
    finally:
        session_target.unlink(missing_ok=True)
    if code:
        print(
            f"\nA execução não concluiu (código {code}, {code & 0xFFFFFFFF:#x}).\n"
            f"Log da abertura: {path}\nDiagnóstico nativo, se disponível: {native}\n"
            "Os logs por arquivo ficam na pasta runtime do fechamento em resultados. Não houve reinício automático.",
            flush=True,
        )
        if desktop and code != 130 and os.environ.get("GOENER_CMD") != "1":
            try:
                input("Pressione Enter para fechar esta janela...")
            except (EOFError, KeyboardInterrupt):
                pass
    return code


def worker():
    path = os.environ.get("GOENER_NATIVE_LOG")
    native = open(path, "a", encoding="utf-8") if path else None
    try:
        if native:
            faulthandler.enable(file=native, all_threads=True)
        for stream in (sys.stdout, sys.stderr):
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        from goener.pipeline import main

        return main()
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        if native:
            faulthandler.disable()
            native.close()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    if sys.argv[1:2] == ["--_goener_worker"]:
        del sys.argv[1]
        raise SystemExit(worker())
    command = [sys.executable, str(Path(__file__).resolve())]
    command += ["--_goener_worker", *sys.argv[1:]]
    desktop = len(sys.argv) == 1 and sys.stdin is not None and sys.stdin.isatty()
    raise SystemExit(supervise(command, Path.cwd(), desktop=desktop))
