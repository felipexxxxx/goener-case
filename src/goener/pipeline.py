from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import os
import platform
import shutil
import sqlite3
import subprocess
import sys
import time
import uuid
from datetime import UTC, datetime
from decimal import Decimal
from pathlib import Path

from . import EXTRACTOR_VERSION, __version__
from .config import Config, load_config
from .models import Result, Statement, classify_duplicates, validate
from .readers import LocalOCR, extract
from .reporting import status_counts, write_html, write_workbook


def atomic_json(path: Path, data: dict | list) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(temporary, path)


def file_hash(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            hasher.update(block)
    return hasher.hexdigest()


class RunLock:
    def __init__(self, path: Path):
        self.path = path

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        try:
            descriptor = os.open(self.path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        except FileExistsError as exc:
            raise RuntimeError(
                f"Execução bloqueada por {self.path}. Confira PID e processo antes de remover um lock abandonado."
            ) from exc
        with os.fdopen(descriptor, "w", encoding="utf-8") as f:
            json.dump({"pid": os.getpid(), "started_at": datetime.now(UTC).isoformat()}, f)
        return self

    def __exit__(self, *_):
        self.path.unlink(missing_ok=True)


class RunLog:
    def __init__(self, root: Path, run_id: str):
        root.mkdir(parents=True, exist_ok=True)
        self.json_path = root / f"{run_id}.jsonl"
        self.text_path = root / f"{run_id}.log"
        self.run_id = run_id

    def event(self, event: str, **fields):
        record = {
            "timestamp": datetime.now(UTC).isoformat(),
            "run_id": self.run_id,
            "event": event,
            **fields,
        }
        with self.json_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
            f.flush()
        detail = " ".join(f"{k}={json.dumps(v, ensure_ascii=False)}" for k, v in fields.items())
        with self.text_path.open("a", encoding="utf-8") as f:
            f.write(f"{record['timestamp']} {event} {detail}\n")


class Cache:
    def __init__(self, path: Path, config: Config):
        source = b"".join(
            Path(__file__).with_name(n).read_bytes() for n in ("readers.py", "models.py")
        )
        options = json.dumps(
            {
                "version": EXTRACTOR_VERSION,
                "ocr": config.ocr_enabled,
                "min_confidence": config.ocr_min_confidence,
                "pages": config.max_pdf_pages,
                "pixels": config.max_image_pixels,
                "mb": config.max_file_mb,
            },
            sort_keys=True,
        ).encode()
        self.fingerprint = hashlib.sha256(source + options).hexdigest()
        self.db = sqlite3.connect(path)
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.execute("PRAGMA synchronous=FULL")
        self.db.execute(
            "CREATE TABLE IF NOT EXISTS extraction (key TEXT PRIMARY KEY, payload TEXT NOT NULL)"
        )

    def key(self, sha256: str, suffix: str) -> str:
        return f"{sha256}:{suffix}:{self.fingerprint}"

    def get(self, key: str) -> Statement | None:
        row = self.db.execute("SELECT payload FROM extraction WHERE key=?", (key,)).fetchone()
        return Statement.from_dict(json.loads(row[0])) if row else None

    def put(self, key: str, statement: Statement) -> None:
        payload = json.dumps(statement.to_dict(), ensure_ascii=False)
        self.db.execute("INSERT OR REPLACE INTO extraction VALUES (?,?)", (key, payload))
        self.db.commit()

    def close(self):
        self.db.close()


def git_revision() -> str:
    try:
        result = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=Path(__file__).resolve().parents[2],
            capture_output=True,
            text=True,
            timeout=5,
            check=True,
        )
        return result.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return "indisponível"


def process_file(
    path: Path, config: Config, ocr: LocalOCR, cache: Cache | None, log: RunLog
) -> Result:
    start = time.perf_counter()
    formats = {
        ".pdf": "PDF",
        ".xlsx": "Excel",
        ".png": "Imagem",
        ".jpg": "Imagem",
        ".jpeg": "Imagem",
    }
    result = Result(path.name, "", formats.get(path.suffix.lower(), "Desconhecido"), Statement())
    log.event("FILE_STARTED", file=path.name)
    for attempt in range(1, config.retries + 2):
        try:
            if path.is_symlink():
                raise ValueError("Links simbólicos não são processados")
            before = path.stat()
            if before.st_size > config.max_file_mb * 1024 * 1024:
                raise ValueError("Arquivo excede o limite de tamanho configurado")
            result.sha256 = file_hash(path)
            key = cache.key(result.sha256, path.suffix.lower()) if cache else ""
            cached = cache.get(key) if cache else None
            result.cached = cached is not None
            result.statement = cached if cached is not None else extract(path, config, ocr)
            after = path.stat()
            if (before.st_size, before.st_mtime_ns) != (
                after.st_size,
                after.st_mtime_ns,
            ) or file_hash(path) != result.sha256:
                raise OSError("Arquivo mudou durante a leitura; aguarde a cópia terminar")
            if cache and not result.cached:
                cache.put(key, result.statement)
            validate(result.statement, config.period)
            break
        except OSError as exc:
            log.event(
                "FILE_ATTEMPT_FAILED", file=path.name, attempt=attempt, error=type(exc).__name__
            )
            if attempt > config.retries:
                result.statement = Statement()
                result.statement.issue("ERRO_LEITURA", str(exc), "falhou")
            else:
                time.sleep(min(0.5 * 2 ** (attempt - 1), 2))
        except Exception as exc:
            # Um documento defeituoso não interrompe o registro dos demais.
            result.statement = Statement()
            result.statement.issue("ERRO_EXTRACAO", f"{type(exc).__name__}: {exc}", "falhou")
            log.event(
                "FILE_EXTRACTION_FAILED", file=path.name, attempt=attempt, error=type(exc).__name__
            )
            break
    result.duration_ms = round((time.perf_counter() - start) * 1000, 2)
    log.event(
        "FILE_EXTRACTED",
        file=path.name,
        sha256=result.sha256,
        cache=result.cached,
        duration_ms=result.duration_ms,
    )
    return result


def run(config: Config) -> tuple[Path, dict, int]:
    if not config.input_dir.is_dir():
        raise ValueError(f"Pasta de entrada inexistente: {config.input_dir}")
    if not config.template.is_file():
        raise ValueError(f"Modelo inexistente: {config.template}")
    files = sorted(
        (p for p in config.input_dir.iterdir() if p.is_file() or p.is_symlink()),
        key=lambda p: p.name.casefold(),
    )
    if not files:
        raise ValueError("Pasta de entrada vazia")
    config.runtime_dir.mkdir(parents=True, exist_ok=True)
    config.output_dir.mkdir(parents=True, exist_ok=True)
    run_id = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ") + "_" + uuid.uuid4().hex[:8]
    log = RunLog(config.runtime_dir / "logs", run_id)
    start = time.perf_counter()
    started_at = datetime.now(UTC).isoformat()
    with RunLock(config.runtime_dir / "execution.lock"):
        log.event("RUN_STARTED", version=__version__, period=config.period, files=len(files))
        cache = (
            Cache(config.runtime_dir / "extractions.sqlite3", config)
            if config.cache_enabled
            else None
        )
        ocr = LocalOCR()
        results: list[Result] = []
        stopped = False
        try:
            for path in files:
                if config.stop_file.exists():
                    stopped = True
                    r = Result(
                        path.name,
                        "",
                        {
                            ".pdf": "PDF",
                            ".xlsx": "Excel",
                            ".png": "Imagem",
                            ".jpg": "Imagem",
                            ".jpeg": "Imagem",
                        }.get(path.suffix.lower(), "Desconhecido"),
                        Statement(),
                    )
                    r.statement.issue(
                        "CANCELADO_STOP", "Não processado: interruptor STOP acionado.", "falhou"
                    )
                    results.append(r)
                    log.event("FILE_CANCELLED", file=path.name)
                    continue
                results.append(process_file(path, config, ocr, cache, log))
            classify_duplicates(results)
            for r in results:
                log.event(
                    "FILE_FINISHED",
                    file=r.filename,
                    status=r.status,
                    codes=[i.code for i in r.statement.issues],
                    duplicate_of=r.duplicate_of,
                )
            statuses = status_counts(results)
            approved = sum(
                (
                    r.statement.closing
                    for r in results
                    if r.status == "ok"
                    and not r.duplicate_of
                    and r.statement.period == config.period
                ),
                Decimal("0.00"),
            )
            exit_code = (
                3
                if stopped
                else 2
                if config.strict and (statuses.get("revisar") or statuses.get("falhou"))
                else 0
            )
            summary = {
                "run_id": run_id,
                "version": __version__,
                "git_commit": git_revision(),
                "started_at": started_at,
                "period": config.period,
                "files": len(files),
                "statuses": statuses,
                "duplicates": sum(r.duplicate_of is not None for r in results),
                "cache_hits": sum(r.cached for r in results),
                "duration_seconds": round(time.perf_counter() - start, 3),
                "approved_balance": str(approved),
                "approved_balance_rule": "Somente status ok, competência esperada e sem cópia de arquivo",
                "stopped": stopped,
                "exit_code": exit_code,
                "python": platform.python_version(),
                "dependencies": {
                    p: importlib.metadata.version(p)
                    for p in ("openpyxl", "pdfplumber", "rapidocr", "onnxruntime", "Pillow")
                },
                "configuration": config.to_dict(),
            }
            stage = config.output_dir / f".{run_id}.pending"
            stage.mkdir()
            write_workbook(config.template, stage / "consolidado.xlsx", results)
            atomic_json(stage / "evidencias.json", [r.to_dict() for r in results])
            summary["duration_seconds"] = round(time.perf_counter() - start, 3)
            atomic_json(stage / "resumo.json", summary)
            write_html(stage / "relatorio.html", summary, results)
            log.event(
                "ARTIFACTS_VALIDATED",
                rows=len(results),
                statuses=json.dumps(statuses, sort_keys=True),
            )
            shutil.copyfile(log.json_path, stage / "execucao.jsonl")
            shutil.copyfile(log.text_path, stage / "execucao.log")
            manifest = {
                "run_id": run_id,
                "algorithm": "SHA-256",
                "inputs": {r.filename: r.sha256 for r in results},
                "template_sha256": file_hash(config.template),
                "artifacts": {p.name: file_hash(p) for p in sorted(stage.iterdir())},
            }
            atomic_json(stage / "manifesto.json", manifest)
            published = config.output_dir / run_id
            os.replace(stage, published)
            atomic_json(
                config.output_dir / "latest.json",
                {"run_id": run_id, "directory": str(published), "exit_code": exit_code},
            )
            log.event(
                "RUN_FINISHED",
                directory=str(published),
                exit_code=exit_code,
                duration_seconds=summary["duration_seconds"],
            )
            return published, summary, exit_code
        except BaseException as exc:
            log.event("RUN_ABORTED", error=type(exc).__name__, message=str(exc))
            raise
        finally:
            if cache:
                cache.close()


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Consolida saldos bancários com rastreabilidade e conciliação local."
    )
    parser.add_argument("--config", type=Path, default=Path("config.toml"))
    parser.add_argument(
        "--period", help="Competência AAAA-MM; sem configuração, mês anterior em São Paulo"
    )
    parser.add_argument("--input", dest="input_dir", help="Pasta de entrada (relativa ao config)")
    parser.add_argument("--output", dest="output_dir", help="Pasta de saída (relativa ao config)")
    parser.add_argument(
        "--no-ocr", action="store_true", help="Desliga OCR; imagens ficam registradas como falha"
    )
    parser.add_argument(
        "--no-cache", action="store_true", help="Reextrai os dados sem usar o checkpoint"
    )
    parser.add_argument(
        "--strict", action="store_true", help="Retorna código 2 quando houver revisar/falhou"
    )
    parser.add_argument("--version", action="version", version=__version__)
    args = parser.parse_args(argv)
    try:
        config = load_config(
            args.config,
            period=args.period,
            input_dir=args.input_dir,
            output_dir=args.output_dir,
            ocr_enabled=False if args.no_ocr else None,
            cache_enabled=False if args.no_cache else None,
            strict=True if args.strict else None,
        )
        directory, summary, code = run(config)
        print(
            json.dumps(
                {
                    "output": str(directory),
                    **{
                        k: summary[k]
                        for k in (
                            "files",
                            "statuses",
                            "duplicates",
                            "cache_hits",
                            "duration_seconds",
                            "exit_code",
                        )
                    },
                },
                ensure_ascii=True,
                indent=2,
            )
        )
        return code
    except (Exception, KeyboardInterrupt) as exc:
        print(f"ERRO: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 130 if isinstance(exc, KeyboardInterrupt) else 1


if __name__ == "__main__":
    raise SystemExit(main())
