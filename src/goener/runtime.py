from __future__ import annotations

import hashlib
import json
import os
import sqlite3
from datetime import UTC, datetime
from pathlib import Path

from . import EXTRACTOR_VERSION
from .config import Config
from .domain import Statement


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
            Path(__file__).with_name(n).read_bytes() for n in ("readers.py", "domain.py")
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
