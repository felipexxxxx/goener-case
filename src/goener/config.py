from __future__ import annotations

import re
import tomllib
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo


@dataclass
class Config:
    input_dir: Path
    output_dir: Path
    runtime_dir: Path
    template: Path
    stop_file: Path
    period: str
    ocr_enabled: bool = True
    ocr_min_confidence: float = 0.90
    cache_enabled: bool = True
    strict: bool = False
    max_file_mb: int = 25
    max_pdf_pages: int = 20
    max_image_pixels: int = 30_000_000
    retries: int = 2
    ai_mode: str = "off"
    ai_provider: str = "openai"
    ai_model: str = "gpt-4.1-mini"
    manual_minutes_per_file: float = 8.0
    review_minutes_per_file: float = 3.0

    def to_dict(self) -> dict:
        return {k: str(v) if isinstance(v, Path) else v for k, v in asdict(self).items()}


def previous_month() -> str:
    today = datetime.now(ZoneInfo("America/Sao_Paulo"))
    return (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")


def load_config(path: Path, **overrides) -> Config:
    path = path.resolve()
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    data.update({k: v for k, v in overrides.items() if v is not None})
    if not data.get("period"):
        data["period"] = previous_month()
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", data["period"]):
        raise ValueError("Competência inválida; use AAAA-MM")
    for key in ("input_dir", "output_dir", "runtime_dir", "template", "stop_file"):
        data[key] = (path.parent / data[key]).resolve()
    result = Config(**data)
    if result.ai_mode not in ("off", "consultivo", "automatico"):
        raise ValueError("Modo de IA inválido")
    if result.ai_provider not in ("chatgpt", "openai"):
        raise ValueError("Provedor de IA inválido")
    if result.manual_minutes_per_file <= 0 or result.review_minutes_per_file < 0:
        raise ValueError("Premissas de tempo inválidas")
    for p in (result.output_dir, result.runtime_dir):
        if p == result.input_dir or p.is_relative_to(result.input_dir):
            raise ValueError("Saída e runtime devem ficar fora da pasta de entrada")
    if not 0 < result.ocr_min_confidence <= 1 or not 0 <= result.retries <= 5:
        raise ValueError("Limite de confiança ou tentativas inválido")
    if min(result.max_file_mb, result.max_pdf_pages, result.max_image_pixels) <= 0:
        raise ValueError("Limites de arquivo, páginas e pixels devem ser positivos")
    return result
