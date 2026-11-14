from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .config import load_config
from .pipeline import run


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
