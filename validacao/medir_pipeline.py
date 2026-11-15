import json
import statistics
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
if len(sys.argv) > 1:
    from dataclasses import replace

    from goener.ai import PROMPT_VERSION
    from goener.core import load_config
    from goener.pipeline import run

    folder = Path(sys.argv[1])
    folder.mkdir(parents=True, exist_ok=True)
    cfg = replace(load_config(ROOT / "config.toml"), output_dir=folder, cache_enabled=False)
    directory, summary, code = run(cfg)
    (folder / "medicao.json").write_text(
        json.dumps(
            {
                "run_id": summary["run_id"],
                "version": summary["version"],
                "pipeline_seconds": summary["duration_seconds"],
                "files": summary["files"],
                "ai": summary["ai_summary"],
                "statuses": summary["statuses"],
                "approved_balance": summary["approved_balance"],
                "published": str(directory),
                "ai_mode": cfg.ai_mode,
                "ai_scope": cfg.ai_scope,
                "thinking": cfg.ai_thinking_level,
                "prompt_version": PROMPT_VERSION,
                "ai_batch_size": cfg.ai_batch_size,
                "ai_max_output_tokens": cfg.ai_max_output_tokens,
                "exit_code": code,
            }
        ),
        encoding="utf-8",
    )
    raise SystemExit(code)
records = []
measurement = datetime.now().strftime("%Y%m%d_%H%M%S")
for i in range(1, 4):
    folder = ROOT / ".local/benchmark_agrupado" / measurement / f"rodada_{i}"
    started = time.perf_counter()
    proc = subprocess.run(
        [sys.executable, str(Path(__file__).resolve()), str(folder)],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=600,
    )
    elapsed = time.perf_counter() - started
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "processo.log").write_text(proc.stdout + "\n" + proc.stderr, encoding="utf-8")
    if proc.returncode != 0:
        raise RuntimeError(f"Rodada {i} falhou; confira {folder}")
    record = json.loads((folder / "medicao.json").read_text(encoding="utf-8"))
    record["wall_seconds"] = round(elapsed, 3)
    evidence = json.loads(
        (Path(record["published"]) / "evidencias.json").read_text(encoding="utf-8")
    )
    record["analyses"] = [
        {
            "file": r["file"],
            "decision": r["ai"].get("decision"),
            "reason": r["ai"].get("reason"),
            "batch_size": r["ai"].get("batch_size"),
            "error_category": r["ai"].get("error", {}).get("category"),
        }
        for r in evidence
        if r.get("ai") and r["ai"].get("decision") != "nao_solicitada"
    ]
    groups = {}
    for row in evidence:
        ai = row.get("ai", {})
        if not ai or ai.get("decision") == "nao_solicitada":
            continue
        error = ai.get("error", {})
        group_id = ai.get("batch_id", error.get("batch_id", row["file"]))
        group = groups.setdefault(
            group_id,
            {
                "files": [],
                "batch_size": ai.get("batch_size", error.get("batch_size")),
                "model": ai.get("selected_model"),
                "usage": ai.get("usage", {}),
                "attempts": ai.get("attempts", error.get("attempts", [])),
            },
        )
        group["files"].append(row["file"])
    record["groups"] = list(groups.values())
    records.append(record)
    print(
        json.dumps(
            {
                "round": i,
                "seconds": record["wall_seconds"],
                "http_attempts": record["ai"]["http_attempts"],
                "valid_decisions": record["ai"]["valid_decisions"],
                "unavailable": record["ai"]["unavailable_files"],
            }
        ),
        flush=True,
    )
result = {
    "version": "1",
    "date": datetime.now().date().isoformat(),
    "method": "Tempo externo entre iniciar e encerrar um processo Python, incluindo importação, leitura/OCR sem cache, IA via HTTP sem cache, validação, geração de destaques/Excel/HTML e publicação. Exclui instalação, navegador aberto aguardando operador e revisão humana.",
    "configuration": {
        "files": records[0]["files"],
        "cache_enabled": False,
        "ai_mode": records[0]["ai_mode"],
        "ai_scope": records[0]["ai_scope"],
        "thinking": records[0]["thinking"],
        "prompt_version": records[0]["prompt_version"],
        "ai_batch_size": records[0]["ai_batch_size"],
        "ai_max_output_tokens": records[0]["ai_max_output_tokens"],
    },
    "rounds": records,
    "n": len(records),
    "mean_seconds": round(statistics.mean(r["wall_seconds"] for r in records), 3),
    "min_seconds": min(r["wall_seconds"] for r in records),
    "max_seconds": max(r["wall_seconds"] for r in records),
    "sample_stdev_seconds": round(statistics.stdev(r["wall_seconds"] for r in records), 3),
    "limitations": "Três execuções no mesmo equipamento e mesmos documentos. API e carga da máquina variam; não é SLA nem teste de novos bancos.",
}
(ROOT / "validacao/tempo_pipeline.json").write_text(
    json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
)
print("Media:", result["mean_seconds"], flush=True)
