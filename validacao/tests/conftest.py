from pathlib import Path

import pytest

from goener.core import load_config

ROOT = Path(__file__).resolve().parents[2]


@pytest.fixture
def config(tmp_path):
    cfg = load_config(ROOT / "config.toml")
    cfg.ai_mode = "off"
    cfg.ai_scope = "todos"
    cfg.input_dir = tmp_path / "input"
    cfg.input_dir.mkdir()
    cfg.output_dir = tmp_path / "output"
    cfg.runtime_dir = tmp_path / "runtime"
    cfg.stop_file = cfg.runtime_dir / "STOP"
    return cfg
