import json
import shutil
from pathlib import Path

from goener.pipeline import run
from goener.reporting import publication_directory

ROOT = Path(__file__).resolve().parents[2]


def test_links_and_relative_index_survive_moving_output_folder(config, tmp_path):
    shutil.copyfile(ROOT / "extratos/pau_ferro_out26.xlsx", config.input_dir / "extrato.xlsx")
    directory, _, _ = run(config)
    pointer = config.output_dir / "historico/latest.json"
    data = json.loads(pointer.read_text(encoding="utf-8"))
    assert data["directory"] == directory.name
    moved = tmp_path / "outro_usuario" / "outro_projeto" / "resultados"
    shutil.copytree(config.output_dir, moved)
    resolved = publication_directory(moved)
    assert resolved == moved / "historico" / directory.name
    html = (resolved / "relatorio.html").read_text(encoding="utf-8")
    for name in ("evidencias.json", "execucao.log", "manifesto.json", "consolidado.xlsx"):
        assert f'href="{name}"' in html
        assert (resolved / name).is_file()


def test_old_absolute_index_is_rebased_to_local_run(config, tmp_path):
    shutil.copyfile(ROOT / "extratos/pau_ferro_out26.xlsx", config.input_dir / "extrato.xlsx")
    directory, _, _ = run(config)
    pointer = config.output_dir / "historico/latest.json"
    data = json.loads(pointer.read_text(encoding="utf-8"))
    data["directory"] = str(tmp_path / "usuario_anterior" / "pasta_antiga" / directory.name)
    pointer.write_text(json.dumps(data), encoding="utf-8")
    assert publication_directory(config.output_dir) == directory


def test_template_shows_explanation_instead_of_broken_links():
    template = (ROOT / "goener/report.html").read_text(encoding="utf-8")
    assert "Este arquivo é o modelo do relatório" in template
    assert 'href="../resultados/relatorio.html"' not in template
    assert "['__','RUN','__'].join('')" in template
