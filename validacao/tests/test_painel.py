import json
import re
import shutil
import threading
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pytest
from openpyxl import load_workbook

from goener.painel import make_server
from goener.pipeline import run
from goener.reporting import publication_directory

ROOT = Path(__file__).resolve().parents[2]
NAME = "extrato_sol_do_sertao_I_out2026.pdf"


@pytest.mark.parametrize("name", [NAME, "pau_ferro_out26.xlsx"])
def test_local_panel_validates_token_and_updates_both_outputs(config, name):
    shutil.copyfile(ROOT / "extratos" / name, config.input_dir / name)
    _, _, _ = run(config)
    with make_server(config) as server:
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        origin = f"http://127.0.0.1:{server.server_port}"
        try:
            with urlopen(origin) as response:
                html = response.read().decode()
            with urlopen(origin + "/evidencias.json") as response:
                assert response.status == 200
                assert json.load(response)[0]["file"] == name
            with urlopen(origin + "/logs/execucao.log") as response:
                assert response.status == 200
            token = re.search("const token='([^']+)'", html)[1]
            assert not token.startswith("__")
            latest = json.loads(
                (config.output_dir / "runtime/historico/latest.json").read_text(encoding="utf-8")
            )
            directory = config.output_dir / "runtime/historico" / latest["run_id"]
            evidence = json.loads((directory / "evidencias.json").read_text(encoding="utf-8"))[0]
            payload = {
                "file": name,
                "sha256": evidence["sha256"],
                "run_id": latest["run_id"],
                "decision": "Corrigir e liberar",
                "balance": "456099.12",
                "reason": "Conferido com documento de suporte",
                "reviewer": "Teste",
                "date": "2026-11-14",
            }

            def post(secret):
                request = Request(
                    origin + "/api/review",
                    data=json.dumps(payload).encode(),
                    headers={
                        "Content-Type": "application/json",
                        "X-Goener-Token": secret,
                        "Origin": origin,
                    },
                )
                try:
                    with urlopen(request) as response:
                        return response.status, json.load(response)
                except HTTPError as error:
                    return error.code, json.load(error)

            assert post("invalid")[0] == 403
            payload["reviewer"] = ""
            assert post(token)[0] == 409
            payload["reviewer"] = "Teste"
            status, _ = post(token)
            assert status == 200
            wb = load_workbook(config.output_dir / "consolidado.xlsx")
            assert wb["Consolidado"]["E8"].value == 456099.12
            wb.close()
            with urlopen(origin) as response:
                assert "Revisão humana registrada" in response.read().decode()
            assert post(token)[0] == 409  # Uma tela antiga não pode sobrescrever outra revisão.
            try:
                urlopen(origin + "/.env")
                raise AssertionError("Credenciais não devem ser servidas")
            except HTTPError as error:
                assert error.code == 404
        finally:
            server.shutdown()
            thread.join(timeout=5)


def test_editing_pinned_session_updates_its_static_html_and_not_other_session(config):
    from goener.sessao import remember_session, select_session

    shutil.copyfile(ROOT / "extratos" / NAME, config.input_dir / NAME)
    first = select_session(config, create=True)
    first_dir, _, _ = run(first)
    second = select_session(config, create=True)
    second_dir, _, _ = run(second)
    remember_session(second, config.output_dir)
    second_before = (second.output_dir / "consolidado.xlsx").read_bytes()
    static_before = (first_dir / "relatorio.html").read_bytes()
    with make_server(first) as server:
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        origin = f"http://127.0.0.1:{server.server_port}"
        try:
            with urlopen(origin) as response:
                page = response.read().decode()
            token = re.search("const token='([^']+)'", page)[1]
            evidence = json.loads((first_dir / "evidencias.json").read_text(encoding="utf-8"))[0]
            summary = json.loads((first_dir / "resumo.json").read_text(encoding="utf-8"))
            payload = {
                "file": NAME,
                "sha256": evidence["sha256"],
                "run_id": summary["run_id"],
                "decision": "Corrigir e liberar",
                "balance": "456099.12",
                "reason": "Conferido",
                "reviewer": "Teste",
                "date": "2026-11-15",
            }
            request = Request(
                origin + "/api/review",
                data=json.dumps(payload).encode(),
                headers={
                    "Content-Type": "application/json",
                    "Origin": origin,
                    "X-Goener-Token": token,
                },
            )
            with urlopen(request) as response:
                assert response.status == 200
            assert (
                publication_directory(first.output_dir) / "relatorio.html"
            ).read_bytes() != static_before
            assert not (first.output_dir / "relatorio.html").exists()
            with urlopen(origin) as response:
                assert "Revisão humana registrada" in response.read().decode()
            assert (second.output_dir / "consolidado.xlsx").read_bytes() == second_before
            assert select_session(config).output_dir == second.output_dir
            assert (first.output_dir / "Abrir relatorio dinamico.cmd").is_file()
            assert (first.output_dir / "Aplicar revisoes (Excel).cmd").is_file()
        finally:
            server.shutdown()
            thread.join(timeout=5)


def test_packaged_reference_or_snapshot_can_open_without_latest_index(config):
    from dataclasses import replace

    shutil.copyfile(ROOT / "extratos" / NAME, config.input_dir / NAME)
    directory, _, _ = run(config)
    reference_config = replace(config, output_dir=directory)
    assert not (directory / "latest.json").exists()
    with make_server(reference_config) as server:
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            with urlopen(f"http://127.0.0.1:{server.server_port}/") as response:
                assert response.status == 200
                assert NAME in response.read().decode()
        finally:
            server.shutdown()
            thread.join(timeout=5)
