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
            with urlopen(origin + "/execucao.log") as response:
                assert response.status == 200
            token = re.search("const token='([^']+)'", html)[1]
            assert not token.startswith("__")
            latest = json.loads(
                (config.output_dir / "historico/latest.json").read_text(encoding="utf-8")
            )
            directory = config.output_dir / "historico" / latest["run_id"]
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
