"""Painel local: usa o mesmo pipeline, sem hospedagem ou dependências web."""

from __future__ import annotations

import json
import mimetypes
import os
import secrets
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import unquote, urlparse

from openpyxl import load_workbook

from .core import Result, Statement
from .reporting import latest_pointer, publication_directory
from .revisoes import apply_reviews, digest, read_reviews


def make_server(config, *, port=0):
    token = secrets.token_urlsafe(32)

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass

        def reply(self, code, content, kind="application/json; charset=utf-8"):
            data = (
                content
                if isinstance(content, bytes)
                else json.dumps(content, ensure_ascii=False).encode()
            )
            self.send_response(code)
            self.send_header("Content-Type", kind)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("X-Frame-Options", "DENY")
            self.end_headers()
            self.wfile.write(data)

        def latest(self):
            pointer = json.loads(latest_pointer(config.output_dir).read_text(encoding="utf-8"))
            directory = publication_directory(config.output_dir)
            if directory.parent not in (
                config.output_dir.resolve(),
                (config.output_dir / "historico").resolve(),
            ):
                raise ValueError("Publicação fora da pasta de resultados.")
            return pointer, directory

        def do_GET(self):
            try:
                _, directory = self.latest()
                route = unquote(urlparse(self.path).path).lstrip("/")
                if route in ("", "relatorio.html"):
                    document = (directory / "relatorio.html").read_text(encoding="utf-8")
                    document = document.replace("__LOCAL_REVIEW_TOKEN__", token)
                    return self.reply(200, document.encode(), "text/html; charset=utf-8")
                if route.startswith("extratos/"):
                    base = config.input_dir.resolve()
                    path = base / route.removeprefix("extratos/")
                elif route.startswith("destaques/") or route in {
                    "consolidado.xlsx",
                    "execucao.log",
                    "evidencias.json",
                    "manifesto.json",
                }:
                    base = directory
                    path = directory / route
                else:
                    return self.reply(404, {"error": "Arquivo não encontrado."})
                if (
                    not path.resolve().is_relative_to(base)
                    or path.is_symlink()
                    or not path.is_file()
                ):
                    return self.reply(404, {"error": "Arquivo não encontrado."})
                self.reply(
                    200,
                    path.read_bytes(),
                    mimetypes.guess_type(path.name)[0] or "application/octet-stream",
                )
            except Exception:
                self.reply(
                    500, {"error": "Não foi possível abrir a publicação. Confira o terminal."}
                )

        def do_POST(self):
            origin = f"http://127.0.0.1:{self.server.server_port}"
            if self.headers.get("X-Goener-Token") != token or self.headers.get("Origin") != origin:
                return self.reply(403, {"error": "Abra o painel local novamente para revisar."})
            if self.path != "/api/review" or self.headers.get_content_type() != "application/json":
                return self.reply(400, {"error": "Requisição inválida."})
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if not 0 < length <= 16000:
                    raise ValueError("Revisão excede o tamanho permitido.")
                request = json.loads(self.rfile.read(length))
                if not isinstance(request, dict):
                    raise ValueError("Revisão inválida.")
                pointer, directory = self.latest()
                if request.get("run_id") != pointer["run_id"]:
                    raise ValueError(
                        "O resultado foi atualizado em outra janela. Recarregue antes de revisar."
                    )
                manifest = json.loads((directory / "manifesto.json").read_text(encoding="utf-8"))
                if (
                    digest(directory / "evidencias.json")
                    != manifest["artifacts"]["evidencias.json"]
                ):
                    raise ValueError("Evidências alteradas; execute o processamento completo.")
                evidence = json.loads((directory / "evidencias.json").read_text(encoding="utf-8"))
                old = next(
                    (
                        r
                        for r in evidence
                        if r["file"] == request.get("file") and r["sha256"] == request.get("sha256")
                    ),
                    None,
                )
                if old is None:
                    raise ValueError("Documento não pertence a esta publicação.")
                source = config.input_dir / old["file"]
                if (
                    source.resolve().parent != config.input_dir.resolve()
                    or source.is_symlink()
                    or digest(source) != old["sha256"]
                ):
                    raise ValueError("O documento mudou. Execute o processamento completo.")
                if config.stop_file.exists():
                    raise ValueError(
                        "O interruptor STOP está acionado; remova-o antes de atualizar."
                    )
                operational = config.output_dir / "consolidado.xlsx"
                rows, before = read_reviews(operational)
                selected = [
                    r for r in rows if tuple(r[:3]) == (old["file"], old["sha256"], config.period)
                ]
                if len(selected) != 1:
                    raise ValueError("Não há uma linha única deste documento na aba Revisões.")
                row = selected[0]
                for field in ("decision", "reason", "reviewer", "date"):
                    if not isinstance(request.get(field), str) or len(request[field]) > 2000:
                        raise ValueError("Preencha os campos da revisão.")
                row[5:10] = [
                    request["decision"],
                    request.get("balance"),
                    request["reason"],
                    request["reviewer"],
                    request["date"],
                ]
                check = Result(
                    old["file"],
                    old["sha256"],
                    old["format"],
                    Statement.from_dict(old["statement"]),
                    duplicate_of=old["duplicate_of"],
                    ai=old["ai"],
                )

                class Log:
                    def event(self, *_args, **_fields):
                        pass

                apply_reviews([check], [row], config.period, Log())
                if (
                    request["decision"] != "Pendente"
                    and check.human_review.get("state") != "applied"
                ):
                    raise ValueError(check.human_review.get("message", "Decisão inválida."))
                wb = load_workbook(operational, keep_links=False)
                temporary = operational.with_name("consolidado.revisao.tmp.xlsx")
                try:
                    ws = wb["Revisões"]
                    index = next(
                        i
                        for i in range(5, ws.max_row + 1)
                        if tuple(ws.cell(i, j).value for j in (1, 2, 3)) == tuple(row[:3])
                    )
                    for col in range(6, 11):
                        cell = ws.cell(index, col, row[col - 1] if row[col - 1] is not None else "")
                        if isinstance(row[col - 1], str):
                            cell.data_type = "s"
                    wb.save(temporary)
                    if digest(operational) != before:
                        raise ValueError(
                            "A planilha mudou durante a revisão. Recarregue e tente novamente."
                        )
                    os.replace(temporary, operational)
                finally:
                    wb.close()
                    temporary.unlink(missing_ok=True)
                from .pipeline import run

                output, summary, _ = run(config, reviews_only=True)
                self.reply(200, {"ok": True, "run_id": summary["run_id"], "directory": str(output)})
            except (ValueError, OSError, KeyError) as exc:
                self.reply(409, {"error": str(exc)})
            except Exception:
                self.reply(
                    500,
                    {
                        "error": "A atualização não concluiu. Confira o terminal e a aba Revisões; decisões já salvas ficam preservadas."
                    },
                )

    return HTTPServer(("127.0.0.1", port), Handler)


def serve(config):
    with make_server(config) as server:
        url = f"http://127.0.0.1:{server.server_port}/"
        print(
            f"Painel local: {url}\nMantenha esta janela aberta. Para encerrar, pressione Ctrl+C.",
            flush=True,
        )
        webbrowser.open(url)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
