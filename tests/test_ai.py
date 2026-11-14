import io
import json

import pytest

from goener.ai import SCHEMA, request_browser, request_decision, review, validate_decision
from goener.models import Result, Statement, money


def test_http_payload_and_schema(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "test-only")

    def opener(req, timeout):
        body = json.loads(req.data)
        assert req.full_url == "https://api.openai.com/v1/responses"
        assert body["store"] is False
        assert body["text"]["format"]["schema"] == SCHEMA
        assert body["text"]["format"]["strict"] is True
        return io.StringIO(
            json.dumps(
                {
                    "status": "completed",
                    "id": "test",
                    "output": [
                        {
                            "content": [
                                {
                                    "type": "output_text",
                                    "text": json.dumps(
                                        {
                                            "decision": "revisar",
                                            "reason": "Conferir origem",
                                            "evidence_codes": [],
                                        }
                                    ),
                                }
                            ]
                        }
                    ],
                }
            )
        )

    assert request_decision({"issues": []}, "test", opener)["decision"] == "revisar"


@pytest.mark.parametrize(
    "bad",
    [
        {},
        {"decision": "executar_codigo", "reason": "x", "evidence_codes": []},
        {"decision": "revisar", "reason": "x", "evidence_codes": ["inventado"]},
    ],
)
def test_invalid_json_rejected(bad):
    with pytest.raises(ValueError):
        validate_decision(bad, set())


def test_two_modes_and_policy(config):
    r = Result("a.pdf", "sha", "PDF", Statement("Banco", "1-0", "2026-10", closing=money(10)))
    config.ai_mode = "consultivo"
    review(
        [r],
        config,
        lambda *_: {
            "decision": "nao_considerar",
            "reason": "Conferir origem",
            "evidence_codes": [],
        },
    )
    assert r.included(config.ai_mode, config.period)
    config.ai_mode = "automatico"
    assert not r.included(config.ai_mode, config.period)
    r.statement.issue("DIVERGENCIA", "Saldo diverge")
    review(
        [r],
        config,
        lambda *_: {"decision": "considerar", "reason": "Liberar", "evidence_codes": []},
    )
    assert not r.included(config.ai_mode, config.period)
    assert r.ai["policy_block"]


def test_missing_key_and_api_failure(config, monkeypatch):
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    config.ai_mode = "automatico"
    config.ai_provider = "openai"
    r = Result("a", "sha", "PDF", Statement("Banco", "1-0", "2026-10", closing=money(10)))
    review([r], config)
    assert r.ai["decision"] == "indisponivel"
    assert not r.included(config.ai_mode, config.period)


def test_browser_requires_prepared_profile_before_loading_selenium(config):
    with pytest.raises(ValueError, match="sessão"):
        request_browser({"issues": []}, config)
