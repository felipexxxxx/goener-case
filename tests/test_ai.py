import io
import json

import pytest

from goener.ai import (
    SCHEMA,
    request_decision,
    request_gemini,
    review,
    validate_decision,
)
from goener.core import Result, Statement, money


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


def test_gemini_http_contract(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-only")

    def opener(req, timeout):
        assert req.full_url.endswith("/models/gemini-3.8-flash:generateContent")
        assert "test-only" not in req.full_url
        assert req.get_header("X-goog-api-key") == "test-only"
        assert timeout == 45
        payload = json.loads(req.data)
        assert payload["generationConfig"]["responseJsonSchema"] == SCHEMA
        assert "tools" not in payload
        return io.StringIO(
            json.dumps(
                {
                    "responseId": "fake-id",
                    "modelVersion": "gemini-3.8-flash",
                    "candidates": [
                        {
                            "finishReason": "STOP",
                            "content": {
                                "parts": [
                                    {"thought": True, "text": "ignorar raciocínio"},
                                    {
                                        "text": json.dumps(
                                            {
                                                "decision": "revisar",
                                                "reason": "Conferir",
                                                "evidence_codes": [],
                                            }
                                        )
                                    },
                                ]
                            },
                        }
                    ],
                }
            )
        )

    result = request_gemini({"issues": []}, "gemini-3.8-flash", opener)
    assert result["decision"] == "revisar"
    assert result["response_id"] == "fake-id"


@pytest.mark.parametrize(
    "body",
    [
        {"promptFeedback": {"blockReason": "SAFETY"}},
        {"candidates": [{"finishReason": "MAX_TOKENS"}]},
        {"candidates": [{"finishReason": "STOP", "content": {"parts": [{"text": "inválido"}]}}]},
    ],
)
def test_gemini_rejects_blocked_partial_or_invalid(monkeypatch, body):
    monkeypatch.setenv("GEMINI_API_KEY", "test-only")
    with pytest.raises(ValueError):
        request_gemini({"issues": []}, "test", lambda *a, **k: io.StringIO(json.dumps(body)))


def test_gemini_quota_is_bounded_and_automatic_mode_retains(config, monkeypatch):
    from urllib.error import HTTPError

    monkeypatch.setenv("GEMINI_API_KEY", "test-only")
    monkeypatch.setattr("goener.ai.time.sleep", lambda _: None)
    calls = []

    def opener(req, timeout):
        calls.append(req)
        raise HTTPError(req.full_url, 429, "quota", {}, None)

    config.ai_mode = "automatico"
    config.ai_provider = "gemini"
    r = Result("a", "sha", "PDF", Statement("Banco", "1-0", "2026-10", closing=money(10)))
    review([r], config, lambda context, model: request_gemini(context, model, opener))
    assert len(calls) == 3
    assert not r.included(config.ai_mode, config.period)
    assert r.ai["decision"] == "indisponivel"
