"""Preparação assistida do perfil do robô; não preenche login nem confirma desafios."""

import json
from datetime import UTC, datetime
from pathlib import Path

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

from goener.ai import browser_driver
from goener.config import load_config

root = Path(__file__).resolve().parents[1]
config = load_config(root / "config.toml")
driver = browser_driver(config, visible=True)
try:
    driver.get("https://chatgpt.com/")
    print(
        "Janela de preparação aberta. Faça o login e eventuais confirmações pessoalmente.",
        flush=True,
    )
    print(
        "Selecione o modo High/Alto na interface. A preparação aguarda por até 10 minutos.",
        flush=True,
    )

    def ready(d):
        editor = any(
            e.is_displayed() for e in d.find_elements(By.CSS_SELECTOR, '[contenteditable="true"]')
        )
        high = any(
            b.is_displayed() and b.text.strip().casefold() in ("high", "alto")
            for b in d.find_elements(By.TAG_NAME, "button")
        )
        return editor and high

    WebDriverWait(driver, 600, poll_frequency=2).until(ready)
    (config.runtime_dir / "chatgpt-ready").write_text(
        json.dumps({"checked_at": datetime.now(UTC).isoformat(), "high_visible": True}),
        encoding="utf-8",
    )
    print(
        "Perfil preparado; High observado. A chamada com envio de mensagem ainda precisa ser validada.",
        flush=True,
    )
finally:
    driver.quit()
