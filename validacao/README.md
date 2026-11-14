# Validação do case

Esta pasta reúne os testes e execuções reais selecionadas para o avaliador. O código operacional está em `goener/`; as saídas de referência abaixo não são entradas do robô. As versões e commits de cada execução estão em `resumo.json`.

| Pasta | O que comprova |
| --- | --- |
| `tests/` | Regras financeiras, leitores, OCR, contrato JSON, cache, retry, failover, segurança de arquivos e entrega. |
| `execucoes/low/` | Cinco solicitações individuais ao 3.5 em low, após bloqueio diário do 3.8. |
| `execucoes/high/` | Os mesmos cinco documentos em high: cinco respostas válidas e comparação do consumo. |
| `execucoes/executavel/` | Validação da distribuição Windows sem usar Python instalado; detalhes em `executavel.json`. |
| `comparacao_tokens.json` | Metadados de consumo informados pela API, por documento, no ensaio low/high. |

Cada execução preservada contém planilha, HTML, resumo, evidências, logs legíveis/JSONL e manifesto de integridade. São registros históricos; caminhos absolutos e versões antigas podem aparecer porque não foram reescritos. Os documentos do case são fictícios. Chaves de API, bancos de cache e `.env` não são versionados.

## Reproduzir

Com Python 3.12, na raiz do projeto:

```powershell
py -3.12 -m pip install -r requirements-lock.txt
py -3.12 -m pytest -q
py -3.12 -m ruff check goener executar.py validacao/tests
```

A suíte usa requisições simuladas e não precisa consumir cota da API. A instalação e os modelos OCR precisam estar disponíveis. Os ensaios HTTP reais estão nos registros, separados da suíte.

Os testes não são necessários para executar `Goener.exe`. Um resultado válido nos documentos fornecidos não garante cobertura de qualquer layout bancário futuro. A rodada high real utilizou o 3.5; o parâmetro HIGH dos três modelos também é conferido por testes de payload.
