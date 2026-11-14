# Validação do case

Esta pasta reúne os testes e execuções reais selecionadas para o avaliador. O código operacional está em `goener/`; as saídas de referência abaixo não são entradas do robô. As versões e commits de cada execução estão em `resumo.json`.

| Pasta | O que comprova |
| --- | --- |
| `tests/` | Regras financeiras, leitores, OCR, contrato JSON, cache, retry, failover, segurança de arquivos e entrega. |
| `execucoes/low/` | Cinco solicitações individuais ao 3.5 em low, após bloqueio diário do 3.8. |
| `execucoes/high/` | Os mesmos cinco documentos em high: cinco respostas válidas e comparação do consumo. |
| `execucoes/executavel_sem_ia/` | Os 17 arquivos com OCR, sem cache, Python fora do PATH e rede bloqueada por proxy. |
| `execucoes/executavel_ia_cache/` | Reutilização de cinco respostas reais em high pelo EXE, sem novas chamadas HTTP. |
| `execucoes/executavel_http/` | Uma chamada HTTP real pelo EXE, com PDF original e chave lida do .env. |
| `executavel.json` e `compilacao.log` | Condições, hash do binário, tempos e registro de construção PyInstaller. |
| `comparacao_tokens.json` | Metadados de consumo informados pela API, por documento, no ensaio low/high. |

Cada execução preservada contém planilha, HTML, resumo, evidências, logs legíveis/JSONL e manifesto de integridade. São registros históricos; caminhos absolutos e versões antigas podem aparecer porque não foram reescritos. Somente o link do HTML no ensaio HTTP foi ajustado para o documento em `extratos/`, atualizando o hash correspondente no manifesto. Os documentos do case são fictícios. Chaves de API, bancos de cache e `.env` não são versionados.

## Reproduzir

Com Python 3.12, na raiz do projeto:

```powershell
py -3.12 -m pip install -r requirements-lock.txt
py -3.12 -m pytest -q
py -3.12 -m ruff check goener executar.py validacao/tests
```

A suíte usa requisições simuladas e não precisa consumir cota da API. A instalação e os modelos OCR precisam estar disponíveis. Os ensaios HTTP reais estão nos registros, separados da suíte.

Os testes não são necessários para executar `Goener.exe`. Um resultado válido nos documentos fornecidos não garante cobertura de qualquer layout bancário futuro. A rodada high real utilizou o 3.5; o parâmetro HIGH dos três modelos também é conferido por testes de payload.

## Resultado da distribuição Windows

O EXE de versão 1.7.0 inclui interpretador, bibliotecas e modelos OCR; tem aproximadamente 146 MiB. Foi executado de outra pasta, sem fontes ao lado, sem Python no PATH e sem variáveis de chave herdadas. Os dois primeiros ensaios bloquearam rede por proxy. A máquina hospedeira possui Python: não se trata de uma VM Windows limpa. O empacotamento e os ensaios comprovam isolamento nessas condições, não compatibilidade com toda configuração de Windows.

- OCR sem cache: 17 arquivos, 12 ok, 4 revisar e 1 falhou; saldo R$ 4.114.372,29; cerca de 30 segundos de processamento, mais inicialização do EXE.
- IA com cache: os mesmos 17 arquivos, cinco análises em cache e zero novas chamadas. Os bancos locais foram copiados para esse ensaio e não são distribuídos.
- HTTP real: um PDF de comunicado, uma chamada ao 3.5 e JSON válido em high. O status financeiro `falhou` é esperado porque o comunicado não é extrato; não significa falha da chamada. Não foi usado cache nessa prova.

O hash exato do executável e os tempos medidos estão em `executavel.json`. Para reproduzir o ensaio offline após construir o EXE: `py -3.12 validacao/validar_executavel.py`. A etapa de cache requer análises reais locais já existentes; ela é omitida quando não há banco de IA/chave. O ensaio HTTP externo está registrado separadamente e consome cota se for repetido.
