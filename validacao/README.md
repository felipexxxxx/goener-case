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

A regra em `.gitattributes` preserva os bytes dos registros técnicos no Git. São versionados resumo, evidências, logs e manifesto; os HTMLs, planilhas e PDFs gerados foram retirados do versionamento e continuam disponíveis somente na entrega local/pacote privado. Os manifestos históricos registram também hashes de saídas que não acompanham mais o repositório. A verificação desses artefatos exige o arquivo local correspondente.

Os registros são históricos e preservam os resultados, versões, caminhos e premissas das respectivas rodadas; não representam a configuração atual em todos os detalhes. Os commits anteriores que já continham relatórios continuam no histórico. Chaves de API, bancos de cache e `.env` não são versionados.

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

## Revisão de interface 1.8.0

`interface.json` registra testes de navegação, filtros, busca, responsividade e ausência de análise de tempo no HTML. O novo EXE foi compilado e contém o template atualizado, mas sua execução foi bloqueada pelo Smart App Control (erro 4551; evento de integridade 3077) por falta de assinatura confiável. Não houve alteração de política ou tentativa de contornar a proteção. A execução pelo Python foi validada; as provas de execução do EXE 1.7.0 são históricas e não homologam o binário 1.8.0.

Após o bloqueio registrado da versão 1.8.0, o usuário iniciou o mesmo binário e a execução concluiu com os 17 arquivos. O resultado e a conferência dos hashes foram registrados em `interface.json`. Não se identificou nos eventos a causa da liberação na segunda tentativa; o registro de bloqueio original foi preservado.

A versão 1.8.1 publica em `resultados/`, com atalhos para HTML e Excel. Foram aprovados 112 testes e uma execução real do novo EXE pelo terminal nesta máquina (17 arquivos, cinco decisões em cache, nenhuma chamada HTTP nova). O registro está em `interface.json`, seção `output_visibility`; não representa validação em uma instalação limpa do Windows nem teste de abertura automática do navegador.


## Revisões e painel local 1.9.0

121 testes passaram na suíte completa. Após os últimos ajustes de apresentação, 14 verificações afetadas passaram novamente. Foram conferidos: correção, liberação, exclusão, saldo zero, campos obrigatórios, mudança do documento, preservação de revisão durante edição concorrente e hashes das cópias destacadas. O endpoint local exige token/origem, rejeita telas antigas e não serve `.env`.

O formulário foi testado no navegador em uma cópia separada de Sol do Sertão: saldo confirmado de R$ 456.099,12, original R$ 457.179,12 preservado, Excel/HTML atualizados e zero novas chamadas à IA. Essa correção é um ensaio, não uma decisão aplicada ao resultado entregue. A base real permanece com 12 ok, 4 revisar e 1 falhou. A prévia da aba Revisões foi conferida, e as fórmulas do Resumo foram recalculadas no artifact-tool. Isso não equivale a homologação em todas as versões do Excel.

Registros em `interface.json`, seção `reviews_panel`. O comando operacional rápido é `py -3.12 executar.py --update-reviews`; o painel é `py -3.12 executar.py --panel`. O teste funcional de HTTP está em `tests/test_painel.py`; ele usa servidor local e não consome API externa.
