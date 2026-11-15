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

122 testes passaram na suíte completa final. Foram conferidos: correção, liberação, exclusão, saldo zero, campos obrigatórios, mudança do documento, preservação de revisão durante edição concorrente e hashes das cópias destacadas. O endpoint local exige token/origem, rejeita telas antigas e não serve `.env`.

O formulário foi testado no navegador em uma cópia separada de Sol do Sertão: saldo confirmado de R$ 456.099,12, original R$ 457.179,12 preservado, Excel/HTML atualizados e zero novas chamadas à IA. Essa correção é um ensaio, não uma decisão aplicada ao resultado entregue. A base real permanece com 12 ok, 4 revisar e 1 falhou. A prévia da aba Revisões foi conferida, e as fórmulas do Resumo foram recalculadas no artifact-tool. Isso não equivale a homologação em todas as versões do Excel.

Registros em `interface.json`, seção `reviews_panel`. O comando operacional rápido é `py -3.12 executar.py --update-reviews`; o painel é `py -3.12 executar.py --panel`. O teste funcional de HTTP está em `tests/test_painel.py`; ele usa servidor local e não consome API externa.

O EXE 1.9.0 foi bloqueado na primeira tentativa (4551) e concluiu na repetição do mesmo comando, sem mudanças na proteção do Windows. O painel também foi iniciado pelo EXE e o formulário conferido no navegador. Os cinco destaques e seus hashes foram verificados. Logs e JSON da execução estão em `execucoes/painel_local/`; PDF, HTML e Excel gerados permanecem fora do Git.


## Conferência final da entrega

`entrega.json` registra o ensaio a partir do ZIP extraído em outra pasta. O EXE processou os 17 arquivos com IA desligada, sem caches distribuídos: 12 ok, 4 revisar, 1 falhou e R$ 4.114.372,29. Foram 38,124 segundos de processamento, além da inicialização. O único reaproveitamento ocorreu dentro do próprio lote, por causa da cópia idêntica. Planilha com oito colunas, fórmulas, revisões e hashes das saídas foram conferidos.

A atualização por Python também foi executada usando o fonte extraído, com os mesmos resultados e zero chamadas HTTP. O teste de empacotamento passou. README, fontes, EXE e documento foram comparados com o conteúdo do ZIP; o PDF foi renderizado e conferido em uma página. A validação não representa máquina sem Python instalado nem política Windows universal. Nenhum novo teste externo de IA foi feito nessa conferência.


## Estimativa por atividade 1.9.1

124 testes passaram. O modelo de tempo foi corrigido para separar extratos únicos da competência de arquivos que exigem apenas triagem inicial. Neste lote: 14 extratos (5 PDF, 5 Excel e 4 imagens) e três triagens. Referência de consolidação: 10 min declarados; triagem: hipótese de 1 min. Cenários são análise de sensibilidade, sem nível de confiança estatístico atribuído. Uma referência manual não sustenta IC de 95%.

`tempo.json` preserva contagens, hipóteses, tempo medido e projeções mensal/anual. Foram testados volumes/frequências diferentes, separação de duplicidade/outro período/comunicado, preservação de ganho negativo e ausência de IC fabricado. O HTML continua sem análise de tempo. As projeções antigas ficam como histórico, substituídas pela metodologia versão 3 nos novos resultados.

O EXE 1.9.1 também concluiu `--update-reviews` na primeira tentativa deste ensaio, com modelo de tempo versão 3 e zero chamadas externas. Fonte/EXE/documentação foram reempacotados juntos após a alteração.


## Abertura e organização 1.9.2

128 testes passaram: supervisor conserva código/log após encerramento abrupto do filho, falha interativa aguarda Enter, sucesso não pausa, trava mantida pelo SO é recuperada após matar o processo e processo anterior ativo não é removido. O histórico agora fica agrupado em resultados/historico; a raiz contém apenas o Excel operacional, o HTML e esse diretório. Não foram adicionadas mensagens de progresso. O diagnóstico nativo é complementar e não garante capturar todo tipo de término forçado do Windows.

## Distribuição Python 1.10.0

PyInstaller e EXE foram retirados da entrega por decisão do responsável. Os registros de binários anteriores são históricos. A versão atual usa Python 3.12 e Executar automacao.cmd, preservando o supervisor, os diagnósticos, a recuperação de trava, as revisões e o histórico organizado. 128 testes passaram.

O CMD da distribuição 1.10.0 foi executado a partir do ZIP extraído em pasta temporária, com IA desligada e sem cache distribuído: 17 arquivos, 12 ok, 4 revisar, 1 falhou, R$ 4.114.372,29 e 34,523 segundos de processamento. Registro em python.json. Nenhum processo Goener.exe permaneceu ativo durante essa validação.

## Portabilidade 1.10.1

131 testes passaram. Foram adicionados testes de mudança de pasta, índice relativo, recuperação de índice absoluto legado e links dos artefatos. O modelo HTML mostra aviso ao ser aberto diretamente. A rota HTTP de evidências e a de log também foram verificadas.

## Compatibilidade Excel Desktop 1.10.2

O usuário identificou reparo de tabela e fórmulas `#REF!` no Excel Desktop. A reabertura via openpyxl dos testes anteriores não era suficiente para comprovar compatibilidade. O defeito foi reproduzido no Excel com uma cópia do arquivo anterior. Removendo somente o AutoFilter da planilha que se sobrepunha ao da tabela, o Excel passou a abrir sem reparo e a calcular as referências estruturadas.

O gerador foi corrigido, a coluna do total ampliada e a saída regenerada por `--update-reviews`, preservando originais e análises, sem chamadas à IA. Uma cópia dessa nova saída foi aberta no Excel Desktop: total R$ 4.114.372,29, 12 aprovados, 4 revisões e 1 falha. Alterar E2 de 271350,74 para 271351,74 atualizou o resumo para R$ 4.114.373,29; a edição foi desfeita. A competência foi alterada/restaurada em outra cópia, com recálculo dos elegíveis. `excel_desktop.json` registra os arquivos e o escopo desse ensaio.

132 testes passaram, incluindo regressão que inspeciona o XML para impedir filtros sobrepostos, mesmo se o modelo de entrada já tiver AutoFilter. Ruff passou. Isso valida o Excel instalado nesta máquina e os cenários descritos, sem prometer compatibilidade universal ou ausência de qualquer outro defeito.

## Planilha simplificada 1.11.0

A saída passou de quatro para duas abas, mantendo os 17 arquivos e as oito colunas exigidas. O resumo com fórmulas ficou no topo do Consolidado; a tabela começa na linha 7. As observações ficaram curtas; a evidência detalhada não foi alterada. Revisões mostra somente pendências e decisões/rascunhos preenchidos, preservando inclusive um saldo zero digitado em revisão ainda incompleta.

134 testes passaram; Ruff passou. No Excel Desktop, uma cópia da saída abriu sem reparo, com total R$ 4.114.372,29 e contagens 12/4/1. Alterar E8 de 271350,74 para 271351,74 recalculou E2 para R$ 4.114.373,29; desfazer restaurou o total anterior. O fluxo pelo painel também foi testado, gravando a correção na nova posição da tabela e regenerando Excel/HTML. `excel_simples.json` identifica a saída validada. Não houve chamada externa de IA nesta atualização.


## Clareza do saldo e edição HTML 1.11.1

136 testes passaram e Ruff passou. O extrato Sol do Sertão permanece revisar, sem decisão humana, com contribuição zero ao total: o saldo declarado exibido é apenas referência. O HTML passou a mostrar explicitamente esse efeito e a espera por revisão. A instrução antiga sobre EXE foi substituída pelo comando Python. No arquivo aberto diretamente, o botão de revisão fica desabilitado; no painel local, fica habilitado. A interface real foi aberta no Chrome, o cartão expandido e o modal de revisão conferido, sem salvar alterações nos dados do responsável. O POST de correção foi testado separadamente com dados temporários, incluindo um arquivo sem pendência inicial cuja linha em Revisões é criada na primeira decisão. Registro em html_revisao.json.


## Iniciador sem parâmetros 1.12.0

Executar automacao.cmd foi chamado sem argumentos e abriu automaticamente o painel local no Chrome. A publicação existente permaneceu a mesma; a resposta HTTP 200 continha a sessão de revisão. O teste unitário impede reprocessamento e leitura de Revisões na abertura de um resultado existente. Outro teste cobre a primeira abertura sem resultados: processa o extrato e depois serve o painel. 138 testes passaram e Ruff passou. A aplicação das revisões salvas no Excel continua explícita com --update-reviews, opcionalmente junto de --panel; as decisões no formulário continuam sendo aplicadas por Salvar e atualizar. Registro em cmd_painel.json.
