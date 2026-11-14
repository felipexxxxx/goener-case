# Consolidação de extratos — Goener

Automação em Python: recebe uma pasta com PDFs, Excel e imagens, lê o saldo final de cada conta e gera a planilha no modelo do case. O fluxo inteiro roda sem copiar valores manualmente.

## Para usar

1. Instale Python 3.12 e baixe este repositório.
2. Coloque os documentos em `extratos/`.
3. Abra o PowerShell nesta pasta e execute:

```powershell
.\executar.ps1
```

O script instala as dependências na primeira execução. Nas seguintes, reaproveita o ambiente. A instalação e o download inicial dos modelos OCR precisam de internet. Com a IA desligada, a leitura dos documentos acontece localmente.

O terminal mostra a pasta do resultado. `.local/outputs/latest.json` aponta para a última execução concluída. Cada execução preserva suas evidências. A pasta `entrega/` contém somente os três arquivos para envio: `goener_automacao.zip`, `consolidado.xlsx` e `documentacao_1_pagina.pdf`. O relatório, os logs e as evidências permanecem no resultado da execução e dentro do ZIP, em `resultado/`.

| Arquivo gerado | Para que serve |
| --- | --- |
| `consolidado.xlsx` | Uma linha por documento nas oito colunas do modelo. |
| `relatorio.html` | Conferência no navegador, com busca, filtros e orientações. |
| `execucao.log` | O que foi processado e o que apresentou problema. |
| `execucao.jsonl` | Eventos de execução para integração e auditoria. |
| `evidencias.json` | Dados extraídos, origem do saldo, conciliação e análise da IA. |
| `resumo.json` | Quantidades, duração medida e saldo aprovado. |
| `manifesto.json` | Hashes para conferir a integridade dos arquivos. |

### Entender os resultados

- **ok:** campos essenciais encontrados e verificações sem ressalvas.
- **revisar:** há uma pendência descrita em Observação; conferir antes de considerar o saldo.
- **falhou:** documento inválido ou dados essenciais indisponíveis.

Nenhum saldo é corrigido para fazer a conta fechar. Os originais são preservados. Valores com ressalvas permanecem na planilha, mas ficam fora do total aprovado.

**Duplicidade registrada:** `EXT_102026_LagoaAzul (1).pdf` é uma cópia exata de `EXT_102026_LagoaAzul.pdf`. A comparação usa SHA-256; ambas as linhas ficam na planilha, e a cópia não entra nas somas. Possíveis lançamentos repetidos dentro de um extrato são sinalizados para conferência, sem apagar movimentações.

O cenário de referência tem 17 arquivos: 12 ok, 4 revisar e 1 falhou. O saldo aprovado pelas regras é R$ 4.114.372,29. O aviso de agência não é um extrato; há também divergência de conciliação e uma imagem de setembro em um lote de outubro.

### Conferência rápida no Excel e no HTML

A planilha mantém as oito colunas do modelo, filtro e cabeçalho congelado. Status usa cor e texto; Observação separa pendência, onde conferir, referência e análise da IA. Os saldos continuam numéricos e os originais são preservados.

O relatório abre na fila de atenção e mostra o motivo antes de expandir o arquivo. Há filtros para recomendações concluídas e indisponibilidades da IA. Dentro de cada item: regras → IA → onde conferir → próximo passo. O botão abre o documento de origem; em PDF, inclui a página quando identificada. A origem da indicação fica explícita: extração local ou sugestão da IA. Quando a localização não é conhecida, o relatório pede conferência do documento completo, sem inventar página.

Os extratos originais não são alterados. A coluna Observação e o HTML informam a medida aplicada: incluir saldo, reter valor, excluir cópia do total, excluir competência diferente ou registrar documento sem saldo. A geração termina com todas as linhas e ressalvas, sem esperar aprovação humana. Revisão financeira posterior é necessária somente para resolver informações que a fonte não permite determinar, como solicitar o extrato correto ou confirmar uma divergência bancária.

O formato original não é alterado para satisfazer a IA. Um alerta consultivo pode exigir atenção mesmo quando as regras passaram, mas não muda automaticamente o total. As indicações geradas pelo modelo precisam ser confirmadas no original. Os links de origem do relatório são ajustados no ZIP para os arquivos do case que acompanham o pacote.

### Configuração e comandos

Edite `config.toml` para trocar pasta, competência, limites e modo de IA. Os caminhos são relativos ao arquivo de configuração. O case usa outubro de 2026; sem `period`, o comando Python usa o mês anterior em São Paulo.

```powershell
.\executar.ps1 -Period 2026-11
.\executar.ps1 -NoCache       # reextrair todos os documentos
.\executar.ps1 -NoOCR         # registrar imagens sem extrair por OCR
.\executar.ps1 -Strict        # avisar pelo código de saída quando houver pendências
.\executar.ps1 -Validate      # executar os testes antes do lote
```

No Linux: `sh executar.sh --period 2026-10`.

## Para manter

A estrutura usa cinco módulos de trabalho, sem camadas de serviços, repositórios ou controladores:

```text
goener/
  core.py           configuração, dados e regras financeiras
  readers.py        leitura de PDF, Excel e OCR
  ai.py             HTTP com OpenAI ou Gemini e contrato JSON
  reporting.py      geração da planilha e do relatório
  pipeline.py       execução, cache, logs e linha de comando
  report.html       apresentação do relatório
  entrega.py        empacotamento dos três arquivos para envio
extratos/           entradas fictícias do case
entrega/            apenas ZIP, planilha e documentação de uma página
tests/             testes automatizados
modelo.xlsx         modelo original fornecido
config.toml         parâmetros da execução
requirements-lock.txt  versões fixadas das dependências
```

`__init__.py` e `__main__.py` permitem executar `python -m goener`. `.github/` contém a integração contínua. `.venv/` guarda o ambiente Python. `.local/` guarda saídas, checkpoints, logs e arquivos locais anteriores; essas pastas não vão para o Git nem para o ZIP. O histórico das versões entregues continua disponível no Git.

### Fluxo e escolhas técnicas

Entrada → limites e hash → leitura direta ou OCR → conciliação e competência → duplicidades → IA opcional → Excel, relatório e logs → conferência dos arquivos gerados → publicação.

- `Decimal` preserva centavos e aceita saldos negativos.
- A conciliação compara saldo inicial + movimentações com o saldo final declarado.
- OCR roda em CPU; confiança baixa e ambiguidade viram ressalvas.
- SQLite reaproveita extrações por conteúdo, código do leitor e opções. As regras de negócio são reaplicadas em cada execução.
- O processamento é **sequencial**. Não há paralelismo entre arquivos; isso mantém a operação simples e evita concorrência na escrita.
- Um lock impede duas execuções no mesmo runtime. As saídas são montadas em uma pasta temporária e publicadas após validação.
- Um documento com erro não derruba os demais. Falha de configuração ou infraestrutura não é declarada como conclusão bem-sucedida.
- Logs registram duração, tentativas e resultados finais. Conte documentos pelo resumo ou pelos eventos `FILE_FINISHED`, sem contar mensagens de tentativa como falhas definitivas.

Para solicitar parada entre arquivos, crie `.local/runtime/STOP`. O arquivo em andamento termina e os restantes são registrados como cancelados. Remova esse arquivo para retomar com uma nova execução. Se um encerramento abrupto deixar `execution.lock`, confira se o PID já terminou antes de remover o lock.

Códigos de saída: `0` fluxo publicado; `1` configuração/infraestrutura; `2` pendências em modo estrito; `3` parada solicitada; `130` interrupção pelo operador.

### IA opcional

O padrão é `ai_mode = "automatico"` e `ai_scope = "excecoes"`. O Python lê todos os documentos, valida campos, competência, duplicidade e conciliação. Só os que ficam como revisar/falhou seguem para o Gemini. Os regulares são concluídos pelas regras, com registro de IA dispensada. A IA exige `GEMINI_API_KEY` no ambiente; sem chave ou serviço disponível, as pendências continuam identificadas, sem impedir a publicação do lote. A configuração inicial da chave ocorre na implantação, sem login ou confirmação durante os lotes.

```powershell
.\executar.ps1                              # automático somente nas exceções
.\executar.ps1 -AiMode off                  # apenas regras locais
.\executar.ps1 -AiMode consultivo -AiScope todos  # conferência adicional de todos
.\executar.ps1 -AiMode automatico            # decisão aplicada às exceções, sem superar bloqueios
```

**Consultivo:** registra recomendação, justificativa e indicação de onde conferir, preservando os totais pelas regras. **Automático:** a IA pode reter o saldo dos arquivos escolhidos pelo escopo; para incluí-los, eles também precisam passar nas verificações obrigatórias. Com escopo excecoes, os arquivos regulares não dependem de uma resposta da IA. Com todos, cada arquivo precisa da decisão válida para entrar no total automático. Duplicidade, competência incorreta, campos ausentes e divergências não são liberados pela IA.

O contrato básico contém `decision` (`considerar`, `nao_considerar` ou `revisar`), `reason` e `evidence_codes`. No modo completo do Gemini, o JSON também exige `source_check`: banco, conta, competência, saldo, legibilidade, localização e um trecho de evidência do original. O Python valida o contrato e compara essa leitura com a extração local. Campos extras e códigos de evidência inventados são rejeitados. Uma citação gerada pela IA não comprova, sozinha, que a leitura está correta.

O Gemini recebe as premissas do fechamento, exemplos de decisão, o resultado das regras, todas as movimentações extraídas, texto/OCR, confiança crítica e evidência do saldo disponíveis. As coordenadas de todos os tokens OCR ficam na auditoria local, sem repetir essa matriz inteira na requisição. As instruções ficam versionadas em `goener/ai.py` (`PROMPT_VERSION = "2"`), sem agente persistente ou treinamento dos pesos. Não há ferramentas para executar código ou seguir links. Documentos, fórmulas e OCR são tratados como dados não confiáveis.

| `ai_document_mode` | O que é enviado ao Gemini |
| --- | --- |
| `completo` | PDF original inteiro ou imagem original; para Excel, células preenchidas de todas as abas, com aba e coordenada. Inclui a extração local para comparação. |
| `dados` | Somente dados, movimentações e evidências extraídos pelo Python/OCR. Não exige uma leitura independente do original. |

O modo completo evita depender da primeira aba que o serviço lê ao receber Excel nativamente. Fórmulas seguem como texto e nunca são executadas pelo robô. Gráficos, imagens incorporadas e formatação do Excel não são enviados nessa conversão. PDFs e imagens seguem por `inlineData`, sem criar arquivos remotos pela Files API. Ao ativar a IA, o conteúdo escolhido é transmitido ao provedor.

Os limites são explícitos: até 8 MB por original, limite de páginas/pixels do projeto, até 30 mil células examinadas no Excel e 100 mil caracteres por contexto/texto de planilha. Se exceder, a análise fica indisponível; não há corte silencioso nem indicação falsa de documento completo. Antes de enviar, o hash deve coincidir com a extração; links simbólicos e caminhos fora da entrada são bloqueados.

A IA nunca sobrescreve banco, conta, período ou saldo do consolidado. No consultivo, diferenças da leitura independente ficam na observação e nas evidências, preservando a seleção pelas regras. No automático, uma diferença ou falta de evidência bloqueia a inclusão mesmo que o modelo recomende considerar.

Gemini usa HTTPS `generateContent`, esquema JSON e chave no cabeçalho. Também é aceita `GOOGLE_API_KEY`. OpenAI usa a Responses API: altere o provedor para `openai`, escolha o modelo e configure `OPENAI_API_KEY`. Não há troca automática entre provedores. No Gemini, até três tentativas atendem HTTP 429/5xx, timeout e falha de rede, com espera crescente, variação aleatória e `Retry-After` numérico limitado a um dia e o `RetryInfo` do provedor. Erros permanentes não são repetidos. Bloqueio, cota esgotada, resposta parcial ou JSON inválido retêm o saldo no modo automático.

Parâmetros editáveis: `ai_thinking_level = "low"` (Gemini 3; também aceita medium/high), `ai_timeout_seconds = 60`, `ai_max_attempts = 3` e `ai_batch_timeout_seconds = 300`. O timeout de rede é por operação; não é uma garantia rígida de duração total. O orçamento impede novas tentativas após o prazo e reduz o timeout ao tempo restante. Após três análises externas falharem seguidamente, as restantes ficam indisponíveis sem novas chamadas. O arquivo STOP é verificado antes dos próximos arquivos e tentativas.

Os eventos `AI_SKIPPED`, `AI_STARTED`, `AI_THROTTLED`, `AI_CACHE_HIT`, `AI_ATTEMPT` e `AI_FINISHED` registram arquivo, duração, tentativa, código HTTP, status seguro do provedor e resultado. Corpos HTTP, documentos e chaves não são copiados para os logs. As evidências registram versão das instruções, hash do contexto, tipo/completude da fonte, resposta validada e discrepâncias. `ai_summary` no resumo separa tentativas HTTP, decisões válidas, arquivos indisponíveis e diferenças encontradas. Sucesso HTTP não equivale a uma análise aprovada.

Não salve chaves no código ou nos logs. A assinatura Google AI de estudante e o faturamento da API são separados. Verifique se o projeto está no Free Tier e tem cota. O programa não ativa faturamento. O experimento de navegador ChatGPT foi retirado para reduzir dependências e porque exigia confirmação humana.

**Histórico da versão anterior (antes da melhoria multimodal):** o lote consultivo de 17 arquivos levou 603,661 segundos; 2 análises retornaram decisões válidas e 15 ficaram indisponíveis (8 timeouts e 7 erros classificados como `ValueError`). Os totais pelas regras foram preservados. Portanto, a conectividade e o contrato tiveram respostas reais válidas, mas a IA externa **ainda não está homologada para decidir o lote inteiro**. Naquela versão, ela ficava desligada por padrão. A configuração atual aplica medidas automáticas apenas nas exceções, por solicitação do responsável pelo case. Os detalhes deste ensaio estão no histórico local da execução `20261114T103109Z_a649ee0e`; os tipos de erro registrados não permitem atribuir todos os casos a uma causa específica.

**Ensaio multimodal desta versão:** três documentos do case (Excel, PDF e PNG), em modo consultivo, foram processados em 127,685 segundos, incluindo a extração local. Para essa rodada foram usados 30 segundos por operação, duas tentativas e orçamento de 120 segundos para IA. Nenhum retornou uma decisão válida: ocorreram HTTP 503 (`UNAVAILABLE`), HTTP 429 (`RESOURCE_EXHAUSTED`) e timeouts, agora discriminados por tentativa. A validação de payload, integridade, todas as abas e bloqueios foi feita por testes automatizados; a leitura multimodal real ainda não está homologada. A evidência local está em `.local/validacao_multimodal/20261114T141520Z_1981ae30/`. Nenhuma nova cobrança foi habilitada.

**Ensaio do roteamento seletivo:** o lote de 17 documentos passou pelas regras; 12 dispensaram a IA e 5 foram classificados como exceção. Houve 3 tentativas HTTP, todas com 429 (`RESOURCE_EXHAUSTED`); o controle de falhas pausou os dois arquivos restantes. O resultado financeiro permaneceu em 12 ok, 4 revisar e 1 falhou. O ensaio usou uma tentativa, 15 segundos por operação e 60 segundos de orçamento para IA, sem alterar os parâmetros padrão de operação. Nenhuma recomendação foi apresentada como concluída quando o serviço não respondeu.

O ajuste supervisionado deve usar exemplos classificados pela equipe e avaliação em casos separados. Ajustar instruções não é treinar os pesos do modelo. O modo consultivo continua disponível para avaliar novas políticas. As ações automáticas atuais são limitadas à seleção e retenção dos valores no consolidado; ampliar a política exige avaliação dos casos e critérios definidos com a gestão.

### Causa dos erros de API e controle de consumo

O painel do projeto Free Tier mostrou para Gemini 3.8 Flash: 5 RPM, 20 RPD e 250 mil TPM. Os picos observados foram 7 RPM e 33 RPD, ultrapassando os limites. Os testes e retries anteriores contribuíram para esse consumo; repetir em poucos segundos não recupera a cota diária. Uma chamada de diagnóstico também retornou 503 com mensagem explícita de alta demanda do modelo.

O modelo passou a ser `gemini-3.1-flash-lite`: o painel atribuiu 15 RPM, 500 RPD e 250 mil TPM a esse modelo no mesmo projeto gratuito. Ele suporta PDF, imagem e JSON estruturado. O raciocínio low reduz latência nessa conferência delimitada; as contas financeiras continuam no Python. Nenhum faturamento foi habilitado e não há rotação de contas, chaves ou projetos.

O arquivo local `gemini.sqlite3` guarda análises válidas, tentativas e bloqueios. A resposta é reutilizada por até 30 dias somente quando modelo, instruções, esquema, conteúdo da requisição e documentos forem iguais. Uma troca de período, evidências, instruções ou modelo invalida a chave. Respostas inválidas não entram no cache. A política consultiva/automática é reaplicada no momento de usar a resposta. `--no-cache` força uma nova análise, respeitando os limites.

Os limites locais são 12 RPM e 450 RPD, deixando margem abaixo da cota observada. Cada tentativa é contada antes do envio, inclusive retries. A reserva e o espaçamento persistem entre execuções; reiniciar o robô não zera os contadores. Erro de cota diária interrompe novas tentativas e registra o horário da próxima janela em vez de insistir. Os limites por dia usam a meia-noite do Pacífico, conforme a API. Esses contadores não enxergam consumo por outros programas; o retorno do Google continua sendo a autoridade e também é tratado.

A disponibilidade externa não é garantida. Quando uma análise não puder ser concluída, o consolidado é entregue com a retenção e a ressalva correspondentes. Em uma nova execução, extrações e análises válidas são reaproveitadas; somente o que ainda necessita de análise demanda novas chamadas. Isso mantém a geração automática sem inventar aprovação ou corrigir o saldo de um extrato por inferência.

**Validação real após a correção:** com Gemini 3.1 Flash-Lite e o modo automático padrão, as cinco exceções do lote receberam respostas JSON válidas, usando os documentos originais. Foram seis tentativas HTTP (um timeout seguido de recuperação), com zero análises indisponíveis no resultado final. O processamento levou 124,315 segundos nesta máquina, incluindo extração local. As leituras de saldo, conta e competência dos quatro extratos conferiram com os documentos; o comunicado foi reconhecido como não sendo extrato. A cópia foi excluída dos totais pela regra determinística mesmo com a recomendação de revisão da IA. Os 17 arquivos permaneceram registrados e o saldo incluído continuou em R$ 4.114.372,29.

**Reexecução validada:** o mesmo lote reaproveitou as cinco análises válidas, com zero novas chamadas HTTP, zero análises indisponíveis e 0,225 segundo de processamento medido (instalação/inicialização do Python não entram nessa duração). Os 12 arquivos regulares continuaram dispensando IA. Os hashes dos 17 documentos de entrada permaneceram inalterados.

### Failover de modelos

A ordem padrão é `gemini-3.8-flash` → `gemini-3.1-flash-lite` → `gemini-3.5-flash-lite`. Ela é definida por `ai_model` e `ai_fallback_models` em `config.toml`. Uma lista vazia desliga a alternância. Não há troca automática de provedor, conta, chave ou projeto.

O principal é tentado primeiro, inclusive verificando o cache daquele modelo. Cota esgotada ou espera de cota superior a cinco segundos permite seguir para o próximo. Também há alternância para timeout, falha de rede, indisponibilidade HTTP 5xx e modelo não encontrado (404). Um modelo que falha por comunicação fica em pausa por 60 segundos; bloqueios de cota diária respeitam a janela diária. Cache válido continua podendo ser usado durante a pausa.

Com alternativos configurados, é feita uma tentativa por modelo antes de seguir a ordem, dentro do orçamento compartilhado do lote. Sem alternativos, `ai_max_attempts` controla as repetições do único modelo. Os limites locais continuam valendo por modelo: para o 3.8 há o teto observado de 5 RPM/20 RPD; para os dois Lite, 15 RPM/500 RPD. O limite configurado pode ser mais restritivo.

Uma decisão válida encerra a busca, inclusive quando é revisar ou não considerar. Falhas de contrato JSON, bloqueios de segurança, autenticação/permissão e HTTP 400 não são contornados por outra resposta de modelo. Os bloqueios financeiros obrigatórios são aplicados pelo Python em todos os casos. Se nenhum modelo responder, o saldo segue retido com ressalva.

A prioridade é preservada: quando o principal voltar a ficar disponível, pode gerar uma nova análise mesmo se existir uma resposta anterior de um alternativo. O cache é específico de cada modelo. `AI_FAILOVER` registra motivo e troca; `AI_MODEL_SELECTED` registra a escolha final. As evidências guardam principal solicitado, modelo selecionado, modelos percorridos e tentativas HTTP de toda a cadeia. Planilha e HTML mostram o modelo utilizado e se houve failover.

**Validação do failover:** no lote real de 17 arquivos, as cinco exceções percorreram 3.8 (bloqueio de cota diária já registrado) → 3.1 Flash-Lite (análise válida em cache), sem novas chamadas HTTP e sem análises indisponíveis. O terceiro modelo, 3.5 Flash-Lite, também respondeu em teste HTTP com um PDF original e identificou corretamente o comunicado como documento que não deve ser considerado no fechamento. Os testes simulam indisponibilidade dos dois primeiros, esgotamento de toda a cadeia e a preservação das esperas indicadas pelo provedor.

### Testar e preparar a entrega

```powershell
.\.venv\Scripts\python.exe -m ruff check goener tests
.\.venv\Scripts\python.exe -m pytest --cov=goener --cov-report=term-missing
.\.venv\Scripts\python.exe -m goener.entrega
```

Validação local da estrutura simplificada: **89 testes passaram**, incluindo OCR dos documentos do case. A versão anterior teve chamadas HTTP reais válidas; a validação da versão multimodal é registrada abaixo, sem presumir homologação do lote inteiro. Os testes automatizados usam HTTP simulado para não depender de chave ou consumir cota na integração contínua.

A instalação manual usa apenas `requirements-lock.txt` e `pip install -e . --no-deps`. Para testar sem inferência OCR, acrescente `-m "not ocr"`. A integração contínua roda os testes e o lote em Windows e Linux, com Python 3.12. O ZIP inclui fonte, entradas, modelo, testes e resultado; não inclui chaves ou dados locais de execução. Nenhum agendamento foi criado nesta máquina.

### Ganho estimado em horas

Hipótese editável: **8 minutos por arquivo** — abrir/localizar (2), copiar campos (3), conferir (2), registrar/organizar (1). Revisão: **3 minutos por pendência**. São premissas, não medições da equipe.

- Manual: `17 × 8 ÷ 60 = 2,27 horas`.
- Assistido: `tempo medido do robô ÷ 3600 + 5 × 3 ÷ 60`.
- Capacidade liberada: horas manuais menos horas assistidas; próxima de 2 horas por lote nas execuções locais sem IA.

O relatório usa a duração efetivamente medida, incluindo a análise de IA quando ativada. Não há conversão da economia para reais nem histórico suficiente para afirmar uma tendência.

### Limites e próximos passos

O projeto atende aos layouts fornecidos. Não autentica extratos nem substitui a conciliação com o razão contábil. A revisão financeira das exceções ocorre depois da geração automática do consolidado.

1. Cadastrar contas esperadas para detectar extratos ausentes.
2. Ampliar leitores e testes com outros layouts autorizados.
3. Avaliar os erros da IA com a gestão antes de adotar decisões automáticas. Power BI permanece adiado.

### Referências

- [RapidOCR](https://rapidai.github.io/RapidOCRDocs/main/en/install_usage/rapidocr/usage/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
- [Gemini: contrato HTTP](https://ai.google.dev/api/generate-content)
- [Google AI: assinatura e API](https://aistudio.google.com/docs/google-ai-plans)
- [Gemini: cotas e faturamento](https://ai.google.dev/gemini-api/docs/billing)

- [Gemini: documentos, imagens e limites da leitura de Excel](https://aistudio.google.com/docs/document-processing?generate-content=true)
- [Gemini: parâmetros de raciocínio](https://ai.google.dev/api/generate-content)

- [Gemini 3.1 Flash-Lite: entradas multimodais e JSON](https://aistudio.google.com/docs/models/gemini-3.1-flash-lite)
- [Gemini: cotas e reinício diário](https://ai.google.dev/gemini-api/docs/rate-limits)

- [Gemini 3.5 Flash-Lite: documentos e saída estruturada](https://ai.google.dev/gemini-api/docs/models/gemini-3.5-flash-lite)
