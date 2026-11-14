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

O terminal mostra a pasta do resultado. `.local/outputs/latest.json` aponta para a última execução concluída. Cada execução preserva suas evidências. A pasta `entrega/` reúne uma única versão pronta para apresentar.

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
extratos/           entradas fictícias do case
entrega/            uma única entrega e gerar.py para empacotá-la
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

`ai_mode = "off"` mantém a análise externa desligada. Para usar Gemini, configure `GEMINI_API_KEY` no ambiente; os parâmetros do provedor já estão em `config.toml`. A configuração inicial da chave ocorre na implantação, sem login ou confirmação durante os lotes.

```powershell
.\.venv\Scripts\python.exe -m goener --ai-mode consultivo
.\.venv\Scripts\python.exe -m goener --ai-mode automatico
```

**Consultivo:** registra recomendação e justificativa, preservando os totais pelas regras. **Automático:** a IA pode reter o saldo; para incluí-lo, ele também precisa passar nas verificações obrigatórias. Duplicidade, competência incorreta, campos ausentes e divergências não são liberados pela IA.

O contrato JSON contém apenas `decision` (`considerar`, `nao_considerar` ou `revisar`), `reason` e `evidence_codes`. Campos extras e evidências inventadas são rejeitados localmente. A chamada recebe dados estruturados mínimos e nenhuma ferramenta para executar ações. Textos de origem são tratados como dados, nunca como instruções. A IA não altera o saldo extraído.

Gemini usa HTTPS `generateContent`, esquema JSON e chave no cabeçalho. Também é aceita `GOOGLE_API_KEY`. OpenAI usa a Responses API: altere o provedor para `openai`, escolha o modelo e configure `OPENAI_API_KEY`. Não há troca automática entre provedores. Até três tentativas atendem erros HTTP transitórios; bloqueio, cota esgotada, resposta parcial ou JSON inválido retêm o saldo no modo automático.

Não salve chaves no código ou nos logs. A assinatura Google AI de estudante e o faturamento da API são separados. Verifique se o projeto está no Free Tier e tem cota. O programa não ativa faturamento. O experimento de navegador ChatGPT foi retirado para reduzir dependências e porque exigia confirmação humana.

O ajuste supervisionado deve usar exemplos classificados pela equipe e avaliação em casos separados. Ajustar instruções não é treinar os pesos do modelo. A gestão define quando os resultados permitem passar do consultivo para o automático.

### Testar e preparar a entrega

```powershell
.\.venv\Scripts\python.exe -m ruff check goener tests entrega/gerar.py
.\.venv\Scripts\python.exe -m pytest --cov=goener --cov-report=term-missing
.\.venv\Scripts\python.exe entrega/gerar.py
```

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
