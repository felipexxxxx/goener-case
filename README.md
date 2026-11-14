# Consolidação de extratos — case Goener

Automação em Python que lê uma pasta com PDFs, planilhas e imagens, extrai os saldos e gera o consolidado no modelo fornecido. A execução também confere as movimentações e registra o que precisa de revisão.

Os arquivos deste repositório são fictícios e pertencem ao case técnico. O programa funciona sem chat, chave de API ou preenchimento manual de saldos.

## Para quem vai usar

### O que a automação entrega

| Arquivo de saída | Para que serve |
| --- | --- |
| `consolidado.xlsx` | Uma linha por arquivo de entrada, nas oito colunas do modelo. |
| `relatorio.html` | Consulta local com filtro por status e evidência de cada arquivo. Abra no navegador. |
| `execucao.log` | Registro legível do processamento e das ocorrências. |
| `execucao.jsonl` | O mesmo acompanhamento em eventos estruturados. |
| `evidencias.json` | Dados extraídos, localização do saldo, confiança do OCR e conciliação. |
| `resumo.json` | Quantidades, duração, versão e saldo aprovado. |
| `manifesto.json` | Hashes dos arquivos de entrada, modelo e entregáveis para conferir integridade. |

### Como executar no Windows

Pré-requisito: **Python 3.12** instalado. Python 3.11 e 3.13 também são aceitos pelo projeto; a validação principal usa 3.12. Internet é necessária na instalação das dependências. Depois disso, o processamento é local.

1. Baixe ou clone o repositório.
2. Coloque os arquivos diretamente em `extratos/`.
3. Abra o PowerShell na pasta do projeto e execute:

```powershell
.\executar.ps1
```

Na primeira vez, o script cria `.venv`, instala as versões fixadas e executa o fluxo inteiro. Nas próximas, reaproveita o ambiente. Não é preciso abrir o Excel, separar PDFs ou informar os valores.

Se a política do computador não permitir scripts PowerShell, use os comandos Python abaixo, sem mudar as políticas do sistema:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-lock.txt
.\.venv\Scripts\python.exe -m pip install -e . --no-deps
.\.venv\Scripts\python.exe -m goener --config config.toml
```

O terminal informa a pasta da execução em `outputs/`. `outputs/latest.json` aponta para a última publicação concluída. As execuções anteriores ficam preservadas.

### Como interpretar a planilha

| Status | Significado | Ação sugerida |
| --- | --- | --- |
| **ok** | Campos essenciais extraídos e verificações sem ressalvas. | Usar no fechamento da competência. |
| **revisar** | Há um valor extraído, mas uma regra encontrou uma ressalva. | Ler a observação e conferir a origem. |
| **falhou** | Não foi possível obter dados essenciais, o documento é inválido ou o processamento foi cancelado. | Solicitar o arquivo correto ou reprocessar após resolver a causa. |

Saldo negativo é válido. Um saldo não é trocado por outro para fazer a conta fechar. Valores com ressalvas permanecem na planilha, mas ficam fora do **saldo aprovado** exibido no relatório.

### A duplicidade do download

`EXT_102026_LagoaAzul (1).pdf` é uma cópia exata de `EXT_102026_LagoaAzul.pdf`. O programa compara o conteúdo com SHA-256, mantém as duas linhas e identifica a cópia em `Observação`. A cópia não participa de somas de acompanhamento.

Isso é diferente de um **possível lançamento repetido dentro do extrato**: nesse caso, a automação sinaliza a semelhança para conferência e não apaga a movimentação.

### Trocar a competência ou a pasta

O case vem configurado para **outubro de 2026**. Para outro mês:

```powershell
.\executar.ps1 -Period 2026-11
```

Também é possível editar `config.toml`. Os caminhos são relativos a esse arquivo. Removendo a linha `period`, o comando Python usa o mês anterior na zona `America/Sao_Paulo`; o atalho PowerShell mantém seu parâmetro padrão do case, que pode ser alterado.

### O que acontece quando um arquivo dá problema

O lote continua, e o arquivo recebe sua própria linha com a causa. O processo não declara publicação concluída quando há falha de infraestrutura, como pasta inexistente, execução concorrente ou problema ao gravar a saída.

Os arquivos originais não são movidos, corrigidos ou apagados. Feche o consolidado se o seu programa de planilhas estiver bloqueando uma gravação de arquivos copiados manualmente.

## Para quem vai manter

### Fluxo

```mermaid
flowchart LR
    A[Pasta de entrada] --> B[Hash e limites]
    B --> C{Extração já disponível?}
    C -->|sim| D[Cache local]
    C -->|não| E[PDF / Excel / OCR local]
    D --> F[Validação e conciliação]
    E --> F
    F --> G[Duplicidades e classificação]
    G --> H[Modelo Excel + evidências + logs]
    H --> I[Reabertura e conferência]
    I --> J[Publicação da execução]
```

### Organização do código

```text
src/goener/
  cli.py          # argumentos e códigos de saída
  config.py       # configuração e competência
  domain.py       # valores monetários e contratos de dados
  readers.py      # leitores de Excel, PDF e imagem
  validation.py   # conciliação e duplicidades
  runtime.py      # lock, cache, hash e log
  reporting.py    # planilha no modelo e relatório local
  pipeline.py     # coordenação do fluxo completo
extratos/         # entradas fictícias do case
templates/        # modelo original
tests/            # testes de regras, leitores e operação
docs/             # premissas e decisões
scripts/          # execução em Linux e preparação da entrega
```

Os leitores retornam o mesmo contrato `Statement`. CLI e uma futura API podem chamar `pipeline.run(config)` sem duplicar as regras financeiras. Não há navegador automatizado: esses arquivos podem ser lidos diretamente, o que reduz dependências e pontos de falha.

### Leitura e validação

- **Excel:** identifica as duas estruturas do case pelos cabeçalhos, preserva contas como texto e lê os valores sem executar fórmulas. Fórmulas, macros e relacionamentos externos são rejeitados.
- **PDF:** usa texto e posições das colunas. O débito é interpretado pelo sufixo `D` ou pela coluna de débito. PDFs digitalizados usam OCR como alternativa; layouts novos podem exigir ajuste do leitor.
- **Imagem:** usa RapidOCR com modelos PP-OCRv4 em CPU. Prints longos são lidos em tiras sobrepostas. Quando um sinal monetário desaparece na primeira leitura, a região é relida ampliada. Se o sinal continuar incerto, o arquivo exige revisão.
- **Conciliação:** compara cada saldo com o anterior mais a movimentação e compara o saldo final declarado com o último saldo e com a soma do período. Cálculos usam `Decimal`, com centavos.
- **Duplicidades:** conteúdo igual é identificado por hash. Conteúdos diferentes para a mesma combinação banco/conta/período também exigem revisão. Lançamentos semelhantes são suspeitas, não exclusões automáticas.
- **Segurança dos dados:** documentos são tratados como dados. Texto do arquivo não vira comando nem fórmula na saída. O HTML escapa os textos recebidos. Há limites de tamanho, páginas, pixels e descompactação.

A confiança do OCR não garante acerto. A conciliação agrega uma verificação independente, mas também não prova a autenticidade do documento. Nenhum arquivo é enviado a uma API de IA.

### Runtime e retomada

`runtime/` é gerado e fica fora do Git:

- `logs/`: eventos com execução, arquivo, hash, tentativa, duração e resultado final;
- `extractions.sqlite3`: checkpoint das extrações, identificado por conteúdo, versão/código do leitor e opções relevantes;
- `execution.lock`: impede duas execuções no mesmo runtime;
- `STOP`: interruptor opcional entre arquivos.

Uma nova execução pode reaproveitar a extração. Competência, conciliação e duplicidade são verificadas novamente. Assim, o cache não reaproveita um `ok` de outro contexto. O programa confere se o arquivo mudou durante a leitura antes de salvar o checkpoint.

Os logs separam tentativas transitórias do resultado final. Para contar falhas de documentos, use os eventos `FILE_FINISHED` ou `resumo.json`, não a quantidade de linhas que contêm a palavra erro.

### Comandos de operação

```powershell
# Reextrair tudo, sem consultar ou preencher o cache
.\.venv\Scripts\python.exe -m goener --no-cache

# Desligar OCR; imagens continuam registradas, com falha explícita
.\.venv\Scripts\python.exe -m goener --no-ocr

# Gerar a entrega e retornar código 2 se houver ressalvas
.\.venv\Scripts\python.exe -m goener --strict

# Solicitar parada: o arquivo atual termina, os restantes ficam cancelados
New-Item -ItemType File -Path runtime\STOP -Force

# Retomar após a parada
Remove-Item -LiteralPath runtime\STOP
.\.venv\Scripts\python.exe -m goener
```

| Código de saída | Interpretação |
| --- | --- |
| `0` | Fluxo concluído e publicado. Pode conter ressalvas de documentos. |
| `1` | Erro de configuração ou infraestrutura. Conferir terminal e runtime. |
| `2` | Fluxo publicado em modo estrito, com `revisar` ou `falhou`. |
| `3` | STOP acionado; entrega parcial publicada com cancelamentos explícitos. |
| `130` | Interrupção pelo operador. Retomar com nova execução. |

Se um encerramento abrupto deixar o lock, confira o PID gravado e confirme que o processo não está mais ativo antes de remover `runtime/execution.lock`. O programa não remove esse bloqueio por conta própria.

### Testes e integração contínua

```powershell
.\.venv\Scripts\python.exe -m ruff check src tests scripts
.\.venv\Scripts\python.exe -m pytest --cov=goener --cov-report=term-missing
```

Para uma rodada rápida sem inferência OCR, acrescente `-m "not ocr"` ao pytest. A suíte completa lê os documentos reais do case e testa centavos, negativos, falhas, duplicidade, reprocessamento, parada, lock, competência, segurança da saída e publicação.

A integração contínua roda em Windows e Linux com Python 3.12, executa os testes e gera uma nova consolidação a partir de `extratos/`. As versões estão registradas em `requirements-lock.txt`. Commits separam estrutura, processamento, testes e entrega.

### Instalação em Linux

```sh
sh scripts/run.sh --period 2026-10
```

Para agendar, use o Agendador de Tarefas do Windows ou `cron` chamando o mesmo comando. Não foi criado um agendamento nesta máquina. No modo estrito, um agendador pode sinalizar ressalvas pelo código de saída sem deixar de gerar a planilha.

### Limites e evolução

Este projeto atende aos layouts fornecidos e rejeita ambiguidade essencial. Não substitui conciliação bancária com o razão contábil, não autentica extratos e não integra contas bancárias. A revisão das exceções é uma etapa financeira posterior; a geração do consolidado é automática.

Próximos passos, por prioridade:

1. **P1 — Cobertura de entrada:** cadastrar as contas esperadas por competência para detectar extratos ausentes.
2. **P2 — Novos layouts:** ampliar o conjunto de testes e leitores com documentos autorizados de outros bancos.
3. **P3 — Aplicação web:** reaproveitar o pipeline numa API Python com upload incremental e interface de acompanhamento. A implementação web está fora desta primeira entrega.

As decisões e os motivos estão em [docs/decisoes.md](docs/decisoes.md). A documentação de uma página fica no pacote `entrega/`.

## Referências técnicas

- [RapidOCR: uso e campos de saída](https://rapidai.github.io/RapidOCRDocs/main/en/install_usage/rapidocr/usage/)
- [openpyxl: leitura e gravação de planilhas](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
